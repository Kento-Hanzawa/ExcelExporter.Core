using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reactive.Disposables;
using ExcelExporterCore.Internal;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporterCore
{
    internal sealed class ExcelPackageCore : IDisposable
    {
        private readonly IComObjectWrappedDisposable<Application> _applicationScope;
        private readonly IComObjectWrappedDisposable<Workbooks> _workbooksScope;
        private readonly IComObjectWrappedDisposable<Workbook> _workbookScope;
        private readonly CompositeDisposable _referencesWorkbookList;



        public ExcelPackageCore(FileInfo excel)
            : this(excel, Array.Empty<FileInfo>())
        {
        }

        public ExcelPackageCore(FileInfo excel, IEnumerable<FileInfo> references)
        {
            if (excel == null)
                throw new ArgumentNullException(nameof(excel));
            if (!excel.Exists)
                throw new FileNotFoundException($"エクセルファイルが見つかりません。", excel.FullName);

            references = references != null ? references.Distinct(new FileFullNameComparer()) : Array.Empty<FileInfo>();
            if (references.Any(file => !file.Exists))
            {
                FileInfo notFound = references.FirstOrDefault(file => !file.Exists);
                throw new FileNotFoundException($"外部参照エクセルファイルのいずれかが見つかりません。", notFound.FullName);
            }

            try
            {
                _applicationScope = ComObjectWrappedDisposable.CreateApplication();
                _workbooksScope = _applicationScope.ComObject.Workbooks.AsWrappedDisposable();

                // 外部参照エクセルは、ターゲットエクセル内のリンクデータが [#REF!] に更新されるのを防ぐために使用します。
                // 先に外部参照エクセルを開いておくことで、データ更新を未然に防げます。
                _referencesWorkbookList = new CompositeDisposable();
                foreach (var file in references)
                {
                    _referencesWorkbookList.Add(_workbooksScope.ComObject.Open(file.FullName, UpdateLinks: XlUpdateLinks.xlUpdateLinksAlways, ReadOnly: true).AsWrappedDisposable());
                }

                _workbookScope = _workbooksScope.ComObject.Open(excel.FullName, UpdateLinks: XlUpdateLinks.xlUpdateLinksAlways, ReadOnly: true).AsWrappedDisposable();
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        public void Dispose()
        {
            // Dispose の順番はコンストラクタ―内の作成順と逆になるようにします。
            _workbookScope?.Dispose();
            _referencesWorkbookList?.Dispose();
            _workbooksScope?.Dispose();
            _applicationScope?.Dispose();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetIndex">1 をベースにしたシート番号。</param>
        /// <param name="dest"></param>
        /// <param name="fileFormat"><see href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat"/> を参照。</param>
        /// <returns></returns>
        public ExcelExportResultCore ExportSheet(int sheetIndex, FileInfo dest, int fileFormat)
        {
            if (dest is null)
                throw new ArgumentNullException(nameof(dest));

            using (var worksheetScope = GetWorksheet(sheetIndex))
            {
                if (worksheetScope == null)
                    throw new Exception($"ワークシート番号 {sheetIndex} は存在しません。");

                ExportCore(worksheetScope, dest, (XlFileFormat)fileFormat, out var rangeString);
                return new ExcelExportResultCore(worksheetScope.ComObject.Name, rangeString, dest.FullName);
            }
        }

        /// <param name="fileFormat"><see href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat"/> を参照。</param>
        public ExcelExportResultCore ExportSheet(string sheetName, FileInfo dest, int fileFormat)
        {
            if (sheetName == null)
                throw new ArgumentNullException(nameof(sheetName));
            if (dest == null)
                throw new ArgumentNullException(nameof(dest));

            using (var worksheetScope = GetWorksheet(sheetName))
            {
                if (worksheetScope == null)
                    throw new Exception($"ワークシート名 {sheetName} は存在しません。");

                ExportCore(worksheetScope, dest, (XlFileFormat)fileFormat, out var rangeString);
                return new ExcelExportResultCore(sheetName, rangeString, dest.FullName);
            }
        }

        /// <param name="fileFormat"><see href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat"/> を参照。</param>
        public IReadOnlyList<ExcelExportResultCore> ExportSheetAny(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, int> fileFormatSelector, Func<RangeInfo, bool> sheetSelector = null)
        {
            if (destSelector == null)
                throw new ArgumentNullException(nameof(destSelector));
            if (fileFormatSelector == null)
                throw new ArgumentNullException(nameof(fileFormatSelector));

            var results = new List<ExcelExportResultCore>();
            foreach (var worksheetScope in GetWorksheetAnyEnumerable())
            {
                var info = new RangeInfo(worksheetScope);
                if (sheetSelector?.Invoke(info) ?? true)
                {
                    var dest = destSelector(info);
                    ExportCore(worksheetScope, dest, (XlFileFormat)fileFormatSelector(info), out var rangeString);
                    results.Add(new ExcelExportResultCore(worksheetScope.ComObject.Name, rangeString, dest.FullName));
                }
            }
            return results;
        }

        /// <param name="fileFormat"><see href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat"/> を参照。</param>
        public ExcelExportResultCore ExportTable(string tableName, FileInfo dest, int fileFormat)
        {
            if (tableName is null)
                throw new ArgumentNullException(nameof(tableName));
            if (dest is null)
                throw new ArgumentNullException(nameof(dest));

            using (var listObjectScope = GetListObject(tableName))
            {
                if (listObjectScope is null)
                    throw new Exception($"エクセルファイルにテーブル {tableName} は存在しません。");

                ExportCore(listObjectScope, dest, (XlFileFormat)fileFormat, out var rangeString);
                return new ExcelExportResultCore(tableName, rangeString, dest.FullName);
            }
        }

        /// <param name="fileFormat"><see href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat"/> を参照。</param>
        public IReadOnlyList<ExcelExportResultCore> ExportTableAny(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, int> fileFormatSelector, Func<RangeInfo, bool> tableSelector = null)
        {
            if (destSelector == null)
                throw new ArgumentNullException(nameof(destSelector));
            if (fileFormatSelector == null)
                throw new ArgumentNullException(nameof(fileFormatSelector));

            var results = new List<ExcelExportResultCore>();
            foreach (var listObjectScope in GetListObjectAnyEnumerable())
            {
                var info = new RangeInfo(listObjectScope);
                if (tableSelector?.Invoke(info) ?? true)
                {
                    var dest = destSelector(info);
                    ExportCore(listObjectScope, dest, (XlFileFormat)fileFormatSelector(info), out var rangeString);
                    results.Add(new ExcelExportResultCore(listObjectScope.ComObject.Name, rangeString, dest.FullName));
                }
            }
            return results;
        }

        private void ExportCore(IComObjectWrappedDisposable<Worksheet> worksheetScope, FileInfo dest, XlFileFormat fileFormat, out string rangeString)
        {
            // 新しく Excel Workbook を作成し、シートの使用領域をコピーします。
            // コピーした Workbook を指定のファイルフォーマットで保存します。
            // この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
            // 読み取り対象の Workbook (managedWorkbook) を閉じるまで出力したファイルにアクセスできなくなるためです。
            using (var usedRangeScope = worksheetScope.ComObject.UsedRange.AsWrappedDisposable())
            using (var tempWorkbookScope = _workbooksScope.ComObject.Add().AsWrappedDisposable())
            using (var tempWorksheetScope = ((Worksheet)tempWorkbookScope.ComObject.ActiveSheet).AsWrappedDisposable())
            using (var destinationRangeScope = tempWorksheetScope.ComObject.Range["A1"].AsWrappedDisposable())
            {
                usedRangeScope.ComObject.Copy();
                destinationRangeScope.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);

                dest.Directory.Create();
                tempWorkbookScope.ComObject.SaveAs(dest.FullName, FileFormat: fileFormat);
                rangeString = usedRangeScope.ComObject.GetRangeString();
            }
        }

        private void ExportCore(IComObjectWrappedDisposable<ListObject> listObjectScope, FileInfo dest, XlFileFormat fileFormat, out string rangeString)
        {
            // 新しく Excel Workbook を作成し、シートの使用領域をコピーします。
            // コピーした Workbook を指定のファイルフォーマットで保存します。
            // この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
            // 読み取り対象の Workbook (managedWorkbook) を閉じるまで出力したファイルにアクセスできなくなるためです。
            using (var tableRangeScope = listObjectScope.ComObject.Range.AsWrappedDisposable())
            using (var tempWorkbookScope = _workbooksScope.ComObject.Add().AsWrappedDisposable())
            using (var tempWorksheetScope = ((Worksheet)tempWorkbookScope.ComObject.ActiveSheet).AsWrappedDisposable())
            using (var destinationRangeScope = tempWorksheetScope.ComObject.Range["A1"].AsWrappedDisposable())
            {
                tableRangeScope.ComObject.Copy();
                destinationRangeScope.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);

                dest.Directory.Create();
                tempWorkbookScope.ComObject.SaveAs(dest.FullName, FileFormat: fileFormat);
                rangeString = tableRangeScope.ComObject.GetRangeString();
            }
        }



        /// <summary>
        /// 指定したシート番号の <see cref="Worksheet"/> を取得します。シートが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        private IComObjectWrappedDisposable<Worksheet> GetWorksheet(int sheetIndex)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsWrappedDisposable())
            {
                var worksheetScope = ((Worksheet)sheetsScope.ComObject[sheetIndex]).AsWrappedDisposable();
                if (worksheetScope.ComObject == null)
                {
                    worksheetScope.Dispose();
                    return null;
                }
                return worksheetScope;
            }
        }

        /// <summary>
        /// 指定したシート名に一致する <see cref="Worksheet"/> を取得します。シートが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        private IComObjectWrappedDisposable<Worksheet> GetWorksheet(string sheetName)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsWrappedDisposable())
            {
                var worksheetScope = ((Worksheet)sheetsScope.ComObject[sheetName]).AsWrappedDisposable();
                if (worksheetScope.ComObject == null)
                {
                    worksheetScope.Dispose();
                    return null;
                }
                return worksheetScope;
            }
        }

        /// <summary>
        /// 全ての <see cref="Worksheet"/> を取得します。（遅延実行専用）
        /// </summary>
        private IEnumerable<IComObjectWrappedDisposable<Worksheet>> GetWorksheetAnyEnumerable()
        {
            return GetWorksheetAnyEnumerable(null);
        }

        /// <summary>
        /// 指定した条件を満たす全ての <see cref="Worksheet"/> を取得します。（遅延実行専用）
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="predicate"/> が <see langword="null"/> です。</exception>
        private IEnumerable<IComObjectWrappedDisposable<Worksheet>> GetWorksheetAnyEnumerable(Predicate<IComObjectWrappedDisposable<Worksheet>> predicate)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsWrappedDisposable())
            {
                for (var i = 1; i <= sheetsScope.ComObject.Count; i++)
                {
                    using (var worksheetScope = ((Worksheet)sheetsScope.ComObject[i]).AsWrappedDisposable())
                    {
                        if ((worksheetScope.ComObject != null) && (predicate?.Invoke(worksheetScope) ?? true))
                        {
                            yield return worksheetScope;
                        }
                    }
                }
            }
        }



        /// <summary>
        /// 指定したテーブル名に一致する <see cref="ListObject"/> を取得します。テーブルが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        private IComObjectWrappedDisposable<ListObject> GetListObject(string tableName)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsWrappedDisposable())
            {
                for (var sheetIndex = 1; sheetIndex <= sheetsScope.ComObject.Count; sheetIndex++)
                {
                    using (var worksheetScope = ((Worksheet)sheetsScope.ComObject[sheetIndex]).AsWrappedDisposable())
                    using (var listObjectsScope = worksheetScope.ComObject.ListObjects.AsWrappedDisposable())
                    {
                        for (var listObjIndex = 1; listObjIndex <= listObjectsScope.ComObject.Count; listObjIndex++)
                        {
                            var listObjectScope = listObjectsScope.ComObject[listObjIndex].AsWrappedDisposable();
                            if ((listObjectScope.ComObject == null) || (listObjectScope.ComObject.Name != tableName))
                            {
                                listObjectScope.Dispose();
                                continue;
                            }
                            return listObjectScope;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 全ての <see cref="ListObject"/> を取得します。（遅延実行専用）
        /// </summary>
        private IEnumerable<IComObjectWrappedDisposable<ListObject>> GetListObjectAnyEnumerable()
        {
            return GetListObjectAnyEnumerable(null);
        }

        /// <summary>
        /// 指定した条件を満たす全ての <see cref="ListObject"/> を取得します。（遅延実行専用）
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="predicate"/> が <see langword="null"/> です。</exception>
        private IEnumerable<IComObjectWrappedDisposable<ListObject>> GetListObjectAnyEnumerable(Predicate<IComObjectWrappedDisposable<ListObject>> predicate)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsWrappedDisposable())
            {
                for (var sheetIndex = 1; sheetIndex <= sheetsScope.ComObject.Count; sheetIndex++)
                {
                    using (var worksheetScope = ((Worksheet)sheetsScope.ComObject[sheetIndex]).AsWrappedDisposable())
                    using (var listObjectsScope = worksheetScope.ComObject.ListObjects.AsWrappedDisposable())
                    {
                        for (var listObjIndex = 1; listObjIndex <= listObjectsScope.ComObject.Count; listObjIndex++)
                        {
                            using (var listObjectScope = listObjectsScope.ComObject[listObjIndex].AsWrappedDisposable())
                            {
                                if ((listObjectScope.ComObject != null) && (predicate?.Invoke(listObjectScope) ?? true))
                                {
                                    yield return listObjectScope;
                                }
                            }
                        }
                    }
                }
            }
        }



        /// <summary>
        /// エクセルファイルに含まれる全てのシート名を取得します。
        /// </summary>
        /// <returns>全シート名の列挙。</returns>
        public string[] GetSheetNames()
        {
            return GetWorksheetAnyEnumerable().Select(x => x.ComObject.Name).ToArray();
        }

        /// <summary>
        /// エクセルファイルに指定したシートが存在するかを判断します。
        /// </summary>
        /// <param name="sheetName">シート名。</param>
        /// <returns>シートが存在する場合は <see langword="true"/>。存在しない場合は <see langword="false"/>。</returns>
        public bool ContainsSheet(string sheetName)
        {
            using (var worksheetScope = GetWorksheet(sheetName))
            {
                return !(worksheetScope is null);
            }
        }



        /// <summary>
        /// エクセルファイルに含まれる全てのテーブルの名前を取得します。
        /// </summary>
        /// <returns>全テーブル名の列挙。</returns>
        public string[] GetTableNames()
        {
            return GetListObjectAnyEnumerable().Select(x => x.ComObject.Name).ToArray();
        }

        /// <summary>
        /// エクセルファイルに指定したテーブルが存在するかを判断します。
        /// </summary>
        /// <param name="tableName">テーブル名。</param>
        /// <returns>テーブルが存在する場合は <see langword="true"/>。存在しない場合は <see langword="false"/>。</returns>
        public bool ContainsTable(string tableName)
        {
            using (var listObjectScope = GetListObject(tableName))
            {
                return !(listObjectScope is null);
            }
        }



        private sealed class FileFullNameComparer : IEqualityComparer<FileInfo>
        {
            public FileFullNameComparer()
            {
            }

            public bool Equals(FileInfo x, FileInfo y)
            {
                return x.FullName == y.FullName;
            }

            public int GetHashCode(FileInfo obj)
            {
                return obj.FullName.GetHashCode();
            }
        }
    }
}
