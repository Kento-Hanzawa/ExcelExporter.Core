using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reactive.Disposables;
using System.Runtime.Versioning;
using ExcelExporterCore.Internal;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporterCore
{
    /// <summary>
    /// エクセルファイルを扱うための機能を提供します。
    /// </summary>
    [SupportedOSPlatform("windows")]
    public sealed class ExcelPackage : IDisposable
    {
        private readonly IComObjectWrappedDisposable<Application> _applicationScope;
        private readonly IComObjectWrappedDisposable<Workbooks> _workbooksScope;
        private readonly IComObjectWrappedDisposable<Workbook> _workbookScope;
        private readonly CompositeDisposable _referencesWorkbookList;

        /// <summary>
        /// <see cref="ExcelPackage"/> クラスの新しいインスタンスを作成します。
        /// </summary>
        /// <param name="excel">エクセルファイル。</param>
        public ExcelPackage(FileInfo excel)
            : this(excel, Array.Empty<FileInfo>())
        {
        }

        /// <summary>
        /// <see cref="ExcelPackage"/> クラスの新しいインスタンスを作成します。
        /// </summary>
        /// <param name="excel">エクセルファイル。</param>
        /// <param name="references">参照維持のためのエクセルファイル。</param>
        public ExcelPackage(FileInfo excel, IEnumerable<FileInfo> references)
        {
            if (excel is null)
                throw new ArgumentNullException(nameof(excel));
            if (!excel.Exists)
                throw new FileNotFoundException($"エクセルファイルが見つかりません。", excel.FullName);

            references = references != null ? references.Distinct(new FileFullNameComparer()) : Array.Empty<FileInfo>();
            if (references.Any(file => !file.Exists))
            {
                FileInfo notFound = references.FirstOrDefault(file => !file.Exists)!;
                throw new FileNotFoundException($"外部参照エクセルファイルのいずれかが見つかりません。", notFound.FullName);
            }

            try
            {
                _applicationScope = ComObjectWrappedDisposable.CreateApplication();
                _workbooksScope = _applicationScope.ComObject.Workbooks.AsDisposable();

                // 外部参照エクセルは、ターゲットエクセル内のリンクデータが [#REF!] に更新されるのを防ぐために使用します。
                // 先に外部参照エクセルを開いておくことで、データ更新を未然に防げます。
                _referencesWorkbookList = new CompositeDisposable();
                foreach (var file in references)
                {
                    _referencesWorkbookList.Add(_workbooksScope.ComObject.Open(file.FullName, UpdateLinks: XlUpdateLinks.xlUpdateLinksAlways, ReadOnly: true).AsDisposable());
                }

                _workbookScope = _workbooksScope.ComObject.Open(excel.FullName, UpdateLinks: XlUpdateLinks.xlUpdateLinksAlways, ReadOnly: true).AsDisposable();
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        /// <summary>
        /// エクセルファイルを閉じます。
        /// </summary>
        public void Dispose()
        {
            // Dispose の順番はコンストラクタ―内の作成順と逆になるようにします。
            _workbookScope?.Dispose();
            _referencesWorkbookList?.Dispose();
            _workbooksScope?.Dispose();
            _applicationScope?.Dispose();
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

        /// <summary>
        /// 指定のインデックスに対応したシートをファイルへエクスポートします。
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="dest"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentException"><paramref name="sheetIndex"/> に対応したシートが存在しません。</exception>
        /// <param name="sheetIndex">出力するシートのインデックス値。( 1 ベース)</param>
        /// <param name="dest">出力先のファイルを表す <see cref="FileInfo"/> のインスタンス。</param>
        /// <param name="fileFormat">出力フォーマット。詳細は <see href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat"/> を参照してください。</param>
        /// <returns>出力の結果を表すオブジェクト。</returns>
        public ExportResult ExportSheet(int sheetIndex, FileInfo dest, int fileFormat)
        {
            if (dest is null)
                throw new ArgumentNullException(nameof(dest));

            using (var worksheetScope = GetWorksheet(sheetIndex))
            {
                if (worksheetScope is null)
                {
                    throw new ArgumentException($"ワークシートインデックス {sheetIndex} は存在しません。", nameof(sheetIndex));
                }

                ExportCore(worksheetScope, dest, (XlFileFormat)fileFormat);
                return new ExportResult(new RangeInfo(worksheetScope.ComObject), dest.FullName);
            }
        }

        /// <summary>
        /// 指定の名前に対応したシートをファイルへエクスポートします。
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="sheetName"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentNullException"><paramref name="dest"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentException"><paramref name="sheetName"/> に対応したシートが存在しません。</exception>
        /// <param name="sheetName">出力するシートの名前。</param>
        /// <param name="dest">出力先のファイルを表す <see cref="FileInfo"/> のインスタンス。</param>
        /// <param name="fileFormat">出力フォーマット。詳細は <see href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat"/> を参照してください。</param>
        /// <returns>出力の結果を表すオブジェクト。</returns>
        public ExportResult ExportSheet(string sheetName, FileInfo dest, int fileFormat)
        {
            if (sheetName is null)
                throw new ArgumentNullException(nameof(sheetName));
            if (dest is null)
                throw new ArgumentNullException(nameof(dest));

            using (var worksheetScope = GetWorksheet(sheetName))
            {
                if (worksheetScope is null)
                {
                    throw new ArgumentException($"ワークシート名 {sheetName} は存在しません。", nameof(sheetName));
                }

                ExportCore(worksheetScope, dest, (XlFileFormat)fileFormat);
                return new ExportResult(new RangeInfo(worksheetScope.ComObject), dest.FullName);
            }
        }

        /// <summary>
        /// <paramref name="sheetSelector"/> の戻り値が <see langword="true"/> のシートをファイルへエクスポートします。
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="destSelector"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentNullException"><paramref name="fileFormatSelector"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentNullException"><paramref name="sheetSelector"/> が <see langword="null"/> です。</exception>
        /// <param name="destSelector">シートごとの出力先ファイルを取得するデリゲートを指定します。</param>
        /// <param name="fileFormatSelector">シートごとの出力フォーマットを取得するデリゲートを指定します。</param>
        /// <param name="sheetSelector">出力するシートを判定するデリゲートを指定します。</param>
        /// <returns>出力の結果を表すオブジェクト。</returns>
        public IReadOnlyList<ExportResult> ExportSheetAny(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, int> fileFormatSelector, Func<RangeInfo, bool> sheetSelector)
        {
            if (destSelector is null)
                throw new ArgumentNullException(nameof(destSelector));
            if (fileFormatSelector is null)
                throw new ArgumentNullException(nameof(fileFormatSelector));
            if (sheetSelector is null)
                throw new ArgumentNullException(nameof(sheetSelector));

            var results = new List<ExportResult>();
            foreach (var worksheetScope in GetWorksheetAnyEnumerable())
            {
                var info = new RangeInfo(worksheetScope.ComObject);
                if (sheetSelector.Invoke(info))
                {
                    var dest = destSelector(info);
                    ExportCore(worksheetScope, dest, (XlFileFormat)fileFormatSelector(info));
                    results.Add(new ExportResult(info, dest.FullName));
                }
            }
            return results;
        }

        /// <summary>
        /// エクセルファイルに含まれる全てのシートをファイルへエクスポートします。
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="destSelector"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentNullException"><paramref name="fileFormatSelector"/> が <see langword="null"/> です。</exception>
        /// <param name="destSelector">シートごとの出力先ファイルを取得するデリゲートを指定します。</param>
        /// <param name="fileFormatSelector">シートごとの出力フォーマットを取得するデリゲートを指定します。</param>
        /// <returns>出力の結果を表すオブジェクト。</returns>
        public IReadOnlyList<ExportResult> ExportSheetAll(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, int> fileFormatSelector)
        {
            if (destSelector is null)
                throw new ArgumentNullException(nameof(destSelector));
            if (fileFormatSelector is null)
                throw new ArgumentNullException(nameof(fileFormatSelector));

            return ExportSheetAny(destSelector, fileFormatSelector, static _ => true);
        }

        /// <summary>
        /// 指定の名前に対応したテーブルをファイルへエクスポートします。
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="tableName"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentNullException"><paramref name="dest"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentException"><paramref name="tableName"/> に対応したシートが存在しません。</exception>
        /// <param name="tableName">出力するテーブルの名前。</param>
        /// <param name="dest">出力先のファイルを表す <see cref="FileInfo"/> のインスタンス。</param>
        /// <param name="fileFormat">出力フォーマット。詳細は <see href="https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat"/> を参照してください。</param>
        /// <returns>出力の結果を表すオブジェクト。</returns>
        public ExportResult ExportTable(string tableName, FileInfo dest, int fileFormat)
        {
            if (tableName is null)
                throw new ArgumentNullException(nameof(tableName));
            if (dest is null)
                throw new ArgumentNullException(nameof(dest));

            using (var listObjectScope = GetListObject(tableName))
            {
                if (listObjectScope is null)
                {
                    throw new ArgumentException($"エクセルファイルにテーブル {tableName} は存在しません。", nameof(tableName));
                }

                ExportCore(listObjectScope, dest, (XlFileFormat)fileFormat);
                return new ExportResult(new RangeInfo(listObjectScope.ComObject), dest.FullName);
            }
        }

        /// <summary>
        /// <paramref name="tableSelector"/> の戻り値が <see langword="true"/> のテーブルをファイルへエクスポートします。
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="destSelector"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentNullException"><paramref name="fileFormatSelector"/> が <see langword="null"/> です。</exception>
        /// <exception cref="ArgumentNullException"><paramref name="tableSelector"/> が <see langword="null"/> です。</exception>
        /// <param name="destSelector">テーブルごとの出力先ファイルを取得するデリゲートを指定します。</param>
        /// <param name="fileFormatSelector">テーブルごとの出力フォーマットを取得するデリゲートを指定します。</param>
        /// <param name="tableSelector">出力するテーブルを判定するデリゲートを指定します。</param>
        /// <returns>出力の結果を表すオブジェクト。</returns>
        public IReadOnlyList<ExportResult> ExportTableAny(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, int> fileFormatSelector, Func<RangeInfo, bool> tableSelector)
        {
            if (destSelector is null)
                throw new ArgumentNullException(nameof(destSelector));
            if (fileFormatSelector is null)
                throw new ArgumentNullException(nameof(fileFormatSelector));
            if (tableSelector is null)
                throw new ArgumentNullException(nameof(tableSelector));

            var results = new List<ExportResult>();
            foreach (var listObjectScope in GetListObjectAnyEnumerable())
            {
                var info = new RangeInfo(listObjectScope.ComObject);
                if (tableSelector.Invoke(info))
                {
                    var dest = destSelector(info);
                    ExportCore(listObjectScope, dest, (XlFileFormat)fileFormatSelector(info));
                    results.Add(new ExportResult(info, dest.FullName));
                }
            }
            return results;
        }

        /// <summary>
        /// エクセルファイルに含まれる全てのテーブルをファイルへエクスポートします。
        /// </summary>
        /// <param name="destSelector">テーブルごとの出力先ファイルを取得するデリゲートを指定します。</param>
        /// <param name="fileFormatSelector">テーブルごとの出力フォーマットを取得するデリゲートを指定します。</param>
        /// <returns>出力の結果を表すオブジェクト。</returns>
        public IReadOnlyList<ExportResult> ExportTableAll(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, int> fileFormatSelector)
        {
            if (destSelector is null)
                throw new ArgumentNullException(nameof(destSelector));
            if (fileFormatSelector is null)
                throw new ArgumentNullException(nameof(fileFormatSelector));

            return ExportSheetAny(destSelector, fileFormatSelector, static _ => true);
        }


        private void ExportCore(IComObjectWrappedDisposable<Worksheet> worksheetScope, FileInfo dest, XlFileFormat fileFormat)
        {
            // 新しく Excel Workbook を作成し、シートの使用領域をコピーします。
            // コピーした Workbook を指定のファイルフォーマットで保存します。
            // この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
            // 読み取り対象の Workbook (managedWorkbook) を閉じるまで出力したファイルにアクセスできなくなるためです。
            using (var usedRangeScope = worksheetScope.ComObject.UsedRange.AsDisposable())
            using (var tempWorkbookScope = _workbooksScope.ComObject.Add().AsDisposable())
            using (var tempWorksheetScope = ((Worksheet)tempWorkbookScope.ComObject.ActiveSheet).AsDisposable())
            using (var destinationRangeScope = tempWorksheetScope.ComObject.Range["A1"].AsDisposable())
            {
                usedRangeScope.ComObject.Copy();
                destinationRangeScope.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);

                dest.Directory?.Create();
                tempWorkbookScope.ComObject.SaveAs(dest.FullName, FileFormat: fileFormat);
            }
        }

        private void ExportCore(IComObjectWrappedDisposable<ListObject> listObjectScope, FileInfo dest, XlFileFormat fileFormat)
        {
            // 新しく Excel Workbook を作成し、シートの使用領域をコピーします。
            // コピーした Workbook を指定のファイルフォーマットで保存します。
            // この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
            // 読み取り対象の Workbook (managedWorkbook) を閉じるまで出力したファイルにアクセスできなくなるためです。
            using (var tableRangeScope = listObjectScope.ComObject.Range.AsDisposable())
            using (var tempWorkbookScope = _workbooksScope.ComObject.Add().AsDisposable())
            using (var tempWorksheetScope = ((Worksheet)tempWorkbookScope.ComObject.ActiveSheet).AsDisposable())
            using (var destinationRangeScope = tempWorksheetScope.ComObject.Range["A1"].AsDisposable())
            {
                tableRangeScope.ComObject.Copy();
                destinationRangeScope.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);

                dest.Directory?.Create();
                tempWorkbookScope.ComObject.SaveAs(dest.FullName, FileFormat: fileFormat);
            }
        }

        /// <summary>
        /// 指定したシート番号の <see cref="Worksheet"/> を取得します。シートが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        private IComObjectWrappedDisposable<Worksheet>? GetWorksheet(int sheetIndex)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsDisposable())
            {
                var worksheetScope = ((Worksheet)sheetsScope.ComObject[sheetIndex]).AsDisposable();
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
        private IComObjectWrappedDisposable<Worksheet>? GetWorksheet(string sheetName)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsDisposable())
            {
                var worksheetScope = ((Worksheet)sheetsScope.ComObject[sheetName]).AsDisposable();
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
        private IEnumerable<IComObjectWrappedDisposable<Worksheet>> GetWorksheetAnyEnumerable(Predicate<IComObjectWrappedDisposable<Worksheet>>? predicate)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsDisposable())
            {
                for (var i = 1; i <= sheetsScope.ComObject.Count; i++)
                {
                    using (var worksheetScope = ((Worksheet)sheetsScope.ComObject[i]).AsDisposable())
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
        private IComObjectWrappedDisposable<ListObject>? GetListObject(string tableName)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsDisposable())
            {
                for (var sheetIndex = 1; sheetIndex <= sheetsScope.ComObject.Count; sheetIndex++)
                {
                    using (var worksheetScope = ((Worksheet)sheetsScope.ComObject[sheetIndex]).AsDisposable())
                    using (var listObjectsScope = worksheetScope.ComObject.ListObjects.AsDisposable())
                    {
                        for (var listObjIndex = 1; listObjIndex <= listObjectsScope.ComObject.Count; listObjIndex++)
                        {
                            var listObjectScope = listObjectsScope.ComObject[listObjIndex].AsDisposable();
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
        private IEnumerable<IComObjectWrappedDisposable<ListObject>> GetListObjectAnyEnumerable(Predicate<IComObjectWrappedDisposable<ListObject>>? predicate)
        {
            using (var sheetsScope = _workbookScope.ComObject.Worksheets.AsDisposable())
            {
                for (var sheetIndex = 1; sheetIndex <= sheetsScope.ComObject.Count; sheetIndex++)
                {
                    using (var worksheetScope = ((Worksheet)sheetsScope.ComObject[sheetIndex]).AsDisposable())
                    using (var listObjectsScope = worksheetScope.ComObject.ListObjects.AsDisposable())
                    {
                        for (var listObjIndex = 1; listObjIndex <= listObjectsScope.ComObject.Count; listObjIndex++)
                        {
                            using (var listObjectScope = listObjectsScope.ComObject[listObjIndex].AsDisposable())
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

        private sealed class FileFullNameComparer : IEqualityComparer<FileInfo>
        {
            public FileFullNameComparer()
            {
            }

            public bool Equals(FileInfo? x, FileInfo? y)
            {
                if ((x, y) is (null, null))
                {
                    return true;
                }

                if ((x, y) is (_, null) or (null, _))
                {
                    return false;
                }

                return x.FullName == y.FullName;
            }

            public int GetHashCode([DisallowNull] FileInfo obj)
            {
                return obj.FullName.GetHashCode();
            }
        }
    }
}
