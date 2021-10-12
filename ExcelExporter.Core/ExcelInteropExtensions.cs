using ExcelExporter.Core.Internal;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporter.Core
{
    internal static class ExcelInteropExtensions
    {
        /// <summary>
        /// 指定した <see cref="Worksheet"/> の使用範囲を、エクセル表現の範囲文字列に変換します。
        /// </summary>
        public static string GetRangeString(this Worksheet worksheet)
        {
            using (var rangeScope = worksheet.UsedRange.AsWrappedDisposable())
            {
                return rangeScope.ComObject.GetRangeString();
            }
        }

        /// <summary>
        /// 指定した <see cref="ListObject"/> の範囲を、エクセル表現の範囲文字列に変換します。
        /// </summary>
        public static string GetRangeString(this ListObject listObject)
        {
            using (var rangeScope = listObject.Range.AsWrappedDisposable())
            {
                return rangeScope.ComObject.GetRangeString();
            }
        }

        /// <summary>
        /// 指定した <see cref="Range"/> の範囲を、エクセル表現の範囲文字列に変換します。
        /// </summary>
        public static string GetRangeString(this Range range)
        {
            using (var columnsScope = range.Columns.AsWrappedDisposable())
            using (var rowsScope = range.Rows.AsWrappedDisposable())
            {
                var column = range.Column;
                var row = range.Row;

                var beginAddress = Util.ToExcelAddressText(column, row);
                var endAddress = Util.ToExcelAddressText(column + (columnsScope.ComObject.Count - 1), row + (rowsScope.ComObject.Count - 1));
                return $"{beginAddress}:{endAddress}";
            }
        }
    }
}
