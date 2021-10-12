using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporterCore
{
    /// <summary>
    /// エクセルワークシートの範囲に関する情報をラップするクラス。
    /// </summary>
    [SupportedOSPlatform("windows")]
    public sealed class RangeInfo
    {
        /// <summary>
        /// セル範囲の名前。例えばシート名やテーブル名などです。
        /// </summary>
        public string RangeName { get; }

        /// <summary>
        /// セル範囲を表すエクセル表現の文字列。( A1:B10, AA10:AZ10 など )
        /// </summary>
        public string RangeString { get; }

        internal RangeInfo(string rangeName, string rangeString)
        {
            RangeName = rangeName;
            RangeString = rangeString;
        }

        internal RangeInfo(Worksheet source)
        {
            RangeName = source.Name;
            RangeString = source.GetRangeString();
        }

        internal RangeInfo(ListObject source)
        {
            RangeName = source.Name;
            RangeString = source.GetRangeString();
        }
    }
}
