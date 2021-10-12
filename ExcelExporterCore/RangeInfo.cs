using Microsoft.Office.Interop.Excel;
using ExcelExporterCore.Internal;

namespace ExcelExporterCore
{
    internal sealed class RangeInfo
    {
        public string RangeName { get; }
        public string RangeString { get; }

        internal RangeInfo(IComObjectWrappedDisposable<Worksheet> source)
        {
            RangeName = source.ComObject.Name;
            RangeString = source.ComObject.GetRangeString();
        }

        internal RangeInfo(IComObjectWrappedDisposable<ListObject> source)
        {
            RangeName = source.ComObject.Name;
            RangeString = source.ComObject.GetRangeString();
        }
    }
}
