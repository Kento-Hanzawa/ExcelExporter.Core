using System.IO;

namespace ExcelExporter.Core
{
    internal readonly struct ExcelExportResultCore
    {
        public string RangeName { get; }
        public string RangeString { get; }
        public string DestFileName { get; }
        public FileInfo DestFileInfo { get { return new FileInfo(DestFileName); } }

        public ExcelExportResultCore(string rangeName, string rangeString, string destFileName)
        {
            this.RangeName = rangeName;
            this.RangeString = rangeString;
            this.DestFileName = destFileName;
        }
    }
}
