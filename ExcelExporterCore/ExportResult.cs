using System.IO;

namespace ExcelExporterCore
{
    /// <summary>
    /// エクスポート操作の結果を表すパラメータをラップしたクラス。
    /// </summary>
    public sealed class ExportResult
    {
        /// <summary>
        /// 出力を行った範囲に関する情報をラップしたクラスのインスタンス。
        /// </summary>
        public RangeInfo RangeInfo { get; }

        /// <summary>
        /// 出力先ファイルの完全パスです。
        /// </summary>
        public string DestFileName { get; }

        /// <summary>
        /// <see cref="DestFileName"/> から作成される <see cref="FileInfo"/> インスタンスを取得します。
        /// </summary>
        public FileInfo DestFileInfo { get { return new FileInfo(DestFileName); } }

        internal ExportResult(RangeInfo rangeInfo, string destFileName)
        {
            this.RangeInfo = rangeInfo;
            this.DestFileName = destFileName;
        }
    }
}
