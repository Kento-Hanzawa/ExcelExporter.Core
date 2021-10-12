using System;

namespace ExcelExporterCore
{
    internal static class Util
    {
        /// <summary>
        /// 指定した列番号と行番号をエクセルのセルアドレスを表す文字列に変換します。
        /// 例えば "A1", "B10", "AA20" などです。
        /// </summary>
        /// <param name="column">列番号を表す値。1 をベースとして計算し、1 = A, 2 = B, 3 = C, ... 26 = Z, 27 = AA, 28 = AB と続きます。</param>
        /// <param name="row">行番号を表す値。</param>
        /// <returns>エクセルのセルアドレスを表す文字列。</returns>
        public static string ToExcelAddressText(int column, int row)
        {
            if (column < 1)
                throw new ArgumentOutOfRangeException(nameof(column), $"{nameof(column)} には 1 以上の値を渡す必要があります。");
            if (row < 1)
                throw new ArgumentOutOfRangeException(nameof(row), $"{nameof(row)} には 1 以上の値を渡す必要があります。");

            return $"{GetColumnText(column)}{row}";
            static string GetColumnText(int value)
            {
                if (value < 1)
                {
                    return string.Empty;
                }

                return GetColumnText((value - 1) / 26) + char.ToString((char)('A' + ((value - 1) % 26)));
            }
        }
    }
}
