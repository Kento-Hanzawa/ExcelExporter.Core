using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporter.Core.Internal
{
    /// <summary>
    /// <see cref="IComObjectWrappedDisposable{T}"/> を作成するためのユーティリティを提供します。
    /// </summary>
    internal static partial class ComObjectWrappedDisposable
    {
        /// <summary>
        /// 新しい <see cref="_Application"/> オブジェクトを作成します。
        /// </summary>
        /// <returns>作成された <see cref="_Application"/> オブジェクトをラップする <see cref="IComObjectWrappedDisposable{T}"/> のインスタンス。</returns>
        public static IComObjectWrappedDisposable<Application> CreateApplication()
        {
            var scope = new Application().AsWrappedDisposable();
            // 以下のパラメータは、バッチ処理の都合上一時的に false に変更しておくものです。
            // ExcelApplicationScope の Dispose 内で true に設定し直す必要があります。
            scope.ComObject.DisplayAlerts = false;
            scope.ComObject.ScreenUpdating = false;
            scope.ComObject.AskToUpdateLinks = false;
            return scope;
        }

        /// <summary>
        /// 対象の COM オブジェクトをラップした <see cref="IComObjectWrappedDisposable{T}"/> のインスタンスを取得します。
        /// </summary>
        /// <exception cref="ArgumentException"><paramref name="comObject"/> が COM 型のオブジェクトではありません。</exception>
        /// <param name="comObject">ラップする COM オブジェクト。</param>
        public static IComObjectWrappedDisposable<T> AsWrappedDisposable<T>(this T comObject)
            where T : class
        {
            if (Marshal.IsComObject(comObject) is false)
            {
                throw new ArgumentException($"引数 {nameof(comObject)} の型である {typeof(T)} は COM 型ではありません。", nameof(comObject));
            }

            return new DefaultScope<T>(comObject);
        }

        /// <summary>
        /// 対象の COM オブジェクトをラップした <see cref="IComObjectWrappedDisposable{T}"/> のインスタンスを取得します。
        /// </summary>
        /// <param name="comObject">ラップする COM オブジェクト。</param>
        public static IComObjectWrappedDisposable<Application> AsWrappedDisposable(this Application comObject)
        {
            return new ExcelApplicationScope(comObject);
        }

        /// <summary>
        /// 対象の COM オブジェクトをラップした <see cref="IComObjectWrappedDisposable{T}"/> のインスタンスを取得します。
        /// </summary>
        /// <param name="comObject">ラップする COM オブジェクト。</param>
        public static IComObjectWrappedDisposable<Workbooks> AsWrappedDisposable(this Workbooks comObject)
        {
            return new ExcelWorkbooksScope(comObject);
        }

        /// <summary>
        /// 対象の COM オブジェクトをラップした <see cref="IComObjectWrappedDisposable{T}"/> のインスタンスを取得します。
        /// </summary>
        /// <param name="comObject">ラップする COM オブジェクト。</param>
        public static IComObjectWrappedDisposable<Workbook> AsWrappedDisposable(this Workbook comObject)
        {
            return new ExcelWorkbookScope(comObject);
        }
    }
}
