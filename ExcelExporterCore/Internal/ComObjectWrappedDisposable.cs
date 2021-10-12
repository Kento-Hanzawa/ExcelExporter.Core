using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporterCore.Internal
{
    /// <summary>
    /// <see cref="IComObjectWrappedDisposable{T}"/> を作成するためのユーティリティを提供します。
    /// </summary>
    [SupportedOSPlatform("windows")]
    internal static partial class ComObjectWrappedDisposable
    {
        /// <summary>
        /// 新しい <see cref="_Application"/> オブジェクトを作成します。
        /// </summary>
        /// <returns>作成された <see cref="_Application"/> オブジェクトをラップする <see cref="IComObjectWrappedDisposable{T}"/> のインスタンス。</returns>
        public static IComObjectWrappedDisposable<Application> CreateApplication()
        {
            var scope = new Application().AsDisposable();
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
        public static IComObjectWrappedDisposable<T> AsDisposable<T>(this T comObject)
            where T : class
        {
            if (Marshal.IsComObject(comObject) is false)
            {
                throw new ArgumentException($"引数 {nameof(comObject)} の型である {typeof(T)} は COM 型ではありません。", nameof(comObject));
            }

            return new DefaultWrappedDisposable<T>(comObject);
        }

        /// <summary>
        /// 対象の COM オブジェクトをラップした <see cref="IComObjectWrappedDisposable{T}"/> のインスタンスを取得します。
        /// </summary>
        /// <param name="comObject">ラップする COM オブジェクト。</param>
        public static IComObjectWrappedDisposable<Application> AsDisposable(this Application comObject)
        {
            return new ApplicationWrappedDisposable(comObject);
        }

        /// <summary>
        /// 対象の COM オブジェクトをラップした <see cref="IComObjectWrappedDisposable{T}"/> のインスタンスを取得します。
        /// </summary>
        /// <param name="comObject">ラップする COM オブジェクト。</param>
        public static IComObjectWrappedDisposable<Workbooks> AsDisposable(this Workbooks comObject)
        {
            return new WorkbooksWrappedDisposable(comObject);
        }

        /// <summary>
        /// 対象の COM オブジェクトをラップした <see cref="IComObjectWrappedDisposable{T}"/> のインスタンスを取得します。
        /// </summary>
        /// <param name="comObject">ラップする COM オブジェクト。</param>
        public static IComObjectWrappedDisposable<Workbook> AsDisposable(this Workbook comObject)
        {
            return new WorkbookWrappedDisposable(comObject);
        }



        private sealed class DefaultWrappedDisposable<T> : IComObjectWrappedDisposable<T>
            where T : class
        {
            private T comObject;
            private bool disposed = false;

            public T ComObject { get { return comObject; } }

            public DefaultWrappedDisposable(T comObject)
            {
                if (Marshal.IsComObject(comObject) is false)
                {
                    throw new ArgumentException($"引数 {nameof(comObject)} の型である {typeof(T)} は COM 型ではありません。", nameof(comObject));
                }

                this.comObject = comObject;
            }

            ~DefaultWrappedDisposable()
            {
                Dispose(false);
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            private void Dispose(bool disposing)
            {
                if (disposed)
                {
                    return;
                }

                if (disposing)
                {
                    if (comObject != null)
                    {
                        _ = Marshal.FinalReleaseComObject(comObject);
                        comObject = null;
                    }
                }

                disposed = true;
            }
        }

        private sealed class ApplicationWrappedDisposable : IComObjectWrappedDisposable<Application>
        {
            private Application application;
            private bool disposed = false;

            public Application ComObject { get { return application; } }

            public ApplicationWrappedDisposable(Application application)
            {
                this.application = application;
            }

            ~ApplicationWrappedDisposable()
            {
                Dispose(false);
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            private void Dispose(bool disposing)
            {
                if (disposed)
                {
                    return;
                }

                if (disposing)
                {
                    // Microsoft.Office.Interop.Excel.Application の COM オブジェクト解放手順は、以下のサイトを参考にしています。
                    // https://blogs.msdn.microsoft.com/office_client_development_support_blog/2012/02/09/office-5/

                    // アプリケーションの終了前に GC を強制します。
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();

                    if (application != null)
                    {
                        // CreateApplication 内で false にした項目を true に変更します。
                        application.DisplayAlerts = true;
                        application.ScreenUpdating = true;
                        application.AskToUpdateLinks = true;

                        application.Quit();
                        _ = Marshal.FinalReleaseComObject(application);
                        application = null;

                        // アプリケーションの終了後に GC を強制します。
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                    }
                }

                disposed = true;
            }
        }

        private sealed class WorkbookWrappedDisposable : IComObjectWrappedDisposable<Workbook>
        {
            private Workbook workbook;
            private bool disposed = false;

            public Workbook ComObject { get { return workbook; } }

            public WorkbookWrappedDisposable(Workbook workbook)
            {
                this.workbook = workbook;
            }

            ~WorkbookWrappedDisposable()
            {
                Dispose(false);
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            private void Dispose(bool disposing)
            {
                if (disposed)
                {
                    return;
                }

                if (disposing)
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        _ = Marshal.FinalReleaseComObject(workbook);
                        workbook = null;
                    }
                }

                disposed = true;
            }
        }

        private sealed class WorkbooksWrappedDisposable : IComObjectWrappedDisposable<Workbooks>
        {
            private Workbooks workbooks;
            private bool disposed = false;

            public Workbooks ComObject { get { return workbooks; } }

            public WorkbooksWrappedDisposable(Workbooks workbooks)
            {
                this.workbooks = workbooks;
            }

            ~WorkbooksWrappedDisposable()
            {
                Dispose(false);
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            private void Dispose(bool disposing)
            {
                if (disposed)
                {
                    return;
                }

                if (disposing)
                {
                    if (workbooks != null)
                    {
                        workbooks.Close();
                        _ = Marshal.FinalReleaseComObject(workbooks);
                        workbooks = null;
                    }
                }

                disposed = true;
            }
        }
    }
}
