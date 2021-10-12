using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporter.Core.Internal
{
    partial class ComObjectWrappedDisposable
    {
        private sealed class ExcelApplicationScope : IComObjectWrappedDisposable<Application>
        {
            private Application application;
            private bool disposed = false;

            public Application ComObject { get { return application; } }

            public ExcelApplicationScope(Application application)
            {
                this.application = application;
            }

            ~ExcelApplicationScope()
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
                        Marshal.FinalReleaseComObject(application);
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
    }
}
