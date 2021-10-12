using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporter.Core.Internal
{
    partial class ComObjectWrappedDisposable
    {
        private sealed class ExcelWorkbooksScope : IComObjectWrappedDisposable<Workbooks>
        {
            private Workbooks workbooks;
            private bool disposed = false;

            public Workbooks ComObject { get { return workbooks; } }

            public ExcelWorkbooksScope(Workbooks workbooks)
            {
                this.workbooks = workbooks;
            }

            ~ExcelWorkbooksScope()
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
                        Marshal.FinalReleaseComObject(workbooks);
                        workbooks = null;
                    }
                }

                disposed = true;
            }
        }
    }
}
