using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelExporterCore.Internal
{
    partial class ComObjectWrappedDisposable
    {
        private sealed class ExcelWorkbookScope : IComObjectWrappedDisposable<Workbook>
        {
            private Workbook workbook;
            private bool disposed = false;

            public Workbook ComObject { get { return workbook; } }

            public ExcelWorkbookScope(Workbook workbook)
            {
                this.workbook = workbook;
            }

            ~ExcelWorkbookScope()
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
                        Marshal.FinalReleaseComObject(workbook);
                        workbook = null;
                    }
                }

                disposed = true;
            }
        }
    }
}
