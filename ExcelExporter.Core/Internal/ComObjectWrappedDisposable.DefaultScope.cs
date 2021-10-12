using System;
using System.Runtime.InteropServices;

namespace ExcelExporter.Core.Internal
{
    partial class ComObjectWrappedDisposable
    {
        private sealed class DefaultScope<T> : IComObjectWrappedDisposable<T>
            where T : class
        {
            private T comObject;
            private bool disposed = false;

            public T ComObject { get { return comObject; } }

            public DefaultScope(T comObject)
            {
                if (Marshal.IsComObject(comObject) is false)
                {
                    throw new ArgumentException($"引数 {nameof(comObject)} の型である {typeof(T)} は COM 型ではありません。", nameof(comObject));
                }

                this.comObject = comObject;
            }

            ~DefaultScope()
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
                        Marshal.FinalReleaseComObject(comObject);
                        comObject = null;
                    }
                }

                disposed = true;
            }
        }
    }
}
