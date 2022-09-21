using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static OutlookAddInWPFTest.Utils.WinAPI;

namespace OutlookAddInWPFTest.Utils
{
    public class DPIContextBlock : IDisposable
    {
        private DPI_AWARENESS_CONTEXT resetContext;
        private bool disposed = false;

        public DPIContextBlock(DPI_AWARENESS_CONTEXT contextSwitchTo)
        {
            resetContext = WinAPI.SetThreadDpiAwarenessContext(contextSwitchTo);

            var cccxtx = WinAPI.GetThreadDpiAwarenessContext();
            var x = WinAPI.GetAwarenessFromDpiAwarenessContext(cccxtx);
            var x2 = WinAPI.GetAwarenessFromDpiAwarenessContext(resetContext);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    WinAPI.SetThreadDpiAwarenessContext(resetContext);
                }
            }
            disposed = true;
        }
    }
}
