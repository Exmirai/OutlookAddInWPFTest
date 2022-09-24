using System;
using System.Globalization;

using Microsoft.Office.Interop.Outlook;

using OutlookAddInWPFTest.Utils;

namespace OutlookAddInWPFTest {
    public class GlobalContext {
        public static bool Initialized => _app != null;
        public static CultureInfo Language { get; internal set; }
        public static IntPtr Handle { get; set; }
        public static int ProcId { get; set; }

        private static Application _app;
        private static System.Windows.Application _appdomain;

        public static Application App => _app ?? throw new System.Exception("GlobalContext was not initialized. Please, call GlobalContext.Init method from your add-in code to initialize global context.");


        internal static void Init(Application application) {
            _app = application;
            _appdomain = new System.Windows.Application();
        }
    }
}
