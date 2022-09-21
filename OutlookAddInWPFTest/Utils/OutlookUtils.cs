using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutlookAddInWPFTest.Enum;

namespace OutlookAddInWPFTest.Utils
{
    public static class OutlookUtils
    {
        public static IntPtr GetOutlookWindow()
        {
            var p1 =  WinAPI.FindChildWindowByClassName(IntPtr.Zero, "rctrl_renwnd32");
            return p1;
        }

        public static IntPtr GetWordWindow()
        {
            return WinAPI.FindChildWindowByClassName(GetOutlookWindow(), "_WwG");
        }
    }
}
