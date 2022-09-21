using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Utils;

namespace OutlookAddInWPFTest.Forms.BaseForm
{
    public class BaseWindow : NativeHelpers.PerMonitorDPIWindow
    {
        public delegate void AttachToD(IntPtr src, AttachFlagEnum flags);
        public void AttachTo(Window src, AttachFlagEnum flags)
        {

        }
        public void AttachTo(IntPtr src, AttachFlagEnum flags)
        {
            var nativeRectangle = new WinAPI.RECT();
            if (!WinAPI.GetWindowRect(src, ref nativeRectangle))
            {
                //throw new Win32Exception(Marshal.GetLastWin32Error());
                return;
            }

            switch (flags)
            {
                case AttachFlagEnum.OVERLAY:
                    this.Left = nativeRectangle.Left;
                    this.Top = nativeRectangle.Top;
                    this.Width = nativeRectangle.Right - nativeRectangle.Left;
                    this.Height = nativeRectangle.Bottom - nativeRectangle.Top;
                    break;
                default:
                    break;
            }
        }
    }
}
