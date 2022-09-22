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
        public delegate Point ScreenToClient(Point src);


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

            if ((flags & AttachFlagEnum.INSIDE) != 0)
            {
                if ((flags & AttachFlagEnum.LEFT) != 0)
                {
                    this.Left = nativeRectangle.Left;
                }
                else if ((flags & AttachFlagEnum.RIGHT) != 0)
                {
                    this.Left = nativeRectangle.Right - this.Width;
                }

                if (flags.HasFlag(AttachFlagEnum.UP))
                {
                    this.Top = nativeRectangle.Top;

                }
                else if (flags.HasFlag(AttachFlagEnum.DOWN))
                {
                    this.Top = nativeRectangle.Bottom - this.Height;
                }
            }
            else if (flags.HasFlag(AttachFlagEnum.OUTSIDE))
            {
                if (flags.HasFlag(AttachFlagEnum.LEFT))
                {
                    this.Left = nativeRectangle.Left - this.Width;
                }
                else if (flags.HasFlag(AttachFlagEnum.RIGHT))
                {
                    this.Left = nativeRectangle.Right;
                }

                if (flags.HasFlag(AttachFlagEnum.UP))
                {
                    this.Top = nativeRectangle.Top - this.Height;
                }
                else if (flags.HasFlag(AttachFlagEnum.DOWN))
                {
                    this.Top = nativeRectangle.Bottom;
                }
            }else if (flags.HasFlag(AttachFlagEnum.OVERLAY))
            {
                this.Left = nativeRectangle.Left;
                this.Top = nativeRectangle.Top;
                this.Width = nativeRectangle.Right - nativeRectangle.Left;
                this.Height = nativeRectangle.Bottom - nativeRectangle.Top;
            }
        }
    }
}
