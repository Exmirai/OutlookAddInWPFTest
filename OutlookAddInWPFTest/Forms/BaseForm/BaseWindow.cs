using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Utils;
using Point = System.Windows.Point;

namespace OutlookAddInWPFTest.Forms.BaseForm
{
    public class BaseWindow : NativeHelpers.PerMonitorDPIWindow
    {
        public delegate Point ScreenToClient(Point src);

        public delegate bool CheckPosition(Point position);

        public delegate WinAPI.RECT GetClientRect();


        public void AttachTo(Window src, AttachFlagEnum flags)
        {
            RectangleF rect = new Rectangle();
            src.Dispatcher.Invoke(() =>
            {
                rect = new RectangleF((float)src.Left, (float)src.Top, (float)src.Width, (float)src.Height);
            });
            AttachToCoords(rect, flags);
        }
        public void AttachTo(IntPtr src, AttachFlagEnum flags)
        {
            var nativeRectangle = new WinAPI.RECT();
            if (!WinAPI.GetWindowRect(src, ref nativeRectangle))
            {
                //throw new Win32Exception(Marshal.GetLastWin32Error());
                return;
            }

            AttachToCoords(new Rectangle(nativeRectangle.Left, nativeRectangle.Top, nativeRectangle.Right - nativeRectangle.Left, nativeRectangle.Bottom - nativeRectangle.Top), flags);
        }

        private void AttachToCoords(RectangleF rect, AttachFlagEnum flags)
        {
            if ((flags & AttachFlagEnum.INSIDE) != 0)
            {
                if ((flags & AttachFlagEnum.LEFT) != 0)
                {
                    this.Left = rect.Left;
                }
                else if ((flags & AttachFlagEnum.RIGHT) != 0)
                {
                    this.Left = rect.Right - this.Width;
                }

                if (flags.HasFlag(AttachFlagEnum.UP))
                {
                    this.Top = rect.Top;

                }
                else if (flags.HasFlag(AttachFlagEnum.DOWN))
                {
                    this.Top = rect.Bottom - this.Height;
                }
            }
            else if (flags.HasFlag(AttachFlagEnum.OUTSIDE))
            {
                if (flags.HasFlag(AttachFlagEnum.LEFT))
                {
                    this.Left = rect.Left - this.Width;
                }
                else if (flags.HasFlag(AttachFlagEnum.RIGHT))
                {
                    this.Left = rect.Right;
                }

                if (flags.HasFlag(AttachFlagEnum.UP))
                {
                    this.Top = rect.Top - this.Height;
                }
                else if (flags.HasFlag(AttachFlagEnum.DOWN))
                {
                    this.Top = rect.Bottom;
                }
            }
            else if (flags.HasFlag(AttachFlagEnum.OVERLAY))
            {
                this.Left = rect.Left;
                this.Top = rect.Top;
                this.Width = rect.Right - rect.Left;
                this.Height = rect.Bottom - rect.Top;
            }
        }
    }
}
