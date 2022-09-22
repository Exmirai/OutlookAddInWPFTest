using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Office.Interop.Outlook;
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Forms.BaseForm;
using OutlookAddInWPFTest.Properties;
using OutlookAddInWPFTest.Utils;
using SharpVectors.Dom.Events;
using Application = System.Windows.Application;
using Exception = System.Exception;
using Point = System.Windows.Point;

namespace OutlookAddInWPFTest.Forms
{
    /// <summary>
    /// Interaction logic for JButton.xaml
    /// </summary>
    public partial class JButton : BaseWindow
    {
        public static JButton Instance { get; set; }
        private readonly Timer _jbuttonThinkTimer;
        private bool isMoving { get; set; }
        public JButton()
        {
            InitializeComponent();
            try
            {
                Uri uri = new Uri(("pack://application:,,,/OutlookAddInWPFTest;component/Resources/JButton.svg"));
                var resourceInfo = Application.GetResourceStream(uri);

                if (resourceInfo != null)
                {
                    using (var resourceStream = resourceInfo.Stream)
                    {
                        ResourceSvgCanvas.StreamSource = resourceStream;
                    }
                }

                DPIChanged += (obj, ev) =>
                {
                    var x = 1 + 1;
                };
                _jbuttonThinkTimer = new Timer(new TimerCallback(JButton_Think), null, 0, 200);
                Instance = this;
            }
            catch (Exception ex)
            {

            }
        }

        private void Window_BeforeMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                var rect = (WinAPI.RECT)Overlay.Instance.Dispatcher.Invoke(new GetClientRect(() =>
                {
                    return new WinAPI.RECT()
                    {
                        Top = (int)Overlay.Instance.Top,
                        Left = (int)Overlay.Instance.Left,
                        Bottom = (int)(Overlay.Instance.Top + Overlay.Instance.Height),
                        Right = (int)(Overlay.Instance.Left + Overlay.Instance.Width),
                    };
                }));
               // var res = WinAPI.ClipCursor(ref rect);
            }
        }
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.isMoving = true;
                this.MouseLeftButtonUp += JButton_MouseUpHandle;
                DragMove();
            }
        }

        private void JButton_LocationChanged(object sender, EventArgs e)
        {
            var overlayRect = new System.Drawing.Rectangle();
            Overlay.Instance.Dispatcher.Invoke(() =>
            {
                overlayRect = new System.Drawing.Rectangle((int)Overlay.Instance.Left, (int)Overlay.Instance.Top,
                    (int)Overlay.Instance.Width, (int)Overlay.Instance.Height);
            });
            if (this.Left < overlayRect.Left)
            {
                this.Left = overlayRect.Left;
            }else if (this.Left + this.Width > overlayRect.Right)
            {
                this.Left = overlayRect.Right - this.Width;
            }

            if (this.Top < overlayRect.Top)
            {
                this.Top = overlayRect.Top;
            }else if (this.Top + this.Height > overlayRect.Bottom)
            {
                this.Top = overlayRect.Bottom - this.Height;
            }
        }

        private void JButton_MouseUpHandle(object sender, EventArgs args)
        {
            var rect = new WinAPI.RECT();
           // if (!WinAPI.ClipCursor(ref rect))
           // {
           //     var x = new Win32Exception(Marshal.GetLastWin32Error());
          //  }
            this.MouseLeftButtonUp -= JButton_MouseUpHandle;
            var clientPos = (Point)Overlay.Instance.Dispatcher.Invoke(new ScreenToClient((pt) =>
            {
                var clPos = Overlay.Instance.PointFromScreen(pt);
                return new Point(Overlay.Instance.Width - clPos.X, Overlay.Instance.Height - clPos.Y);
            }), DispatcherPriority.Normal, new Point(this.Left, this.Top));
            if (!ValidateJButtonPosition(clientPos))
            {
                ResetJButtonPosition();
            }

            Properties.Settings.Default.JButtonPositionX = clientPos.X;
            Properties.Settings.Default.JButtonPositionY = clientPos.Y;
            isMoving = false;
        }

        private void JButton_Think(object obj)
        {
            if (Managers.StateManager.OutlookState == OutlookStateEnum.MINIMIZED)
            {
                if (this.IsVisible)
                {
                    this.Dispatcher.Invoke(() => this.Hide());
                }
                return;
            }

            if (this.isMoving)
            {
                return;
            }
            this.Dispatcher.Invoke(() =>
            {
                if (System.Windows.Input.Mouse.LeftButton == MouseButtonState.Pressed)
                {
                    return;
                }
                var buttonPosition = new Point(Properties.Settings.Default.JButtonPositionX, Properties.Settings.Default.JButtonPositionY);
                if (!ValidateJButtonPosition(buttonPosition))
                {
                    ResetJButtonPosition();
                }
                else
                {
                    SetJButtonPositionRelative(new Point(buttonPosition.X, buttonPosition.Y));
                }

                this.Show();
            });
        }

        private bool ValidateJButtonPosition(Point position)
        {
            return (bool)Overlay.Instance.Dispatcher.Invoke(new CheckPosition((pos) =>
            {
                if (pos.X < 0.0f ||
                    pos.Y < 0.0f ||
                    pos.X > Overlay.Instance.Width ||
                    pos.Y > Overlay.Instance.Height
                   )
                {
                    return false;
                }
                return true;
            }), DispatcherPriority.Normal, position);
        }

        private void ResetJButtonPosition()
        {
            this.AttachTo(Utils.OutlookUtils.GetWordWindow(),
                AttachFlagEnum.RIGHT | AttachFlagEnum.DOWN | AttachFlagEnum.INSIDE);
            var screenButtonPosition = new Point(this.Left, this.Top);
            var overlayButtonPosition = (Point)this.Dispatcher.Invoke(
                new ScreenToClient((scrPos) =>
                {
                    var clCoord = Overlay.Instance.PointFromScreen(scrPos);

                    return new Point(Overlay.Instance.Width - clCoord.X, Overlay.Instance.Height - clCoord.Y);
                }), DispatcherPriority.Normal, screenButtonPosition);
            Properties.Settings.Default.JButtonPositionX = overlayButtonPosition.X;
            Properties.Settings.Default.JButtonPositionY = overlayButtonPosition.Y;
        }

        private void SetJButtonPositionRelative(Point relativePosition)
        {
            Point nativePosition = new Point();
            Overlay.Instance.Dispatcher.Invoke(() =>
            {
                var clientPosition = new Point(Overlay.Instance.Width - relativePosition.X, Overlay.Instance.Height - relativePosition.Y);
                nativePosition = Overlay.Instance.PointToScreen(clientPosition);
            });
            this.Top = nativePosition.Y;
            this.Left = nativePosition.X;
        }
    }
}
