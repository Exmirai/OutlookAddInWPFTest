using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Forms.BaseForm;
using OutlookAddInWPFTest.Managers;
using OutlookAddInWPFTest.Utils;

namespace OutlookAddInWPFTest.Forms
{
    /// <summary>
    /// Interaction logic for Overlay.xaml
    /// </summary>
    public partial class Overlay : BaseWindow
    {
        public static Overlay Instance { get; private set; }
        private readonly Timer _overlayThinkTimer;
        public Overlay()
        {
            InitializeComponent();
            UpdateAlertList();
            _overlayThinkTimer = new Timer(new TimerCallback(OverlayThink), null, 0, 200);
            Instance = this;
        }

        public void UpdateAlertList()
        {
            ClearAlertList();
            var alerts = AlertManager.GetAlerts();
            foreach (var alert in alerts)
            {
                var rect = new System.Windows.Shapes.Rectangle();
                rect.Stroke = new SolidColorBrush(Colors.Black);
                rect.StrokeThickness = 2;
                rect.Fill = new SolidColorBrush(Colors.Black);
                rect.Width = alert.rect.Width;
                rect.Height = alert.rect.Height;
                rect.Opacity = 0.5;
                rect.MouseEnter += (ev, ez) =>
                {
                    alert.ProcessHover();
                    rect.Fill = new SolidColorBrush(Colors.Red);
                };
                rect.MouseLeave += (ev, ez) =>
                {
                    alert.ProcessHover();
                    rect.Fill = new SolidColorBrush(Colors.Black);
                };
                rect.MouseLeftButtonUp += (ev, ez) =>
                {
                    alert.ProcessClick();
                    rect.Fill = new SolidColorBrush(Colors.Yellow);
                };
                var pt = new Point(alert.rect.Left, alert.rect.Top); 
                pt = (Point)this.Dispatcher.Invoke(new ScreenToClient(this.PointFromScreen), DispatcherPriority.Normal, pt);
                Canvas.SetLeft(rect, pt.X);
                Canvas.SetTop(rect, pt.Y);
                RenderList.Children.Add(rect);
            }
        }

        public void ClearAlertList()
        {
            RenderList.Dispatcher.Invoke(() => RenderList.Children.Clear());
        }


        private void OverlayThink(object ob)
        {
            if (Managers.StateManager.OutlookState == OutlookStateEnum.MINIMIZED || Managers.StateManager.UiState == UIStateEnum.DESCWND)
            {
                if (this.IsVisible)
                {
                    this.Dispatcher.Invoke(() => this.Hide());
                }
                return;
            }

            this.Dispatcher.Invoke(() => UpdateAlertList());
            this.Dispatcher.Invoke(() => this.AttachTo(Utils.OutlookUtils.GetWordWindow(), AttachFlagEnum.OVERLAY));
            this.Dispatcher.Invoke(() => this.Show());
        }

        private void Overlay_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var nativePoint = this.PointToScreen(e.GetPosition(this));
            WinAPI.ClickMouseButton(OutlookUtils.GetWordWindow(), true,new WinAPI.POINT((int)nativePoint.X, (int)nativePoint.Y));
        }

        private void Overlay_OnMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            var nativePoint = this.PointToScreen(e.GetPosition(this));
            WinAPI.ClickMouseButton(OutlookUtils.GetWordWindow(), false, new WinAPI.POINT((int)nativePoint.X, (int)nativePoint.Y));
        }
    }
}
