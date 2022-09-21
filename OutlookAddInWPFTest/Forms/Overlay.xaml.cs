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
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Forms.BaseForm;
using OutlookAddInWPFTest.Utils;

namespace OutlookAddInWPFTest.Forms
{
    /// <summary>
    /// Interaction logic for Overlay.xaml
    /// </summary>
    public partial class Overlay : BaseWindow
    {
        private readonly Timer _overlayThinkTimer;
        public Overlay()
        {
            InitializeComponent();
            UpdateAlertList();
            _overlayThinkTimer = new Timer(new TimerCallback(OverlayThink), null, 0, 200);
        }

        public void UpdateAlertList()
        {
            var rect = new System.Windows.Shapes.Rectangle();
            rect.Stroke = new SolidColorBrush(Colors.Black);
            rect.StrokeThickness = 2;
            rect.Fill = new SolidColorBrush(Colors.Black);
            rect.Width = 60;
            rect.Height = 30;
            rect.Opacity = 0.5f;
            rect.MouseEnter += (ev, ez) =>
            {
                var x = 1 + 1;
            };
            Canvas.SetLeft(rect, 0);
            Canvas.SetTop(rect, 0);
            RenderList.Children.Add(rect);
        }

        public void ClearAlertList()
        {
            RenderList.Children.Clear();
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
            this.Dispatcher.Invoke(() => this.AttachTo(Utils.OutlookUtils.GetWordWindow(), AttachFlagEnum.OVERLAY));
            this.Dispatcher.Invoke(() => this.Show());
        }

        private void Overlay_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var nativePoint = this.PointToScreen(e.GetPosition(this));
            WinAPI.ClickMouseButton(OutlookUtils.GetOutlookWindow(), true,new WinAPI.POINT((int)nativePoint.X, (int)nativePoint.Y));
        }

        private void Overlay_OnMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            var nativePoint = this.PointToScreen(e.GetPosition(this));
            WinAPI.ClickMouseButton(OutlookUtils.GetOutlookWindow(), false, new WinAPI.POINT((int)nativePoint.X, (int)nativePoint.Y));
        }
    }
}
