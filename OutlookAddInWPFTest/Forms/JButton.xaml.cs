using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Forms.BaseForm;

namespace OutlookAddInWPFTest.Forms
{
    /// <summary>
    /// Interaction logic for JButton.xaml
    /// </summary>
    public partial class JButton : BaseWindow
    {
        private readonly Timer _jbuttonThinkTimer;
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
            }
            catch (Exception ex)
            {

            }
        }
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
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
            this.Dispatcher.Invoke(() => this.AttachTo(Utils.OutlookUtils.GetWordWindow(), AttachFlagEnum.RIGHT | AttachFlagEnum.DOWN | AttachFlagEnum.INSIDE));
            this.Dispatcher.Invoke(() => this.Show());
        }
    }
}
