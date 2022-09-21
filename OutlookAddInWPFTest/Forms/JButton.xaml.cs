using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
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
using OutlookAddInWPFTest.Forms.BaseForm;

namespace OutlookAddInWPFTest.Forms
{
    /// <summary>
    /// Interaction logic for JButton.xaml
    /// </summary>
    public partial class JButton : BaseWindow
    {
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
    }
}
