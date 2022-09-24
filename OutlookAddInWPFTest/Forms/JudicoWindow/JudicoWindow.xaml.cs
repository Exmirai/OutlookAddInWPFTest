using System;
using System.Collections.Generic;
using System.Linq;
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
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Forms.BaseForm;
using OutlookAddInWPFTest.Forms.JudicoWindow.LoginWnd;

namespace OutlookAddInWPFTest.Forms.JudicoWindow
{
    /// <summary>
    /// Interaction logic for JudicoWindow.xaml
    /// </summary>
    public partial class JudicoWindow : BaseWindow
    {
        private readonly LoginContent _loginContent;
        public static JudicoWindow Instance { get; private set; }
        public JudicoWindow()
        {
            InitializeComponent();
            Instance = this;
            _loginContent = new LoginContent();
            ///bla bla bla
            this.SwitchToLoginWnd();
            ///
        }

        public void ToggleWindow()
        {
            if (this.IsVisible)
            {
                this.HideWindow();
            }
            else
            {
                this.ShowWindow();
            }
        }

        public void ShowWindow()
        {
            Overlay.Instance.Topmost = false;
            this.AttachTo(JButton.Instance, AttachFlagEnum.OUTSIDE | AttachFlagEnum.LEFT | AttachFlagEnum.UP);
        this.Show();
        }
        public void HideWindow()
        {
            this.Hide();
            Overlay.Instance.Topmost = true;
        }
        public void SwitchToLoginWnd()
        {
            this.ContentGrid.Children.Clear();
            this.ContentGrid.Children.Add(_loginContent);
        }

        public void SwitchToUserMenu()
        {

        }
    }
}
