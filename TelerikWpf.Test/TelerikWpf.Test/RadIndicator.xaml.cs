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
using System.Windows.Shapes;
using Telerik.Windows.Controls;

namespace TelerikWpf.Test
{
    /// <summary>
    /// RadIndicator.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class RadIndicator : Window
    {
        string strContent = string.Empty;
        public RadIndicator(string showText)
        {
            strContent = showText.ToString();
            InitializeComponent();
            this.Loaded += RadIndicator_Loaded;
        }

        private void RadIndicator_Loaded(object sender, RoutedEventArgs e)
        {
            showIndicator.BusyContent = strContent.ToString();
        }

        public void thisClose()
        {
            Window.GetWindow(this).DialogResult = true;
        }
    }
}
