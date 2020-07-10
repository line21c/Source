using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using Telerik.Windows.Controls;
using Telerik.Windows.Controls.Navigation;

namespace TelerikWpf.Test
{
    /// <summary>
    /// Interaction logic for RadIndicator01.xaml
    /// </summary>
    public partial class RadIndicator01
    {
        string strContent = string.Empty;
        public RadIndicator01(string showText)
        {
            strContent = showText.ToString();
            InitializeComponent();

            this.Loaded += RadIndicator01_Loaded;
        }

        private void RadIndicator01_Loaded(object sender, RoutedEventArgs e)
        {
            this.Header = strContent.ToString();
            
            showIndicator.BusyContent = strContent.ToString();
            RadWindowInteropHelper.SetShowInTaskbar(this, false);
            showIndicator.Width = 290;
            //Height = 0;
            //Width = 0;
            //Visibility = Visibility.Collapsed;
        }

        public void CloseThis()
        {
            this.Close();
        }
    }
}
