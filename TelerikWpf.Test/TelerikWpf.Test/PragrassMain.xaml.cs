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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Telerik.Windows.Controls;

namespace TelerikWpf.Test
{
    /// <summary>
    /// PragrassMain.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class PragrassMain : Page
    {
        public PragrassMain()
        {
            InitializeComponent();
            this.btnProgress.Click += BtnProgress_Click;
            this.btnIndicator.Click += BtnIndicator_Click;
        }

        private void BtnIndicator_Click(object sender, RoutedEventArgs e)
        {
            //RadIndicator01 ict = new RadIndicator01("Testing...");
            //ict.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            //ict.HideMaximizeButton = true;
            //ict.HideMinimizeButton = true;
            //ict.ShowDialog();
            //Thread.Sleep(100);
            //ict.CloseThis();

            RadIndicator idc = new RadIndicator("Testing...");
            idc.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            idc.Owner = Application.Current.MainWindow;
            Thread.Sleep(100);

            if (idc.ShowDialog() == true)
            {
                idc.thisClose();
            }
        }

        private void BtnProgress_Click(object sender, RoutedEventArgs e)
        {
            ProgressSample pg = new ProgressSample();
            pg.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            pg.HideMaximizeButton = true;
            pg.HideMinimizeButton = false;
            pg.doMyWork();
            pg.ShowDialog();
           
            pg.Close();
        }
    }
}
