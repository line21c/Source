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
using System.ComponentModel;
using System.Threading;

using Telerik.Windows.Controls;

namespace TelerikWpf.Test
{
    //https://www.technical-recipes.com/2016/running-a-wpf-progress-bar-as-a-background-worker-thread
    public partial class ProgressSample
    {
        System.ComponentModel.BackgroundWorker MyWorker = new System.ComponentModel.BackgroundWorker();
        private delegate void UpdateMyDelegatedelegate(int i);

        public ProgressSample()
        {
            InitializeComponent();
            MyWorker.DoWork += MyWorker_DoWork;
            MyWorker.RunWorkerCompleted += MyWorker_RunWorkerCompleted;
        }

        private void MyWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                //MessageBox.Show("Complete My Background Work", "Complate");
                this.Close();
            }
            else
            {
                // MessageBox.Show("Complete My Background Work", "Complate");
            }
        }

        private void MyWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int i = 0; i <= 100; i++)
            {
                Thread.Sleep(10);
                if (MyWorker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                UpdateMyDelegatedelegate UpdateMyDelegate = new UpdateMyDelegatedelegate(UpdateMyDelegateLabel);
                pbStatus.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, UpdateMyDelegate, i);
            }
        }

        private void UpdateMyDelegateLabel(int i)
        {
            pbStatus.Value = i;
        }


        public void doMyWork()
        {
            MyWorker.RunWorkerAsync();
        }

    }
}
