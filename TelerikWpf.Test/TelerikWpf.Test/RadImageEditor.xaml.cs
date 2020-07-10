using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Shapes;
using Telerik.Windows.Controls;
using System.Drawing;
using Telerik.Windows.Media.Imaging.FormatProviders;


namespace TelerikWpf.Test
{
    /// <summary>
    /// RadImageEditor.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class RadImageEditor : Page
    {
        #region 01. 화면 전역 변수.함수
        private string SampleImageFolder = @"C:\Users\samsung\Pictures\";
        #endregion
        public RadImageEditor()
        {
            this.Tag = "8129345";
            InitializeComponent();
            this.Loaded += JfrmProcImgView_Loaded;
            this.btnShow.Click += BtnShow_Click;

            //ImageEditorUI.
        }

        private void JfrmProcImgView_Loaded(object sender, RoutedEventArgs e)
        {
            string strTag = this.Tag.ToString();
            if (string.IsNullOrEmpty(strTag)) return;


            try
            {
                this.LoadSampleImage(strTag);
            }
            catch (Exception ex) { }
        }

        #region 03. 기타 이벤트
        private void BtnShow_Click(object sender, RoutedEventArgs e)
        {
            string strName = "8125219_C.JPG";
            LoadSampleImage(this.ImageEditorUI, strName);
        }

        #endregion

        #region 04. 사용자 정의 함수
        private void LoadSampleImage(string sImgDwgPath)
        {
            if (System.IO.File.Exists(SampleImageFolder + sImgDwgPath + ".JPG"))
            {
                Stream stream1 = File.OpenRead(SampleImageFolder + sImgDwgPath + ".JPG");
                string extension = System.IO.Path.GetExtension(SampleImageFolder + sImgDwgPath + ".JPG").ToLower();

                IImageFormatProvider formatProvider = ImageFormatProviderManager.GetFormatProviderByExtension(extension);
                if (formatProvider != null)
                {
                    ImageEditorUI.Image = formatProvider.Import(stream1);
                    ImageEditorUI.ApplyTemplate();
                    ImageEditorUI.ImageEditor.ScaleFactor = 0;
                }
            }
            else
            {
                MessageBox.Show("미등록 작업표준서입니다.", "확인");
            }
        }

        #region  Resource
        //Error
        private void LoadSampleImage(RadImageEditorUI imageEditorUI, string image)
        {
            //string strPath = server
            //using (Stream stream = Application.GetResourceStream(GetResourceUri(image)).Stream)
            //{
            //    imageEditorUI.Image = new Telerik.Windows.Media.Imaging.RadBitmap(stream);
            //    imageEditorUI.ApplyTemplate();
            //    imageEditorUI.ImageEditor.ScaleFactor = 0;
            //}
        }

        public static Uri GetResourceUri(string resource)
        {
            AssemblyName assemblyName = new AssemblyName(typeof(MainWindow).Assembly.FullName);
            string resourcePath = "/" + assemblyName.Name + ";component/" + resource;
            Uri resourceUri = new Uri(resourcePath, UriKind.Relative);

            return resourceUri;
        }
        #endregion

        #endregion
    }
}
