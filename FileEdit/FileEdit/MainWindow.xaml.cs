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
using System.Text.RegularExpressions;
using Xceed.Words.NET;

namespace FileEdit
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // textbox.Text = "a" + "\r\n" + "b";
          
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            using (DocX doc = DocX.Load(@"..\..\File\TempDoc.docx"))
            {
                var regularBookmark1 = doc.Bookmarks["Name"];
                var regularBookmark2 = doc.Bookmarks["Address"];
                var regularBookmark3 = doc.Bookmarks["Postcode"];
                var regularBookmark4 = doc.Bookmarks["OBNumber"];
                var regularBookmark5 = doc.Bookmarks["DP1"];
                var regularBookmark6 = doc.Bookmarks["DP2"];
                var regularBookmark7 = doc.Bookmarks["DP3"];

           
             


                if (regularBookmark1 != null)
                {
                    var Name = textName.Text;
                    regularBookmark1.SetText(Name);
                }



                if (regularBookmark2 != null)
                {
                    var Address = textAddress.Text;
                    regularBookmark2.SetText(Address);
                }



                if (regularBookmark3 != null)
                {
                    var Postcode = textPostcode.Text;
                    regularBookmark3.SetText(Postcode);
                }


                if (regularBookmark4 != null)
                {
                    var OBNumber = textOBNo.Text;
                    regularBookmark4.SetText(OBNumber);
                }


                if (regularBookmark5 != null)
                {
                    regularBookmark5.SetText(Convert.ToDateTime(DP1.Text).ToString("dd/MM/yyyy"));
                }

                if (regularBookmark6 != null)
                {
                    regularBookmark6.SetText(Convert.ToDateTime(DP2.Text).ToString("dd/MM/yyyy"));
                }


                if (regularBookmark7 != null)
                {
                    regularBookmark7.SetText(Convert.ToDateTime(DP3.Text).ToString("dd/MM/yyyy"));
                }

                doc.Save();

            
            }
        }

       
    }
}
