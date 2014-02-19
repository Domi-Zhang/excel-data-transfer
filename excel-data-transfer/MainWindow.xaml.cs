using Microsoft.Win32;
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

namespace excel_data_transfer
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn_fileName_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                using (FileStream stream = newFileStream(@"C:\Users\Administrator\Desktop\Book1.xls",FileMode.Open, FileAccess.Read))

            {

                //创建一个工作薄

                HSSFWorkbook workbook = newHSSFWorkbook(stream);

                //创建表的实例指向文件流工作薄的第一个表

                HSSFSheet hs =workbook.GetSheet(workbook.GetSheetName(0));

                for (int i = 0; i < 5; i++)

                {

                    for (int j = 0; j < 2;j++)

                    {

                        创建行的实例

                        HSSFRow hr =hs.GetRow(i);

                        //创建列的实例

                        HSSFCell hc =hr.GetCell(j);

                        Console.Write(hc.ToString()+"   ");

                    }

                    Console.WriteLine();

                }

            }

        }
            }
        }
    }
}
