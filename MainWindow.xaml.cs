using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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

        private void btn_addTgtFileName_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btn_addSrcFileName_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                using (FileStream stream = new FileStream(dlg.FileName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new HSSFWorkbook(stream);
                    ISheet hs = workbook.GetSheet(workbook.GetSheetName(0));

                    IRow header = hs.GetRow(0);
                    List<ICell> headerCells = header.Cells;
                    List<ColumnMapping> mappingList = new List<ColumnMapping>();

                    for (int i = 0; i < headerCells.Count; i++)
                    {
                        mappingList.Add(new ColumnMapping() { SourceIndex=i, SourceName=headerCells[i].ToString() });
                    }

                    GridView gvMapping = new GridView();
                    gvMapping.Columns.Add(new GridViewColumn(){ DisplayMemberBinding=new Binding("SourceName"), Header="原始列名"});

                    ListView lvMapping=new ListView();
                    lvMapping.ItemsSource=mappingList;
                    lvMapping.View=gvMapping;

                    sp_columnMapping.Children.Add(lvMapping);
                }
            }
        }

    }

}

