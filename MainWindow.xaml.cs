using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
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
        private List<ColumnMapping> mappingConfigs = new List<ColumnMapping>();
        private Dictionary<string, ColumnMapping> srcColumnConfigMapping = new Dictionary<string, ColumnMapping>();

        public MainWindow()
        {
            InitializeComponent();

            try
            {
                init();
            }
            catch (Exception ex)
            {
                MessageBox.Show("程序初始化失败: \r\n\r\n" + ex);
                return;
            }
        }

        private void init()
        {
            string[] configs = File.ReadAllLines("mapping.txt");
            for (int i = 0; i < configs.Length; i++)
            {
                string[] config = configs[i].Split(new char[] { ' ','\t' }, StringSplitOptions.RemoveEmptyEntries);
                ColumnMapping columnMapping = new ColumnMapping() { SourceFile=config[0], SourceName = config[1], TargetName = config[2] };
                mappingConfigs.Add(columnMapping);
            }
            foreach (ColumnMapping mapping in mappingConfigs)
            {
                srcColumnConfigMapping.Add(mapping.SourceName, mapping);
            }

            GridView gvMapping = new GridView();
            gvMapping.Columns.Add(new GridViewColumn() { DisplayMemberBinding = new Binding("SourceFile") { Mode = BindingMode.TwoWay }, Header = "原始文件名" });
            gvMapping.Columns.Add(new GridViewColumn() { DisplayMemberBinding = new Binding("SourceName") { Mode = BindingMode.TwoWay }, Header = "原始列名" });
            gvMapping.Columns.Add(new GridViewColumn() { DisplayMemberBinding = new Binding("TargetName") { Mode = BindingMode.TwoWay }, Header = "目标列名" });

            ListView lvMapping = new ListView();
            lvMapping.ItemsSource = mappingConfigs;
            lvMapping.View = gvMapping;

            sp_columnMapping.Children.Add(lvMapping);
        }

        private void btn_addTgtFileName_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                sp_targetFileNames.Children.Add(new TextBlock() { Text = dlg.SafeFileName });
            }
        }

        private void btn_addSrcFileName_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                sp_sourceFileNames.Children.Add(new TextBlock() { Text = dlg.SafeFileName });
                readExcelToDatabase(dlg, sourceDatabase);
            }
        }

        private Dictionary<string, Dictionary<string, object>> sourceDatabase = new Dictionary<string, Dictionary<string, object>>();

        private void readExcelToDatabase(OpenFileDialog dlg, Dictionary<string, Dictionary<string, object>> databse)
        {
            using (FileStream stream = new FileStream(dlg.FileName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new HSSFWorkbook(stream);
                ISheet hs = workbook.GetSheet(workbook.GetSheetName(0));

                List<ICell> headerCells=null;

                IEnumerator rowEnumerator = hs.GetRowEnumerator();
                while (rowEnumerator.MoveNext())
                {
                    IRow currentRow = (IRow)rowEnumerator.Current;
                    List<ICell> cells = currentRow.Cells;
                    if (headerCells==null) 
                    {
                        headerCells = cells;
                        continue;
                    }

                    Dictionary<string, object> extractedRow=null;
                    for (int i = 0; i < cells.Count; i++)
                    {
                        ICell header = headerCells[i];
                        string headerName = header.ToString();
                        if (headerName == mappingConfigs[0].SourceName) 
                        {
                            string keyName=cells[i].StringCellValue;
                            if (!databse.ContainsKey(keyName))
                            {
                                extractedRow = new Dictionary<string, object>();
                                databse.Add(cells[i].StringCellValue, extractedRow);
                            }
                            else 
                            {
                                extractedRow=databse[keyName];
                            }
                            break;
                        }
                    }

                    for (int i = 0; i < cells.Count; i++)
                    {
                        ICell header = headerCells[i];
                        string headerName = header.ToString();
                        ColumnMapping columnMappingConfig = srcColumnConfigMapping[headerName];
                        if (columnMappingConfig.SourceFile == dlg.SafeFileName)
                        {
                            extractedRow.Add(columnMappingConfig.TargetName, cells[i].StringCellValue);
                        }
                    }
                }
            }
        }
    }

}

