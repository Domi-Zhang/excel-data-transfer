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
using FolderBrowserDialog = System.Windows.Forms.FolderBrowserDialog;

namespace excel_data_transfer
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private ColumnMapping keyMapping;
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
                string[] config = configs[i].Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (keyMapping == null)
                {
                    keyMapping = new ColumnMapping() { SourceFile = config[0], SourceName = config[1], TargetName = config[2] };
                }
                else
                {
                    ColumnMapping columnMapping = new ColumnMapping() { SourceFile = config[0], SourceName = config[1], TargetName = config[2] };
                    mappingConfigs.Add(columnMapping);
                }
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

        private Dictionary<string, string> tgtFileDict = new Dictionary<string, string>();
        private void btn_addTgtFileName_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                sp_targetFileNames.Children.Add(new TextBlock() { Text = dlg.SafeFileName });
                tgtFileDict.Add(dlg.SafeFileName, dlg.FileName);
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
            stepThroughExcel(dlg.FileName, (Dictionary<string, object> extractedRow, string rowKey, List<ICell> headerCells, List<ICell> cells) =>
            {
                if (extractedRow == null)
                {
                    sourceDatabase.Add(rowKey, (extractedRow = new Dictionary<string, object>()));
                }
                extractDataToRow(dlg.SafeFileName, headerCells, cells, extractedRow);
            },
            null);
        }

        private void stepThroughExcel(string fileName, Action<Dictionary<string, object>, string, List<ICell>, List<ICell>> action, Action<IWorkbook> finish)
        {
            using (FileStream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new HSSFWorkbook(stream);
                ISheet hs = workbook.GetSheet(workbook.GetSheetName(0));
                List<ICell> headerCells = hs.GetRow(0).Cells;
                int keyColumn = getKeyColumnIndex(headerCells, keyMapping.SourceFile);

                IEnumerator rowEnumerator = hs.GetRowEnumerator();
                rowEnumerator.MoveNext();//跳过列首
                while (rowEnumerator.MoveNext())
                {
                    List<ICell> cells = ((IRow)rowEnumerator.Current).Cells;
                    string rowKey = cells[keyColumn].StringCellValue;
                    Dictionary<string, object> extractedRow = sourceDatabase[rowKey];
                    action(extractedRow, rowKey, headerCells, cells);
                }

                if (finish != null) 
                {
                    finish(workbook);
                }
            }
        }

        private void extractDataToRow(string fileName, List<ICell> headerCells, List<ICell> cells, Dictionary<string, object> extractedRow)
        {
            for (int i = 0; i < cells.Count; i++)
            {
                ICell header = headerCells[i];
                string headerName = header.ToString();
                ColumnMapping columnMappingConfig = srcColumnConfigMapping[headerName];
                if (columnMappingConfig.SourceFile == fileName)
                {
                    extractedRow.Add(columnMappingConfig.TargetName, cells[i].StringCellValue);
                }
            }
        }

        private void btn_transfer_Click(object sender, RoutedEventArgs e)
        {
            foreach (KeyValuePair<string,string> tgtFile in tgtFileDict)
            {
                stepThroughExcel(tgtFile.Value, (Dictionary<string, object> extractedRow, string rowKey, List<ICell> headerCells, List<ICell> cells) =>
                {
                    for (int i = 0; i < cells.Count; i++)
                    {
                        object srcValue = extractedRow[headerCells[i].StringCellValue];
                        if (srcValue != null)
                        {
                            cells[i].SetCellValue(srcValue.ToString());
                        }
                    }
                }, (IWorkbook workbook) => 
                {
                    FileStream writeStream = new FileStream(targetFolder+"/"+tgtFile.Key, FileMode.OpenOrCreate, FileAccess.Write);
                    workbook.Write(writeStream);
                    writeStream.Close();
                });
            }
        }

        private int getKeyColumnIndex(List<ICell> headerCells, string keyColumnName)
        {
            int keyColumn = -1;
            for (int i = 0; i < headerCells.Count; i++)
            {
                string columnName = headerCells[i].StringCellValue;
                if (columnName == keyColumnName)
                {
                    keyColumn = i;
                    break;
                }
            }
            return keyColumn;
        }

        private string targetFolder;
        private void btn_addTgtFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                targetFolder = fbd.SelectedPath;
                txt_targetFolder.Text = targetFolder;
            }
        }
    }

}

