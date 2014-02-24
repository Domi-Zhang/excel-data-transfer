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
        private Dictionary<string, ColumnMapping> srcColumnConfigMapping = new Dictionary<string, ColumnMapping>();
        private Dictionary<string, ExcelConfig> excelConfigDict = new Dictionary<string, ExcelConfig>();

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
            string[] mappingConfigs = File.ReadAllLines("column-mapping.txt");
            for (int i = 0; i < mappingConfigs.Length; i++)
            {
                string[] configInfo = mappingConfigs[i].Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (keyMapping == null)
                {
                    keyMapping = new ColumnMapping() { SourceFile = configInfo[0], SourceNames = configInfo[1].Split('|'), TargetNames = configInfo[2].Split('|') };
                }
                else
                {
                    ColumnMapping columnMapping = new ColumnMapping() { SourceFile = configInfo[0], SourceNames = configInfo[1].Split('|'), TargetNames = configInfo[2].Split('|') };
                    foreach (string srcColumn in columnMapping.SourceNames)
                    {
                        srcColumnConfigMapping.Add(srcColumn, columnMapping);
                    } 
                }
            }

            string[] excelConfigs = File.ReadAllLines("excel-config.txt");
            for (int i = 0; i < excelConfigs.Length; i++)
            {
                string[] configInfo = excelConfigs[i].Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                ExcelConfig tgtConfig = new ExcelConfig() { FileName=configInfo[0], HeaderRow=int.Parse(configInfo[1]), SheetIndex=int.Parse(configInfo[2]) };
                excelConfigDict.Add(tgtConfig.FileName, tgtConfig);
            }

            lv_columnMapping.ItemsSource = srcColumnConfigMapping.Values;
            lv_excelConfig.ItemsSource = excelConfigDict.Values;
        }

        private Dictionary<string, string> tgtFileDict = new Dictionary<string, string>();
        private void btn_addTgtFileName_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            dlg.Multiselect = true;

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                for (int i = 0; i < dlg.SafeFileNames.Length; i++)
                {
                    sp_targetFileNames.Children.Add(new TextBlock() { Text = dlg.SafeFileNames[i] });
                    tgtFileDict.Add(dlg.SafeFileNames[i], dlg.FileNames[i]);
                }
            }
        }

        private void btn_addSrcFileName_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            dlg.Multiselect = true;

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                for (int i = 0; i < dlg.SafeFileNames.Length; i++)
                {
                    sp_sourceFileNames.Children.Add(new TextBlock() { Text = dlg.SafeFileNames[i] });
                    readExcelToDatabase(dlg.FileNames[i], dlg.SafeFileNames[i], sourceDatabase);
                }
            }
        }

        private Dictionary<string, Dictionary<string, ICell>> sourceDatabase = new Dictionary<string, Dictionary<string, ICell>>();

        private void readExcelToDatabase(string fileName, string safeFileName, Dictionary<string, Dictionary<string, ICell>> databse)
        {
            stepThroughExcel(fileName, safeFileName, keyMapping.SourceNames, true, (string rowKey, List<ICell> headerCells, List<ICell> cells) =>
            {
                Dictionary<string, ICell> extractedRow;
                if (!sourceDatabase.ContainsKey(rowKey))
                {
                    sourceDatabase.Add(rowKey, (extractedRow = new Dictionary<string, ICell>()));
                }
                else 
                {
                    extractedRow = sourceDatabase[rowKey];
                }
                extractDataToRow(safeFileName, headerCells, cells, extractedRow);
            },
            null);
        }

        private void stepThroughExcel(string fullFileName, string breifFileName, string[] keyColumnNames, bool skipTitle, Action<string, List<ICell>, List<ICell>> action, Action<IWorkbook> finish)
        {
            using (FileStream stream = new FileStream(fullFileName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                if (fullFileName.EndsWith(".xlsx"))
                {
                    workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(stream);
                }
                else 
                {
                    workbook = new HSSFWorkbook(stream);
                }

                int sheetIndex = 1;
                if (excelConfigDict.ContainsKey(breifFileName)) 
                {
                    sheetIndex = excelConfigDict[breifFileName].SheetIndex;
                }
                ISheet hs = workbook.GetSheet(workbook.GetSheetName(sheetIndex - 1));

                int headerRow = 0;
                if(skipTitle&&excelConfigDict.ContainsKey(breifFileName))
                {
                    headerRow=excelConfigDict[breifFileName].HeaderRow-1;
                }
                List<ICell> headerCells = hs.GetRow(headerRow).Cells;
                int keyColumn = getKeyColumnIndex(headerCells, keyColumnNames);

                IEnumerator rowEnumerator = hs.GetRowEnumerator();
                rowEnumerator.MoveNext();//跳过列首
                while (rowEnumerator.MoveNext())
                {
                    List<ICell> cells = ((IRow)rowEnumerator.Current).Cells;
                    string rowKey = cells[keyColumn].StringCellValue;
                    if (string.IsNullOrEmpty(rowKey)) 
                    {
                        continue;
                    }
                    action(rowKey, headerCells, cells);
                }

                if (finish != null) 
                {
                    finish(workbook);
                }
            }
        }

        private void extractDataToRow(string fileName, List<ICell> headerCells, List<ICell> cells, Dictionary<string, ICell> extractedRow)
        {
            for (int i = 0; i < cells.Count; i++)
            {
                ICell header = headerCells[i];
                string headerName = header.ToString();
                if (srcColumnConfigMapping.ContainsKey(headerName))
                {
                    ColumnMapping columnMappingConfig = srcColumnConfigMapping[headerName];
                    if (columnMappingConfig.SourceFile == fileName)
                    {
                        foreach (var name in columnMappingConfig.TargetNames)
                        {
                            extractedRow.Add(name, cells[i]);
                        }
                    }
                }
            }
        }

        private void btn_transfer_Click(object sender, RoutedEventArgs e)
        {
            btn_transfer.Content = "处理中...";
            int handleProgress = 0;
            foreach (KeyValuePair<string,string> tgtFile in tgtFileDict)
            {
                stepThroughExcel(tgtFile.Value, tgtFile.Key, keyMapping.TargetNames, true, (string rowKey, List<ICell> headerCells, List<ICell> cells) =>
                {
                    if (sourceDatabase.ContainsKey(rowKey))
                    {
                        Dictionary<string, ICell> extractedRow = sourceDatabase[rowKey];
                        for (int i = 0; i < cells.Count && i < headerCells.Count; i++)
                        {
                            if (extractedRow.ContainsKey(headerCells[i].StringCellValue))
                            {
                                ICell srcValue = extractedRow[headerCells[i].StringCellValue];
                                switch (srcValue.CellType)
                                {
                                    case CellType.Boolean:
                                        cells[i].SetCellValue(srcValue.BooleanCellValue);
                                        break;
                                    case CellType.Numeric:
                                        cells[i].SetCellValue(srcValue.NumericCellValue);
                                        break;
                                    case CellType.String:
                                        cells[i].SetCellValue(srcValue.StringCellValue);
                                        break;
                                    case CellType.Blank:
                                        cells[i].SetCellValue(srcValue.StringCellValue);
                                        break;
                                    default:
                                        cells[i].SetCellValue(srcValue.StringCellValue);
                                        break;
                                }
                            }
                        }
                    }
                }, (IWorkbook workbook) =>
                {
                    FileStream writeStream = new FileStream(targetFolder + "/" + tgtFile.Key, FileMode.OpenOrCreate, FileAccess.Write);
                    workbook.Write(writeStream);
                    writeStream.Close();
                });
                txt_handleProgress.Text = (handleProgress++) + "/" + tgtFileDict.Count;
            }
            btn_transfer.Content = "处理完成";
        }

        private int getKeyColumnIndex(List<ICell> headerCells, string[] keyColumnNames)
        {
            for (int i = 0; i < headerCells.Count; i++)
            {
                string columnName = headerCells[i].StringCellValue;
                foreach (string name in keyColumnNames)
                {
                    if (name == columnName)
                    {
                        return i;
                    }
                }
            }
            return -1;
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

