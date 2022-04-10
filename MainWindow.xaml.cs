using HtmlAgilityPack;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace XLSXtoHTML
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string[,] _result;

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Получение данных из excel
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="result"></param>
        public string[,] LoadXLSX(string filename)
        {
            string[,] result = null;

            string FileName = filename;
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            try
            {
                workbooks = app.Workbooks as Excel.Workbooks;
                workbook = workbooks.Open(FileName, MissingObj, rOnly);

                // Получение всех страниц докуента
                sheets = workbook.Sheets as Excel.Sheets;

                foreach (Excel.Worksheet worksheet in sheets)
                {
                    // Получаем диапазон используемых на странице ячеек
                    Excel.Range usedRange = worksheet.UsedRange as Excel.Range;
                    // Получаем строки в используемом диапазоне
                    Excel.Range urRows = usedRange.Rows as Excel.Range;
                    // Получаем столбцы в используемом диапазоне
                    Excel.Range urColums = usedRange.Columns as Excel.Range;


                    // Количества строк и столбцов
                    int rowsCount = urRows.Count;
                    int columnsCount = urColums.Count;
                    result = new string[rowsCount, columnsCount];
                    for (int i = 1; i <= rowsCount; i++)
                    {
                        for (int j = 1; j <= columnsCount; j++)
                        {
                            Excel.Range cellRange = usedRange.Cells[i, j] as Excel.Range;
                            // Получение текста ячейки
                            string cellText = (cellRange == null || cellRange.Value2 == null) ? null :
                                                (cellRange as Excel.Range).Value2.ToString();

                            result[i - 1, j - 1] = cellText;
                        }
                    }
                    // Очистка неуправляемых ресурсов на каждой итерации
                    if (urRows != null) Marshal.ReleaseComObject(urRows);
                    if (urColums != null) Marshal.ReleaseComObject(urColums);
                    if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                    if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                /* Очистка оставшихся неуправляемых ресурсов */
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (workbook != null)
                {
                    workbook.Close(SaveChanges);
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }

                if (workbooks != null)
                {
                    workbooks.Close();
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }

            return result;
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new();
                openFileDialog.Filter = "Excel files (*.xlsx) | *.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                _result = LoadXLSX(openFileDialog.FileName);

                FilePath.Text = "Выбран файл: " + openFileDialog.FileName;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (_result == null)
                return;

            SaveFileDialog saveFileDialog = new();
            if (saveFileDialog.ShowDialog() == true)
            {
                StreamWriter streamWriter = new(saveFileDialog.FileName);

                string table = "";
                table += "<table>\n";
                table += "<tbody>\n";
                for (int i = 0; i < _result.GetLength(0); i++)
                {
                    table += "<tr>\n";
                    for (int j = 0; j < _result.GetLength(1); j++)
                    {
                        table += "<td>\n";

                        table += "" + _result[i, j] + "\n";

                        table += "</td>\n";
                    }    
                    table += "</tr>\n";
                }
                table += "</tbody>\n";
                table += "</table>\n";

                streamWriter.WriteLine(table);
                streamWriter.Close();
            }
        }
    }
}
