using System;
using System.Collections.Generic;
using System.IO;
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
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Page = System.Windows.Controls.Page;
using System.Net.NetworkInformation;
using System.Threading;
using System.Diagnostics;
using Application = Microsoft.Office.Interop.Excel.Application;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Reflection;
using static Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System;
using Task = System.Threading.Tasks.Task;

namespace SerpCollPoj
{
    /// <summary>
    /// Логика взаимодействия для Year.xaml
    /// </summary>
    public partial class Year : Page
    {
        delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);

        
    public static string path = ""; //Путь к директории
        public Year()
        {
            InitializeComponent();
            pb.Visibility = Visibility.Hidden;
        }

        private void Button_Cl1ick22(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new CommonOpenFileDialog();
                dialog.IsFolderPicker = true;
                CommonFileDialogResult result = dialog.ShowDialog();
                path = dialog.FileName;
                
            }
            catch { }

        }
        public void stat()
        {
           


            double value1 = 100 / 14;
            double value = 0;
            UpdateProgressBarDelegate updProgress = new UpdateProgressBarDelegate(pb.SetValue);
            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value += value1 });
            if (path != "" && tyt != "")
            {
                string[] days = new string[7] { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };
                string[] daysrus = new string[7] { "пн", "вт", "ср", "чт", "пт", "сб", "вс" };

                string[,] mass = new string[12, 32];
                for (int i = 9; i <= 12; i++)
                {
                    mass[i - 9, 0] = Convert.ToString(System.DateTime.DaysInMonth(Convert.ToInt32(tyt), i));
                    for (int j = 0; j < Convert.ToInt32(mass[i - 9, 0]); j++)
                    {
                        DateTime dt = new DateTime(Convert.ToInt32(tyt), i, j + 1);
                        string p = dt.DayOfWeek.ToString();
                        for (int x = 0; x < 7; x++)
                        {
                            if (days[x] == p)
                            {
                                mass[i - 9, j + 1] = Convert.ToString(daysrus[x]);
                            }
                        }
                    }
                }
                for (int i = 0; i < 8; i++)
                {
                    mass[i + 4, 0] = Convert.ToString(System.DateTime.DaysInMonth(Convert.ToInt32(tyt) + 1, i + 1));
                    for (int j = 0; j < Convert.ToInt32(mass[i + 4, 0]); j++)
                    {
                        DateTime dt = new DateTime(Convert.ToInt32(tyt) + 1, i + 1, j + 1);
                        string p = dt.DayOfWeek.ToString();
                        for (int x = 0; x < 7; x++)
                        {
                            if (days[x] == p)
                            {
                                mass[i + 4, j + 1] = Convert.ToString(daysrus[x]);
                            }
                        }
                    }
                }
                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value += value1 });

                string[] months = new string[] { "сентябрь", "октябрь", "ноябрь", "декабрь", "январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август" };
                string[] lan = new string[37] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK" };
                Console.WriteLine();
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWbSource;
                xlWbSource = xlApp.Workbooks.Add(Type.Missing);
                Excel.Worksheet sheet = xlWbSource.Sheets.Add(Type.Missing, Type.Missing, 12, Type.Missing)
                                as Excel.Worksheet;
                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


                    sheet = (Excel.Worksheet)xlWbSource.Sheets[i + 1];
                    sheet.Name = months[i];






                    sheet.Cells[5][5].Value2 = months[i];
                    sheet.Cells[1][5].Value2 = "№";
                    sheet.Cells[2][5].Value2 = "Преподаватель";
                    sheet.Cells[3][5].Value2 = "Индекс";
                    sheet.Cells[4][5].Value2 = "Дисциплина";
                    sheet.Cells[5][2].Value2 = "Ведомость учета часов работы преподавателей";
                    for (int j = 0; j < Convert.ToInt32(mass[i, 0]); j++)
                    {
                        sheet.Cells[j + 5][6].Value2 = mass[i, j + 1];

                        if (mass[i, j + 1] == "сб" || mass[i, j + 1] == "вс")
                        {
                            sheet.get_Range($"{lan[j + 4]}6", $"{lan[j + 4]}33").Font.Color = Excel.XlRgbColor.rgbBlue;
                            sheet.get_Range($"{lan[j + 4]}6", $"{lan[j + 4]}33").Interior.ColorIndex = 34;

                        }

                        sheet.Cells[j + 5][7].Value2 = j + 1;
                    }
                    Excel.Range aRange = sheet.get_Range("E6", "AI33");
                    aRange.Columns.AutoFit();


                    sheet.Cells[Convert.ToInt32(mass[i, 0]) + 5][5].Value2 = "Всего";

                    sheet.Cells[Convert.ToInt32(mass[i, 0]) + 6][5].Value2 = "остаток";

                    Excel.Range ost = (Excel.Range)sheet.get_Range($"{lan[Convert.ToInt32(mass[i, 0]) + 5]}5", $"{lan[Convert.ToInt32(mass[i, 0]) + 5]}7").Cells;
                    ost.Merge(Type.Missing);
                    Excel.Range vs = (Excel.Range)sheet.get_Range($"{lan[Convert.ToInt32(mass[i, 0]) + 4]}5", $"{lan[Convert.ToInt32(mass[i, 0]) + 4]}7").Cells;
                    vs.Merge(Type.Missing);


                    Excel.Range _excelCellss = (Excel.Range)sheet.get_Range("A5", $"{lan[Convert.ToInt32(mass[i, 0]) + 5]}33");


                    _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
                    _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные            
                    _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                    _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                    _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                    _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    Excel.Range _excelCells111 = (Excel.Range)sheet.get_Range($"A5", $"A7").Cells;
                    _excelCells111.Merge(Type.Missing);
                    Excel.Range _excelCells11 = (Excel.Range)sheet.get_Range($"B5", $"B7").Cells;
                    _excelCells11.Merge(Type.Missing);
                    Excel.Range _excelCells1 = (Excel.Range)sheet.get_Range($"C5", $"C7").Cells;
                    _excelCells1.Merge(Type.Missing);
                    Excel.Range _excelCells = (Excel.Range)sheet.get_Range($"D5", $"D7").Cells;

                    _excelCells.Merge(Type.Missing);
                    Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value += value1 });
                }
                xlWbSource.SaveAs(path + $"/Часы {tyt} - {Convert.ToString(Convert.ToInt32(tyt) + 1)}" + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//
                xlWbSource.Close(true, path + "/" + $"Часы {tyt} - {Convert.ToString(Convert.ToInt32(tyt) + 1)}" + ".xlsx", Type.Missing);

            }
            else
            {
                MessageBox.Show("Некоторые данные введены неверно");
            }
            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value += 100 });
        }
        public string tyt = "";
        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            tyt = year.Text;
            pb.Visibility = Visibility.Visible;
            await Task.Run(() => stat());
        }

        private void year_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                tbye.Text = year.Text + " - " + Convert.ToString(Convert.ToInt32(year.Text) + 1);
            }
            catch { }
        }
    }
}
