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

namespace SerpCollPoj
{
    /// <summary>
    /// Логика взаимодействия для YearTeacher.xaml
    /// </summary>
    public partial class YearTeacher : Page
    {
        delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);
        
        public string path = "";
        public List<string> MultiPaths = new List<string>();
        public List<string> prepods = new List<string>();
        public string[] mass = new string[] { "сентябрь - 1", "октябрь - 2", "ноябрь - 3", "декабрь - 4", "январь - 5", "февраль - 6", "март - 7", "апрель - 8", "май - 9", "июнь - 10", "июль - 11", "август - 12" };
        public string mounth = "";
        
        public YearTeacher()
        {
            InitializeComponent();
            
            pb.Visibility = Visibility.Hidden;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                if (openFileDialog.ShowDialog() == true)
                    MultiPaths = openFileDialog.FileNames.ToList();
                
            }
            catch { }
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

        public class ved
        {
            public string name { get; set; }
            public string id { get; set; }
            public string item { get; set; }

            public string data { get; set; }
            public string number { get; set; }
            public string group { get; set; }
            public string vsego { get; set; }
            public string ostatok { get; set; }
        }
        public List<ved> veds = new List<ved>();
        int datamax = 0;
        public void workMethod()
        {
            
            double value1 = 100 / (MultiPaths.Count * 5);
            double value = 0;
            UpdateProgressBarDelegate updProgress = new UpdateProgressBarDelegate(pb.SetValue);
            
            int y1 = 1;
            for (int j = 0; j < 12; j++)
            {
                if (mounth == mass[j])
                {
                    y1 = j;
                }
            }
                int isch = 0;
           
            for (int i1 = 0; i1 < isch; i1++)
            {
                int y = 0;

                for (int j = 0; j < 12; j++)
                {
                    if (mounth == mass[j])
                    {
                        y = j;
                    }
                }
                for (int i = 0; i < MultiPaths.Count; i++)
                {
                    Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value += value1 });

                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWbSource;
                    xlWbSource = xlApp.Workbooks.Open(MultiPaths[i]);
                    Excel.Worksheet sheet = (Excel.Worksheet)xlWbSource.Sheets[y + 2];
                    int p = 8;
                    for (int j = 8; j < 70; j++)
                    {
                        try
                        {
                            if (sheet.Cells[4][j].Value2 != "" || sheet.Cells[4][j].Value2 != "0" || sheet.Cells[4][j].Value2 != null)
                            {
                                p++;
                            }
                            else
                            {
                                break;
                            }

                        }
                        catch
                        {

                            break;
                        }
                    }
                    int o = 4;
                    for (int j = 4; j < 50; j++)
                    {
                        try
                        {
                            if (Convert.ToString(sheet.Cells[j][6].Value2).ToLower() == "остаток")
                            {
                                break;
                            }
                            else
                            {
                                o++;

                            }
                        }
                        catch
                        {

                        }
                    }
                    datamax = o;
                    int[] massive = new int[31];
                    int u = 0;
                    Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value += value1 });
                    for (int j = 5; j < o + 1; j++)
                    {
                        try
                        {
                            if (sheet.Cells[j][6].Value2 != "сб" && sheet.Cells[j][6].Value2 != "вс" && sheet.Cells[j][6].Value2 != "" && sheet.Cells[j][6].Value2 != " " && sheet.Cells[j][6].Value2 != null)
                            {
                                massive[u] = j;
                                string pp = sheet.Cells[j][6].Value2;
                                u++;
                            }
                        }
                        catch { }
                    }
                    string[] massstr = sheet.Cells[5][4].Value2.Split(' ');
                    Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value += value1 });
                    for (int r = 8; r <= p; r++)
                    {
                        for (int j = 0; j < u; j++)
                        {
                            if (massive[j] == 33)
                            {
                                Console.WriteLine();
                            }
                            try
                            {
                                if (sheet.Cells[massive[j]][r].Value2 == 2 || sheet.Cells[massive[j]][r].Value2 == 4 || sheet.Cells[massive[j]][r].Value2 == 6 || sheet.Cells[massive[j]][r].Value2 == 8)
                                {
                                    veds.Add(new ved
                                    {
                                        name = Convert.ToString(sheet.Cells[2][r].Value2),
                                        id = Convert.ToString(sheet.Cells[3][r].Value2),
                                        item = Convert.ToString(sheet.Cells[4][r].Value2),
                                        data = Convert.ToString(sheet.Cells[massive[j]][7].Value2),
                                        number = Convert.ToString(sheet.Cells[massive[j]][r].Value2),
                                        group = massstr[massstr.Length - 1],
                                        vsego = Convert.ToString(sheet.Cells[36][r].Value2),
                                        ostatok = Convert.ToString(sheet.Cells[37][r].Value2)

                                    });
                                }
                            }
                            catch (Exception ex)
                            {

                            }

                        }
                    }
                    Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value += value1 });


                    xlWbSource.Close();

                }


                int pp1 = 0;
                for (int i = 0; i < prepods.Count; i++)
                {
                    string h = "";
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWbSource;
                    xlWbSource = xlApp.Workbooks.Open(prepods[i]);
                    Excel.Worksheet sheet = (Excel.Worksheet)xlWbSource.Sheets[y + 1];
                    for (int j = 1; j < 50; j++)
                    {
                        if (sheet.Cells[j][5].Value2 != "" && sheet.Cells[j][5].Value2 != null)
                        {
                            h = sheet.Cells[j][5].Value2;
                            if (h.Contains("."))
                            {
                                pp1 = j;
                                break;
                            }
                        }
                    }
                    var cd = veds.Where(x => x.name == h).ToList();
                    int op = 0;
                    for (int j = 9; j < 40; j++)
                    {
                        if (Convert.ToString(sheet.Cells[9][j].Value2) == "" || Convert.ToString(sheet.Cells[9][j].Value2) == " " || Convert.ToString(sheet.Cells[9][j].Value2) == null)
                        {
                            op = j;
                            break;

                        }
                    }
                    for (int j = 9; j < op; j++)
                    {

                        for (int c = 0; c < cd.Count; c++)
                        {






                            if (Convert.ToString(sheet.Cells[8][j].Value2) == cd[c].id && Convert.ToString(sheet.Cells[9][j].Value2) == cd[c].item && Convert.ToString(sheet.Cells[10][j].Value2) == cd[c].group)
                            {
                                sheet.Cells[Convert.ToInt32(cd[c].data) + 11][j].Value2 = cd[c].number;
                                sheet.Cells[datamax + 9][j].Value2 = cd[c].vsego;
                                sheet.Cells[datamax + 10][j].Value2 = cd[c].ostatok;
                            }
                        }

                    }




                    xlWbSource.SaveAs(prepods[i]);
                    xlWbSource.Close();
                }
                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value = 100 });


            }

        }
        //public class ved
        //{
        //    public string name { get; set; }
        //    public string id { get; set; }
        //    public string item { get; set; }

        //    public string data { get; set; }
        //    public string number { get; set; }
        //    public string group { get; set; }
        //    public string vsego { get; set; }
        //    public string ostatok { get; set; }
        //}
        private async void Inputted_Click(object sender, RoutedEventArgs e)
        {
            
            
            pb.Visibility = Visibility.Visible;
            await System.Threading.Tasks.Task.Run(() =>workMethod());
            
        }

        private void Button_Click2(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                if (openFileDialog.ShowDialog() == true)
                    prepods = openFileDialog.FileNames.ToList();
                
            }
            catch { }
        }

 
    }
}
