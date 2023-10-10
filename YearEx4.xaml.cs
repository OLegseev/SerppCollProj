using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace SerpCollPoj
{
    /// <summary>
    /// Логика взаимодействия для YearEx4.xaml
    /// </summary>
    public partial class YearEx4 : Page
    {

        public static string VedDir; //папка
        public static List<string> Syrok = new List<string>(); //Путь к расписанию
        public static List<string> GroupsForVed = new List<string>();
        public static List<string> AllLastName = new List<string>();

        public class DataVed
        {
            public double Groups;
            public string Fio;
            public double firstsem;
            public double secsem;
            public double HourFirst;
            public double HourSec;

        }
        public static List<DataVed> dataVeds = new List<DataVed>();
        public static string PathTo;
        public static string Pathfrom;
        public static string YearL;
        public static bool bobl = false;
        public YearEx4()
        {
            InitializeComponent();
            IOpen.Visibility = Visibility.Hidden;
            Persent.Visibility = Visibility.Hidden;

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == true)
                Syrok = openFileDialog.FileNames.ToList();
            bt1.Background = Brushes.Green;

        }
        private async void Inputted_Click(object sender, RoutedEventArgs e)
        {
            Inputted.Visibility = Visibility.Hidden;
            bt1.Visibility = Visibility.Hidden;
            bts.Visibility = Visibility.Hidden;
            if (Syrok.Count > 0)
            {
                PathTo = path.Text;
                Pathfrom = YearLS_Copy.Text;
                Persent.Visibility = Visibility.Visible;
                foreach (var item in Syrok)
                {
                    ReadAllVed(item);
                }
                Persent.Text = "Происходит формирование файла";
                await Task.Run(() => CreateDoc());

                Persent.Text = "Формирование файла завершено!Ведомость сохранена в папке C//";
                IOpen.Visibility = Visibility.Visible;

            }
            else
            {
                Inputted.Visibility = Visibility.Visible;
                bt1.Visibility = Visibility.Visible;
                bts.Visibility = Visibility.Visible;
                MessageBox.Show("Отсутствуют ведомости");
            }
            Inputted.Visibility = Visibility.Visible;
            bt1.Visibility = Visibility.Visible;
            bts.Visibility = Visibility.Visible;
        }



        public static void ReadAllVed(string path)
        {
            string[] groups = new string[2];
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            xlWorkBook = xlApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); // 1 секунда
            Excel.Worksheet sheet = (Excel.Worksheet)xlApp.Worksheets.get_Item(1);
            for (int i = 0; i < 2; i++)
            {
                if (sheet.Cells[i + 2][3].Text != "" && sheet.Cells[i + 2][3].Text != null)
                {
                    groups[i] = sheet.Cells[i + 2][3].Text;
                    GroupsForVed.Add(sheet.Cells[i + 2][3].Text);
                }
            }// здесь он должен взять все номера групп и записать их в индекс
            string dataAll = "";
            int indexOfall = 0;
            while (dataAll.ToLower().Replace(" ", "").Replace(" ", "").Replace(" ", "") != "всего")
            {
                indexOfall += 1;
                dataAll = sheet.Cells[11][indexOfall].Value2;
                if (dataAll == null)
                {
                    dataAll = "";
                }
            }//ищет колонку всего чтоб понять до скольки идти по фамилиям
            //Первое это колонка второе строчка
            try
            {
                for (int i = 2; i < groups.Length + 2; i++)
                {

                    for (int j = 9; j < indexOfall; j++)
                    {
                        if (sheet.Cells[i][j].Value2 != null && sheet.Cells[i][j].Value2 != "")
                        {
                            var SecondaryName = AllLastName.IndexOf(sheet.Cells[i][j].Value2);
                            if (SecondaryName == -1)
                            {
                                AllLastName.Add(sheet.Cells[i][j].Value2);
                            }
                            dataVeds.Add(new DataVed { Groups = Convert.ToDouble(groups[i - 2]), Fio = sheet.Cells[i][j].Value2, firstsem = sheet.Cells[8][j].Value2, secsem = sheet.Cells[9][j].Value2, HourFirst = sheet.Cells[12][j].Value2, HourSec = sheet.Cells[13][j].Value2 });
                        }
                    }

                }
            }
            catch (Exception e)
            {
                MessageBox.Show("      Наиболее вероятная ошибка: \n     В задействованных полях присутствует пустое поле(если в нем нет данных лучше всего его заполнить 0) \n      " + e);
            }

            xlWorkBook.Close();
        }









        public static void Fulling(Excel.Worksheet worksheet)
        {
            worksheet.Cells[4][1].Value2 = "1 семестр";
            worksheet.Cells[3][1].Value2 = DateTime.Now.ToString("dd.MM.yyyy") + "Г";
        }
        //public class IdVed
        //{
        //    public double Row;
        //    public string Fio;
        //    public string item;
        //}
        //public static List<IdVed> idVed = new List<IdVed>();


        //public class DataVed1
        //{
        //    public string NameOfItem;
        //    public string index;
        //}
        //public static List<DataVed1> dataVeds1 = new List<DataVed1>();
        public static List<string> Predmet = new List<string>();




        public static void CreateDoc()
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                double sumsum1 = 0;
                double sumsum2 = 0;
                app = new Excel.Application();
                app.Visible = false;
                workbook = app.Workbooks.Add(1);
                worksheet = workbook.Sheets[1];
                Fulling(worksheet);


                double summCELL = 0;
                double summ2CELL = 0;

                for (int g = 0; g < GroupsForVed.Count; g++)
                {
                    worksheet.Cells[g + 4][2].Value2 = GroupsForVed[g];
                    worksheet.Cells[g + 4 + GroupsForVed.Count + 1][2].Value2 = GroupsForVed[g];
                }
                worksheet.Cells[4 + GroupsForVed.Count][2].Value2 = "Итог";
                worksheet.Cells[4 + GroupsForVed.Count + 1 + GroupsForVed.Count][2].Value2 = "Итог";
                int rowid = 3;

                for (int i = 0; i < AllLastName.Count; i++)
                {

                    double schet = 1;
                    worksheet.Cells[2][rowid].Value2 = schet;
                    schet++;
                    worksheet.Cells[3][rowid].Value2 = AllLastName[i];
                    double summ = 0;
                    double summ2 = 0;

                    var curs = dataVeds.Where(x => x.Fio == AllLastName[i]).ToList();
                    int itemm = 0;
                    foreach (var item in GroupsForVed)
                    {
                        var surs = curs.Where(x => x.Groups == Convert.ToDouble(item)).ToList();

                        if (bobl == false)
                        {
                            for (int j = 0; j < surs.Count; j++)
                            {
                                summCELL += surs[j].firstsem;
                                summ2CELL += surs[j].secsem;
                            }
                        }
                        else
                        {
                            for (int j = 0; j < surs.Count; j++)
                            {
                                summCELL += surs[j].HourFirst;
                                summ2CELL += surs[j].HourSec;
                            }
                        }
                        worksheet.Cells[4 + itemm][rowid].Value2 = summCELL;
                        summ += summCELL;

                        worksheet.Cells[4 + itemm + GroupsForVed.Count + 1][rowid].Value2 = summ2CELL;
                        summ2 += summ2CELL;
                        itemm++;
                        summCELL = 0;
                        summ2CELL = 0;

                    }





                    //for (int j = 0; j < curs.Count; j++)
                    //{

                    //    var groop = GroupsForVed.IndexOf(Convert.ToString(curs[j].Groups));
                    //    worksheet.Cells[4 + groop][rowid].Value2 = curs[j].firstsem;
                    //    summ += curs[j].firstsem;
                    //    worksheet.Cells[4 + groop + GroupsForVed.Count + 1][rowid].Value2 = curs[j].secsem;
                    //    summ2 += curs[j].secsem;
                    //}
                    worksheet.Cells[4 + GroupsForVed.Count][rowid].Value2 = summ;
                    worksheet.Cells[4 + GroupsForVed.Count + 1 + GroupsForVed.Count][rowid].Value2 = summ2;
                    sumsum1 += summ;
                    sumsum2 += summ2;
                    summ = 0;
                    summ2 = 0;
                    rowid++;

                }
                worksheet.Cells[4 + GroupsForVed.Count][rowid].Value2 = sumsum1;
                worksheet.Cells[4 + GroupsForVed.Count + 1 + GroupsForVed.Count][rowid].Value2 = sumsum2;


                for (int h = 0; h < GroupsForVed.Count; h++)
                {
                    double hc = 0; double hc1 = 0;


                    var hoursCurs = dataVeds.Where(x => x.Groups == Convert.ToDouble(GroupsForVed[h])).ToList();


                    if (bobl == false)
                    {
                        for (int l = 0; l < hoursCurs.Count; l++)
                        {
                            hc += hoursCurs[l].firstsem;
                            hc1 += hoursCurs[l].secsem;
                        }
                    }
                    else
                    {
                        for (int l = 0; l < hoursCurs.Count; l++)
                        {
                            hc += hoursCurs[l].HourFirst;
                            hc1 += hoursCurs[l].HourSec;
                        }
                    }




                    worksheet.Cells[4 + h][rowid].Value2 = hc;
                    worksheet.Cells[4 + GroupsForVed.Count + 1 + h][rowid].Value2 = hc1;
                }
                workbook.SaveAs(PathTo + Pathfrom + @".xlsx");
                workbook.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show("" + e);
                workbook.Close();
            }
            try
            {

                Predmet.Clear();
                VedDir = null;
                Syrok.Clear();
                GroupsForVed.Clear();
                AllLastName.Clear();
                dataVeds.Clear();

            }
            catch { }


        }

        private void path_Copy_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void path_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void IOpen_Click(object sender, RoutedEventArgs e)
        {

            Process.Start($"{PathTo}{Pathfrom}.xlsx");
        }

        private void checkB_Checked(object sender, RoutedEventArgs e)
        {
            bobl = true;
        }

        private void checkB_Unchecked(object sender, RoutedEventArgs e)
        {
            bobl = false;
        }
    }
}
