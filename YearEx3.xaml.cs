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
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Page = System.Windows.Controls.Page;

namespace SerpCollPoj
{
    /// <summary>
    /// Логика взаимодействия для YearEx3.xaml
    /// </summary>
    public partial class YearEx3 : Page
    {
        public static string VedDir; //папка
        public static List<string> Syrok = new List<string>(); //Путь к расписанию
        public static List<string> GroupsForVed = new List<string>();
        public static List<string> AllLastName = new List<string>();

        public class DataVed
        {
            public double Groups;
            public string Fio;
            public double RasCroup;
            public string Index;
            public double HourseByCurse;
            public string NameOfItem;
            public double Ecsam;
            public double consult;

        }
        public static List<DataVed> dataVeds = new List<DataVed>();
        public static string PathTo;
        public static string Pathfrom;
        public static string YearL;

        public YearEx3()
        {
            InitializeComponent();

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
                YearL = YearLS.Text;
                Persent.Visibility = Visibility.Visible;
                foreach (var item in Syrok)
                {
                    ReadAllVed(item);
                }
                Persent.Text = "Происходит формирование файла";
                await Task.Run(() => CreateDoc());

                Persent.Text = "Формирование файла завершено!Ведомость сохранена в папке C//";
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
                dataAll = Convert.ToString(sheet.Cells[11][indexOfall].Value2);
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
                            dataVeds.Add(new DataVed { Groups = Convert.ToDouble(groups[i - 2]), Fio = sheet.Cells[i][j].Value2, RasCroup = sheet.Cells[5][j].Value2, Index = sheet.Cells[10][j].Value2, HourseByCurse = sheet.Cells[4][j].Value2, NameOfItem = sheet.Cells[11][j].Value2, consult = Convert.ToInt32(sheet.Cells[6][j].Value2), Ecsam = Convert.ToInt32(sheet.Cells[7][j].Value2) });
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
            worksheet.Cells[1][2].Value2 = "Расчет часов и нагрузки преподавателей на " + YearL + " учебный год";
            worksheet.Cells[1][3].Value2 = "ПО ОБЩЕПРОФЕССИОНАЛЬНЫМ И СПЕЦИАЛЬНЫМ ДИСЦИПЛИНАМ    ";
            worksheet.Cells[2][4].Value2 = "Фамилия, имя, отчество преподавателей";
            worksheet.Cells[3][4].Value2 = "ИНДЕКС";
            worksheet.Cells[4][4].Value2 = "Дисциплины";
            worksheet.Cells[5][5].Value2 = "Код группы";
            worksheet.Cells[5][6].Value2 = "Курс";
            worksheet.Cells[5][7].Value2 = "Численность";
            worksheet.Cells[6][4].Value2 = "НАИМЕНОВАНИЕ  ГРУПП";
            worksheet.Cells[5 + GroupsForVed.Count + 1][4].Value2 = "разукрупнения";
            worksheet.Cells[5 + GroupsForVed.Count + 3][4].Value2 = "экзамены";
            worksheet.Cells[5 + GroupsForVed.Count + 2][4].Value2 = "консультации";
            worksheet.Cells[5 + GroupsForVed.Count + 4][4].Value2 = "Всего часов";
            worksheet.Cells[5 + GroupsForVed.Count + 5][4].Value2 = "итог";
        }
        //public class IdVed
        //{
        //    public double Row;
        //    public string Fio;
        //    public string item;
        //}
        //public static List<IdVed> idVed = new List<IdVed>();


        public class DataVed1
        {
            public string NameOfItem;
            public string index;
        }
        public static List<DataVed1> dataVeds1 = new List<DataVed1>();
        public static List<string> Predmet = new List<string>();




        public static void CreateDoc()
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                app = new Excel.Application();
                app.Visible = false;
                workbook = app.Workbooks.Add(1);
                worksheet = workbook.Sheets[1];
                Fulling(worksheet);




                for (int g = 0; g < GroupsForVed.Count; g++)
                {
                    worksheet.Cells[g + 6][5].Value2 = GroupsForVed[g];

                }



                int rowid = 8;
                double hours2 = 0;
                for (int i = 0; i < AllLastName.Count; i++)
                {
                    double summ = 0;
                    worksheet.Cells[2][rowid].Value2 = AllLastName[i];
                    var fam1 = dataVeds.Where(x => x.Fio == AllLastName[i]).ToList();
                    for (int j = 0; j < fam1.Count; j++)
                    {
                        //var povtor = fam1.Where(x => x.NameOfItem == fam1[j].NameOfItem).ToList();
                        var datanoname = dataVeds1.Where(x => x.NameOfItem == fam1[j].NameOfItem && x.index == fam1[j].Index).ToList();
                        if (datanoname.Count == 0)
                        {
                            dataVeds1.Add(new DataVed1 { NameOfItem = fam1[j].NameOfItem, index = fam1[j].Index });
                        }////////////////////////
                    }
                    double hours3 = 0;
                    for (int j = 0; j < dataVeds1.Count; j++)
                    {

                        worksheet.Cells[3][rowid].Value2 = dataVeds1[j].index;
                        worksheet.Cells[4][rowid].Value2 = dataVeds1[j].NameOfItem;
                        double rasucrupn = 0;
                        double ecsam = 0;
                        double consT = 0;
                        double hours = 0;

                        var datanoname1 = fam1.Where(x => x.NameOfItem == dataVeds1[j].NameOfItem && x.Index == dataVeds1[j].index).ToList();
                        for (int k = 0; k < datanoname1.Count; k++)
                        {
                            summ += datanoname1[k].HourseByCurse;

                            rasucrupn += datanoname1[k].RasCroup;
                            ecsam += datanoname1[k].Ecsam;
                            consT += datanoname1[k].consult;
                            hours += datanoname1[k].HourseByCurse;
                            var group = GroupsForVed.IndexOf(datanoname1[k].Groups.ToString());
                            worksheet.Cells[5 + group + 1][rowid].Value2 = datanoname1[k].HourseByCurse;
                            //worksheet.Cells[4 + group+2 + 4][rowid].Value2 = datanoname1[k].HourseByCurse;


                        }
                        hours2 += hours + ecsam + consT + rasucrupn;
                        hours3 += hours + ecsam + consT + rasucrupn;
                        worksheet.Cells[5 + GroupsForVed.Count + 1][rowid].Value2 = rasucrupn;
                        worksheet.Cells[5 + GroupsForVed.Count + 3][rowid].Value2 = ecsam;
                        worksheet.Cells[5 + GroupsForVed.Count + 4][rowid].Value2 = hours + ecsam + consT + rasucrupn;
                        worksheet.Cells[5 + GroupsForVed.Count + 2][rowid].Value2 = consT;

                        rowid++;

                    }
                    worksheet.Cells[5 + GroupsForVed.Count + 5][rowid - 1].Value2 = hours3;
                    hours3 = 0;


                    rowid++;
                    dataVeds1.Clear();




                }
                for (int h = 0; h < GroupsForVed.Count; h++)
                {
                    double hc = 0;


                    var hoursCurs = dataVeds.Where(x => x.Groups == Convert.ToDouble(GroupsForVed[h])).ToList();

                    for (int l = 0; l < hoursCurs.Count; l++)
                    {
                        hc += hoursCurs[l].HourseByCurse;
                    }





                    worksheet.Cells[5 + h + 1][rowid].Value2 = hc;
                }
                double ras = 0;
                double acs = 0;
                double cos = 0;
                for (int i = 0; i < dataVeds.Count; i++)
                {
                    ras += dataVeds[i].RasCroup;
                    acs += dataVeds[i].Ecsam;
                    cos += dataVeds[i].consult;

                }
                worksheet.Cells[5 + GroupsForVed.Count + 1][rowid].Value2 = ras;
                worksheet.Cells[5 + GroupsForVed.Count + 3][rowid].Value2 = acs;
                worksheet.Cells[5 + GroupsForVed.Count + 2][rowid].Value2 = cos;

                worksheet.Cells[5 + GroupsForVed.Count + 5][rowid].Value2 = hours2;


                //Excel.Range _excelCellss = (Excel.Range)worksheet.get_Range("A2", $"AZ{GroupsForVed.Count + dataVeds1.Count}");



                //_excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
                //_excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные            
                //_excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                //_excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                //_excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                //_excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                Excel.Range aRange = worksheet.get_Range("A2", $"AZ{200}");
                aRange.Columns.AutoFit();
                workbook.SaveAs(PathTo + Pathfrom + @"xlsx");
                workbook.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("" + e);
                workbook.Close();
            }









            try
            {
                dataVeds1.Clear();
                Predmet.Clear();
                VedDir = null;
                Syrok.Clear();
                GroupsForVed.Clear();
                AllLastName.Clear();
                dataVeds.Clear();
                PathTo = null;
                YearL = null;
            }
            catch { }
        }

        private void path_Copy_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void path_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
