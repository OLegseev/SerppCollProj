using Microsoft.Win32;
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
using static SerpCollPoj.CurrentData;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Page = System.Windows.Controls.Page;

namespace SerpCollPoj
{
    /// <summary>
    /// Логика взаимодействия для YearEx5.xaml
    /// </summary>
    public partial class YearEx5 : Page
    {
        public static List<string> Syrok = new List<string>(); //Путь к Ведомости
        public static List<string> Potolok = new List<string>(); //Путь к списку
        public static string Path = ""; //Путь к директории

        public YearEx5()
        {
            InitializeComponent();


        }
        private void Button_Cl1ick(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                if (openFileDialog.ShowDialog() == true)
                    Potolok = openFileDialog.FileNames.ToList();

                bt1.Background = Brushes.Green;
            }
            catch { }
        }
        //86у
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                if (openFileDialog.ShowDialog() == true)
                    Syrok = openFileDialog.FileNames.ToList();
                if (Syrok.Count > 1)
                {
                    MessageBox.Show("Нельзя вносить больше 1 ведомости за раз");
                    Syrok = null;
                }
                else
                    bt11.Background = Brushes.Green;
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
                Path = dialog.FileName;
                bt22.Background = Brushes.Green;
            }
            catch { }
        }

        private void Inputted_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show(@"Ошибка версии excel\word");
            Inputted.Visibility = Visibility.Hidden;

            if (Syrok.Count > 0F && Potolok.Count > 0 && Path != "")
            {
                ReadAllVed(Syrok[0]);
                for (int i = 0; i < Potolok.Count; i++)
                {
                    ReadList(Potolok[i], i);
                    for (int j = 0; j < Subject.Count; j++)
                    {
                        CreateDoc(i, j);
                        //CreateDocEcz(i, j);
                    }
                    studentList.Clear();
                }

                ReadAllVedEcz(Syrok[0]);
                for (int i = 0; i < Potolok.Count; i++)
                {
                    ReadList(Potolok[i], i);
                    for (int j = 0; j < Subject.Count; j++)
                    {

                        CreateDocEcz(i, j);
                    }
                    studentList.Clear();
                }

                foo();





            }
            else
            {
                Inputted.Visibility = Visibility.Visible;
                Path = "";
                Syrok = null;
                Potolok = null;
                MessageBox.Show("Отсутствуют ведомости");
            }

        }
        public void foo()
        {
            Inputted.Visibility = Visibility.Hidden;
          
            if (Syrok.Count > 0F && Potolok.Count > 0 && Path != "")
            {
                ReadUsless(Syrok[0]);

                for (int j = 0; j < groups1.Length; j++)
                {
                    CreateDocUssl(j);
                }
            }
            else
            {
                Inputted.Visibility = Visibility.Visible;
                Path = "";
                Syrok = null;
                Potolok = null;
                MessageBox.Show("Отсутствуют ведомости");
            }
        }

        public class UsslVed
        {

            public string NumerVed { get; set; }
            public string NumerEcz { get; set; }
            public string Hours1 { get; set; }
            public string Hours2 { get; set; }
            public string NameFio1Group { get; set; }
            public string NameFio2Group { get; set; }
            public string Subject { get; set; }
            public string Reg { get; set; }
        }
        public static List<UsslVed> usslVed = new List<UsslVed>();


        public void CreateDocUssl(int id)
        {//Метод создающий документ
            DirectoryInfo dirInfo = new DirectoryInfo(Path + @"//Часы//");
            if (!dirInfo.Exists)
            {
                Directory.CreateDirectory(Path + @"//Часы//");
            }

            DirectoryInfo di1;
            di1 = new DirectoryInfo(@"..\..\..");
            string di2 = di1.FullName + @"\Test0.xlsx";

            Excel.Workbook xlWorkBook;
            Excel.Application xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(di2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true); // 1 секунда
            Excel.Worksheet sheet = xlWorkBook.Sheets[1];
            sheet = xlWorkBook.Sheets[1];

            sheet.Cells[2][2] = groups1[id] + "+";
            sheet.Cells[3][2] = "Ф.И.О. \r\nпреподавателя";
            sheet.Cells[5][2] = "Наименование\r\nучебной дисциплины\r\n";
            sheet.Cells[4][2] = "Индекс";
            sheet.Cells[6][2] = "Экзамен";
            sheet.Cells[7][2] = "Зачет";
            sheet.Cells[9][2] = sem;
            sheet.Cells[10][2] = sem1;
            sheet.Cells[8][2] = "Ссылка";
            Excel.Range _excelCellss = (Excel.Range)sheet.get_Range($"B2", $"J{usslVed.Count + 2}").Cells;
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные            
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            string prep;

            for (int i = 0; i < usslVed.Count; i++)
            {
                if (id == 0)
                {
                    prep = usslVed[i].NameFio1Group;
                }
                else
                {
                    prep = usslVed[i].NameFio2Group;
                }
                //usslVed.Add(new UsslVed { Hours = sheet.Cells[4][i].Value2, NumerEcz = ecz, NumerVed = zach, NameFio1Group = sheet.Cells[2][i].Value2, NameFio2Group = sheet.Cells[3][i].Value2, Subject = sheet.Cells[11][i].Value2, Reg = sheet.Cells[10][i].Value2, });

                sheet.Cells[3][3 + i] = prep;
                sheet.Cells[4][3 + i] = usslVed[i].Subject;
                sheet.Cells[5][3 + i] = usslVed[i].Reg;
                sheet.Cells[6][3 + i] = usslVed[i].NumerEcz;
                sheet.Cells[7][3 + i] = usslVed[i].NumerVed;
                sheet.Cells[9][3 + i] = usslVed[i].Hours1;
                sheet.Cells[10][3 + i] = usslVed[i].Hours2;





            }


            xlWorkBook.SaveAs(Path + "/Часы" + "/Часы " + groups1[id] + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//
            xlWorkBook.Close(true, Path + "/Часы" + "/Часы " + groups1[id] + ".xlsx", Type.Missing);

        }










        public static string sem;
        public static string sem1;
        public static string[] groups1 = new string[2];
        public static void ReadUsless(string path)
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            xlWorkBook = xlApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); // 1 секунда
            Excel.Worksheet sheet = (Excel.Worksheet)xlApp.Worksheets.get_Item(1);
            string dataAll = "";
            int indexOfall = 0;
            GroupsForVed.Clear();
            Numer.Clear();
            NameFio1Group.Clear();
            NameFio2Group.Clear();
            Subject.Clear();
            Reg.Clear();
            examVed.Clear();
            for (int i = 0; i < 2; i++)
            {
                if (sheet.Cells[i + 2][3].Text != "" && sheet.Cells[i + 2][3].Text != null)
                {
                    groups1[i] = sheet.Cells[i + 2][3].Text;
                    GroupsForVed.Add(sheet.Cells[i + 2][3].Text);
                }
            }
            spec = "";
            spec = sheet.Cells[11][1].Value2;

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
            YearEx5 workPage = new YearEx5();

            sem = sheet.Cells[12][2].Value2;
            sem1 = sheet.Cells[13][2].Value2;




            string ecz;
            string zach;
            for (int i = 9; i < indexOfall; i++)
            {

                string a = Convert.ToString(sheet.Cells[14][i].Value2);
                int num;
                bool isNum = int.TryParse(Convert.ToString(sheet.Cells[14][i].Value2), out num);
                string a1 = Convert.ToString(sheet.Cells[16][i].Value2);
                int num1;
                bool isNum1 = int.TryParse(Convert.ToString(sheet.Cells[16][i].Value2), out num1);

                if (isNum)
                {
                    ecz = Convert.ToString(sheet.Cells[14][i].Value2);
                }
                else
                {
                    ecz = "";
                }
                if (isNum1)
                {
                    zach = Convert.ToString(sheet.Cells[16][i].Value2);
                }
                else
                {
                    zach = "";
                }
                if (isNum || isNum1)
                {
                    usslVed.Add(new UsslVed { Hours1 = Convert.ToString(sheet.Cells[12][i].Value2), Hours2 = Convert.ToString(sheet.Cells[13][i].Value2), NumerEcz = Convert.ToString(ecz), NumerVed = Convert.ToString(zach), NameFio1Group = Convert.ToString(sheet.Cells[2][i].Value2), NameFio2Group = Convert.ToString(sheet.Cells[3][i].Value2), Subject = Convert.ToString(sheet.Cells[11][i].Value2), Reg = Convert.ToString(sheet.Cells[10][i].Value2), });

                }



            }
            xlWorkBook.Close();
        }

































































        public static List<double> Numer = new List<double>(); //Путь к списку
        public static List<string> NameFio1Group = new List<string>(); //Путь к списку
        public static List<string> NameFio2Group = new List<string>(); //Путь к списку
        public static List<string> Subject = new List<string>(); //Путь к списку
        public static List<string> Reg = new List<string>(); //Путь к списку

        public static void ReadAllVedEcz(string path)
        {
            string[] groups = new string[2];
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            xlWorkBook = xlApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); // 1 секунда
            Excel.Worksheet sheet = (Excel.Worksheet)xlApp.Worksheets.get_Item(1);
            string dataAll = "";
            int indexOfall = 0;
            GroupsForVed.Clear();
            Numer.Clear();
            NameFio1Group.Clear();
            NameFio2Group.Clear();
            Subject.Clear();
            Reg.Clear();
            examVed.Clear();
            for (int i = 0; i < 2; i++)
            {
                if (sheet.Cells[i + 2][3].Text != "" && sheet.Cells[i + 2][3].Text != null)
                {
                    groups[i] = sheet.Cells[i + 2][3].Text;
                    GroupsForVed.Add(sheet.Cells[i + 2][3].Text);
                }
            }
            spec = "";
            spec = sheet.Cells[11][1].Value2;

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
            for (int i = 9; i < indexOfall; i++)
            {
                string a = Convert.ToString(sheet.Cells[14][i].Value2);
                int num;
                bool isNum = int.TryParse(Convert.ToString(sheet.Cells[14][i].Value2), out num);
                if (isNum)
                {
                    Numer.Add(num);
                    NameFio1Group.Add(sheet.Cells[2][i].Value2);
                    NameFio2Group.Add(sheet.Cells[3][i].Value2);
                    Subject.Add(sheet.Cells[11][i].Value2);
                    Reg.Add(sheet.Cells[10][i].Value2);
                    examVed.Add(new ExsamVed { Numer = Convert.ToDouble(sheet.Cells[14][i].Value2), NameFio1Group = sheet.Cells[2][i].Value2, NameFio2Group = sheet.Cells[3][i].Value2, Subject = sheet.Cells[11][i].Value2, Reg = sheet.Cells[10][i].Value2, });
                }
            }
            xlWorkBook.Close();
        }

        public void CreateDocEcz(int id, int subb)
        {//Метод создающий документ
            DirectoryInfo dirInfo = new DirectoryInfo(Path + @"//Экзамен//" + GroupsForVed[id]);
            if (!dirInfo.Exists)
            {
                Directory.CreateDirectory(Path + @"//Экзамен//" + GroupsForVed[id]);
            }

            DirectoryInfo di1;
            di1 = new DirectoryInfo(@"..\..\..");
            string di2 = di1.FullName + @"\Test2.xlsx";

            Excel.Workbook xlWorkBook;
            Excel.Application xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(di2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true); // 1 секунда
            Excel.Worksheet sheet = xlWorkBook.Sheets[1];
            sheet = xlWorkBook.Sheets[1];
            Excel.Range _excelCells = (Excel.Range)sheet.get_Range($"B13", $"E15").Cells;
            _excelCells.Merge(Type.Missing);
            Excel.Range _excelCells4 = (Excel.Range)sheet.get_Range($"A13", $"A15").Cells;
            _excelCells4.Merge(Type.Missing);
            Excel.Range _excelCells3 = (Excel.Range)sheet.get_Range($"F13", $"F15").Cells;
            _excelCells3.Merge(Type.Missing);
            Excel.Range _excelCells31 = (Excel.Range)sheet.get_Range($"G13", $"J14").Cells;
            _excelCells31.Merge(Type.Missing);
            Excel.Range _excelCells111 = (Excel.Range)sheet.get_Range($"I15", $"J15").Cells;
            _excelCells111.Merge(Type.Missing);

            sheet.Cells[1][13] = "№ п/п";
            sheet.Cells[2][13] = "Ф.И.О. Студента";
            sheet.Cells[6][13] = $"Номер экзаменацион\n-ного билета";
            sheet.Cells[7][13] = "Оценки по экзамену";
            sheet.Cells[7][15] = "письменно";
            sheet.Cells[8][15] = "устно";
            sheet.Cells[9][15] = "общая";
            int range = 15 + studentList.Count;

            try
            {
                for (int i = 16; i < 16 + studentList.Count; i++)
                {
                    sheet.Cells[1][i] = i - 15;
                    Excel.Range _excelCells1 = (Excel.Range)sheet.get_Range($"B{i}", $"E{i}").Cells;
                    _excelCells1.Merge(Type.Missing);
                    sheet.Cells[2][i] = studentList[i - 16];
                    sheet.Cells[10][i].Formula = $"=IF(I{i}=5,\"(отлично)\",IF(I{i}=4,\"(хорошо)\",IF(I{i}=3,\"(удовлетворительно)\",IF(I{i}=2,\"(неудовлетворительно)\",))))";

                }
            }
            catch { }
            sheet.Cells[4][7] = Reg[subb] + " " + Subject[subb];
            sheet.Cells[2][8] = GroupsForVed[id];
            sheet.Cells[3][9] = spec;
            if (id == 0)
            {
                sheet.Cells[4][10] = NameFio1Group[subb];
            }
            else
            {
                sheet.Cells[4][10] = NameFio2Group[subb];
            }

            Excel.Range _excelCellss = (Excel.Range)sheet.get_Range($"A14", $"J" + range).Cells;

            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные            
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            Excel.Range _excelCells8 = (Excel.Range)sheet.get_Range($"A{range + 2}", $"E{range + 2}").Cells;
            _excelCells8.Merge(Type.Missing);
            sheet.Cells[1][range + 2] = "Подпись преподавателей _____________________________";
            sheet.Cells[1][range + 3] = "5 «отлично»";
            sheet.Cells[1][range + 4] = "4 «хорошо»";
            sheet.Cells[1][range + 5] = "3 «удовлетворительно»";
            sheet.Cells[1][range + 6] = "2 «неудовлетворительно»";
            sheet.Cells[1][range + 7] = "Качество:";
            for (int i = range + 2; i < range + 8; i++)
            {
                Excel.Range _excelCells7 = (Excel.Range)sheet.get_Range($"A{i}", $"B{i}").Cells;
                _excelCells7.Merge(Type.Missing);
            }
            sheet.Cells[4][range + 3] = "чел.";
            sheet.Cells[4][range + 4] = "чел.";
            sheet.Cells[4][range + 5] = "чел.";
            sheet.Cells[4][range + 6] = "чел.";
            sheet.Cells[4][range + 7] = "чел.";

            sheet.Cells[6][range + 3] = "%";
            sheet.Cells[6][range + 4] = "%";
            sheet.Cells[6][range + 5] = "%";
            sheet.Cells[6][range + 6] = "%";
            sheet.Cells[6][range + 7] = "%";

            sheet.Cells[1][range + 8] = "Средний балл";
            sheet.Cells[1][range + 9] = "«___» _______ 202_г";

            sheet.Cells[3][range + 3].Formula = $"=COUNTIF(I16:I100,\"5\")";
            sheet.Cells[3][range + 4] = $"=COUNTIF( I16:I100,\"4\")";
            sheet.Cells[3][range + 5] = $"=COUNTIF(I16:I100,\"3\")";
            sheet.Cells[3][range + 6] = $"=COUNTIF(I16:I100,\"2\")";
            sheet.Cells[3][range + 7] = $"=C{range + 3}+C{range + 4}";
            sheet.Cells[3][range + 8] = $"=(C{range + 3}*5+C{range + 4}*4+C{range + 5}*3+C{range + 6}*2)/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6})";

            sheet.Cells[5][range + 3] = $"=(C{range + 3}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";
            sheet.Cells[5][range + 4] = $"=(C{range + 4}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";
            sheet.Cells[5][range + 5] = $"=(C{range + 5}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";
            sheet.Cells[5][range + 6] = $"=(C{range + 6}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";
            sheet.Cells[5][range + 7] = $"=(C{range + 7}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";





            (sheet.Cells[3][range + 3] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 4] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 5] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 6] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 7] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 8] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 3] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 4] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 5] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 6] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 7] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 8] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 9] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 10] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 11] as Excel.Range).NumberFormat = "0.0";






            if (Subject[subb].Length > 30)
            {
                Subject[subb] = Subject[subb].Remove(30, Subject[subb].Length - 30);
            }



            xlWorkBook.SaveAs(Path + "/Экзамен" + "/Экзамен " + Reg[subb] + " " + Subject[subb] + " " + GroupsForVed[id] + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//
            xlWorkBook.Close(true, Path + "/Экзамен" + "/Экзамен " + Reg[subb] + " " + Subject[subb] + " " + GroupsForVed[id] + ".xlsx", Type.Missing);



            ////Excel.Application app = null;
            ////Excel.Workbook workbook = null;
            ////Excel.Worksheet worksheet = null;
            //try
            //{
            //    //    app = new Excel.Application();
            //    //    app.Visible = false;
            //    //    workbook = app.Workbooks.Add(1);
            //    //    worksheet = workbook.Sheets[1];

            //    //    worksheet.Cells[4 + GroupsForVed.Count + 2][rowid].Value2 = consT;


            //    workbook.SaveAs(Path + TBName.Text + id + @"xlsx");
            //    workbook.Close();
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show("" + e);
            //    workbook.Close();
            //}

        }
        public void CreateDoc(int id, int subb)
        {//Метод создающий документ
            DirectoryInfo dirInfo = new DirectoryInfo(Path + @"//Зачет//" + GroupsForVed[id]);
            if (!dirInfo.Exists)
            {
                Directory.CreateDirectory(Path + @"//Зачет//" + GroupsForVed[id]);
            }

            DirectoryInfo di1;
            di1 = new DirectoryInfo(@"..\..\..");
            string di2 = di1.FullName + @"\Test1.xlsx";

            Excel.Workbook xlWorkBook;
            Excel.Application xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(di2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true); // 1 секунда
            Excel.Worksheet sheet = xlWorkBook.Sheets[1];
            sheet = xlWorkBook.Sheets[1];
            Excel.Range _excelCells = (Excel.Range)sheet.get_Range($"B14", $"G15").Cells;
            _excelCells.Merge(Type.Missing);
            Excel.Range _excelCells3 = (Excel.Range)sheet.get_Range($"H14", $"I15").Cells;
            _excelCells3.Merge(Type.Missing);
            Excel.Range _excelCells4 = (Excel.Range)sheet.get_Range($"A14", $"A15").Cells;
            _excelCells4.Merge(Type.Missing);
            sheet.Cells[1][14] = "№ п/п";
            sheet.Cells[2][14] = "Ф.И.О. Студента";
            sheet.Cells[8][14] = "Оценка";
            int range = 15 + studentList.Count;


            for (int i = 16; i < 16 + studentList.Count; i++)
            {
                sheet.Cells[1][i] = i - 15;
                Excel.Range _excelCells1 = (Excel.Range)sheet.get_Range($"B{i}", $"G{i}").Cells;
                _excelCells1.Merge(Type.Missing);
                //Excel.Range _excelCells2 = (Excel.Range)sheet.get_Range($"H{i}", $"I{i}").Cells;
                //_excelCells2.Merge(Type.Missing);
                sheet.Cells[2][i] = studentList[i - 16];
                sheet.Cells[9][i].Formula = $"=IF(H{i}=5,\"(отлично)\",IF(H{i}=4,\"(хорошо)\",IF(H{i}=3,\"(удовлетворительно)\",IF(H{i}=2,\"(неудовлетворительно)\",))))";

            }

            sheet.Cells[4][7] = Reg[subb] + " " + Subject[subb];
            sheet.Cells[2][8] = GroupsForVed[id];
            sheet.Cells[3][9] = spec;
            if (id == 0)
            {
                sheet.Cells[4][10] = NameFio1Group[subb];
            }
            else
            {
                sheet.Cells[4][10] = NameFio2Group[subb];
            }

            Excel.Range _excelCellss = (Excel.Range)sheet.get_Range($"A14", $"I" + range).Cells;

            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные            
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
            _excelCellss.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            Excel.Range _excelCells8 = (Excel.Range)sheet.get_Range($"A{range + 2}", $"E{range + 2}").Cells;
            _excelCells8.Merge(Type.Missing);
            sheet.Cells[1][range + 2] = "Подпись преподавателей _____________________________";
            sheet.Cells[1][range + 3] = "5 «отлично»";
            sheet.Cells[1][range + 4] = "4 «хорошо»";
            sheet.Cells[1][range + 5] = "3 «удовлетворительно»";
            sheet.Cells[1][range + 6] = "2 «неудовлетворительно»";
            sheet.Cells[1][range + 7] = "Качество:";
            for (int i = range + 2; i < range + 8; i++)
            {
                Excel.Range _excelCells7 = (Excel.Range)sheet.get_Range($"A{i}", $"B{i}").Cells;
                _excelCells7.Merge(Type.Missing);
            }
            sheet.Cells[4][range + 3] = "чел.";
            sheet.Cells[4][range + 4] = "чел.";
            sheet.Cells[4][range + 5] = "чел.";
            sheet.Cells[4][range + 6] = "чел.";
            sheet.Cells[4][range + 7] = "чел.";

            sheet.Cells[6][range + 3] = "%";
            sheet.Cells[6][range + 4] = "%";
            sheet.Cells[6][range + 5] = "%";
            sheet.Cells[6][range + 6] = "%";
            sheet.Cells[6][range + 7] = "%";

            sheet.Cells[1][range + 8] = "Средний балл";
            sheet.Cells[1][range + 9] = "«___» _______ 202_г";

            sheet.Cells[3][range + 3].Formula = $"=COUNTIF(H16:H100,\"5\")";
            sheet.Cells[3][range + 4].Formula = $"=COUNTIF(H16:H100,\"4\")";
            sheet.Cells[3][range + 5].Formula = $"=COUNTIF(H16:H100,\"3\")";
            sheet.Cells[3][range + 6].Formula = $"=COUNTIF(H16:H100,\"2\")";
            sheet.Cells[3][range + 7].Formula = $"=C{range + 3}+C{range + 4}";
            sheet.Cells[3][range + 8].Formula = $"=(C{range + 3}*5+C{range + 4}*4+C{range + 5}*3+C{range + 6}*2)/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6})";

            sheet.Cells[5][range + 3].Formula = $"=(C{range + 3}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";
            sheet.Cells[5][range + 4].Formula = $"=(C{range + 4}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";
            sheet.Cells[5][range + 5].Formula = $"=(C{range + 5}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";
            sheet.Cells[5][range + 6].Formula = $"=(C{range + 6}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";
            sheet.Cells[5][range + 7].Formula = $"=(C{range + 7}/(C{range + 3}+C{range + 4}+C{range + 5}+C{range + 6}))*100";





            (sheet.Cells[3][range + 3] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 4] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 5] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 6] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 7] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[3][range + 8] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 3] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 4] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 5] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 6] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 7] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 8] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 9] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 10] as Excel.Range).NumberFormat = "0.0";
            (sheet.Cells[5][range + 11] as Excel.Range).NumberFormat = "0.0";



            if (Subject[subb].Length > 30)
            {
                Subject[subb] = Subject[subb].Remove(30, Subject[subb].Length - 30);
            }



            xlWorkBook.SaveAs(Path + "/Зачет" + "/Зачет " + Reg[subb] + " " + Subject[subb] + " " + GroupsForVed[id] + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//
            xlWorkBook.Close(true, Path + "/Зачет" + "/Зачет " + Reg[subb] + " " + Subject[subb] + " " + GroupsForVed[id] + ".xlsx", Type.Missing);



            ////Excel.Application app = null;
            ////Excel.Workbook workbook = null;
            ////Excel.Worksheet worksheet = null;
            //try
            //{
            //    //    app = new Excel.Application();
            //    //    app.Visible = false;              G:\Kursach 2002\CreateDoxVed/Зачет/Зачет ОГСЭ.06 Основы духовно-нравственной культуры народов России 1101.xlsx
            //    //    workbook = app.Workbooks.Add(1);  G:\Kursach 2002\CreateDoxVed/Зачет/Зачет МДК.02.01 Разработка технологических процессов для сборки узлов и изделий в механосборочном производстве, в том числе в Технологический процесс и технологическая документация по сборке узлов и изделий с применением систем автоматизированного проектирования 1591.xlsx
            //    //    worksheet = workbook.Sheets[1];

            //    //    worksheet.Cells[4 + GroupsForVed.Count + 2][rowid].Value2 = consT;


            //    workbook.SaveAs(Path + TBName.Text + id + @"xlsx");
            //    workbook.Close();
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show("" + e);
            //    workbook.Close();
            //}

        }
        public class ExsamVed
        {
            public double Numer { get; set; }
            public string NameFio1Group { get; set; }
            public string NameFio2Group { get; set; }
            public string Subject { get; set; }
            public string Reg { get; set; }
        }
        public static List<ExsamVed> examVed = new List<ExsamVed>(); //Путь к списку
        public static List<string> GroupsForVed = new List<string>();
        public static string spec;
        public static void ReadAllVed(string path)
        {
            string[] groups = new string[2];
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            xlWorkBook = xlApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); // 1 секунда
            Excel.Worksheet sheet = (Excel.Worksheet)xlApp.Worksheets.get_Item(1);
            string dataAll = "";
            int indexOfall = 0;


            for (int i = 0; i < 2; i++)
            {
                if (sheet.Cells[i + 2][3].Text != "" && sheet.Cells[i + 2][3].Text != null)
                {
                    groups[i] = sheet.Cells[i + 2][3].Text;
                    GroupsForVed.Add(sheet.Cells[i + 2][3].Text);
                }
            }
            spec = sheet.Cells[11][1].Value2;

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
            for (int i = 9; i < indexOfall; i++)
            {
                string a = Convert.ToString(sheet.Cells[16][i].Value2);
                int num;
                bool isNum = int.TryParse(Convert.ToString(sheet.Cells[16][i].Value2), out num);
                if (isNum)
                {
                    Numer.Add(num);
                    NameFio1Group.Add(sheet.Cells[2][i].Value2);
                    NameFio2Group.Add(sheet.Cells[3][i].Value2);
                    Subject.Add(sheet.Cells[11][i].Value2);
                    Reg.Add(sheet.Cells[10][i].Value2);
                    examVed.Add(new ExsamVed { Numer = Convert.ToDouble(sheet.Cells[16][i].Value2), NameFio1Group = sheet.Cells[2][i].Value2, NameFio2Group = sheet.Cells[3][i].Value2, Subject = sheet.Cells[11][i].Value2, Reg = sheet.Cells[10][i].Value2, });
                }
            }
            xlWorkBook.Close();
        }


        public static List<string> studentList = new List<string>();
        public static void ReadList(string path, int num)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            xlWorkBook = xlApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); // 1 секунда
            Excel.Worksheet sheet = (Excel.Worksheet)xlApp.Worksheets.get_Item(1);
            string dataAll = "";
            int indexOfall = 1;
            do
            {
                studentList.Add(sheet.Cells[1][indexOfall].Value2);
                dataAll = sheet.Cells[1][indexOfall].Value2;
                indexOfall += 1;
            }
            while (dataAll != null && dataAll != "");

            studentList.RemoveAt(studentList.Count - 1);

            xlWorkBook.Close();
        }


        private void IOpen_Click(object sender, RoutedEventArgs e)
        {

        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            foo();
        }
    }
}
