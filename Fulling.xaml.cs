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
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Page = System.Windows.Controls.Page;
using System.Net.NetworkInformation;
using System.Threading;
using System.Diagnostics;
using Application = Microsoft.Office.Interop.Excel.Application;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace SerpCollPoj
{
    /// <summary>
    /// Логика взаимодействия для Fulling.xaml
    /// </summary>
    public partial class Fulling : Page
    {
        public List<string> paths = new List<string>();
        public string files;
        public Fulling()
        {
            InitializeComponent();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string filename in openFileDialog.FileNames)
                    paths.Add(filename);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            files = dialog.FileName;
        }
        public void bb()
        {


        }
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWbSource, xlWbTarget;
            xlWbTarget = xlApp.Workbooks.Add();
            for (int i = 0; i < paths.Count; i++)
            {
                xlWbSource = xlApp.Workbooks.Open(paths[i]);
                //Новая книга
                //Вставка первого листа из книги xlWbSource перед первым листом книги xlWbTarget
                (xlWbSource.Worksheets[1]).Copy(xlWbTarget.Worksheets[1]);
                xlApp.Visible = false;
                xlWbSource.Close(false);
            }



            xlWbTarget.SaveAs(files + @"\" + tbname.Text + @".xlsx");
            xlWbTarget.Close(true);
            xlApp.Quit();
            System.Windows.MessageBox.Show("rrrfrf");





         
        }

    }
}
