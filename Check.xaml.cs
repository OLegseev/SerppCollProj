using System;
using System.Collections.Generic;

using System.Linq;
using System.Text;
using System.Threading;
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

namespace SerpCollPoj
{
    /// <summary>
    /// Логика взаимодействия для Check.xaml
    /// </summary>
    public partial class Check : Page
    {
        public Check()
        {
            InitializeComponent();
            b1.Visibility = Visibility.Hidden;
            b2.Visibility = Visibility.Hidden;
            b3.Visibility = Visibility.Hidden;
            
            b1.Visibility = Visibility.Visible;
            
            b2.Visibility = Visibility.Visible;
            
            b3.Visibility = Visibility.Visible;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            NavigationService.Navigate(new Year());
        }

        private void Button_Click1(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Fulling());
        }

        private void Button_Click2(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new YearTeacher());
        }

        private void Button_Click21(object sender, RoutedEventArgs e)
        {

            NavigationService.Navigate(new YearEx1());
        }

        private void Button_Click22(object sender, RoutedEventArgs e)
        {

            NavigationService.Navigate(new YearEx2());
        }

        private void Button_Click23(object sender, RoutedEventArgs e)
        {

            NavigationService.Navigate(new YearEx3());
        }

        private void Button_Click24(object sender, RoutedEventArgs e)
        {

            NavigationService.Navigate(new YearEx4());
        }

        private void Button_Click25(object sender, RoutedEventArgs e)
        {

            NavigationService.Navigate(new YearEx5());
        }
    }
}
