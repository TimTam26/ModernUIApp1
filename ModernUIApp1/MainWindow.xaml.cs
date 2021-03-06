using FirstFloor.ModernUI.Windows.Controls;
using System;
using System.Windows;
using static ModernUIApp1.MainService;

namespace ModernUIApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : ModernWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Execute();

            resultMessage.Text = DateTime.Now.ToString() + " Executed";
        }

        private void Button_Click_Closing(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
