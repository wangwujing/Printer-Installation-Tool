using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace 打印机安装工具
{
    /// <summary>
    /// Window1.xaml 的交互逻辑
    /// </summary>
    public partial class Window1 : Window
    {
        public int win1 { get; set; }
        public Window1()
        {
            InitializeComponent();          
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            win1 = 1;
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            win1 = 2;
            this.Close();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            win1 = 3;
            this.Close();
        }
    }
}
