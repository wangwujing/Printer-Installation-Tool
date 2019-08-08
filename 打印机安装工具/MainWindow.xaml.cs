using System;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.ServiceProcess;
using System.Diagnostics;
using Microsoft.Win32;
using System.IO;
using System.Runtime.InteropServices;

namespace 打印机安装工具
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        string server_IP = @"\\10.40.1.160\";
        public MainWindow()
        {
            File.Delete(Environment.CurrentDirectory + @"\打印机安装工具.old");
            FileVersionInfo file1 = FileVersionInfo.GetVersionInfo(server_IP + @"Printer tools\打印机安装工具");
            if (file1.FileMajorPart != 3)//判断版本号，注意！！！！！
            {
                MessageBoxResult Updata = MessageBox.Show("检查到新版本，是否现在更新？", "有新版本！", MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (Updata == MessageBoxResult.Yes)
                {
                    string path = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                    try { 
                    File.Move(path, Environment.CurrentDirectory + @"\打印机安装工具.old");
                    File.Copy(server_IP + @"Printer tools\打印机安装工具", path);
                    }
                    catch
                    {
                        MessageBox.Show("更新失败！当前程序运行路径没有写入权限！","Error",MessageBoxButton.OK,MessageBoxImage.Error);
                        Environment.Exit(0);
                    }
                    DelayRun(path, 3);//延迟执行
                    this.Close();
                }
                else
                {
                    this.Close();
                }
            }
            InitializeComponent();
        }
        public static void DelayRun(string path, int time) //延迟运行
        {
            ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", "/C ping 2.0.0.1 -n " + time + " -w 1000 > Nul & \"" + path + "\"");
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            psi.CreateNoWindow = true;
            Process.Start(psi);
        }
        private void Window_ContentRendered(object sender, EventArgs e)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", @"/C sc config spooler start= auto & net start spooler & pause");
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                psi.CreateNoWindow = true;
                Process.Start(psi);
            }
            catch { };
            Thread thread1 = new Thread(new ThreadStart(Thread1));
            thread1.SetApartmentState(ApartmentState.STA);
            thread1.IsBackground = true;
            thread1.Start();
        }
        private void Thread1()
        {
            string Print_Name="";
            string ip = GetClientLocalIPv4Address();
            string[] ip_item = Regex.Split(ip, "\\.");
            int ip_2= Convert.ToInt32(ip_item[1]);
            switch (ip_2)
            {
                case 31:
                    Print_Name = "3栋1楼打印机";
                    break;
                case 32:
                    Print_Name = "3栋2楼打印机";
                    break;
                case 33:
                    Print_Name = "3栋3楼打印机";
                    break;
                case 34:
                    Print_Name = "3栋4楼打印机";
                    break;
                case 35:
                    Print_Name = "3栋5楼打印机";
                    break;
                case 36:
                    Print_Name = "3栋6楼打印机";
                    break;
                case 37:
                    Print_Name = "3栋7楼打印机";
                    break;
                case 38:
                    Print_Name = "3栋8楼打印机";
                    break;
                case 27:
                    Print_Name = "2栋7楼打印机";
                    break;
                case 28:
                    Print_Name = "2栋8楼打印机";
                    break;
                case 29:
                    Print_Name = "2栋9楼打印机";
                    break;
                case 23:
                    Print_Name = "2栋13楼打印机";
                    break;
                case 24:
                    Print_Name = "2栋14楼打印机";
                    break;
                case 17:
                    Print_Name = "2栋17楼打印机";
                    break;
                case 18:
                    Print_Name = "2栋18楼打印机";
                    break;
                case 19:
                    Print_Name = "2栋19楼打印机";
                    break;
            }
            if (ip_2==38)
            {
                Window1 win = new Window1();                 
                win.ShowDialog();
                string Print_Name1;
                switch (win.win1)
                {
                    case 1:
                        Print_Name = "3栋8楼人事部打印机";
                        Print_Name1 = "3栋8楼人事部打印机（备用）";
                        Print_Name = server_IP + Print_Name;
                        Print_Name1 = server_IP + Print_Name1;
                        bool P1=IsPrinterInstalled(Print_Name);
                        bool P2=IsPrinterInstalled(Print_Name1);
                        bool Add_p=false;
                        bool Add_p1=false;
                        if (!(P1&&P2))
                        {
                            if (!P1)
                            {
                                Add_p = AddPrinter(Print_Name);
                            }
                            if (!P2)
                            {
                                Add_p1 = AddPrinter(Print_Name1);
                            }
                            if (Add_p||Add_p1)
                            {
                                if (IsPrinterInstalled(Print_Name))
                                {
                                    SetDefaultPrinter(Print_Name);
                                }
                                else
                                {
                                    SetDefaultPrinter(Print_Name1);
                                }                        
                                Action action1 = () =>
                                {
                                    this.Hide();
                                    MessageBox.Show("您的打印机已经安装成功！", "success", MessageBoxButton.OK, MessageBoxImage.Information);
                                    this.Close();
                                };
                                this.Dispatcher.BeginInvoke(action1);
                            }
                            else
                            {
                                Action action1 = () =>
                                {
                                    this.Hide();
                                    MessageBox.Show("添加打印机失败！", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                                    this.Close();
                                };
                                this.Dispatcher.BeginInvoke(action1);
                            }
                        }
                        else
                        {
                            Action action1 = () =>
                            {
                                this.Hide();
                                MessageBox.Show("您已安装对应楼层打印机！", "", MessageBoxButton.OK, MessageBoxImage.Warning);
                                this.Close();
                            };
                            this.Dispatcher.BeginInvoke(action1);
                        }
                        break;
                    case 2:
                        Print_Name = "3栋8楼财务一部打印机";
                        Print_Name1 = "3栋8楼财务一部打印机（备用）";
                        Print_Name = server_IP + Print_Name;
                        Print_Name1 = server_IP + Print_Name1;
                        bool P1_2 = IsPrinterInstalled(Print_Name);
                        bool P2_2 = IsPrinterInstalled(Print_Name1);
                        bool Add_p_2 = false;
                        bool Add_p1_2 = false;
                        if (!(P1_2 && P2_2))
                        {
                            if (!P1_2)
                            {
                                Add_p_2 = AddPrinter(Print_Name);
                            }
                            if (!P2_2)
                            {
                                Add_p1_2 = AddPrinter(Print_Name1);
                            }
                            if (Add_p_2 || Add_p1_2)
                            {
                                if (IsPrinterInstalled(Print_Name))
                                {
                                    SetDefaultPrinter(Print_Name);
                                }
                                else
                                {
                                    SetDefaultPrinter(Print_Name1);
                                }
                                Action action1 = () =>
                                {
                                    this.Hide();
                                    MessageBox.Show("您的打印机已经安装成功！", "success", MessageBoxButton.OK, MessageBoxImage.Information);
                                    this.Close();
                                };
                                this.Dispatcher.BeginInvoke(action1);
                            }
                            else
                            {
                                Action action1 = () =>
                                {
                                    this.Hide();
                                    MessageBox.Show("添加打印机失败！", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                                    this.Close();
                                };
                                this.Dispatcher.BeginInvoke(action1);
                            }
                        }
                        else
                        {
                            Action action1 = () =>
                            {
                                this.Hide();
                                MessageBox.Show("您已安装对应楼层打印机！", "", MessageBoxButton.OK, MessageBoxImage.Warning);
                                this.Close();
                            };
                            this.Dispatcher.BeginInvoke(action1);
                        }
                        break;
                    case 3:
                        Print_Name = "3栋8楼财务二部打印机";
                        AddPrint(Print_Name);
                        break;
                    default:
                        System.Environment.Exit(0);
                        break;
                }
            }
            else
            {
                AddPrint(Print_Name);
            }
                 
        }
        public void AddPrint(string Print_Name)
        {
            Print_Name = server_IP + Print_Name;
            if (!IsPrinterInstalled(Print_Name))
            {
                if (AddPrinter(Print_Name))
                {
                     SetDefaultPrinter(Print_Name);                 
                    Action action1 = () =>
                    {
                        this.Hide();
                        MessageBox.Show("您的打印机已经安装成功！", "success", MessageBoxButton.OK, MessageBoxImage.Information);
                        this.Close();
                    };
                    this.Dispatcher.BeginInvoke(action1);
                }
                else
                {
                    Action action1 = () =>
                    {
                        this.Hide();
                        MessageBox.Show("添加打印机失败！", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                        this.Close();
                    };
                    this.Dispatcher.BeginInvoke(action1);
                }
            }
            else
            {
                Action action1 = () =>
                {
                    this.Hide();
                    MessageBox.Show("您已安装对应楼层打印机！", "", MessageBoxButton.OK, MessageBoxImage.Warning);
                    this.Close();
                };
                this.Dispatcher.BeginInvoke(action1);
            }
        }
        //public static string GetClientLocalIPv4Address()//获取IPV4的地址
        //{
        //    string strLocalIP = string.Empty;
        //    try
        //    {
        //        IPHostEntry ipHost = Dns.GetHostEntry(Dns.GetHostName());
        //        for (int i = 0; i < ipHost.AddressList.Length; i++)
        //        {
        //            if (ipHost.AddressList[i].AddressFamily == AddressFamily.InterNetwork)
        //            {
        //                strLocalIP = ipHost.AddressList[i].ToString();
        //            }
        //        }
        //        return strLocalIP;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("获取IPV4的地址出错，信息：" + ex.Message);
        //        return "";
        //    }
        //}
        public static string GetClientLocalIPv4Address()//获取IPV4的地址,排除不是10开头的
        {
            string strLocalIP = string.Empty;
            try
            {
                IPHostEntry ipHost = Dns.GetHostEntry(Dns.GetHostName());
                for (int i = 0; i < ipHost.AddressList.Length; i++)
                {
                    if (ipHost.AddressList[i].AddressFamily == AddressFamily.InterNetwork)
                    {
                        strLocalIP = ipHost.AddressList[i].ToString();
                        string[] ip_item = Regex.Split(strLocalIP, "\\.");
                        if (ip_item[0] == "10")//判断IP第一位是否为10，直接排除虚拟机等多余网卡获取到的地址。
                        {
                            break;//跳出FOR循环，执行弹窗，返回IP参数。
                        }
                    }
                }
                return strLocalIP;
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取IPV4的地址出错，信息：" + ex.Message);
                return "";
            }
        }
        private static ManagementScope oManagementScope = null;//全局变量，添加打印机，需要注意！易出BUG。
        public static bool AddPrinter(string sPrinterName)//添加打印机
        {
            try
            {
                oManagementScope = new ManagementScope(ManagementPath.DefaultPath);
                oManagementScope.Connect();

                ManagementClass oPrinterClass = new ManagementClass
                               (new ManagementPath("Win32_Printer"), null);
                ManagementBaseObject oInputParameters =
                   oPrinterClass.GetMethodParameters("AddPrinterConnection");
                oInputParameters.SetPropertyValue("Name", sPrinterName);
                oPrinterClass.InvokeMethod("AddPrinterConnection", oInputParameters, null);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("安装打印机出错，信息：" + ex.Message);
                return false;
            }
        }
        public static void SetDefaultPrinter(string sPrinterName)//将打印机设为默认值
        {
            oManagementScope = new ManagementScope(ManagementPath.DefaultPath);
            oManagementScope.Connect();

            SelectQuery oSelectQuery = new SelectQuery();
            oSelectQuery.QueryString = @"SELECT * FROM Win32_Printer WHERE Name = 
			'" + sPrinterName.Replace("\\", "\\\\") + "'";

            ManagementObjectSearcher oObjectSearcher =
               new ManagementObjectSearcher(oManagementScope, oSelectQuery);
            ManagementObjectCollection oObjectCollection = oObjectSearcher.Get();

            if (oObjectCollection.Count != 0)
            {
                foreach (ManagementObject oItem in oObjectCollection)
                {
                    oItem.InvokeMethod("SetDefaultPrinter", new object[] { sPrinterName });
                    return;
                }
            }
        }
        public static bool IsPrinterInstalled(string sPrinterName)//检测打印机是否安装
        {
            oManagementScope = new ManagementScope(ManagementPath.DefaultPath);
            oManagementScope.Connect();

            SelectQuery oSelectQuery = new SelectQuery();
            oSelectQuery.QueryString = @"SELECT * FROM Win32_Printer WHERE Name = '" +
                           sPrinterName.Replace("\\", "\\\\") + "'";

            ManagementObjectSearcher oObjectSearcher =
               new ManagementObjectSearcher(oManagementScope, oSelectQuery);
            ManagementObjectCollection oObjectCollection = oObjectSearcher.Get();

            return oObjectCollection.Count > 0;
        }

        private void Label_MouseEnter(object sender, MouseEventArgs e)
        {
            EXIT.Foreground= new SolidColorBrush(Colors.Red);
        }

        private void Label_MouseLeave(object sender, MouseEventArgs e)
        {
            EXIT.Foreground = new SolidColorBrush(Colors.Black);
        }

        private void EXIT_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Application.Current.Shutdown();
        }       
    }
}

