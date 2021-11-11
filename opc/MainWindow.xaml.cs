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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Drawing;
using OPCAutomation;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Timers;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;

namespace opc
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public OPCServer Server;
        public OPCGroups Groups;
        public OPCGroup Group;
        public OPCItem Item;
        public OPCItems Items;
        public JObject jsonObject;
        public OleDbConnection cn;
        public FileStream fs1;
        public WindowState ws;
        public WindowState wsl;
        public NotifyIcon notifyIcon;
        public class ListItem
        {
            public ListItem(int ID,string SensorName, string StationName,string ChannelName,string ItemName, string Values, string Time, string StationNo)
            {
                this.ID = ID;
                this.SensorName = SensorName;
                this.StationName = StationName;
                this.ChannelName = ChannelName;
                this.ItemName = ItemName;
                this.Values = Values;
                this.Time = Time;
                this.StationNo = StationNo;
            }
            public int ID { get; set; }
            public string StationName { get; set; }
            public string SensorName { get; set; }
            public string ChannelName { get; set; }
            public string Values { get; set; }
            public string Time { get; set; }
            public string ItemName { get; set; }
            public string StationNo { get; set; }
        }
        public List<ListItem> dictitem = new List<ListItem>();
        public System.Timers.Timer timer;
        public int tesk = 2;
        public string apppath = "";
        private async void Getlist()
        {
            await Task.Factory.StartNew(() =>
            {
                //Dictionary<int, string> dictitem = new Dictionary<int, string>();
                try
                {
                    dictitem.Clear();
                    string strComm = "SELECT * from Station,Sensor where Sensor.StationID=Station.ID";
                    OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                    OleDbDataAdapter oleDbDataAda = new OleDbDataAdapter(oleDbCmd);
                    DataSet daSet = new DataSet();
                    oleDbDataAda.Fill(daSet);
                    foreach (DataRow r in daSet.Tables[0].Rows)
                    {
                        dictitem.Add(new ListItem(Convert.ToInt16(r["Sensor.ID"]), r["SensorName"].ToString(), r["StationName"].ToString(), r["ChannelName"].ToString(), r["ItemName"].ToString(), "未采集", "",r["StationNo"].ToString()));
                    }
                    this.dg.Dispatcher.Invoke(new Action(() => { dg.ItemsSource = null; dg.ItemsSource = dictitem; dg.Columns[0].Header = "序号"; dg.Columns[1].Header = "站点"; dg.Columns[2].Header = "传感器名"; dg.Columns[3].Header = "位号"; dg.Columns[4].Header = "采集值"; dg.Columns[5].Header = "采集时间"; dg.Columns[6].Header = "采集标签"; dg.Columns[7].Header = "站点编号"; }));
                }
                catch
                {
                    System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
                }
                if (tesk==2||tesk==1)
                {
                    tesk = 1;
                    GoOPC();
                    this.stat.Dispatcher.Invoke(new Action(() => { stat.Text = "采集中"; }));
                    Update(Convert.ToInt32(Convert.ToDouble(jsonObject["Time"]) * 1000));
                    Upfs(Convert.ToInt32(Convert.ToDouble(jsonObject["Time"]) * 1000));
                }
            });
        }
        private async void GoOPC()
        {
            await Task.Factory.StartNew(() =>
            {
                int run = 1;
                try
                {
                    Server = new OPCServer();
                }
                catch (Exception ex)
                {
                    run = 0;
                    System.Windows.MessageBox.Show("OPC ERRO: "+ex.ToString());
                    tesk = 0;
                }
                if (run == 1)
                {
                    try
                    {
                        Server.Connect(jsonObject["OPCserver"].ToString());
                        if (Server.ServerState == Convert.ToInt32(OPCServerState.OPCRunning))
                        {
                            Groups = Server.OPCGroups;
                            Groups.DefaultGroupDeadband = 0;
                            Groups.DefaultGroupIsActive = true;
                            Groups.DefaultGroupUpdateRate = 250;
                            Group = Groups.Add("group1");
                            Group.IsSubscribed = true;
                            Items = Group.OPCItems;
                            int a = dg.Items.Count;
                            for (int i = 0; i < a; i++)
                            {
                                try
                                {
                                    Item = Items.AddItem(dictitem[i].ItemName.ToString(), i + 1);
                                }
                                catch
                                {
                                    string str = dictitem[i].ItemName.ToString() + " 标签无法取数。";
                                    str = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + ": " + str + "\r\n";
                                    byte[] t = System.Text.Encoding.UTF8.GetBytes(str);
                                    fs1.Write(t, 0, t.Length);
                                    fs1.Flush();
                                }
                            }
                            string str1 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + ": 开始采集\r\n";
                            byte[] t1 = System.Text.Encoding.UTF8.GetBytes(str1);
                            fs1.Write(t1, 0, t1.Length);
                            fs1.Flush();
                            Group.DataChange += Group_DataChange;
                        }
                        else
                        {
                            System.Windows.MessageBox.Show("OPC server is not connected");
                        }
                    }
                    catch
                    {
                        System.Windows.MessageBox.Show("OPC 服务不存在。");
                        tesk = 0;
                        this.stat.Dispatcher.Invoke(new Action(() => { stat.Text = "错误"; }));
                    }
                }
            });
        }
        private async void Update(int time)
        {
            await Task.Factory.StartNew(() =>
            {
                timer = new System.Timers.Timer();
                System.Threading.Thread.Sleep(1000);
                timer.Enabled = true;
                timer.Elapsed += new ElapsedEventHandler(Getdata);
                timer.Interval = time;
            });
        }
        private async void Upfs(int time)
        {
            await Task.Factory.StartNew(() =>
            {
                System.Timers.Timer timer2 = new System.Timers.Timer();
                System.Threading.Thread.Sleep(2000);
                timer2.Enabled = true;
                timer2.Elapsed += new ElapsedEventHandler(Fs);
                timer2.Interval = time;
            });
        }
        private async void Upui()
        {
            await Task.Factory.StartNew(() =>
            {
                System.Timers.Timer timer1 = new System.Timers.Timer();
                System.Threading.Thread.Sleep(2000);
                timer1.Enabled = true;
                timer1.Elapsed += new ElapsedEventHandler(UI);
                timer1.Interval = 5000;
            });
        }
        private async void Up(string str)
        {
            await Task.Factory.StartNew(() =>
            {
                IPEndPoint ipep = new IPEndPoint(IPAddress.Parse(jsonObject["IP"].ToString()), Convert.ToInt32(jsonObject["Port"]));
                byte[] b = System.Text.Encoding.UTF8.GetBytes(str);
                if (jsonObject["Type"].ToString()=="tcp")
                {
                    TcpClient tc = new TcpClient();
                    tc.Connect(ipep);
                    NetworkStream ns = tc.GetStream();
                    ns.Write(b, 0, b.Length);
                    ns.Close();
                    tc.Close();
                }
                else
                {
                    UdpClient uc = new UdpClient();
                    uc.Send(b, b.Length, ipep);
                    uc.Close();
                }
                str = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + ": " + str + "\r\n";
                byte[] t = System.Text.Encoding.UTF8.GetBytes(str);
                fs1.Write(t, 0, t.Length);
            });
        }
        private void Getdata(object sender, ElapsedEventArgs e)
        {
            if(tesk!=0)
            {
                try
                {
                    var names = dictitem.Select(t => t.StationNo).Distinct<string>().ToList();
                    foreach (string row in names)
                    {
                        string str = "从dictitem中获取数据并传递到UP方法上传数据。";
                        Up(str);
                    }
                }
                catch
                {
                    Console.WriteLine("erro");
                }
            }
        }
        private void UI(object sender, ElapsedEventArgs e)
        {
            if (tesk != 0)
            {
                try
                {
                    this.dg.Dispatcher.Invoke(new Action(() => { this.dg.ItemsSource = null; this.dg.ItemsSource = dictitem; dg.Columns[0].Header = "序号"; dg.Columns[1].Header = "站点"; dg.Columns[2].Header = "传感器名"; dg.Columns[3].Header = "位号"; dg.Columns[4].Header = "采集值"; dg.Columns[5].Header = "采集时间"; dg.Columns[6].Header = "采集标签"; dg.Columns[7].Header = "站点编号"; }));
                }
                catch
                {
                    Console.WriteLine("erro");
                }
            }
        }
        private void Fs(object sender, ElapsedEventArgs e)
        {
            if (tesk != 0)
            {
                try
                {
                    fs1.Flush();
                }
                catch
                {
                    Console.WriteLine("erro");
                }
            }
        }
        private void Group_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            for (int i = 1; i <= NumItems; i++)
            {
                if (tesk == 1)
                {
                    int tmpClientHandle = Convert.ToInt32(ClientHandles.GetValue(i));
                    string tmpValue = ItemValues.GetValue(i).ToString();
                    string tmpTime = ((DateTime)(TimeStamps.GetValue(i))).AddHours(8).ToString();
                    dictitem[tmpClientHandle-1].Values = tmpValue;
                    dictitem[tmpClientHandle-1].Time = tmpTime;
                }
                else
                {
                    Console.WriteLine("wo");
                }
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            icon();
            contextMenu();
            wsl = this.WindowState;
            apppath = Environment.CurrentDirectory;
            string con = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + apppath + "\\db.mdb";
            if (!File.Exists(apppath + "\\config.json"))
            {
                FileStream fs1 = new FileStream(apppath + "\\config.json", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs1);
                sw.WriteLine("{\"OPCserver\":\"\",\"Type\":\"udp\",\"IP\":\"127.0.0.1\",\"Port\":\"80\",\"Time\":\"60\"}");//开始写入值
                sw.Close();
                fs1.Close();
            }
            try
            {
                string file = File.ReadAllText(apppath+"\\config.json");
                jsonObject = (JObject)JsonConvert.DeserializeObject(file);
            }
            catch
            {
                System.Windows.MessageBox.Show("JSON解析失败。");
            }
            try
            {
                if (!File.Exists(apppath + "\\db.mdb"))
                {
                    System.Windows.MessageBox.Show("MDB文件不存在。");
                }
                else
                {
                    try
                    {
                        cn = new OleDbConnection(con);
                        cn.Open();
                    }
                    catch(Exception ex)
                    {
                        System.Windows.MessageBox.Show("erro: " + ex);
                    }
                    Getlist();
                    Upui();
                }
            }
            catch
            {
                System.Windows.MessageBox.Show("MDB文件不存在。");
            }
            string date = "update.log";
            if (!System.IO.File.Exists(apppath + "\\" + date))
            {
                fs1 = new FileStream(apppath + "\\" + date, FileMode.Create, FileAccess.Write);
                System.IO.File.SetAttributes(apppath + "\\" + date, FileAttributes.ReadOnly);
            }
            else
            {
                new FileInfo(apppath + "\\" + date).Attributes = FileAttributes.Normal;
                fs1 = new FileStream(apppath + "\\" + date, FileMode.Append, FileAccess.Write);
            }


        }
        
        private void icon()
        {
            string path = System.IO.Path.GetFullPath(@"Icon\1.ico");
            if (File.Exists(path))
            {
                this.notifyIcon = new NotifyIcon();
                this.notifyIcon.BalloonTipText = "OPC采集";
                this.notifyIcon.Text = "OPC采集";
                System.Drawing.Icon icon = new System.Drawing.Icon(path);
                this.notifyIcon.Icon = icon;
                this.notifyIcon.Visible = true;
                notifyIcon.MouseDoubleClick += onNotifyIconDoubleClick;
            }
        }
        private void contextMenu()
        {
            ContextMenuStrip cms = new ContextMenuStrip();
            notifyIcon.ContextMenuStrip = cms;
            ToolStripMenuItem exitMuItem = new ToolStripMenuItem();
            exitMuItem.Text = "退出";
            exitMuItem.Click += new EventHandler(ExitMuItem_Click);
            ToolStripMenuItem hideMenumItem = new ToolStripMenuItem();
            hideMenumItem.Text = "隐藏";
            hideMenumItem.Click += new EventHandler(HideMenumItem_Click);
            ToolStripMenuItem showMenuItem = new ToolStripMenuItem();
            showMenuItem.Text = "显示";
            showMenuItem.Click += new EventHandler(ShowMenuItem_Click);
            cms.Items.Add(exitMuItem);
            cms.Items.Add(hideMenumItem);
            cms.Items.Add(showMenuItem);
        }

        private void ExitMuItem_Click(object sender, EventArgs e)
        {
            fs1.Flush();
            notifyIcon.Visible = false;
            System.Windows.Application.Current.Shutdown();
        }

        private void HideMenumItem_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void ShowMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
            this.Activate();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            Window1 window1 = new Window1();
            window1.apppath = apppath;
            window1.cn = cn;
            window1.sendMessage1 += ReceivedMessage1;
            window1.ShowDialog();
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            if (tesk != 1)
            {
                GoOPC();
                tesk = 1;
                stat.Text = "采集中";
                Update(Convert.ToInt32(Convert.ToDouble(jsonObject["Time"]) * 1000));
                Upfs(Convert.ToInt32(Convert.ToDouble(jsonObject["Time"]) * 1000));
            }
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                if (tesk == 1)
                {
                    Server.Disconnect();
                    tesk = 0;
                    timer.Stop();
                    stat.Text = "停止中";
                    string str = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff") + ": 采集结束\r\n";
                    byte[] t = System.Text.Encoding.UTF8.GetBytes(str);
                    fs1.Write(t, 0, t.Length);
                    fs1.Flush();
                }
                else
                {
                    System.Windows.MessageBox.Show("OPC 服务未连接。");
                }
            }
            catch
            {
                System.Windows.MessageBox.Show("OPC 服务未连接。");
            }
        }

        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            Window3 window3 = new Window3();
            window3.apppath = apppath;
            window3.cn = cn;
            window3.sendMessage1 += ReceivedMessage1;
            window3.ShowDialog();
        }

        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            Window2 window2 = new Window2(jsonObject);
            window2.sendMessage += ReceivedMessage;
            window2.ShowDialog();
        }
        private void ReceivedMessage(JObject value)
        {
            jsonObject = value;
            string str = JsonConvert.SerializeObject(jsonObject);
            System.IO.File.WriteAllText(apppath + "\\config.json", str, Encoding.UTF8);
        }
        private void ReceivedMessage1(string value)
        {
            Getlist();
            this.dg.Dispatcher.Invoke(new Action(() => { this.dg.ItemsSource = null; this.dg.ItemsSource = dictitem; dg.Columns[0].Header = "序号"; dg.Columns[1].Header = "站点"; dg.Columns[2].Header = "传感器名"; dg.Columns[3].Header = "位号"; dg.Columns[4].Header = "采集值"; dg.Columns[5].Header = "采集时间"; dg.Columns[6].Header = "采集标签"; dg.Columns[7].Header = "站点编号"; }));
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            ws = this.WindowState;
            if (ws == WindowState.Minimized)
            {
                this.Hide();
            }
        }
        private void onNotifyIconDoubleClick(object sender, EventArgs e)
        {
            this.Show();
            WindowState = wsl;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            ws = this.WindowState;
            this.Hide();
        }
    }
}