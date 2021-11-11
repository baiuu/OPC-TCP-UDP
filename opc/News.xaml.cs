using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using OPCAutomation;
using System.Data;
using System.Data.OleDb;

namespace opc
{
    /// <summary>
    /// Window1.xaml 的交互逻辑
    /// </summary>
    /// 
    public delegate void SendMessage1(string str);
    public partial class Window1 : Window
    {
        public OPCServer Server;
        public OPCGroups Groups;
        public OPCGroup Group;
        public OPCItem Item;
        public OPCItems Items;
        public string apppath;
        public OleDbConnection cn;

        public SendMessage1 sendMessage1;
        public class ListItem
        {
            public ListItem(int ID, string StationName)
            {
                this.ID = ID;
                this.StationName = StationName;
            }
            public int ID { get; set; }
            public string StationName { get; set; }
        }
        public class OPCitem
        {
            public OPCitem(int ID, string OPCName)
            {
                this.ID = ID;
                this.OPCName = OPCName;
            }
            public int ID { get; set; }
            public string OPCName { get; set; }
        }
        private async void Getlist()
        {
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    //Dictionary<int, string> dictitem = new Dictionary<int, string>();
                    List<ListItem> dictitem = new List<ListItem>();
                    string strComm = "select * from Station";
                    OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                    OleDbDataAdapter oleDbDataAda = new OleDbDataAdapter(oleDbCmd);
                    DataSet daSet = new DataSet();
                    oleDbDataAda.Fill(daSet);
                    foreach (DataRow r in daSet.Tables[0].Rows)
                    {
                        dictitem.Add(new ListItem(Convert.ToInt16(r["ID"]), r["StationName"].ToString()));
                        if (combo.Items.IndexOf(r["ServerName"]) == -1)
                            this.combo.Dispatcher.Invoke(new Action(() => { combo.Items.Add(r["ServerName"]); }));
                    }
                    this.list.Dispatcher.Invoke(new Action(() => { list.ItemsSource = null; list.ItemsSource = dictitem; list.DisplayMemberPath = "StationName"; list.SelectedValuePath = "ID"; list.SelectedIndex = -1; }));
                }
                catch
                {
                    System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
                }
            });
        }
        private void Dellist(int ID)
        {
            try
            {
                string strComm = "DELETE from Station where ID=" + ID;
                OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                oleDbCmd.ExecuteNonQuery();
            }
            catch
            {
                System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
            }

        }

        private async void Getopc()
        {
            await Task.Factory.StartNew(() =>
            {
                List<OPCitem> dictitem = new List<OPCitem>();
                Server = new OPCServer();
                object serverList = Server.GetOPCServers("127.0.0.1");
                foreach (string test2 in (Array)serverList)
                {
                    if(combo.Items.IndexOf(test2)==-1)
                        this.combo.Dispatcher.Invoke(new Action(() => { combo.Items.Add(test2); }));
                }
            });
        }
        private void Addllist()
        {
            try
            {
                string chek;
                if (Istrue.IsChecked == true)
                {
                    chek = "True";
                }
                else
                {
                    chek = "False";
                }
                string strComm = "INSERT INTO Station  (StationName, [Interval], StationNo, ServerName, IsSend, ComputerName) VALUES ('" + stname.Text + "',60,'" + st.Text + "','" + combo.Text + "'," + chek + ",'127.0.0.1')";
                OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                oleDbCmd.ExecuteNonQuery();
                Getlist();
            }
            catch
            {
                System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
            }

        }
        private void Update(int ID)
        {
            try
            {
                string chek;
                if (Istrue.IsChecked == true)
                {
                    chek = "True";
                }
                else
                {
                    chek = "False";
                }
                string strComm = "UPDATE Station SET StationName = '" + stname.Text + "', StationNo = '" + st.Text + "', ServerName = '" + combo.Text + "', IsSend = " + chek + " WHERE ID =" + ID;
                OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                oleDbCmd.ExecuteNonQuery();
                Getlist();
            }
            catch
            {
                System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
            }

        }
        private async void Getitem(int ID)
        {
            await Task.Factory.StartNew(() =>
            {
                try
                {
                    string strComm = "select * from Station WHERE ID=" + ID;
                    OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                    OleDbDataAdapter oleDbDataAda = new OleDbDataAdapter(oleDbCmd);
                    DataSet daSet = new DataSet();
                    oleDbDataAda.Fill(daSet);
                    foreach (DataRow r in daSet.Tables[0].Rows)
                    {
                        if (combo.Items.IndexOf(r["ServerName"]) == -1)
                            this.combo.Dispatcher.Invoke(new Action(() => { combo.Items.Add(r["ServerName"]); }));
                        this.combo.Dispatcher.Invoke(new Action(() => { combo.SelectedIndex = combo.Items.IndexOf(r["ServerName"]); }));
                        this.st.Dispatcher.Invoke(new Action(() => { st.Text = r["StationNo"].ToString(); }));
                        this.stname.Dispatcher.Invoke(new Action(() => { stname.Text = r["StationName"].ToString(); }));
                        if (r["IsSend"].ToString() == "True")
                        {
                            this.Istrue.Dispatcher.Invoke(new Action(() => { this.Istrue.IsChecked = true; }));
                        }
                        else
                        {
                            this.Istrue.Dispatcher.Invoke(new Action(() => { this.Istrue.IsChecked = false; }));
                        }
                    }
                }
                catch
                {
                    System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
                }
                
            });
        }

        public Window1()
        {
            InitializeComponent();
            Getlist();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Getopc();
            list.SelectedIndex = -1;
            st.IsEnabled = true;
            stname.IsEnabled = true;
            Istrue.IsEnabled = true;
            cl.IsEnabled = true;
            check.IsEnabled = true;
            combo.IsEnabled = true;
            st.Text = "";
            stname.Text = "";
            Istrue.IsChecked = false;
            combo.Items.Clear();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Dellist(Convert.ToInt32(list.SelectedValue));
            this.combo.Dispatcher.Invoke(new Action(() => { combo.Items.RemoveAt(combo.SelectedIndex); }));
            Getlist();
        }


        private void List_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (list.SelectedIndex != -1)
            {
                Getitem(Convert.ToInt32(list.SelectedValue));
                st.IsEnabled = true;
                stname.IsEnabled = true;
                Istrue.IsEnabled = true;
                cl.IsEnabled = true;
                check.IsEnabled = true;
                combo.IsEnabled = true;
                del.IsEnabled = true;
            }
            else
            {
                st.IsEnabled = false;
                stname.IsEnabled = false;
                Istrue.IsEnabled = false;
                cl.IsEnabled = false;
                check.IsEnabled = false;
                combo.IsEnabled = false;
                del.IsEnabled = false;
                st.Text = "";
                stname.Text = "";
                Istrue.IsChecked = false;
                combo.SelectedIndex = -1;
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (list.SelectedIndex != -1)
                Getitem(Convert.ToInt32(list.SelectedValue));
            else
            {
                st.Text = "";
                stname.Text = "";
                Istrue.IsChecked = false;
            }

        }

        private void check_Click(object sender, RoutedEventArgs e)
        {
            if (list.SelectedIndex != -1)
            {
                Update(Convert.ToInt32(list.SelectedValue));
            }
            else
            {
                Addllist();
            }
        }
        private void Window_Closed(object sender, EventArgs e)
        {
            sendMessage1("2");
        }
    }
}


