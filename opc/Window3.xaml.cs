using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
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
using System.Windows.Shapes;

namespace opc
{
    /// <summary>
    /// Window3.xaml 的交互逻辑
    /// </summary>
    /// 
    public partial class Window3 : Window
    {
        public class ListItem
        {
            public ListItem(int ID, string SensorName)
            {

                this.ID = ID;
                this.SensorName = SensorName;
            }
            public int ID { get; set; }
            public string SensorName { get; set; }

        }
        public class ListItem1
        {
            public ListItem1(int ID, string StationName)
            {
                this.ID = ID;
                this.StationName = StationName;
            }
            public int ID { get; set; }
            public string StationName { get; set; }
        }
        public string apppath;
        public SendMessage1 sendMessage1;
        public OleDbConnection cn;
        private async void Getlist()
        {
            await Task.Factory.StartNew(() =>
            {
                //Dictionary<int, string> dictitem = new Dictionary<int, string>();
                try
                {
                    List<ListItem> dictitem = new List<ListItem>();
                    string strComm = "SELECT * from Sensor";
                    OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                    OleDbDataAdapter oleDbDataAda = new OleDbDataAdapter(oleDbCmd);
                    DataSet daSet = new DataSet();
                    oleDbDataAda.Fill(daSet);
                    foreach (DataRow r in daSet.Tables[0].Rows)
                    {
                        dictitem.Add(new ListItem(Convert.ToInt16(r["ID"]), r["SensorName"].ToString()));
                    }
                    this.list.Dispatcher.Invoke(new Action(() => { list.ItemsSource = dictitem; list.DisplayMemberPath = "SensorName"; list.SelectedValuePath = "ID"; list.SelectedIndex = -1; }));
                    strComm = "SELECT * from Station";
                    List<ListItem1> dictitem1 = new List<ListItem1>();
                    oleDbCmd = new OleDbCommand(strComm, cn);
                    oleDbDataAda = new OleDbDataAdapter(oleDbCmd);
                    daSet = new DataSet();
                    oleDbDataAda.Fill(daSet);
                    foreach (DataRow r in daSet.Tables[0].Rows)
                    {
                        dictitem1.Add(new ListItem1(Convert.ToInt16(r["ID"]), r["StationName"].ToString()));
                    }
                    this.combo.Dispatcher.Invoke(new Action(() => { combo.ItemsSource = dictitem1; combo.DisplayMemberPath = "StationName"; combo.SelectedValuePath = "ID"; combo.SelectedIndex = -1; }));
                }
                catch
                {
                    System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
                }

            });
        }
        private async void Getlist1()
        {
            await Task.Factory.StartNew(() =>
            {
                //Dictionary<int, string> dictitem = new Dictionary<int, string>();
                try
                {
                    List<ListItem> dictitem = new List<ListItem>();
                    string strComm = "SELECT * from Sensor";
                    OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                    OleDbDataAdapter oleDbDataAda = new OleDbDataAdapter(oleDbCmd);
                    DataSet daSet = new DataSet();
                    oleDbDataAda.Fill(daSet);
                    foreach (DataRow r in daSet.Tables[0].Rows)
                    {
                        dictitem.Add(new ListItem(Convert.ToInt16(r["ID"]), r["SensorName"].ToString()));
                    }
                    this.list.Dispatcher.Invoke(new Action(() => { list.ItemsSource = null; list.ItemsSource = dictitem; list.DisplayMemberPath = "SensorName"; list.SelectedValuePath = "ID"; }));
                }
                catch
                {
                    System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
                }

            });
        }


        //触发事件改变MainWindow的值
        private void Dellist(int ID)
        {
            try
            {
                string strComm = "DELETE from Sensor where ID=" + ID;
                OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                oleDbCmd.ExecuteNonQuery();
            }
            catch
            {
                System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
            }

        }
        private void Addllist()
        {
            try
            {
                string strComm = "INSERT INTO Sensor (SensorName, ItemName, ChannelName, StationID, [Decimal], SensorType) VALUES ('" + stname.Text + "','"+ time.Text + "','"+st.Text+"',"+combo.SelectedValue+",3,0)";
                OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                oleDbCmd.ExecuteNonQuery();
                Getlist1();
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
                string strComm = "UPDATE Sensor SET SensorName = '" + stname.Text + "', ItemName = '" + time.Text + "', ChannelName = '" + st.Text + "', StationID = "+ combo.SelectedValue + " WHERE ID =" + ID;
                OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                oleDbCmd.ExecuteNonQuery();
                Getlist1();
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
                    string strComm = "select * from Sensor WHERE ID=" + ID;
                    OleDbCommand oleDbCmd = new OleDbCommand(strComm, cn);
                    OleDbDataAdapter oleDbDataAda = new OleDbDataAdapter(oleDbCmd);
                    DataSet daSet = new DataSet();
                    oleDbDataAda.Fill(daSet);
                    foreach (DataRow r in daSet.Tables[0].Rows)
                    {
                        this.combo.Dispatcher.Invoke(new Action(() => { combo.SelectedValue = r["StationID"]; }));
                        this.st.Dispatcher.Invoke(new Action(() => { st.Text = r["ChannelName"].ToString(); }));
                        this.stname.Dispatcher.Invoke(new Action(() => { stname.Text = r["SensorName"].ToString(); }));
                        this.time.Dispatcher.Invoke(new Action(() => { time.Text = r["ItemName"].ToString(); }));
                    }
                }
                catch
                {
                    System.Windows.MessageBox.Show("MDB文件不存在，无法使用。");
                }

            });
        }
        public Window3()
        {
            InitializeComponent();
            Getlist();
        }

        private void List_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (list.SelectedIndex != -1)
            {
                Getitem(Convert.ToInt32(list.SelectedValue));
                st.IsEnabled = true;
                time.IsEnabled = true;
                stname.IsEnabled = true;
                cl.IsEnabled = true;
                check.IsEnabled = true;
                combo.IsEnabled = true;
                del.IsEnabled = true;
            }
            else
            {
                st.IsEnabled = false;
                time.IsEnabled = false;
                stname.IsEnabled = false;
                cl.IsEnabled = false;
                check.IsEnabled = false;
                combo.IsEnabled = false;
                del.IsEnabled = false;
                st.Text = "";
                time.Text = "";
                stname.Text = "";
                combo.SelectedIndex = -1;
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

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (list.SelectedIndex != -1)
                Getitem(Convert.ToInt32(list.SelectedValue));
            else
            {
                st.Text = "";
                time.Text = "";
                stname.Text = "";
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            list.SelectedIndex = -1;
            st.IsEnabled = true;
            time.IsEnabled = true;
            stname.IsEnabled = true;
            cl.IsEnabled = true;
            check.IsEnabled = true;
            combo.IsEnabled = true;
            st.Text = "";
            time.Text = "";
            stname.Text = "";
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Dellist(Convert.ToInt32(list.SelectedValue));
            Getlist();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            sendMessage1("2");
        }
    }
}
