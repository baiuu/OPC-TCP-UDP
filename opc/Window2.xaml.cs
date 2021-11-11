using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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

namespace opc
{
    /// <summary>
    /// Window2.xaml 的交互逻辑
    /// </summary>
    /// 

    public delegate void SendMessage(JObject str);
    public partial class Window2 : Window
    {
        public JObject jsonObject1;
        public SendMessage sendMessage;
        public Window2(JObject jsonObject)
        {
            InitializeComponent();
            jsonObject1 = jsonObject;
            cb.Items.Add("udp");
            cb.Items.Add("tcp");
            cb.SelectedItem=jsonObject["Type"].ToString();
            ip.Text = jsonObject["IP"].ToString();
            port.Text = jsonObject["Port"].ToString();
            opc.Text = jsonObject["OPCserver"].ToString();
            time.Text= jsonObject["Time"].ToString();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            jsonObject1["Type"] = cb.SelectedItem.ToString();
            jsonObject1["IP"] = ip.Text;
            jsonObject1["Port"] = port.Text;
            jsonObject1["OPCserver"] = opc.Text;
            jsonObject1["Time"] = time.Text;
            sendMessage(jsonObject1);
            Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
