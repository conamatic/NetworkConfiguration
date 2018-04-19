using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.Net.NetworkInformation;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;

namespace NetworkConfiguration
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string InterfaceToUse = "";
        List<Config> Configurations = new List<Config>();

        public MainWindow()
        {
            InitializeComponent();
            NetworkInterface[] properties = NetworkInterface.GetAllNetworkInterfaces();

            foreach (NetworkInterface adapter in properties)
            {
                interfaceList.Items.Add(adapter.Description);
            }

            if (File.Exists("configuration.json"))
            {
                using (StreamReader reader = new StreamReader("configuration.json"))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        Config configs = JsonConvert.DeserializeObject<Config>(line);

                        ConfigurationList.Items.Add(configs.Name);
                        Configurations.Add(configs);
                    }
                }
            }

            interfaceList.MouseDoubleClick += new MouseButtonEventHandler(ListDoubleClick);
            ConfigurationList.MouseDoubleClick += new MouseButtonEventHandler(ConfigDoubleClick);
        }

        private void ListDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (interfaceList.SelectedItem != null)
            {
                InterfaceToUse = interfaceList.SelectedItem.ToString();
                interface_label.Content = interfaceList.SelectedItem.ToString();
            }
        }

        private void ConfigDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ConfigurationList.Items.Count > 0)
            {
                string ToLookFor = ConfigurationList.SelectedItem.ToString();
                foreach (Config collection in Configurations)
                {
                    try
                    {
                        string[] ip = collection.IP.Split('.');
                        ip_1.Text = ip[0]; ip_2.Text = ip[1]; ip_3.Text = ip[2]; ip_4.Text = ip[3];
                    }
                    catch (Exception) { }

                    try
                    {
                        string[] subnet = collection.Subnet.Split('.');
                        subnet_1.Text = subnet[0]; subnet_2.Text = subnet[1]; subnet_3.Text = subnet[2]; subnet_4.Text = subnet[3];
                    }
                    catch (Exception) { }

                    try
                    {
                        string[] gateway = collection.Gateway.Split('.');
                        gateway_1.Text = gateway[0]; gateway_2.Text = gateway[1]; gateway_3.Text = gateway[2]; gateway_4.Text = gateway[3];
                    }
                    catch (Exception) { }

                    try
                    {
                        string[] dns = collection.DNS.Split('.');
                        dns_1.Text = dns[0]; dns_2.Text = dns[1]; dns_3.Text = dns[2]; dns_4.Text = dns[3];
                    }
                    catch (Exception) { }

                    try
                    {
                        string dns_suffix = collection.DNS_Suffix;
                        dnssuffix.Text = dns_suffix;
                    }
                    catch (Exception) { }

                    try
                    {
                        InterfaceToUse = collection.@Interface;
                        interface_label.Content = collection.@Interface;
                    }
                    catch (Exception) { }
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            PrepareIPChange();
        }

        public void PrepareIPChange()
        {
            string IP = ip_1.Text + "." + ip_2.Text + "." + ip_3.Text + "." + ip_4.Text;
            string subnet = subnet_1.Text + "." + subnet_2.Text + "." + subnet_3.Text + "." + subnet_4.Text;
            string gateway = gateway_1.Text + "." + gateway_2.Text + "." + gateway_3.Text + "." + gateway_4.Text;
            string dns = dns_1.Text + "." + dns_2.Text + "." + dns_3.Text + "." + dns_4.Text;
            string dns_suffix = dnssuffix.Text;

            setIP(IP, subnet);
            setGateway(gateway);
            if (!String.IsNullOrEmpty(InterfaceToUse))
                setDNS(InterfaceToUse, dns, dns_suffix);
        }

        public void setIP(string ip_address, string subnet_mask)
        {
            ManagementClass objMC = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection objMOC = objMC.GetInstances();

            foreach (ManagementObject objMO in objMOC)
            {
                if ((bool)objMO["IPEnabled"])
                {
                    try
                    {
                        ManagementBaseObject setIP;
                        ManagementBaseObject newIP =
                            objMO.GetMethodParameters("EnableStatic");

                        newIP["IPAddress"] = new string[] { ip_address };
                        newIP["SubnetMask"] = new string[] { subnet_mask };

                        setIP = objMO.InvokeMethod("EnableStatic", newIP, null);
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
            }
        }

        public void setGateway(string gateway)
        {
            ManagementClass objMC = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection objMOC = objMC.GetInstances();

            foreach (ManagementObject objMO in objMOC)
            {
                if ((bool)objMO["IPEnabled"])
                {
                    try
                    {
                        ManagementBaseObject setGateway;
                        ManagementBaseObject newGateway =
                            objMO.GetMethodParameters("SetGateways");

                        newGateway["DefaultIPGateway"] = new string[] { gateway };
                        newGateway["GatewayCostMetric"] = new int[] { 1 };

                        setGateway = objMO.InvokeMethod("SetGateways", newGateway, null);
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
            }
        }

        public void setDNS(string NIC, string DNS, string suffix)
        {
            ManagementClass objMC = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection objMOC = objMC.GetInstances();

            foreach (ManagementObject objMO in objMOC)
            {
                if ((bool)objMO["IPEnabled"])
                {
                    if (objMO["Caption"].Equals(NIC))
                    {
                        try
                        {
                            ManagementBaseObject newDNS = objMO.GetMethodParameters("SetDNSServerSearchOrder");
                            newDNS["DNSServerSearchOrder"] = DNS.Split(',');

                            ManagementBaseObject setDNS = objMO.InvokeMethod("SetDNSServerSearchOrder", newDNS, null);

                            string[] SearchOrder = new string[] { suffix, "8.8.8.8" };

                            ManagementBaseObject newSuffix = objMO.GetMethodParameters("SetDNSDomainSuffixSearchOrder");
                            newSuffix["DNSDomainSuffixSearchOrder"] = SearchOrder;
                            ManagementBaseObject setSuffix = objMO.InvokeMethod("SetDNSDomainSuffixSearchOrder", newSuffix, null);
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                    }
                }
            }
        }

        public void setWINS(string NIC, string priWINS, string secWINS)
        {
            ManagementClass objMC = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection objMOC = objMC.GetInstances();

            foreach (ManagementObject objMO in objMOC)
            {
                if ((bool)objMO["IPEnabled"])
                {
                    if (objMO["Caption"].Equals(NIC))
                    {
                        try
                        {
                            ManagementBaseObject setWINS;
                            ManagementBaseObject wins =
                            objMO.GetMethodParameters("SetWINSServer");
                            wins.SetPropertyValue("WINSPrimaryServer", priWINS);
                            wins.SetPropertyValue("WINSSecondaryServer", secWINS);

                            setWINS = objMO.InvokeMethod("SetWINSServer", wins, null);
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                    }
                }
            }
        }

        public void EnableDHCP()
        {
            ManagementClass objMC = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection objMOC = objMC.GetInstances();

            foreach (ManagementObject objMO in objMOC)
            {
                if ((bool)objMO["IPEnabled"])
                {
                    try
                    {
                        ManagementBaseObject newDNS = objMO.GetMethodParameters("SetDNSServerSearchOrder");
                        newDNS["DNSServerSearchOrder"] = null;
                        ManagementBaseObject enableDHCP = objMO.InvokeMethod("EnableDHCP", null, null);
                        ManagementBaseObject setDNS = objMO.InvokeMethod("SetDNSServerSearchOrder", newDNS, null);
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
            }
        }

        public void FieldChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            try
            {
                int value = Convert.ToInt32(textBox.Text);
                if (value > 255)
                    textBox.BorderBrush = Brushes.Red;
                else
                    textBox.ClearValue(BorderBrushProperty);
            }
            catch (Exception)
            {
                textBox.BorderBrush = Brushes.Red;
            }
        }

        private void MoveToNext()
        {
            TraversalRequest traversalRequest = new TraversalRequest(FocusNavigationDirection.Next);
            UIElement keyboardFocus = Keyboard.FocusedElement as UIElement;

            if (keyboardFocus != null)
                keyboardFocus.MoveFocus(traversalRequest);
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            string ip = ip_1.Text + "." + ip_2.Text + "." + ip_3.Text + "." + ip_4.Text;
            string subnet = subnet_1.Text + "." + subnet_2.Text + "." + subnet_3.Text + "." + subnet_4.Text;
            string gateway = gateway_1.Text + "." + gateway_2.Text + "." + gateway_3.Text + "." + gateway_4.Text;
            string dns = dns_1.Text + "." + dns_2.Text + "." + dns_3.Text + "." + dns_4.Text;
            string dns_suffix = dnssuffix.Text;
            string name = NameInput.Text;
            
            if (Configurations.Any(c => c.Name == name))
            {
                MessageBoxResult result = MessageBox.Show("A record for this address already exists, overwrite?", "Overwrite", MessageBoxButton.OKCancel, MessageBoxImage.Question);

                if (result == MessageBoxResult.OK)
                {
                    int index = Configurations.FindIndex(c => c.Name == name);
                    Configurations.RemoveAt(index);
                    ConfigurationList.Items.RemoveAt(index);
                    ConfigurationList.Items.Refresh();
                }
                else
                {
                    return;
                }
            }

            Config config = new Config
            {
                IP = ip,
                Subnet = subnet,
                Gateway = gateway,
                DNS = dns,
                DNS_Suffix = dns_suffix,
                @Interface = InterfaceToUse,
                Name = name
            };

            var json = JsonConvert.SerializeObject(config);
            using (StreamWriter writer = new StreamWriter("configuration.json", true))
            {
                writer.WriteLine(json.ToString());
            }

            Configurations.Add(config);
            ConfigurationList.Items.Add(name);
        }

        private void DHCPButton_Click(object sender, RoutedEventArgs e)
        {
            EnableDHCP();
        }

        private void miscInput_KeyDown(object sender, KeyEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            if (textBox.Text.Length == 3 && !textBox.Name.EndsWith("4"))
            {
                MoveToNext();
                e.Handled = true;
            }
        }
    }

    public class Config
    {
        public string IP { get; set; }
        public string Subnet { get; set; }
        public string Gateway { get; set; }
        public string DNS { get; set; }
        public string DNS_Suffix { get; set; }
        public string @Interface { get; set; }
        public string Name { get; set; }
    }
}
