using System;
using System.Collections.Generic;
using System.Management;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Net.NetworkInformation;
using Leitor;

namespace Leitor.Model
{
    //Classe principal que irá se preocupar sómente em manter as informações nescessarias do computador
    public class InfoComputer
    {
        public string hostName { get; set; }
        public string serialNumber { get; set; }
        public string modelo { get; set; }
        public string status { get; set; }
        public string localPadrao { get; set; }
        public string local { get; set; }
        public string usuario { get; set; }
        public string macLAN { get; set; }
        public string macWIFI { get; set; }
        public string memoria { get; set; }
        public string hd { get; set; }
        public string processador { get; set; }
    }
    public interface IComputerProvider
    {
        InfoComputer GetComputerInfo();
    }
    public interface IInfoProvider
    {
        string GetHost();
        string GetSerialNumber();
        string GetModelo();
        string GetStatus();
        string GetLocalPadrao();
        string GetLocal();
        string GetUsuario();
        string GetMacLAN();
        string GetMacWIFI();
        string GetMemoria();
        string GetHD();
        string GetProcessador();
    }
    public class MOS : IComputerProvider, IInfoProvider
    {
        public InfoComputer GetComputerInfo()
        {
            InfoComputer computer = new InfoComputer();
            computer.hostName = this.GetHost();
            computer.serialNumber = this.GetSerialNumber();
            computer.hd = this.GetHD();
            computer.local = this.GetLocal();
            computer.localPadrao = this.GetLocalPadrao();
            computer.macLAN = this.GetMacLAN();
            computer.macWIFI = this.GetMacWIFI();
            computer.memoria = this.GetMemoria();
            computer.modelo = this.GetModelo();
            computer.status = this.GetStatus();
            computer.processador = this.GetProcessador();
            computer.usuario = this.GetUsuario();

            return computer;
        }
        public string GetHD()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return $"{Util.GetTipoDisco()} - {Util.FormatarTamanhoMemoria(Convert.ToInt64(result?["Size"]))}";
        }
        public string GetHost()
        {
            return Environment.MachineName;
        }
        public string GetLocal()
        {
            return "---";
        }
        public string GetLocalPadrao()
        {
            return "Seller LEAD";
        }
        public string GetMacLAN()
        {
            //var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'");
            //var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();
            //return result?["MACAddress"]?.ToString();
            var ethernet = NetworkInterface
                        .GetAllNetworkInterfaces()
                        .FirstOrDefault(x => x.NetworkInterfaceType == NetworkInterfaceType.Ethernet);
            return Util.FormatarMAC(ethernet.GetPhysicalAddress().ToString());
        }
        public string GetMacWIFI()
        {

            //var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'");
            //var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault
            //return result?["MACAddress"]?.ToString();
            var wifi = NetworkInterface
                        .GetAllNetworkInterfaces()
                        .FirstOrDefault(x => x.NetworkInterfaceType == NetworkInterfaceType.Wireless80211);
            return Util.FormatarMAC(wifi.GetPhysicalAddress().ToString());
        }
        public string GetMemoria()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return Util.FormatarTamanhoMemoria(Convert.ToInt64(result?["TotalPhysicalMemory"])).ToString();
        }
        public string GetModelo()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return result?["Model"]?.ToString();
        }
        public string GetProcessador()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return result?["Name"]?.ToString();
        }
        public string GetSerialNumber()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return result?["SerialNumber"]?.ToString();
        }
        public string GetStatus()
        {
            return "Ok";
        }
        public string GetUsuario()
        {
            return Environment.UserName;
        }
    }
}
