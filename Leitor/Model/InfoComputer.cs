using System;
using System.Management;
using System.Linq;
using System.Net.NetworkInformation;

namespace Leitor.Model
{
    /// <summary>
    /// InfoComputer essa classe tem afunção de guradar as informações do computador.
    /// 
    /// Todas as suas variaveis são strings entção todo valor obtido deve ser obrigatóriamente convertido em string.
    /// 
    /// Suas propriedades são hostName, serialNumber, modelo, status, localPadrao, local, usuario, macLAN, macWIFI, memoria, hd e processador.
    /// </summary>
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
    /// <summary>
    ///  Essa interface é usada como basa para qualquer classe que for realizar a busca das informações no computador.
    /// </summary>
    public interface IComputerProvider
    {
        /// <summary>
        /// O Método em questão ira guardar as informações do computador
        /// </summary>
        /// <returns>O retorno obrigatóriamente será o tipo InfoComputer</returns>
        InfoComputer GetComputerInfo();
    }

    /// <summary>
    /// Essa interface é usada como abistração para obter as informações do computador tendo como possibilidade de utilizar qualquer método para buscar tais informações.
    /// </summary>
    public interface IInfoProvider
    {
        /// <summary>
        /// Metodo que buscara o Host do computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetHost();
        /// <summary>
        /// Metodo que buscara o Serial do computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetSerialNumber();
        /// <summary>
        /// Metodo que buscara o Modelo do computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetModelo();
        /// <summary>
        /// Metodo que buscara o Status do computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetStatus();
        /// <summary>
        /// Metodo que buscara o Local do computador.
        /// Obs: Esse método foi colocado para escrever em Hardcode pós ouve uma nescessidade de colocar o Local padrão.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetLocalPadrao();
        /// <summary>
        /// Metodo que buscara o Local do computador.
        /// Obs: Esse método foi colocado para escrever em Hardcode pós ouve uma nescessidade de colocar o Local padrão.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetLocal();
        /// <summary>
        /// Metodo que buscara o Usuário conectado no computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetUsuario();
        /// <summary>
        /// Metodo que buscara o Mac Adress LAN do computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetMacLAN();
        /// <summary>
        /// Metodo que buscara o Mac Adress WIFI do computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetMacWIFI();
        /// <summary>
        /// Metodo que buscara a Quantidade da Memoria RAM do computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetMemoria();
        /// <summary>
        /// Metodo que buscara o tipo do disco e a quantidade de memoria.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetHD();
        /// <summary>
        /// Metodo que buscara o nome do processador instalado no computador.
        /// </summary>
        /// <returns>O retorno deve ser uma string.</returns>
        string GetProcessador();
    }
    /// <summary>
    /// Essa classe utiliza métodos atravez do ManagementObjectSearcher para buscar as informações do computador em que o programa é executado.
    /// </summary>
    public class MOS : IComputerProvider, IInfoProvider
    {
        /// <summary>
        /// Esse Método chama todos os métodos de busca na classe e atribui os valores devidamente ao objeto InfoComputer.
        /// </summary>
        /// <returns>Retorna um InfoComputer com todas as informações encontrads do computador.</returns>
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
        /// <summary>
        /// Esse Método Retorna o tipo do disco e a quantidade de memoria contida no disco.
        /// </summary>
        /// <returns>Retorna uma string formatado com tipo do disco e quantidade memoria.</returns>
        public string GetHD()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return $"{Util.GetTipoDisco()} - {Util.FormatarTamanhoMemoria(Convert.ToInt64(result?["Size"]))}";
        }
        /// <summary>
        /// Esse Método retorna o nome do computador
        /// </summary>
        /// <returns>Retorna uma string contendo o nome do computador.</returns>
        public string GetHost()
        {
            return Environment.MachineName;
        }
        /// <summary>
        /// Esse método retorna um texto escrito no próprio código para preenchimento nescessario.
        /// </summary>
        /// <returns>Retorna uma string de valor prédefinido</returns>
        public string GetLocal()
        {
            return "---";
        }
        /// <summary>
        /// Esse método retorna um texto escrito no próprio código para preenchimento nescessario.
        /// </summary>
        /// <returns>Retorna uma string de valor prédefinido</returns>
        public string GetLocalPadrao()
        {
            return "Seller LEAD";
        }
        /// <summary>
        /// Esse Método busca e retorna o MAC (LAN) do computador
        /// </summary>
        /// <returns>Retorna uma String com o Mac do computador já formatado</returns>
        public string GetMacLAN()
        {
            var ethernet = NetworkInterface
                        .GetAllNetworkInterfaces()
                        .FirstOrDefault(x => x.NetworkInterfaceType == NetworkInterfaceType.Ethernet);
            return Util.FormatarMAC(ethernet.GetPhysicalAddress().ToString());
        }
        /// <summary>
        /// Esse Método busca e retorna o MAC (WIFI) do computador
        /// </summary>
        /// <returns>Retorna uma String com o Mac do computador já formatado</returns>
        public string GetMacWIFI()
        {
            var wifi = NetworkInterface
                        .GetAllNetworkInterfaces()
                        .FirstOrDefault(x => x.NetworkInterfaceType == NetworkInterfaceType.Wireless80211);
            return Util.FormatarMAC(wifi.GetPhysicalAddress().ToString());
        }
        /// <summary>
        /// Esse Método busca e retorna a quantidade de memória RAM no computador.
        /// </summary>
        /// <returns>Retorna uma String com a Memória já formatada.</returns>
        public string GetMemoria()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return Util.FormatarTamanhoMemoria(Convert.ToInt64(result?["TotalPhysicalMemory"])).ToString();
        }
        /// <summary>
        /// Esse Método busca e retorna o modelo do computador
        /// </summary>
        /// <returns>Retorna uma string com o modelo do computador</returns>
        public string GetModelo()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return result?["Model"]?.ToString();
        }
        /// <summary>
        /// Esse Método busca e retorna o nome do processador instalado no computador.
        /// </summary>
        /// <returns>Retorna Uma String com o nome do processador</returns>
        public string GetProcessador()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return result?["Name"]?.ToString();
        }
        /// <summary>
        /// Essé mètodo busca e retorna o serial do computador.
        /// </summary>
        /// <returns>Retorna uma String contendo o serial do computador</returns>
        public string GetSerialNumber()
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS");
            var result = searcher.Get().Cast<ManagementObject>().FirstOrDefault();

            return result?["SerialNumber"]?.ToString();
        }
        /// <summary>
        /// Esse Método sempre retornara Ativo pós foi escrito no proprio código o valor a ser retornado.
        /// </summary>
        /// <returns>Retorna uma String com o valor Ativo.</returns>
        public string GetStatus()
        {
            return "Ativo";
        }
        /// <summary>
        /// Esse Método busca e retorna o usuário logado no computador.
        /// </summary>
        /// <returns>Retorna uma String com o nome do usuário.</returns>
        public string GetUsuario()
        {
            return Environment.UserName;
        }
    }
}
