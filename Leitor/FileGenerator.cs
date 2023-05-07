using Microsoft.Office.Interop.Excel;
using Leitor.Model;

namespace Leitor
{
    /// <summary>
    /// Essa Abistração cuidara de generaizar a criação do arquivo podendo colocar alem de Excel uma PDF ou Word ou txt o que vier a nescessidade.
    /// </summary>
    public interface IFileGenerator
    {
        void Generate(string filePath, InfoComputer computer);
    }
    /// <summary>
    /// Essa Gerador de Excel é uma classe que trata de obter as informações contidas no objeto InfoComputer e gerar um arqui vo do Excel con as informações devidamente organizadas.
    /// </summary>
    public class ExcelGenerator : IFileGenerator
    {
        public Application App { get; set; }
        public Workbook wb { get; set;}
        public Worksheet sheet { get; set;}
        public ExcelGenerator()
        {
            App = new Application();
            wb = App.Workbooks.Add();
            App.DisplayAlerts = false;
            sheet = wb.ActiveSheet;
            sheet.Name = "Informações Computador";   
        }
        /// <summary>
        /// Essa função trata de simplificar a atribuição das celulas.
        /// </summary>
        /// <param name="line">Valor numérico da linha a ser alterada</param>
        /// <param name="col">Valor numérico da coluna a ser alterada</param>
        /// <param name="value">Valor a ser atribuida a celula</param>
        public void SetCelAndValue(int line,int col, string value)
        {
            this.sheet.Cells[line,col] = value;
        }
        /// <summary>
        /// Essa função cria a tabela a adiciona uma linha com os devidos valores organizados.
        /// </summary>
        /// <param name="filePath">Caminho do arquivo a ser criado.</param>
        /// <param name="computer">Objeto que contem os valores a serem escritos na planilha.</param>
        public void Generate(string filePath, InfoComputer computer)
        {
            /*
             * Cabeçalho
             * 
             * Atribuição dos nomes das colunas.
             */
            this.SetCelAndValue(1, 1, "HostName");
            this.SetCelAndValue(1, 2, "Serial Number");
            this.SetCelAndValue(1, 3, "Modelo");
            this.SetCelAndValue(1, 4, "Status");
            this.SetCelAndValue(1, 5, "LocalPadrão");
            this.SetCelAndValue(1, 6, "Local");
            this.SetCelAndValue(1, 7, "Usuário");
            this.SetCelAndValue(1, 8, "MAC (LAN)");
            this.SetCelAndValue(1, 9, "MAC (WIFI)");
            this.SetCelAndValue(1, 10,"Memória");
            this.SetCelAndValue(1, 11, "HD");
            this.SetCelAndValue(1, 12, "Processador");
           /*
            * Corpo
            * 
            * Criação da linha com os valores devidamente organizados.
            */
            this.SetCelAndValue(2, 1, computer.hostName);       // Host do computador
            this.SetCelAndValue(2, 2, computer.serialNumber);   // Serial do computador
            this.SetCelAndValue(2, 3, computer.modelo);         // Modelo do computador
            this.SetCelAndValue(2, 4, computer.status);         // Status do computador
            this.SetCelAndValue(2, 5, computer.localPadrao);    // Local Padrão do computador
            this.SetCelAndValue(2, 6, computer.local);          // Local do computador
            this.SetCelAndValue(2, 7, computer.usuario);        // Usuário do computador
            this.SetCelAndValue(2, 8, computer.macLAN);         // MAC (LAN) do computador
            this.SetCelAndValue(2, 9, computer.macWIFI);        // MAC (WIFI) do computador
            this.SetCelAndValue(2, 10, computer.memoria);       // Memoria RAM do computador
            this.SetCelAndValue(2, 11, computer.hd);            // Disco do computador
            this.SetCelAndValue(2, 12, computer.processador);   // Processador do computador

            /*
             * Salvamento do arquivo gerado
             * 
             * Aqui o arquivo será salvo no local informado pela variavel filePath.
             */
            try
            {
                wb.SaveAs(filePath);
                wb.Close();
            }
            catch (System.IO.IOException ex)
            {
                throw ex;
            }
        }
    }
}
