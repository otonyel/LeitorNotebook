using Microsoft.Office.Interop.Excel;
using Leitor.Model;
using System.Linq;

namespace Leitor
{
    public interface IFileGenerator
    {
        void Generate(string filePath, InfoComputer computer);
    }
    public class ExcelGenerator : IFileGenerator
    {
        public Application App { get; set; }
        public Workbook wb { get; set;}
        public Worksheet sheet { get; set;}
        public ExcelGenerator()
        {
            App = new Application();

            wb = App.Workbooks.Add();
            sheet = wb.ActiveSheet;
            sheet.Name = "Informações Computador";   
        }
        public void SetCelAndValue(int line,int col, string value)
        {
            this.sheet.Cells[line,col] = value;
        }
        public void Generate(string filePath, InfoComputer computer)
        {
            //Nome das colunas
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
           //Adicionando uma linha
            this.SetCelAndValue(2, 1, computer.hostName);
            this.SetCelAndValue(2, 2, computer.serialNumber);
            this.SetCelAndValue(2, 3, computer.modelo);
            this.SetCelAndValue(2, 4, computer.status);
            this.SetCelAndValue(2, 5, computer.localPadrao);
            this.SetCelAndValue(2, 6, computer.local);
            this.SetCelAndValue(2, 7, computer.usuario);
            this.SetCelAndValue(2, 8, computer.macLAN);
            this.SetCelAndValue(2, 9, computer.macWIFI);
            this.SetCelAndValue(2, 10, computer.memoria);
            this.SetCelAndValue(2, 11, computer.hd);
            this.SetCelAndValue(2, 12, computer.processador);
            // Salve o Excel como arquivo .xlsx.
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
