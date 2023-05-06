using System;
using System.IO;
using System.Reflection;
using Leitor;
using Leitor.Model;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Capturando informações do computador!");
        InfoComputer computer = new InfoComputer();
        MOS mos = new MOS();
        computer = mos.GetComputerInfo();
        Console.WriteLine("Informações Capturadas!!\n\nAquarde anquanto o arquivo é gerado.");

        try
        {
            Console.Clear();
            IFileGenerator file = new ExcelGenerator();
            file.Generate(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\ExcelGerado\\Infos.xlsx"), computer);
            Console.WriteLine("Arquivo Excel gerado com sucesso!!");
            Console.ReadKey();
        } catch
        {
            Console.WriteLine("O arquivo esta sendo usado no momento!\nFavor feche e execute o Leitor novamente!!");
            Console.ReadKey();
        }

    }
}
