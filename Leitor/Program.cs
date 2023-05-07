using System;
using System.IO;
using System.Reflection;
using Leitor;
using Leitor.Model;

class Program
{
    static void Main(string[] args)
    {
        /*
         * Inicio do Pregrama!
         * 
         * Aqui será chamado as funções e métodos para buscar e guardar as 
         * informações do computador. 
         */

        Console.WriteLine("Capturando informações do computador!");
        //Instanciando objeto InfoComputer.
        InfoComputer computer = new InfoComputer();
        MOS mos = new MOS(); //Instaciando o MOS que por sua vez irá chamar o métodos de busca.

        //chamado do GetComputer que por suavez ira retorna um InfoComputer com todos os dados preenchidos.
        computer = mos.GetComputerInfo(); 

        Console.WriteLine("Informações Capturadas!!\n\nAquarde anquanto o arquivo é gerado.");
        Console.Clear();
        /*
         * Aqui criaremos o arquivo
         * 
         * Nesse programa está no momento sómente a opção de gerar um Excel
         */
        try
        {
            //Instanciando um Gerador de arquivo Excel.
            IFileGenerator file = new ExcelGenerator();
            //Criando o Arquivo com as demais informações do Computador.
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
