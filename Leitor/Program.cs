using System;
using System.IO;
using System.Reflection;
using System.Threading;
using Leitor;
using Leitor.Model;

using System.Timers;

class Program
{
    private static string author = "Otonyel Carvalho da Silva Brandão".ToUpper();
    private static string date = "06/05/2023".ToUpper();
    private static string programName = "Leitor".ToUpper();
    static void Main(string[] args)
    {
        /*
         * Inicio do Pregrama!
         * 
         * Apresentação do cebeçalho do pragrama.
         * 
         * Aqui será chamado as funções e métodos para buscar e guardar as 
         * informações do computador. 
         */
        //Chamando o Caeçalho
        Cabecalho();
        print("  Capturando informações do computador, aguarde enquanto o processo é executado");
        Carregamento();
        //Instanciando objeto InfoComputer.
        InfoComputer computer = new InfoComputer();
        MOS mos = new MOS(); //Instaciando o MOS que por sua vez irá chamar o métodos de busca.

        //chamado do GetComputer que por suavez ira retorna um InfoComputer com todos os dados preenchidos.
        computer = mos.GetComputerInfo();
        Clear();

        print("  Informações Capturadas!!\n\n  Aquarde anquanto o arquivo é gerado");
        Carregamento();
        Clear();
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
            file.Generate(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\ExcelGerado"), computer);

            print("  Arquivo Excel gerado com sucesso!!");
            print("\n\n  Aperte Qualquer tecla para finalizar o programa.");
            Console.ReadKey();
        } catch
        {
            print("  O arquivo esta sendo usado no momento!\n\tFavor feche e execute o Leitor novamente!!");
            Console.ReadKey();
        }

    }
    /// <summary>
    /// Essa função cria e escreve o Cabeçalho do pragrama.
    /// </summary>
    private static void Cabecalho()
    {
        string authorAndDate = " Autor: " + author + "\n Data da Criação: " + date;
        int totalSpaces = programName.Length - authorAndDate.Length;

        if (totalSpaces < 0)
            totalSpaces = 0;

        int leftSpaces = totalSpaces / 2;
        int rightSpaces = totalSpaces - leftSpaces;

        string header = "" + new string(' ', leftSpaces) + authorAndDate + new string(' ', rightSpaces);

        Console.SetCursorPosition(0, 0); // posiciona o cursor no início do console
        Console.WriteLine("+" + new string('-', programName.Length + 110) + "+\n");
        Console.WriteLine("" + " Nome do Pragrama: ".ToUpper() + programName.ToUpper());
        Console.WriteLine(header.ToUpper()+"\n");
        Console.WriteLine("+" + new string('-', programName.Length + 110) + "+\n");
        Thread.Sleep(1000);
    }
    /// <summary>
    /// Essa função limpa o código e escreve o Cabeçalho do pragrama.
    /// </summary>
    private static void Clear()
    {
        Console.SetCursorPosition(0, 8);
        Console.Write(new string(' ', 1000));
        Console.SetCursorPosition(0, 8);
    }
    /// <summary>
    /// Essa simula o carregamento do programa.
    /// </summary>
    private static void Carregamento()
    {
        int count = 1;
        System.Timers.Timer timer = new System.Timers.Timer(300);
        timer.Elapsed += (sender, e) => {
            Console.Write(".");
            count++;
            if (count == 4)
            {
                Console.Write("\b\b\b   \b\b\b"); // Apaga os três pontos
                count = 1;
            }
        };
        timer.Enabled = true;

        System.Threading.Thread.Sleep(6000);
        timer.Enabled = false;
    }
    /// <summary>
    /// Essa função suavisa ao escrever o texto na tela
    /// </summary>
    /// <param name="texto">String com o texto a ser escrito</param>
    private static void print(string texto)
    {
        foreach (char c in texto)
        {
            Console.Write(c);
            Thread.Sleep(75);
        }
    }
}
