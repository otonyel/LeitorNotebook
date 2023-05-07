using System;
using System.IO;

namespace Leitor
{
    /// <summary>
    /// Essa classe é usada para os utilitarios dentre todo o código, Nela deve ser mantido calculos ou formatações que possam ser reutilizados no mais variados codigos o nas novas implementações a serem criadas.
    /// </summary>
    public class Util
    {
        /// <summary>
        /// Esse Método trata de formata a quantidade de memoria podendo ser utilizado com HD ou até memoria RAM.
        /// </summary>
        /// <param name="bytes">A quantidade de bytes que o disco ou memoria Ram contem</param>
        /// <returns>Retorna uma string com o valor formatado conforma a quantidade de bytes.</returns>
        public static string FormatarTamanhoMemoria(long bytes)
        {
            string[] suf = { "B", "KB", "MB", "GB", "TB", "PB", "EB" };
            if (bytes == 0)
                return "0" + suf[0];
            long bytesAbs = Math.Abs(bytes);
            int lugar = Math.Min(suf.Length - 1, Convert.ToInt32(Math.Floor(Math.Log(bytesAbs, 1024))));
            double tamanho = Math.Round(bytesAbs / Math.Pow(1024, lugar), 1);
            if (tamanho >= 1000)
            {
                tamanho /= 1024;
                lugar++;
            }
            return string.Format("{0:N1} {1}", tamanho, suf[lugar]);
        }
        /// <summary>
        /// Esse Método informa o tipo de disco se é HD ou SSD
        /// </summary>
        /// <returns>Retorna uma String com o valor formatado indicando HDD para HD ou SSD para SSD</returns>
        public static string GetTipoDisco()
        {
            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if (drive.DriveType == DriveType.Fixed && drive.IsReady)
                {
                    string tipo = drive.DriveFormat.Contains("NTFS") ? "HDD" : "SSD";
                    return tipo;
                }
            }
            return "";
        }
        /// <summary>
        /// Esse Método formata o MAC seguindo o padrã de duas casas dois pontos.
        /// Exemplo: 00:00:00:00:00:00
        /// </summary>
        /// <param name="mac">String com o valor MAC a ser formatado</param>
        /// <returns>Retorna uma String com o valor formatado do MAC.</returns>
        public static string FormatarMAC(string mac)
        {
            string macFormatado = string.Empty;

            for (int i = 0; i < mac.Length; i++)
            {
                if (i % 2 == 0 && i > 0)
                {
                    macFormatado += ":";
                }

                macFormatado += mac[i];
            }

            return macFormatado;
        }
    }
}
