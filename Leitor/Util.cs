using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Leitor
{
    
    public class Util
    {
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
