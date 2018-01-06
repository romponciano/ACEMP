using ACEMP.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACEMP.Services
{
    class FileService
    {
        public static string gerarSchemaCsv(string caminho)
        {
            using (FileStream fs = new FileStream(Path.GetDirectoryName(caminho) + "\\schema.ini", FileMode.Create, FileAccess.Write))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    string[] nomeArquivo = caminho.Split('\\');
                    sw.WriteLine("[" + nomeArquivo.GetValue(nomeArquivo.Length - 1) + "]");
                    sw.WriteLine("ColNameHeader=True");
                    sw.WriteLine("Format=Delimited(;)");
                    sw.WriteLine("DecimalSymbol=,");
                    sw.WriteLine("DateTimeFormat=DD-MM-YYYY");
                    sw.Close();
                    sw.Dispose();
                }

                fs.Close();
                fs.Dispose();
            }
            return Path.GetDirectoryName(caminho) + "\\schema.ini";
        }

        public static void deletarArquivo(string caminho)
        {
            if (File.Exists(@caminho))
            {
                File.Delete(@caminho);
            }
        }
    }
}
