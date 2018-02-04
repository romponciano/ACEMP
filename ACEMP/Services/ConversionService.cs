using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACEMP.Services;

namespace ACEMP.Services
{
    class ConversionService
    {
        public static void datatable2xls(DataTable final, string local)
        {
            ExcelEngine ExcelEngineObject = new ExcelEngine();
            IApplication Application = ExcelEngineObject.Excel;
            Application.DefaultVersion = ExcelVersion.Excel2013;
            IWorkbook Workbook = Application.Workbooks.Create(1);
            Workbook.StandardFont = "Verdana";
            Workbook.StandardFontSize = 11;
            IWorksheet Worksheet = Workbook.Worksheets[0];
            Worksheet.ImportDataTable(final, true, 1, 1);
            Workbook.SaveAs(local);
            Workbook.Close();
            ExcelEngineObject.Dispose();
        }

        public static DataTable csv2datatable(string caminho)
        {
            string caminhoIni = FileService.gerarSchemaCsv(caminho);

            DataTable dt = new DataTable("data");
            using (OleDbConnection conexao = new OleDbConnection(
                    "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=\"" + Path.GetDirectoryName(caminho) + "\";" +
                    "Extended Properties='text;HDR=yes;'"
                )
            )
            {
                using (OleDbCommand cmd = new OleDbCommand(
                        string.Format("select * from [{0}]", new FileInfo(caminho).Name),
                        conexao
                    )
                )
                {
                    conexao.Open();
                    using (OleDbDataAdapter adaptador = new OleDbDataAdapter(cmd))
                    {
                        adaptador.Fill(dt);
                    }
                }
            }

            FileService.deletarArquivo(caminhoIni);

            return dt;
        }

        public static DataTable csv2numeronfs(string caminho)
        {
            string caminhoIni = FileService.gerarSchemaCsvNumeroNfs(caminho);

            DataTable dt = new DataTable("data");
            using (OleDbConnection conexao = new OleDbConnection(
                    "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=\"" + Path.GetDirectoryName(caminho) + "\";" +
                    "Extended Properties='text;HDR=yes;'"
                )
            )
            {
                using (OleDbCommand cmd = new OleDbCommand(
                        string.Format("select * from [{0}]", new FileInfo(caminho).Name),
                        conexao
                    )
                )
                {
                    conexao.Open();
                    using (OleDbDataAdapter adaptador = new OleDbDataAdapter(cmd))
                    {
                        adaptador.Fill(dt);
                    }
                }
            }

            FileService.deletarArquivo(caminhoIni);

            return dt;
        }
    }
}
