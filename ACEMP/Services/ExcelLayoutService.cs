using ACEMP.Model;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACEMP.Services
{
    class ExcelLayoutService
    {
        public static void formatarValores(IWorkbook workbook, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = workbook.Worksheets[0];
            planilha.Range["D7" + ":F" + linhas["linhaIrpj"]].NumberFormat = "#,###,##0.00";
            planilha.Range["G7:L" + (linhas["ultimaLinha"] + 1)].NumberFormat = "#,##0.00";
            planilha.Range["C" + linhas["linhaPis"] + ":C" + linhas["linhaIrpj"]].NumberFormat = "0.00%";
            workbook.SaveAs(caminho);
        }

        public static void ajustarDatas(IWorkbook workbook, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = workbook.Worksheets[0];
            for (int i = 7; i < linhas["ultimaLinha"]; i++)
            {
                string[] data = planilha.Range["A" + i].Value.Split('/');
                planilha.Range["A" + i].Text = data.GetValue(0) + "/" + data.GetValue(1);
            }
        }

        public static void aplicarEstilos(IWorkbook wb, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = wb.Worksheets[0];
            // estilo header main
            IStyle headerMainStyle = wb.Styles.Add("HeaderMainStyle");
            headerMainStyle.Font.Bold = true;
            headerMainStyle.Font.Color = ExcelKnownColors.Blue;
            headerMainStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            planilha.Range["A1:L5"].CellStyle.PatternColor = Color.White;
            planilha.Range["A2:A4"].CellStyle = headerMainStyle;
            // estilo header legenda de dados
            IStyle headerStyle = wb.Styles.Add("HeaderDadosStyle");
            headerStyle.Font.Bold = true;
            headerStyle.FillPattern = ExcelPattern.Solid;
            headerStyle.Interior.Color = Color.LightGray;
            planilha.Range["A6:L6"].CellStyle = headerStyle;
            headerStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            planilha.Range["A" + linhas["ultimaLinha"] + ":L" + linhas["ultimaLinha"]].CellStyle = headerStyle;
            // estilo header impostos
            IStyle headerImpostosStyle = wb.Styles.Add("HeaderImpostosStyle");
            headerImpostosStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            headerImpostosStyle.Font.Bold = true;
            planilha.Range["A" + linhas["ultimaLinha"] + ":L" + (linhas["ultimaLinha"] + 11)].CellStyle.PatternColor = Color.White;
            headerImpostosStyle.FillPattern = ExcelPattern.Solid;
            headerImpostosStyle.Interior.Color = Color.LightGray;
            planilha.Range["B" + linhas["primeiraLinha"] + ":F" + linhas["primeiraLinha"]].CellStyle = headerImpostosStyle;
            // salvar
            wb.SaveAs(caminho);
        }

        public static void alinharDados(IWorkbook wb, string caminho)
        {
            IWorksheet planilha = wb.Worksheets[0];
            int ultimaLinha = planilha.Rows.Count() - 1;
            planilha.Range["B7:B" + ultimaLinha].HorizontalAlignment = ExcelHAlign.HAlignCenter;
            planilha.Range["G7:L" + ultimaLinha].HorizontalAlignment = ExcelHAlign.HAlignCenter;
            wb.SaveAs(caminho);
        }

        public static void autofitColunas(IWorkbook wb, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = wb.Worksheets[0];
            planilha.AutofitColumn(1);
            planilha.AutofitColumn(2);
            planilha.AutofitColumn(7);
            planilha.AutofitColumn(8);
            planilha.AutofitColumn(9);
            planilha.AutofitColumn(10);
            planilha.AutofitColumn(11);
            planilha.AutofitColumn(12);
            planilha.Range["B" + linhas["primeiraLinha"] + ":H" + linhas["linhaIrpj"]].AutofitColumns();
            wb.SaveAs(caminho);
        }

        public static void aplicarBorda(IWorkbook wb, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = wb.Worksheets[0];
            planilha.Range["A6:L" + linhas["ultimaLinha"]].BorderInside();
            planilha.Range["A6:L" + linhas["ultimaLinha"]].BorderAround();
            planilha.Range["B" + linhas["primeiraLinha"] + ":F" + linhas["linhaIrpj"]].BorderInside();
            planilha.Range["B" + linhas["primeiraLinha"] + ":F" + linhas["linhaIrpj"]].BorderAround();
            wb.SaveAs(caminho);
        }

        public static string verificarMes(string m)
        {
            if (m.Equals("01")) return "JANEIRO";
            else if (m.Equals("02")) return "FEVEREIRO";
            else if (m.Equals("03")) return "MARÇO";
            else if (m.Equals("04")) return "ABRIL";
            else if (m.Equals("05")) return "MAIO";
            else if (m.Equals("06")) return "JUNHO";
            else if (m.Equals("07")) return "JULHO";
            else if (m.Equals("08")) return "AGOSTO";
            else if (m.Equals("09")) return "SETEMBRO";
            else if (m.Equals("10")) return "OUTUBRO";
            else if (m.Equals("11")) return "NOVEMBRO";
            else return "DEZEMBRO";
        }
    }
}
