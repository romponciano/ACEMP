using ACEMP.Model;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACEMP.Services;
using System.Windows.Forms;

namespace ACEMP.Services
{
    class ExcelFunctionalService
    {
        public static void criarModelo(string caminho, CSV csv)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2013;
            application.EnableIncrementalFormula = true;

            IWorkbook workbook = application.Workbooks.Open(caminho, ExcelOpenType.Automatic);
            IWorksheet planilha = workbook.Worksheets[0];
            planilha.EnableSheetCalculations();
            planilha.Name = ExcelLayoutService.verificarMes(csv.mes);
            workbook.SaveAs(caminho);

            gerarModeloBase(workbook, caminho);

            Dictionary<string, int> linhas = new Dictionary<string, int>();
            linhas.Add("ultimaLinha", planilha.Rows.Count());
            linhas.Add("primeiraLinha", linhas["ultimaLinha"] + 4);
            linhas.Add("linhaPis", linhas["primeiraLinha"] + 1);
            linhas.Add("linhaCofins", linhas["primeiraLinha"] + 2);
            linhas.Add("linhaCsll", linhas["primeiraLinha"] + 3);
            linhas.Add("linhaIrpj", linhas["primeiraLinha"] + 4);
            
            gerarHeader(workbook, csv, caminho);

            ExcelLayoutService.formatarValores(workbook, caminho, linhas);

            gerarColunaCliente(workbook, caminho, linhas, csv);
            
            gerarLiquido(workbook, caminho, linhas);

            gerarTotal(workbook, caminho, linhas);

            if (csv.temExterior) gerarNacional(workbook, caminho, linhas, csv);

            gerarColunaTributo(workbook, caminho, linhas);

            gerarColunaVrImposto(workbook, caminho, linhas, csv);

            gerarColunaCompensar(workbook, caminho, linhas);

            gerarColunaImposto(workbook, caminho, linhas);

            //gerarColunaVencimento(workbook, caminho, linhas);

            ExcelLayoutService.ajustarDatas(workbook, caminho, linhas);
            
            ExcelLayoutService.aplicarEstilos(workbook, caminho, linhas);

            ExcelLayoutService.autofitColunas(workbook, caminho, linhas);

            ExcelLayoutService.alinharDados(workbook, caminho);

            ExcelLayoutService.aplicarBorda(workbook, caminho, linhas);

            planilha = workbook.Worksheets[0];
            planilha.Zoom = 90;

            workbook.SaveAs(caminho);
            workbook.Close();
            excelEngine.Dispose();
        }

        private static void gerarModeloBase(IWorkbook wb, string caminho)
        {
            IWorksheet planilha = wb.Worksheets[0];
            int qtdTamanhoClienteNome = 3;
            for (int i = 1; i <= qtdTamanhoClienteNome; i++)
            {
                planilha.InsertColumn(4);
            }
            int qtdLinhasCabecalho = 5;
            for (int i = 1; i <= qtdLinhasCabecalho; i++)
            {
                planilha.InsertRow(1);
            }
            wb.SaveAs(caminho);
        }

        private static void gerarHeader(IWorkbook wb, CSV csv, string caminho)
        {
            IWorksheet planilha = wb.Worksheets[0];
            planilha.Range["A2"].Text = csv.nomeEmpresa.ToUpper();
            planilha.Range["A2:L2"].Merge();
            planilha.Range["A3"].Text = "CNPJ: " + csv.cnpj;
            planilha.Range["A3:L3"].Merge();
            planilha.Range["A4"].Text = "FATURAMENTO DE " + ExcelLayoutService.verificarMes(csv.mes) + " " + csv.ano;
            planilha.Range["A4:L4"].Merge();
            wb.SaveAs(caminho);
        }

        private static void gerarColunaCliente(IWorkbook wb, string caminho, Dictionary<string, int> linhas, CSV csv)
        {
            IWorksheet planilha = wb.Worksheets[0];
            if(csv.temExterior)
            {
                for (int i = 6; i < linhas["ultimaLinha"]; i++)
                {
                    planilha.Range["C" + i].HorizontalAlignment = ExcelHAlign.HAlignLeft;
                    planilha.Range["C" + i + ":F" + i].Merge();
                    if ( csv.clientesExterior.Contains(i) )
                    {
                        planilha.Range["C" + i].Text = "** EXTERIOR ** " + planilha.Range["C" + i].Value.ToString();
                        planilha.Range["C" + i].CellStyle.Font.Bold = true;
                    }
                }
            }
            else
            {
                for (int i = 6; i < linhas["ultimaLinha"]; i++)
                {
                    planilha.Range["C" + i].HorizontalAlignment = ExcelHAlign.HAlignLeft;
                    planilha.Range["C" + i + ":F" + i].Merge();
                }
            }
            wb.SaveAs(caminho);
        }

        private static void gerarLiquido(IWorkbook wb, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = wb.Worksheets[0];
            planilha.Range["L7:L" + linhas["ultimaLinha"]].Formula = "=$G7-$H7-$I7-$J7-$K7";
            wb.SaveAs(caminho);
        }

        private static void gerarTotal(IWorkbook wb, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = wb.Worksheets[0];
            planilha.Range["A" + linhas["ultimaLinha"]].Text = "TOTAL";
            planilha.Range["A" + linhas["ultimaLinha"] + ":F" + linhas["ultimaLinha"]].Merge();
            planilha.Range["G" + linhas["ultimaLinha"]].Formula = "=SUM(G7:G" + (linhas["ultimaLinha"] - 1).ToString() + ")";
            planilha.Range["H" + linhas["ultimaLinha"]].Formula = "=SUM(H7:H" + (linhas["ultimaLinha"] - 1).ToString() + ")";
            planilha.Range["I" + linhas["ultimaLinha"]].Formula = "=SUM(I7:I" + (linhas["ultimaLinha"] - 1).ToString() + ")";
            planilha.Range["J" + linhas["ultimaLinha"]].Formula = "=SUM(J7:J" + (linhas["ultimaLinha"] - 1).ToString() + ")";
            planilha.Range["K" + linhas["ultimaLinha"]].Formula = "=SUM(K7:K" + (linhas["ultimaLinha"] - 1).ToString() + ")";
            planilha.Range["L" + linhas["ultimaLinha"]].Formula = "=SUM(L7:L" + (linhas["ultimaLinha"] - 1).ToString() + ")";
            wb.SaveAs(caminho);
        }

        private static void gerarNacional(IWorkbook workbook, string caminho, Dictionary<string, int> linhas, CSV csv)
        {
            IWorksheet planilha = workbook.Worksheets[0];
            int linhaNacional = linhas["ultimaLinha"] + 1;
            planilha.Range["A" + linhaNacional].Text = "NACIONAL";
            planilha.Range["A" + linhaNacional].HorizontalAlignment = ExcelHAlign.HAlignCenter;
            planilha.Range["A" + linhaNacional + ":F" + linhaNacional].Merge();
            planilha.Range["A" + linhaNacional + ":L" + linhaNacional].CellStyle.Font.Bold = true;
            planilha.Range["A" + linhaNacional + ":L" + linhaNacional].CellStyle.FillPattern = ExcelPattern.Solid;
            planilha.Range["A" + linhaNacional + ":L" + linhaNacional].CellStyle.Interior.Color = Color.LightGray;
            planilha.Range["A" + linhaNacional + ":L" + linhaNacional].BorderInside();
            planilha.Range["A" + linhaNacional + ":L" + linhaNacional].BorderAround();
            planilha.Range["G" + linhaNacional].Text = planilha.Range["G"+(linhaNacional-1)].CalculatedValue;
            for (int i=7; i < linhas["ultimaLinha"]; i++)
            {
                if (csv.clientesExterior.Contains(i)) {
                    planilha.Range["G" + linhaNacional].Text = planilha.Range["G" + linhaNacional].Value + "-G" + i;
                }
            }
            planilha.Range["G" + linhaNacional].Formula = "=" + planilha.Range["G" + linhaNacional].Value;
            planilha.Range["H" + linhaNacional].Formula = "=H" + (linhaNacional - 1);
            planilha.Range["I" + linhaNacional].Formula = "=I" + (linhaNacional - 1);
            planilha.Range["J" + linhaNacional].Formula = "=J" + (linhaNacional - 1);
            planilha.Range["K" + linhaNacional].Formula = "=K" + (linhaNacional - 1);
            planilha.Range["L" + linhaNacional].Formula =
                "=G" + linhaNacional + "-H" + linhaNacional + "-I" + linhaNacional + "-J" + linhaNacional + "-K" + linhaNacional;
        }

        private static void gerarColunaTributo(IWorkbook workbook, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = workbook.Worksheets[0];
            planilha.Range["B" + linhas["primeiraLinha"]].Text = "Tributo";
            planilha.Range["B" + linhas["primeiraLinha"] + ":C" + linhas["primeiraLinha"]].Merge();
            planilha.Range["B" + linhas["linhaPis"]].Text = "PIS";
            planilha.Range["B" + linhas["linhaCofins"]].Text = "COFINS";
            planilha.Range["B" + linhas["linhaCsll"]].Text = "CSLL";
            planilha.Range["B" + linhas["linhaIrpj"]].Text = "IRPJ";
            planilha.Range["C" + linhas["linhaPis"]].Number = 0.0065; // 0.65 = 65% -> 0,65% = 0.0065
            planilha.Range["C" + linhas["linhaCofins"]].Number = 0.030;
            planilha.Range["C" + linhas["linhaCsll"]].Number = 0.0288;
            planilha.Range["C" + linhas["linhaIrpj"]].Number = 0.048;
            workbook.SaveAs(caminho);
        }

        private static void gerarColunaVrImposto(IWorkbook workbook, string caminho, Dictionary<string, int> linhas, CSV csv)
        {
            IWorksheet planilha = workbook.Worksheets[0];
            planilha.Range["D" + linhas["primeiraLinha"]].Text = "Vr Imposto";
            
            if(csv.temExterior)
            {
                planilha.Range["D" + linhas["linhaPis"]].Formula = "=G" + (linhas["ultimaLinha"]+1) + "*C" + linhas["linhaPis"];
                planilha.Range["D" + linhas["linhaCofins"]].Formula = "=G" + (linhas["ultimaLinha"]+1) + "*C" + linhas["linhaCofins"];
            }
            else
            {
                planilha.Range["D" + linhas["linhaPis"]].Formula = "=G" + linhas["ultimaLinha"] + "*C" + linhas["linhaPis"];
                planilha.Range["D" + linhas["linhaCofins"]].Formula = "=G" + linhas["ultimaLinha"] + "*C" + linhas["linhaCofins"];
            }
            planilha.Range["D" + linhas["linhaCsll"]].Formula = "=G" + linhas["ultimaLinha"] + "*C" + linhas["linhaCsll"];
            planilha.Range["D" + linhas["linhaIrpj"]].Formula = "=G" + linhas["ultimaLinha"] + "*C" + linhas["linhaIrpj"];
            workbook.SaveAs(caminho);
        }

        private static void gerarColunaCompensar(IWorkbook workbook, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = workbook.Worksheets[0];
            planilha.Range["E" + linhas["primeiraLinha"]].Text = "Compensar";
            planilha.Range["E" + linhas["linhaPis"]].Formula = "=I" + linhas["ultimaLinha"];
            planilha.Range["E" + linhas["linhaCofins"]].Formula = "=J" + linhas["ultimaLinha"];
            planilha.Range["E" + linhas["linhaCsll"]].Formula = "=K" + linhas["ultimaLinha"];
            planilha.Range["E" + linhas["linhaIrpj"]].Formula = "=H" + linhas["ultimaLinha"];
            workbook.SaveAs(caminho);
        }

        private static void gerarColunaImposto(IWorkbook workbook, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = workbook.Worksheets[0];
            planilha.Range["F" + linhas["primeiraLinha"]].Text = "Imposto a pagar";
            planilha.Range["F" + linhas["linhaPis"]].Formula = "=D" + linhas["linhaPis"] + "-E" + linhas["linhaPis"];
            planilha.Range["F" + linhas["linhaCofins"]].Formula = "=D" + linhas["linhaCofins"] + "-E" + linhas["linhaCofins"];
            planilha.Range["F" + linhas["linhaCsll"]].Formula = "=D" + linhas["linhaCsll"] + "-E" + linhas["linhaCsll"];
            planilha.Range["F" + linhas["linhaIrpj"]].Formula = "=D" + linhas["linhaIrpj"] + "-E" + linhas["linhaIrpj"];
        }

        private static void gerarColunaVencimento(IWorkbook workbook, string caminho, Dictionary<string, int> linhas)
        {
            IWorksheet planilha = workbook.Worksheets[0];
            planilha.Range["G" + linhas["primeiraLinha"]].Text = "Vencimento";
            planilha.Range["G" + linhas["primeiraLinha"] + ":H" + linhas["primeiraLinha"]].Merge();
            planilha.Range["G" + linhas["linhaPis"] + ":H" + linhas["linhaPis"]].Merge();
            planilha.Range["G" + linhas["linhaCofins"] + ":H" + linhas["linhaCofins"]].Merge();
            planilha.Range["G" + linhas["linhaCsll"] + ":H" + linhas["linhaCsll"]].Merge();
            planilha.Range["G" + linhas["linhaIrpj"] + ":H" + linhas["linhaIrpj"]].Merge();
            workbook.SaveAs(caminho);
        }
    }
}
