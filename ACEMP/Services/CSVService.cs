using ACEMP.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ACEMP.Services
{
    class CSVService
    {
        public static CSV gerarcsv(DataTable original, String arquivo)
        {
            DataTable f = new DataTable();
            f.Columns.Add("Data");
            f.Columns.Add("NF");
            f.Columns.Add("Cliente");
            f.Columns.Add("R$");
            f.Columns.Add("Vr Irf");
            f.Columns.Add("PIS");
            f.Columns.Add("COFINS");
            f.Columns.Add("CSLL");
            f.Columns.Add("Líquido");

            int aux = 0;
            List<string> nfsAux = gerarNumerosNfs(arquivo);
            for (int i = 0; i < original.Rows.Count; i++)
            {
                f.Rows.Add();
                f.Rows[i]["Data"] = original.Rows[i][4];
                f.Rows[i]["NF"] = nfsAux[i];
                f.Rows[i]["Cliente"] = original.Rows[i][29];
                f.Rows[i]["R$"] = original.Rows[i][50];
                f.Rows[i]["Vr Irf"] = original.Rows[i][57];
                f.Rows[i]["PIS"] = original.Rows[i][58];
                f.Rows[i]["COFINS"] = original.Rows[i][54];
                f.Rows[i]["CSLL"] = original.Rows[i][55];
                f.Rows[i]["Líquido"] =
                    float.Parse(f.Rows[i]["R$"].ToString()) -
                    float.Parse(f.Rows[i]["Vr Irf"].ToString()) -
                    float.Parse(f.Rows[i]["COFINS"].ToString()) -
                    float.Parse(f.Rows[i]["CSLL"].ToString());
                aux = i;
            }

            if (aux <= 2) aux = 0;
            else aux = 1;
            string nome = original.Rows[aux][13].ToString();
            string cnpj = original.Rows[aux][10].ToString();
            string[] dataCompleta = original.Rows[aux][4].ToString().Split('/');
            string mes = dataCompleta.GetValue(1).ToString();
            string ano = dataCompleta.GetValue(2).ToString().Split(' ').GetValue(0).ToString();

            CSV csv = new CSV(f, nome, mes, ano, cnpj);

            aux = 7;
            foreach (DataRow linha in original.Rows)
            {
                if (linha[26].ToString().Equals(""))
                {
                    csv.clientesExterior.Add(aux);
                }
                aux++;
            }
            csv.clientesExterior.RemoveAt(csv.clientesExterior.Count() - 1);
            if (csv.clientesExterior.Count > 0) csv.temExterior = true;

            return csv;
        }

        public static List<string> gerarNumerosNfs(String arquivo)
        {
            DataTable numerosnfs = ConversionService.csv2numeronfs(arquivo);
            List<string> aux = new List<string>();
            for (int i = 0; i < numerosnfs.Rows.Count; i++)
                aux.Add(numerosnfs.Rows[i][1].ToString());
            return aux;
        }
    }
}
