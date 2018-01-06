using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ACEMP.Services;
using ACEMP.Model;
using System.Text.RegularExpressions;

namespace ACEMP
{
    public partial class Form1 : Form
    {
        string CAMINHO_SALVAR;

        public Form1()
        {
            InitializeComponent();
        }

        private static string verificarNome(string fileName)
        {
            return Regex.Replace(fileName.Trim(), "[^A-Za-z0-9_. ]+", "").ToLower();
        }

        private void btnSelecionarCsv_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog caminhoDialogo = new OpenFileDialog()
                {
                    Filter = "CSV|*.csv",
                    ValidateNames = true,
                    Multiselect = true
                })
                {
                    if (caminhoDialogo.ShowDialog() == DialogResult.OK)
                    {
                        foreach(String arquivo in caminhoDialogo.FileNames) {
                            DataTable original = ConversionService.csv2datatable(arquivo);

                            CSV csv = CSVService.gerarcsv(original);

                            string caminho = CAMINHO_SALVAR + verificarNome(csv.nomeEmpresa) + ".xlsx";

                            ConversionService.datatable2xls(csv.csvFinal, caminho);

                            ExcelFunctionalService.criarModelo(caminho, csv);
                        }

                        MessageBox.Show("Processo finalizado com sucesso!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSelecionarSalvar_Click(object sender, EventArgs e)
        {
            try
            {
                using (FolderBrowserDialog caminhoPasta = new FolderBrowserDialog())
                {
                    if (caminhoPasta.ShowDialog() == DialogResult.OK)
                    {
                        CAMINHO_SALVAR = caminhoPasta.SelectedPath + "\\";
                        lblLocalSalvar.Text = CAMINHO_SALVAR;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDoar_click(object sender, EventArgs e)
        {
            DialogResult escolha = MessageBox.Show(
                "Gostou e quer ajudar a manter a manutenção do sistema? \nFaça uma doação em qualquer valor! :)",
                "Doe para ajudar o projeto",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (escolha == DialogResult.Yes)
            {

            }
        }

        private void btnSobre_Click(object sender, EventArgs e)
        {
            DialogResult escolha = MessageBox.Show(
                "Sua versão é: v1.0. \n\nPara se manter atualizado, visite a página do projeto clicando em \"Sim\".",
                "Acesse a página do projeto",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (escolha == DialogResult.Yes)
            {
                
            }
        }

        private void btnAjuda_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Para reportar algum projeto, envie um print descrevendo a situação para o email: \nromuloponciano@id.uff.br",
                "Reporte um problema",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
    }
}

