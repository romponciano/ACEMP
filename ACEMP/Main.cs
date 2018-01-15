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
using System.Deployment.Application;

namespace ACEMP
{
    public partial class Main : Form
    {
        string CAMINHO_SALVAR;

        public Main()
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
                System.Diagnostics.Process.Start("https://www.vakinha.com.br/vaquinha/projeto-acemp");
            }
        }

        private void btnSobre_Click(object sender, EventArgs e)
        {
            DialogResult escolha = MessageBox.Show(
                "Sua versão é: v1.1.3 \n\nPara se manter atualizado, visite a página do projeto clicando em \"Sim\".",
                "Acesse a página do projeto",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (escolha == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start("https://github.com/rponciano/ACEMP");
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

        private void btnUpdate_click(object sender, EventArgs e)
        {
            UpdateCheckInfo info = null;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                try
                {
                    info = ad.CheckForDetailedUpdate();

                }
                catch (DeploymentDownloadException dde)
                {
                    MessageBox.Show("The new version of the application cannot be downloaded at this time. \n\nPlease check your network connection, or try again later. Error: " + dde.Message);
                    return;
                }
                catch (InvalidDeploymentException ide)
                {
                    MessageBox.Show("Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error: " + ide.Message);
                    return;
                }
                catch (InvalidOperationException ioe)
                {
                    MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message);
                    return;
                }

                if (info.UpdateAvailable)
                {
                    Boolean doUpdate = true;

                    if (!info.IsUpdateRequired)
                    {
                        DialogResult dr = MessageBox.Show("An update is available. Would you like to update the application now?", "Update Available", MessageBoxButtons.OKCancel);
                        if (!(DialogResult.OK == dr))
                        {
                            doUpdate = false;
                        }
                    }
                    else
                    {
                        // Display a message that the app MUST reboot. Display the minimum required version.
                        MessageBox.Show("This application has detected a mandatory update from your current " +
                            "version to version " + info.MinimumRequiredVersion.ToString() +
                            ". The application will now install the update and restart.",
                            "Update Available", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (doUpdate)
                    {
                        try
                        {
                            ad.Update();
                            MessageBox.Show("The application has been upgraded, and will now restart.");
                            Application.Restart();
                        }
                        catch (DeploymentDownloadException dde)
                        {
                            MessageBox.Show("Cannot install the latest version of the application. \n\nPlease check your network connection, or try again later. Error: " + dde);
                            return;
                        }
                    }
                }
            }
        }
    }
}

/*
Personally I'm using a very simple methodology for any kind of auto-update:

Have an installer
Check the new version (simple WebClient and compare numbers with your current AssemblyVersion)
If the version is higher download the latest installer (should be over SSL for security reasons)
Run the downloaded installer and close the application. (in this stage you need to have admin privileges if your installer requires you to be an admin)
The installer should take care of the rest. This way you'll always have the latest version and an installer with a latest version.
 */