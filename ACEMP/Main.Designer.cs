namespace ACEMP
{
    partial class Main
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.btnSelecionarCsv = new System.Windows.Forms.Button();
            this.btnSeleionarSaida = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblLocalSalvar = new System.Windows.Forms.Label();
            this.lblLocalSalvo = new System.Windows.Forms.Label();
            this.lblSalvo = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuAjuda = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSobre = new System.Windows.Forms.ToolStripMenuItem();
            this.menuDoar = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSelecionarCsv
            // 
            this.btnSelecionarCsv.Location = new System.Drawing.Point(12, 113);
            this.btnSelecionarCsv.Name = "btnSelecionarCsv";
            this.btnSelecionarCsv.Size = new System.Drawing.Size(131, 51);
            this.btnSelecionarCsv.TabIndex = 0;
            this.btnSelecionarCsv.Text = "Selecionar CSV\'s de entrada";
            this.btnSelecionarCsv.UseVisualStyleBackColor = true;
            this.btnSelecionarCsv.Click += new System.EventHandler(this.btnSelecionarCsv_Click);
            // 
            // btnSeleionarSaida
            // 
            this.btnSeleionarSaida.Location = new System.Drawing.Point(12, 27);
            this.btnSeleionarSaida.Name = "btnSeleionarSaida";
            this.btnSeleionarSaida.Size = new System.Drawing.Size(131, 51);
            this.btnSeleionarSaida.TabIndex = 1;
            this.btnSeleionarSaida.Text = "Selecionar local de saída";
            this.btnSeleionarSaida.UseVisualStyleBackColor = true;
            this.btnSeleionarSaida.Click += new System.EventHandler(this.btnSelecionarSalvar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(149, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(133, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Arquivos serão salvos em: ";
            // 
            // lblLocalSalvar
            // 
            this.lblLocalSalvar.AutoSize = true;
            this.lblLocalSalvar.Location = new System.Drawing.Point(149, 56);
            this.lblLocalSalvar.Name = "lblLocalSalvar";
            this.lblLocalSalvar.Size = new System.Drawing.Size(0, 13);
            this.lblLocalSalvar.TabIndex = 4;
            // 
            // lblLocalSalvo
            // 
            this.lblLocalSalvo.AutoSize = true;
            this.lblLocalSalvo.Location = new System.Drawing.Point(149, 142);
            this.lblLocalSalvo.Name = "lblLocalSalvo";
            this.lblLocalSalvo.Size = new System.Drawing.Size(0, 13);
            this.lblLocalSalvo.TabIndex = 6;
            // 
            // lblSalvo
            // 
            this.lblSalvo.AutoSize = true;
            this.lblSalvo.Location = new System.Drawing.Point(149, 123);
            this.lblSalvo.Name = "lblSalvo";
            this.lblSalvo.Size = new System.Drawing.Size(0, 13);
            this.lblSalvo.TabIndex = 5;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuAjuda,
            this.menuSobre,
            this.menuDoar});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.menuStrip1.Size = new System.Drawing.Size(559, 24);
            this.menuStrip1.TabIndex = 7;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuAjuda
            // 
            this.menuAjuda.Name = "menuAjuda";
            this.menuAjuda.Size = new System.Drawing.Size(50, 20);
            this.menuAjuda.Text = "Ajuda";
            this.menuAjuda.Click += new System.EventHandler(this.btnAjuda_Click);
            // 
            // menuSobre
            // 
            this.menuSobre.Name = "menuSobre";
            this.menuSobre.Size = new System.Drawing.Size(49, 20);
            this.menuSobre.Text = "Sobre";
            this.menuSobre.Click += new System.EventHandler(this.btnSobre_Click);
            // 
            // menuDoar
            // 
            this.menuDoar.Name = "menuDoar";
            this.menuDoar.Size = new System.Drawing.Size(44, 20);
            this.menuDoar.Text = "Doar";
            this.menuDoar.Click += new System.EventHandler(this.btnDoar_click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(559, 183);
            this.Controls.Add(this.lblLocalSalvo);
            this.Controls.Add(this.lblSalvo);
            this.Controls.Add(this.lblLocalSalvar);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSeleionarSaida);
            this.Controls.Add(this.btnSelecionarCsv);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ACEMP";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelecionarCsv;
        private System.Windows.Forms.Button btnSeleionarSaida;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblLocalSalvar;
        private System.Windows.Forms.Label lblLocalSalvo;
        private System.Windows.Forms.Label lblSalvo;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuSobre;
        private System.Windows.Forms.ToolStripMenuItem menuAjuda;
        private System.Windows.Forms.ToolStripMenuItem menuDoar;
    }
}

