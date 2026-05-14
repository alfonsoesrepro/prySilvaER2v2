namespace prySilvaER2v2
{
    partial class frmMigracion
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMigracion));
            this.cmdIniciar = new System.Windows.Forms.Button();
            this.lblInfo = new System.Windows.Forms.Label();
            this.txtInfo = new System.Windows.Forms.TextBox();
            this.pbGH = new System.Windows.Forms.PictureBox();
            this.cmdImportar = new System.Windows.Forms.Button();
            this.lbInfo = new System.Windows.Forms.ListBox();
            ((System.ComponentModel.ISupportInitialize)(this.pbGH)).BeginInit();
            this.SuspendLayout();
            // 
            // cmdIniciar
            // 
            this.cmdIniciar.Location = new System.Drawing.Point(186, 109);
            this.cmdIniciar.Name = "cmdIniciar";
            this.cmdIniciar.Size = new System.Drawing.Size(120, 27);
            this.cmdIniciar.TabIndex = 0;
            this.cmdIniciar.Text = "Iniciar &Migración";
            this.cmdIniciar.UseVisualStyleBackColor = true;
            this.cmdIniciar.Click += new System.EventHandler(this.cmdIniciar_Click);
            // 
            // lblInfo
            // 
            this.lblInfo.AutoSize = true;
            this.lblInfo.Location = new System.Drawing.Point(25, 167);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(65, 13);
            this.lblInfo.TabIndex = 1;
            this.lblInfo.Text = "Información:";
            // 
            // txtInfo
            // 
            this.txtInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtInfo.Location = new System.Drawing.Point(28, 194);
            this.txtInfo.Multiline = true;
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.Size = new System.Drawing.Size(278, 146);
            this.txtInfo.TabIndex = 2;
            // 
            // pbGH
            // 
            this.pbGH.Image = global::prySilvaER2v2.Properties.Resources.ChatGPT_Image_14_may_2026__04_51_57_a_m_;
            this.pbGH.Location = new System.Drawing.Point(132, 12);
            this.pbGH.Name = "pbGH";
            this.pbGH.Size = new System.Drawing.Size(71, 65);
            this.pbGH.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbGH.TabIndex = 3;
            this.pbGH.TabStop = false;
            // 
            // cmdImportar
            // 
            this.cmdImportar.Location = new System.Drawing.Point(28, 109);
            this.cmdImportar.Name = "cmdImportar";
            this.cmdImportar.Size = new System.Drawing.Size(120, 27);
            this.cmdImportar.TabIndex = 5;
            this.cmdImportar.Text = "Importar &Archivos";
            this.cmdImportar.UseVisualStyleBackColor = true;
            this.cmdImportar.Click += new System.EventHandler(this.cmdImportar_Click);
            // 
            // lbInfo
            // 
            this.lbInfo.FormattingEnabled = true;
            this.lbInfo.Location = new System.Drawing.Point(28, 190);
            this.lbInfo.Name = "lbInfo";
            this.lbInfo.Size = new System.Drawing.Size(278, 160);
            this.lbInfo.TabIndex = 6;
            // 
            // frmMigracion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(334, 361);
            this.Controls.Add(this.lbInfo);
            this.Controls.Add(this.cmdImportar);
            this.Controls.Add(this.pbGH);
            this.Controls.Add(this.txtInfo);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.cmdIniciar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMigracion";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Migración de Datos";
           
            ((System.ComponentModel.ISupportInitialize)(this.pbGH)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cmdIniciar;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.TextBox txtInfo;
        private System.Windows.Forms.PictureBox pbGH;
        private System.Windows.Forms.Button cmdImportar;
        private System.Windows.Forms.ListBox lbInfo;
    }
}

