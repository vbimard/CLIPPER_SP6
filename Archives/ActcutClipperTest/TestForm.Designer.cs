namespace ActcutClipperTest
{
    partial class TestForm
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.BtnInitApi = new System.Windows.Forms.Button();
            this.BtnGetQuoteApi = new System.Windows.Forms.Button();
            this.BtnExportQuoteApi = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Init = new System.Windows.Forms.Button();
            this.Reset = new System.Windows.Forms.Button();
            this.ExportUI = new System.Windows.Forms.Button();
            this.Export = new System.Windows.Forms.Button();
            this.txtQuoteid = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboDataBaseList = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.linkemf = new System.Windows.Forms.LinkLabel();
            this.linkquote = new System.Windows.Forms.LinkLabel();
            this.status = new System.Windows.Forms.Label();
            this.Select_Quote = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnInitApi
            // 
            this.BtnInitApi.Location = new System.Drawing.Point(0, 0);
            this.BtnInitApi.Name = "BtnInitApi";
            this.BtnInitApi.Size = new System.Drawing.Size(75, 23);
            this.BtnInitApi.TabIndex = 0;
            // 
            // BtnGetQuoteApi
            // 
            this.BtnGetQuoteApi.Location = new System.Drawing.Point(0, 0);
            this.BtnGetQuoteApi.Name = "BtnGetQuoteApi";
            this.BtnGetQuoteApi.Size = new System.Drawing.Size(75, 23);
            this.BtnGetQuoteApi.TabIndex = 0;
            // 
            // BtnExportQuoteApi
            // 
            this.BtnExportQuoteApi.Location = new System.Drawing.Point(0, 0);
            this.BtnExportQuoteApi.Name = "BtnExportQuoteApi";
            this.BtnExportQuoteApi.Size = new System.Drawing.Size(75, 23);
            this.BtnExportQuoteApi.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Init);
            this.groupBox1.Controls.Add(this.Reset);
            this.groupBox1.Controls.Add(this.ExportUI);
            this.groupBox1.Location = new System.Drawing.Point(289, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(142, 151);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            // 
            // Init
            // 
            this.Init.Location = new System.Drawing.Point(13, 12);
            this.Init.Name = "Init";
            this.Init.Size = new System.Drawing.Size(112, 23);
            this.Init.TabIndex = 16;
            this.Init.Text = "Init";
            this.Init.UseVisualStyleBackColor = true;
            this.Init.Click += new System.EventHandler(this.Init_Click);
            // 
            // Reset
            // 
            this.Reset.Location = new System.Drawing.Point(14, 120);
            this.Reset.Name = "Reset";
            this.Reset.Size = new System.Drawing.Size(112, 23);
            this.Reset.TabIndex = 21;
            this.Reset.Text = "Reset";
            this.Reset.UseVisualStyleBackColor = true;
            this.Reset.Click += new System.EventHandler(this.Reset_Click);
            // 
            // ExportUI
            // 
            this.ExportUI.Location = new System.Drawing.Point(14, 41);
            this.ExportUI.Name = "ExportUI";
            this.ExportUI.Size = new System.Drawing.Size(112, 23);
            this.ExportUI.TabIndex = 22;
            this.ExportUI.Text = " Export UI";
            this.ExportUI.UseVisualStyleBackColor = true;
            this.ExportUI.Click += new System.EventHandler(this.ExportUI_Click);
            // 
            // Export
            // 
            this.Export.Location = new System.Drawing.Point(302, 70);
            this.Export.Name = "Export";
            this.Export.Size = new System.Drawing.Size(112, 23);
            this.Export.TabIndex = 17;
            this.Export.Text = "Export";
            this.Export.UseVisualStyleBackColor = true;
            this.Export.Click += new System.EventHandler(this.Export_Click);
            // 
            // txtQuoteid
            // 
            this.txtQuoteid.Location = new System.Drawing.Point(213, 39);
            this.txtQuoteid.Name = "txtQuoteid";
            this.txtQuoteid.Size = new System.Drawing.Size(52, 20);
            this.txtQuoteid.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(78, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "IdDevis";
            // 
            // comboDataBaseList
            // 
            this.comboDataBaseList.FormattingEnabled = true;
            this.comboDataBaseList.Location = new System.Drawing.Point(81, 12);
            this.comboDataBaseList.Name = "comboDataBaseList";
            this.comboDataBaseList.Size = new System.Drawing.Size(184, 21);
            this.comboDataBaseList.TabIndex = 15;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 102);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(37, 13);
            this.label3.TabIndex = 19;
            this.label3.Text = "Devis:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 124);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(28, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "Emf:";
            // 
            // linkemf
            // 
            this.linkemf.AutoSize = true;
            this.linkemf.Location = new System.Drawing.Point(58, 124);
            this.linkemf.Name = "linkemf";
            this.linkemf.Size = new System.Drawing.Size(45, 13);
            this.linkemf.TabIndex = 17;
            this.linkemf.TabStop = true;
            this.linkemf.Text = "C:\\temp";
            // 
            // linkquote
            // 
            this.linkquote.AutoSize = true;
            this.linkquote.Location = new System.Drawing.Point(58, 102);
            this.linkquote.Name = "linkquote";
            this.linkquote.Size = new System.Drawing.Size(45, 13);
            this.linkquote.TabIndex = 16;
            this.linkquote.TabStop = true;
            this.linkquote.Text = "C:\\temp";
            // 
            // status
            // 
            this.status.AutoSize = true;
            this.status.Location = new System.Drawing.Point(14, 81);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(16, 13);
            this.status.TabIndex = 14;
            this.status.Text = "...";
            // 
            // Select_Quote
            // 
            this.Select_Quote.Enabled = false;
            this.Select_Quote.Location = new System.Drawing.Point(81, 70);
            this.Select_Quote.Name = "Select_Quote";
            this.Select_Quote.Size = new System.Drawing.Size(184, 23);
            this.Select_Quote.TabIndex = 21;
            this.Select_Quote.Text = "Select_Quote";
            this.Select_Quote.UseVisualStyleBackColor = true;
            this.Select_Quote.Click += new System.EventHandler(this.Select_Quote_Click);
            // 
            // TestForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(449, 155);
            this.Controls.Add(this.Select_Quote);
            this.Controls.Add(this.Export);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtQuoteid);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboDataBaseList);
            this.Controls.Add(this.linkemf);
            this.Controls.Add(this.linkquote);
            this.Controls.Add(this.status);
            this.Controls.Add(this.groupBox1);
            this.Name = "TestForm";
            this.Text = "Test Form";
            this.Load += new System.EventHandler(this.TestForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnInitApi;
        private System.Windows.Forms.Button BtnGetQuoteApi;
        private System.Windows.Forms.Button BtnExportQuoteApi;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.LinkLabel linkemf;
        private System.Windows.Forms.LinkLabel linkquote;
        private System.Windows.Forms.ComboBox comboDataBaseList;
        private System.Windows.Forms.Label status;
        private System.Windows.Forms.TextBox txtQuoteid;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button Init;
        private System.Windows.Forms.Button Export;
        private System.Windows.Forms.Button Reset;
        private System.Windows.Forms.Button ExportUI;
        private System.Windows.Forms.Button Select_Quote;
    }
}

