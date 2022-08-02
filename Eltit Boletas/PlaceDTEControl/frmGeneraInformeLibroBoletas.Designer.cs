namespace Eltit
{
    partial class frmGeneraInformeLibroBoletas
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGeneraInformeLibroBoletas));
            this.telerikMetroBlueTheme1 = new Telerik.WinControls.Themes.TelerikMetroBlueTheme();
            this.radStatusStrip1 = new Telerik.WinControls.UI.RadStatusStrip();
            this.lblInformacion = new Telerik.WinControls.UI.RadLabelElement();
            this.RadPageView1 = new Telerik.WinControls.UI.RadPageView();
            this.RadPageViewPage1 = new Telerik.WinControls.UI.RadPageViewPage();
            this.rbRectifica = new Telerik.WinControls.UI.RadRadioButton();
            this.rbEspecial = new Telerik.WinControls.UI.RadRadioButton();
            this.label6 = new System.Windows.Forms.Label();
            this.rbMensual = new Telerik.WinControls.UI.RadRadioButton();
            this.radButton1 = new Telerik.WinControls.UI.RadButton();
            this.ddLAno = new Telerik.WinControls.UI.RadDropDownList();
            this.label5 = new System.Windows.Forms.Label();
            this.ddlMes = new Telerik.WinControls.UI.RadDropDownList();
            this.label4 = new System.Windows.Forms.Label();
            this.ddLocales = new Telerik.WinControls.UI.RadDropDownList();
            this.label2 = new System.Windows.Forms.Label();
            this.lblRut = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblNombreEmpresa = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGenerar = new Telerik.WinControls.UI.RadButton();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RadPageView1)).BeginInit();
            this.RadPageView1.SuspendLayout();
            this.RadPageViewPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rbRectifica)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rbEspecial)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rbMensual)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ddLAno)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ddlMes)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ddLocales)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnGenerar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // radStatusStrip1
            // 
            this.radStatusStrip1.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.lblInformacion});
            this.radStatusStrip1.Location = new System.Drawing.Point(0, 265);
            this.radStatusStrip1.Name = "radStatusStrip1";
            this.radStatusStrip1.Size = new System.Drawing.Size(733, 25);
            this.radStatusStrip1.TabIndex = 1;
            this.radStatusStrip1.Text = "radStatusStrip1";
            this.radStatusStrip1.ThemeName = "TelerikMetroBlue";
            // 
            // lblInformacion
            // 
            this.lblInformacion.Name = "lblInformacion";
            this.radStatusStrip1.SetSpring(this.lblInformacion, false);
            this.lblInformacion.Text = "";
            this.lblInformacion.TextWrap = true;
            // 
            // RadPageView1
            // 
            this.RadPageView1.BackColor = System.Drawing.Color.DarkGray;
            this.RadPageView1.Controls.Add(this.RadPageViewPage1);
            this.RadPageView1.Location = new System.Drawing.Point(12, 4);
            this.RadPageView1.Name = "RadPageView1";
            this.RadPageView1.SelectedPage = this.RadPageViewPage1;
            this.RadPageView1.Size = new System.Drawing.Size(709, 247);
            this.RadPageView1.TabIndex = 7;
            this.RadPageView1.Text = "Documento";
            this.RadPageView1.ThemeName = "TelerikMetroBlue";
            ((Telerik.WinControls.UI.RadPageViewStripElement)(this.RadPageView1.GetChildAt(0))).StripButtons = Telerik.WinControls.UI.StripViewButtons.None;
            // 
            // RadPageViewPage1
            // 
            this.RadPageViewPage1.Controls.Add(this.rbRectifica);
            this.RadPageViewPage1.Controls.Add(this.rbEspecial);
            this.RadPageViewPage1.Controls.Add(this.label6);
            this.RadPageViewPage1.Controls.Add(this.rbMensual);
            this.RadPageViewPage1.Controls.Add(this.radButton1);
            this.RadPageViewPage1.Controls.Add(this.ddLAno);
            this.RadPageViewPage1.Controls.Add(this.label5);
            this.RadPageViewPage1.Controls.Add(this.ddlMes);
            this.RadPageViewPage1.Controls.Add(this.label4);
            this.RadPageViewPage1.Controls.Add(this.ddLocales);
            this.RadPageViewPage1.Controls.Add(this.label2);
            this.RadPageViewPage1.Controls.Add(this.lblRut);
            this.RadPageViewPage1.Controls.Add(this.label3);
            this.RadPageViewPage1.Controls.Add(this.lblNombreEmpresa);
            this.RadPageViewPage1.Controls.Add(this.label1);
            this.RadPageViewPage1.Controls.Add(this.btnGenerar);
            this.RadPageViewPage1.Controls.Add(this.pictureBox1);
            this.RadPageViewPage1.ItemSize = new System.Drawing.SizeF(120F, 25F);
            this.RadPageViewPage1.Location = new System.Drawing.Point(5, 31);
            this.RadPageViewPage1.Name = "RadPageViewPage1";
            this.RadPageViewPage1.Size = new System.Drawing.Size(699, 211);
            this.RadPageViewPage1.Text = "Procesando Cliente: ";
            // 
            // rbRectifica
            // 
            this.rbRectifica.Location = new System.Drawing.Point(315, 113);
            this.rbRectifica.Name = "rbRectifica";
            this.rbRectifica.Size = new System.Drawing.Size(68, 19);
            this.rbRectifica.TabIndex = 27;
            this.rbRectifica.TabStop = false;
            this.rbRectifica.Text = "Rectifica";
            this.rbRectifica.ThemeName = "TelerikMetroBlue";
            // 
            // rbEspecial
            // 
            this.rbEspecial.CheckState = System.Windows.Forms.CheckState.Checked;
            this.rbEspecial.Location = new System.Drawing.Point(240, 113);
            this.rbEspecial.Name = "rbEspecial";
            this.rbEspecial.Size = new System.Drawing.Size(66, 19);
            this.rbEspecial.TabIndex = 26;
            this.rbEspecial.Text = "Especial";
            this.rbEspecial.ThemeName = "TelerikMetroBlue";
            this.rbEspecial.ToggleState = Telerik.WinControls.Enumerations.ToggleState.On;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label6.Location = new System.Drawing.Point(64, 106);
            this.label6.Name = "label6";
            this.label6.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label6.Size = new System.Drawing.Size(85, 25);
            this.label6.TabIndex = 25;
            this.label6.Text = "Tipo Libro";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // rbMensual
            // 
            this.rbMensual.Location = new System.Drawing.Point(162, 113);
            this.rbMensual.Name = "rbMensual";
            this.rbMensual.Size = new System.Drawing.Size(69, 19);
            this.rbMensual.TabIndex = 24;
            this.rbMensual.TabStop = false;
            this.rbMensual.Text = "Mensual";
            this.rbMensual.ThemeName = "TelerikMetroBlue";
            // 
            // radButton1
            // 
            this.radButton1.Enabled = false;
            this.radButton1.Image = ((System.Drawing.Image)(resources.GetObject("radButton1.Image")));
            this.radButton1.Location = new System.Drawing.Point(238, 156);
            this.radButton1.Name = "radButton1";
            this.radButton1.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.radButton1.Size = new System.Drawing.Size(162, 41);
            this.radButton1.TabIndex = 17;
            this.radButton1.Text = "       Generar Libro";
            this.radButton1.ThemeName = "TelerikMetroBlue";
            this.radButton1.Visible = false;
            this.radButton1.Click += new System.EventHandler(this.radButton1_Click);
            // 
            // ddLAno
            // 
            this.ddLAno.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList;
            this.ddLAno.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.ddLAno.Location = new System.Drawing.Point(365, 69);
            this.ddLAno.Name = "ddLAno";
            this.ddLAno.Size = new System.Drawing.Size(82, 28);
            this.ddLAno.TabIndex = 23;
            this.ddLAno.ThemeName = "TelerikMetroBlue";
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label5.Location = new System.Drawing.Point(274, 71);
            this.label5.Name = "label5";
            this.label5.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label5.Size = new System.Drawing.Size(85, 24);
            this.label5.TabIndex = 22;
            this.label5.Text = "Año";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // ddlMes
            // 
            this.ddlMes.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList;
            this.ddlMes.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.ddlMes.Location = new System.Drawing.Point(155, 68);
            this.ddlMes.Name = "ddlMes";
            this.ddlMes.Size = new System.Drawing.Size(113, 28);
            this.ddlMes.TabIndex = 21;
            this.ddlMes.ThemeName = "TelerikMetroBlue";
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label4.Location = new System.Drawing.Point(64, 70);
            this.label4.Name = "label4";
            this.label4.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label4.Size = new System.Drawing.Size(85, 25);
            this.label4.TabIndex = 20;
            this.label4.Text = "Mes";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // ddLocales
            // 
            this.ddLocales.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList;
            this.ddLocales.Font = new System.Drawing.Font("Segoe UI", 11F);
            this.ddLocales.Location = new System.Drawing.Point(155, 5);
            this.ddLocales.Name = "ddLocales";
            this.ddLocales.Size = new System.Drawing.Size(476, 28);
            this.ddLocales.TabIndex = 19;
            this.ddLocales.ThemeName = "TelerikMetroBlue";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label2.Location = new System.Drawing.Point(64, 5);
            this.label2.Name = "label2";
            this.label2.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label2.Size = new System.Drawing.Size(85, 25);
            this.label2.TabIndex = 18;
            this.label2.Text = "Local";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblRut
            // 
            this.lblRut.AutoEllipsis = true;
            this.lblRut.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblRut.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblRut.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRut.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblRut.Location = new System.Drawing.Point(155, 39);
            this.lblRut.Name = "lblRut";
            this.lblRut.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblRut.Size = new System.Drawing.Size(113, 24);
            this.lblRut.TabIndex = 4;
            this.lblRut.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label3.Location = new System.Drawing.Point(64, 38);
            this.label3.Name = "label3";
            this.label3.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label3.Size = new System.Drawing.Size(85, 24);
            this.label3.TabIndex = 3;
            this.label3.Text = "Rut";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblNombreEmpresa
            // 
            this.lblNombreEmpresa.AutoEllipsis = true;
            this.lblNombreEmpresa.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblNombreEmpresa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNombreEmpresa.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNombreEmpresa.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblNombreEmpresa.Location = new System.Drawing.Point(365, 39);
            this.lblNombreEmpresa.Name = "lblNombreEmpresa";
            this.lblNombreEmpresa.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblNombreEmpresa.Size = new System.Drawing.Size(266, 24);
            this.lblNombreEmpresa.TabIndex = 2;
            this.lblNombreEmpresa.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label1.Location = new System.Drawing.Point(274, 39);
            this.label1.Name = "label1";
            this.label1.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label1.Size = new System.Drawing.Size(85, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "Razon Social";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnGenerar
            // 
            this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
            this.btnGenerar.Location = new System.Drawing.Point(60, 156);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.btnGenerar.Size = new System.Drawing.Size(160, 41);
            this.btnGenerar.TabIndex = 16;
            this.btnGenerar.Text = "       Generar Informe";
            this.btnGenerar.ThemeName = "TelerikMetroBlue";
            this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(3, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(55, 56);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // timer1
            // 
            this.timer1.Interval = 2000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyIcon1.Text = "Generando Documento Electrónicos";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseDoubleClick);
            // 
            // frmGeneraInformeLibroBoletas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(733, 290);
            this.Controls.Add(this.RadPageView1);
            this.Controls.Add(this.radStatusStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.ImageScalingSize = new System.Drawing.Size(5, 5);
            this.MaximizeBox = false;
            this.Name = "frmGeneraInformeLibroBoletas";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.SmallImageScalingSize = new System.Drawing.Size(5, 5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Genera Libro de Boletas Electrónicas";
            this.ThemeName = "TelerikMetroBlue";
            this.Load += new System.EventHandler(this.frmGeneraInformeLibroBoletas_Load);
            this.Resize += new System.EventHandler(this.frmGeneraDocumentos_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RadPageView1)).EndInit();
            this.RadPageView1.ResumeLayout(false);
            this.RadPageViewPage1.ResumeLayout(false);
            this.RadPageViewPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rbRectifica)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rbEspecial)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rbMensual)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ddLAno)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ddlMes)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ddLocales)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnGenerar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Telerik.WinControls.Themes.TelerikMetroBlueTheme telerikMetroBlueTheme1;
        private Telerik.WinControls.UI.RadStatusStrip radStatusStrip1;
        private Telerik.WinControls.UI.RadLabelElement lblInformacion;
        internal Telerik.WinControls.UI.RadPageView RadPageView1;
        internal Telerik.WinControls.UI.RadPageViewPage RadPageViewPage1;
        private System.Windows.Forms.Label lblRut;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblNombreEmpresa;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private Telerik.WinControls.UI.RadButton btnGenerar;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private Telerik.WinControls.UI.RadButton radButton1;
        private Telerik.WinControls.UI.RadDropDownList ddLocales;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private Telerik.WinControls.UI.RadDropDownList ddLAno;
        private System.Windows.Forms.Label label5;
        private Telerik.WinControls.UI.RadDropDownList ddlMes;
        private Telerik.WinControls.UI.RadRadioButton rbMensual;
        private System.Windows.Forms.Label label6;
        private Telerik.WinControls.UI.RadRadioButton rbRectifica;
        private Telerik.WinControls.UI.RadRadioButton rbEspecial;
    }
}
