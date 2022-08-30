namespace SamplesDTE
{
    partial class frmPrincipal
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPrincipal));
            this.telerikMetroBlueTheme1 = new Telerik.WinControls.Themes.TelerikMetroBlueTheme();
            this.radStatusStrip1 = new Telerik.WinControls.UI.RadStatusStrip();
            this.lblInformacion = new Telerik.WinControls.UI.RadLabelElement();
            this.lblRecinto = new Telerik.WinControls.UI.RadLabelElement();
            this.lblusuario = new Telerik.WinControls.UI.RadLabelElement();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.radMenu1 = new Telerik.WinControls.UI.RadMenu();
            this.radMenuItem1 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem2 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem13 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem3 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem6 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem7 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem10 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem12 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem14 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem15 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem4 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem9 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem11 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem16 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem5 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem8 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem17 = new Telerik.WinControls.UI.RadMenuItem();
            this.radMenuItem18 = new Telerik.WinControls.UI.RadMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radMenu1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // radStatusStrip1
            // 
            this.radStatusStrip1.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.lblInformacion,
            this.lblRecinto,
            this.lblusuario});
            this.radStatusStrip1.Location = new System.Drawing.Point(0, 668);
            this.radStatusStrip1.Name = "radStatusStrip1";
            this.radStatusStrip1.Size = new System.Drawing.Size(1305, 25);
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
            // lblRecinto
            // 
            this.lblRecinto.Name = "lblRecinto";
            this.radStatusStrip1.SetSpring(this.lblRecinto, false);
            this.lblRecinto.Text = ".";
            this.lblRecinto.TextWrap = true;
            // 
            // lblusuario
            // 
            this.lblusuario.Name = "lblusuario";
            this.radStatusStrip1.SetSpring(this.lblusuario, false);
            this.lblusuario.Text = "";
            this.lblusuario.TextWrap = true;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyIcon1.Text = "Generando Documento Electrónicos";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseDoubleClick);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(12, 567);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(337, 112);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // radMenu1
            // 
            this.radMenu1.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.radMenuItem1,
            this.radMenuItem3,
            this.radMenuItem4,
            this.radMenuItem5,
            this.radMenuItem17});
            this.radMenu1.Location = new System.Drawing.Point(0, 0);
            this.radMenu1.Name = "radMenu1";
            this.radMenu1.Padding = new System.Windows.Forms.Padding(2, 2, 0, 0);
            this.radMenu1.Size = new System.Drawing.Size(1305, 29);
            this.radMenu1.TabIndex = 2;
            this.radMenu1.Text = "radMenu1";
            this.radMenu1.ThemeName = "TelerikMetroBlue";
            this.radMenu1.Click += new System.EventHandler(this.radMenu1_Click);
            // 
            // radMenuItem1
            // 
            this.radMenuItem1.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.radMenuItem2,
            this.radMenuItem13});
            this.radMenuItem1.Name = "radMenuItem1";
            this.radMenuItem1.Text = "Mantención";
            // 
            // radMenuItem2
            // 
            this.radMenuItem2.Enabled = false;
            this.radMenuItem2.Name = "radMenuItem2";
            this.radMenuItem2.Text = "Ingreso de Timbraje Boletas Electrónicas";
            this.radMenuItem2.Click += new System.EventHandler(this.radMenuItem2_Click);
            // 
            // radMenuItem13
            // 
            this.radMenuItem13.Enabled = false;
            this.radMenuItem13.Name = "radMenuItem13";
            this.radMenuItem13.Text = "Ingreso Timbraje Facturas Notas y Guias";
            this.radMenuItem13.Click += new System.EventHandler(this.radMenuItem13_Click);
            // 
            // radMenuItem3
            // 
            this.radMenuItem3.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.radMenuItem6,
            this.radMenuItem7,
            this.radMenuItem10,
            this.radMenuItem12,
            this.radMenuItem14,
            this.radMenuItem15});
            this.radMenuItem3.Name = "radMenuItem3";
            this.radMenuItem3.Text = "Operaciones";
            // 
            // radMenuItem6
            // 
            this.radMenuItem6.Name = "radMenuItem6";
            this.radMenuItem6.Text = "Actualizar Información de Intercambio";
            this.radMenuItem6.Click += new System.EventHandler(this.radMenuItem6_Click);
            // 
            // radMenuItem7
            // 
            this.radMenuItem7.Enabled = false;
            this.radMenuItem7.Name = "radMenuItem7";
            this.radMenuItem7.SerializeChildren = false;
            this.radMenuItem7.Text = "Genera Reporte de Consumo de Folios";
            this.radMenuItem7.Click += new System.EventHandler(this.radMenuItem7_Click);
            // 
            // radMenuItem10
            // 
            this.radMenuItem10.Enabled = false;
            this.radMenuItem10.Name = "radMenuItem10";
            this.radMenuItem10.Text = "Genera Libro de Boletas Electrónicas";
            this.radMenuItem10.Click += new System.EventHandler(this.radMenuItem10_Click);
            // 
            // radMenuItem12
            // 
            this.radMenuItem12.Enabled = false;
            this.radMenuItem12.Name = "radMenuItem12";
            this.radMenuItem12.Text = "Procesa Correos de Intercambio";
            // 
            // radMenuItem14
            // 
            this.radMenuItem14.Enabled = false;
            this.radMenuItem14.Name = "radMenuItem14";
            this.radMenuItem14.Text = "Regenarar XML Masivo Boletas";
            this.radMenuItem14.Click += new System.EventHandler(this.radMenuItem14_Click);
            // 
            // radMenuItem15
            // 
            this.radMenuItem15.Enabled = false;
            this.radMenuItem15.Name = "radMenuItem15";
            this.radMenuItem15.Text = "Regenerar Facturas, Notas de Crédito y Guías";
            this.radMenuItem15.Click += new System.EventHandler(this.radMenuItem15_Click);
            // 
            // radMenuItem4
            // 
            this.radMenuItem4.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.radMenuItem9,
            this.radMenuItem11,
            this.radMenuItem16});
            this.radMenuItem4.Name = "radMenuItem4";
            this.radMenuItem4.Text = "Informes";
            // 
            // radMenuItem9
            // 
            this.radMenuItem9.Enabled = false;
            this.radMenuItem9.Name = "radMenuItem9";
            this.radMenuItem9.Text = "Informe de DTEs Recibidos";
            // 
            // radMenuItem11
            // 
            this.radMenuItem11.Enabled = false;
            this.radMenuItem11.Name = "radMenuItem11";
            this.radMenuItem11.Text = "Revisa Folios Disponibles";
            // 
            // radMenuItem16
            // 
            this.radMenuItem16.Enabled = false;
            this.radMenuItem16.Name = "radMenuItem16";
            this.radMenuItem16.Text = "Informe de Boletas Generadas";
            this.radMenuItem16.Click += new System.EventHandler(this.radMenuItem16_Click);
            // 
            // radMenuItem5
            // 
            this.radMenuItem5.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.radMenuItem8});
            this.radMenuItem5.Name = "radMenuItem5";
            this.radMenuItem5.Text = "Acerca de ...";
            // 
            // radMenuItem8
            // 
            this.radMenuItem8.Enabled = false;
            this.radMenuItem8.Name = "radMenuItem8";
            this.radMenuItem8.Text = "Acerca de DTE Manager";
            // 
            // radMenuItem17
            // 
            this.radMenuItem17.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.radMenuItem18});
            this.radMenuItem17.Name = "radMenuItem17";
            this.radMenuItem17.Text = "Utilidades";
            // 
            // radMenuItem18
            // 
            this.radMenuItem18.Name = "radMenuItem18";
            this.radMenuItem18.Text = "Imprimir documentos";
            this.radMenuItem18.Click += new System.EventHandler(this.radMenuItem18_Click);
            // 
            // frmPrincipal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(1305, 693);
            this.Controls.Add(this.radMenu1);
            this.Controls.Add(this.radStatusStrip1);
            this.Controls.Add(this.pictureBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.ImageScalingSize = new System.Drawing.Size(5, 5);
            this.MaximizeBox = false;
            this.Name = "frmPrincipal";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.SmallImageScalingSize = new System.Drawing.Size(5, 5);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Módulo de Control de Documentación Electrónica";
            this.ThemeName = "TelerikMetroBlue";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmPrincipal_FormClosed);
            this.Load += new System.EventHandler(this.frmGeneraDocumentos_Load);
            this.Resize += new System.EventHandler(this.frmGeneraDocumentos_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radMenu1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Telerik.WinControls.Themes.TelerikMetroBlueTheme telerikMetroBlueTheme1;
        private Telerik.WinControls.UI.RadStatusStrip radStatusStrip1;
        private Telerik.WinControls.UI.RadLabelElement lblInformacion;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private Telerik.WinControls.UI.RadMenu radMenu1;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem1;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem2;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem3;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem4;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem5;
        private Telerik.WinControls.UI.RadLabelElement lblRecinto;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem6;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem7;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem9;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem8;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem10;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem11;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem12;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem13;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem14;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem15;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem16;
        private Telerik.WinControls.UI.RadLabelElement lblusuario;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem17;
        private Telerik.WinControls.UI.RadMenuItem radMenuItem18;
    }
}
