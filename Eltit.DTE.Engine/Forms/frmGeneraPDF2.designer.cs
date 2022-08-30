namespace PlaceDTE
{
    partial class frmGeneraPDF2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGeneraPDF2));
            this.telerikMetroBlueTheme1 = new Telerik.WinControls.Themes.TelerikMetroBlueTheme();
            this.radStatusStrip1 = new Telerik.WinControls.UI.RadStatusStrip();
            this.lblInformacion = new Telerik.WinControls.UI.RadLabelElement();
            this.RadPageView1 = new Telerik.WinControls.UI.RadPageView();
            this.RadPageViewPage1 = new Telerik.WinControls.UI.RadPageViewPage();
            this.txtLocalActivo = new Telerik.WinControls.UI.RadTextBox();
            this.lblCiudad = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.lblComuna = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblDireccion = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblRut = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblNombreEmpresa = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.radWaitingBar1 = new Telerik.WinControls.UI.RadWaitingBar();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.radPageView2 = new Telerik.WinControls.UI.RadPageView();
            this.radPageViewPage2 = new Telerik.WinControls.UI.RadPageViewPage();
            this.pictureBoxTimbre = new System.Windows.Forms.PictureBox();
            this.label13 = new System.Windows.Forms.Label();
            this.btnGenerar = new Telerik.WinControls.UI.RadButton();
            this.radGroupBox1 = new Telerik.WinControls.UI.RadGroupBox();
            this.radCheckBox1 = new Telerik.WinControls.UI.RadCheckBox();
            this.lblMonto = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lbTipo = new System.Windows.Forms.Label();
            this.lblFolio = new System.Windows.Forms.Label();
            this.label25 = new System.Windows.Forms.Label();
            this.lblFecha = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.lblNombreDocumento = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.btnBuscar = new Telerik.WinControls.UI.RadButton();
            this.txtFilePath = new Telerik.WinControls.UI.RadTextBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.btnSalir = new Telerik.WinControls.UI.RadButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.radButton1 = new Telerik.WinControls.UI.RadButton();
            this.chkImprimeDirecto = new Telerik.WinControls.UI.RadCheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RadPageView1)).BeginInit();
            this.RadPageView1.SuspendLayout();
            this.RadPageViewPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtLocalActivo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radWaitingBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radPageView2)).BeginInit();
            this.radPageView2.SuspendLayout();
            this.radPageViewPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxTimbre)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnGenerar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGroupBox1)).BeginInit();
            this.radGroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radCheckBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnBuscar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtFilePath)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSalir)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkImprimeDirecto)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // radStatusStrip1
            // 
            this.radStatusStrip1.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.lblInformacion});
            this.radStatusStrip1.Location = new System.Drawing.Point(0, 363);
            this.radStatusStrip1.Name = "radStatusStrip1";
            this.radStatusStrip1.Size = new System.Drawing.Size(967, 25);
            this.radStatusStrip1.TabIndex = 2;
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
            this.RadPageView1.Location = new System.Drawing.Point(12, 13);
            this.RadPageView1.Name = "RadPageView1";
            this.RadPageView1.SelectedPage = this.RadPageViewPage1;
            this.RadPageView1.Size = new System.Drawing.Size(567, 164);
            this.RadPageView1.TabIndex = 8;
            this.RadPageView1.Text = "Documento";
            this.RadPageView1.ThemeName = "TelerikMetroBlue";
            ((Telerik.WinControls.UI.RadPageViewStripElement)(this.RadPageView1.GetChildAt(0))).StripButtons = Telerik.WinControls.UI.StripViewButtons.None;
            // 
            // RadPageViewPage1
            // 
            this.RadPageViewPage1.Controls.Add(this.txtLocalActivo);
            this.RadPageViewPage1.Controls.Add(this.lblCiudad);
            this.RadPageViewPage1.Controls.Add(this.label7);
            this.RadPageViewPage1.Controls.Add(this.lblComuna);
            this.RadPageViewPage1.Controls.Add(this.label5);
            this.RadPageViewPage1.Controls.Add(this.lblDireccion);
            this.RadPageViewPage1.Controls.Add(this.label2);
            this.RadPageViewPage1.Controls.Add(this.lblRut);
            this.RadPageViewPage1.Controls.Add(this.label3);
            this.RadPageViewPage1.Controls.Add(this.lblNombreEmpresa);
            this.RadPageViewPage1.Controls.Add(this.label1);
            this.RadPageViewPage1.Controls.Add(this.radWaitingBar1);
            this.RadPageViewPage1.Controls.Add(this.pictureBox1);
            this.RadPageViewPage1.ItemSize = new System.Drawing.SizeF(173F, 25F);
            this.RadPageViewPage1.Location = new System.Drawing.Point(5, 31);
            this.RadPageViewPage1.Name = "RadPageViewPage1";
            this.RadPageViewPage1.Size = new System.Drawing.Size(557, 128);
            this.RadPageViewPage1.Text = "Certificando a Cliente Cliente: ";
            // 
            // txtLocalActivo
            // 
            this.txtLocalActivo.Enabled = false;
            this.txtLocalActivo.Location = new System.Drawing.Point(504, 0);
            this.txtLocalActivo.Name = "txtLocalActivo";
            this.txtLocalActivo.Size = new System.Drawing.Size(45, 24);
            this.txtLocalActivo.TabIndex = 19;
            this.txtLocalActivo.ThemeName = "TelerikMetroBlue";
            // 
            // lblCiudad
            // 
            this.lblCiudad.AutoEllipsis = true;
            this.lblCiudad.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblCiudad.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCiudad.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCiudad.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblCiudad.Location = new System.Drawing.Point(433, 83);
            this.lblCiudad.Name = "lblCiudad";
            this.lblCiudad.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblCiudad.Size = new System.Drawing.Size(116, 24);
            this.lblCiudad.TabIndex = 10;
            this.lblCiudad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label7.Location = new System.Drawing.Point(367, 83);
            this.label7.Name = "label7";
            this.label7.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label7.Size = new System.Drawing.Size(60, 24);
            this.label7.TabIndex = 9;
            this.label7.Text = "Ciudad";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblComuna
            // 
            this.lblComuna.AutoEllipsis = true;
            this.lblComuna.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblComuna.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblComuna.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblComuna.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblComuna.Location = new System.Drawing.Point(218, 83);
            this.lblComuna.Name = "lblComuna";
            this.lblComuna.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblComuna.Size = new System.Drawing.Size(145, 24);
            this.lblComuna.TabIndex = 8;
            this.lblComuna.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label5.Location = new System.Drawing.Point(127, 83);
            this.label5.Name = "label5";
            this.label5.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label5.Size = new System.Drawing.Size(85, 24);
            this.label5.TabIndex = 7;
            this.label5.Text = "Comuna";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblDireccion
            // 
            this.lblDireccion.AutoEllipsis = true;
            this.lblDireccion.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblDireccion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDireccion.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDireccion.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblDireccion.Location = new System.Drawing.Point(218, 55);
            this.lblDireccion.Name = "lblDireccion";
            this.lblDireccion.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblDireccion.Size = new System.Drawing.Size(331, 24);
            this.lblDireccion.TabIndex = 6;
            this.lblDireccion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label2.Location = new System.Drawing.Point(127, 55);
            this.label2.Name = "label2";
            this.label2.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label2.Size = new System.Drawing.Size(85, 24);
            this.label2.TabIndex = 5;
            this.label2.Text = "Dirección";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblRut
            // 
            this.lblRut.AutoEllipsis = true;
            this.lblRut.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblRut.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblRut.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRut.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblRut.Location = new System.Drawing.Point(218, 0);
            this.lblRut.Name = "lblRut";
            this.lblRut.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblRut.Size = new System.Drawing.Size(130, 24);
            this.lblRut.TabIndex = 4;
            this.lblRut.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label3.Location = new System.Drawing.Point(127, 0);
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
            this.lblNombreEmpresa.Location = new System.Drawing.Point(218, 27);
            this.lblNombreEmpresa.Name = "lblNombreEmpresa";
            this.lblNombreEmpresa.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblNombreEmpresa.Size = new System.Drawing.Size(331, 24);
            this.lblNombreEmpresa.TabIndex = 2;
            this.lblNombreEmpresa.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label1.Location = new System.Drawing.Point(127, 27);
            this.label1.Name = "label1";
            this.label1.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label1.Size = new System.Drawing.Size(85, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "Razon Social";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // radWaitingBar1
            // 
            this.radWaitingBar1.Location = new System.Drawing.Point(127, 111);
            this.radWaitingBar1.Name = "radWaitingBar1";
            this.radWaitingBar1.Size = new System.Drawing.Size(422, 14);
            this.radWaitingBar1.TabIndex = 0;
            this.radWaitingBar1.Text = "radWaitingBar1";
            this.radWaitingBar1.ThemeName = "TelerikMetroBlue";
            this.radWaitingBar1.WaitingStyle = Telerik.WinControls.Enumerations.WaitingBarStyles.Throbber;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(3, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(118, 122);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // radPageView2
            // 
            this.radPageView2.Controls.Add(this.radPageViewPage2);
            this.radPageView2.Location = new System.Drawing.Point(590, 13);
            this.radPageView2.Name = "radPageView2";
            this.radPageView2.SelectedPage = this.radPageViewPage2;
            this.radPageView2.Size = new System.Drawing.Size(373, 184);
            this.radPageView2.TabIndex = 10;
            this.radPageView2.Text = "Documento";
            this.radPageView2.ThemeName = "TelerikMetroBlue";
            ((Telerik.WinControls.UI.RadPageViewStripElement)(this.radPageView2.GetChildAt(0))).StripButtons = Telerik.WinControls.UI.StripViewButtons.None;
            // 
            // radPageViewPage2
            // 
            this.radPageViewPage2.Controls.Add(this.pictureBoxTimbre);
            this.radPageViewPage2.ItemSize = new System.Drawing.SizeF(90F, 25F);
            this.radPageViewPage2.Location = new System.Drawing.Point(5, 31);
            this.radPageViewPage2.Name = "radPageViewPage2";
            this.radPageViewPage2.Size = new System.Drawing.Size(363, 148);
            this.radPageViewPage2.Text = "Listado de Pdf";
            // 
            // pictureBoxTimbre
            // 
            this.pictureBoxTimbre.Location = new System.Drawing.Point(25, 8);
            this.pictureBoxTimbre.Name = "pictureBoxTimbre";
            this.pictureBoxTimbre.Size = new System.Drawing.Size(318, 125);
            this.pictureBoxTimbre.TabIndex = 14;
            this.pictureBoxTimbre.TabStop = false;
            // 
            // label13
            // 
            this.label13.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label13.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label13.Location = new System.Drawing.Point(12, 31);
            this.label13.Name = "label13";
            this.label13.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label13.Size = new System.Drawing.Size(114, 24);
            this.label13.TabIndex = 3;
            this.label13.Text = "Ruta de Archivo";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnGenerar
            // 
            this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
            this.btnGenerar.Location = new System.Drawing.Point(590, 215);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.btnGenerar.Size = new System.Drawing.Size(182, 43);
            this.btnGenerar.TabIndex = 16;
            this.btnGenerar.Text = "Generar PDF";
            this.btnGenerar.ThemeName = "TelerikMetroBlue";
            this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
            // 
            // radGroupBox1
            // 
            this.radGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this.radGroupBox1.BackColor = System.Drawing.Color.White;
            this.radGroupBox1.Controls.Add(this.radCheckBox1);
            this.radGroupBox1.Controls.Add(this.lblMonto);
            this.radGroupBox1.Controls.Add(this.label4);
            this.radGroupBox1.Controls.Add(this.lbTipo);
            this.radGroupBox1.Controls.Add(this.lblFolio);
            this.radGroupBox1.Controls.Add(this.label25);
            this.radGroupBox1.Controls.Add(this.lblFecha);
            this.radGroupBox1.Controls.Add(this.label23);
            this.radGroupBox1.Controls.Add(this.lblNombreDocumento);
            this.radGroupBox1.Controls.Add(this.label21);
            this.radGroupBox1.Controls.Add(this.btnBuscar);
            this.radGroupBox1.Controls.Add(this.txtFilePath);
            this.radGroupBox1.Controls.Add(this.label13);
            this.radGroupBox1.HeaderText = "Detalle del Documento";
            this.radGroupBox1.Location = new System.Drawing.Point(12, 191);
            this.radGroupBox1.Name = "radGroupBox1";
            this.radGroupBox1.Size = new System.Drawing.Size(567, 163);
            this.radGroupBox1.TabIndex = 17;
            this.radGroupBox1.Text = "Detalle del Documento";
            this.radGroupBox1.ThemeName = "TelerikMetroBlue";
            this.radGroupBox1.Click += new System.EventHandler(this.radGroupBox1_Click);
            // 
            // radCheckBox1
            // 
            this.radCheckBox1.Location = new System.Drawing.Point(390, 127);
            this.radCheckBox1.Name = "radCheckBox1";
            this.radCheckBox1.Size = new System.Drawing.Size(125, 19);
            this.radCheckBox1.TabIndex = 31;
            this.radCheckBox1.Text = "PDF de Simulación";
            this.radCheckBox1.ThemeName = "TelerikMetroBlue";
            // 
            // lblMonto
            // 
            this.lblMonto.AutoEllipsis = true;
            this.lblMonto.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblMonto.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblMonto.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMonto.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblMonto.Location = new System.Drawing.Point(132, 122);
            this.lblMonto.Name = "lblMonto";
            this.lblMonto.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblMonto.Size = new System.Drawing.Size(105, 24);
            this.lblMonto.TabIndex = 30;
            this.lblMonto.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label4.Location = new System.Drawing.Point(12, 122);
            this.label4.Name = "label4";
            this.label4.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label4.Size = new System.Drawing.Size(114, 24);
            this.label4.TabIndex = 29;
            this.label4.Text = "Monto Total";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lbTipo
            // 
            this.lbTipo.AutoEllipsis = true;
            this.lbTipo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lbTipo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lbTipo.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbTipo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lbTipo.Location = new System.Drawing.Point(132, 62);
            this.lbTipo.Name = "lbTipo";
            this.lbTipo.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lbTipo.Size = new System.Drawing.Size(37, 24);
            this.lbTipo.TabIndex = 28;
            this.lbTipo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblFolio
            // 
            this.lblFolio.AutoEllipsis = true;
            this.lblFolio.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblFolio.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFolio.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFolio.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblFolio.Location = new System.Drawing.Point(132, 92);
            this.lblFolio.Name = "lblFolio";
            this.lblFolio.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblFolio.Size = new System.Drawing.Size(105, 24);
            this.lblFolio.TabIndex = 26;
            this.lblFolio.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label25
            // 
            this.label25.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label25.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label25.Location = new System.Drawing.Point(12, 92);
            this.label25.Name = "label25";
            this.label25.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label25.Size = new System.Drawing.Size(114, 24);
            this.label25.TabIndex = 24;
            this.label25.Text = "Folio";
            this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblFecha
            // 
            this.lblFecha.AutoEllipsis = true;
            this.lblFecha.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblFecha.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFecha.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFecha.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblFecha.Location = new System.Drawing.Point(390, 92);
            this.lblFecha.Name = "lblFecha";
            this.lblFecha.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblFecha.Size = new System.Drawing.Size(105, 24);
            this.lblFecha.TabIndex = 23;
            this.lblFecha.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label23
            // 
            this.label23.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label23.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label23.Location = new System.Drawing.Point(270, 92);
            this.label23.Name = "label23";
            this.label23.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label23.Size = new System.Drawing.Size(114, 24);
            this.label23.TabIndex = 22;
            this.label23.Text = "Fecha";
            this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblNombreDocumento
            // 
            this.lblNombreDocumento.AutoEllipsis = true;
            this.lblNombreDocumento.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblNombreDocumento.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNombreDocumento.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNombreDocumento.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblNombreDocumento.Location = new System.Drawing.Point(175, 62);
            this.lblNombreDocumento.Name = "lblNombreDocumento";
            this.lblNombreDocumento.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblNombreDocumento.Size = new System.Drawing.Size(320, 24);
            this.lblNombreDocumento.TabIndex = 21;
            this.lblNombreDocumento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label21
            // 
            this.label21.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label21.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label21.Location = new System.Drawing.Point(12, 62);
            this.label21.Name = "label21";
            this.label21.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label21.Size = new System.Drawing.Size(114, 24);
            this.label21.TabIndex = 20;
            this.label21.Text = "Tipo Documento";
            this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnBuscar
            // 
            this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
            this.btnBuscar.Location = new System.Drawing.Point(501, 21);
            this.btnBuscar.Name = "btnBuscar";
            this.btnBuscar.Padding = new System.Windows.Forms.Padding(10, 5, 0, 5);
            this.btnBuscar.Size = new System.Drawing.Size(45, 34);
            this.btnBuscar.TabIndex = 19;
            this.btnBuscar.ThemeName = "TelerikMetroBlue";
            this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Enabled = false;
            this.txtFilePath.Location = new System.Drawing.Point(132, 31);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(363, 24);
            this.txtFilePath.TabIndex = 18;
            this.txtFilePath.ThemeName = "TelerikMetroBlue";
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(801, 264);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(143, 90);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 19;
            this.pictureBox2.TabStop = false;
            // 
            // btnSalir
            // 
            this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
            this.btnSalir.Location = new System.Drawing.Point(590, 313);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.btnSalir.Size = new System.Drawing.Size(182, 41);
            this.btnSalir.TabIndex = 20;
            this.btnSalir.Text = "     Salir";
            this.btnSalir.ThemeName = "TelerikMetroBlue";
            this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // radButton1
            // 
            this.radButton1.Image = ((System.Drawing.Image)(resources.GetObject("radButton1.Image")));
            this.radButton1.Location = new System.Drawing.Point(590, 264);
            this.radButton1.Name = "radButton1";
            this.radButton1.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.radButton1.Size = new System.Drawing.Size(182, 43);
            this.radButton1.TabIndex = 21;
            this.radButton1.Text = "Limpiar";
            this.radButton1.ThemeName = "TelerikMetroBlue";
            this.radButton1.Click += new System.EventHandler(this.radButton1_Click);
            // 
            // chkImprimeDirecto
            // 
            this.chkImprimeDirecto.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImprimeDirecto.Location = new System.Drawing.Point(801, 212);
            this.chkImprimeDirecto.Name = "chkImprimeDirecto";
            this.chkImprimeDirecto.Size = new System.Drawing.Size(112, 19);
            this.chkImprimeDirecto.TabIndex = 32;
            this.chkImprimeDirecto.Text = "Imprime Directo";
            this.chkImprimeDirecto.ThemeName = "TelerikMetroBlue";
            this.chkImprimeDirecto.ToggleState = Telerik.WinControls.Enumerations.ToggleState.On;
            // 
            // frmGeneraPDF2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(967, 388);
            this.Controls.Add(this.chkImprimeDirecto);
            this.Controls.Add(this.radButton1);
            this.Controls.Add(this.btnSalir);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.radGroupBox1);
            this.Controls.Add(this.radPageView2);
            this.Controls.Add(this.RadPageView1);
            this.Controls.Add(this.radStatusStrip1);
            this.Controls.Add(this.btnGenerar);
            this.MaximizeBox = false;
            this.Name = "frmGeneraPDF2";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Formulario de  Generación de PDF";
            this.ThemeName = "TelerikMetroBlue";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.Load += new System.EventHandler(this.frmGeneraPDF2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RadPageView1)).EndInit();
            this.RadPageView1.ResumeLayout(false);
            this.RadPageViewPage1.ResumeLayout(false);
            this.RadPageViewPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtLocalActivo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radWaitingBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radPageView2)).EndInit();
            this.radPageView2.ResumeLayout(false);
            this.radPageViewPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxTimbre)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnGenerar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGroupBox1)).EndInit();
            this.radGroupBox1.ResumeLayout(false);
            this.radGroupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radCheckBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnBuscar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtFilePath)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSalir)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkImprimeDirecto)).EndInit();
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
        private System.Windows.Forms.Label lblCiudad;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblComuna;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblDireccion;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblRut;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblNombreEmpresa;
        private System.Windows.Forms.Label label1;
        private Telerik.WinControls.UI.RadWaitingBar radWaitingBar1;
        private System.Windows.Forms.PictureBox pictureBox1;
        internal Telerik.WinControls.UI.RadPageView radPageView2;
        internal Telerik.WinControls.UI.RadPageViewPage radPageViewPage2;
        private System.Windows.Forms.Label label13;
        private Telerik.WinControls.UI.RadGroupBox radGroupBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label lblNombreDocumento;
        private System.Windows.Forms.Label label21;
        private Telerik.WinControls.UI.RadButton btnBuscar;
        private Telerik.WinControls.UI.RadTextBox txtFilePath;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label lblFecha;
        private System.Windows.Forms.Label label23;
        private Telerik.WinControls.UI.RadButton btnSalir;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label lbTipo;
        private System.Windows.Forms.Label lblMonto;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblFolio;
        private System.Windows.Forms.PictureBox pictureBoxTimbre;
        private Telerik.WinControls.UI.RadButton radButton1;
        private Telerik.WinControls.UI.RadCheckBox radCheckBox1;
        public Telerik.WinControls.UI.RadButton btnGenerar;
        private Telerik.WinControls.UI.RadCheckBox chkImprimeDirecto;
        public Telerik.WinControls.UI.RadTextBox txtLocalActivo;
    }
}
