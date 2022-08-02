namespace SchoolManagementAdmin
{
    partial class PopLeeCorreos
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
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn11 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn12 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn13 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn14 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn15 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.TableViewDefinition tableViewDefinition3 = new Telerik.WinControls.UI.TableViewDefinition();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PopLeeCorreos));
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn16 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn17 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn18 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn19 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn20 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.TableViewDefinition tableViewDefinition4 = new Telerik.WinControls.UI.TableViewDefinition();
            this.gvPagos = new Telerik.WinControls.UI.RadGridView();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.RadPageView1 = new Telerik.WinControls.UI.RadPageView();
            this.RadPageViewPage1 = new Telerik.WinControls.UI.RadPageViewPage();
            this.lblPath = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.lblpassemail = new System.Windows.Forms.Label();
            this.lblHost = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.lblemail = new System.Windows.Forms.Label();
            this.lblNroResolucion = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.lblFechaResolucion = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lblCertificadoNombre = new System.Windows.Forms.Label();
            this.lblCertificadoRut = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblCodigo = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblComuna = new System.Windows.Forms.Label();
            this.lblDireccion = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblRut = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblNombreEmpresa = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGenerar = new Telerik.WinControls.UI.RadButton();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.gvEmpresas = new Telerik.WinControls.UI.RadGridView();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.txtInfo = new System.Windows.Forms.RichTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.gvPagos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPagos.MasterTemplate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RadPageView1)).BeginInit();
            this.RadPageView1.SuspendLayout();
            this.RadPageViewPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnGenerar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvEmpresas)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvEmpresas.MasterTemplate)).BeginInit();
            this.SuspendLayout();
            // 
            // gvPagos
            // 
            this.gvPagos.EnableTheming = false;
            this.gvPagos.EnterKeyMode = Telerik.WinControls.UI.RadGridViewEnterKeyMode.EnterMovesToNextRow;
            this.gvPagos.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gvPagos.Location = new System.Drawing.Point(512, 574);
            // 
            // 
            // 
            this.gvPagos.MasterTemplate.AllowAddNewRow = false;
            this.gvPagos.MasterTemplate.AllowColumnReorder = false;
            this.gvPagos.MasterTemplate.AllowDeleteRow = false;
            gridViewTextBoxColumn11.HeaderText = "ID";
            gridViewTextBoxColumn11.MaxLength = 5;
            gridViewTextBoxColumn11.MinWidth = 1;
            gridViewTextBoxColumn11.Name = "codigo";
            gridViewTextBoxColumn11.ReadOnly = true;
            gridViewTextBoxColumn11.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            gridViewTextBoxColumn11.Width = 40;
            gridViewTextBoxColumn12.HeaderText = "Correo Origen";
            gridViewTextBoxColumn12.Name = "origen";
            gridViewTextBoxColumn12.ReadOnly = true;
            gridViewTextBoxColumn12.Width = 200;
            gridViewTextBoxColumn13.HeaderText = "Server";
            gridViewTextBoxColumn13.MaxLength = 15;
            gridViewTextBoxColumn13.Name = "server";
            gridViewTextBoxColumn13.TextAlignment = System.Drawing.ContentAlignment.MiddleRight;
            gridViewTextBoxColumn13.Width = 120;
            gridViewTextBoxColumn14.HeaderText = "Subject";
            gridViewTextBoxColumn14.MinWidth = 1;
            gridViewTextBoxColumn14.Name = "destino";
            gridViewTextBoxColumn14.Width = 200;
            gridViewTextBoxColumn15.HeaderText = "Mensaje";
            gridViewTextBoxColumn15.Name = "mensaje";
            gridViewTextBoxColumn15.TextAlignment = System.Drawing.ContentAlignment.MiddleRight;
            gridViewTextBoxColumn15.Width = 5;
            this.gvPagos.MasterTemplate.Columns.AddRange(new Telerik.WinControls.UI.GridViewDataColumn[] {
            gridViewTextBoxColumn11,
            gridViewTextBoxColumn12,
            gridViewTextBoxColumn13,
            gridViewTextBoxColumn14,
            gridViewTextBoxColumn15});
            this.gvPagos.MasterTemplate.EnableAlternatingRowColor = true;
            this.gvPagos.MasterTemplate.EnableGrouping = false;
            this.gvPagos.MasterTemplate.EnableSorting = false;
            this.gvPagos.MasterTemplate.MultiSelect = true;
            this.gvPagos.MasterTemplate.ViewDefinition = tableViewDefinition3;
            this.gvPagos.Name = "gvPagos";
            this.gvPagos.PrintStyle.SummaryCellBackColor = System.Drawing.Color.Green;
            this.gvPagos.ShowGroupPanel = false;
            this.gvPagos.Size = new System.Drawing.Size(598, 42);
            this.gvPagos.TabIndex = 24;
            this.gvPagos.ThemeName = "TelerikMetroBlue";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // RadPageView1
            // 
            this.RadPageView1.BackColor = System.Drawing.Color.DarkGray;
            this.RadPageView1.Controls.Add(this.RadPageViewPage1);
            this.RadPageView1.Location = new System.Drawing.Point(512, 12);
            this.RadPageView1.Name = "RadPageView1";
            this.RadPageView1.SelectedPage = this.RadPageViewPage1;
            this.RadPageView1.Size = new System.Drawing.Size(598, 207);
            this.RadPageView1.TabIndex = 26;
            this.RadPageView1.Text = "Datos de Emisor Electrónico";
            this.RadPageView1.ThemeName = "TelerikMetroBlue";
            ((Telerik.WinControls.UI.RadPageViewStripElement)(this.RadPageView1.GetChildAt(0))).StripButtons = Telerik.WinControls.UI.StripViewButtons.None;
            // 
            // RadPageViewPage1
            // 
            this.RadPageViewPage1.Controls.Add(this.lblPath);
            this.RadPageViewPage1.Controls.Add(this.label10);
            this.RadPageViewPage1.Controls.Add(this.lblpassemail);
            this.RadPageViewPage1.Controls.Add(this.lblHost);
            this.RadPageViewPage1.Controls.Add(this.label4);
            this.RadPageViewPage1.Controls.Add(this.label7);
            this.RadPageViewPage1.Controls.Add(this.lblemail);
            this.RadPageViewPage1.Controls.Add(this.lblNroResolucion);
            this.RadPageViewPage1.Controls.Add(this.label11);
            this.RadPageViewPage1.Controls.Add(this.lblFechaResolucion);
            this.RadPageViewPage1.Controls.Add(this.label9);
            this.RadPageViewPage1.Controls.Add(this.lblCertificadoNombre);
            this.RadPageViewPage1.Controls.Add(this.lblCertificadoRut);
            this.RadPageViewPage1.Controls.Add(this.label5);
            this.RadPageViewPage1.Controls.Add(this.lblCodigo);
            this.RadPageViewPage1.Controls.Add(this.label6);
            this.RadPageViewPage1.Controls.Add(this.lblComuna);
            this.RadPageViewPage1.Controls.Add(this.lblDireccion);
            this.RadPageViewPage1.Controls.Add(this.label2);
            this.RadPageViewPage1.Controls.Add(this.lblRut);
            this.RadPageViewPage1.Controls.Add(this.label3);
            this.RadPageViewPage1.Controls.Add(this.lblNombreEmpresa);
            this.RadPageViewPage1.Controls.Add(this.label1);
            this.RadPageViewPage1.Controls.Add(this.btnGenerar);
            this.RadPageViewPage1.Controls.Add(this.pictureBox1);
            this.RadPageViewPage1.ItemSize = new System.Drawing.SizeF(156F, 28F);
            this.RadPageViewPage1.Location = new System.Drawing.Point(10, 37);
            this.RadPageViewPage1.Name = "RadPageViewPage1";
            this.RadPageViewPage1.Size = new System.Drawing.Size(577, 159);
            this.RadPageViewPage1.Text = "Datos de Emisor Electrónico";
            // 
            // lblPath
            // 
            this.lblPath.AutoEllipsis = true;
            this.lblPath.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblPath.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPath.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblPath.Location = new System.Drawing.Point(527, 1);
            this.lblPath.Name = "lblPath";
            this.lblPath.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblPath.Size = new System.Drawing.Size(55, 24);
            this.lblPath.TabIndex = 34;
            this.lblPath.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label10
            // 
            this.label10.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label10.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label10.Location = new System.Drawing.Point(457, -1);
            this.label10.Name = "label10";
            this.label10.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label10.Size = new System.Drawing.Size(68, 24);
            this.label10.TabIndex = 33;
            this.label10.Text = "Carpeta";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblpassemail
            // 
            this.lblpassemail.AutoEllipsis = true;
            this.lblpassemail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblpassemail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblpassemail.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblpassemail.ForeColor = System.Drawing.Color.White;
            this.lblpassemail.Location = new System.Drawing.Point(3, 141);
            this.lblpassemail.Name = "lblpassemail";
            this.lblpassemail.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblpassemail.Size = new System.Drawing.Size(118, 24);
            this.lblpassemail.TabIndex = 32;
            this.lblpassemail.Text = "...";
            this.lblpassemail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblHost
            // 
            this.lblHost.AutoEllipsis = true;
            this.lblHost.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblHost.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblHost.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHost.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblHost.Location = new System.Drawing.Point(218, 141);
            this.lblHost.Name = "lblHost";
            this.lblHost.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblHost.Size = new System.Drawing.Size(103, 24);
            this.lblHost.TabIndex = 31;
            this.lblHost.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label4.Location = new System.Drawing.Point(127, 141);
            this.label4.Name = "label4";
            this.label4.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label4.Size = new System.Drawing.Size(85, 24);
            this.label4.TabIndex = 30;
            this.label4.Text = "Host";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label7.Location = new System.Drawing.Point(327, 141);
            this.label7.Name = "label7";
            this.label7.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label7.Size = new System.Drawing.Size(85, 24);
            this.label7.TabIndex = 29;
            this.label7.Text = "Email";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblemail
            // 
            this.lblemail.AutoEllipsis = true;
            this.lblemail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblemail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblemail.Font = new System.Drawing.Font("Franklin Gothic Medium", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblemail.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblemail.Location = new System.Drawing.Point(418, 141);
            this.lblemail.Name = "lblemail";
            this.lblemail.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblemail.Size = new System.Drawing.Size(162, 24);
            this.lblemail.TabIndex = 28;
            this.lblemail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblNroResolucion
            // 
            this.lblNroResolucion.AutoEllipsis = true;
            this.lblNroResolucion.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblNroResolucion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNroResolucion.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNroResolucion.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblNroResolucion.Location = new System.Drawing.Point(418, 112);
            this.lblNroResolucion.Name = "lblNroResolucion";
            this.lblNroResolucion.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblNroResolucion.Size = new System.Drawing.Size(100, 24);
            this.lblNroResolucion.TabIndex = 27;
            this.lblNroResolucion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label11
            // 
            this.label11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label11.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label11.Location = new System.Drawing.Point(327, 111);
            this.label11.Name = "label11";
            this.label11.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label11.Size = new System.Drawing.Size(85, 24);
            this.label11.TabIndex = 26;
            this.label11.Text = "N° Resol.";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblFechaResolucion
            // 
            this.lblFechaResolucion.AutoEllipsis = true;
            this.lblFechaResolucion.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblFechaResolucion.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFechaResolucion.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFechaResolucion.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblFechaResolucion.Location = new System.Drawing.Point(218, 112);
            this.lblFechaResolucion.Name = "lblFechaResolucion";
            this.lblFechaResolucion.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblFechaResolucion.Size = new System.Drawing.Size(103, 24);
            this.lblFechaResolucion.TabIndex = 25;
            this.lblFechaResolucion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label9.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label9.Location = new System.Drawing.Point(127, 111);
            this.label9.Name = "label9";
            this.label9.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label9.Size = new System.Drawing.Size(85, 24);
            this.label9.TabIndex = 24;
            this.label9.Text = "FechaResol.";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblCertificadoNombre
            // 
            this.lblCertificadoNombre.AutoEllipsis = true;
            this.lblCertificadoNombre.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblCertificadoNombre.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCertificadoNombre.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCertificadoNombre.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblCertificadoNombre.Location = new System.Drawing.Point(327, 83);
            this.lblCertificadoNombre.Name = "lblCertificadoNombre";
            this.lblCertificadoNombre.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblCertificadoNombre.Size = new System.Drawing.Size(255, 24);
            this.lblCertificadoNombre.TabIndex = 23;
            this.lblCertificadoNombre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblCertificadoRut
            // 
            this.lblCertificadoRut.AutoEllipsis = true;
            this.lblCertificadoRut.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblCertificadoRut.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCertificadoRut.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCertificadoRut.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblCertificadoRut.Location = new System.Drawing.Point(218, 84);
            this.lblCertificadoRut.Name = "lblCertificadoRut";
            this.lblCertificadoRut.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblCertificadoRut.Size = new System.Drawing.Size(103, 24);
            this.lblCertificadoRut.TabIndex = 22;
            this.lblCertificadoRut.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label5.Location = new System.Drawing.Point(127, 83);
            this.label5.Name = "label5";
            this.label5.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label5.Size = new System.Drawing.Size(85, 24);
            this.label5.TabIndex = 21;
            this.label5.Text = "Certificado";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblCodigo
            // 
            this.lblCodigo.AutoEllipsis = true;
            this.lblCodigo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblCodigo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCodigo.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCodigo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblCodigo.Location = new System.Drawing.Point(408, 0);
            this.lblCodigo.Name = "lblCodigo";
            this.lblCodigo.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblCodigo.Size = new System.Drawing.Size(40, 24);
            this.lblCodigo.TabIndex = 20;
            this.lblCodigo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.label6.Location = new System.Drawing.Point(328, -1);
            this.label6.Name = "label6";
            this.label6.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.label6.Size = new System.Drawing.Size(74, 24);
            this.label6.TabIndex = 19;
            this.label6.Text = "Cód.Conta";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblComuna
            // 
            this.lblComuna.AutoEllipsis = true;
            this.lblComuna.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblComuna.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblComuna.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblComuna.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblComuna.Location = new System.Drawing.Point(497, 55);
            this.lblComuna.Name = "lblComuna";
            this.lblComuna.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblComuna.Size = new System.Drawing.Size(85, 24);
            this.lblComuna.TabIndex = 8;
            this.lblComuna.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            this.lblDireccion.Size = new System.Drawing.Size(273, 24);
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
            this.lblRut.Size = new System.Drawing.Size(103, 24);
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
            this.lblNombreEmpresa.Size = new System.Drawing.Size(364, 24);
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
            // btnGenerar
            // 
            this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
            this.btnGenerar.Location = new System.Drawing.Point(541, 110);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.btnGenerar.Size = new System.Drawing.Size(39, 27);
            this.btnGenerar.TabIndex = 16;
            this.btnGenerar.ThemeName = "TelerikMetroBlue";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(1, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(123, 122);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // gvEmpresas
            // 
            this.gvEmpresas.EnableTheming = false;
            this.gvEmpresas.EnterKeyMode = Telerik.WinControls.UI.RadGridViewEnterKeyMode.EnterMovesToNextRow;
            this.gvEmpresas.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gvEmpresas.Location = new System.Drawing.Point(11, 15);
            // 
            // 
            // 
            this.gvEmpresas.MasterTemplate.AllowAddNewRow = false;
            this.gvEmpresas.MasterTemplate.AllowColumnReorder = false;
            this.gvEmpresas.MasterTemplate.AllowDeleteRow = false;
            gridViewTextBoxColumn16.HeaderText = "Cod.";
            gridViewTextBoxColumn16.MaxLength = 200;
            gridViewTextBoxColumn16.MinWidth = 50;
            gridViewTextBoxColumn16.Name = "codigo";
            gridViewTextBoxColumn16.ReadOnly = true;
            gridViewTextBoxColumn16.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            gridViewTextBoxColumn17.HeaderText = "Nombre Empresa";
            gridViewTextBoxColumn17.Name = "nombre";
            gridViewTextBoxColumn17.ReadOnly = true;
            gridViewTextBoxColumn17.Width = 165;
            gridViewTextBoxColumn18.HeaderText = "Server";
            gridViewTextBoxColumn18.MaxLength = 15;
            gridViewTextBoxColumn18.Name = "server";
            gridViewTextBoxColumn18.TextAlignment = System.Drawing.ContentAlignment.MiddleRight;
            gridViewTextBoxColumn18.Width = 80;
            gridViewTextBoxColumn19.HeaderText = "Path";
            gridViewTextBoxColumn19.Name = "path";
            gridViewTextBoxColumn19.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            gridViewTextBoxColumn19.Width = 60;
            gridViewTextBoxColumn20.HeaderText = "Rut";
            gridViewTextBoxColumn20.Name = "rut";
            gridViewTextBoxColumn20.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            gridViewTextBoxColumn20.Width = 100;
            this.gvEmpresas.MasterTemplate.Columns.AddRange(new Telerik.WinControls.UI.GridViewDataColumn[] {
            gridViewTextBoxColumn16,
            gridViewTextBoxColumn17,
            gridViewTextBoxColumn18,
            gridViewTextBoxColumn19,
            gridViewTextBoxColumn20});
            this.gvEmpresas.MasterTemplate.EnableAlternatingRowColor = true;
            this.gvEmpresas.MasterTemplate.EnableGrouping = false;
            this.gvEmpresas.MasterTemplate.EnableSorting = false;
            this.gvEmpresas.MasterTemplate.MultiSelect = true;
            this.gvEmpresas.MasterTemplate.ViewDefinition = tableViewDefinition4;
            this.gvEmpresas.Name = "gvEmpresas";
            this.gvEmpresas.PrintStyle.SummaryCellBackColor = System.Drawing.Color.Green;
            this.gvEmpresas.ShowGroupPanel = false;
            this.gvEmpresas.Size = new System.Drawing.Size(489, 601);
            this.gvEmpresas.TabIndex = 28;
            this.gvEmpresas.ThemeName = "TelerikMetroBlue";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 650);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1122, 22);
            this.statusStrip1.TabIndex = 29;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // txtInfo
            // 
            this.txtInfo.BackColor = System.Drawing.SystemColors.InfoText;
            this.txtInfo.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtInfo.ForeColor = System.Drawing.Color.Lime;
            this.txtInfo.Location = new System.Drawing.Point(512, 226);
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.Size = new System.Drawing.Size(598, 342);
            this.txtInfo.TabIndex = 30;
            this.txtInfo.Text = "";
            // 
            // PopLeeCorreos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightBlue;
            this.ClientSize = new System.Drawing.Size(1122, 672);
            this.Controls.Add(this.txtInfo);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.gvEmpresas);
            this.Controls.Add(this.RadPageView1);
            this.Controls.Add(this.gvPagos);
            this.Name = "PopLeeCorreos";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Procesando Correos de Intercambio Electrónico";
            this.Activated += new System.EventHandler(this.PopBoletas_Activated);
            this.Load += new System.EventHandler(this.PopLeeCorreos_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gvPagos.MasterTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPagos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RadPageView1)).EndInit();
            this.RadPageView1.ResumeLayout(false);
            this.RadPageViewPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btnGenerar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvEmpresas.MasterTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvEmpresas)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal Telerik.WinControls.UI.RadGridView gvPagos;
        private System.Windows.Forms.Timer timer1;
        internal Telerik.WinControls.UI.RadPageView RadPageView1;
        internal Telerik.WinControls.UI.RadPageViewPage RadPageViewPage1;
        public System.Windows.Forms.Label lblNroResolucion;
        private System.Windows.Forms.Label label11;
        public System.Windows.Forms.Label lblFechaResolucion;
        private System.Windows.Forms.Label label9;
        public System.Windows.Forms.Label lblCertificadoNombre;
        public System.Windows.Forms.Label lblCertificadoRut;
        private System.Windows.Forms.Label label5;
        public System.Windows.Forms.Label lblCodigo;
        private System.Windows.Forms.Label label6;
        public System.Windows.Forms.Label lblComuna;
        private Telerik.WinControls.UI.RadButton btnGenerar;
        public System.Windows.Forms.Label lblDireccion;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label lblRut;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label lblNombreEmpresa;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        internal Telerik.WinControls.UI.RadGridView gvEmpresas;
        private System.Windows.Forms.Label label7;
        public System.Windows.Forms.Label lblemail;
        private System.Windows.Forms.StatusStrip statusStrip1;
        public System.Windows.Forms.Label lblHost;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.Label lblpassemail;
        private System.Windows.Forms.RichTextBox txtInfo;
        public System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.Label label10;
    }
}