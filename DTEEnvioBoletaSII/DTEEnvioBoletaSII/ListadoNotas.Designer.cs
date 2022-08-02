namespace SchoolManagementAdmin
{
    partial class frmListado
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle31 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle32 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle33 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle34 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle35 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle36 = new System.Windows.Forms.DataGridViewCellStyle();
            this.RightOptions = new System.Windows.Forms.Timer(this.components);
            this.pnlRightMain = new System.Windows.Forms.Panel();
            this.pbExit = new System.Windows.Forms.PictureBox();
            this.pbHome = new System.Windows.Forms.PictureBox();
            this.pbLogout = new System.Windows.Forms.PictureBox();
            this.pnlRightOptions = new System.Windows.Forms.Panel();
            this.Options = new System.Windows.Forms.Timer(this.components);
            this.pnlMain = new System.Windows.Forms.Panel();
            this.btnSalir = new MetroFramework.Controls.MetroButton();
            this.metroTabControl1 = new MetroFramework.Controls.MetroTabControl();
            this.metroTabPage1 = new MetroFramework.Controls.MetroTabPage();
            this.txtdia = new System.Windows.Forms.TextBox();
            this.listaNotas = new MetroFramework.Controls.MetroGrid();
            this.panel16 = new System.Windows.Forms.Panel();
            this.lblcaja = new System.Windows.Forms.Label();
            this.txtaño = new System.Windows.Forms.TextBox();
            this.txtmes = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.metroTabPage2 = new MetroFramework.Controls.MetroTabPage();
            this.panel19 = new System.Windows.Forms.Panel();
            this.lblcuentaitems = new System.Windows.Forms.Label();
            this.gdrollo = new MetroFramework.Controls.MetroGrid();
            this.panel11 = new System.Windows.Forms.Panel();
            this.lbliva = new System.Windows.Forms.Label();
            this.panel10 = new System.Windows.Forms.Panel();
            this.lbltotal = new System.Windows.Forms.Label();
            this.panel9 = new System.Windows.Forms.Panel();
            this.lblneto = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.lblmensaje = new System.Windows.Forms.Label();
            this.btnInsertar = new MetroFramework.Controls.MetroButton();
            this.pnlRightMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbExit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbHome)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbLogout)).BeginInit();
            this.pnlRightOptions.SuspendLayout();
            this.pnlMain.SuspendLayout();
            this.metroTabControl1.SuspendLayout();
            this.metroTabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listaNotas)).BeginInit();
            this.panel16.SuspendLayout();
            this.metroTabPage2.SuspendLayout();
            this.panel19.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gdrollo)).BeginInit();
            this.panel11.SuspendLayout();
            this.panel10.SuspendLayout();
            this.panel9.SuspendLayout();
            this.SuspendLayout();
            // 
            // RightOptions
            // 
            this.RightOptions.Interval = 1;
            this.RightOptions.Tick += new System.EventHandler(this.RightOptions_Tick);
            // 
            // pnlRightMain
            // 
            this.pnlRightMain.Controls.Add(this.pbExit);
            this.pnlRightMain.Controls.Add(this.pbHome);
            this.pnlRightMain.Controls.Add(this.pbLogout);
            this.pnlRightMain.Location = new System.Drawing.Point(3, 52);
            this.pnlRightMain.Name = "pnlRightMain";
            this.pnlRightMain.Size = new System.Drawing.Size(70, 396);
            this.pnlRightMain.TabIndex = 1;
            // 
            // pbExit
            // 
            this.pbExit.Image = global::SchoolManagementAdmin.Properties.Resources.appbar_power;
            this.pbExit.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.pbExit.Location = new System.Drawing.Point(-3, 320);
            this.pbExit.Name = "pbExit";
            this.pbExit.Size = new System.Drawing.Size(76, 76);
            this.pbExit.TabIndex = 5;
            this.pbExit.TabStop = false;
            this.pbExit.Click += new System.EventHandler(this.pbExit_Click);
            // 
            // pbHome
            // 
            this.pbHome.Image = global::SchoolManagementAdmin.Properties.Resources.appbar_home;
            this.pbHome.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.pbHome.Location = new System.Drawing.Point(-3, 160);
            this.pbHome.Name = "pbHome";
            this.pbHome.Size = new System.Drawing.Size(76, 76);
            this.pbHome.TabIndex = 4;
            this.pbHome.TabStop = false;
            this.pbHome.Click += new System.EventHandler(this.pbHome_Click);
            // 
            // pbLogout
            // 
            this.pbLogout.Image = global::SchoolManagementAdmin.Properties.Resources.appbar_lock;
            this.pbLogout.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.pbLogout.Location = new System.Drawing.Point(-3, 0);
            this.pbLogout.Name = "pbLogout";
            this.pbLogout.Size = new System.Drawing.Size(76, 76);
            this.pbLogout.TabIndex = 3;
            this.pbLogout.TabStop = false;
            // 
            // pnlRightOptions
            // 
            this.pnlRightOptions.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(20)))), ((int)(((byte)(20)))));
            this.pnlRightOptions.Controls.Add(this.pnlRightMain);
            this.pnlRightOptions.Location = new System.Drawing.Point(1025, 1);
            this.pnlRightOptions.Name = "pnlRightOptions";
            this.pnlRightOptions.Size = new System.Drawing.Size(77, 644);
            this.pnlRightOptions.TabIndex = 15;
            // 
            // pnlMain
            // 
            this.pnlMain.BackColor = System.Drawing.Color.Transparent;
            this.pnlMain.Controls.Add(this.btnSalir);
            this.pnlMain.Controls.Add(this.metroTabControl1);
            this.pnlMain.Location = new System.Drawing.Point(12, 12);
            this.pnlMain.Name = "pnlMain";
            this.pnlMain.Size = new System.Drawing.Size(1007, 710);
            this.pnlMain.TabIndex = 13;
            this.pnlMain.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlMain_Paint);
            // 
            // btnSalir
            // 
            this.btnSalir.BackColor = System.Drawing.Color.LightSkyBlue;
            this.btnSalir.ForeColor = System.Drawing.SystemColors.Control;
            this.btnSalir.Location = new System.Drawing.Point(499, 648);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Size = new System.Drawing.Size(196, 30);
            this.btnSalir.TabIndex = 23;
            this.btnSalir.Text = "SALIR";
            this.btnSalir.Theme = MetroFramework.MetroThemeStyle.Light;
            this.btnSalir.UseCustomBackColor = true;
            this.btnSalir.UseCustomForeColor = true;
            this.btnSalir.UseSelectable = true;
            this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
            // 
            // metroTabControl1
            // 
            this.metroTabControl1.Controls.Add(this.metroTabPage1);
            this.metroTabControl1.Controls.Add(this.metroTabPage2);
            this.metroTabControl1.FontWeight = MetroFramework.MetroTabControlWeight.Bold;
            this.metroTabControl1.Location = new System.Drawing.Point(77, 16);
            this.metroTabControl1.Margin = new System.Windows.Forms.Padding(3, 3, 10, 3);
            this.metroTabControl1.Name = "metroTabControl1";
            this.metroTabControl1.SelectedIndex = 0;
            this.metroTabControl1.Size = new System.Drawing.Size(663, 626);
            this.metroTabControl1.Style = MetroFramework.MetroColorStyle.Blue;
            this.metroTabControl1.TabIndex = 22;
            this.metroTabControl1.Theme = MetroFramework.MetroThemeStyle.Light;
            this.metroTabControl1.UseCustomForeColor = true;
            this.metroTabControl1.UseSelectable = true;
            // 
            // metroTabPage1
            // 
            this.metroTabPage1.Controls.Add(this.lblmensaje);
            this.metroTabPage1.Controls.Add(this.txtdia);
            this.metroTabPage1.Controls.Add(this.listaNotas);
            this.metroTabPage1.Controls.Add(this.panel16);
            this.metroTabPage1.Controls.Add(this.txtaño);
            this.metroTabPage1.Controls.Add(this.txtmes);
            this.metroTabPage1.Controls.Add(this.button3);
            this.metroTabPage1.HorizontalScrollbarBarColor = true;
            this.metroTabPage1.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.HorizontalScrollbarSize = 0;
            this.metroTabPage1.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage1.Name = "metroTabPage1";
            this.metroTabPage1.Size = new System.Drawing.Size(655, 584);
            this.metroTabPage1.TabIndex = 0;
            this.metroTabPage1.Text = "Notas       ";
            this.metroTabPage1.VerticalScrollbarBarColor = true;
            this.metroTabPage1.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.VerticalScrollbarSize = 0;
            // 
            // txtdia
            // 
            this.txtdia.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtdia.Font = new System.Drawing.Font("Segoe UI", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtdia.ForeColor = System.Drawing.Color.Gray;
            this.txtdia.Location = new System.Drawing.Point(199, 4);
            this.txtdia.MaxLength = 2;
            this.txtdia.Name = "txtdia";
            this.txtdia.Size = new System.Drawing.Size(53, 50);
            this.txtdia.TabIndex = 29;
            this.txtdia.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // listaNotas
            // 
            this.listaNotas.AllowUserToAddRows = false;
            this.listaNotas.AllowUserToDeleteRows = false;
            this.listaNotas.AllowUserToResizeRows = false;
            this.listaNotas.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.listaNotas.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listaNotas.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.listaNotas.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle31.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle31.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle31.Font = new System.Drawing.Font("Arial Rounded MT Bold", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle31.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle31.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle31.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle31.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.listaNotas.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle31;
            this.listaNotas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle32.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle32.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle32.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle32.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(136)))), ((int)(((byte)(136)))), ((int)(((byte)(136)))));
            dataGridViewCellStyle32.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle32.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle32.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.listaNotas.DefaultCellStyle = dataGridViewCellStyle32;
            this.listaNotas.EnableHeadersVisualStyles = false;
            this.listaNotas.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.listaNotas.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.listaNotas.Location = new System.Drawing.Point(3, 60);
            this.listaNotas.MultiSelect = false;
            this.listaNotas.Name = "listaNotas";
            this.listaNotas.ReadOnly = true;
            this.listaNotas.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle33.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle33.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle33.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle33.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle33.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle33.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle33.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.listaNotas.RowHeadersDefaultCellStyle = dataGridViewCellStyle33;
            this.listaNotas.RowHeadersVisible = false;
            this.listaNotas.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.listaNotas.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.listaNotas.Size = new System.Drawing.Size(621, 484);
            this.listaNotas.TabIndex = 28;
            this.listaNotas.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.listaNotas_CellContentClick);
            this.listaNotas.DoubleClick += new System.EventHandler(this.listaNotas_DoubleClick);
            // 
            // panel16
            // 
            this.panel16.BackColor = System.Drawing.Color.LightSkyBlue;
            this.panel16.Controls.Add(this.lblcaja);
            this.panel16.Location = new System.Drawing.Point(3, 3);
            this.panel16.Name = "panel16";
            this.panel16.Size = new System.Drawing.Size(190, 50);
            this.panel16.TabIndex = 27;
            // 
            // lblcaja
            // 
            this.lblcaja.BackColor = System.Drawing.Color.Transparent;
            this.lblcaja.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblcaja.ForeColor = System.Drawing.Color.White;
            this.lblcaja.Location = new System.Drawing.Point(8, 5);
            this.lblcaja.Name = "lblcaja";
            this.lblcaja.Size = new System.Drawing.Size(122, 38);
            this.lblcaja.TabIndex = 3;
            this.lblcaja.Text = "PERIODO";
            this.lblcaja.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtaño
            // 
            this.txtaño.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtaño.Font = new System.Drawing.Font("Segoe UI", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtaño.ForeColor = System.Drawing.Color.Gray;
            this.txtaño.Location = new System.Drawing.Point(317, 4);
            this.txtaño.MaxLength = 4;
            this.txtaño.Name = "txtaño";
            this.txtaño.Size = new System.Drawing.Size(95, 50);
            this.txtaño.TabIndex = 25;
            this.txtaño.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtmes
            // 
            this.txtmes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtmes.Font = new System.Drawing.Font("Segoe UI", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtmes.ForeColor = System.Drawing.Color.Gray;
            this.txtmes.Location = new System.Drawing.Point(258, 4);
            this.txtmes.MaxLength = 2;
            this.txtmes.Name = "txtmes";
            this.txtmes.Size = new System.Drawing.Size(53, 50);
            this.txtmes.TabIndex = 24;
            this.txtmes.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.RoyalBlue;
            this.button3.FlatAppearance.BorderColor = System.Drawing.Color.Gainsboro;
            this.button3.FlatAppearance.CheckedBackColor = System.Drawing.Color.Gainsboro;
            this.button3.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Gainsboro;
            this.button3.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gainsboro;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.Color.White;
            this.button3.Location = new System.Drawing.Point(418, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(206, 50);
            this.button3.TabIndex = 26;
            this.button3.Text = "BUSCAR";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // metroTabPage2
            // 
            this.metroTabPage2.Controls.Add(this.btnInsertar);
            this.metroTabPage2.Controls.Add(this.panel19);
            this.metroTabPage2.Controls.Add(this.gdrollo);
            this.metroTabPage2.Controls.Add(this.panel11);
            this.metroTabPage2.Controls.Add(this.panel10);
            this.metroTabPage2.Controls.Add(this.panel9);
            this.metroTabPage2.Controls.Add(this.button1);
            this.metroTabPage2.HorizontalScrollbarBarColor = true;
            this.metroTabPage2.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.HorizontalScrollbarSize = 0;
            this.metroTabPage2.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage2.Name = "metroTabPage2";
            this.metroTabPage2.Size = new System.Drawing.Size(655, 584);
            this.metroTabPage2.Style = MetroFramework.MetroColorStyle.Silver;
            this.metroTabPage2.TabIndex = 1;
            this.metroTabPage2.Text = " Detalle Producto   ";
            this.metroTabPage2.VerticalScrollbarBarColor = true;
            this.metroTabPage2.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.VerticalScrollbarSize = 0;
            // 
            // panel19
            // 
            this.panel19.BackColor = System.Drawing.Color.CornflowerBlue;
            this.panel19.Controls.Add(this.lblcuentaitems);
            this.panel19.Location = new System.Drawing.Point(1, -5);
            this.panel19.Name = "panel19";
            this.panel19.Size = new System.Drawing.Size(131, 65);
            this.panel19.TabIndex = 29;
            // 
            // lblcuentaitems
            // 
            this.lblcuentaitems.BackColor = System.Drawing.Color.Transparent;
            this.lblcuentaitems.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblcuentaitems.Location = new System.Drawing.Point(0, 11);
            this.lblcuentaitems.Name = "lblcuentaitems";
            this.lblcuentaitems.Size = new System.Drawing.Size(128, 38);
            this.lblcuentaitems.TabIndex = 3;
            this.lblcuentaitems.Text = "ITEMS 32";
            this.lblcuentaitems.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // gdrollo
            // 
            this.gdrollo.AllowUserToAddRows = false;
            this.gdrollo.AllowUserToDeleteRows = false;
            this.gdrollo.AllowUserToResizeRows = false;
            this.gdrollo.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.gdrollo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gdrollo.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.gdrollo.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle34.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle34.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle34.Font = new System.Drawing.Font("Arial Rounded MT Bold", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle34.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle34.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle34.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle34.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gdrollo.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle34;
            this.gdrollo.ColumnHeadersHeight = 5;
            this.gdrollo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridViewCellStyle35.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle35.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle35.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle35.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(136)))), ((int)(((byte)(136)))), ((int)(((byte)(136)))));
            dataGridViewCellStyle35.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle35.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle35.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.gdrollo.DefaultCellStyle = dataGridViewCellStyle35;
            this.gdrollo.EnableHeadersVisualStyles = false;
            this.gdrollo.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.gdrollo.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.gdrollo.Location = new System.Drawing.Point(4, 66);
            this.gdrollo.MultiSelect = false;
            this.gdrollo.Name = "gdrollo";
            this.gdrollo.ReadOnly = true;
            this.gdrollo.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle36.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle36.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle36.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle36.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle36.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(174)))), ((int)(((byte)(219)))));
            dataGridViewCellStyle36.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle36.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gdrollo.RowHeadersDefaultCellStyle = dataGridViewCellStyle36;
            this.gdrollo.RowHeadersVisible = false;
            this.gdrollo.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.gdrollo.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gdrollo.Size = new System.Drawing.Size(649, 467);
            this.gdrollo.TabIndex = 28;
            // 
            // panel11
            // 
            this.panel11.BackColor = System.Drawing.Color.YellowGreen;
            this.panel11.Controls.Add(this.lbliva);
            this.panel11.Location = new System.Drawing.Point(279, -7);
            this.panel11.Name = "panel11";
            this.panel11.Size = new System.Drawing.Size(119, 65);
            this.panel11.TabIndex = 27;
            // 
            // lbliva
            // 
            this.lbliva.BackColor = System.Drawing.Color.Transparent;
            this.lbliva.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbliva.Location = new System.Drawing.Point(3, 7);
            this.lbliva.Name = "lbliva";
            this.lbliva.Size = new System.Drawing.Size(116, 51);
            this.lbliva.TabIndex = 4;
            this.lbliva.Text = "0";
            this.lbliva.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel10
            // 
            this.panel10.BackColor = System.Drawing.Color.YellowGreen;
            this.panel10.Controls.Add(this.lbltotal);
            this.panel10.Location = new System.Drawing.Point(404, -7);
            this.panel10.Name = "panel10";
            this.panel10.Size = new System.Drawing.Size(249, 65);
            this.panel10.TabIndex = 26;
            // 
            // lbltotal
            // 
            this.lbltotal.BackColor = System.Drawing.Color.Transparent;
            this.lbltotal.Font = new System.Drawing.Font("Segoe UI", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbltotal.Location = new System.Drawing.Point(3, 7);
            this.lbltotal.Name = "lbltotal";
            this.lbltotal.Size = new System.Drawing.Size(236, 51);
            this.lbltotal.TabIndex = 2;
            this.lbltotal.Text = "0";
            this.lbltotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel9
            // 
            this.panel9.BackColor = System.Drawing.Color.YellowGreen;
            this.panel9.Controls.Add(this.lblneto);
            this.panel9.Location = new System.Drawing.Point(135, -7);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(141, 65);
            this.panel9.TabIndex = 25;
            // 
            // lblneto
            // 
            this.lblneto.BackColor = System.Drawing.Color.Transparent;
            this.lblneto.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblneto.Location = new System.Drawing.Point(3, 7);
            this.lblneto.Name = "lblneto";
            this.lblneto.Size = new System.Drawing.Size(134, 51);
            this.lblneto.TabIndex = 4;
            this.lblneto.Text = "0";
            this.lblneto.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.RoyalBlue;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.Gainsboro;
            this.button1.FlatAppearance.CheckedBackColor = System.Drawing.Color.Gainsboro;
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Gainsboro;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gainsboro;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(469, 477);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(141, 39);
            this.button1.TabIndex = 24;
            this.button1.Text = "SALIR";
            this.button1.UseVisualStyleBackColor = false;
            // 
            // lblmensaje
            // 
            this.lblmensaje.AutoSize = true;
            this.lblmensaje.Location = new System.Drawing.Point(56, 553);
            this.lblmensaje.Name = "lblmensaje";
            this.lblmensaje.Size = new System.Drawing.Size(18, 30);
            this.lblmensaje.TabIndex = 30;
            this.lblmensaje.Text = ".";
            // 
            // btnInsertar
            // 
            this.btnInsertar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btnInsertar.Enabled = false;
            this.btnInsertar.ForeColor = System.Drawing.SystemColors.Control;
            this.btnInsertar.Location = new System.Drawing.Point(414, 552);
            this.btnInsertar.Name = "btnInsertar";
            this.btnInsertar.Size = new System.Drawing.Size(195, 30);
            this.btnInsertar.TabIndex = 30;
            this.btnInsertar.Text = "INSERTAR EN ROLLO";
            this.btnInsertar.Theme = MetroFramework.MetroThemeStyle.Light;
            this.btnInsertar.UseCustomBackColor = true;
            this.btnInsertar.UseCustomForeColor = true;
            this.btnInsertar.UseSelectable = true;
            this.btnInsertar.Click += new System.EventHandler(this.btnInsertar_Click);
            // 
            // frmListado
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 30F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1003, 740);
            this.Controls.Add(this.pnlRightOptions);
            this.Controls.Add(this.pnlMain);
            this.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(6, 7, 6, 7);
            this.Name = "frmListado";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.frmListado_Load);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.frmTemplate_MouseMove);
            this.pnlRightMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbExit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbHome)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbLogout)).EndInit();
            this.pnlRightOptions.ResumeLayout(false);
            this.pnlMain.ResumeLayout(false);
            this.metroTabControl1.ResumeLayout(false);
            this.metroTabPage1.ResumeLayout(false);
            this.metroTabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listaNotas)).EndInit();
            this.panel16.ResumeLayout(false);
            this.metroTabPage2.ResumeLayout(false);
            this.panel19.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gdrollo)).EndInit();
            this.panel11.ResumeLayout(false);
            this.panel10.ResumeLayout(false);
            this.panel9.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer RightOptions;
        private System.Windows.Forms.PictureBox pbExit;
        private System.Windows.Forms.PictureBox pbHome;
        private System.Windows.Forms.Panel pnlRightMain;
        private System.Windows.Forms.PictureBox pbLogout;
        private System.Windows.Forms.Panel pnlRightOptions;
        private System.Windows.Forms.Timer Options;
        private System.Windows.Forms.Panel pnlMain;
        private MetroFramework.Controls.MetroTabControl metroTabControl1;
        private MetroFramework.Controls.MetroTabPage metroTabPage1;
        internal MetroFramework.Controls.MetroGrid listaNotas;
        private System.Windows.Forms.Panel panel16;
        private System.Windows.Forms.Label lblcaja;
        private System.Windows.Forms.TextBox txtaño;
        private System.Windows.Forms.TextBox txtmes;
        private System.Windows.Forms.Button button3;
        private MetroFramework.Controls.MetroTabPage metroTabPage2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel19;
        private System.Windows.Forms.Label lblcuentaitems;
        internal MetroFramework.Controls.MetroGrid gdrollo;
        private System.Windows.Forms.Panel panel11;
        private System.Windows.Forms.Label lbliva;
        private System.Windows.Forms.Panel panel10;
        private System.Windows.Forms.Label lbltotal;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.Label lblneto;
        private MetroFramework.Controls.MetroButton btnSalir;
        private System.Windows.Forms.TextBox txtdia;
        private System.Windows.Forms.Label lblmensaje;
        private MetroFramework.Controls.MetroButton btnInsertar;
    }
}

