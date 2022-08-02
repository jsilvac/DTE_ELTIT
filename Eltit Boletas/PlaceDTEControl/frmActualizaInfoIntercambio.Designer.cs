namespace SamplesDTE
{
    partial class frmActualizaInfoIntercambio
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmActualizaInfoIntercambio));
            this.telerikMetroBlueTheme1 = new Telerik.WinControls.Themes.TelerikMetroBlueTheme();
            this.radStatusStrip1 = new Telerik.WinControls.UI.RadStatusStrip();
            this.lblInformacion = new Telerik.WinControls.UI.RadLabelElement();
            this.RadPageView1 = new Telerik.WinControls.UI.RadPageView();
            this.RadPageViewPage1 = new Telerik.WinControls.UI.RadPageViewPage();
            this.lblCiudad = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.label7 = new System.Windows.Forms.Label();
            this.lblComuna = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblDireccion = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblNombreEmpresa = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.radWaitingBar1 = new Telerik.WinControls.UI.RadWaitingBar();
            this.label13 = new System.Windows.Forms.Label();
            this.radGroupBox1 = new Telerik.WinControls.UI.RadGroupBox();
            this.lblInfo = new System.Windows.Forms.Label();
            this.btnBuscar = new Telerik.WinControls.UI.RadButton();
            this.txtFilePath = new Telerik.WinControls.UI.RadTextBox();
            this.btnSalir = new Telerik.WinControls.UI.RadButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.radButton1 = new Telerik.WinControls.UI.RadButton();
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RadPageView1)).BeginInit();
            this.RadPageView1.SuspendLayout();
            this.RadPageViewPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radWaitingBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGroupBox1)).BeginInit();
            this.radGroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnBuscar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtFilePath)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSalir)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // radStatusStrip1
            // 
            this.radStatusStrip1.Items.AddRange(new Telerik.WinControls.RadItem[] {
            this.lblInformacion});
            this.radStatusStrip1.Location = new System.Drawing.Point(0, 400);
            this.radStatusStrip1.Name = "radStatusStrip1";
            this.radStatusStrip1.Size = new System.Drawing.Size(590, 25);
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
            this.RadPageViewPage1.Controls.Add(this.lblCiudad);
            this.RadPageViewPage1.Controls.Add(this.pictureBox2);
            this.RadPageViewPage1.Controls.Add(this.label7);
            this.RadPageViewPage1.Controls.Add(this.lblComuna);
            this.RadPageViewPage1.Controls.Add(this.label5);
            this.RadPageViewPage1.Controls.Add(this.lblDireccion);
            this.RadPageViewPage1.Controls.Add(this.label2);
            this.RadPageViewPage1.Controls.Add(this.lblNombreEmpresa);
            this.RadPageViewPage1.Controls.Add(this.label1);
            this.RadPageViewPage1.Controls.Add(this.radWaitingBar1);
            this.RadPageViewPage1.ItemSize = new System.Drawing.SizeF(127F, 25F);
            this.RadPageViewPage1.Location = new System.Drawing.Point(5, 31);
            this.RadPageViewPage1.Name = "RadPageViewPage1";
            this.RadPageViewPage1.Size = new System.Drawing.Size(557, 128);
            this.RadPageViewPage1.Text = "Datos de La empresa";
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
            this.lblCiudad.Text = "PUCON";
            this.lblCiudad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(0, 0);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(121, 101);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 19;
            this.pictureBox2.TabStop = false;
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
            this.lblComuna.Text = "PUCON";
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
            this.lblDireccion.Text = "PUCON";
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
            this.lblNombreEmpresa.Text = "EMPRESAS ELTIT";
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
            // radGroupBox1
            // 
            this.radGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this.radGroupBox1.BackColor = System.Drawing.Color.White;
            this.radGroupBox1.Controls.Add(this.lblInfo);
            this.radGroupBox1.Controls.Add(this.btnBuscar);
            this.radGroupBox1.Controls.Add(this.txtFilePath);
            this.radGroupBox1.Controls.Add(this.label13);
            this.radGroupBox1.HeaderText = "Detalle del Documento";
            this.radGroupBox1.Location = new System.Drawing.Point(12, 191);
            this.radGroupBox1.Name = "radGroupBox1";
            this.radGroupBox1.Size = new System.Drawing.Size(567, 125);
            this.radGroupBox1.TabIndex = 17;
            this.radGroupBox1.Text = "Detalle del Documento";
            this.radGroupBox1.ThemeName = "TelerikMetroBlue";
            this.radGroupBox1.Click += new System.EventHandler(this.radGroupBox1_Click);
            // 
            // lblInfo
            // 
            this.lblInfo.AutoEllipsis = true;
            this.lblInfo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.lblInfo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblInfo.Font = new System.Drawing.Font("Franklin Gothic Medium", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblInfo.Location = new System.Drawing.Point(10, 74);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblInfo.Size = new System.Drawing.Size(544, 24);
            this.lblInfo.TabIndex = 20;
            this.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnBuscar
            // 
            this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
            this.btnBuscar.Location = new System.Drawing.Point(509, 21);
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
            // btnSalir
            // 
            this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
            this.btnSalir.Location = new System.Drawing.Point(395, 353);
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
            this.radButton1.Location = new System.Drawing.Point(198, 351);
            this.radButton1.Name = "radButton1";
            this.radButton1.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.radButton1.Size = new System.Drawing.Size(182, 43);
            this.radButton1.TabIndex = 21;
            this.radButton1.Text = "Limpiar";
            this.radButton1.ThemeName = "TelerikMetroBlue";
            this.radButton1.Click += new System.EventHandler(this.radButton1_Click);
            // 
            // frmActualizaInfoIntercambio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(590, 425);
            this.Controls.Add(this.radButton1);
            this.Controls.Add(this.btnSalir);
            this.Controls.Add(this.radGroupBox1);
            this.Controls.Add(this.RadPageView1);
            this.Controls.Add(this.radStatusStrip1);
            this.MaximizeBox = false;
            this.Name = "frmActualizaInfoIntercambio";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Módulo de Ingreso de Contribuyentes";
            this.ThemeName = "TelerikMetroBlue";
            this.Load += new System.EventHandler(this.frmActualizaInfoIntercambio_Load);
            ((System.ComponentModel.ISupportInitialize)(this.radStatusStrip1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RadPageView1)).EndInit();
            this.RadPageView1.ResumeLayout(false);
            this.RadPageViewPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radWaitingBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGroupBox1)).EndInit();
            this.radGroupBox1.ResumeLayout(false);
            this.radGroupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnBuscar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtFilePath)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSalir)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).EndInit();
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
        private System.Windows.Forms.Label lblNombreEmpresa;
        private System.Windows.Forms.Label label1;
        private Telerik.WinControls.UI.RadWaitingBar radWaitingBar1;
        private System.Windows.Forms.Label label13;
        private Telerik.WinControls.UI.RadGroupBox radGroupBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private Telerik.WinControls.UI.RadButton btnBuscar;
        private Telerik.WinControls.UI.RadTextBox txtFilePath;
        private Telerik.WinControls.UI.RadButton btnSalir;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private Telerik.WinControls.UI.RadButton radButton1;
        private System.Windows.Forms.Label lblInfo;
    }
}
