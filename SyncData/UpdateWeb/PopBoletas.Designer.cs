namespace SchoolManagementAdmin
{
    partial class PopBoletas
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
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn21 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn22 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn23 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn24 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn25 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.TableViewDefinition tableViewDefinition5 = new Telerik.WinControls.UI.TableViewDefinition();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn26 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn27 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn28 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn29 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn30 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.TableViewDefinition tableViewDefinition6 = new Telerik.WinControls.UI.TableViewDefinition();
            this.gvPagos = new Telerik.WinControls.UI.RadGridView();
            this.lblNombreLocal = new System.Windows.Forms.Label();
            this.gvGrilla2 = new Telerik.WinControls.UI.RadGridView();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.gvPagos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPagos.MasterTemplate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvGrilla2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvGrilla2.MasterTemplate)).BeginInit();
            this.SuspendLayout();
            // 
            // gvPagos
            // 
            this.gvPagos.EnableTheming = false;
            this.gvPagos.EnterKeyMode = Telerik.WinControls.UI.RadGridViewEnterKeyMode.EnterMovesToNextRow;
            this.gvPagos.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gvPagos.Location = new System.Drawing.Point(12, 53);
            // 
            // 
            // 
            this.gvPagos.MasterTemplate.AllowAddNewRow = false;
            this.gvPagos.MasterTemplate.AllowColumnReorder = false;
            this.gvPagos.MasterTemplate.AllowDeleteRow = false;
            gridViewTextBoxColumn21.HeaderText = "Cod.";
            gridViewTextBoxColumn21.MaxLength = 5;
            gridViewTextBoxColumn21.MinWidth = 1;
            gridViewTextBoxColumn21.Name = "codigo";
            gridViewTextBoxColumn21.ReadOnly = true;
            gridViewTextBoxColumn21.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            gridViewTextBoxColumn21.Width = 40;
            gridViewTextBoxColumn22.HeaderText = "Nombre Caja";
            gridViewTextBoxColumn22.Name = "local";
            gridViewTextBoxColumn22.ReadOnly = true;
            gridViewTextBoxColumn22.Width = 150;
            gridViewTextBoxColumn23.HeaderText = "Server";
            gridViewTextBoxColumn23.MaxLength = 15;
            gridViewTextBoxColumn23.Name = "server";
            gridViewTextBoxColumn23.TextAlignment = System.Drawing.ContentAlignment.MiddleRight;
            gridViewTextBoxColumn23.Width = 150;
            gridViewTextBoxColumn24.HeaderText = "Destino";
            gridViewTextBoxColumn24.MinWidth = 1;
            gridViewTextBoxColumn24.Name = "destino";
            gridViewTextBoxColumn24.Width = 1;
            gridViewTextBoxColumn25.HeaderText = "Caf 39";
            gridViewTextBoxColumn25.Name = "caf39";
            gridViewTextBoxColumn25.TextAlignment = System.Drawing.ContentAlignment.MiddleRight;
            gridViewTextBoxColumn25.Width = 80;
            this.gvPagos.MasterTemplate.Columns.AddRange(new Telerik.WinControls.UI.GridViewDataColumn[] {
            gridViewTextBoxColumn21,
            gridViewTextBoxColumn22,
            gridViewTextBoxColumn23,
            gridViewTextBoxColumn24,
            gridViewTextBoxColumn25});
            this.gvPagos.MasterTemplate.EnableAlternatingRowColor = true;
            this.gvPagos.MasterTemplate.EnableGrouping = false;
            this.gvPagos.MasterTemplate.EnableSorting = false;
            this.gvPagos.MasterTemplate.MultiSelect = true;
            this.gvPagos.MasterTemplate.ViewDefinition = tableViewDefinition5;
            this.gvPagos.Name = "gvPagos";
            this.gvPagos.PrintStyle.SummaryCellBackColor = System.Drawing.Color.Green;
            this.gvPagos.ShowGroupPanel = false;
            this.gvPagos.Size = new System.Drawing.Size(457, 637);
            this.gvPagos.TabIndex = 24;
            this.gvPagos.ThemeName = "TelerikMetroBlue";
            // 
            // lblNombreLocal
            // 
            this.lblNombreLocal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNombreLocal.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNombreLocal.ForeColor = System.Drawing.Color.Gray;
            this.lblNombreLocal.Location = new System.Drawing.Point(12, 14);
            this.lblNombreLocal.Name = "lblNombreLocal";
            this.lblNombreLocal.Size = new System.Drawing.Size(938, 33);
            this.lblNombreLocal.TabIndex = 25;
            this.lblNombreLocal.Text = ".";
            this.lblNombreLocal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // gvGrilla2
            // 
            this.gvGrilla2.EnableTheming = false;
            this.gvGrilla2.EnterKeyMode = Telerik.WinControls.UI.RadGridViewEnterKeyMode.EnterMovesToNextRow;
            this.gvGrilla2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gvGrilla2.Location = new System.Drawing.Point(491, 53);
            // 
            // 
            // 
            this.gvGrilla2.MasterTemplate.AllowAddNewRow = false;
            this.gvGrilla2.MasterTemplate.AllowColumnReorder = false;
            this.gvGrilla2.MasterTemplate.AllowDeleteRow = false;
            gridViewTextBoxColumn26.HeaderText = "Cod.";
            gridViewTextBoxColumn26.MaxLength = 5;
            gridViewTextBoxColumn26.MinWidth = 1;
            gridViewTextBoxColumn26.Name = "codigo";
            gridViewTextBoxColumn26.ReadOnly = true;
            gridViewTextBoxColumn26.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            gridViewTextBoxColumn26.Width = 40;
            gridViewTextBoxColumn27.HeaderText = "Nombre Caja";
            gridViewTextBoxColumn27.Name = "local";
            gridViewTextBoxColumn27.ReadOnly = true;
            gridViewTextBoxColumn27.Width = 150;
            gridViewTextBoxColumn28.HeaderText = "Server";
            gridViewTextBoxColumn28.MaxLength = 15;
            gridViewTextBoxColumn28.Name = "server";
            gridViewTextBoxColumn28.TextAlignment = System.Drawing.ContentAlignment.MiddleRight;
            gridViewTextBoxColumn28.Width = 150;
            gridViewTextBoxColumn29.HeaderText = "Destino";
            gridViewTextBoxColumn29.MinWidth = 1;
            gridViewTextBoxColumn29.Name = "destino";
            gridViewTextBoxColumn29.Width = 1;
            gridViewTextBoxColumn30.HeaderText = "Caf 39";
            gridViewTextBoxColumn30.Name = "caf39";
            gridViewTextBoxColumn30.TextAlignment = System.Drawing.ContentAlignment.MiddleRight;
            gridViewTextBoxColumn30.Width = 80;
            this.gvGrilla2.MasterTemplate.Columns.AddRange(new Telerik.WinControls.UI.GridViewDataColumn[] {
            gridViewTextBoxColumn26,
            gridViewTextBoxColumn27,
            gridViewTextBoxColumn28,
            gridViewTextBoxColumn29,
            gridViewTextBoxColumn30});
            this.gvGrilla2.MasterTemplate.EnableAlternatingRowColor = true;
            this.gvGrilla2.MasterTemplate.EnableGrouping = false;
            this.gvGrilla2.MasterTemplate.EnableSorting = false;
            this.gvGrilla2.MasterTemplate.MultiSelect = true;
            this.gvGrilla2.MasterTemplate.ViewDefinition = tableViewDefinition6;
            this.gvGrilla2.Name = "gvGrilla2";
            this.gvGrilla2.PrintStyle.SummaryCellBackColor = System.Drawing.Color.Green;
            this.gvGrilla2.ShowGroupPanel = false;
            this.gvGrilla2.Size = new System.Drawing.Size(459, 637);
            this.gvGrilla2.TabIndex = 26;
            this.gvGrilla2.ThemeName = "TelerikMetroBlue";
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // PopBoletas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(962, 702);
            this.Controls.Add(this.gvGrilla2);
            this.Controls.Add(this.lblNombreLocal);
            this.Controls.Add(this.gvPagos);
            this.Name = "PopBoletas";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Informe de Boletas Por Local";
            this.Activated += new System.EventHandler(this.PopBoletas_Activated);
            this.Load += new System.EventHandler(this.PopBoletas_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gvPagos.MasterTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPagos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvGrilla2.MasterTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvGrilla2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        internal Telerik.WinControls.UI.RadGridView gvPagos;
        public System.Windows.Forms.Label lblNombreLocal;
        internal Telerik.WinControls.UI.RadGridView gvGrilla2;
        private System.Windows.Forms.Timer timer1;
    }
}