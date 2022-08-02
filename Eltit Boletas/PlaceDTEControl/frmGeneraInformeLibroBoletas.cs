using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using SamplesDTE.Clases;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Globalization;
using Eltit.Clases;


namespace Eltit
{
    public partial class frmGeneraInformeLibroBoletas : Telerik.WinControls.UI.RadForm
    {
        private Icon[] icons = new Icon[2];
        private int currentIcon = 0;

        public frmGeneraInformeLibroBoletas()
        {
            //InitializeComponent();
        }

        private void frmGeneraInformeLibroBoletas_Load(object sender, EventArgs e)
        {
            FuncionesClass config = new FuncionesClass();
            config.CargaConfiguracionInicial();
            this.InicializaControlesDeEmpresa();
            //dtInicio.Value = DateTime.Today;
            //dtFin.Value = DateTime.Today;

            lblInformacion.Text = "EMPRESA: " + FuncionesClass.G_EMPRESANOMBRE;

            icons[0] = new Icon("factura.ico");
            icons[1] = new Icon("xml.ico");

        

            GetLocales();
            lblRut.Text = Convert.ToDouble(FuncionesClass.G_EMPRESARUT.Substring(0, 9)) + "-" + FuncionesClass.G_EMPRESARUT.Substring(9, 1);
            CargaMeses();
            CargaAños();
            
        }
        private void CargaAños()
        {
            int desde = 2010 ;
            int hasta = DateTime.Now.Year;
            int indice = 0;

            for (desde = 2010; desde <= hasta; desde++)
            {
                ddLAno.Items.Add(desde.ToString());
                if (DateTime.Now.Year.ToString() == desde.ToString())
                {
                    indice = desde - 1;
                }
                indice++;
            }
            ddLAno.SelectedIndex = indice;

        }
        private void CargaMeses()
        {
            int i = 1;
            string nombre = "";
            int indice = 0;
            for(i=1;i<=12;i++)
            {
                nombre = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i);
                ddlMes.Items.Add(nombre.ToUpper());

                if(DateTime.Now.Month.ToString() == i.ToString())
                {
                    indice = i-1;
                }
            }

            ddlMes.SelectedIndex = indice;
        }
        private void GetLocales()
        {
            //LocalesClass loc = new LocalesClass(FuncionesClass.G_SERVIDOR);
            //MySqlDataReader dr;

            //dr = loc.getLocalesByCodigoContable("00");
            //if(dr.HasRows == true)
            //{
            //    while(dr.Read())
            //    {
            //        ddLocales.Items.Add(dr["codigo"].ToString() + " " + dr["nombre"].ToString());
            //    }
            //}
            //dr.Close();
            //loc.CerrarTransaccion();

            //ddLocales.SelectedIndex = 0;
        }
        private void InicializaControlesDeEmpresa()
        {
            lblRut.Text = FuncionesClass.G_EMPRESARUT;
            lblNombreEmpresa.Text = FuncionesClass.G_EMPRESANOMBRE;
              
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            notifyIcon1.Icon = icons[currentIcon];
            currentIcon++;
            if (currentIcon == 2)
                currentIcon = 0;
        }

        private void BuscaVentas()
        {
 
        }
    
        private void btnGenerar_Click(object sender, EventArgs e)
        {
           
            string fecha = "";
            string local = "";
            string tipo = "";
    
            local = ddLocales.Text;
            if(rbMensual.CheckState == CheckState.Checked)
            {
                tipo = "MENSUAL";
            }
            if (rbEspecial.CheckState == CheckState.Checked)
            {
                tipo = "ESPECIAL";
            }
            if (rbRectifica.CheckState == CheckState.Checked)
            {
                tipo = "RECTIFICA";
            }
            //frmPopGeneraLibro frm = new frmPopGeneraLibro();
            //frm.lblNombreEmpresa.Text = local;
            //frm.lblFecha.Text = ddLAno.Text + "-" + (ddlMes.SelectedIndex + 1).ToString().PadLeft(2,Convert.ToChar("0") );
            //frm.lblEnvios.Text = tipo;
            //frm.ShowDialog();

        }
        private void CargaFechas()
        {
            //int dias = 20;

            //DateTime date = ddLfecha.Value;
            //DateTime endDate = date.AddDays(-dias);
            //DateTime paso = date.AddDays(1);
            //while (endDate <= date )
            //{
            //    paso = paso.AddDays(-1);
            //    endDate = endDate.AddDays(1);
            //    gvInforme.Rows.Add(paso.ToShortDateString(), ddLocales.Text, "", "","NO ENVIADO",null);
                
            //    this.VerificaRcof(FuncionesClass.GetFechaMysql(paso.ToShortDateString()), gvInforme.Rows.Count - 1);
            //}
        }
        private void VerificaRcof(string xFecha, int xIndice)
        {
            //CafFoliosClass fo = new CafFoliosClass(FuncionesClass.G_SERVIDOR,FuncionesClass.BASE_DTE);
            //MySqlDataReader dr = fo.BuscaRCOF(ddLocales.Text.Substring(0, 2), xFecha);

            //if(dr.HasRows == true)
            //{
            //    if(dr.Read())
            //    {
            //        //gvInforme.Rows[xIndice].Cells[2].Value = dr["fae_fechaenvio_sii"].ToString();
            //        //gvInforme.Rows[xIndice].Cells[3].Value = dr["fae_horaenvio_sii"].ToString();
            //        //gvInforme.Rows[xIndice].Cells[4].Value = dr["fae_trackenvio_sii"].ToString();
                   
            //        //if (dr["fae_GLOSA_sii"].ToString() == "CORRECTO")
            //        //{
            //        //    gvInforme.Rows[xIndice].Cells[5].Value = Properties.Resources.OK_48;
            //        //}
            //        //else
            //        //{
            //        //    gvInforme.Rows[xIndice].Cells[5].Value = Properties.Resources.icons8_exclamacion;
            //        //}
                    
            //    }
            //}
            //dr.Close();
            //fo.CerrarTransaccion();
        }
        private void frmGeneraDocumentos_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
                timer1.Enabled = true;
                Hide();
                notifyIcon1.Visible = true;
                //notifyIcon1.Icon = SystemIcons.Information;
                notifyIcon1.BalloonTipText = "Esta aplicación se está ejecutando en segundo plano.";
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon1.ShowBalloonTip(100);
                
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            //this.GeneraLibro("");
        }
        private void GeneraInforme()
        {

        }
        //private void GeneraLibro(string xtipo)
        //{
        //    ChileSystems.DTE.Engine.InformacionElectronica.LCV.Detalle detalle ;
        //    List<ChileSystems.DTE.Engine.InformacionElectronica.LCV.Detalle> Detalles = 
        //        new List<ChileSystems.DTE.Engine.InformacionElectronica.LCV.Detalle>();

        //    ChileSystems.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo resumen;
        //    List<ChileSystems.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo> Resumenes =
        //        new List<ChileSystems.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo>();

        //Int32 i = 0;
        //    string xmlLibro="";
        //    LibroCVHandler myLibro = new LibroCVHandler();
        //    int neto = 0;
        //    int exento = 0;
        //    int iva = 0;
        //    int total = 0;
        //    string idLibro ="";
        //    string[] tipos;

        //    int cant33 = 0;
        //    double neto33 = 0;
        //    double exento33 = 0;
        //    double iva33 = 0;
        //    double total33 = 0;

        //    int cant34 = 0;
        //    double neto34 = 0;
        //    double exento34 = 0;
        //    double iva34 = 0;
        //    double total34 = 0;

        //    int cant46 = 0;
        //    double neto46 = 0;
        //    double exento46 = 0;
        //    double iva46 = 0;
        //    double total46 = 0;

        //    int cant56 = 0;
        //    double neto56 = 0;
        //    double exento56 = 0;
        //    double iva56 = 0;
        //    double total56 = 0;

        //    int cant61 = 0;
        //    double neto61 = 0;
        //    double exento61 = 0;
        //    double iva61 = 0;
        //    double total61 = 0;

        //    /************** LLENADO DE DATOS DE LA CARÁTULA  *********/
        //    string prefijoLibro = ""; // ddLMes.SelectedIndex.ToString().PadLeft(2, Convert.ToChar("0")) + ddLAno.Text;
        //    if (xtipo == "VENTA")
        //    {
        //        idLibro = "LV-" + prefijoLibro;
        //        myLibro.tipoOperacion = 1;
        //    }
        //    else
        //    {
        //        idLibro = "LC-" + prefijoLibro;
        //        myLibro.tipoOperacion = 2;
        //    }
        //    myLibro.Id = idLibro;
        //    myLibro.rutEmpresa = lblRut.Text;
        //    myLibro.rutCertificado = FuncionesClass.G_DTE_RUT_ENVIA;
        //    myLibro.Periodo = "";
        //    myLibro.FechaResolucion = FuncionesClass.G_DTE_FECHARESOLUCION;
        //    myLibro.NumeroResolucion = FuncionesClass.G_DTE_NUMERO_RESOLUCION;
        //    myLibro.tipoLibro = 2;
        //    myLibro.tipoEnvio = 1;
        //    myLibro.FolioNotificacion = 1;

        //    //for (i=0; i<= gvInforme.Rows.Count -1;i++)
        //    //{
        //    //    detalle = new ChileSystems.DTE.Engine.InformacionElectronica.LCV.Detalle();
        //    //    neto = Convert.ToInt32(gvInforme.Rows[i].Cells[6].Value.ToString().Replace(".", ""));
        //    //    exento = Convert.ToInt32(gvInforme.Rows[i].Cells[7].Value.ToString().Replace(".", ""));
        //    //    iva = Convert.ToInt32(gvInforme.Rows[i].Cells[8].Value.ToString().Replace(".", ""));
        //    //    total = Convert.ToInt32(gvInforme.Rows[i].Cells[9].Value.ToString().Replace(".", ""));

        //    //    detalle.TipoDocumento = (ChileSystems.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro) Convert.ToInt32(gvInforme.Rows[i].Cells[0].Value);
        //    //    detalle.NumeroDocumento = Convert.ToInt32(gvInforme.Rows[i].Cells[1].Value);
        //    //    detalle.FechaDocumento = Convert.ToDateTime(FuncionesClass.GetFechaMysql(gvInforme.Rows[i].Cells[3].Value.ToString()));
        //    //    detalle.RutDocumento = gvInforme.Rows[i].Cells[4].Value.ToString();
        //    //    detalle.RazonSocial = gvInforme.Rows[i].Cells[5].Value.ToString();
        //    //    detalle.MontoExento = exento;
        //    //    detalle.MontoNeto = neto;
        //    //    detalle.MontoIva = iva;
        //    //    detalle.MontoTotal = total;

        //    //    Detalles.Add(detalle);

        //    //    if(gvInforme.Rows[i].Cells[0].Value.ToString() == "33")
        //    //    {
        //    //        cant33 = cant33 + 1;
        //    //        neto33 = neto33 + neto;
        //    //        exento33 = exento33 + exento;
        //    //        iva33 = iva33 + iva;
        //    //        total33 = total33 + total;
        //    //    }
        //    //    if (gvInforme.Rows[i].Cells[0].Value.ToString() == "34")
        //    //    {
        //    //        cant34 = cant34 + 1;
        //    //        neto34 = neto34 + neto;
        //    //        exento34 = exento34 + exento;
        //    //        iva34 = iva34 + iva;
        //    //        total34 = total34 + total;
        //    //    }
        //    //    if (gvInforme.Rows[i].Cells[0].Value.ToString() == "46")
        //    //    {
        //    //        cant46 = cant46 + 1;
        //    //        neto46 = neto46 + neto;
        //    //        exento46 = exento46 + exento;
        //    //        iva46 = iva46 + iva;
        //    //        total46 = total46 + total;
        //    //    }
        //    //    if (gvInforme.Rows[i].Cells[0].Value.ToString() == "56")
        //    //    {
        //    //        cant56 = cant56 + 1;
        //    //        neto56 = neto56 + neto;
        //    //        exento56 = exento56 + exento;
        //    //        iva56 = iva56 + iva;
        //    //        total56 = total56 + total;
        //    //    }
        //    //    if (gvInforme.Rows[i].Cells[0].Value.ToString() == "61")
        //    //    {
        //    //        cant61 = cant61 + 1;
        //    //        neto61 = neto61 + neto;
        //    //        exento61 = exento61 + exento;
        //    //        iva61 = iva61 + iva;
        //    //        total61 = total61 + total;
        //    //    }

               
        //    //}


        //    if(cant33 > 0)
        //    {
        //        resumen = new ChileSystems.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant33;
        //        resumen.TipoDocumento = (ChileSystems.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)33;
        //        resumen.TotalMontoExento = Convert.ToInt32( exento33);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto33);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva33);
        //        resumen.TotalMonto = Convert.ToInt32(total33);
        //        Resumenes.Add(resumen);
        //    }

        //    if (cant34 > 0)
        //    {
        //        resumen = new ChileSystems.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant34;
        //        resumen.TipoDocumento = (ChileSystems.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)34;
        //        resumen.TotalMontoExento = Convert.ToInt32(exento34);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto34);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva34);
        //        resumen.TotalMonto = Convert.ToInt32(total34);
        //        Resumenes.Add(resumen);
        //    }

        //    if (cant46 > 0)
        //    {
        //        resumen = new ChileSystems.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant46;
        //        resumen.TipoDocumento = (ChileSystems.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)46;
        //        resumen.TotalMontoExento = Convert.ToInt32(exento46);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto46);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva46);
        //        resumen.TotalMonto = Convert.ToInt32(total46);
        //        Resumenes.Add(resumen);
        //    }

        //    if (cant56 > 0)
        //    {
        //        resumen = new ChileSystems.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant56;
        //        resumen.TipoDocumento = (ChileSystems.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)56;
        //        resumen.TotalMontoExento = Convert.ToInt32(exento56);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto56);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva56);
        //        resumen.TotalMonto = Convert.ToInt32(total56);
        //        Resumenes.Add(resumen);
        //    }

        //    if (cant61 > 0)
        //    {
        //        resumen = new ChileSystems.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant61;
        //        resumen.TipoDocumento = (ChileSystems.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)61;
        //        resumen.TotalMontoExento = Convert.ToInt32(exento61);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto61);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva61);
        //        resumen.TotalMonto = Convert.ToInt32(total61);
        //        Resumenes.Add(resumen);
        //    }




        //    myLibro.Detalles = Detalles;
        //    myLibro.Resumenes = Resumenes;
        //    xmlLibro = myLibro.GenerateLibroVentas();

        //    if (File.Exists(xmlLibro))
        //    {
        //        FileInfo fi = new FileInfo(xmlLibro);
        //        string destino = @"C:\FAE\LibrosCV\" + idLibro + ".xml";
        //        fi.CopyTo(destino, true);
        //       // string xml = File.ReadAllText(destino, Encoding.GetEncoding("ISO-8859-1"));
               
        //        MessageBox.Show("Libro Generado  Exitosamente");
        //        System.Diagnostics.Process proc = new System.Diagnostics.Process();
        //        proc.EnableRaisingEvents = false;
        //        proc.StartInfo.FileName = destino;
        //        proc.Start();

        //    }

        //}

  

           }
}
