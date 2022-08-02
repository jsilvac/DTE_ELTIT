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
using System.Xml.Serialization;
using System.IO;
using Eltit.Clases;

namespace SamplesDTE
{
    public partial class frmGeneraInformeRCOF : Telerik.WinControls.UI.RadForm
    {
        private Icon[] icons = new Icon[2];
        private int currentIcon = 0;
      

        public frmGeneraInformeRCOF()
        {
            InitializeComponent();
        }

        private void frmGeneraInformeRCOF_Load(object sender, EventArgs e)
        {
            FuncionesClass config = new FuncionesClass();
            config.CargaConfiguracionInicial();
            this.InicializaControlesDeEmpresa();
            //dtInicio.Value = DateTime.Today;
            //dtFin.Value = DateTime.Today;

            lblInformacion.Text = "EMPRESAS ELTIT";

            icons[0] = new Icon("factura.ico");
            icons[1] = new Icon("xml.ico");

            ddLfecha.Value = DateTime.Now;

            //CargaClientes();
            CargaEmpresas();

            lblservidor.Text = FuncionesClass.host_direccion;

            FuncionesClass fu = new FuncionesClass();

            if(fu.PingToHost(lblservidor.Text) == true)
            {
                picConectado.Image = Eltit.Properties.Resources.icons8_exclamacion;
            }
            else
            {
                picConectado.Image = Eltit.Properties.Resources.icons8_exclamacion;
            }

            gvEmpresas.TableElement.Font = new Font("Arial", 8);
            RadPageView1.SelectedPage = RadPageViewPage1;


        }
        private void CargaEmpresas()
        {
            Empresas empresas = new Empresas(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = empresas.GetEmpresasBoleta();
            object img = null;
            FuncionesClass fu = new FuncionesClass();

            if(dr.HasRows == true)
            {
                while (dr.Read())
                {

                    if (fu.PingToHost(dr["ip_servidor"].ToString()) == true)
                    {
                        img = Eltit.Properties.Resources.OK_48;
                    } else {
                        img = Eltit.Properties.Resources.icons8_exclamacion;

                    }


                    gvEmpresas.Rows.Add(dr["codigo_contable"].ToString(), dr["razon_social"].ToString(), dr["ip_servidor"].ToString(), dr["codigo_contable"].ToString(), img, dr["rut"].ToString());
                }
            }

            dr.Close();
            empresas.CerrarTransaccion();               


        }
        private void CargaClientes()
        {
            Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR,FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS );
            MySqlDataReader dr;
            int i = 0;
            dr = cli.GetClientesDTE();
            FuncionesClass fu = new FuncionesClass();
            object img;
           

            if (dr.HasRows == true)
            {
                if(dr.HasRows == true)
                {
                    while (dr.Read())
                    {
                        if (fu.PingToHost(dr["ip_servidor"].ToString()) == true)
                        {
                            img = Eltit.Properties.Resources.OK_48;
                        }
                        else
                        {
                            img = Eltit.Properties.Resources.icons8_exclamacion;
                        }

                        gvEmpresas.Rows.Add((i + 1), dr["prefijo"].ToString(), dr["ip_servidor"].ToString(), dr["servidor_destino"].ToString(), img, dr["rut"].ToString());
                        i++;
                    }
                }
            }

            dr.Close();
            
          
        }
        private void gvPagos_CellClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            CargaCliente();
        }
        private void CargaCliente()
        {
            gvInforme.Rows.Clear();
            gvLocales.Rows.Clear();

            if (gvEmpresas.Rows.Count > 0)
            {
                int index = 0;
                string cliente = "";
                string rut = "";
                string cod_conta = "";
                string ip_cliente = "";

                index = gvEmpresas.CurrentRow.Index;

                cliente = FuncionesClass.G_CLIENTE_PREFIJO;
                cod_conta = gvEmpresas.Rows[index].Cells[0].Value.ToString();
                rut = gvEmpresas.Rows[index].Cells[5].Value.ToString();
                ip_cliente = gvEmpresas.Rows[index].Cells[2].Value.ToString();
                CargaDatosCertificado(cliente, rut);
                CargaLocalesEmpresa(cod_conta, FuncionesClass.G_SERVIDOR);

                btnGenerar.Enabled = true;

            }
        }
        private void CargaLocalesEmpresa(string xCodigoConta, string xServidor)
        {

            //lblRoot.Text = "adminerp_general";
            //lblPassword.Text = "fran061cony252agus203elba214";

            Locales loc = new Locales(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS );
            MySqlDataReader dr = loc.getLocalesByCodigoContable(xCodigoConta);

            if(dr.HasRows == true)
            {
                while(dr.Read())
                {
                    gvLocales.Rows.Add(dr["codigo"].ToString(), dr["nombre"].ToString() + "[" + dr["nombrelocal"].ToString() + "]" );
                }
            }

            dr.Close();
            loc.CerrarTransaccion();

        }
        private void CargaDatosCertificado(string xCliente, string xRut)
        {
            //Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            //MySqlDataReader dr;
            //dr = cli.GetClientesByPrefijoRut(xCliente, xRut);

            Empresas emp = new Empresas(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = emp.GetDatoEmpresaByRut(xCliente, xRut);


            if(dr.HasRows == true)
            {
                if (dr.Read())
                {
                    lblCliente.Text = xCliente;
                    lblRut.Text = dr["rut"].ToString();
                    lblNombreEmpresa.Text = dr["razon_social"].ToString();
                    //CargaLocalesByCliente(dr["razon_social"].ToString());

                    lblRutCertificado.Text = dr["rut_certificado"].ToString();
                    lblNombreCertificado.Text = dr["nombre_certificado"].ToString();
                    lblFechaResolucion.Text = dr["fecha_resolucion"].ToString();
                    lblNumeroResolucion.Text = dr["numero_resolucion"].ToString();
                    
                    lblusermysql.Text =   dr["mysql_user"].ToString(); 
                    lblpassmysql.Text = dr["mysql_pass"].ToString();
                    lblIPCliente.Text = dr["ip_servidor"].ToString();

                    FuncionesClass fu = new FuncionesClass();

                    if (fu.PingToHost(dr["ip_servidor"].ToString()) == false)
                    {
                        btnGenerar.Enabled = false;
                    }
                    else
                    {
                        btnGenerar.Enabled = true;
                    }
                }
            }
            dr.Close();
            emp.CerrarTransaccion();

        }

        private void InicializaControlesDeEmpresa()
        {
            //lblRut.Text = FuncionesClass.G_EMPRESARUT;
            //lblNombreEmpresa.Text = FuncionesClass.G_EMPRESANOMBRE;
              
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
            Ventas ventas = new Ventas();

            int totalDocs = 0;

            double neto = 0;
            double exento = 0;
            double iva = 0;
            double total =0;

            MySqlDataReader dr = ventas.getXmlByNroAtencion(FuncionesClass.G_LOCAL,"");

            if(dr.HasRows == true)
            {
                while(dr.Read())
                {
                    string xml = dr["fae_xml"].ToString(); // File.ReadAllText(dr["fae_xml"], Encoding.GetEncoding("ISO-8859-1"));
                    var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);

                    neto = dte.Documento.Encabezado.Totales.MontoNeto;
                    exento = dte.Documento.Encabezado.Totales.MontoExento;
                    iva = dte.Documento.Encabezado.Totales.IVA;
                    total = dte.Documento.Encabezado.Totales.MontoTotal;

                    gvInforme.Rows.Add(dr["fae_tipo"].ToString(), dr["fae_folio"].ToString(), dte.Documento.Encabezado.Totales.TasaIVA,
                                      dte.Documento.Encabezado.IdentificacionDTE.FechaEmision.ToShortDateString(), dte.Documento.Encabezado.Receptor.Rut,
                                      dte.Documento.Encabezado.Receptor.RazonSocial, String.Format("{0:N0}", neto),
                                      String.Format("{0:N0}", exento), String.Format("{0:N0}", iva), String.Format("{0:N0}", total));                    totalDocs = totalDocs + 1;
                }
            }


            lbInfo.Text = "Documentos Generados " + totalDocs;


        }
    
        private void btnGenerar_Click(object sender, EventArgs e)
        {

            //FuncionesClass fu = new FuncionesClass();

            //if (fu.PingToHost(lblservidor.Text) == true)
            //{
            //    picConectado.Image = Properties.Resources.OK_48;
            //    this.Enabled = false;
            //    gvInforme.Rows.Clear();
            //    this.CargaFechas();
            //    this.gvInforme.GridNavigator.Select(this.gvInforme.Rows[0], this.gvInforme.Columns[0]);
            //    this.Enabled = true;
            //    this.Refresh();
            //}
            //else
            //{
            //    MessageBox.Show("No se puede Establecer conexión con el Host Remoto: " + lblservidor.Text,"Error de Conexión");
            //    picConectado.Image = Properties.Resources.icons8_exclamacion;
            //}


            picConectado.Image = Eltit.Properties.Resources.OK_48;
            this.Enabled = false;
            gvInforme.Rows.Clear();
            this.CargaFechas();
            this.gvInforme.GridNavigator.Select(this.gvInforme.Rows[0], this.gvInforme.Columns[0]);
            this.Enabled = true;
            this.Refresh();


        }
        private void CargaFechas()
        {
            int dias = Convert.ToInt32(txtDias.Text);

            DateTime date = ddLfecha.Value;
            DateTime endDate = date.AddDays(-dias);
            DateTime paso = date.AddDays(1);

            int index = gvEmpresas.CurrentRow.Index;
            string cod_empresa = gvEmpresas.Rows[index].Cells[0].Value.ToString();
            string nombre_empresa = gvEmpresas.Rows[index].Cells[1].Value.ToString();
            string servidor = gvEmpresas.Rows[index].Cells[2].Value.ToString();

            while (endDate <= date )
            {
                paso = paso.AddDays(-1);
                endDate = endDate.AddDays(1);
                gvInforme.Rows.Add(paso.ToShortDateString(), cod_empresa + " " + nombre_empresa, "", "","NO ENVIADO",null);
                
                this.VerificaRcof( cod_empresa, FuncionesClass.GetFechaMysql(paso.ToShortDateString()), gvInforme.Rows.Count - 1, servidor);
            }
        }
        private void VerificaRcof(string cod_empresa, string xFecha, int xIndice, string xServidor)
        {
          
            string bdatos = lblCliente.Text + "dte_" + Convert.ToDouble(lblRut.Text.Substring(0,9)) ;
            PlaceSoft.Eltit.Class.clases.DTEClass dte = new PlaceSoft.Eltit.Class.clases.DTEClass(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS );
           // DTEClass dte = new DTEClass(lblCliente.Text, "127.0.0.1", "root", "1121");

            MySqlDataReader dr = dte.BuscaRCOF(cod_empresa, xFecha, bdatos);
            
            if(dr.HasRows == true)
            {
                if(dr.Read())
                {
                    gvInforme.Rows[xIndice].Cells[2].Value = dr["fae_fechaenvio_sii"].ToString();
                    gvInforme.Rows[xIndice].Cells[3].Value = dr["fae_horaenvio_sii"].ToString();
                    gvInforme.Rows[xIndice].Cells[4].Value = dr["fae_trackenvio_sii"].ToString();
                   
                    if (dr["fae_GLOSA_sii"].ToString() == "CORRECTO")
                    {
                        gvInforme.Rows[xIndice].Cells[5].Value = Eltit.Properties.Resources.OK_48;
                    }
                    else
                    {
                        gvInforme.Rows[xIndice].Cells[5].Value = Eltit.Properties.Resources.icons8_exclamacion;
                    }
                    
                }
            }
            dr.Close();
            dte.CerrarTransaccion();
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

    
        private void GeneraInforme()
        {

        }
        //private void GeneraLibro(string xtipo)
        //{
        //    PlaceSoft.DTE.Engine.InformacionElectronica.LCV.Detalle detalle ;
        //    List<PlaceSoft.DTE.Engine.InformacionElectronica.LCV.Detalle> Detalles = 
        //        new List<PlaceSoft.DTE.Engine.InformacionElectronica.LCV.Detalle>();

        //    PlaceSoft.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo resumen;
        //    List<PlaceSoft.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo> Resumenes =
        //        new List<PlaceSoft.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo>();

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

        //    for (i=0; i<= gvInforme.Rows.Count -1;i++)
        //    {
        //        detalle = new PlaceSoft.DTE.Engine.InformacionElectronica.LCV.Detalle();
        //        neto = Convert.ToInt32(gvInforme.Rows[i].Cells[6].Value.ToString().Replace(".", ""));
        //        exento = Convert.ToInt32(gvInforme.Rows[i].Cells[7].Value.ToString().Replace(".", ""));
        //        iva = Convert.ToInt32(gvInforme.Rows[i].Cells[8].Value.ToString().Replace(".", ""));
        //        total = Convert.ToInt32(gvInforme.Rows[i].Cells[9].Value.ToString().Replace(".", ""));

        //        detalle.TipoDocumento = (PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro) Convert.ToInt32(gvInforme.Rows[i].Cells[0].Value);
        //        detalle.NumeroDocumento = Convert.ToInt32(gvInforme.Rows[i].Cells[1].Value);
        //        detalle.FechaDocumento = Convert.ToDateTime(FuncionesClass.GetFechaMysql(gvInforme.Rows[i].Cells[3].Value.ToString()));
        //        detalle.RutDocumento = gvInforme.Rows[i].Cells[4].Value.ToString();
        //        detalle.RazonSocial = gvInforme.Rows[i].Cells[5].Value.ToString();
        //        detalle.MontoExento = exento;
        //        detalle.MontoNeto = neto;
        //        detalle.MontoIva = iva;
        //        detalle.MontoTotal = total;

        //        Detalles.Add(detalle);

        //        if(gvInforme.Rows[i].Cells[0].Value.ToString() == "33")
        //        {
        //            cant33 = cant33 + 1;
        //            neto33 = neto33 + neto;
        //            exento33 = exento33 + exento;
        //            iva33 = iva33 + iva;
        //            total33 = total33 + total;
        //        }
        //        if (gvInforme.Rows[i].Cells[0].Value.ToString() == "34")
        //        {
        //            cant34 = cant34 + 1;
        //            neto34 = neto34 + neto;
        //            exento34 = exento34 + exento;
        //            iva34 = iva34 + iva;
        //            total34 = total34 + total;
        //        }
        //        if (gvInforme.Rows[i].Cells[0].Value.ToString() == "46")
        //        {
        //            cant46 = cant46 + 1;
        //            neto46 = neto46 + neto;
        //            exento46 = exento46 + exento;
        //            iva46 = iva46 + iva;
        //            total46 = total46 + total;
        //        }
        //        if (gvInforme.Rows[i].Cells[0].Value.ToString() == "56")
        //        {
        //            cant56 = cant56 + 1;
        //            neto56 = neto56 + neto;
        //            exento56 = exento56 + exento;
        //            iva56 = iva56 + iva;
        //            total56 = total56 + total;
        //        }
        //        if (gvInforme.Rows[i].Cells[0].Value.ToString() == "61")
        //        {
        //            cant61 = cant61 + 1;
        //            neto61 = neto61 + neto;
        //            exento61 = exento61 + exento;
        //            iva61 = iva61 + iva;
        //            total61 = total61 + total;
        //        }

               
        //    }


        //    if(cant33 > 0)
        //    {
        //        resumen = new PlaceSoft.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant33;
        //        resumen.TipoDocumento = (PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)33;
        //        resumen.TotalMontoExento = Convert.ToInt32( exento33);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto33);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva33);
        //        resumen.TotalMonto = Convert.ToInt32(total33);
        //        Resumenes.Add(resumen);
        //    }

        //    if (cant34 > 0)
        //    {
        //        resumen = new PlaceSoft.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant34;
        //        resumen.TipoDocumento = (PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)34;
        //        resumen.TotalMontoExento = Convert.ToInt32(exento34);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto34);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva34);
        //        resumen.TotalMonto = Convert.ToInt32(total34);
        //        Resumenes.Add(resumen);
        //    }

        //    if (cant46 > 0)
        //    {
        //        resumen = new PlaceSoft.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant46;
        //        resumen.TipoDocumento = (PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)46;
        //        resumen.TotalMontoExento = Convert.ToInt32(exento46);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto46);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva46);
        //        resumen.TotalMonto = Convert.ToInt32(total46);
        //        Resumenes.Add(resumen);
        //    }

        //    if (cant56 > 0)
        //    {
        //        resumen = new PlaceSoft.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant56;
        //        resumen.TipoDocumento = (PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)56;
        //        resumen.TotalMontoExento = Convert.ToInt32(exento56);
        //        resumen.TotalMontoNeto = Convert.ToInt32(neto56);
        //        resumen.TotalMontoIva = Convert.ToInt32(iva56);
        //        resumen.TotalMonto = Convert.ToInt32(total56);
        //        Resumenes.Add(resumen);
        //    }

        //    if (cant61 > 0)
        //    {
        //        resumen = new PlaceSoft.DTE.Engine.InformacionElectronica.LCV.TotalPeriodo();
        //        resumen.CantidadDocumentos = cant61;
        //        resumen.TipoDocumento = (PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)(int)61;
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

        private void gvInforme_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            int index = gvInforme.CurrentCell.RowIndex;
            string fecha = "";
            string local = "";
            string track = "";

            fecha = gvInforme.Rows[index].Cells[0].Value.ToString();
            local = gvInforme.Rows[index].Cells[1].Value.ToString();
            track = gvInforme.Rows[index].Cells[4].Value.ToString();

            frmPopEnviaRCOF frm = new frmPopEnviaRCOF();
            frm.lblCliente.Text = lblCliente.Text.Replace("_","");
            frm.lblNombreEmpresa.Text = local;
            frm.lblFecha.Text = fecha;
            frm.lblRut.Text = lblRut.Text;
            frm.lblServidor.Text = this.lblIPCliente.Text;
            frm.lblRutCertificado.Text = lblRutCertificado.Text;
            frm.lblNombreCertificado.Text = lblNombreCertificado.Text;
            frm.lblFechaResolucion.Text = lblFechaResolucion.Text;
            frm.lblNumeroResolucion.Text = lblNumeroResolucion.Text;
            frm.mysql_root = lblusermysql.Text;
            frm.mysql_pass = lblpassmysql.Text;
            
            if(chbMuestraDetalle.Checked == true)
            {
                frm.muestraDetalles = true;
            }
            else
            {
                frm.muestraDetalles = false;
                
            }
            frm.CargaLocales(this.gvLocales);
    

            if(track != "NO ENVIADO")
            {
                frm.lblTrack.Text = track;
                frm.txtSecuencia.Text = "2";
            }
            else
            {
                frm.lblTrack.Text = "";
                frm.txtSecuencia.Text = "1";
            }
            frm.ShowDialog();
            frm.Dispose();
            this.btnGenerar_Click(null,null);
            
        }

        private void gvInforme_Click(object sender, EventArgs e)
        {

        }

        private void gvPagos_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void gvInforme_DoubleClick(object sender, EventArgs e)
        {

        }
    }
}
