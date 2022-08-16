using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using MySql.Data.MySqlClient;
using System.IO;
using Telerik.WinControls.UI;
using Eltit.Clases;
using PlaceSoft.Eltit.Class.clases;
using PlaceSoft.Eltit.Functions;
using PlaceSoftDTE.clases;
using Eltit.DTE.Forms;
using PlaceDTE;

namespace Eltit
{

    public partial class frmReimprimir : Telerik.WinControls.UI.RadForm
    {
        public string mysql_root;
        public string mysql_pass;
        private Icon[] icons = new Icon[2];
        private int currentIcon = 0;
        Handler handler = new Handler();
        List<PlaceSoft.DTE.Engine.Documento.DTE> dtes = new List<PlaceSoft.DTE.Engine.Documento.DTE>();
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        string BASE_DTE = "";
        string BASE_VENTAS = "";
        private string cliente_sistema;
        private string XML_DETALLE ="";
        private int countGenerados = 0;
        List<string> rangoUtilizado39 = new List<string>();
        List<string> rangoAnulado39 = new List<string>();

        List<string> rangoUtilizado41 = new List<string>();
        List<string> rangoAnulado41 = new List<string>();

        List<string> rangoUtilizado61 = new List<string>();
        List<string> rangoAnulado61 = new List<string>();

        public bool muestraDetalles;
        public DataTable dt_locales;

        List<PlaceSoft.DTE.Engine.Documento.DTE> dtesSII = new List<PlaceSoft.DTE.Engine.Documento.DTE>();
        List<string> xmlDtes = new List<string>();


        Funciones fun = new Funciones("eltit_");

        private Ventas dc;
        //private VentasClass cj;
        private Cajera cj;
        Documentos doc;
        ClienteFactura clienteFactura;


        public frmReimprimir()
        {
            InitializeComponent();
        }

        private void frmReimprimir_Load(object sender, EventArgs e)
        {
       
            try
            {
                 GetEmpresasContables();
                gvInforme.Columns[5].ReadOnly = false;


            }
            catch(Exception ex)
            {
                MessageBox.Show("Error:[" + ex.Message.ToString() + "]");
            }

            LlenaTipoDocumento();
            dtpHasta.Value = DateTime.Now;
            DateTime date = DateTime.Now;
            date = date.AddDays(-15);
            dtpdesde.Text = date.ToString();
        }

        private void ddLlocales_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            if (ddLlocales.SelectedIndex > 0)
            {
                LeeDatosLocal();
                txtCaja.Focus();

            }
        }
        private void CargaLocales()
        {
            Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);

            MySqlDataReader dr = cli.GetLocalesByEmpresa(ddlEmpresas.Text.Substring(0, 2));

            ddLlocales.Items.Clear();
            ddLlocales.Items.Add("-- SELECCIONE UN LOCAL --");

            if (dr.HasRows == true)
            {
                while (dr.Read())
                {
                    ddLlocales.Items.Add(dr["codigo"].ToString() + " " + dr["nombrelocal"].ToString());
                }
            }

            dr.Close();
            cli.CerrarTransaccion();

            ddLlocales.SelectedIndex = 0;

        }
        private void ddlEmpresas_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            if (ddlEmpresas.SelectedIndex != 0)
            {
                LeeDatosEmpresa();
                CargaLocales();

            }
        }
        private void LeeDatosLocal()
        {
            Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = cli.GetDatosLocalByCodigo(ddLlocales.Text.Substring(0, 2));
            FuncionesClass fu = new FuncionesClass();

            if (dr.HasRows == true)
            {
                if (dr.Read())
                {
                    lblDireccion.Text = dr["direccion"].ToString();            
                    lblServidorVentas.Text = dr["servidor_ventas"].ToString();
                    lblRutEmpresa.Text = dr["rut"].ToString();
                    if (fu.PingToHost("192.168.4.9") == true)
                    {
                            if(lblServidorVentas.Text == "192.168.4.9")
                            {
                                lblRoot.Text = "adminerp_general";
                                lblPassword.Text = "fran061cony252agus203elba214";
                            }
                            else
                            {
                                lblRoot.Text = "conta";
                                lblPassword.Text = "conta";
                            }

                            txxCajaFolios.Enabled = true;
                            txtDesde.Enabled = true;
                            txtHasta.Enabled = true;

                            pbStatus.Image = global::Eltit.Properties.Resources.OK_48;
                   
                            txtNumero.Enabled = true;
                            btnVer.Enabled = true;
                            txtCaja.Enabled = true;
                            ddlTipoDocumento.Enabled = true;
                            ddlTipoDocumento.Focus();

                            //txxCajaFolios.Focus();
                    }
                    else
                    {
                          btnVer.Enabled = false;
                            pbStatus.Image = global::Eltit.Properties.Resources.icons8_exclamacion;
                            RadMessageBox.Show(this, "No se puede establecer conexión con el Host de Destino [" + lblServidorVentas.Text + "]", "Atencion", MessageBoxButtons.OK);
                    }

                }
            }

            dr.Close();
            cli.CerrarTransaccion();

        }

        private void LlenaTipoDocumento(){

               // ddlTipoDocumento.Items.Clear();
                ddlTipoDocumento.Items.Add("-- SELECCIONE TIPO --");
                ddlTipoDocumento.Items.Add("33 - Factura Electrónica ");
                ddlTipoDocumento.Items.Add("39 - Boleta Electrónica ");
                ddlTipoDocumento.Items.Add("61 - Notas de Credtito facturas ");
                ddlTipoDocumento.Items.Add("61 - Notas de Credtito boletas ");
                ddlTipoDocumento.Items.Add("52 - Guias ");

                ddlTipoDocumento.SelectedIndex = 0;
        }
    


        private void GetEmpresasContables()
        {
            Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = cli.GetClientesDTE();

            ddlEmpresas.Items.Clear();
            ddlEmpresas.Items.Add("-- SELECCIONE EMPRESA --");

            if (dr.HasRows == true)
            {
                while (dr.Read())
                {
                    ddlEmpresas.Items.Add(dr["codigo_contable"].ToString() + " " + dr["razon_social"].ToString());
                }
            }

            dr.Close();
            cli.CerrarTransaccion();

            ddlEmpresas.SelectedIndex = 0;
            
        }
        private void LeeDatosEmpresa()
        {
            Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = cli.GetEmpresaByCodigo(ddlEmpresas.Text.Substring(0, 2));
            FuncionesClass fu = new FuncionesClass();

            if (dr.HasRows == true)
            {
                if (dr.Read())
                {
                    lblComuna.Text = dr["comuna"].ToString();
                    lblCiudad.Text = dr["ciudad"].ToString();
                    lblDireccionEmpresa.Text = dr["direccion"].ToString();
                    lblGiro.Text = dr["giro"].ToString();

                    lblRutCertificado.Text    = dr["rut_certificado"].ToString();
                    lblNombreCertificado.Text = dr["nombre_certificado"].ToString();
                    lblFechaResolucion.Text   = dr["fecha_resolucion"].ToString();
                    lblNumeroResolucion.Text  = dr["numero_resolucion"].ToString();
                    lblCodigoSucursal.Text    = dr["codigo_sucursal_sii"].ToString();
                    lblActeco.Text            = dr["acteco_principal"].ToString();
                }
            }

            dr.Close();
            cli.CerrarTransaccion();

        }

        private void BuscaBoletas()
        {

        }
        private void btnVer_Click(object sender, EventArgs e)
        {
            //if(ddlEmpresas.SelectedIndex > 0 && ddLlocales.SelectedIndex > 0 )
            txtNumero.Text = txtNumero.Text.PadLeft(10, Convert.ToChar("0"));
            if (ddlEmpresas.SelectedIndex > 0 &&   ddLlocales.SelectedIndex > 0 && ddlTipoDocumento.SelectedIndex > 0)
            {
                if (txxCajaFolios.Text.Length != 2)
                {
                RadMessageBox.Show(this, "Debe Selecionar Una Caja de venta Válida. [" + txxCajaFolios.Text + "]", "Atencion", MessageBoxButtons.OK);
                }
                else
                {

                    GetDocumentos();
                }
            }
            else
            {
                RadMessageBox.Show(this, "Debe Selecionar Una Empresa, un Local y Tipo de Documento. [" + ddLlocales.Text + "]", "Atencion", MessageBoxButtons.OK);
            }
        }
        private void BuscaCaf(string xTipo)
        {
        
            PlaceSoft.Eltit.Class.clases.Caf myCaf = new PlaceSoft.Eltit.Class.clases.Caf(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);

             MySqlDataReader dr = myCaf.GetCafByCajaLocal(ddLlocales.Text.Substring(0, 2), "", xTipo);
     

            string filePath = "";
            string xmlString = "";
           

            if (dr.HasRows == true)
            {
                while(dr.Read())
                {
                    filePath = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + Convert.ToDouble(lblRutEmpresa.Text.Substring(0, 9)) + @"\Produccion\Caf\" + ddLlocales.Text.Substring(0, 2) + @"\";
                    filePath =  filePath + string.Format("{0}_{1}_{2}.dat", Convert.ToInt32(xTipo),dr["desde"].ToString(), dr["hasta"].ToString());

                    if(!File.Exists(filePath))
                    {

                        FileStream fst;
                        BinaryWriter bw;
                        string tmp_path = @"C:\temp\" + DateTime.Now.Ticks + ".xml";

                        fst = new FileStream(tmp_path, FileMode.OpenOrCreate, FileAccess.Write);
                        bw = new BinaryWriter(fst);
                        string strxml = dr["xml"].ToString().Replace("±	","");

                        Encoding ByteConverter = Encoding.GetEncoding("ISO-8859-1");
                        byte[] textEnBytes = ByteConverter.GetBytes(strxml);

                        bw.Write(textEnBytes);
                        bw.Flush();
                        bw.Close();
                        bw.Dispose();

                        using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                        {
                            var xml = File.ReadAllBytes(tmp_path);
                            xmlString = File.ReadAllText(tmp_path, Encoding.GetEncoding("ISO-8859-1"));
                            
                            fs.Write(xml, 0, xml.Length);
                            fs.Flush();
                            fs.Close();
                        }
                    }
             
                }
            }

            dr.Close();
            myCaf.CerrarTransaccion();
            
        }
        private void GetDocumentos()
        {
            PlaceSoft.Eltit.Class.clases.Documentos dc = new PlaceSoft.Eltit.Class.clases.Documentos(FuncionesClass.G_SERVIDORMASTER, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            //PlaceSoft.Eltit.Class.clases.Documentos dc = new PlaceSoft.Eltit.Class.clases.Documentos(lblPrincipal.Text, lblRootPrincipal.Text, lblPassPrincipal.Text);
            MySqlDataReader dr = null;
            string base_datos = FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + ddLlocales.Text.Substring(0,2) ;
            object img = null;
            int count = 0;
            string tipo = "";
            string tipoFiscal = "";
                string fechadesde = "";
                string fechahasta = "";

            if (ddlTipoDocumento.SelectedIndex > 0)
            {
                    switch (ddlTipoDocumento.SelectedItem.Text.Substring(0,2))
                    {
                    case "33":
                            tipo = "FV";
                            tipoFiscal = "33";
                            break;
                    case "61":
                            tipo = "NB";
                            tipoFiscal = "61";
                            break;
                    case "52":
                            tipo = "G4";// "GV";
                            tipoFiscal = "52";
                            break;
                    case "39":
                            tipo = "BV";
                            tipoFiscal = "39";
                            break;
                    default:

                            MessageBox.Show("Documento sin tipo");
                            break;

                    }

                    txtNumero.Focus();
            }

                fechadesde = fun.FechaMysql(dtpdesde.Text);
                fechahasta = fun.FechaMysql(dtpHasta.Text);

                this.BuscaCaf(tipoFiscal);


            if (tipo == "G4")
            {
                dr = dc.GetDocumentosGuasByLocalNroInternoCajaDesdeHasta(ddLlocales.Text.Substring(0, 2), txxCajaFolios.Text, tipo, txtDesde.Text,
                    txtHasta.Text, base_datos);
            }
            else
            {

                    dr = dc.GetDocumentosCabezaByLocalNroInternoCajaFechadesdeFechahasta(ddLlocales.Text.Substring(0, 2), txxCajaFolios.Text,tipo,txtNumero.Text,fechadesde,fechahasta,base_datos);
            }



            string Nombre_Doc = "";
            string xml = "0";
            gvInforme.Rows.Clear();
            if(dr.HasRows == true)
            {
                ddlEmpresas.Enabled = false;
                ddLlocales.Enabled = false;
                btnGenera.Enabled = true;
                FuncionesClass fun = new FuncionesClass();
               
                while (dr.Read())
                {
                    Nombre_Doc = FuncionesClass.getNombredocumentoByCodigo2(dr["tipo"].ToString());
                    xml = dr["xml"].ToString();

                    //  img = Properties.Resources.icon_imprimir_31;


                    gvInforme.Rows.Add(dr["tipo"].ToString() , dr["numero"].ToString(), dr["fecha"].ToString(), dr["caja"].ToString(),
                            String.Format("{0:N0}", dr["total"]), dr["foliosii"].ToString(), false, dr["indicador_traslado"].ToString(), 
                            dr["rut"].ToString(), dr["foliosii"].ToString(), Properties.Resources.icon_xml_15, Properties.Resources.icons_imprimir_15, false,dr["cajera"].ToString(),
                            "", Properties.Resources.icon_pdf_25, Properties.Resources.icon_sobre);
                    count++;
                    
                    fun.ColoreaCelda(gvInforme.Rows[gvInforme.CurrentRow.Index].Cells[6], Color.LightBlue);
                }


            }
            else
            {
                RadMessageBox.Show(this, "No se encontraon resultados con los parametros indicados.", "Atencion", MessageBoxButtons.OK);
            }

            dr.Close();
            dc.CerrarTransaccion();

            lblInfo.Text = "Total Registros " + count.ToString();

            btnLimpiar.Enabled = true;
        }

        private void SugiereUltimoCaf( string xTipo, string xCaja)
        {
            PlaceSoft.Eltit.Class.clases.DTEClass dte = new PlaceSoft.Eltit.Class.clases.DTEClass(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            int folio = dte.GetUltimoFolioDTEByLocalCaja(ddLlocales.Text.Substring(0, 2), xTipo, "", FuncionesClass.G_CLIENTE_PREFIJO + "fae" + ddLlocales.Text.Substring(0, 2));

            txtFolioGenera.Text = folio.ToString();
        }

        private void InicializaControlesDeEmpresa()
        {
           
              
        }

        private void btnGenera_Click(object sender, EventArgs e)
        {
            if(gvInforme.Rows.Count > 0)
            {
                if(RadMessageBox.Show(this, "Atención, va Re Generar los XML selecionados y generar unos nuevos con los mismos datos " +
                    "¿Desea Realmente Hacerlo?", "Atencion", MessageBoxButtons.YesNo, RadMessageIcon.Exclamation) == DialogResult.Yes)
                {
                    countGenerados = 0;
                    RegeneraDocumentos();
                }
               
            }
        }


        private void RegeneraDocumentos()
        {
            int i = 0;
            bool check = false;
            string tipo = "";
            string numero = "";
            string fecha = "";
            string caja = "";
            int count = 0;
            double MONTO = 0;
            string tipo_traslado = "";
     
            for (i=0; i <= gvInforme.Rows.Count-1;i++ )
            {
                check =  Convert.ToBoolean(gvInforme.Rows[i].Cells[7].Value);
                if(check == true)
                {
                    tipo = gvInforme.Rows[i].Cells[0].Value.ToString();
                    numero = gvInforme.Rows[i].Cells[1].Value.ToString();
                    fecha = gvInforme.Rows[i].Cells[2].Value.ToString();
                    caja = gvInforme.Rows[i].Cells[3].Value.ToString();
                    MONTO = Convert.ToDouble(gvInforme.Rows[i].Cells[4].Value.ToString());
                    tipo_traslado = gvInforme.Rows[i].Cells[3].Value.ToString();

                    this.GenerarDocumentoElectronico(tipo, numero, fecha, caja, MONTO, i);
                    count++;
                    countGenerados++;
                    
                }

            }


            if(count == 0)
            {
                RadMessageBox.Show(this, "Debe Seleccionar Al menos un Documento.", "Atencion", MessageBoxButtons.OK);
            }
            else
            {
                RadMessageBox.Show(this, "Se Re Generaron " + count + " Documentos tipo [" + tipo +"]", "Atencion", MessageBoxButtons.OK);
                btnGenera.Enabled = false;


                if (RadMessageBox.Show(this, "¿Desea Enviar los Documentos Procesados a Impuestos Internos Ahora?", "Atencion", MessageBoxButtons.YesNo) == DialogResult.Yes )
                {
                    //  PlaceSoft.Eltit.Handler.Handler handler = new PlaceSoft.Eltit.Handler.Handler(FuncionesClass.G_CLIENTE_PREFIJO);
                    //var EnvioSII = handler.GenerarEnvioDTEToSII(dtes, xmlDtes);

                    PlaceSoft.Eltit.Handler.Handler handler = new PlaceSoft.Eltit.Handler.Handler(FuncionesClass.G_CLIENTE_PREFIJO);
                    handler.rutCertificado = lblRutCertificado.Text;
                    handler.nombreCertificado = lblNombreCertificado.Text;
                    handler.fechaResolucion = Convert.ToDateTime(lblFechaResolucion.Text);
                    handler.numero_resolucion = Convert.ToInt32(lblNumeroResolucion.Text);
                    handler.emisor_rut = Convert.ToDouble(lblRutEmpresa.Text.Substring(0,9)) + "-" + lblRutEmpresa.Text.Substring(9,1);

                    var EnvioSII = handler.GenerarEnvioDTEToSII(dtesSII, xmlDtes);
                    var filePath = handler.FirmarEnvioDTE(EnvioSII);

                    if (File.Exists(filePath))
                    {
                        FileInfo fi = new FileInfo(filePath);
                        //  string destino =  @"C:\Envios\ENVSETBASICO_" + FuncionesClass.G_DTE_CASO_BASICO +"_"+ DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xml";
                        string destino = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + Convert.ToDouble(lblRutEmpresa.Text.Substring(0, 9)) + @"\Produccion\envios\" + ddLlocales.Text.Substring(0, 2) + @"\ENVIO_DTE_" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xml";
                        
                        fi.CopyTo(destino, true);
                        string xml = File.ReadAllText(destino, Encoding.GetEncoding("ISO-8859-1"));


                        //this.GrabaXML(ddTipoDocumento.Text.Substring(0, 2), txtFolioFiscal.Text, FuncionesClass.GetFechaMysql(dtInicio.Text), xml);
                        //MessageBox.Show("Sobre Generado exitosamente");
                        //this.Retorno();

                        string pathFile = filePath;
                        long trackId = handler.EnviarEnvioDTEToSII(destino, "XH1F-EFZ5-ZH93", true);
                        this.GrabaEnvio(trackId);
                        RadMessageBox.Show(this, "Sobre enviado correctamente. TrackID: " + trackId.ToString(), "Atencion", MessageBoxButtons.OK);
                       
                    }

                }





            }



        }

        private void GrabaEnvio(long xTrack)
        {
            DTEClass mydte = new DTEClass(lblPrincipal.Text,lblRootPrincipal.Text,lblPassPrincipal.Text);
            int i = 0;
            bool val = false;
            string tipo = "";
            string numero = "";
            string local = ddLlocales.Text.Substring(0, 2);

            for (i = 0; i <= gvInforme.Rows.Count - 1; i++)
            {
                val = Convert.ToBoolean(gvInforme.Rows[i].Cells[7].Value);
                tipo = gvInforme.Rows[i].Cells[0].Value.ToString().Substring(0, 2);
                numero = gvInforme.Rows[i].Cells[10].Value.ToString();
                if (val == true)
                {
                    mydte.ActualizaTrackEnDTE(local, tipo, Convert.ToDouble(numero).ToString(), DateTime.Now.ToString("yyyy-MM-dd"), xTrack.ToString());
                }
            }
                
        }


        private void GenerarDocumentoElectronico(string xTipoInterno, string xNumeroInterno, string xFecha , string xCaja, 
                                            double xMonto, int indice)
        {
            Handler handler;
            handler = new Handler();
            dc = new Ventas( FuncionesClass.G_CLIENTE_PREFIJO, lblPrincipal.Text, lblRootPrincipal.Text, lblPassPrincipal.Text);
            MySqlDataReader dr = null;
            string xtipo = xTipoInterno.Substring(0,2);
            string base_venta = FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + ddLlocales.Text.Substring(0,2);
     
            string rut_venta = gvInforme.Rows[indice].Cells[9].Value.ToString();

            List<ItemBoleta> items;
            items = new List<ItemBoleta>();
            int xTipo_Traslado = Convert.ToInt32(gvInforme.Rows[indice].Cells[8].Value) ;
            string tipoFiscal = "";

            PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronica;

            if (rbAfectas.CheckState == CheckState.Checked)
            {
                xtipo = "FV";
                tipoFiscal = "33";
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronica;
            } 

            if(rbNotasBoletas.CheckState == CheckState.Checked)
            {
                xtipo = "NB";
                tipoFiscal = "61";
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica;
            }
            if (rbNotasFactura.CheckState == CheckState.Checked)
            {
                xtipo = "NF";
                tipoFiscal = "61";
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica;
            }
            if (rbGuias.CheckState == CheckState.Checked)
            {
                xtipo = "G4";//"GV";
                tipoFiscal = "52";
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica;
            }

            if (tipoFiscal == "52")
            {
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica;
                if (xTipo_Traslado == 1)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.OperacionConstituyeVenta;
                }
                if (xTipo_Traslado == 2)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.VentaPorEfectuar;
                }
                if (xTipo_Traslado == 3)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.Consignaciones;
                }
                if (xTipo_Traslado == 4)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.EntregaGratuita;
                }
                if (xTipo_Traslado == 5)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.TrasladosInternos;
                }
                if (xTipo_Traslado == 6)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.OtrosTrasladosNoVenta;
                }
                if (xTipo_Traslado == 7)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.GuiaDeDevolucion;
                }
                if (xTipo_Traslado == 8)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.TrasladoParaExportacion;
                }
                if (xTipo_Traslado == 9)
                {
                    handler.emisor_tipo_traslado = PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.VentaParaExportacion;
                }

            }
                       

            xFecha = FuncionesClass.GetFechaMysql(xFecha);

            if (xtipo == "G4") //"GV"
            {
                dr = dc.getDetalleGuiasByTipoNroCaja(ddLlocales.Text.Substring(0, 2), xtipo, xNumeroInterno, xCaja, xFecha);
            }
            else
            {
                dr = dc.getVentaDetalleDocumentosByTipoNroCaja(ddLlocales.Text.Substring(0, 2), xtipo, xNumeroInterno, xCaja, xFecha);
            }

            if (dr.HasRows == true)
            {


                /********** GET ULTIMO FOLIO FISCAL DTE ****************/
                PlaceSoft.Eltit.Class.clases.DTEClass DTE = new PlaceSoft.Eltit.Class.clases.DTEClass(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
                int FOLIO_CAF = DTE.GetUltimoFolioDTEByLocalCaja(ddLlocales.Text.Substring(0, 2), tipoFiscal, "", base_venta);

                if (txtFolioGenera.Text == "")
                {
                    txtFolioGenera.Text = "0";
                }



                FOLIO_CAF = Convert.ToInt32(gvInforme.Rows[indice].Cells[6].Value);

                /*
                    * Las facturas notas y crédito y guias no tienen correlativos asignados por caga
                    * se manejan por local global                 
                    */


                if (FOLIO_CAF == 0)
                {
                    RadMessageBox.Show(this, "No se ha podido Regenerar el Documento por que el CAF asignado no se encuntra autorizado[" + FOLIO_CAF + "]", "Atencion", MessageBoxButtons.OK);
                    return;
                }

                handler.nombreCertificado = lblNombreCertificado.Text;
                handler.tipo = typeDTE;
                handler.casoPruebas = string.Empty;
                handler.Folio = (int)FOLIO_CAF;
                handler.idDte = "DTE_" + FOLIO_CAF + "T" + tipoFiscal;
                handler.rutcliente = Convert.ToDouble(rut_venta.Substring(0, 9)) + "-" + rut_venta.Substring(9, 1);

                handler.emisor_rut = Convert.ToDouble(lblRutEmpresa.Text.Substring(0, 9)) + "-" + lblRutEmpresa.Text.Substring(9, 1);
                handler.emisor_razon_social = ddlEmpresas.Text.Substring(2, ddlEmpresas.Text.Length - 2);
                handler.emisor_giro = lblGiro.Text;
                handler.emisor_comuna = lblComuna.Text;
                handler.emisor_ciudad = lblCiudad.Text;
                handler.emisor_direccion = lblCiudad.Text;
                handler.cod_sucursal_sii = ""; // no aplica
                handler.emisor_acteco = lblActeco.Text;
                handler.fechaEmision = Convert.ToDateTime(xFecha);
                MySqlDataReader drCliente = null;
                Clientes cliente = new Clientes(lblPrincipal.Text, lblRootPrincipal.Text, lblPassPrincipal.Text);
                drCliente = cliente.getClienteByRutSucursal(rut_venta, "0");

                var dte = handler.GenerateDTEFacturas(drCliente, xFecha, lblActeco.Text);
                Ventas venta = new Ventas(FuncionesClass.G_CLIENTE_PREFIJO, lblPrincipal.Text, lblRootPrincipal.Text, lblPassPrincipal.Text);
                string pago = venta.LeeformaPago(ddLlocales.Text.Substring(0, 2), xNumeroInterno, xtipo, xCaja, xFecha);

                handler.GenerateDetailsFacturas(dte, dr, pago);

                if (tipoFiscal == "56" || tipoFiscal == "61")
                {
                    //string ref_fecha = venta.GetFechaReferencia(ddLlocales.Text.Substring(0, 2), dr["ref_numero"].ToString(), dr["ref_tipo"].ToString());
                    handler.ReferenciaFacturas(dte, ddLlocales.Text.Substring(0, 2), lblPrincipal.Text, lblRootPrincipal.Text, lblPassPrincipal.Text);
                }

                dr.Close();

                var path = handler.TimbrarYFirmarXMLDTE(dte, @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + Convert.ToDouble(lblRutEmpresa.Text.Substring(0, 9)) + @"\Produccion\Caf\" + ddLlocales.Text.Substring(0, 2) + @"\");

                if (File.Exists(path))
                {
                    FileInfo fi = new FileInfo(path);
                    //string destino = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\xml\" + FuncionesClass.G_LOCAL + @"\DTE39F" + FOLIO_CAF + ".xml";
                    string destino = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + Convert.ToDouble(lblRutEmpresa.Text.Substring(0, 9)) + @"\Produccion\xml\" + ddLlocales.Text.Substring(0, 2) + @"\DTE" + tipoFiscal + "F" + FOLIO_CAF + ".xml";
                    fi.CopyTo(destino, true);
                    string XML = File.ReadAllText(destino, Encoding.GetEncoding("ISO-8859-1"));

                    this.GrabaXML(ddLlocales.Text.Substring(0, 2), xTipoInterno, xNumeroInterno, tipoFiscal, FOLIO_CAF,
                                        xCaja, XML, xFecha, dte.Documento.Encabezado.Receptor.RazonSocial, dte.Documento.Encabezado.Receptor.Rut, xMonto);
                    dtesSII.Add(dte);
                    xmlDtes.Add(XML);
                }



            }
            else
            {
                RadMessageBox.Show(this, " No se encontraron los detalles del documento o no se encontraron los impuestos en el documento ! ");
            }


            dr.Close();
            dc.CerrarTransaccion();

        }

        private void GrabaXML(string xLocal, string xTipoInterno,string xNroInterno, string xTipoFiscal, int xFolioFiscal,
                    string xCaja, string XML, string xFecha, string xNombre, string xRut, double xMonto)
        {
            string basedte = FuncionesClass.G_CLIENTE_PREFIJO + "fae" + ddLlocales.Text.Substring(0, 2);
            PlaceSoft.Eltit.Class.clases.DTEClass dte = new PlaceSoft.Eltit.Class.clases.DTEClass(lblPrincipal.Text, lblRootPrincipal.Text,lblPassPrincipal.Text);

            ////////////////////// by jaimiko   2021-08-12    ///////////////////////////
            PlaceSoft.Eltit.Class.clases.DTEClass dte2 = new PlaceSoft.Eltit.Class.clases.DTEClass( lblServidorVentas.Text, lblRootPrincipal.Text, lblPassPrincipal.Text);

            dte.GrabaXML(xLocal, xTipoFiscal, Convert.ToInt32(xFolioFiscal), xTipoInterno.Substring(0,2),  xNroInterno, xFecha, basedte, XML, xCaja, xRut, xNombre, xMonto);

            ////////////////////// by jaimiko   2021-08-12    ///////////////////////////
            dte2.GrabaXML(xLocal, xTipoFiscal, Convert.ToInt32(xFolioFiscal), xTipoInterno.Substring(0, 2), xNroInterno, xFecha, basedte, XML, xCaja, xRut, xNombre, xMonto);

            ////////////////////// by jaimiko   2021-08-12    ///////////////////////////
            ActualizaFolioSII(xNroInterno, xCaja, xTipoInterno, xFecha, xFolioFiscal.ToString().PadLeft(10, Convert.ToChar("0")));
        }
        private void ActualizaFolioSII(string xnrointerno, string xCaja, string xTipointerno,string xFecha ,string xFolioSII)
        {
            string BASE_VENTA = FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + ddLlocales.Text.Substring(0, 2);

            ////////////////////// by jaimiko   2021-08-12    ///////////////////////////
            Documentos dc = new Documentos(lblPrincipal.Text, lblRootPrincipal.Text, lblPassPrincipal.Text);
            Documentos dc2 = new Documentos(lblServidorVentas.Text, lblRootPrincipal.Text, lblPassPrincipal.Text);
            dc.ActualizaFolioSII(ddLlocales.Text.Substring(0, 2), BASE_VENTA, xnrointerno, xTipointerno.Substring(0, 2), xFecha, xCaja, xFolioSII);
            dc2.ActualizaFolioSII(ddLlocales.Text.Substring(0, 2), BASE_VENTA, xnrointerno, xTipointerno.Substring(0, 2), xFecha, xCaja, xFolioSII);
        }

        private int GetUltimoFolio(string xtipo, int xFolio, string xCaja)
        {            
            int ultimo = xFolio;
            string base_dte = FuncionesClass.G_CLIENTE_PREFIJO + "fae" + ddLlocales.Text.Substring(0, 2);
            PlaceSoft.Eltit.Class.clases.DTEClass caf = new PlaceSoft.Eltit.Class.clases.DTEClass(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
     
            /***************** VERIFICA SI EL NUMERO ESTA AUTORIZADO ***********************/
            if (caf.VerificaCaf(ddLlocales.Text.Substring(0,2), xtipo, ultimo.ToString(), xCaja, base_dte) == true)
            {
                return ultimo;
            }
            else
            {
                ultimo = 0;// caf.BuscaRangoSiguiente(ddLlocales.Text.Substring(0, 2),xtipo, ultimo.ToString(),xCaja, base_dte);
            }
                       
            return ultimo;
        }
            private void btn_pdf_clik(object sender, StateChangedEventArgs args)
            {
                MessageBox.Show("estamos hasta el lolli");
            }

        private void radCheckBox1_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if(radCheckBox1.CheckState == CheckState.Checked)
            {
                MarcarCheck(true);
            }
            else
            {
                MarcarCheck(false);
            }
        }

        private void MarcarCheck(bool val)
        {
            int i = 0;

            for(i=0; i <= gvInforme.Rows.Count -1; i++)
            {

                gvInforme.Rows[i].Cells[7].Value = val;
            }

        }

        private void rbFolioNuevo_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if(rbFolioNuevo.CheckState == CheckState.Checked)
            {
                string xtipo = "";
                string tipoFiscal = "";

                if (rbAfectas.CheckState == CheckState.Checked)
                {
                    xtipo = "FV";
                    tipoFiscal = "33";
               
                }

                if (rbNotasBoletas.CheckState == CheckState.Checked)
                {
                    xtipo = "NB";
                    tipoFiscal = "61";
                    
                }
                if (rbNotasFactura.CheckState == CheckState.Checked)
                {
                    xtipo = "NF";
                    tipoFiscal = "61";

                }

                if (rbGuias.CheckState == CheckState.Checked)
                {
                    xtipo = "GV";
                    tipoFiscal = "52";                 
                }
                    SugiereUltimoCaf(tipoFiscal, txxCajaFolios.Text);
                this.BuscaCaf(tipoFiscal);
                if (RadMessageBox.Show(this, "Ha Seleccionado Asignar Folios Comenzando de el Folio fiscal Sugerido ["+ txtFolioGenera.Text +"] por el Sistema ¿Desea Asignar ese Correlativo a los Documentos Actuales?", "Atencion", MessageBoxButtons.YesNo, RadMessageIcon.Exclamation) == DialogResult.Yes)
                {
                    CargarfoliosNuevos();
                }

                //txtFolioGenera.Enabled = true;
                //txtFolioGenera.Focus();
                //txtFolioGenera.SelectAll();

            }
        }
        private void CargarfoliosNuevos()
        {
            int n;
            bool isNumeric = int.TryParse(txtFolioGenera.Text, out n);

            if(isNumeric == true)
            {
                int inicial = Convert.ToInt32(txtFolioGenera.Text);
                int i = 0;


                for (i = 0; i <= gvInforme.Rows.Count - 1; i++)
                {
                    gvInforme.Rows[i].Cells[6].Value = inicial.ToString().PadLeft(10, Convert.ToChar("0"));
                    inicial++;
                }
            }
            

        }
        private void rbMismoFolio_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if(rbMismoFolio.CheckState == CheckState.Checked)
            {
                txtFolioGenera.Text = "0";
                txtFolioGenera.Enabled = false;

                int i = 0;


                for (i = 0; i <= gvInforme.Rows.Count - 1; i++)
                {
                    gvInforme.Rows[i].Cells[6].Value = gvInforme.Rows[i].Cells[10].Value;
                  
                }

            }
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            ddlTipoDocumento.Text = "-- SELECCIONE TIPO --";
            txtNumero.Text = "";
            txxCajaFolios.Text = "";
            Retorno();
        }

        private void Retorno()
        {
            ddlEmpresas.SelectedIndex = 0;
            ddLlocales.Items.Clear();
            lblDireccion.Text = "";
            lblServidorVentas.Text = "";
            gvInforme.Rows.Clear();
            lblInfo.Text = "";
            rbMismoFolio.CheckState = CheckState.Checked;
            radCheckBox1.CheckState = CheckState.Unchecked;
            txtFolioGenera.Text = "0";
            rbAfectas.CheckState = CheckState.Checked;
            txxCajaFolios.Text = "";
            txtDesde.Text = "";
            txtHasta.Text = "";
            lblRutEmpresa.Text = "";
            lblDireccionEmpresa.Text = "";
            lblComuna.Text = "";
            lblCiudad.Text = "";
            lblGiro.Text = "";
            lblCodigoSucursal.Text = "";
            lblActeco.Text = "";
            ddlEmpresas.Enabled = true;
            ddLlocales.Enabled = true;
            


        }

        private void gvInforme_CellClick(object sender, GridViewCellEventArgs e)
        {
            if (gvInforme.Rows.Count > 0)
            {
                int fila = gvInforme.CurrentRow.Index;

                string local = ddLlocales.Text.Substring(0, 2);
                string tipoDTE = ddlTipoDocumento.Text.Substring(0, 2);
                string folioSII = Convert.ToInt32(gvInforme.Rows[fila].Cells[1].Value.ToString()).ToString();
                string fechaEmision = gvInforme.Rows[fila].Cells[2].Value.ToString();
                fechaEmision = fechaEmision.Substring(6, 4) + "-" + fechaEmision.Substring(3, 2) + "-" + fechaEmision.Substring(0, 2);
                doc = new Documentos(FuncionesClass.G_SERVIDORMASTER,FuncionesClass.G_MYSQL_USER,FuncionesClass.G_MYSQL_PASS);
                string caja         = gvInforme.Rows[fila].Cells[3].Value.ToString();
                string rutEmpresa   = Convert.ToDouble(lblRutEmpresa.Text.Substring(0,9)) + "-" + lblRutEmpresa.Text.Substring(9, 1);
                string tipoInterno = gvInforme.Rows[fila].Cells[0].Value.ToString();


                if (gvInforme.CurrentCell.ColumnIndex == 10)
                {

                    this.CargaXML(local, tipoDTE, folioSII, fechaEmision);
                    
                }

                if (gvInforme.CurrentCell.ColumnIndex == 11)
                {   
                    //string nombreImpresora = "POS-80";
                    
                    
                    Boolean impCedible = Convert.ToBoolean(gvInforme.Rows[fila].Cells[12].Value);
                    //Boolean exportaPDF = Convert.ToBoolean(gvInforme.Rows[fila].Cells[15].Value);
                    string cajera = gvInforme.Rows[fila].Cells[13].Value.ToString();
                    List<string> formaPago = new List<string>();
                    formaPago = doc.GetPagosByDocumento(local, tipoInterno, folioSII, caja, fechaEmision);



                    this.imprimeFactura(tipoDTE,fechaEmision,folioSII,local,caja,rutEmpresa,formaPago,tipoInterno, impCedible,cajera);
                }

                if (gvInforme.CurrentCell.ColumnIndex == 15)
                {
                    this.generaPDF(local, tipoDTE, folioSII, fechaEmision, rutEmpresa);
                }

                if (gvInforme.CurrentCell.ColumnIndex == 16)
                {
                    MessageBox.Show("mete aki el correo");
                }

            }
        }

        private void generaPDF(string xLocal, string xTipoDTE, string xFolio, string xFecha, string xRutEmpresa)
        {
            DTEClass dte = new DTEClass(FuncionesClass.G_SERVIDORMASTER, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            string XML = dte.GetXMLFacturas(xLocal, xTipoDTE, xFolio, xFecha);

            int fila = gvInforme.CurrentRow.Index;

            frmGeneraPDF2 este = new frmGeneraPDF2();
            este.DOC_RUT_BASE = xRutEmpresa;
            este.DOC_XML= XML;
            este.DOC_FOLIOSII = xFolio;
            este.DOC_TIPOSII = xTipoDTE;
            este.DOC_LOCAL = xLocal;
            este.DOC_CONTACTO = GetClienteSucursal( gvInforme.Rows[fila].Cells[8].Value.ToString(),"0","contacto");
            este.DOC_FONO = GetClienteSucursal(gvInforme.Rows[fila].Cells[8].Value.ToString(), "0", "fono1");
            este.DOC_CORREO_CONTACTO = GetClienteSucursal(gvInforme.Rows[fila].Cells[8].Value.ToString(), "0", "email");
            este.ShowDialog();


        }

        private string GetClienteSucursal(string xRut, string xSucursal, string xparametro)
        {
            string cli="";
           clienteFactura = new ClienteFactura(FuncionesClass.G_SERVIDORMASTER, FuncionesClass.G_MYSQL_USER);

            MySqlDataReader obtieneCliente = clienteFactura.GetClienteByRutSucursal(xRut, xSucursal);

            if (obtieneCliente.HasRows)
            {
                if (obtieneCliente.Read())
                {
                    cli = obtieneCliente[xparametro].ToString();
                }
            }

            return cli;
        }



        private void CargaXML(string xLocal,string xTipoDTE, string xFolio, string xFecha)
        {

            DTEClass dte = new DTEClass(FuncionesClass.G_SERVIDORMASTER, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            string XML = dte.GetXMLFacturas(xLocal, xTipoDTE, xFolio, xFecha);

            XmlGridViewSample.Form1 frm = new XmlGridViewSample.Form1();

            if (!XML.Contains("<?xml version"))
            {
                XML = "<?xml version='1.0' encoding='ISO-8859-1' ?>" + XML;
            }


            FileStream fst;
            BinaryWriter bw;
            string tmp_path = @"C:\temp\" + DateTime.Now.Ticks + ".xml";
            string tmp_path2 = @"C:\temp\" + DateTime.Now.Ticks + ".xml";
            fst = new FileStream(tmp_path, FileMode.OpenOrCreate, FileAccess.Write);
            bw = new BinaryWriter(fst);
            string strxml = XML;
            Encoding ByteConverter = Encoding.GetEncoding("ISO-8859-1");
            byte[] textEnBytes = ByteConverter.GetBytes(strxml);
            bw.Write(textEnBytes);
            bw.Flush();
            bw.Close();
            bw.Dispose();


            frm.txtFile.Text = tmp_path;
            frm.ShowDialog();
        }
        private string getNombreCajera(string xRut)
        {
            string salida = "";
            string nombreCajera = "";
            Cajera cajera = new Cajera(xRut,"eltit_",FuncionesClass.G_SERVIDORMASTER,"");

            nombreCajera = cajera.GetCajera(xRut);

            return salida = nombreCajera;
        }


        private void imprimeFactura(string xTipoDTE, string xFechaEmision, string xFolioSII, string xLocal, string xCaja, string xRutEmpresa, List<string> xFormaDePago, string xTipoInterno, Boolean xImpCedible, string xCajera) {

            Eltit.DTE.Forms.Form1 formulario = new Form1();
            formulario.DOC_FOLIOSII = xFolioSII;
            formulario.DOC_LOCAL = xLocal;
            formulario.DOC_FECHA_EMISION = xFechaEmision;
            formulario.DOC_TIPO_DTE = xTipoDTE;
            formulario.DOC_RUT = xRutEmpresa;
            formulario.DOC_NOMBRE_CAJERA = getNombreCajera(xCajera);
            formulario.DOC_FORMA_PAGO = xFormaDePago;
            formulario.DOC_CAJA = xCaja;
            formulario.DOC_TIPO_INTERNO = xTipoInterno;
            formulario.IMPRIME_CEDIBLE = xImpCedible;
            formulario.DOC_NOMBRE_IMPRESORA = FuncionesClass.G_IMPRESORA_TICKET;
            formulario.DOC_SERVIDOR = FuncionesClass.G_SERVIDORMASTER;
            formulario.DOC_ROOT_MYSQL = FuncionesClass.G_MYSQL_USER;
            formulario.DOC_PASSWORD_MYSQL = FuncionesClass.G_MYSQL_PASS;

            formulario.ShowDialog();

        }

        private void gvInforme_Click(object sender, EventArgs e)
        {

        }

        private void txtNumero_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtNumero_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ( (int)e.KeyChar == (int)Keys.Enter && txtNumero.Text != "") {
                txtNumero.Text = txtNumero.Text.PadLeft(10, Convert.ToChar("0"));
                txxCajaFolios.Focus();
            }

        }

        private void ddlTipoDocumento_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {

        }

        private void txxCajaFolios_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar == (int)Keys.Enter && txxCajaFolios.Text != "")
            {
                btnVer.Focus();
                btnVer.Enabled = true;
            }
        }

        private void txxCajaFolios_TextChanged(object sender, EventArgs e)
        {

        }
    }



}