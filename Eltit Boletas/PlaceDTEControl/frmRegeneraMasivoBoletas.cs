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
using Telerik.WinControls.UI;
using System.Xml;
using Eltit.Clases;
using System.util;
using PlaceSoft.Eltit.Class.clases;

namespace Eltit { 

public partial class frmRegeneraMasivoBoletas : Telerik.WinControls.UI.RadForm
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

        public frmRegeneraMasivoBoletas()
        {
            InitializeComponent();
        }

        private void frmRegeneraMasivoBoletas_Load(object sender, EventArgs e)
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
                if (fu.PingToHost(lblServidorVentas.Text) == true)
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
                    btnVer.Enabled = true;
                        txxCajaFolios.Focus();
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
            if(ddlEmpresas.SelectedIndex > 0 && ddLlocales.SelectedIndex > 0 )
            {
                if(txxCajaFolios.Text.Length != 2 )
                {
                    RadMessageBox.Show(this, "Debe Selecionar Una Caja de venta Válida. [" + txxCajaFolios.Text + "]", "Atencion", MessageBoxButtons.OK);
                }
                else
                {
                    GetDocumentosByCajaaLocalDesdeHata();
                }
            }
            else
            {
                RadMessageBox.Show(this, "Debe Selecionar Una Empresa y Local Válidos. [" + ddLlocales.Text + "]", "Atencion", MessageBoxButtons.OK);
            }
        }
        private void BuscaCaf(string xCaja, string xTipo)
        {
            Eltit.Clases.Caf myCaf = new Eltit.Clases.Caf(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            MySqlDataReader dr = myCaf.GetCafByCajaLocal(ddLlocales.Text.Substring(0, 2), xCaja, xTipo);
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
        private void GetDocumentosByCajaaLocalDesdeHata()
        {
            Documentos dc = new Documentos(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            MySqlDataReader dr = null;
            string base_datos = FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + ddLlocales.Text.Substring(0,2) ;
            object img = null;
            int count = 0;
            string tipo = "";
            string tipoFiscal = "";
            if(rbAfectas.CheckState == CheckState.Checked)
            {
                tipo = "BV";
                tipoFiscal = "39";
            }
            else
            {
                tipo = "BE";
                tipoFiscal = "41";
            }
            /**************     VERIFICA SI EL CAF EXISTE SI NO LOS TRAE **************/
            this.BuscaCaf(txxCajaFolios.Text, tipoFiscal);

            dr = dc.GetDocumentosCabezaByLocalNroInternoCajaDesdeHasta(ddLlocales.Text.Substring(0, 2), txxCajaFolios.Text, tipo, txtDesde.Text, 
                    txtHasta.Text, base_datos);
            string xml = "0";
            gvInforme.Rows.Clear();
            if(dr.HasRows == true)
            {
                ddlEmpresas.Enabled = false;
                ddLlocales.Enabled = false;
                btnGenera.Enabled = true;

                while(dr.Read())
                {
                    xml = dr["xml"].ToString();
                    if (xml != "0")
                    {
                        img = Properties.Resources.OK_48;
                    }
                    else
                    {
                        img = Properties.Resources.icons8_exclamacion;
                    }
                    gvInforme.Rows.Add(dr["tipo"].ToString(), dr["numero"].ToString(), dr["fecha"].ToString(), dr["caja"].ToString(), String.Format("{0:N0}", dr["total"]), img, false );
                    count++;
                }
            }
            dr.Close();
            dc.CerrarTransaccion();

            lblInfo.Text = "Total Registros " + count.ToString();
            SugiereUltimoCaf(tipoFiscal, txxCajaFolios.Text);

        }

        private void SugiereUltimoCaf( string xTipo, string xCaja)
        {
            PlaceSoft.Eltit.Class.clases.DTEClass dte = new PlaceSoft.Eltit.Class.clases.DTEClass(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            int folio = dte.GetUltimoFolioDTEByLocalCaja(ddLlocales.Text.Substring(0, 2), xTipo, xCaja, FuncionesClass.G_CLIENTE_PREFIJO + "fae" + ddLlocales.Text.Substring(0, 2));

            txtFolioGenera.Text = folio.ToString();
        }

        private void InicializaControlesDeEmpresa()
        {
           
              
        }

  

        // private void BuscaVentas(string xTipo)
        //{
        //    DTEClass idte = new DTEClass(this.cliente_sistema,   lblServidor.Text, this.mysql_root,this.mysql_pass);
            
        //    int inicial = 0;
        //    int final = 0;
        //    int count = 0;
        //    double neto = 0;
        //    double exento = 0;
        //    double iva = 0;
        //    double total =0;
        //    double totaldte = 0;
        //    double div = (FuncionesClass.G_IVA / 100) + 1;
        //    double cant39 = 0;
        //    double cant41 = 0;
        //     string nombredoc = "";
        //    string tipoFiscal = "";
        //    int anterior = 0;
        //    double diferencia = 0;

        //    int nulosInicial39 = 0;
        //    int nulosFinal39 = 0;
        //    int nulos39 = 0;

        //    int nulosInicial41 = 0;
        //    int nulosFinal41 = 0;
        //    int nulos41 = 0;

          
        //    int Usadoinicial39 = 0;
        //    int Usadofinal39 = 0;
        //    int Anuladoinicial39 = 0;
        //    int Anuladofinal39 = 0;
        //    int TotalUsados39 = 0;
        //    int TotalAnulados39 = 0;


        //    int Usadoinicial41 = 0;
        //    int Usadofinal41 = 0;
        //    int Anuladoinicial41 = 0;
        //    int Anuladofinal41 = 0;
        //    int TotalUsados41 = 0;
        //    int TotalAnulados41 = 0;



        //    //string Base_dte =  lblCliente.Text + "_dte_" + Convert.ToDouble(lblRut.Text.Substring(0, 9));
        //    //string base_ventas =  lblCliente.Text + "_local" + lblNombreEmpresa.Text.Substring(0, 2);

        //    string Base_dte =  lblCliente.Text + "_fae";
        //    string base_ventas = lblCliente.Text + "_ventas";


        //    MySqlDataReader dr = idte.GetBoletasByLocalDia(lblNombreEmpresa.Text.Substring(0, 2), Base_dte, base_ventas, FuncionesClass.GetFechaMysql(lblFecha.Text), xTipo, dt_locales);
        //    object img;
            
        //    if (dr.HasRows == true)
        //    {
        //        while(dr.Read())
        //        {

        //            if(count == 0)
        //            {
        //                inicial = Convert.ToInt32(dr["foliosii"].ToString());
        //                anterior = inicial;
        //            }
        //            else
        //            {
        //                diferencia = Convert.ToInt32(dr["foliosii"].ToString()) - anterior;
        //                anterior = Convert.ToInt32(dr["foliosii"].ToString());

        //            }  

        //            if (dr["tipo_doc"].ToString() == "39")
        //            {
        //                cant39 = cant39 + 1;
        //                tipoFiscal = "39";
        //                //if(Convert.ToDouble(dr["monto_total"].ToString()) > 0 && diferencia == 1 || Convert.ToDouble(dr["monto_total"].ToString()) > 0 && diferencia == 0)
        //                if ( diferencia == 1 )

        //                {
        //                    if(Usadoinicial39 == 0) //SI EL USADO FINAL NO CORRESPONDE A LSIGUENTE CORRELATIVO SALTRSE AL CORRELATIVO FINAL.
        //                    {
        //                        if(Anuladoinicial39 !=0)
        //                        {
        //                            //rangoAnulado39.Add(Anuladoinicial39 + "," + Anuladofinal39);
        //                            Anuladoinicial39 = 0;
        //                            Anuladofinal39 = 0;
        //                        }
                               

        //                        Usadoinicial39 = Convert.ToInt32(dr["foliosii"].ToString());
        //                        Usadofinal39   =  Convert.ToInt32(dr["foliosii"].ToString());
        //                    }
        //                    else
        //                    {
                                
        //                        Usadofinal39 = Convert.ToInt32(dr["foliosii"].ToString());

        //                    }
        //                    TotalUsados39++;
                          
        //                }
        //                else
        //                {
        //                    if (Anuladoinicial39 == 0 || diferencia > 1)
        //                    {
        //                        if(Usadoinicial39 !=0)
        //                        {
                                    
                                   
        //                            rangoUtilizado39.Add(Usadoinicial39 + "," + Usadofinal39);
        //                            Usadoinicial39 = Convert.ToInt32(dr["foliosii"].ToString());
        //                            //Usadofinal39 = 0;
        //                            Usadofinal39 = Convert.ToInt32(dr["foliosii"].ToString()); 
        //                        }
                               
        //                        if(Convert.ToDouble(dr["monto_total"].ToString()) == 0)
        //                        {
        //                            //Anuladoinicial39 = Convert.ToInt32(dr["foliosii"].ToString());
        //                            //Anuladofinal39 = Convert.ToInt32(dr["foliosii"].ToString());
        //                        }
                                
        //                    }
        //                    else
        //                    {
        //                        //Anuladofinal39 = Convert.ToInt32(dr["foliosii"].ToString());
        //                        TotalAnulados39++;
        //                    }

        //                    TotalUsados39++;

        //                }
        //            }


        //            if (dr["tipo_doc"].ToString() == "41")
        //            {
        //                cant41 = cant41 + 1;
        //                tipoFiscal = "41";
        //                if (Convert.ToDouble(dr["monto_total"].ToString()) > 0 && diferencia == 1)
        //                {
        //                    if (Usadoinicial41 == 0) //SI EL USADO FINAL NO CORRESPONDE A LSIGUENTE CORRELATIVO SALTRSE AL CORRELATIVO FINAL.
        //                    {
        //                        if (Anuladoinicial41 != 0)
        //                        {
        //                            rangoAnulado41.Add(Anuladoinicial41 + "," + Anuladofinal41);
        //                            Anuladoinicial41 = 0;
        //                            Anuladofinal41 = 0;
        //                        }


        //                        Usadoinicial41 = Convert.ToInt32(dr["foliosii"].ToString());
        //                        Usadofinal41 = Convert.ToInt32(dr["foliosii"].ToString());
        //                    }
        //                    else
        //                    {

        //                        Usadofinal41 = Convert.ToInt32(dr["foliosii"].ToString());

        //                    }
        //                    TotalUsados41++;

        //                }
        //                else
        //                {
        //                    if (Anuladoinicial41 == 0 || diferencia > 1)
        //                    {
        //                        if (Usadoinicial41 != 0)
        //                        {


        //                            rangoUtilizado41.Add(Usadoinicial41 + "," + Usadofinal41);
        //                            Usadoinicial41 = Convert.ToInt32(dr["foliosii"].ToString());
        //                            //Usadofinal39 = 0;
        //                            Usadofinal41 = Convert.ToInt32(dr["foliosii"].ToString()); ;
        //                        }

        //                        if (Convert.ToDouble(dr["monto_total"].ToString()) == 0)
        //                        {
        //                            Anuladoinicial41 = Convert.ToInt32(dr["foliosii"].ToString());
        //                            Anuladofinal41 = Convert.ToInt32(dr["foliosii"].ToString());
        //                        }

        //                    }
        //                    else
        //                    {
        //                        Anuladofinal41 = Convert.ToInt32(dr["foliosii"].ToString());
        //                        TotalAnulados41++;
        //                    }

        //                    TotalUsados41++;

        //                }



        //            }// end if 41


        //            nombredoc = FuncionesClass.getNombredocumentoByCodigo(dr["tipo_doc"].ToString());
        //            string xml = dr["fae_xml"].ToString();
        //            XML_DETALLE += "[ TD"+ dr["tipo_doc"].ToString() + "-FF" + dr["foliosii"].ToString() + "-FE"+ dr["fecha_emision"].ToString() + "-TO"+ dr["monto_total"] + "]" + Environment.NewLine;

        //            totaldte = 0;
        //            if (xml != "")
        //            {
        //                /**************** admin erp no firma electronicamente las boletas '''''''
        //                 por lo que solo se sacara el monto del campo total sin serializar
        //                 */
        //                //var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);                       
        //                //totaldte = dte.Documento.Encabezado.Totales.MontoTotal;
                        
        //                var dte = ""; //PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);
        //                xml = "<?xml version='1.0' encoding='iso - 8859 - 1'?>" + xml;
        //                xml = xml.Replace("Descuentopct", "DescuentoPct");
        //                totaldte =  this.TotalDoc(xml);

        //            }

        //            double totalDoc = Convert.ToDouble(dr["monto_total"]);
        //            if ( totalDoc == totaldte)
        //            {
        //                img = Resources.OK_48;
        //            }
        //            else
        //            {
        //                img = Resources.icons8_exclamacion;
        //                gvDiferencias.Rows.Add(dr["localdocumento"].ToString(), "Folio ["+ dr["foliosii"].ToString() + "] Doc.Cabeza[" + totalDoc + "] <> XML[" + totaldte + "]");

        //            }

        //           if(muestraDetalles == true)
        //           {
        //            gvInforme.Rows.Add(tipoFiscal + " " + dr["localdocumento"].ToString(), dr["foliosii"].ToString(),
        //                               dr["fecha_emision"].ToString(), dr["caja_doc"].ToString(),
        //                               String.Format("{0:N0}", dr["monto_total"]), img);

        //            if (diferencia != 1 && count != 0)
        //            {
        //                FuncionesClass fun = new FuncionesClass();
        //                fun.ColoreaCelda(gvInforme.Rows[gvInforme.CurrentRow.Index].Cells[0], Color.Yellow);
        //            }

        //        }
                  
                   

                  

        //            exento = exento + Convert.ToDouble(dr["monto_exento"]);
        //            total = total + Convert.ToDouble(dr["monto_total"]);


        //            count++;
        //            final = Convert.ToInt32(dr["foliosii"].ToString());



        //        } // FIN WHILE
        //    } // FIN IF


        //    /*************** GENERA RANGOS DE FOLIOS CONSECUTIVOS ****************/
        //    if(Usadoinicial39 > 0)
        //    {
        //        rangoUtilizado39.Add(Usadoinicial39 + "," + Usadofinal39);
        //    }

        //    if(Anuladoinicial39 > 0)
        //    {
        //        rangoAnulado39.Add(Anuladoinicial39 + "," + Anuladofinal39);
        //    }
        //    /*************** GENERA RANGOS DE FOLIOS CONSECUTIVOS ****************/

        //    if (Usadoinicial41 > 0)
        //    {
        //        rangoUtilizado41.Add(Usadoinicial41 + "," + Usadofinal41);
        //    }

        //    if (Anuladoinicial41 > 0)
        //    {
        //        rangoAnulado41.Add(Anuladoinicial41 + "," + Anuladofinal41);
        //    }

        ///*****MUESTRA RANGOS *******/
        //  if (xTipo == "39")
        //{
        //    foreach (string var in rangoUtilizado39)
        //    {
        //        txtXML.Text += "Util 39: " + var.ToString() + Environment.NewLine;
        //    }
        //    foreach (string var2 in rangoAnulado39)
        //    {
        //        txtXML.Text += "NULO 39: " + var2.ToString() + Environment.NewLine;
        //    }
        //}
        //  if (xTipo == "41")
        //{
        //    foreach (string var in rangoUtilizado41)
        //    {
        //        txtXML.Text += "Util 41: " + var.ToString() + Environment.NewLine;
        //    }
        //    foreach (string var2 in rangoAnulado41)
        //    {
        //        txtXML.Text += "Util 41: " + var2.ToString() + Environment.NewLine;
        //    }
        //}

        ///******* GENERA TOTALES ****/
        //  if (xTipo == "39")
        //    {
        //        neto = total / div;
        //        neto = Math.Round(neto, 4);
        //        iva = total - neto;

        //        neto = Math.Round(neto, 0, MidpointRounding.AwayFromZero);
        //        iva = Math.Round(iva, 0, MidpointRounding.AwayFromZero);
        //        total = Math.Round(neto + 0 + iva, 0, MidpointRounding.AwayFromZero);

        //        lblDesde39.Text = inicial.ToString();
        //        lblHasta39.Text = final.ToString();

        //        lblNeto39.Text = String.Format("{0:N0}", neto);
        //        lblExento39.Text = String.Format("{0:N0}", 0);
        //        lblIva39.Text = String.Format("{0:N0}", iva);
        //        lbltotal39.Text = String.Format("{0:N0}", total);
              

        //        lblEmitidos39.Text = TotalUsados39.ToString();
        //        lblAnulado39.Text = TotalAnulados39.ToString();
        //        lblUtilizados39.Text = (TotalUsados39 + TotalAnulados39).ToString();
               
        //    }
        //  if (xTipo == "41")
        //    {                
        //        neto = 0;
        //        iva = 0;

        //        neto = Math.Round(neto, 0, MidpointRounding.AwayFromZero);
        //        iva = Math.Round(iva, 0, MidpointRounding.AwayFromZero);
        //        total = Math.Round(total, MidpointRounding.AwayFromZero);

        //        lbldesde41.Text = inicial.ToString();
        //        lblHasta41.Text = final.ToString();

        //        lblNeto41.Text = String.Format("{0:N0}", neto);
        //        lblExento41.Text = String.Format("{0:N0}", total);
        //        lblIva41.Text = String.Format("{0:N0}", iva);
        //        lblTotal41.Text = String.Format("{0:N0}", total);


        //        lblEmitidos41.Text = TotalUsados41.ToString();
        //        lblAnulados41.Text = TotalAnulados41.ToString();
        //        lblUtilizados41.Text = (TotalUsados41 + TotalAnulados41).ToString();

        //    }
           
         
               

        //    dr.Close();
        //    idte.CerrarTransaccion();
          
        //}


        private void btnGenera_Click(object sender, EventArgs e)
        {
            if(gvInforme.Rows.Count > 0)
            {
                if(RadMessageBox.Show(this, "Atención, va a eliminar XML y generar unos nuevos con los mismos datos ¿Desea Realmente Hacerlo?", "Atencion",MessageBoxButtons.YesNo) == DialogResult.Yes)
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
            for (i=0; i <= gvInforme.Rows.Count-1;i++ )
            {
                check =  Convert.ToBoolean(gvInforme.Rows[i].Cells[6].Value);
                if(check == true)
                {
                    tipo = gvInforme.Rows[i].Cells[0].Value.ToString();
                    numero = gvInforme.Rows[i].Cells[1].Value.ToString();
                    fecha = gvInforme.Rows[i].Cells[2].Value.ToString();
                    caja = gvInforme.Rows[i].Cells[3].Value.ToString();
                    MONTO = Convert.ToDouble(gvInforme.Rows[i].Cells[4].Value.ToString());
                    this.GenerarDocumentoBoleta(tipo, numero, fecha, caja, "66666666-6", "PARTICULAR",MONTO);
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
            }



        }

        private void GenerarDocumentoBoleta(string xTipoInterno, string xNumeroInterno, string xFecha , string xCaja, string xRut, string xNombre,
                                            double xMonto)
        {
            Handler handler;
            handler = new Handler();
            Documentos dc = new Documentos(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            MySqlDataReader dr = null;
            string xtipo = "";
            string base_venta = FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + ddLlocales.Text.Substring(0,2);
            double precio = 0;
            double cantidad = 0;
            double dcto = 0;
            double DctoLinea = 0;
            int totalLinea = 0;
            string codigoArticulo = "";
            string rut_venta = "";
            List<ItemBoleta> items;
            items = new List<ItemBoleta>();

            string tipoFiscal = "";

            PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronica;

            if (rbAfectas.CheckState == CheckState.Checked)
            {
                xtipo = "BV";
                tipoFiscal = "39";
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica;
            } else {
                xtipo = "BE";
                tipoFiscal = "41";
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta;
            }

            xFecha = FuncionesClass.GetFechaMysql(xFecha);
            dr = dc.GetDoucumentosDetalleByTipoCajaNroOnternoFechaLocal(ddLlocales.Text.Substring(0,2), xTipoInterno, xNumeroInterno, xCaja, xFecha, base_venta);

            if(dr.HasRows == true)
            {
               
                while (dr.Read())
                {
                    rut_venta = dr["rut"].ToString();
                    xFecha = FuncionesClass.GetFechaMysql(dr["fecha"].ToString());
                    //formaPago = dr["formapago"].ToString();
                    ItemBoleta item = new ItemBoleta();
                    codigoArticulo = dr["codigo"].ToString();
                    precio = Convert.ToInt32(dr["precio"].ToString());
                    cantidad = Convert.ToDouble(dr["cantidad"].ToString());
                    dcto = Convert.ToDouble(dr["descuentopesos"]);
                    DctoLinea = dcto; 
                    totalLinea = Convert.ToInt32(Math.Round(cantidad * precio, 4) - DctoLinea);

                    item.Nombre = dr["descripcion"].ToString();
                    item.Cantidad = cantidad;
                    item.Codigo = codigoArticulo;
                    //if (dr["impuesto"].ToString() == "00008")
                    //{
                    //    item.Afecto = false;
                    //}
                    //else
                    //{
                    //    item.Afecto = true;
                    //}
                    
                    item.Afecto = true;
                    item.Precio = Convert.ToInt32(Math.Round(precio, 4));
                    item.Porce_Descuento = Convert.ToInt32(dr["descuento"].ToString());
                    item.Monto_Descuento = Convert.ToInt32(dcto);
                    
                    item.Total = totalLinea;
                    item.UnidadMedida = string.Empty;
                    items.Add(item);
                }// END WHILE

                /********** GET ULTIMO FOLIO FISCAL DTE ****************/
                PlaceSoft.Eltit.Class.clases.DTEClass DTE = new PlaceSoft.Eltit.Class.clases.DTEClass(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
                int FOLIO_CAF = DTE.GetUltimoFolioDTEByLocalCaja(ddLlocales.Text.Substring(0, 2), tipoFiscal, xCaja, base_venta);

                if(txtFolioGenera.Text != "" )
                {
                    FOLIO_CAF =  Convert.ToInt32(txtFolioGenera.Text) + countGenerados;
                }

                FOLIO_CAF = GetUltimoFolio(tipoFiscal, FOLIO_CAF, xCaja);

                handler.tipo = typeDTE;
                handler.casoPruebas = string.Empty;
                handler.Folio = (int)FOLIO_CAF;
                handler.idDte = "ENVIOFOLIO_"+ FOLIO_CAF +"T"+ xtipo; // "DTE" + DOC_TIPO + "F" + FOLIO_CAF;
                handler.rutcliente = Convert.ToDouble(rut_venta.Substring(0,9)) + "-" + rut_venta.Substring(9, 1);

                handler.emisor_rut = Convert.ToDouble( lblRutEmpresa.Text.Substring(0, 9)) + "-" + lblRutEmpresa.Text.Substring(9, 1);
                handler.emisor_razon_social = ddlEmpresas.Text.Substring(2, ddlEmpresas.Text.Length - 2);
                handler.emisor_giro = lblGiro.Text;
                handler.emisor_comuna = lblComuna.Text;
                handler.emisor_ciudad = lblCiudad.Text;
                handler.emisor_direccion = lblCiudad.Text;
                handler.cod_sucursal_sii = lblCodigoSucursal.Text;
                handler.fechaEmision = Convert.ToDateTime(xFecha);
                var dte = handler.GenerateDTEBoleta();

                handler.GenerateDetails(dte, items);
                handler.ReferenciasBoleta(dte);

                var path = handler.TimbrarYFirmarXMLDTE(dte,  @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + Convert.ToDouble(lblRutEmpresa.Text.Substring(0, 9)) + @"\Produccion\Caf\" + ddLlocales.Text.Substring(0, 2) + @"\");
                if (File.Exists(path))
                {
                    FileInfo fi = new FileInfo(path);
                    //string destino = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\xml\" + FuncionesClass.G_LOCAL + @"\DTE39F" + FOLIO_CAF + ".xml";
                    string destino = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + Convert.ToDouble(lblRutEmpresa.Text.Substring(0,9)) +  @"\Produccion\xml\" + ddLlocales.Text.Substring(0,2) + @"\DTE" + tipoFiscal + "F" + FOLIO_CAF + ".xml";
                    fi.CopyTo(destino, true);
                    string XML = File.ReadAllText(destino, Encoding.GetEncoding("ISO-8859-1"));

                    this.GrabaXML(ddLlocales.Text.Substring(0, 2), xTipoInterno, xNumeroInterno, tipoFiscal, FOLIO_CAF,
                                     xCaja, XML, xFecha, xNombre, xRut, xMonto);

                }


            }

            dr.Close();
            dc.CerrarTransaccion();


        }

        private void GrabaXML(string xLocal, string xTipoInterno,string xNroInterno, string xTipoFiscal, int xFolioFiscal,
                 string xCaja, string XML, string xFecha, string xNombre, string xRut, double xMonto)
        {
            string basedte = FuncionesClass.G_CLIENTE_PREFIJO + "fae" + ddLlocales.Text.Substring(0, 2);
            PlaceSoft.Eltit.Class.clases.DTEClass dte = new PlaceSoft.Eltit.Class.clases.DTEClass(lblServidorVentas.Text, lblRoot.Text,lblPassword.Text);
            dte.GrabaXML(xLocal, xTipoFiscal, xFolioFiscal, xTipoInterno, xNroInterno, xFecha, basedte, XML, xCaja, xRut, xNombre, xMonto);

            ActualizaFolioSII(xNroInterno, xCaja, xTipoInterno, xFecha, xFolioFiscal.ToString().PadLeft(10, Convert.ToChar("0")));
        }
        private void ActualizaFolioSII(string xnrointerno, string xCaja, string xTipointerno,string xFecha ,string xFolioSII)
        {
            string BASE_VENTA = FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + ddLlocales.Text.Substring(0, 2);
            Documentos dc = new Documentos(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            dc.ActualizaFolioSII(ddLlocales.Text.Substring(0, 2), BASE_VENTA, xnrointerno, xTipointerno, xFecha, xCaja, xFolioSII);
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
                ultimo = caf.BuscaRangoSiguiente(ddLlocales.Text.Substring(0, 2),xtipo, ultimo.ToString(),xCaja, base_dte);
            }
                       
            return ultimo;
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

                gvInforme.Rows[i].Cells[6].Value = val;
            }

        }

    }



}