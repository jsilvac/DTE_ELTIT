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
using Eltit;

public partial class frmPopEnviaRCOF : Telerik.WinControls.UI.RadForm
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

        List<string> rangoUtilizado39 = new List<string>();
        List<string> rangoAnulado39 = new List<string>();

        List<string> rangoUtilizado41 = new List<string>();
        List<string> rangoAnulado41 = new List<string>();

        List<string> rangoUtilizado61 = new List<string>();
        List<string> rangoAnulado61 = new List<string>();

        public bool muestraDetalles;
        public DataTable dt_locales;

        public frmPopEnviaRCOF()
        {
            InitializeComponent();
        }

        private void frmPopEnviaRCOF_Load(object sender, EventArgs e)
        {
            FuncionesClass fu = new FuncionesClass();

            if (fu.PingToHost(lblServidor.Text) == false)
            {
                MessageBox.Show("No se pudo establecer conección con el host Remoto[Host:" + lblServidor.Text + "]");
                this.Close();
            }

            try
            {
                handler.rutEmpresa = lblRut.Text;
                BASE_DTE = lblCliente.Text + "_dte_" + Convert.ToDouble(lblRut.Text.Substring(0, 9));
                BASE_VENTAS = lblCliente.Text + "_local" + lblNombreEmpresa.Text.Substring(0, 2);
                cliente_sistema = lblCliente.Text + "";
                this.BuscaVentas("39");
                //this.BuscaVentas("41");
                //this.BuscaNCBoletas();

                //VerificaRcof(FuncionesClass.GetFechaMysql(lblFecha.Text));
                chbProduccion.Checked = true;

                //this.BuscaVentasCliente();
                gvLocales.TableElement.Font = new Font("Arial", 6);
                gvLocales.TableElement.RowHeight = 18;

                gvDiferencias.TableElement.Font = new Font("Arial", 6);
                gvDiferencias.TableElement.RowHeight = 18;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error:[" + ex.Message.ToString() + "]");
            }
          


        }
        public void CargaLocales(RadGridView gv)
        {
            int i = 0;

            dt_locales = new DataTable();           
            dt_locales.Columns.Add("codigo", typeof(string));
            dt_locales.Columns.Add("nombre", typeof(string));
            DataRow row = dt_locales.NewRow();

            for (i=0; i <= gv.Rows.Count - 1;i++)
            {
                gvLocales.Rows.Add(gv.Rows[i].Cells[0].Value.ToString(), gv.Rows[i].Cells[1].Value.ToString());
                dt_locales.Rows.Add(gv.Rows[i].Cells[0].Value.ToString(), gv.Rows[i].Cells[1].Value.ToString());
            }

        }
        
        private void InicializaControlesDeEmpresa()
        {
            lblRut.Text = FuncionesClass.G_EMPRESARUT;
            lblNombreEmpresa.Text =   FuncionesClass.G_EMPRESANOMBRE;
              
        }

        private void BuscaNCBoletas()
        {
        PlaceSoft.Eltit.Class.clases.DTEClass ventas = new PlaceSoft.Eltit.Class.clases.DTEClass(lblServidor.Text, this.mysql_root, this.mysql_pass);
           // ventas.setBaseDTE(this.BASE_DTE);
            int inicial = 0;
            int final = 0;
            int count = 0;
            double neto = 0;
            double exento = 0;
            double iva = 0;
            double total = 0;
            double totaldte = 0;
            double div = (FuncionesClass.G_IVA / 100) + 1;
            double cant = 0;   
            double cantnulas = 0;
            string nombredoc = "";
            string tipoFiscal = "";


            MySqlDataReader dr = ventas.GetNotasCreditoBoletaByLocalDia(lblNombreEmpresa.Text.Substring(0, 2), FuncionesClass.GetFechaMysql(lblFecha.Text), this.BASE_VENTAS,  this.BASE_DTE);
            object img;

            if (dr.HasRows == true)
            {
                while (dr.Read())
                {

                    if (count == 0)
                    {
                        inicial = Convert.ToInt32(dr["foliosii"].ToString());
                    }

                    if (dr["tipo_doc"].ToString() == "NFE")
                    {
                        cant = cant + 1;
                        tipoFiscal = "61";
                    }


                    nombredoc = FuncionesClass.getNombredocumentoByCodigo(dr["tipo_doc"].ToString());
                    string xml = dr["fae_xml"].ToString();
                    txtXML.Text += xml + Environment.NewLine;
                    XML_DETALLE += "[ TD" + dr["tipo_doc"].ToString() + "-FF" + dr["foliosii"].ToString() + "-FE" + dr["fecha_emision"].ToString() + "-TO" + dr["monto_total"] + "]" + Environment.NewLine;

                    var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);
                    totaldte = dte.Documento.Encabezado.Totales.MontoTotal;
                    dtes.Add(dte);
                    if (Convert.ToDouble(dr["monto_total"]) == totaldte)
                    {
                        img = Eltit.Properties.Resources.OK_48;
                    }
                    else
                    {
                        img = Eltit.Properties.Resources.icons8_exclamacion;
                    }

                    gvInforme.Rows.Add(tipoFiscal + " " + nombredoc, dr["foliosii"].ToString(),
                                        dr["fecha_emision"].ToString(), dr["caja_doc"].ToString(),
                                        String.Format("{0:N0}", dr["monto_total"]), img);

                    exento = exento + Convert.ToDouble(dr["monto_exento"]);
                    total = total + Convert.ToDouble(dr["monto_total"]);


                    count++;
                    final = Convert.ToInt32(dr["foliosii"].ToString());
                } // FIN WHILE
            } // FIN IF

            neto = total / div;
            neto = Math.Round(neto, 4);
            iva = total - neto;

            neto = Math.Round(neto, 0, MidpointRounding.AwayFromZero);
            iva = Math.Round(iva, 0, MidpointRounding.AwayFromZero);
            total = Math.Round(neto + exento + iva, 0, MidpointRounding.AwayFromZero);

            lblNeto61.Text = String.Format("{0:N0}", neto);
            lblExento61.Text = String.Format("{0:N0}", exento);
            lblIva61.Text = String.Format("{0:N0}", iva);
            lblTotal61.Text = String.Format("{0:N0}", total);
            lblCantidad61.Text = cant.ToString();
            lblEmitidos61.Text = count.ToString();
            lblAnulados61.Text = cantnulas.ToString();

            lblInicial61.Text = inicial.ToString();
            lblFinal61.Text = final.ToString();


        }

         private void BuscaVentas(string xTipo)
        {
        PlaceSoft.Eltit.Class.clases.DTEClass idte = new PlaceSoft.Eltit.Class.clases.DTEClass(lblServidor.Text, this.mysql_root,this.mysql_pass);
            
            int inicial = 0;
            int final = 0;
            int count = 0;
            double neto = 0;
            double exento = 0;
            double iva = 0;
            double total =0;
            double totaldte = 0;
            double div = (FuncionesClass.G_IVA / 100) + 1;
            double cant39 = 0;
            double cant41 = 0;
             string nombredoc = "";
            string tipoFiscal = "";
            int anterior = 0;
            double diferencia = 0;

            int nulosInicial39 = 0;
            int nulosFinal39 = 0;
            int nulos39 = 0;

            int nulosInicial41 = 0;
            int nulosFinal41 = 0;
            int nulos41 = 0;

          
            int Usadoinicial39 = 0;
            int Usadofinal39 = 0;
            int Anuladoinicial39 = 0;
            int Anuladofinal39 = 0;
            int TotalUsados39 = 0;
            int TotalAnulados39 = 0;

        string TipoInterno = "";
            int Usadoinicial41 = 0;
            int Usadofinal41 = 0;
            int Anuladoinicial41 = 0;
            int Anuladofinal41 = 0;
            int TotalUsados41 = 0;
            int TotalAnulados41 = 0;

           if(xTipo == "39")
            {
               TipoInterno = "BV";
            }
            if (xTipo == "41")
            {
                TipoInterno = "BE";
            }
            if (xTipo == "61")
            {
                TipoInterno = "NF";
            }

            string Base_dte =  lblCliente.Text + "_fae";
            string base_ventas = lblCliente.Text + "_ventas";


            MySqlDataReader dr = idte.GetBoletasByLocalDia(lblNombreEmpresa.Text.Substring(0, 2), Base_dte, base_ventas, FuncionesClass.GetFechaMysql(lblFecha.Text), xTipo,TipoInterno, dt_locales);
            object img;


        try
        {
            if (dr.HasRows == true)
            {
                while (dr.Read())
                {

                    if (count == 0)
                    {
                        inicial = Convert.ToInt32(dr["foliosii"].ToString());
                        anterior = inicial;
                    }
                    else
                    {
                        diferencia = Convert.ToInt32(dr["foliosii"].ToString()) - anterior;
                        anterior = Convert.ToInt32(dr["foliosii"].ToString());

                    }

                    if (dr["tipo_doc"].ToString() == "39")
                    {
                        cant39 = cant39 + 1;
                        tipoFiscal = "39";
                        //if(Convert.ToDouble(dr["monto_total"].ToString()) > 0 && diferencia == 1 || Convert.ToDouble(dr["monto_total"].ToString()) > 0 && diferencia == 0)
                        if (diferencia == 1)

                        {
                            if (Usadoinicial39 == 0) //SI EL USADO FINAL NO CORRESPONDE A LSIGUENTE CORRELATIVO SALTRSE AL CORRELATIVO FINAL.
                            {
                                if (Anuladoinicial39 != 0)
                                {
                                    //rangoAnulado39.Add(Anuladoinicial39 + "," + Anuladofinal39);
                                    Anuladoinicial39 = 0;
                                    Anuladofinal39 = 0;
                                }


                                Usadoinicial39 = Convert.ToInt32(dr["foliosii"].ToString());
                                Usadofinal39 = Convert.ToInt32(dr["foliosii"].ToString());
                            }
                            else
                            {

                                Usadofinal39 = Convert.ToInt32(dr["foliosii"].ToString());

                            }
                            TotalUsados39++;

                        }
                        else
                        {
                            if (Anuladoinicial39 == 0 || diferencia > 1)
                            {
                                if (Usadoinicial39 != 0)
                                {


                                    rangoUtilizado39.Add(Usadoinicial39 + "," + Usadofinal39);
                                    Usadoinicial39 = Convert.ToInt32(dr["foliosii"].ToString());
                                    //Usadofinal39 = 0;
                                    Usadofinal39 = Convert.ToInt32(dr["foliosii"].ToString());
                                }

                                if (Convert.ToDouble(dr["monto_total"].ToString()) == 0)
                                {
                                    //Anuladoinicial39 = Convert.ToInt32(dr["foliosii"].ToString());
                                    //Anuladofinal39 = Convert.ToInt32(dr["foliosii"].ToString());
                                }

                            }
                            else
                            {
                                //Anuladofinal39 = Convert.ToInt32(dr["foliosii"].ToString());
                                TotalAnulados39++;
                            }

                            TotalUsados39++;

                        }
                    }


                    if (dr["tipo_doc"].ToString() == "41")
                    {
                        cant41 = cant41 + 1;
                        tipoFiscal = "41";
                        if (Convert.ToDouble(dr["monto_total"].ToString()) > 0 && diferencia == 1)
                        {
                            if (Usadoinicial41 == 0) //SI EL USADO FINAL NO CORRESPONDE A LSIGUENTE CORRELATIVO SALTRSE AL CORRELATIVO FINAL.
                            {
                                if (Anuladoinicial41 != 0)
                                {
                                    rangoAnulado41.Add(Anuladoinicial41 + "," + Anuladofinal41);
                                    Anuladoinicial41 = 0;
                                    Anuladofinal41 = 0;
                                }


                                Usadoinicial41 = Convert.ToInt32(dr["foliosii"].ToString());
                                Usadofinal41 = Convert.ToInt32(dr["foliosii"].ToString());
                            }
                            else
                            {

                                Usadofinal41 = Convert.ToInt32(dr["foliosii"].ToString());

                            }
                            TotalUsados41++;

                        }
                        else
                        {
                            if (Anuladoinicial41 == 0 || diferencia > 1)
                            {
                                if (Usadoinicial41 != 0)
                                {


                                    rangoUtilizado41.Add(Usadoinicial41 + "," + Usadofinal41);
                                    Usadoinicial41 = Convert.ToInt32(dr["foliosii"].ToString());
                                    //Usadofinal39 = 0;
                                    Usadofinal41 = Convert.ToInt32(dr["foliosii"].ToString()); ;
                                }

                                if (Convert.ToDouble(dr["monto_total"].ToString()) == 0)
                                {
                                    Anuladoinicial41 = Convert.ToInt32(dr["foliosii"].ToString());
                                    Anuladofinal41 = Convert.ToInt32(dr["foliosii"].ToString());
                                }

                            }
                            else
                            {
                                Anuladofinal41 = Convert.ToInt32(dr["foliosii"].ToString());
                                TotalAnulados41++;
                            }

                            TotalUsados41++;

                        }



                    }// end if 41


                    nombredoc = FuncionesClass.getNombredocumentoByCodigo(dr["tipo_doc"].ToString());
                    string xml = dr["fae_xml"].ToString();
                    XML_DETALLE += "[ TD" + dr["tipo_doc"].ToString() + "-FF" + dr["foliosii"].ToString() + "-FE" + dr["fecha_emision"].ToString() + "-TO" + dr["monto_total"] + "]" + Environment.NewLine;

                    totaldte = 0;
                    if (xml != "")
                    {
                        /**************** admin erp no firma electronicamente las boletas '''''''
                         por lo que solo se sacara el monto del campo total sin serializar
                         */
                        //var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);                       
                        //totaldte = dte.Documento.Encabezado.Totales.MontoTotal;

                        var dte = ""; //PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);

                        if (xml.Contains("<?xml version"))
                        {

                        }
                        else
                        {
                            xml = "<?xml version='1.0' encoding='iso - 8859 - 1'?>" + xml;

                        }

                        xml = xml.Replace("Descuentopct", "DescuentoPct");
                        totaldte = this.TotalDoc(xml);

                    }

                    double totalDoc = Convert.ToDouble(dr["monto_total"]);
                    if (totalDoc == totaldte)
                    {
                        img = Eltit.Properties.Resources.OK_48;
                    }
                    else
                    {
                        img = Eltit.Properties.Resources.icons8_exclamacion;
                        gvDiferencias.Rows.Add(dr["localdocumento"].ToString(), "Folio [" + dr["foliosii"].ToString() + "] Doc.Cabeza[" + totalDoc + "] <> XML[" + totaldte + "]");

                    }

                    if (muestraDetalles == true)
                    {
                        gvInforme.Rows.Add(tipoFiscal + " " + dr["localdocumento"].ToString(), dr["foliosii"].ToString(),
                                           dr["fecha_emision"].ToString(), dr["caja_doc"].ToString(),
                                           String.Format("{0:N0}", dr["monto_total"]), img);

                        if (diferencia != 1 && count != 0)
                        {
                            FuncionesClass fun = new FuncionesClass();
                            fun.ColoreaCelda(gvInforme.Rows[gvInforme.CurrentRow.Index].Cells[0], Color.Yellow);
                        }

                    }





                    exento = exento + Convert.ToDouble(dr["monto_exento"]);
                    total = total + Convert.ToDouble(dr["monto_total"]);


                    count++;
                    final = Convert.ToInt32(dr["foliosii"].ToString());



                } // FIN WHILE
            } // FIN IF
        }
        catch(Exception ex)
        {
            MessageBox.Show(ex.Message.ToString());
        }

        


            /*************** GENERA RANGOS DE FOLIOS CONSECUTIVOS ****************/
            if(Usadoinicial39 > 0)
            {
                rangoUtilizado39.Add(Usadoinicial39 + "," + Usadofinal39);
            }

            if(Anuladoinicial39 > 0)
            {
                rangoAnulado39.Add(Anuladoinicial39 + "," + Anuladofinal39);
            }
            /*************** GENERA RANGOS DE FOLIOS CONSECUTIVOS ****************/

            if (Usadoinicial41 > 0)
            {
                rangoUtilizado41.Add(Usadoinicial41 + "," + Usadofinal41);
            }

            if (Anuladoinicial41 > 0)
            {
                rangoAnulado41.Add(Anuladoinicial41 + "," + Anuladofinal41);
            }

        /*****MUESTRA RANGOS *******/
          if (xTipo == "39")
        {
            foreach (string var in rangoUtilizado39)
            {
                txtXML.Text += "Util 39: " + var.ToString() + Environment.NewLine;
            }
            foreach (string var2 in rangoAnulado39)
            {
                txtXML.Text += "NULO 39: " + var2.ToString() + Environment.NewLine;
            }
        }
          if (xTipo == "41")
        {
            foreach (string var in rangoUtilizado41)
            {
                txtXML.Text += "Util 41: " + var.ToString() + Environment.NewLine;
            }
            foreach (string var2 in rangoAnulado41)
            {
                txtXML.Text += "Util 41: " + var2.ToString() + Environment.NewLine;
            }
        }

        /******* GENERA TOTALES ****/
          if (xTipo == "39")
            {
                neto = total / div;
                neto = Math.Round(neto, 4);
                iva = total - neto;

                neto = Math.Round(neto, 0, MidpointRounding.AwayFromZero);
                iva = Math.Round(iva, 0, MidpointRounding.AwayFromZero);
               // total = Math.Round(neto + 0 + iva, 0, MidpointRounding.AwayFromZero);

                lblDesde39.Text = inicial.ToString();
                lblHasta39.Text = final.ToString();

                lblNeto39.Text = String.Format("{0:N0}", neto);
                lblExento39.Text = String.Format("{0:N0}", 0);
                lblIva39.Text = String.Format("{0:N0}", iva);
                lbltotal39.Text = String.Format("{0:N0}", total);
              

                lblEmitidos39.Text = TotalUsados39.ToString();
                lblAnulado39.Text = TotalAnulados39.ToString();
                lblUtilizados39.Text = (TotalUsados39 + TotalAnulados39).ToString();
               
            }
          if (xTipo == "41")
            {                
                neto = 0;
                iva = 0;

                neto = Math.Round(neto, 0, MidpointRounding.AwayFromZero);
                iva = Math.Round(iva, 0, MidpointRounding.AwayFromZero);
                total = Math.Round(total, MidpointRounding.AwayFromZero);

                lbldesde41.Text = inicial.ToString();
                lblHasta41.Text = final.ToString();

                lblNeto41.Text = String.Format("{0:N0}", neto);
                lblExento41.Text = String.Format("{0:N0}", total);
                lblIva41.Text = String.Format("{0:N0}", iva);
                lblTotal41.Text = String.Format("{0:N0}", total);


                lblEmitidos41.Text = TotalUsados41.ToString();
                lblAnulados41.Text = TotalAnulados41.ToString();
                lblUtilizados41.Text = (TotalUsados41 + TotalAnulados41).ToString();

            }
           
         
               

            dr.Close();
            idte.CerrarTransaccion();

        this.CargaTotalesConta();
          
        }
     private void CargaTotalesConta()
    {
        Contabilidad conta = new Contabilidad(lblServidor.Text, this.mysql_root, this.mysql_pass);
        MySqlDataReader dr = conta.GetTotalByDiaEmpresa(lblNombreEmpresa.Text.Substring(0, 2), FuncionesClass.GetFechaMysql(lblFecha.Text));
        double total = 0;
        double exento = 0;

        if(dr.HasRows == true)
        {
            if (dr.Read())
            {
                total = Convert.ToDouble(dr["total"].ToString());
                exento = Convert.ToDouble(dr["exento"].ToString());
            }
        }

        dr.Close();

        conta.CerrarTransaccion();

        lblConta39.Text = String.Format("{0:N0}",total);
        lblConta41.Text = String.Format("{0:N0}", exento);


    }
        private double TotalDoc(string xml)
        {
            double total = 0;
            XmlDocument xmlDoc = xmlDoc = new XmlDocument();
            bool salida = false;
            string tag = "";
            try
            {
                xmlDoc.LoadXml(xml);
                string TIPO_XML = "";
                XmlNode node = xmlDoc.DocumentElement.FirstChild;
                XmlNodeList lstNodos = node.ChildNodes;

                TIPO_XML = xmlDoc.DocumentElement.Name;

                if (TIPO_XML == "DTE")
                {
                    for (int i = 0; i < lstNodos.Count; i++)
                    {
                        if (lstNodos[i].Name == "Encabezado")
                        {
                            XmlNodeList nodototal = ((XmlElement)lstNodos[0]).GetElementsByTagName("Totales");
                            // XmlNodeList receptor = ((XmlElement)lstNodos[0]).GetElementsByTagName("RutReceptor");
                            XmlNodeList lstChilds = lstNodos[i].ChildNodes;
                            tag = nodototal[0].InnerXml;
                            //total = Convert.ToDouble(nodototal[0].Value);

                            for (int j = 0; j < lstChilds.Count; j++)
                            {
                                XmlNode node2 = lstChilds[j];
                                if (node2.Name == "Totales")
                                {
                                    XmlNodeList fist = lstChilds[j].ChildNodes;
                                    for (int k = 0; k < fist.Count; k++)
                                    {
                                        XmlNode node3 = fist[k];
                                        if (node3.Name == "MntTotal")
                                        {
                                            tag = node3.InnerText;
                                            total = Convert.ToDouble(tag);
                                            break;
                                            //return 0;
                                        }
                                    }

                                }



                            }

                        }
                    }

                }

            }//end try
            catch(Exception  ex)
            {
                log.Error(ex);
            }
           

            return total;
        }


         private void GeneraArchivoRCOF()
        {
            DateTime xfecha = Convert.ToDateTime(lblFecha.Text);

            handler.secuenciaEnvio = txtSecuencia.Text;
            handler.rutEmpresa = Convert.ToDouble(lblRut.Text.Substring(0,9)) + "-" + lblRut.Text.Substring(9,1) ;
            handler.rutCertificado = lblRutCertificado.Text;
            handler.nombreCertificado = lblNombreCertificado.Text;
            handler.fechaResolucion = Convert.ToDateTime(lblFechaResolucion.Text);
            handler.numero_resolucion = Convert.ToInt32(lblNumeroResolucion.Text);

            //var rcof = handler.GenerarRCOF(dtes ,xfecha, Convert.ToInt32(lblEmitidos39.Text), Convert.ToInt32(lblAnulado39.Text) ,rangoUtilizado, rangoAnulado);
            var rcof = handler.GenerarRCOF(xfecha, Convert.ToInt32(lblEmitidos39.Text), Convert.ToInt32(lblAnulado39.Text), Convert.ToInt32(lblUtilizados39.Text),
                                            Convert.ToDouble(lbltotal39.Text), rangoUtilizado39, rangoAnulado39,
                                            Convert.ToInt32(lblEmitidos41.Text),Convert.ToInt32(lblAnulados41.Text),Convert.ToInt32(lblUtilizados41.Text), 
                                            Convert.ToDouble(lblTotal41.Text), rangoUtilizado41, rangoAnulado41,
                                            Convert.ToInt32(lblEmitidos61.Text),Convert.ToInt32(lblAnulados61.Text), Convert.ToInt32(lblUtilizado61.Text),
                                            Convert.ToDouble(lblTotal61.Text), rangoUtilizado61, rangoAnulado61);


            rcof.DocumentoConsumoFolios.Id = "RCOF_" + lblFecha.Text + "_N" + txtSecuencia.Text  ;
            btnGenerar.Enabled = false;
            string xmlString = string.Empty;
            var filePathArchivo = rcof.Firmar(lblNombreCertificado.Text, out xmlString);
            if (File.Exists(filePathArchivo))
            {
                FileInfo fi = new FileInfo(filePathArchivo);
                string destino = FuncionesClass._BASE_FOLDER_PROD + lblCliente.Text + @"\" + Convert.ToDouble(lblRut.Text.Substring(0, 9)) +  @"\Produccion\envios\RCOF\RCOF_D" +lblFecha.Text + "N" + txtSecuencia.Text + ".xml";
                fi.CopyTo(destino, true);
                string xml =  File.ReadAllText(filePathArchivo, Encoding.GetEncoding("ISO-8859-1"));

                long trackId = 0;
                if (chbRegistrar.Checked == true)
                {
                     trackId = 1111111111; 
                }
                else
                {
                     trackId =  handler.EnviarEnvioDTEToSII(filePathArchivo, "XH1F-EFZ5-ZH93", chbProduccion.Checked);
                }

                if (trackId != 0)
                {
                    this.GrabaXML(lblNombreEmpresa.Text.Substring(0, 2), FuncionesClass.GetFechaMysql(lblFecha.Text), xml,
                    lblDesde39.Text, lblHasta39.Text, Convert.ToDouble(lbltotal39.Text), trackId.ToString());

                    GrabaXMLDetalles(FuncionesClass.GetFechaMysql(lblFecha.Text), Convert.ToInt32(txtSecuencia.Text), Convert.ToDouble(lbltotal39.Text));
                    RadMessageBox.SetThemeName("TelerikMetroBlue");
                    RadMessageBox.Show(this, "Informe RCOF Generado y enviado satisfactoriamente [TrackID:" + trackId.ToString() + "]", "OK", MessageBoxButtons.OK, RadMessageIcon.Info);
                    log.Info("-> Envio RCOF Track " + trackId + " Fecha Contable[" + lblFecha.Text + "]");
                    this.Close();
                }             

            }
        }
        private void GrabaXMLDetalles(string xFecha, int xSecuencia, double xTotal)
        {
            PlaceSoft.Eltit.Class.clases.DTEClass dte = new PlaceSoft.Eltit.Class.clases.DTEClass(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS );
            string base_dte = lblCliente.Text + "_dte_" + Convert.ToDouble(lblRut.Text.Substring(0, 9));

            dte.GrabaRCOFDetalles(lblNombreEmpresa.Text.Substring(0, 2), xFecha, xSecuencia,  XML_DETALLE, "EDUVERGARA", base_dte, xTotal);

        }
        private void GrabaXML(string xlocal, string xfecha, string xml, string xdesde, string xhasta, double xTotal,string TrackID)
        {
        PlaceSoft.Eltit.Class.clases.DTEClass dte = new PlaceSoft.Eltit.Class.clases.DTEClass(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER , FuncionesClass.G_MYSQL_PASS);

            string base_dte =  lblCliente.Text + "_dte_" + Convert.ToDouble(lblRut.Text.Substring(0, 9));


            xml = xml.Replace("'", "~");
            dte.GrabaRCOF(xlocal, xfecha, Convert.ToInt32(txtSecuencia.Text), xml,
                xdesde, xhasta, Convert.ToDouble(lbltotal39.Text),TrackID, base_dte);

        }
        private void VerificaRcof(string xFecha)
        {
            string base_dte = "eltit_dte_" + Convert.ToDouble(lblRut.Text.Substring(0, 9));
            Caf fo = new Caf(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS );
            MySqlDataReader dr = fo.BuscaRCOF(lblNombreEmpresa.Text.Substring(0, 2), xFecha, base_dte);

            if (dr.HasRows == true)
            {
                if (dr.Read())
                {                  

                    if (dr["fae_GLOSA_sii"].ToString() == "CORRECTO")
                    {
                        pictureBox2.Image = Eltit.Properties.Resources.OK_48;
                        btnGenerar.Enabled = false;
                    }
                    else
                    {
                        pictureBox2.Image = Eltit.Properties.Resources.icons8_exclamacion;
                    }

                }
            }
            else
            {
                pictureBox2.Image = Eltit.Properties.Resources.icons8_exclamacion;
                txtSecuencia.Text = "1";
            }
            dr.Close();
            fo.CerrarTransaccion();
        }
        private void btnGenerar_Click(object sender, EventArgs e)
        {
            if(lblTrack.Text != "")
            {
                if(RadMessageBox.Show(this, "Este RCOF ya Contiene un Envío ¿Desea realmente enviarlo con una nueva Secuencia ["+ txtSecuencia.Text +"] ?", "OK", MessageBoxButtons.YesNo, RadMessageIcon.Error) == DialogResult.Yes)
                {
                    FuncionesClass fu = new FuncionesClass();
                    if (fu.PingToHost(lblServidor.Text) == true)
                    {
                        this.GeneraArchivoRCOF();
                    }
                    else
                    {
                        RadMessageBox.SetThemeName("TelerikMetroBlue");
                        RadMessageBox.Show(this, "No se pudo establecer conexión con el host de destino[" + lblServidor.Text + "]", "OK", MessageBoxButtons.OK, RadMessageIcon.Error);
                    }
                }
            }
            else
            {
                FuncionesClass fu = new FuncionesClass();
                if (fu.PingToHost(lblServidor.Text) == true)
                {
                    this.GeneraArchivoRCOF();
                }
                else
                {
                    RadMessageBox.SetThemeName("TelerikMetroBlue");
                    RadMessageBox.Show(this, "No se pudo establecer conexión con el host de destino[" + lblServidor.Text + "]", "OK", MessageBoxButtons.OK, RadMessageIcon.Error);
                }
            }
           
            
        }
   
         private void RadPageView1_SelectedPageChanged(object sender, EventArgs e)
        {

        }
    }
