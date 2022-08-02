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
using Eltit;

namespace SamplesDTE
{
    public partial class frmPopGeneraLibro : Telerik.WinControls.UI.RadForm
    {
        private Icon[] icons = new Icon[2];
        private int currentIcon = 0;
        Handler handler = new Handler();
        List<PlaceSoft.DTE.Engine.Documento.DTE> dtes = new List<PlaceSoft.DTE.Engine.Documento.DTE>();
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public frmPopGeneraLibro()
        {
            InitializeComponent();
        }

        private void frmPopGeneraLibro_Load(object sender, EventArgs e)
        {           
            lblInformacion.Text = "EMPRESA: " + FuncionesClass.G_EMPRESANOMBRE;
            lblRut.Text = Convert.ToDouble(FuncionesClass.G_EMPRESARUT.Substring(0, 9)) + "-" + FuncionesClass.G_EMPRESARUT.Substring(9, 1);
            handler.rutEmpresa = lblRut.Text;

            //this.BuscaVentas();
            
        }
    
        private void InicializaControlesDeEmpresa()
        {
            lblRut.Text = FuncionesClass.G_EMPRESARUT;
            lblNombreEmpresa.Text = FuncionesClass.G_EMPRESANOMBRE;
              
        }

        // private void BuscaVentas()
        //{
        //    VentasClass ventas = new VentasClass(FuncionesClass.G_SERVIDOR);
        //    ventas.setBaseDTE(FuncionesClass.BASE_DTE);
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
        //    double cantnulas = 0;
        //    string nombredoc = "";
        //    string tipoFiscal = "";
        //    string desde = lblFecha.Text + "-01";
        //    string hasta = lblFecha.Text + "-31";
        //    MySqlDataReader dr = ventas.GetBoletasByLocalDesdeHasta(lblNombreEmpresa.Text.Substring(0, 2),desde,hasta, "BEL");
        //    object img;
            
        //    if (dr.HasRows == true)
        //    {
        //        while(dr.Read())
        //        {

        //            //string xml = dr["fae_xml"].ToString(); // File.ReadAllText(dr["fae_xml"], Encoding.GetEncoding("ISO-8859-1"));
        //            //var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);

        //            //neto = dte.Documento.Encabezado.Totales.MontoNeto;
        //            //exento = dte.Documento.Encabezado.Totales.MontoExento;
        //            //iva = dte.Documento.Encabezado.Totales.IVA;
        //            //total = dte.Documento.Encabezado.Totales.MontoTotal;

        //            if(count == 0)
        //            {
        //                inicial = Convert.ToInt32( dr["foliosii"].ToString());

        //            }
        //            else
        //            {
        //                final = Convert.ToInt32(dr["foliosii"].ToString()); 
        //            }

        //            if (dr["tipo_doc"].ToString() == "BEL")
        //            {
        //                cant39 = cant39 + 1;
        //                tipoFiscal = "39";
        //            }
        //            if (dr["tipo_doc"].ToString() == "BEE")
        //            {
        //                cant41 = cant41 + 1;
        //                tipoFiscal = "41";
        //            }
        //            if (dr["tipo_doc"].ToString() == "NBE")
        //            {
        //                cantnulas = cantnulas + 1;
        //                tipoFiscal = "61";
        //            }

        //            nombredoc = FuncionesClass.getNombredocumentoByCodigo(dr["tipo_doc"].ToString());
        //            string xml = dr["fae_xml"].ToString();
        //            txtXML.Text += xml + Environment.NewLine;
        //            var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);
        //            totaldte = dte.Documento.Encabezado.Totales.MontoTotal;
        //            dtes.Add(dte);
        //            if (Convert.ToDouble(dr["monto_total"]) == totaldte)
        //            {
        //                img = Eltit.Properties.Resources.OK_48;
        //            }
        //            else
        //            {
        //                img = Eltit.Properties.Resources.icons8_exclamacion;
        //            }

        //            gvInforme.Rows.Add(tipoFiscal + " " + nombredoc, dr["foliosii"].ToString(), 
        //                                dr["fecha_emision"].ToString(), dr["caja_doc"].ToString(),
        //                                String.Format("{0:N0}", dr["monto_total"]),img);

        //            exento = exento + Convert.ToDouble(dr["monto_exento"]);
        //            total = total + Convert.ToDouble(dr["monto_total"]);


        //            count++;

        //        }
        //    }

            
        //    neto = total / div ;
        //    neto = Math.Round(neto, 4);
        //    iva = total - neto;

        //    neto = Math.Round(neto , 0, MidpointRounding.AwayFromZero);
        //    iva = Math.Round( iva, 0, MidpointRounding.AwayFromZero);
        //    total = Math.Round(neto + exento + iva, 0, MidpointRounding.AwayFromZero);

        //    lblNeto.Text = String.Format("{0:N0}", neto);
        //    lblExento.Text = String.Format("{0:N0}", exento);
        //    lblIva.Text = String.Format("{0:N0}", iva);
        //    lbltotal.Text = String.Format("{0:N0}", total);

        //    lblElectronicas.Text = cant39.ToString();
        //    lblExentas.Text = cant41.ToString();
        //    lblEmitidos.Text = (gvInforme.Rows.Count).ToString();
        //    lblAnulado.Text = cantnulas.ToString();

        //    lblDesde.Text = inicial.ToString();
        //    lblHasta.Text = final.ToString();

                                          
        //    VerificaRcof(FuncionesClass.GetFechaMysql(lblFecha.Text));
        //    dr.Close();
        //    ventas.CerrarTransaccion();
          
        //}

        private void GeneraLibroBoletas()
        {
            //var rcof = handler.GenerarRCOF(dtes);
            //rcof.DocumentoConsumoFolios.Id = "LibroBoletas_P" + lblFecha.Text + "_T" + lblEnvios.Text;
            //string xmlString = string.Empty;
            //var filePathArchivo = rcof.Firmar(FuncionesClass.G_DTE_NOMBRE_CERTIFICADO, out xmlString);
            //if (File.Exists(filePathArchivo))
            //{
            //    FileInfo fi = new FileInfo(filePathArchivo);
            //    string destino = FuncionesClass._BASE_FOLDER_PROD + @"\envios\RCOF_D" +lblFecha.Text + "N" + lblEnvios.Text + ".xml";
            //    fi.CopyTo(destino, true);
            //    string xml =  File.ReadAllText(filePathArchivo, Encoding.GetEncoding("ISO-8859-1"));
            //    long trackId = handler.EnviarEnvioDTEToSII(filePathArchivo, "XH1F-EFZ5-ZH93", chbProduccion.Checked);
            //    if(trackId > 0)
            //    {
            //        this.GrabaXML(lblNombreEmpresa.Text.Substring(0, 2), FuncionesClass.GetFechaMysql(lblFecha.Text), xml,
            //       lblDesde.Text, lblHasta.Text, Convert.ToDouble(lbltotal.Text) , trackId.ToString());

            //        MessageBox.Show("Informe RCOF Generado y enviado satisfactoriamente [TrackID:" + trackId.ToString() + "]");
            //        log.Info("-> Envio RCOF Track " + trackId + " Fecha Contable[" + lblFecha.Text + "]");
            //    }
            //}


           // List<PlaceSoft.DTE.Engine.Documento.DTE> dtes = new List<PlaceSoft.DTE.Engine.Documento.DTE>();
            //foreach (string pathFile in pathFiles)
            //{
            //    string xml = File.ReadAllText(pathFile, Encoding.GetEncoding("ISO-8859-1"));
            //    var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);
            //    dtes.Add(dte);
            //}

            var libro = handler.GenerateLibroBoletas(dtes, lblFecha.Text);
            libro.EnvioLibro.Caratula.FolioNotificacion = Convert.ToInt32(txtNroNotificacion.Text);
            libro.EnvioLibro.Id = "LibroBoletas_" + lblFecha.Text + "_" + lblEnvios.Text ;
            var filePathArchivo = libro.Firmar(FuncionesClass.G_DTE_NOMBRE_CERTIFICADO);
            if (File.Exists(filePathArchivo))
            {
                FileInfo fi = new FileInfo(filePathArchivo);
                string destino = FuncionesClass._BASE_FOLDER_PROD + @"\LibrosCV\" + libro.EnvioLibro.Id + "_" + DateTime.Now.Ticks.ToString() + ".xml";
                fi.CopyTo(destino, true);
                string xml = File.ReadAllText(filePathArchivo, Encoding.GetEncoding("ISO-8859-1"));
                this.GrabaXML(xml, Convert.ToDouble(lbltotal.Text));
                MessageBox.Show("Libro Boletas Generado correctamente");

                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = destino;
                proc.Start();

            }
        




        }
        private void GrabaXML( string xml, double xTotal)
        {
            //VentasClass ven = new VentasClass(FuncionesClass.G_SERVIDOR);
            //ven.setBaseDTE(FuncionesClass.BASE_DTE);
            //xml = xml.Replace("'", "~");
            //ven.GrabaXMLLibroBoletas(lblNombreEmpresa.Text.Substring(0, 2), lblFecha.Text, lblEnvios.Text, xml,xTotal);
            //caf.GrabaRCOF(xlocal, xfecha, Convert.ToInt32(lblEnvios.Text), xml,
            //    xdesde, xhasta, Convert.ToDouble(lbltotal.Text),TrackID);

        }
        private void VerificaRcof(string xFecha)
        {
            Caf fo = new Caf(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER,FuncionesClass.G_MYSQL_PASS);
            string base_dte = "eltit_fae_" + Convert.ToDouble(lblRut.Text.Substring(0, 9));
            MySqlDataReader dr = fo.BuscaRCOF(lblNombreEmpresa.Text.Substring(0, 2), base_dte, xFecha);

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
                lblEnvios.Text = "1";
            }
            dr.Close();
            fo.CerrarTransaccion();
        }
        private void btnGenerar_Click(object sender, EventArgs e)
        {
            this.GeneraLibroBoletas();
        }
   
  

    }
}
