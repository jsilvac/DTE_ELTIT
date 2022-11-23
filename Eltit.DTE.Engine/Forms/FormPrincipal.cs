using Eltit.DTE.clases;
using MySql.Data.MySqlClient;
using PlaceSoft.DTE.Engine.Documento;
using PlaceSoft.Eltit.Class.clases;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Eltit.DTE.Forms
{
    public partial class Form1 : Form
    {
        private static readonly log4net.ILog log =
         log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //List<PlaceSoft.DTE.Engine.Documento.DTE> dtes = new List<PlaceSoft.DTE.Engine.Documento.DTE>();
        public PlaceSoft.DTE.Engine.Documento.DTE dte;
        iTextSharp.text.Image jpg;
        public DatosEmisor emisor;
        private int CONTADOR = 0;
        public string DOC_TIPO_DTE;
        public string DOC_TIPO_INTERNO;
        public string DOC_FOLIOSII;
        public string DOC_LOCAL;
        public string DOC_FECHA_EMISION;
        public string DOC_NOMBRE_IMPRESORA;
        // VICTOR
        public string DOC_NOMBRE_CAJERA;
        public string DOC_CAJA;
        public string DOC_VUELTO;
        public List<string> DOC_FORMA_PAGO;
        public double DOC_TOTAL_A_PAGAR;
        public double DOC_TASA_IVA = 19;
        public bool DOC_REIMPRESO;
        public bool IMPRIME_CEDIBLE;
        public int DOC_REDONDEO;
        public int DOC_TIPO_TRASLADO;
        public string DOC_RUT;
        public string Observaciones;

        public string DOC_SERVIDOR;
        public string DOC_PASSWORD_MYSQL;
        public string DOC_ROOT_MYSQL;

        VentasClass miClase;
        Cajera xcajera;
        Documentos doc;

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            log.Debug("Inicializando formulario...");
            this.roundCorners(this);
            this.Refresh();

            timer1.Enabled = true;
            timer1.Start();
            timer1.Interval = 3;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                if(CONTADOR == 0)
                {
                    lblInformacion.Text = "IMPRIMIENDO FACTURA  ELECTRÒNICA Nº" + DOC_FOLIOSII;
                    lblInformacion.Refresh();
                    ImprimirDTE();
                    CONTADOR++;
                    Application.Exit();
                    this.Close();
                }
            }
            catch(Exception ex)
            {
                log.Error("El error es: ", ex);
                Application.Exit();
                this.Close();
            }
        }

        private void ImprimirDTE()
        {

            string XML = "";
            /********** AQUI RESCATAR DATOS DEL EMISOR ******************/
            emisor = new DatosEmisor(DOC_RUT, "eltit_", "192.168.4.9", DOC_LOCAL);


            /*************************************************************/

            /*********** aqui rescatar los datos del el xml **************/

            doc = new Documentos(this.DOC_SERVIDOR, this.DOC_ROOT_MYSQL, this.DOC_PASSWORD_MYSQL);
           // miClase = new VentasClass(this.DOC_SERVIDOR, this.DOC_PASSWORD_MYSQL,this.DOC_ROOT_MYSQL);


            if (DOC_TIPO_DTE == "39" || DOC_TIPO_DTE == "41")
            {
                DTEClass dte = new DTEClass(this.DOC_SERVIDOR, this.DOC_ROOT_MYSQL, this.DOC_PASSWORD_MYSQL);
                XML = dte.GetXMLBoletas(DOC_LOCAL, DOC_TIPO_DTE, DOC_FOLIOSII, DOC_FECHA_EMISION);
            }
            if (DOC_TIPO_DTE == "33" || DOC_TIPO_DTE == "61" )
            {
                miClase = new VentasClass(this.DOC_SERVIDOR, this.DOC_PASSWORD_MYSQL, this.DOC_ROOT_MYSQL);
                XML = miClase.GetXMLFacturas(DOC_LOCAL, DOC_TIPO_DTE, DOC_FOLIOSII, DOC_FECHA_EMISION);
            }

            if ( DOC_TIPO_DTE == "52")
            {
                miClase = new VentasClass(this.DOC_SERVIDOR, this.DOC_PASSWORD_MYSQL, this.DOC_ROOT_MYSQL);
                XML = miClase.GetXMLGuias(DOC_LOCAL, DOC_TIPO_DTE, DOC_FOLIOSII, DOC_FECHA_EMISION);
            }


            //string XML = miClase.GetXMLFacturas(DOC_LOCAL, DOC_TIPO_DTE, DOC_FOLIOSII, DOC_FECHA_EMISION);

            /******************** aqui rescata del ventas *****************************************/


            /*************************************************************/

            dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(XML);
            pictureBoxTimbre.Image = dte.Documento.TimbrePDF417;

            
            this.ImprimeTicket();

        }

        private void ImprimeTicket()
        {
            Boolean isFinish;
            PrintDocument pdPrint = new PrintDocument();

            if (DOC_TIPO_DTE == "33" || DOC_TIPO_DTE == "61" || DOC_TIPO_DTE == "52") // SI ES CASO DE FACTURAS O NOTAS DE CRÉDITO ELECTRONICAS
            {
                pdPrint.PrintPage += new PrintPageEventHandler(pdPrint_PrintPage_FAE);
            }

            if(DOC_TIPO_DTE == "39" || DOC_TIPO_DTE == "41")
            {
                pdPrint.PrintPage += new PrintPageEventHandler(pdPrint_PrintPage_BOLETAS);
            }
            pdPrint.PrinterSettings.PrinterName = DOC_NOMBRE_IMPRESORA;

            try
            {

                if (1 == 1)
                {
                    if (pdPrint.PrinterSettings.IsValid)
                    {
                        pdPrint.DocumentName = "Imprimiendo Comprobante...";
                        // Start printing.
                        pdPrint.PrintController = new StandardPrintController();
                        pdPrint.Print();

                        isFinish = false;
                    }
                    else
                        MessageBox.Show("IMPRESORA NO DISPONIBLE.", "Program06", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }
                else
                {
                    MessageBox.Show("Failed to open printer status monitor.", "Program06", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Imprimicion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void pdPrint_PrintPage_BOLETAS(object sender, PrintPageEventArgs e)
        {
            float x, y;

            Font printFoot = new Font("Arial", (float)6.3, FontStyle.Regular, GraphicsUnit.Point);
            Font printdETALLE = new Font("Arial", (float)7, FontStyle.Bold, GraphicsUnit.Point);
            // Instantiate font objects used in printing.
            Font printFont = new Font("Arial", (float)7, FontStyle.Regular, GraphicsUnit.Point); // Substituted to FontA Font
            Font gridFont = new Font("Arial", (float)6, FontStyle.Regular, GraphicsUnit.Point); // Substituted to FontA Font
            Font gridTotal = new Font("Arial", (float)11, FontStyle.Bold, GraphicsUnit.Point); // Substituted to FontA Font

            e.Graphics.PageUnit = GraphicsUnit.Point;

            /************ IMPRIME IMAGEN A LA IZQUIERDA *******************/
            x = 2;
            y = 0;
            e.Graphics.DrawImage(pbImage.Image, x, y, pbImage.Width, pbImage.Height);
            /************ IMPRIME RECTANGULO ***********************/
            string text2 = "R.U.T. " + emisor.Rut;
            Font font2 = new Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point);
            Font font5 = new Font("Arial", (float)6.7, GraphicsUnit.Point);


            Rectangle rect2;
            if (DOC_TIPO_DTE == "41")
            {
                rect2 = new Rectangle(70, 2, 120, 60);
            }
            else
            {
                rect2 = new Rectangle(70, 2, 120, 50);
            }

            e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rect2));
            y = 7;
            e.Graphics.DrawString(text2, font2, Brushes.Black, 90, y);
            y = y + 14;
            if (DOC_TIPO_DTE == "39")
            {
                e.Graphics.DrawString("BOLETA  ELECTRONICA", font2, Brushes.Black, 76, y);
            }
            else
            {
                //e.Graphics.DrawString(txtTipoDocumento.Text, new Font("Arial", 8, FontStyle.Bold, GraphicsUnit.Point), Brushes.Black, 76, y);
            }

            if (DOC_TIPO_DTE == "41")
            {
                e.Graphics.DrawString("      BOLETA EXENTA", font2, Brushes.Black, 76, y);
                y = y + 10;
                e.Graphics.DrawString("        ELECTRÓNICA", font2, Brushes.Black, 76, y);
            }
            y = y + 14;
            e.Graphics.DrawString("N° " + dte.Documento.Encabezado.IdentificacionDTE.Folio.ToString().PadLeft(10, Convert.ToChar("0")), font2, Brushes.Black, 97, y);
            y = y + 20;
            e.Graphics.DrawString("S.I.I. - " + emisor.Sii, printFont, Brushes.Black, 90, y);
            y = y + 15;
            Font FNTGRANDE = new Font("Arial", (float)8.6, FontStyle.Bold, GraphicsUnit.Point);
            e.Graphics.DrawString(emisor.Razon, FNTGRANDE, Brushes.Black, 1, y);
            /************* imprime membrete *********************/
            y = y + 20;
            e.Graphics.DrawString("Giro :", font2, Brushes.Black, 1, y);
            e.Graphics.DrawString(emisor.Giro_1, printFont, Brushes.Black, 27, y);
            if (emisor.Giro_2 != "")
            {
                y = y + 10;
                e.Graphics.DrawString(emisor.Giro_2, printFont, Brushes.Black, 27, y);
            }
            if (emisor.Giro_3 != "")
            {
                y = y + 10;
                e.Graphics.DrawString(emisor.Giro_3, printFont, Brushes.Black, 32, y);
            }
            if (emisor.Giro_4 != "")
            {
                y = y + 10;
                e.Graphics.DrawString(emisor.Giro_4, printFont, Brushes.Black, 32, y);
            }


            y = y + 12;
            e.Graphics.DrawString("Dirección :", font2, Brushes.Black, 1, y);
            e.Graphics.DrawString(emisor.Direccion + ", " + emisor.Comuna + ".", font5, Brushes.Black, 48, y + 2);
            y = y + 12;
            e.Graphics.DrawString("Fono :", font2, Brushes.Black, 1, y);
            e.Graphics.DrawString(emisor.Fono, font5, Brushes.Black, 30, y + 2);
            y = y + 12;
            e.Graphics.DrawString("Email :", font2, Brushes.Black, 1, y);
            e.Graphics.DrawString(emisor.Email, font5, Brushes.Black, 30, y + 2);
            y = y + 12;
            e.Graphics.DrawString("Vendedor :", font2, Brushes.Black, 1, y);
            e.Graphics.DrawString(DOC_NOMBRE_CAJERA, font5, Brushes.Black, 50, y + 2);
            y = y + 12;
            e.Graphics.DrawString("Fecha Emisión :", font2, Brushes.Black, 1, y);
            e.Graphics.DrawString(dte.Documento.Encabezado.IdentificacionDTE.FechaEmision.ToShortDateString(), font5, Brushes.Black, 70, y + 2);
            e.Graphics.DrawString("Caja :", font2, Brushes.Black, 122, y);
            e.Graphics.DrawString(DOC_CAJA, font5, Brushes.Black, 152, y + 2);
            y = y + 5;


            /*************** REGION QUE IMPRIME LOS DETALLES  ***************/
            Font font3 = new Font("Arial", 6, FontStyle.Bold, GraphicsUnit.Point);
            y = y + 6;
            e.Graphics.DrawString("____________________________________________________________", font3, Brushes.Black, x + 1, y);
            y = y + 8;
            e.Graphics.DrawString("CANT. DESCRIPCION.                                    PRECIO   DESC.   TOTAL", font3, Brushes.Black, x + 2, y);
            y = y + 3;
            e.Graphics.DrawString("____________________________________________________________", font3, Brushes.Black, x + 1, y);

            /*************** RECORRER ROLLO DE LA GRILLA ***************/
            y += 8;
            List<PlaceSoft.DTE.Engine.Documento.Detalle> listDetalles = new List<PlaceSoft.DTE.Engine.Documento.Detalle>();
            listDetalles = dte.Documento.Detalles;

            foreach (PlaceSoft.DTE.Engine.Documento.Detalle detalle in listDetalles)
            {

                double cantidad = detalle.Cantidad;
                string descripcion = detalle.Nombre;
                string descripcion2 = "";
                if (detalle.Nombre.Length > 25)
                {
                    descripcion = detalle.Nombre.Substring(0, 25);
                    descripcion2 = detalle.Nombre.Substring(25, detalle.Nombre.Length - 25);
                }

                string totallinea = string.Format(CultureInfo.CurrentCulture, "{0:C0}", detalle.MontoItem);
                string preciolinea = string.Format(CultureInfo.CurrentCulture, "{0:C0}", detalle.Precio);
                int dctolinea = detalle.Descuento;
                string porDescuento = "0%";
                if (detalle.DescuentoPorcentaje > 0)
                {
                    porDescuento = detalle.DescuentoPorcentaje + "%"; //@String.Format(new CultureInfo("es-CL"), "{0:P2}", detalle.DescuentoPorcentaje);
                }

                int pad = 183 + (6 - totallinea.Length) * 3;
                int pad2 = 133 + (6 - preciolinea.Length) * 3;
                int pad3 = 5 + (3 - cantidad.ToString().Length) * 3;
                e.Graphics.DrawString(cantidad.ToString(), printFoot, Brushes.Black, pad3, y);

                // detalle.CodigosItem = new List<PlaceSoft.DTE.Engine.Documento.CodigoItem>();

                foreach (CodigoItem cod in detalle.CodigosItem)
                {
                    e.Graphics.DrawString(cod.ValorCodigo, printdETALLE, Brushes.Black, 23, y);
                }
                e.Graphics.DrawString(detalle.Nombre, printdETALLE, Brushes.Black, 23, y + 8);
                //e.Graphics.DrawString(descripcion , printFoot, Brushes.Black, 23, y + 6);
                //e.Graphics.DrawString(descripcion2, printFoot, Brushes.Black, 23, y+12);
                e.Graphics.DrawString(preciolinea, printFoot, Brushes.Black, pad2, y + 15);
                e.Graphics.DrawString(porDescuento, printFoot, Brushes.Black, 158, y + 15);
                e.Graphics.DrawString(totallinea, printFoot, Brushes.Black, pad, y + 15);
                y += 21;
            }

            y = y - 1;


            // y += 3;
            e.Graphics.DrawString("____________________________________________________________", font3, Brushes.Black, x, y);

            double TOTALEXENTO = dte.Documento.Encabezado.Totales.MontoExento;
            if (TOTALEXENTO > 0)
            {
                y = y + 10;
                Font font4 = new Font("Arial", 7, FontStyle.Bold, GraphicsUnit.Point);
                //double TOTALEXENTO = dte.Documento.Encabezado.Totales.MontoExento;
                e.Graphics.DrawString("Monto Exento :", font4, Brushes.Black, 1, y);
                string srtexento = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALEXENTO);
                string totalfinalex = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtexento.Replace(",00", ""));
                int padfinalex = 182 + (4 - totalfinalex.Length) * 3;
                e.Graphics.DrawString(totalfinalex, font4, Brushes.Black, padfinalex, y + 1);
            }



            /************************** REGION NETO *****************************/
            y = y + 22;
            double TOTALNETO = dte.Documento.Encabezado.Totales.MontoNeto;
            e.Graphics.DrawString("Neto :", font2, Brushes.Black, 1, y);
            string srtNeto = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALNETO);
            string netofinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtNeto.Replace(",00", ""));
            int padnetofinal = 178 + (4 - netofinal.Length) * 3;
            e.Graphics.DrawString(netofinal, font2, Brushes.Black, padnetofinal, y + 1);
            /********************************************************************/
            /************************** REGION IVA *****************************/
            y = y + 12;
            double TOTALIVA = dte.Documento.Encabezado.Totales.IVA;
            e.Graphics.DrawString("IVA :", font2, Brushes.Black, 1, y);
            string srtIva = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALIVA);
            string ivafinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtIva.Replace(",00", ""));
            int padivafinal = 178 + (4 - ivafinal.Length) * 3;
            e.Graphics.DrawString(ivafinal, font2, Brushes.Black, padivafinal, y + 1);
            /********************************************************************/


            /************************* REGION TOTAL VENTA ***********************/
            y = y + 12;
            //double TOTALVENTA = (dte.Documento.Encabezado.Totales.MontoTotal - DOC_MONTO_PROPINA) ; // Reemplazado por el total de ventas
            double TOTALVENTA = (dte.Documento.Encabezado.Totales.MontoTotal);
            // de la pantalla de pagos, por la ley del redondeo. 08-09-2019 por eduvergara
            //double TOTALVENTA = DOC_TOTAL_A_PAGAR - DOC_MONTO_PROPINA  ;
            e.Graphics.DrawString("Monto Total :", font2, Brushes.Black, 1, y);
            string srtTotal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALVENTA);
            string totalfinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtTotal.Replace(",00", ""));
            int padfinal = 178 + (4 - totalfinal.Length) * 3;
            e.Graphics.DrawString(totalfinal, font2, Brushes.Black, padfinal, y + 1);
            /*********************************************************************/

            /************************* MONTO PROPINA ****************************/
            //y = y + 15;
            //if (DOC_MONTO_PROPINA != 0)
            //{
            //    double TOTALPROPINA = DOC_MONTO_PROPINA;
            //    e.Graphics.DrawString("Monto Propina :", font2, Brushes.Black, 1, y);
            //    string srtPropina = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALPROPINA);
            //    string propinafinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtPropina.Replace(",00", ""));
            //    int padpropina = 178 + (4 - propinafinal.Length) * 3;
            //    e.Graphics.DrawString(propinafinal, font2, Brushes.Black, padpropina, y + 1);
            //}
            /********************************************************************/
            /************************* REGION TOTAL A PAGAR ***********************/
            //y = y + 12;

            //double TOTALVENTA = DOC_TOTAL_A_PAGAR;
            //e.Graphics.DrawString("Monto Total :", font2, Brushes.Black, 5, y);
            //string srtTotal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALVENTA);
            //string totalfinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtTotal.Replace(",00", ""));
            //int padfinal = 178 + (4 - totalfinal.Length) * 3;
            //e.Graphics.DrawString(totalfinal, font2, Brushes.Black, padfinal, y + 1);
            /*********************************************************************/


            /********************* REGION DEL REDONDEO *************************/
            if (this.DOC_REDONDEO != 0)
            {
                y = y + 22;
                double TOTALREDONDEO = this.DOC_REDONDEO;
                e.Graphics.DrawString("Ley 20.956 :", font2, Brushes.Black, 1, y);
                string srtRedondeo = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALREDONDEO);
                string redondeofinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtRedondeo.Replace(",00", ""));
                e.Graphics.DrawString(redondeofinal, font2, Brushes.Black, 185, y + 1);

            }

            y = y + 16;
            double TOTALAPAGAR = DOC_TOTAL_A_PAGAR;
            e.Graphics.DrawString("TOTAL A PAGAR :", font2, Brushes.Black, 1, y);
            string srtTotalpagar = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALAPAGAR);
            string totalpagarfinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtTotalpagar.Replace(",00", ""));
            int padfinalpagar = 178 + (4 - totalpagarfinal.Length) * 3;
            e.Graphics.DrawString(totalpagarfinal, font2, Brushes.Black, padfinalpagar, y + 1);

            /******************************************************************/
            y = y + 7;
            e.Graphics.DrawString("------------------------------------------------------------------", font2, Brushes.Black, 5, y);

            y = y + 12;
            //string strFormaPago = DOC_FORMA_PAGO;
            //List<string> strFormaPago = DOC_FORMA_PAGO;
            //int padfinalPago = (185 - (strFormaPago.Length * 4));
            //e.Graphics.DrawString("Formas de Pago :", font2, Brushes.Black, 1, y);
            ////e.Graphics.DrawString(strFormaPago, font2, Brushes.Black, padfinalPago + ((15-strFormaPago.Length)*2)  , y );
            //y = y + 6;
            //VentasClass ve = new VentasClass();
            //MySqlDataReader dr = ve.GetPagosDocumento(DOC_LOCAL, TIPO_INTERNO, DOC_NUMERO, DOC_CAJA, DOC_FECHA_EMISION);
            //Font fontFormaPago = new Font("Arial", 8, FontStyle.Bold, GraphicsUnit.Point);

            //if (dr.HasRows == true)
            //{
            //    while (dr.Read())
            //    {
            //        y = y + 10;
            //        padfinalPago = (176 - (dr["nombre"].ToString().Length * 4));
            //        strFormaPago = dr["nombre"].ToString();
            //        e.Graphics.DrawString(strFormaPago, fontFormaPago, Brushes.Black, 1, y);
            //        e.Graphics.DrawString(string.Format(CultureInfo.CurrentCulture, "{0:C0}", dr["monto_total"]), fontFormaPago, Brushes.Black, 165, y);

            //    }
            //}
            //else
            //{
            //    y = y + 10;
            //    e.Graphics.DrawString(strFormaPago, font2, Brushes.Black, 1, y);
            //    //e.Graphics.DrawString(strFormaPago, font2, Brushes.Black, padfinalPago + ((15 - strFormaPago.Length) * 2), y);
            //}

            if (DOC_TIPO_DTE != "G4")
            {     List<string> strFormaPago = DOC_FORMA_PAGO;
                //int padfinalPago = (176 - (strFormaPago.Lengt * 4));
                e.Graphics.DrawString("Forma de Pago :", font2, Brushes.Black, 5, y);
                y = y + 3;
                foreach (string row in strFormaPago)
                {

                    e.Graphics.DrawString(row.ToString(), font5, Brushes.Black, 77, y);
                    y = y + 10;
                }
           

            }


            //dr.Close();
            //ve.CerrarTransaccion();


            y = y + 20;
            string strVuelto = DOC_VUELTO;
            e.Graphics.DrawString("Vuelto :", font2, Brushes.Black, 1, y);
            string srtvuelto = string.Format(CultureInfo.CurrentCulture, "{0:C0}", DOC_VUELTO);
            string totalVuelto = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtvuelto.Replace(",00", ""));
            int padVuelto = 178 + (4 - totalVuelto.Length) * 3;
            e.Graphics.DrawString(totalVuelto, font2, Brushes.Black, padVuelto, y + 1);



            y = y + 20;
            if (DOC_REIMPRESO == false)
            {
                e.Graphics.DrawString("Hora Venta: " + DateTime.Now.ToString("HH:mm:ss"), font5, Brushes.Black, 1, y + 2);
            }
            else
            {
                e.Graphics.DrawString("Documento Reimpreso: " + DateTime.Now.ToString("HH:mm:ss"), font5, Brushes.Black, 1, y + 2);
            }


            y = y + 14;

            // wi = (float)pictureBoxTimbre.Width * (float)0.7;
            // este si funciona
            //e.Graphics.DrawImage(pictureBoxTimbre.Image, 5, y, pictureBoxTimbre.Width * (float)0.6, pictureBoxTimbre.Height * (float)0.6);
            e.Graphics.DrawImage(pictureBoxTimbre.Image, 5, y, pictureBoxTimbre.Width * (float)0.6, pictureBoxTimbre.Height * (float)0.8);

            y = y + 90;
            Font printFont2 = new Font("Arial", (float)6, FontStyle.Bold, GraphicsUnit.Point);
            //
            e.Graphics.DrawString("Timbre Electrónico S.I.I.", printFont2, Brushes.Black, 68, y);
            y = y + 15;
            e.Graphics.DrawString(emisor.Glosa_res + " Verifique su documento en:", printFoot, Brushes.Black, 40, y);
            //y = y + 7;
            //e.Graphics.DrawString("Verifique su documento en:", printFoot, Brushes.Black, 58, y);
            y = y + 7;
            e.Graphics.DrawString(emisor.Web_verificacion, printFoot, Brushes.Black, 63, y);
            y = y + 20;
            //e.Graphics.DrawString("Contrate ya su documentación electrónica en http://www.placesoft.cl/", printFoot, Brushes.Black, 5, y);
            e.HasMorePages = false;
        }

        private void pdPrint_PrintPage_FAE(object sender, PrintPageEventArgs e)
        {
            float x, y;

            Font printFoot = new Font("Arial", (float)6.3, FontStyle.Regular, GraphicsUnit.Point);
            // Instantiate font objects used in printing.
            Font printFont = new Font("Arial", (float)7, FontStyle.Regular, GraphicsUnit.Point); // Substituted to FontA Font
            Font gridFont = new Font("Arial", (float)6, FontStyle.Regular, GraphicsUnit.Point); // Substituted to FontA Font
            Font gridTotal = new Font("Arial", (float)11, FontStyle.Bold, GraphicsUnit.Point); // Substituted to FontA Font

            e.Graphics.PageUnit = GraphicsUnit.Point;

            /************ IMPRIME IMAGEN A LA IZQUIERDA *******************/
            x = 2;
            y = 0;
            e.Graphics.DrawImage(pbImage.Image, x, y, pbImage.Width, pbImage.Height);
            /************ IMPRIME RECTANGULO ***********************/
            string text2 = "R.U.T. " + emisor.Rut;
            Font font2 = new Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point);
            Font font5 = new Font("Arial", (float)6.7, GraphicsUnit.Point);

            Rectangle rect2 = new Rectangle(70, 2, 120, 55);
            e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rect2));
            y = 7;
            e.Graphics.DrawString(text2, font2, Brushes.Black, 90, y);
            y = y + 14;
            if (DOC_TIPO_DTE == "33")
            {
                e.Graphics.DrawString("FACTURA ELECTRONICA", font2, Brushes.Black, 75, y + 3);
                y = y + 9;
            }
            if (DOC_TIPO_DTE == "61")
            {
                e.Graphics.DrawString("NOTA DE CRÉDITO", new Font("Arial", 8, FontStyle.Bold, GraphicsUnit.Point), Brushes.Black, 92, y - 1);
                y = y + 12;
                e.Graphics.DrawString("ELECTRÓNICA", new Font("Arial", 8, FontStyle.Bold, GraphicsUnit.Point), Brushes.Black, 99, y - 3);
            }
            if (DOC_TIPO_DTE == "52")
            {
                e.Graphics.DrawString("GUIA DE DESPACHO", new Font("Arial", 8, FontStyle.Bold, GraphicsUnit.Point), Brushes.Black, 92, y - 1);
                y = y + 12;
                e.Graphics.DrawString("ELECTRÓNICA", new Font("Arial", 8, FontStyle.Bold, GraphicsUnit.Point), Brushes.Black, 99, y - 3);
            }

            y = y + 10;
            e.Graphics.DrawString("N° " + dte.Documento.Encabezado.IdentificacionDTE.Folio.ToString().PadLeft(10, Convert.ToChar("0")), font2, Brushes.Black, 97, y);
            y = y + 20;
            e.Graphics.DrawString("S.I.I. - " + emisor.Sii, printFont, Brushes.Black, 90, y);
            y = y + 20;
            Font FNTGRANDE = new Font("Arial", (float)8.5, FontStyle.Bold, GraphicsUnit.Point);
            e.Graphics.DrawString(emisor.Razon, FNTGRANDE, Brushes.Black, 1, y);
            /************* imprime membrete *********************/
            y = y + 25;
            e.Graphics.DrawString("Giro  :", font2, Brushes.Black, 5, y);

            if(emisor.Giro_1.Length > 35)
            {
                string linea2 = "";
                e.Graphics.DrawString(emisor.Giro_1.Substring(0,38), printFont, Brushes.Black, 32, y);

                y = y + 10;
                e.Graphics.DrawString("", font2, Brushes.Black, 5, y);
                linea2 = emisor.Giro_1.Substring(38, emisor.Giro_1.Length -38);
                e.Graphics.DrawString(linea2, printFont, Brushes.Black, 32, y);
                
            }
            else
            {

                e.Graphics.DrawString(emisor.Giro_1, printFont, Brushes.Black, 32, y);
            }

            if (emisor.Giro_2 != "")
            {
                y = y + 10;
                e.Graphics.DrawString(emisor.Giro_2, printFont, Brushes.Black, 32, y);
            }
            if (emisor.Giro_3 != "")
            {
                y = y + 10;
                e.Graphics.DrawString(emisor.Giro_3, printFont, Brushes.Black, 32, y);
            }
            if (emisor.Giro_4 != "")
            {
                y = y + 10;
                e.Graphics.DrawString(emisor.Giro_4, printFont, Brushes.Black, 32, y);
            }


            y = y + 12;
            e.Graphics.DrawString("Dirección  :", font2, Brushes.Black, 5, y);
            e.Graphics.DrawString(emisor.Direccion + ", " + emisor.Comuna + ".", font5, Brushes.Black, 55, y + 2);
            y = y + 12;
            e.Graphics.DrawString("Fono  :", font2, Brushes.Black, 5, y);
            e.Graphics.DrawString(emisor.Fono, font5, Brushes.Black, 40, y + 2);
            y = y + 12;
            e.Graphics.DrawString("Email  :", font2, Brushes.Black, 5, y);
            e.Graphics.DrawString(emisor.Email, font5, Brushes.Black, 40, y + 2);
            y = y + 12;
            e.Graphics.DrawString("Cajera :", font2, Brushes.Black, 5, y);
            e.Graphics.DrawString(DOC_NOMBRE_CAJERA , font5, Brushes.Black, 40, y + 2);
            y = y + 12;
            e.Graphics.DrawString("Fecha Emisión  :", font2, Brushes.Black, 5, y);
            e.Graphics.DrawString(dte.Documento.Encabezado.IdentificacionDTE.FechaEmision.ToShortDateString(), font5, Brushes.Black, 77, y + 2);

            e.Graphics.DrawString("Caja  :", font2, Brushes.Black, 122, y);
            e.Graphics.DrawString(DOC_CAJA, font5, Brushes.Black, 152, y + 2);

            /****************************   DATOS RECEPTOR   *****************************************/
            y = y + 5;
            e.Graphics.DrawString("____________________________________________________________", font2, Brushes.Black, x + 1, y);
            y = y + 11;
            e.Graphics.DrawString("Datos Receptor:", font2, Brushes.Black, 5, y);
            y = y + 3;
            e.Graphics.DrawString("____________________________________________________________", font2, Brushes.Black, x + 1, y);

            y = y + 14;
            Font fontEmisor = new Font("Arial", (float)7, FontStyle.Regular, GraphicsUnit.Point);
            e.Graphics.DrawString("Rut :", fontEmisor, Brushes.Black, 5, y);
            e.Graphics.DrawString(dte.Documento.Encabezado.Receptor.Rut, fontEmisor, Brushes.Black, 40, y);
            y = y + 10;
            e.Graphics.DrawString("R.Social:", fontEmisor, Brushes.Black, 5, y);
            e.Graphics.DrawString(dte.Documento.Encabezado.Receptor.RazonSocial, fontEmisor, Brushes.Black, 40, y);
            y = y + 10;
            e.Graphics.DrawString("Giro:", fontEmisor, Brushes.Black, 5, y);
            e.Graphics.DrawString(dte.Documento.Encabezado.Receptor.Giro, fontEmisor, Brushes.Black, 40, y);
            y = y + 10;
            e.Graphics.DrawString("Dirección:", fontEmisor, Brushes.Black, 5, y);
            e.Graphics.DrawString(dte.Documento.Encabezado.Receptor.Direccion + ", " + dte.Documento.Encabezado.Receptor.Comuna, fontEmisor, Brushes.Black, 40, y);
            y = y + 7;




            /*************** REGION QUE IMPRIME LOS DETALLES  ***************/
            Font font3 = new Font("Arial", 6, FontStyle.Bold, GraphicsUnit.Point);
            y = y + 6;
            e.Graphics.DrawString("____________________________________________________________", font3, Brushes.Black, x + 1, y);
            y = y + 8;
            e.Graphics.DrawString("CANT. DESCRIPCION.                                    PRECIO   DESC.   TOTAL", font3, Brushes.Black, x + 2, y);
            y = y + 3;
            e.Graphics.DrawString("____________________________________________________________", font3, Brushes.Black, x + 1, y);

            /*************** RECORRER ROLLO DE LA GRILLA ***************/
            y += 8;
            List<PlaceSoft.DTE.Engine.Documento.Detalle> listDetalles = new List<PlaceSoft.DTE.Engine.Documento.Detalle>();
            listDetalles = dte.Documento.Detalles;

            foreach (PlaceSoft.DTE.Engine.Documento.Detalle detalle in listDetalles)
            {
                List<PlaceSoft.DTE.Engine.Documento.CodigoItem> listCodigo = new List<PlaceSoft.DTE.Engine.Documento.CodigoItem>();
                listCodigo = detalle.CodigosItem;
                string codigo ="";
                foreach (PlaceSoft.DTE.Engine.Documento.CodigoItem fila in listCodigo)
                {
                    codigo = fila.ValorCodigo;
                }


                double cantidad = detalle.Cantidad;
                string descripcion = detalle.Nombre;
                string descripcion2 = "";
                if (detalle.Nombre.Length > 25)
                {
                    descripcion = detalle.Nombre.Substring(0, 25);
                   // descripcion2 = detalle.Nombre.Substring(25, detalle.Nombre.Length - 25);
                }

                string totallinea = string.Format(CultureInfo.CurrentCulture, "{0:C0}", detalle.MontoItem);
                string preciolinea = string.Format(CultureInfo.CurrentCulture, "{0:C0}", detalle.Precio);
                int dctolinea = detalle.Descuento;
                string porDescuento = "0%";
                if (detalle.DescuentoPorcentaje > 0)
                {
                    porDescuento = detalle.DescuentoPorcentaje + "%"; //@String.Format(new CultureInfo("es-CL"), "{0:P2}", detalle.DescuentoPorcentaje);
                }

                int pad = 183 + (6 - totallinea.Length) * 3;
                int pad2 = 133 + (6 - preciolinea.Length) * 3;
                int pad3 = 5 + (3 - cantidad.ToString().Length) * 3;


                e.Graphics.DrawString(cantidad.ToString(), printFoot, Brushes.Black, pad3, y);


                e.Graphics.DrawString(codigo.ToString(), printFoot, Brushes.Black, 23, y);
                e.Graphics.DrawString(descripcion , printFoot, Brushes.Black, 23, y + 6);
               // e.Graphics.DrawString(descripcion2, printFoot, Brushes.Black, 23, y + 6);
                e.Graphics.DrawString(preciolinea, printFoot, Brushes.Black, pad2, y + 6);
                e.Graphics.DrawString(porDescuento, printFoot, Brushes.Black, 162, y + 6);
                e.Graphics.DrawString(totallinea, printFoot, Brushes.Black, pad, y + 6);
                y += 14;
            }

            y = y - 1;


            // y += 3;
            e.Graphics.DrawString("____________________________________________________________", font3, Brushes.Black, x, y);
            Font font4 = new Font("Arial", 7, FontStyle.Bold, GraphicsUnit.Point);
            /***************** REGION PARA PONER LOS ILAS ************************/
            /*************************************  INICIO REGION IMPUESTOS ILAS ******************************/
            string GlosaImpuestoAdicional = "";
            if (dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Count > 0)
            {
                string srtimp = "";
                string totalfinalimp = "";
                int padfinalimp = 0;
                foreach (var imp in dte.Documento.Encabezado.Totales.ImpuestosRetenciones)
                {
                    GlosaImpuestoAdicional = imp.TipoImpuesto.ToString() + " " + imp.TasaImpuesto.ToString();

                    if (imp.TipoImpuesto.ToString() == "Licores")
                    {
                        GlosaImpuestoAdicional = "Imp.Licores   " + imp.TasaImpuesto + "%";
                    }
                    if (imp.TipoImpuesto.ToString() == "Vinos")
                    {
                        GlosaImpuestoAdicional = "Imp.Vinos     " + imp.TasaImpuesto + "%";
                    }
                    if (imp.TipoImpuesto.ToString() == "Cervezas")
                    {
                        GlosaImpuestoAdicional = "Imp.Cervezas  " + imp.TasaImpuesto + "%";
                    }
                    //// 27
                    if (imp.TipoImpuesto.ToString() == "BebidasAnalcoholicasYMinerales")
                    {
                        GlosaImpuestoAdicional = "Imp.NO Azucar " + imp.TasaImpuesto + "%";
                    }
                    ///271
                    if (imp.TipoImpuesto.ToString() == "BebidasAnalcoholicasYMineralesAltaAzucar")
                    {
                        GlosaImpuestoAdicional = "Imp.Azucar    " + imp.TasaImpuesto + "%";
                    }


                    y = y + 10;
                    e.Graphics.DrawString(GlosaImpuestoAdicional, font4, Brushes.Black, 90, y);
                    srtimp = string.Format(CultureInfo.CurrentCulture, "{0:C0}", @String.Format(new CultureInfo("es-CL"), "{0:C0}", imp.MontoImpuesto));
                    totalfinalimp = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtimp.Replace(",00", ""));
                    padfinalimp = 182 + (4 - totalfinalimp.Length) * 3;
                    e.Graphics.DrawString(totalfinalimp, font4, Brushes.Black, padfinalimp, y + 1);

                }
            }




            y = y + 8;

            /****************** fin region ilas ******************************/
            double TOTALEXENTO = dte.Documento.Encabezado.Totales.MontoExento;
            if (TOTALEXENTO > 0)
            {
                y = y + 10;
                e.Graphics.DrawString("Monto Exento :", font4, Brushes.Black, 105, y);
                string srtexento = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALEXENTO);
                string totalfinalex = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtexento.Replace(",00", ""));
                int padfinalex = 182 + (4 - totalfinalex.Length) * 3;
                e.Graphics.DrawString(totalfinalex, font4, Brushes.Black, padfinalex, y + 1);
            }


            double TOTALNETO = dte.Documento.Encabezado.Totales.MontoNeto;
            if (TOTALNETO > 0)
            {
                y = y + 10;
                e.Graphics.DrawString("Monto Neto :", font4, Brushes.Black, 105, y);
                string srtneto = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALNETO);
                string totalfinalneto = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtneto.Replace(",00", ""));
                int padfinalex = 182 + (4 - totalfinalneto.Length) * 3;
                e.Graphics.DrawString(totalfinalneto, font4, Brushes.Black, padfinalex, y + 1);
            }
            double TOTALIVA = dte.Documento.Encabezado.Totales.IVA;
            y = y + 10;
            e.Graphics.DrawString("I.V.A. :" + DOC_TASA_IVA + "%", font4, Brushes.Black, 105, y);
            string srtiva = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALIVA);
            string totalfinaliva = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtiva.Replace(",00", ""));
            int padfinaliva = 182 + (4 - totalfinaliva.Length) * 3;
            e.Graphics.DrawString(totalfinaliva, font4, Brushes.Black, padfinaliva, y + 1);

            double TOTALVENTA = dte.Documento.Encabezado.Totales.MontoTotal;
            y = y + 10;
            e.Graphics.DrawString("Monto Total :", font4, Brushes.Black, 105, y);
            string srtTotal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALVENTA);
            string totalfinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtTotal.Replace(",00", ""));
            int padfinal = 182 + (4 - totalfinal.Length) * 3;
            e.Graphics.DrawString(totalfinal, font4, Brushes.Black, padfinal, y + 1);
            /********************* REGION DEL REDONDEO *************************/
            y = y + 5;
            if (this.DOC_REDONDEO != 0)
            {
                y = y + 22;
                double TOTALREDONDEO = this.DOC_REDONDEO;
                e.Graphics.DrawString("Ley 20.956 :", font2, Brushes.Black, 5, y);
                string srtRedondeo = string.Format(CultureInfo.CurrentCulture, "{0:C0}", TOTALREDONDEO);
                string redondeofinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtRedondeo.Replace(",00", ""));
                e.Graphics.DrawString(redondeofinal, font2, Brushes.Black, 180, y + 1);

                /// A PAGAR /////
                double APAGAR = DOC_TOTAL_A_PAGAR;
                y = y + 10;
                e.Graphics.DrawString("Total a Pagar :", font4, Brushes.Black, 105, y);
                string srtApagar = string.Format(CultureInfo.CurrentCulture, "{0:C0}", APAGAR);
                string apagarfinal = string.Format(CultureInfo.CurrentCulture, "{0:C0}", srtApagar.Replace(",00", ""));
                int padapagar = 182 + (4 - apagarfinal.Length) * 3;
                e.Graphics.DrawString(apagarfinal, font4, Brushes.Black, padapagar, y + 1);

            }
            /******************************************************************/


            y = y + 6;
            e.Graphics.DrawString("____________________________________________________________", font3, Brushes.Black, x, y);
            if (DOC_TIPO_DTE != "61" && DOC_TIPO_DTE != "52")
            {
                y = y + 13;
                List<string> strFormaPago = DOC_FORMA_PAGO;
                e.Graphics.DrawString("Forma de Pago :", font2, Brushes.Black, 5, y);
                y = y + 3;
                foreach(string row in strFormaPago) {

                    e.Graphics.DrawString(row.ToString(), font5, Brushes.Black, 77, y);
                    y = y + 10;
                }
                y = y + 15;
            }

            /****************** IMPRIME COPIA CEDIBLE *************/

            if (dte.Documento.Referencias.Count > 0)
            {
                y = y + 10;
                Rectangle rect3 = new Rectangle(5, (int)y, 197, 50);
                e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rect3));
                y = y + 1;
                e.Graphics.DrawString("REFERENCIAS ", font2, Brushes.Black, 76, y);
                y = y + 10;
                e.Graphics.DrawString("TIPO DOCUMENTO                  NUMERO            FECHA", font4, Brushes.Black, 5, y);
                y = y + 9;
                int lineaRef = 1;
                string refe = "";
                string refe_tipo = "";

                
                    foreach (var referencia in dte.Documento.Referencias)
                    {
                        if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.FacturaElectronica)
                        {
                            refe_tipo = "FAC. ELECTRÓNICA(33)";
                        }
                        if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.NotaCreditoElectronica)
                        {
                            refe_tipo = "N.C. ELECTRÓNICA(61)";
                        }
                        if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.NotaDebitoElectronica)
                        {
                            refe_tipo = "N.D. ELECTRÓNICA(56)";
                        }
                        if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.BoletaElectronica)
                        {
                            refe_tipo = "BOLETA ELECTRÓNICA(39)";
                        }
                        if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.OrdenCompra)
                        {
                            refe_tipo = "ORDEN DE COMPRA(801)";
                        }
                        if (refe_tipo != "")
                        {

                            e.Graphics.DrawString(refe_tipo, printFoot, Brushes.Black, 6, y);
                            e.Graphics.DrawString(referencia.FolioReferencia.ToString().PadLeft(10, Convert.ToChar("0")), printFoot, Brushes.Black, 108, y);
                            e.Graphics.DrawString(referencia.FechaDocumentoReferencia.ToShortDateString(), printFoot, Brushes.Black, 165, y);
                       
                            y = y + 7;
                        }
                        /******************* IMPRIME LA RAZON DE LA REFERENCIA EN EL CAJON DE ABAJO *********************/
                        refe = referencia.RazonReferencia;
                        if (refe.Length > 55)
                        {
                            refe = referencia.RazonReferencia.Substring(0, 55);
                        }
                        //ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(lineaRef + "- " + refe, fontBlack), (float)234, (float)LINEA - 35, 0);
                        y = y + 5;
                        e.Graphics.DrawString(lineaRef + "- " + refe, font5, Brushes.Black, 6, y + 7);
                        lineaRef++;
                    }
            }

            y = y + 25;
            if (Observaciones != "")
            {
                y = y + 10;
                Rectangle rect3 = new Rectangle(5, (int)y, 197, 50);
                e.Graphics.DrawRectangle(Pens.Black, Rectangle.Round(rect3));
                y = y + 1;
                e.Graphics.DrawString("OBSEVACIONES ", font2, Brushes.Black, 76, y);
                y = y + 10;
                //e.Graphics.DrawString("TIPO DOCUMENTO                  NUMERO            FECHA", font4, Brushes.Black, 5, y);
                //y = y + 9;
                int lineaRef = 1;
                string refe = "";
                string refe_tipo = "";
                Observaciones = "POR CONGRESO CIRUGIA SEGUN CONVENIO";

                //foreach (var referencia in dte.Documento.Referencias)
                //{
                //    if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.FacturaElectronica)
                //    {
                //        refe_tipo = "FAC. ELECTRÓNICA(33)";
                //    }
                //    if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.NotaCreditoElectronica)
                //    {
                //        refe_tipo = "N.C. ELECTRÓNICA(61)";
                //    }
                //    if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.NotaDebitoElectronica)
                //    {
                //        refe_tipo = "N.D. ELECTRÓNICA(56)";
                //    }
                //    if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.BoletaElectronica)
                //    {
                //        refe_tipo = "BOLETA ELECTRÓNICA(39)";
                //    }
                //    if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.OrdenCompra)
                //    {
                //        refe_tipo = "ORDEN DE COMPRA(801)";
                //    }
                //    if (refe_tipo != "")
                //    {

                //        e.Graphics.DrawString(refe_tipo, printFoot, Brushes.Black, 6, y);
                //        e.Graphics.DrawString(referencia.FolioReferencia.ToString().PadLeft(10, Convert.ToChar("0")), printFoot, Brushes.Black, 108, y);
                //        e.Graphics.DrawString(referencia.FechaDocumentoReferencia.ToShortDateString(), printFoot, Brushes.Black, 165, y);

                //        y = y + 7;
                //    }
                    /******************* IMPRIME LA RAZON DE LA REFERENCIA EN EL CAJON DE ABAJO *********************/
                 //   refe = referencia.RazonReferencia;
                    if (Observaciones.Length > 55)
                    {
                        refe = Observaciones.Substring(0, 55);
                    }
                    //ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(lineaRef + "- " + refe, fontBlack), (float)234, (float)LINEA - 35, 0);
                    y = y + 5;
                    e.Graphics.DrawString(lineaRef + "- " + Observaciones, font5, Brushes.Black, 6, y + 7);
                    lineaRef++;
                //}
            }


            y = y + 35;
            if (DOC_REIMPRESO == false)
            {
                e.Graphics.DrawString("Hora Reimpresión: " + DateTime.Now.ToString("HH:mm:ss"), font5, Brushes.Black, 65, y + 2);
            }
            else
            {
                e.Graphics.DrawString("Documento Reimpreso: " + DateTime.Now.ToString("HH:mm:ss"), font5, Brushes.Black, 5, y + 2);
            }

            y = y + 25;
            if (DOC_TIPO_DTE == "52")
            {
                //if(DOC_TIPO_TRASLADO == 1)
                //{
                //    e.Graphics.DrawString("1 O" , font5, Brushes.Black, 5, y + 2);
                //}
                if (DOC_TIPO_TRASLADO == 6)
                {
                    e.Graphics.DrawString("           6 Operación no constituye venta", font5, Brushes.Black, 5, y + 2);
                }
            }






            y = y + 14;

            e.Graphics.DrawImage(pictureBoxTimbre.Image, 5, y, pictureBoxTimbre.Width * (float)0.6, pictureBoxTimbre.Height * (float)0.8);

            y = y + 90;
            Font printFont2 = new Font("Arial", (float)6, FontStyle.Bold, GraphicsUnit.Point);
            //
            e.Graphics.DrawString("Timbre Electrónico S.I.I.", printFont2, Brushes.Black, 68, y);
            y = y + 10;
            e.Graphics.DrawString(emisor.Glosa_res + " Verifique su documento en: " + emisor.Web_verificacion, printFoot, Brushes.Black, 20, y);
            //y = y + 7;
            //e.Graphics.DrawString(emisor.Web_verificacion, printFoot, Brushes.Black, 63, y);
            /****************
            *      AQUI PONER LAS REFERENCIAS DE LA NOTA DE CRÉDITO SI FUESE EL CASO
            *      E IMPRIMIR LA COPIA CEDIBLE
            * 
            * ***************/
            if (IMPRIME_CEDIBLE == true)
            {
                y = y + 15;
                Rectangle rect4 = new Rectangle(5, (int)y, 197, 12);
                Pen pen = new Pen(Color.Black, 2);
                e.Graphics.DrawRectangle(pen, Rectangle.Round(rect4));
                e.Graphics.DrawString("Acuse de Recibo ", font2, Brushes.Black, 72, y + 1);
                y = y + 26;
                e.Graphics.DrawString("Rut:______________________________________", font2, Brushes.Black, 5, y);
                y = y + 22;
                e.Graphics.DrawString("Nombre:___________________________________", font2, Brushes.Black, 5, y);
                y = y + 22;
                e.Graphics.DrawString("Fecha:____________________________________", font2, Brushes.Black, 5, y);
                y = y + 26;
                e.Graphics.DrawString("Firma:____________________________________", font2, Brushes.Black, 5, y);
                y = y + 18;
                e.Graphics.DrawString("El acuse de recibo que se declara en este acto, de acuerdo a lo ", printFoot, Brushes.Black, 5, y);
                y = y + 6;
                e.Graphics.DrawString("dispueto en la letra b) del Art. 4° y letra c) del Art 5° de  ", printFoot, Brushes.Black, 5, y);
                y = y + 6;
                e.Graphics.DrawString("la Ley 19.983, acredita que la entrega de las mercaderías o ", printFoot, Brushes.Black, 5, y);
                y = y + 6;
                e.Graphics.DrawString("servicio(s) prestado(s) ha(n) sido recibido(s)", printFoot, Brushes.Black, 5, y);

                y = y + 5;
                e.Graphics.DrawString("__________________________________________", font2, Brushes.Black, 5, y);
                y = y + 12;
                e.Graphics.DrawString("COPIA CEDIBLE", FNTGRANDE, Brushes.Black, 120, y);
            }

            y = y + 20;
            //e.Graphics.DrawString("Contrate ya su documentación electrónica en http://www.placesoft.cl/", printFoot, Brushes.Black, 5, y);
            e.HasMorePages = false;
        }

        public void roundCorners(Form obj)
        {
            obj.FormBorderStyle = FormBorderStyle.None;
            // obj.BackColor = Color.Cyan


            System.Drawing.Drawing2D.GraphicsPath DGP = new System.Drawing.Drawing2D.GraphicsPath();
            DGP.StartFigure();
            // top left corner
            DGP.AddArc(new Rectangle(0, 0, 40, 40), 180, 90);
            DGP.AddLine(40, 0, obj.Width - 40, 0);

            // top right corner
            DGP.AddArc(new Rectangle(obj.Width - 40, 0, 40, 40), -90, 90);
            DGP.AddLine(obj.Width, 40, obj.Width, obj.Height - 40);

            // buttom right corner
            DGP.AddArc(new Rectangle(obj.Width - 40, obj.Height - 40, 40, 40), 0, 90);
            DGP.AddLine(obj.Width - 40, obj.Height, 40, obj.Height);

            // buttom left corner
            DGP.AddArc(new Rectangle(0, obj.Height - 40, 40, 40), 90, 90);
            DGP.CloseFigure();

            obj.Region = new Region(DGP);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            try
            {
                if (CONTADOR == 0)
                {
                    lblInformacion.Text = "IMPRIMIENDO FACTURA  ELECTRÓNICA Nº " + DOC_FOLIOSII.PadLeft(10,Convert.ToChar("0"));
                    lblInformacion.Refresh();
                    CONTADOR++;
                    ImprimirDTE();
                    System.Threading.Thread.Sleep(3000);
                    this.Close();

                }
            }
            catch (Exception ex)
            {
                log.Error("El error es: ", ex);
                this.Close();
            }
        }
    }
}
