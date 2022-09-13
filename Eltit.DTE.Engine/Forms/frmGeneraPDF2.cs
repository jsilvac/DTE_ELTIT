using PlaceSoft.DTE.Engine.Documento;
using PlaceSoft.DTE.Engine.XML;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
//using SamplesDTE.Clases;
using System.IO;
using log4net;
using MySql.Data;
using MySql.Data.MySqlClient;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Globalization;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using PdfSharp.Pdf.Printing;
using Eltit.DTE.clases;
using PlaceSoftDTE.clases;
using System.Xml;

namespace PlaceDTE
{
    public partial class frmGeneraPDF2 : Telerik.WinControls.UI.RadForm
    {
        public string OBSERVACION_VENTA = "";
        private string _CURR_COMPANY;
        private static readonly ILog log =
          LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        string PDFPathCedible = "";
        string PDFPathTributario = "";

        iTextSharp.text.Image jpg;
        public string DOC_SERVIDOR = "";
        public string DOC_PREFIJO = "";
        public string DOC_EMPRESA_CONTABLE = "";
        public string IMPRESORA_FACTURA = "";
        public string DOC_FOLIOSII = "";
        public string DOC_TIPOSII = "";
        public string DOC_LOCAL = "";
        public string DOC_NUMERO = "";
        public string DOC_CAJA = "";
        public string DOC_CLIENTE = "";
        public string DOC_TIPO = "";
        public string DOC_FONO = "";
        public string DOC_VENDEDOR = "";
        public bool ENVIAR_IMPRIMIR = false;
        public string DOC_RUT_BASE = "";
        public string DOC_XML;
        public string DOC_CONTACTO = "";
        public string DOC_CORREO_CONTACTO = "";
        public bool SOLOENVIO = false;
        public string destinoDTE1 = "";
        public string destinoDTE2 = "";

        // Cabecera PDF
        //public string DOC_RAZONSOCIAL = "";
        //public string DOC_GIRO = "";
        //public string DOC_CMATRIZ = "";
        //public string DOC_COMUNA = "";
        //public string DOC_SUCURSALSII = "";

        // DOC_FONO

        public PlaceSoft.DTE.Engine.Documento.DTE dte;
        public DatosEmisor emisor;
        public string _BASE_FOLDER_PROD = "";


        [DllImport("shell32.dll")]




        private static extern long FindExecutable(string lpFile, string lpDirectory, [Out] StringBuilder lpResult);

        public frmGeneraPDF2()
        {
            InitializeComponent();
        }

        private void frmGeneraPDF2_Load(object sender, EventArgs e)
        {

            this.InicializaControlesDeEmpresa();
            log.Debug(" InicializaControlesDeEmpresa()");

            _BASE_FOLDER_PROD = @"C:\PlaceDTE\eltit\"+ DOC_RUT_BASE.Substring(0,8) +@"\Produccion";

            radGroupBox1.GroupBoxElement.Header.Font = new System.Drawing.Font("Arial", 6);

            //string rut = lblRut.Text.Substring(0, 9);
            //rut = Convert.ToDouble(rut).ToString();

            PDFPathCedible = @"C:\PlaceDTE\eltit\comun\plantillas\eltit-cedible.pdf";
            PDFPathTributario = @"C:\PlaceDTE\eltit\comun\plantillas\eltit-tributario.pdf";

            generaDTE();

           // this.btnGenerar_Click(null, null);
           // System.Threading.Thread.Sleep(2000);
            
        }

        private void generaDTE()
        {
            dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(DOC_XML);
            //this.CapturaGiro();
            pictureBoxTimbre.Image = dte.Documento.TimbrePDF417;
            emisor = new DatosEmisor(DOC_RUT_BASE, "eltit_","192.168.4.9", DOC_LOCAL);



            this.btnGenerar_Click(null, null);
        }

        private void InicializaControlesDeEmpresa()
        {
            //lblRut.Text = FuncionesClass.G_EMPRESARUT;
            //lblNombreEmpresa.Text = FuncionesClass.G_EMPRESANOMBRE;
            //lblDireccion.Text = FuncionesClass.G_EMPRESADIRECCION;
            //lblComuna.Text = FuncionesClass.G_EMPRESACOMUNA;
            //lblCiudad.Text = FuncionesClass.G_EMPRESACIUDAD;

        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            //this.CargaTimbraje();
        }
        public void btnGenerar_Click(object sender, EventArgs e)
        {
            

       try
            {
                if(1==1)
                {
                         destinoDTE1 = _BASE_FOLDER_PROD + @"\pdf\" + DOC_LOCAL + @"\DTE" + DOC_TIPOSII + "F" + Convert.ToInt32(DOC_FOLIOSII) + ".pdf";
                         destinoDTE2 = _BASE_FOLDER_PROD + @"\pdf\" + DOC_LOCAL + @"\DTE" + DOC_TIPOSII + "F" + Convert.ToInt32(DOC_FOLIOSII) + "-Cedible.pdf";

                    jpg = iTextSharp.text.Image.GetInstance(dte.Documento.TimbrePDF417, System.Drawing.Imaging.ImageFormat.Jpeg);

                    if (!File.Exists(destinoDTE1))
                    {
                        this.CreaPDF(this.PDFPathTributario, destinoDTE1, this.jpg, this.dte);
                        log.Debug("-Generando " + destinoDTE1);
                        //ENVIAR_IMPRIMIR = true;
                    }
                    else
                    {
                        try
                        {
                            File.Delete(destinoDTE1);                        
                            System.Threading.Thread.Sleep(1000);
                            this.CreaPDF(this.PDFPathTributario, destinoDTE1, this.jpg, this.dte);
                            log.Debug("-Generando " + destinoDTE1);
                            //ENVIAR_IMPRIMIR = true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Archivo no existe...");
                        }
                    }

                    if (!File.Exists(destinoDTE2))
                    {
                        this.CreaPDF(this.PDFPathCedible, destinoDTE2, this.jpg, this.dte);
                        log.Debug("-Generando " + destinoDTE2);
                        //ENVIAR_IMPRIMIR = true;
                    }
                    else
                    {
                        try
                        {
                            File.Delete(destinoDTE2);
                            System.Threading.Thread.Sleep(1000);
                            this.CreaPDF(this.PDFPathCedible, destinoDTE2, this.jpg, this.dte);
                            log.Debug("-Generando " + destinoDTE2);
                            //ENVIAR_IMPRIMIR = true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Archivo no existe...");
                        }
                    }

                    if(ENVIAR_IMPRIMIR == true )
                    {
                        //this.ActualizaImpreso(txtLocalActivo.Text, lbTipo.Text, lblFolio.Text);
                        System.Threading.Thread.Sleep(1000);
                        if (chkImprimeDirecto.Checked == true)
                        {
                                if (lbTipo.Text == "61")
                                {
                                    this.SendToPrinter(destinoDTE1, 1);
                                    log.Debug("-IMPRIMIENDO IMPRIMIR_PDF_DTE " + destinoDTE2);
                                }
                                else
                                {
                                    this.SendToPrinter(destinoDTE1, 1);
                                    System.Threading.Thread.Sleep(3000);
                                    this.SendToPrinter(destinoDTE2, 1);
                                    log.Debug("-IMPRIMIENDO IMPRIMIR_PDF_DTE " + destinoDTE2);
                                }      
                        }
                    }

                }          

            }           
            catch (Exception ex)
            {
                log.Debug(destinoDTE1 + "," + destinoDTE2);
                log.Error("Error:", ex);
               // RadMessageBox.Show(this, "Error:" + ex.Message.ToString(), "Atencion", MessageBoxButtons.OK);
            }

            this.Close();
        }
        private void ActualizaImpreso(string xLocal,string xTipo, string xCaf)
        {
            //VentasClass ve = new VentasClass(FuncionesClass.G_SERVIDOR);
            //ve.setBaseDTE(FuncionesClass.BASE_DTE);
            //int ok = ve.ActualizaImpresa(xLocal, xTipo, xCaf, "1");
        }
        private void SendToPrinter(string xPath, int xTry)
        {

            if(IMPRESORA_FACTURA != "" && xTry != 2 )
            {
                PrinterSettings settings = new PrinterSettings();
                string defaultPrinterName = settings.PrinterName;

                if (defaultPrinterName != IMPRESORA_FACTURA)
                {
                    SetDefaultPrinter(IMPRESORA_FACTURA);
                    log.Debug("SET FACTURA POR DEFECTO " + IMPRESORA_FACTURA);
                }
            }
            else
            {
                PrinterSettings settings = new PrinterSettings();
                string defaultPrinterName = settings.PrinterName;
                //SetDefaultPrinter(IMPRESORA_FACTURA);
            }
           
            /***********************************************+
             * 
             * ESTA APLICACION NO IMPRIME EN WINDOWS 10 SI
             * EL PROGRAMA ESTA EJECUTADA COMO ADMINISTRADOR 
             * 
             * ***********************************************/
            if (xTry == 1)
            {
                try { 
                ProcessStartInfo info = new ProcessStartInfo();
                info.Verb = "print";
                info.FileName = xPath;
                info.CreateNoWindow = true;
 
                Process p = new Process();

                p.StartInfo = info;
                p.Start();

                p.WaitForInputIdle();
                System.Threading.Thread.Sleep(3000);
                    log.Debug("Imprimiendo en " + IMPRESORA_FACTURA + " op: " + xTry);
                if (false == p.CloseMainWindow())
                    p.Kill();
                    log.Debug("Matando Proceso " + IMPRESORA_FACTURA + " op: " + xTry);

                }
                catch (Exception ex)
                {                    
                    log.Error("Error: " , ex);
                }
            }

            if(xTry == 2)
            {
                if(Environment.Is64BitOperatingSystem)
                {
                    PdfFilePrinter.AdobeReaderPath = @"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe";
                    
                }
                else
                {
                    PdfFilePrinter.AdobeReaderPath = @"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe";
                    log.Debug("- try 2 impresion " + PdfFilePrinter.AdobeReaderPath);
                }

                // Set the file to print and the Windows name of the printer. C:\Program Files\Adobe\Acrobat Reader DC\Reader
                // At my home office I have an old Laserjet 6L under my desk.
                PdfFilePrinter printer = new PdfFilePrinter(xPath, "激光");

                printer.Print();
            }
           


            if (xTry == 5)
            {

                PrintDialog pd = new PrintDialog();

                pd.PrinterSettings.PrinterName = IMPRESORA_FACTURA;
                ProcessStartInfo info = new ProcessStartInfo();
                info.Verb = "print";
                info.FileName = xPath;
                info.CreateNoWindow = true;
              //  info.WindowStyle = ProcessWindowStyle.Hidden;

                Process p = new Process();
                p.StartInfo.Arguments = pd.PrinterSettings.PrinterName.ToString();
                pd.PrinterSettings.MaximumPage = 1;
                p.StartInfo = info;
                p.Start();
                p.WaitForInputIdle();
              //  System.Threading.Thread.Sleep(3000);
                if (false == p.CloseMainWindow())
                    p.Kill();

                log.Debug("Imprimiendo en " + IMPRESORA_FACTURA + " opcion " + xTry);
            }
                
            if (xTry == 6)
            {
                Process proc = new Process();
                proc.StartInfo.FileName = @"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe";
                proc.StartInfo.Arguments = @" /t /h " + "\"" + xPath + "\"" + " " + "\"" + IMPRESORA_FACTURA + "\"";
                proc.StartInfo.UseShellExecute = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                System.Threading.Thread.Sleep(1000);
                proc.WaitForInputIdle();
                log.Debug("Imprimiendo en " + IMPRESORA_FACTURA + " opcion " + xTry);
                proc.Kill();
            }


        }

        public static void SetDefaultPrinter(string printername)
        {
            var type = Type.GetTypeFromProgID("WScript.Network");
            var instance = Activator.CreateInstance(type);
            type.InvokeMember("SetDefaultPrinter", System.Reflection.BindingFlags.InvokeMethod, null, instance, new object[] { printername });
        }
        public void CargaTimbraje(string XML)
        {
            try
            {
                string pathxml = "";

               // txtFilePath.Text = xPathXML; //openFileDialog1.FileName;

                string xmlString = "";
                pathxml = @"C:\PlaceDTE\" + DOC_PREFIJO.Replace("_","") + @"\"+ DOC_RUT_BASE + @"\Produccion\xml\" + DOC_LOCAL + @"\DTE" + DOC_TIPOSII + "F" + DOC_FOLIOSII + ".xml";

                if (!File.Exists(pathxml))
                {
                    FileStream fst;
                    BinaryWriter bw;
                    string tmp_path = @"C:\temp\" + DateTime.Now.Ticks + ".xml";

                    fst = new FileStream(tmp_path, FileMode.OpenOrCreate, FileAccess.Write);
                    bw = new BinaryWriter(fst);
                    string strxml = XML.Replace("±	", "");
                    Encoding ByteConverter = Encoding.GetEncoding("ISO-8859-1");
                    byte[] textEnBytes = ByteConverter.GetBytes(strxml);

                    bw.Write(textEnBytes);
                    bw.Flush();
                    bw.Close();
                    bw.Dispose();

                    using (FileStream fs = new FileStream(pathxml, FileMode.Create, FileAccess.Write))
                    {
                        var xml2 = File.ReadAllBytes(tmp_path);
                        xmlString = File.ReadAllText(tmp_path, Encoding.GetEncoding("ISO-8859-1"));

                        fs.Write(xml2, 0, xml2.Length);
                        fs.Flush();
                        fs.Close();
                    }

                    System.Threading.Thread.Sleep(500);

                }





                System.Drawing.Point textPoint = new System.Drawing.Point(50, 100);
                    string pathFile = pathxml; //openFileDialog1.FileName;
                    string xml = File.ReadAllText(pathFile, Encoding.GetEncoding("ISO-8859-1"));
                    dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);

                     jpg = iTextSharp.text.Image.GetInstance(dte.Documento.TimbrePDF417, System.Drawing.Imaging.ImageFormat.Jpeg);
                    pictureBoxTimbre.Image = dte.Documento.TimbrePDF417;
                
                    lblFecha.Text = dte.Documento.Encabezado.IdentificacionDTE.FechaEmision.ToShortDateString();
                    lblFolio.Text = dte.Documento.Encabezado.IdentificacionDTE.Folio.ToString();
                    lblMonto.Text = dte.Documento.Encabezado.Totales.MontoTotal.ToString("N0");

                    string tipo = string.Empty;
                    switch (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE)
                    {
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaCompraElectronica:
                            tipo = "FACTURA DE COMPRA ELECTRÓNICA";
                            lbTipo.Text = "46";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronica:
                            tipo = "FACTURA ELECTRÓNICA";
                            lbTipo.Text = "33";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronicaExenta:
                            tipo = "FACTURA ELECTRÓNICA EXENTA";
                            lbTipo.Text = "34";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica:
                            tipo = "GUIA DE DESPACHO ELECTRÓNICA";
                            lbTipo.Text = "52";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica:
                            tipo = "NOTA DE CRÉDITO ELECTRÓNICA";
                            lbTipo.Text = "61";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaDebitoElectronica:
                            tipo = "NOTA DE DÉBITO ELECTRÓNICA";
                            lbTipo.Text = "56";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica:
                            tipo = "BOLETA ELECTRÓNICA";
                            lbTipo.Text = "39";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta:
                            tipo = "BOLETA ELECTRÓNICA EXENTA";
                            break;
                    }
                    lblNombreDocumento.Text = tipo;
                   
                
            }
            catch(Exception ex)
            {
                log.Error("Error:", ex);
                RadMessageBox.Show(this, "Error:" + ex.Message.ToString(), "Atencion", MessageBoxButtons.OK);
            }
           
        }

        private void CreaPDF(string orige, string destino, iTextSharp.text.Image jpg,
                               PlaceSoft.DTE.Engine.Documento.DTE dte)
        {

            using (Stream inputPdfStream = new FileStream(orige, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (Stream outputPdfStream = new FileStream(destino, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                var reader = new PdfReader(inputPdfStream);
                var stamper = new PdfStamper(reader, outputPdfStream);
                var pdfContentByte = stamper.GetOverContent(1);
                /**************** MODIFICA DATOS DE LA CABEZA ***********************/
                BaseFont FontNroytipo = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.WINANSI, BaseFont.EMBEDDED);
                BaseFont FontNormal = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
                BaseFont FontItalic = BaseFont.CreateFont(BaseFont.TIMES_ITALIC, BaseFont.WINANSI, BaseFont.EMBEDDED);
                BaseFont FonstStrong = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.WINANSI, BaseFont.EMBEDDED);
                BaseFont FonstBaseSii = BaseFont.CreateFont(BaseFont.COURIER_BOLD, BaseFont.WINANSI, BaseFont.EMBEDDED);
                BaseColor bColorRed = new BaseColor(Color.Red);
                BaseColor bColorBlack = new BaseColor(Color.Black);
                BaseColor bColorRazonSocial = new BaseColor(Color.DarkBlue);

                iTextSharp.text.Font fontRed = new iTextSharp.text.Font(FontNroytipo, 9, 0, bColorRed);
                iTextSharp.text.Font fontBlack = new iTextSharp.text.Font(FontNormal, 9, 0, bColorBlack);
                iTextSharp.text.Font fontSii = new iTextSharp.text.Font(FonstBaseSii, 9, 0, bColorRed);
                iTextSharp.text.Font fontBlackUltra = new iTextSharp.text.Font(FonstStrong, 9, 0, bColorRazonSocial);
                iTextSharp.text.Font fontTotales = new iTextSharp.text.Font(FontNroytipo, 9, 0, bColorBlack);
                iTextSharp.text.Font fontdescripcion = new iTextSharp.text.Font(FontItalic,6, 0, bColorBlack); // new Font(bfTimes, 12, Font.ITALIC, Color.RED)
                fontRed.Size = 9;
                fontBlack.Size = 8;
                fontBlackUltra.Size = 11;
                fontSii.Size = 11;

                iTextSharp.text.Font fontNumero = new iTextSharp.text.Font(FontNroytipo, 9, 0, bColorRed);
                fontNumero.Size = (float)11.5;

                // TIPO DE DOCUMENTO
                if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronica)
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("FACTURA ELECTRÓNICA", fontRed), 485f, 785f, 0);
                }
                if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronicaExenta)
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("FACTURA NO AFECTA O EXENTA", fontRed), 485f, 785f, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("ELECTRÓNICA", fontRed), 490f, 775f, 0);
                }
                if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica)
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("NOTA DE CRÉDITO", fontRed), 490f, 785f, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("ELECTRÓNICA", fontRed), 490f, 775f, 0);
                }
                if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaDebitoElectronica)
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("NOTA DE DÉBITO", fontRed), 490f, 785f, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("ELECTRÓNICA", fontRed), 490f, 775f, 0);
                }
                if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica)
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("GUÍA DE DESPACHO", fontRed), 490f, 785f, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("ELECTRÓNICA", fontRed), 490f, 775f, 0);
                }
                ///////////////////////////////////// sector encabezado emisor///////////////////////////
                ///
                float espacio =0;

                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(emisor.Razon, fontBlackUltra ), (float)95, 818f, 0);

                // GIRO EMISOR
                if (emisor.Giro_1.Length>50)
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("GIRO: " , fontTotales), (float)95, 791f, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase( emisor.Giro_1.Substring(0,50), fontBlack), (float)124, 791f, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase( emisor.Giro_1.Substring(50, emisor.Giro_1.Length-50), fontBlack), (float)121, 781f, 0);
                    espacio = 8;
                }
                else
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("GIRO: ", fontTotales), (float)95, 791f, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(emisor.Giro_1, fontBlack), (float)124, 791f, 0);
                }

                ///// direccion emisor

                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("DIRECCION: ", fontTotales), (float)95, 776f - espacio, 0);
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(emisor.Direccion, fontBlack), (float)153, 776f - espacio, 0);
               // espacio = espacio + 15;
                //////// direcciones sucursales

                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("SUCURSALES: ", fontTotales), (float)95, 761f - espacio, 0);
                int y=0;
                string direcionesX = "";
                foreach (string xdireccion in emisor.Direcciones)
                {
                    direcionesX = direcionesX + " - " + xdireccion;
                    
                }
                direcionesX = direcionesX.Substring(2, direcionesX.Length - 2);
                if (direcionesX.Length > 48)
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(direcionesX.Substring(0, 48), fontBlack), (float)161, 761f -espacio , 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(direcionesX.Substring(48, direcionesX.Length - 48), fontBlack), (float)161, 751f - espacio , 0);
                 
                }
                else
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(direcionesX, fontBlack), (float)158, 761f - y, 0);
                 
                }
                espacio = espacio + 8;


                ///// TELEFONO
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("TELEFONO: ", fontTotales), (float)95, 746f - espacio, 0);
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("452-379500", fontBlack), (float)153, 746f - espacio, 0);

                ///////////////////////////////// fin emisor//////////////////////////////////////////////////
                ///

                // RUT EMISOR
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase(emisor.RutFormat, fontSii), 504f, 804f, 0);

                // CODIGO SII
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase(emisor.Sii, fontSii), 498f, 732f, 0);


                // FOLIO DOCUMENTO
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_CENTER, new Phrase("N° " + dte.Documento.Encabezado.IdentificacionDTE.Folio.ToString().PadLeft(10, Convert.ToChar("0")), fontNumero), 489f, 760f, 0);
                // RAZON SOCIAL
                
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(dte.Documento.Encabezado.Receptor.RazonSocial, fontBlack), (float)85, (float)688, 0);
                // DIRECCION

                string sector = this.GetSectorCliente(dte.Documento.Encabezado.Receptor.Rut.Replace("-", "").PadLeft(10,Convert.ToChar("0")) );

                if(sector != "")
                {
                    sector = " / " + sector;
                }
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(dte.Documento.Encabezado.Receptor.Direccion + sector, fontBlack), (float)85, (float)673, 0);
                // GIRO
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(dte.Documento.Encabezado.Receptor.Giro, fontBlack), (float)85, (float)658, 0);
                //// COMUNA
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(dte.Documento.Encabezado.Receptor.Comuna, fontBlack), (float)85, (float)643, 0);
                //// CIUDAD
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(dte.Documento.Encabezado.Receptor.Ciudad, fontBlack), (float)265, (float)643, 0);
                //// RUT
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(dte.Documento.Encabezado.Receptor.Rut, fontBlack), (float)455, (float)688, 0);
                //// fecha emision
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(dte.Documento.Encabezado.IdentificacionDTE.FechaEmision.ToShortDateString(), fontBlack), (float)455, (float)673, 0);
                //// FORMA DE PAGO
                ///
                string formaPago = "";
                if (dte.Documento.Encabezado.IdentificacionDTE.FormaPago == PlaceSoft.DTE.Engine.Enum.FormaPago.FormaPagoEnum.Contado)
                {
                    formaPago = "CONTADO";
                } else
                {
                    formaPago = "CREDITO";
                }

                //// nueva region para mostrar la forma de pago ////////////////
                //VentasClass ve = new VentasClass();
                
                string medio = "";
                //if(formaPago == "CONTADO")
                //{
                //    medio = ve.LeeFormaPago(DOC_LOCAL, DOC_TIPO, DOC_NUMERO, DOC_CAJA, DOC_CLIENTE.Replace("-", "").PadLeft(10, Convert.ToChar("0")));
                //    formaPago = formaPago + " / " + medio;
                //}
                //else
                //{
                //    medio = ve.LeeFormaPago(DOC_LOCAL, DOC_TIPO, DOC_NUMERO, DOC_CAJA, DOC_CLIENTE.Replace("-", "").PadLeft(10, Convert.ToChar("0")));
                //    formaPago = "CREDITO";
                //    if (medio == "CONTRA ENTREGA")
                //    {
                //        formaPago = medio;
                //    }

                   
                //}

                if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica)
                {
                    formaPago = "-";
                }
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(formaPago, fontBlack), (float)455, (float)658, 0);
                //// FECHA VENCIMIENTO
                string xfechaVEnce = dte.Documento.Encabezado.IdentificacionDTE.FechaVencimiento.ToShortDateString();
                if (xfechaVEnce == "01-01-0001" || dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica)
                {
                    xfechaVEnce = "-";
                }
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(xfechaVEnce, fontBlack), (float)455, (float)643, 0);
                
                
                /////// INFORMACION DE CONTACTO 
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(this.DOC_CONTACTO, fontBlack), (float)85, (float)605, 0);
                // FONO
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(this.DOC_FONO, fontBlack), (float)265, (float)605, 0);
                // VENDEDOR
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(this.DOC_CORREO_CONTACTO, fontBlack), (float)430, (float)605, 0);


                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("__", fontBlack), (float)1, (float)420, 0);
                //************************************* CICLO QUE AGREGA LOS DETALLES ****************************************************//
                fontBlack.Size = 7;
                int LINEA = 565;
                int CANT_DETALLE = dte.Documento.Detalles.Count;
                List<PlaceSoft.DTE.Engine.Documento.Detalle> listDetalles = new List<PlaceSoft.DTE.Engine.Documento.Detalle>();
                listDetalles = dte.Documento.Detalles;
                foreach (PlaceSoft.DTE.Engine.Documento.Detalle detalle in listDetalles)
                {
                    List<PlaceSoft.DTE.Engine.Documento.CodigoItem> codigos = new List<PlaceSoft.DTE.Engine.Documento.CodigoItem>();
                    codigos = detalle.CodigosItem;
                    string barra = "";
                    if (codigos.Count > 0)
                    {
                        foreach (PlaceSoft.DTE.Engine.Documento.CodigoItem codigo in codigos)
                        {
                            barra = codigo.ValorCodigo;
                        }
                    }
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(barra, fontBlack), (float)27, (float)LINEA, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(detalle.Nombre, fontBlack), (float)90, (float)LINEA, 0);
                    // AQUI IMPRIMIR OBSERVACIONES DEL DETALLE '''''''
                    if(detalle.Descripcion.ToString()  != "")
                    {
                        // //MAXIMO DEBE TENER 70 CARACTERES
                        if(detalle.Descripcion.ToString().Length > 70)
                        {
                            LINEA = LINEA - 11;
                            ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(detalle.Descripcion.Substring(0,70), fontdescripcion), (float)95, (float)LINEA , 0);
                            LINEA = LINEA - 11;
                            ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(detalle.Descripcion.Substring(70, detalle.Descripcion.Length - 70), fontdescripcion), (float)95, (float)LINEA , 0);
                           
                        }
                        else
                        {
                            LINEA = LINEA - 11;
                            ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(detalle.Descripcion, fontdescripcion), (float)100, (float)LINEA , 0);
                        }
                        LINEA = LINEA - 6;
                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("--------------------------------------------------------------------------------------------------------------------", fontdescripcion), (float)95, (float)LINEA - 26, 0);

                    }


                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(detalle.Cantidad.ToString(), fontBlack), (float)385, (float)LINEA, 0);
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C}", detalle.Precio), fontBlack), (float)452, (float)LINEA, 0);
                    string porDescuento = "0 %";
                    if (detalle.DescuentoPorcentaje > 0)
                    {
                        porDescuento = detalle.DescuentoPorcentaje.ToString() + " %"; //@String.Format(new CultureInfo("es-CL"), "{0:P2}", detalle.DescuentoPorcentaje);
                    }
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(porDescuento, fontBlack), (float)488, (float)LINEA, 0);

                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C}", detalle.MontoItem), fontBlack), (float)568, (float)LINEA, 0);

                    LINEA = LINEA - 10;
                }
                /********************************* INICIO REGION DECUENTOS SI ES QUE HAY *************************/
                fontBlack.Size = 7;
                LINEA = 230;

                LINEA = LINEA - 13;
                if (dte.Documento.DescuentosRecargos.Count > 0)
                {
                    foreach (var dec in dte.Documento.DescuentosRecargos)
                    {
                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("Descuento", fontBlack), (float)470, (float)LINEA, 0);
                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C0}", dec.Valor), fontBlack), (float)568, (float)LINEA, 0);
                    }
                    LINEA = LINEA - 13;
                }
                //ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase("Descuento", fontBlack), (float)470, (float)LINEA, 0);
                //ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C0}", "$ 1000"), fontBlack), (float)568, (float)LINEA, 0);
                /************************************* FIN REGION DESCUENTOS *************************************/

                /*************************************  INICIO REGION IMPUESTOS ILAS ******************************/
                if (dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Count > 0)
                {
                    foreach (var imp in dte.Documento.Encabezado.Totales.ImpuestosRetenciones)
                    {

                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase("Ret." + imp.TipoImpuesto.ToString(), fontBlack), (float)426, (float)LINEA, 0);
                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C0}", imp.MontoImpuesto), fontBlack), (float)568, (float)LINEA, 0);
                        LINEA = LINEA - 9;
                    }
                }




                fontBlack.Size = 9;
                /********************************************* INICIO REGION TOTALES ***********************************/
                LINEA = 167;
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C0}", dte.Documento.Encabezado.Totales.MontoExento), fontBlack), (float)568, (float)LINEA, 0);
                LINEA = LINEA - 13;
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C0}", dte.Documento.Encabezado.Totales.MontoNeto), fontBlack), (float)568, (float)LINEA, 0);
                LINEA = LINEA - 13;
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(dte.Documento.Encabezado.Totales.TasaIVA.ToString() + "%", fontBlack), (float)472, (float)LINEA - 1, 0);
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C0}", dte.Documento.Encabezado.Totales.IVA), fontBlack), (float)568, (float)LINEA, 0);
                LINEA = LINEA - 13;
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(@String.Format(new CultureInfo("es-CL"), "{0:C0}", dte.Documento.Encabezado.Totales.MontoTotal), fontBlack), (float)568, (float)LINEA, 0);
                /*********************************************** FIN REGIÓN TOTALES *************************************/


                /****************************************** IMPRIME PDF417 CON CODIGO QR ********************************/
                iTextSharp.text.Image image = jpg;
                image.SetAbsolutePosition(49, 148f);
                image.ScaleToFit(250f, 87f);
                pdfContentByte.AddImage(image);
                /******************************************* FIN REGION CON IMAGEN DE QR ********************************/
                fontBlack.Size = 6;
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(emisor.Glosa_res , fontBlack), (float)93, 130f, 0);


                /************************************** INICIO DE REGION PARA LAS REFERENCIAS ***************************/
                LINEA = LINEA + 72;
                int lineaRef = 1;
                string refe = "";
                string refe_tipo = "";
                if (dte.Documento.Referencias.Count > 0)
                {
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
                            refe_tipo = "BOL. ELECTRÓNICA(39)";
                        }
                        if (referencia.TipoDocumento == PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.OrdenCompra)
                        {
                            refe_tipo = "O. DE COMPRA(801)";
                        }

                        fontBlack.Size = 6;
                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(refe_tipo, fontBlack), (float)234, (float)LINEA, 0);
                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(referencia.FolioReferencia.ToString().PadLeft(10, Convert.ToChar("0")), fontBlack), (float)305, (float)LINEA, 0);
                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(referencia.FechaDocumentoReferencia.ToShortDateString(), fontBlack), (float)380, (float)LINEA, 0);
                        LINEA = LINEA - 7;
                        /******************* IMPRIME LA RAZON DE LA REFERENCIA EN EL CAJON DE ABAJO *********************/
                        refe = referencia.RazonReferencia;
                        if (referencia.RazonReferencia.Length > 53)
                        {
                            refe = referencia.RazonReferencia.Substring(0, 53);
                        }
                        ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(lineaRef + "- " + refe, fontBlack), (float)234, (float)LINEA - 35, 0);
                        lineaRef++;
                    }
                }

                /********************** OBSERVACIONES *****************/
                string obs1 = this.OBSERVACION_VENTA;
                string obs2 = "";
                if (obs1.Length > 62)
                {
                    obs1 = obs1.Substring(0, 61);
                    obs2 = OBSERVACION_VENTA.Substring(61, this.OBSERVACION_VENTA.Length - 61);
                }

                if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica)
                {
                    obs1 = "6 OPERACIÓN NO CONSTITUYE VENTA";
                    obs2 = "SOLO TRASLADO DE MERCADERIA.";
                }

                if (obs1 != "")
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(obs1, fontBlack), (float)236, (float)LINEA - 45, 0);

                }
                if (obs2 != "")
                {
                    ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_LEFT, new Phrase(obs2, fontBlack), (float)236, (float)LINEA - 55, 0);

                }
                
               



                stamper.Close();
            }
        }

        private string GetSectorCliente(string xRut)
        {
            string salida = "";
            //ClientesClass cli = new ClientesClass(,);
            //MySqlDataReader dr = null;

            //dr = cli.getClienteByRut(xRut, "000");
            //if (dr.HasRows == true)
            //{
            //    if(dr.Read())
            //    {
            //        salida = dr["sector"].ToString();
            //    }
            //}

            //dr.Close();
            //cli.CerrarTransaccion();

            return salida;
        }


        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

     
        private void Return()
        {
            txtFilePath.Text = "";
            lbTipo.Text = "";
            lblNombreDocumento.Text = "";
            lblFecha.Text = "";         
           
        }
        
        private void radGroupBox1_Click(object sender, EventArgs e)
        {

        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            this.Retorno();
        }

        private void Retorno()
        {
            txtFilePath.Text = "";
            lblFolio.Text = "";
            lblFecha.Text = "";
            lblMonto.Text = "0";
            lbTipo.Text = "";
            lblNombreDocumento.Text = "";
            pictureBoxTimbre.Image = null;
        }
    }
}
