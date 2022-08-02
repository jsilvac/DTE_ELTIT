using Eltit.Clases;
using MySql.Data.MySqlClient;
using OpenPop.Mime;
using OpenPop.Pop3;
using PlaceSoft.Eltit.Class;
using SchoolManagementAdmin.objetos;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace SchoolManagementAdmin
{
    public partial class PopLeeCorreos : Form
    {
        public string LOCAL_ACTIVO;
        public DataTable DT_LOCALES;
        int i = 0;
        private bool cargo = false;
        Pop3Client pop3Client;
        private string  MYSQL_USER = "sistema";            
        private string   MYSQL_PASS = "desarrollo_1990";
        private static readonly log4net.ILog log =
log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private int paso_end = 0;
        public PopLeeCorreos()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {


            if (i > 5)
            {
                //timer1.Stop();
                //timer1.Enabled = false;
                //this.Close();
                RecorrerEmpresas();
                i = 0;
                paso_end++;
            }
            i++;

            if (paso_end >= 3)
            {
                this.Close();
            }

        }
        private void PopBoletas_Activated(object sender, EventArgs e)
        {
            if (cargo == false)
            {
                cargo = true;
                this.CargaEmpresas();
                timer1.Enabled = true;
                timer1.Start();
            }
        }

        private void PopLeeCorreos_Load(object sender, EventArgs e)
        {
            gvPagos.TableElement.Font = new Font("Arial", 8);     
            gvPagos.TableElement.RowHeight = 21;
         
        }
        private void CargaEmpresas()
        {
            Empresas empresas = new Empresas("192.168.4.9", this.MYSQL_USER, this.MYSQL_PASS);
            MySqlDataReader dr = empresas.GEtEmpresasConta();
            object img = null;
            FuncionesClass fu = new FuncionesClass();

            if (dr.HasRows == true)
            {
                while (dr.Read())
                {
                    gvEmpresas.Rows.Add(dr["codigoempresa"].ToString(), dr["nombre"].ToString(), "192.168.4.9", dr["empresafae"].ToString(), dr["rut"].ToString());
                }
            }

            dr.Close();
            empresas.CerrarTransaccion();


        }
        private void RecorrerEmpresas()
        {
            int i = 0;
            string rut = "";
            Inicial ini = new Inicial();
            for(i=0; i<= gvEmpresas.Rows.Count -1;i++)
            {
                rut = gvEmpresas.Rows[i].Cells[4].Value.ToString();
                lblemail.Text = "Procesando Empresa " + gvEmpresas.Rows[i].Cells[1].Value.ToString();
                lblemail.Refresh();

                ini.ColoreaCeldaYTexto(gvEmpresas.Rows[i].Cells[0], Color.YellowGreen, Color.Black, new Font("Arial", 8));
                ini.ColoreaCeldaYTexto(gvEmpresas.Rows[i].Cells[1], Color.YellowGreen, Color.Black, new Font("Arial", 8));
                gvEmpresas.Refresh();
                this.ProcesaCorreosEmpresa(rut);
                System.Threading.Thread.Sleep(1000);
                ini.ColoreaCeldaYTexto(gvEmpresas.Rows[i].Cells[0], Color.White, Color.Black, new Font("Arial", 8));
                ini.ColoreaCeldaYTexto(gvEmpresas.Rows[i].Cells[1], Color.White, Color.Black, new Font("Arial", 8));
                gvEmpresas.Refresh();
            }
        }

        private void ProcesaCorreosEmpresa(string xrut)
        {
            Empresas empresas = new Empresas("192.168.4.9", this.MYSQL_USER, this.MYSQL_PASS);
            MySqlDataReader dr = empresas.GEtEmpresasContaByRut(xrut);

            if(dr.HasRows == true)
            {
                if(dr.Read())
                {
                    lblRut.Text = dr["rut"].ToString();
                    lblNombreEmpresa.Text = dr["nombre"].ToString();
                    lblDireccion.Text = dr["direccion"].ToString();
                    lblComuna.Text = dr["comuna"].ToString();
                    lblCodigo.Text = dr["codigoempresa"].ToString();
                    lblCertificadoRut.Text = dr["rutenviasii"].ToString();
                    lblCertificadoNombre.Text = "juan patricio eltit jadue";

                    lblFechaResolucion.Text = dr["fecharesolucion"].ToString();
                    lblNroResolucion.Text = dr["numeroresolucion"].ToString();
                    lblHost.Text = dr["servermail"].ToString();
                    lblemail.Text = dr["mailsalida"].ToString();
                    lblpassemail.Text = dr["clavemail"].ToString();
                    lblPath.Text = dr["empresafae"].ToString();

                    RadPageView1.Refresh();

                    this.Read_Emails();
                }
            }

            dr.Close();
            empresas.CerrarTransaccion();
        }
        protected List<Email> Emails
        {
            get; set;

        }
        private void Read_Emails()
        {


            if (lblemail.Text == "" || lblpassemail.Text == "" )
            {
                return;
            }

            try
            {
                string ultimoCorreo = "";
                txtInfo.Text += DateTime.Now + " ==> Abriendo Bandeja Empresa " + lblNombreEmpresa.Text + " <==" + Environment.NewLine;
                txtInfo.ScrollToCaret();
                txtInfo.Refresh();
                pop3Client = new Pop3Client();
               pop3Client.Connect("pop.gmail.com", 995, true);
                pop3Client.Authenticate(lblemail.Text, lblpassemail.Text, AuthenticationMethod.UsernameAndPassword);

                txtInfo.Text += DateTime.Now + " ==> AUTENTICACION OK CON [" +  lblemail.Text.ToUpper() + "] <==" + Environment.NewLine;
                txtInfo.SelectionStart = txtInfo.Text.Length;
                txtInfo.ScrollToCaret();
                txtInfo.Refresh();

                int count = pop3Client.GetMessageCount();
                this.Emails = new List<Email>();
                int counter = 0;

                for (int i = count; i >= 1; i--)
                {
                    OpenPop.Mime.Message message = pop3Client.GetMessage(i);
                    Email email = new Email()
                    {
                        MessageNumber = i,
                        Subject = message.Headers.Subject,
                        DateSent = message.Headers.DateSent,
                        From = string.Format("<a href = 'mailto:{1}'>{0}</a>", message.Headers.From.DisplayName, message.Headers.From.Address),
                    };

                    ultimoCorreo = message.Headers.From.DisplayName + " " + email.Subject.ToString();
                    log.Debug(ultimoCorreo + " " + message.Headers.DateSent.ToString());
                    List<MessagePart> attachments = message.FindAllAttachments();
                    bool grabado = false;
                    if (message.Headers.Subject.Contains("Mail Delivery") || message.Headers.Subject.Contains("Mail delivery"))
                    {

                        pop3Client.DeleteMessage(i);
                        log.Debug("DELETE: " + ultimoCorreo + " " + message.Headers.DateSent.ToString());
                    }
                    else
                    {
                        foreach (MessagePart attachment in attachments)
                        {
                            string extension = Path.GetExtension(attachment.FileName);
                            log.Debug(attachment.FileName);
                            if (extension == ".xml" || extension == ".XML") 
                            {
                                string filename = string.Format(@"{0}{1}_{2}{3}", @"C:\Temp\", Path.GetFileNameWithoutExtension(attachment.FileName), "", Path.GetExtension(attachment.FileName));
                                string filename2 = @"\\192.168.4.6\fae\" + lblPath.Text + @"\correos_recibidos\" + attachment.FileName;

                                attachment.Save(new FileInfo(filename));

                                FileInfo fi = new FileInfo(filename);
                                 fi.CopyTo(filename2, true);


                                log.Debug("GRABA:" + filename);
                                grabado = LeeXML2(lblRut.Text, "eltit_", filename, message.Headers.From.Address, attachment.FileName, message.Headers.Date);
                                log.Debug("GRABO DTE " + lblRut.Text + " COD:"+ lblCodigo.Text + " Subj:" + message.Headers.From.Address + " en " + lblPath.Text );
                                System.Threading.Thread.Sleep(500);
                            }


                        }
                    }


                    if ( attachments.Count == 0)
                    {
                        counter++;
                        txtInfo.Text += DateTime.Now + " => Elimnando  Correo From [" + message.Headers.From + "] <==" + Environment.NewLine;
                        txtInfo.Text += DateTime.Now + " ===> No Se encontraron Adjuntos =" + Environment.NewLine;
                        txtInfo.ScrollToCaret();
                        txtInfo.Refresh();
                        pop3Client.DeleteMessage(i);
                    }
                    if (grabado == true )
                    {
                        counter++;
                        txtInfo.Text += DateTime.Now + " => Importando  Correo From [" + message.Headers.From + "] <==" + Environment.NewLine;
                        txtInfo.SelectionStart = txtInfo.Text.Length;
                        txtInfo.ScrollToCaret();
                        txtInfo.Refresh();
                        pop3Client.DeleteMessage(i);
                    }
                    if (counter > 50)
                    {
                        log.Debug("BREAK " + counter);
                        break;

                    }
                }
                pop3Client.Disconnect();
                txtInfo.Text += DateTime.Now + " ********* Se importaron  [" + counter + "] a Empresa [" + lblNombreEmpresa.Text + "] *********" + Environment.NewLine;
                txtInfo.Refresh();
            }
            catch (Exception ex)
            {
                log.Debug("Excepcion no controlada en " + lblemail.Text);
                pop3Client.Disconnect(); 
                log.Error(ex);
                MessageBox.Show("Error en Read_Email() " + ex.Message.ToString());
                return;
            }


        }

        private bool LeeXML2(string xRutCliente, string xPRefijo, string xml, string xCorreo, string xNombreArchivo, string xFechaCorreo)
        {
            XmlDocument xmlDoc = xmlDoc = new XmlDocument();
            bool salida = false;
            xmlDoc.Load(xml);
            XmlNode node = xmlDoc.DocumentElement.FirstChild;
            XmlNodeList lstNodos = node.ChildNodes;
            string XML_DTE = "";
            string rutEmisor = "";
            string razonSocialEmisor = "";
            string rutReceptor = "";
            string tipoDoc = "";
            string numeroDoc = "";
            string FechaEmision = "";
            double Totaldoc = 0;
            string TIPO_XML = "";
            int cuentaInterior = 0;


            try
            {

                TIPO_XML = xmlDoc.DocumentElement.Name;
                //xmlDoc.OuterXml
                string RutFormat = xRutCliente;// Convert.ToDouble(xRutCliente.Substring(0, 9)) + "-" + xRutCliente.Substring(9, 1);
                if (TIPO_XML == "EnvioDTE" || TIPO_XML == "EnvioBOLETA")
                {
                    for (int i = 0; i < lstNodos.Count; i++)
                    {

                        if (lstNodos[i].Name == "Caratula")
                        {
                            XmlNodeList nodorut = ((XmlElement)lstNodos[0]).GetElementsByTagName("RutEmisor");
                            XmlNodeList receptor = ((XmlElement)lstNodos[0]).GetElementsByTagName("RutReceptor");
                            rutEmisor = nodorut[0].InnerXml;
                            rutReceptor = receptor[0].InnerXml;

                            if (RutFormat != rutReceptor)
                            {
                                return false;
                            }

                        }
                        if (lstNodos[i].Name == "DTE")
                        {
                            XML_DTE = lstNodos[i].OuterXml;
                            XmlNodeList lstChilds = lstNodos[i].ChildNodes;
                            for (int j = 0; j < lstChilds.Count; j++)
                            {
                                XmlNode node2 = lstChilds[j];
                                if (node2.Name == "Documento")
                                {
                                    XmlNodeList fist = lstChilds[j].ChildNodes;
                                    for (int k = 0; k < fist.Count; k++)
                                    {
                                        XmlNode node3 = fist[k];
                                        if (node3.Name == "Encabezado")
                                        {
                                            XmlNodeList fist2 = fist[k].ChildNodes;
                                            for (int l = 0; l < fist2.Count; l++)
                                            {
                                                XmlNode node4 = fist2[l];
                                                if (node4.Name == "IdDoc")
                                                {
                                                    XmlNodeList fist3 = fist2[l].ChildNodes;
                                                    for (int h = 0; h < fist3.Count; h++)
                                                    {
                                                        XmlNode node5 = fist3[h];
                                                        if (node5.Name == "TipoDTE")
                                                        {
                                                            tipoDoc = node5.InnerText;
                                                        }
                                                        if (node5.Name == "Folio")
                                                        {
                                                            numeroDoc = node5.InnerText;
                                                        }
                                                        if (node5.Name == "FchEmis")
                                                        {
                                                            FechaEmision = node5.InnerText;
                                                        }
                                                    }
                                                }
                                                if (node4.Name == "Emisor")
                                                {
                                                    XmlNodeList ndeEmisor = fist2[l].ChildNodes;
                                                    for (int t = 0; t < ndeEmisor.Count; t++)
                                                    {
                                                        XmlNode node6 = ndeEmisor[t];
                                                        if (node6.Name == "RznSoc")
                                                        {
                                                            razonSocialEmisor = node6.InnerText;
                                                        }

                                                    }
                                                }
                                                if (node4.Name == "Totales")
                                                {
                                                    XmlNodeList ndeTotal = fist2[l].ChildNodes;
                                                    for (int m = 0; m < ndeTotal.Count; m++)
                                                    {
                                                        XmlNode node7 = ndeTotal[m];
                                                        if (node7.Name == "MntTotal")
                                                        {
                                                            Totaldoc = Convert.ToDouble(node7.InnerText);
                                                        }

                                                    }
                                                }

                                            }
                                        }
                                    }
                                }


                            }


                            if (rutEmisor != "" && tipoDoc != "" && numeroDoc != "" && FechaEmision != "")
                            {
                                //this.GrabaDteProveedor(xPRefijo, xRutCliente, "00", rutEmisor, razonSocialEmisor, tipoDoc, numeroDoc, FechaEmision,
                                //                                          xCorreo, Totaldoc, XML_DTE, xNombreArchivo);
                                salida = true;
                                cuentaInterior++;
                            }
                            else
                            {
                                salida = false;
                            }

                        }// END IF NODO DTE
                    }

                    /***********************  GRABA ENVIO PROVEEDOR, SII Y RESPUESTA EN FORMATO BRUTO  *****************/

                    this.GrabaRecibidos(xCorreo, xNombreArchivo, xmlDoc.OuterXml, rutEmisor, "PROVEEDOR", lblPath.Text, cuentaInterior);
                    salida = true;

                } // end if envio





                if (TIPO_XML == "RespuestaDTE" || TIPO_XML == "EnvioRecibos")
                {
                    for (int i = 0; i < lstNodos.Count; i++)
                    {
                        if (lstNodos[i].Name == "Caratula")
                        {
                            XmlNodeList nodoresponde = ((XmlElement)lstNodos[0]).GetElementsByTagName("RutResponde");
                            XmlNodeList nodorecibe = ((XmlElement)lstNodos[0]).GetElementsByTagName("RutRecibe");
                            rutEmisor = nodoresponde[0].InnerXml;
                            rutReceptor = nodorecibe[0].InnerXml;

                            if (RutFormat != rutReceptor)
                            {
                                return false;
                            }
                            else
                            {
                               // this.GrabaXMLIntercambio(xPRefijo, xRutCliente, "00", TIPO_XML, xNombreArchivo, xCorreo, xmlDoc.InnerXml, rutEmisor, rutReceptor, xFechaCorreo);
                                salida = true;
                            }

                        }
                    }
                    this.GrabaRecibidos(xCorreo, xNombreArchivo, xmlDoc.OuterXml, rutEmisor, "CLIENTE", lblPath.Text, 1);
                    salida = true;
                }

                if (TIPO_XML == "RESULTADO_ENVIO")
                {
                    XML_DTE = xmlDoc.InnerXml;
                    string track = "";
                    string estado = "OK";
                    for (int i = 0; i < lstNodos.Count; i++)
                    {

                        if (lstNodos[i].Name == "RUTEMISOR")
                        {
                            rutEmisor = lstNodos[i].InnerText;

                            if (RutFormat != rutEmisor)
                            {
                                return false;
                            }

                        }

                        if (lstNodos[i].Name == "TRACKID")
                        {
                            track = lstNodos[i].InnerText;
                        }
                    }


                    XmlNodeList lstNodosRes = xmlDoc.DocumentElement.ChildNodes;


                    for (int x = 0; x < lstNodosRes.Count; x++)
                    {
                        if (lstNodosRes[x].Name == "REVISIONENVIO")
                        {
                            if (lstNodosRes[x].OuterXml.Contains("Rechazado"))
                            {
                                estado = "ERROR"; 
                            }
                        }

                    }

                    this.GrabaRecibidos(xCorreo, xNombreArchivo, xmlDoc.OuterXml, rutEmisor, "SII", lblPath.Text, 1);
                    salida = true;

                    /*************** AQUI ENVIR CORREO DE AVISO ********************/

                }



            }
            catch (Exception ex)
            {
                log.Error(ex);
                return false;
            }




            return salida;

        }
        
        private void GrabaRecibidos(string xCorreo, string xNombreArchivo,string xArchivo,
                                    string xRutEmisor,string xTipoRecibo, 
                                    string xCarpetaFAE, int xCantFiles)
                                   
        {
            xArchivo = xArchivo.Replace("'", "~");
           // xArchivo = xArchivo.Replace("`", "~");
           // xArchivo = xArchivo.Replace("DirRecep> ", "DirRecep>O");

            Empresas empresas = new Empresas("192.168.4.9", this.MYSQL_USER, this.MYSQL_PASS);
            empresas.GrabaDteRecepcion(xRutEmisor, xTipoRecibo, xCorreo, xArchivo, xNombreArchivo,
               xCarpetaFAE, xCantFiles, DateTime.Now.ToString("yyyy-MM-dd"));



        }





      
        private void EnviarEmailFolios(string xRut, string xLocal, string xServidor, string xTipo, double xCritico, double xCurrCaf, 
                                       string xCaja, string xcorreo_soporte)
        {

            string htmlString = @"<html>";
         
            htmlString = htmlString + "<body>";
            htmlString = htmlString + "<img src='http://www.placesoft.cl/images/eltit/header_eltit.png' border='0'  />";
            htmlString = htmlString + "<p>Se Han detectado Folios Críticos</p>";
            htmlString = htmlString + "--------------------------------------------<br>";

            htmlString = htmlString + "<div style='overflow-x:auto; font-family:Arial, Helvetica, sans-serif; color:#666; font-size:12px;'> ";
            htmlString = htmlString + " <table  border='0'> ";
            htmlString = htmlString + " <tr> ";
            htmlString = htmlString + " <td><strong>Rut</strong></td> ";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + Convert.ToDouble(xRut.Substring(0, 9)).ToString() + "-" + xRut.Substring(9, 1) + "</td> ";
            htmlString = htmlString + " </tr> ";
            htmlString = htmlString + " <tr> ";
            htmlString = htmlString + " <td><strong>Local</strong></td> ";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + xLocal + "</td> ";
            htmlString = htmlString + " </tr> ";
            htmlString = htmlString + " <tr> ";
            htmlString = htmlString + " <td><strong>Servidor</strong></td> ";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + xServidor + "</td> ";
            htmlString = htmlString + " </tr>";
            htmlString = htmlString + " <tr>";
            htmlString = htmlString + " <td><strong>Tipo Doc</strong></td>";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + xTipo + "</td>";
            htmlString = htmlString + " </tr>";
            htmlString = htmlString + " <tr>";
            htmlString = htmlString + " <td><strong>Caja Doc</strong></td>";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + xCaja + "</td>";
            htmlString = htmlString + " </tr>";
            htmlString = htmlString + " <tr>";
            htmlString = htmlString + " <td><strong>Critico</strong></td>";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>[" + xCritico + "] Quedan[" + xCurrCaf + "]</td>";
            htmlString = htmlString + " </tr>";
    
            htmlString = htmlString + " </table>";
            htmlString = htmlString + " </div> <br>";



            //htmlString = htmlString + "--------------------------------------------<br>";
            //htmlString = htmlString + " Local &nbsp;&nbsp;&nbsp;    :" + xLocal + "<br>";
            //htmlString = htmlString + " Servidor   :" + xServidor + "<br>";
            //htmlString = htmlString + " Tipo Doc :" + xTipo + "<br>";
            //htmlString = htmlString + " Caja Doc :" + xCaja + "<br>";
            //htmlString = htmlString + " Critico         :[" + xCritico + "] Quedan[" + xCurrCaf + "]  <br>";
            htmlString = htmlString + "--------------------------------------------<br>";
            htmlString = htmlString + "<p>Enviado el " + DateTime.Now.ToString("dd-MM-yyyy") + " a las " + DateTime.Now.ToString("HH:mm:ss tt") + "</p>";
            htmlString = htmlString + "<img src='http://www.placesoft.cl/images/eltit/footer_eltit.png' border='0'  />";
            htmlString = htmlString + "</body>";
            htmlString = htmlString + " </html>";

            Inicial.EnviarEmail(xcorreo_soporte, Inicial.G_CORREO_SOPORTE_COPIA, "CAF " + xTipo, "CAF CRITICOS EN " + xLocal, htmlString);

        }
    
      
    }
}
