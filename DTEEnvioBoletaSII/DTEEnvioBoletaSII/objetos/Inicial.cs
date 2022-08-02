using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.ServiceProcess;
using System.Net.Mail;
using System.Net;
using MySql.Data.MySqlClient;
using System.Reflection;
using System.Net.NetworkInformation;
using Telerik.WinControls.UI;
using System.Drawing;
using Telerik.WinControls;

namespace SchoolManagementAdmin.objetos
{
   public class Inicial
    {
        private string ruta = "";
        public static string G_SERVIDOR = "";
        public static string G_CAJA = "";
        public static string G_VENDEDOR = "";
        public static string G_LOCAL = "";
        public static string G_LOCAL_NOMBRE = "";
        public static string G_LOCAL_CLOUD_ECOMMERCE = "";
        public static string G_CLIENTE_SISTEMA = "";
        public static string G_RUBRO = "";
        public static string G_LOCAL_RUT = "";
        private string error = "";

        /************************************************
         *   SERVIDORES DE REPLICACION  */
        public static string CLOUD_01;
        public static string CLOUD_02;
        public static string CLOUD_03;
        public static string CLOUD_04;
        public static string G_INTERVAL;
        public static string G_WHERE_CONSULTAS = "";


        public static string G_MYSQL_USER;
        public static string G_MYSQL_PASS;
        public static string G_WEBSERVER = "";
        public static string G_WEBDATABASE = "";
        public static string G_WEBUSER = "";
        public static string G_WEBPASSWORD = "";
        public static string G_CENTRAL = "";
        public static string G_NROREGISTROS = "";
        public static string G_USER_MYSQL = "";
        public static string G_PASSWORD_MYSQL = "";
        public static string G_CORREO_SOPORTE_PRINCIPAL = "";
        public static string G_CORREO_SOPORTE_COPIA = "";
        public static string G_CORREO_SOPORTE_COPIA2 = "";

        public static string G_SMTPSERVER = "";
        public static string G_SMTPUSUARIO = "";
        public static string G_SMTPCLAVE = "";
         private static readonly log4net.ILog log =
           log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static bool G_ERROR = false;
        //public static string G_SERVIDOR_XML_DIRECCION = "192.168.4.200";
        //public static string G_SERVIDOR_XML_ROOT = "sistema";
        //public static string G_SERVIDOR_XML_PASS = "desarrollo_1990";

        // public static string G_SERVIDOR_XML_DIRECCION = "192.168.4.200";
        public static string G_SERVIDOR_XML_DIRECCION = "192.168.4.23";
        public static string G_SERVIDOR_XML_ROOT = "sistema";
        public static string G_SERVIDOR_XML_PASS = "desarrollo_1990";

        public void CargaConfiguracion()
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                ruta = "updateweb.xml";
                xmlDoc.Load(ruta);

                XmlNodeList nodeList = xmlDoc.GetElementsByTagName("config");
                XmlNodeList lista = ((XmlElement)nodeList[0]).GetElementsByTagName("datos");

                foreach (XmlElement nodo in lista)
                {
                    XmlNodeList server = nodo.GetElementsByTagName("SERVIDOR");                    
                    XmlNodeList local = nodo.GetElementsByTagName("LOCAL");
                    XmlNodeList registros = nodo.GetElementsByTagName("REGISTROS");

                    XmlNodeList svr01 = nodo.GetElementsByTagName("CLOUD_01");
                    XmlNodeList svr02 = nodo.GetElementsByTagName("CLOUD_02");
                    XmlNodeList svr03 = nodo.GetElementsByTagName("CLOUD_03");
                    XmlNodeList interval = nodo.GetElementsByTagName("INTERVAL");
                    G_SERVIDOR = server[0].InnerText;              
                    G_LOCAL = local[0].InnerText;                  
                    G_NROREGISTROS = registros[0].InnerText;

                    CLOUD_01 = svr01[0].InnerText;
                    CLOUD_02 = svr02[0].InnerText;
                    CLOUD_03 = svr03[0].InnerText;
                    G_INTERVAL= interval[0].InnerText;

                }
                G_CLIENTE_SISTEMA = "aliupos_";
                G_RUBRO = "00";
                G_MYSQL_USER = "sistema";
                G_MYSQL_PASS = "desarrollo_1990";

                /**************** datos del email **************/

                G_SMTPSERVER = "mail.placesoft.cl";
                G_SMTPUSUARIO = "eltit_dte@placesoft.cl";
                G_SMTPCLAVE = "eltit12345";

                //G_SMTPSERVER = "corre.eltit.cl";
                //G_SMTPUSUARIO = "envio250@eltit.cl";
                //G_SMTPCLAVE = "Eltit2020.";


                G_CORREO_SOPORTE_PRINCIPAL = "edu.vergara@hotmail.com";
              //  G_CORREO_SOPORTE_COPIA     = "jsilva@eltit.cl";



                //this.LeeDatosLocal(G_LOCAL);

                //G_WHERE_CONSULTAS = " ( basedatos LIKE '%_dte_%' OR ";
                //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " basedatos LIKE '%_local%' OR ";
                //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " basedatos = 'false' ) And ";

                G_WHERE_CONSULTAS = " ( query_str LIKE '%local_venta_detalle_%' OR ";
                G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%dte_fae_sobres_envios%'  OR ";
                G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%local_venta_cabeza_%'  OR ";
                G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%local_venta_mediodepago_%'  OR ";
                G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%dte_boe_local%'  OR ";
                G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%dte_fae_local%' ) And ";
                G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str NOT LIKE '%local_movimientos_detalle%' AND ";






                //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%%'  OR ";
                //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%%'  OR ";


                //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + ") AND";
            }
            catch(Exception ex)
            {
                this.error = ex.Message;
            }
            

        }
        private void LeeDatosLocal(string xCodigo)
        {
            try
            {
                LocalesClass loc = new LocalesClass();
                MySqlDataReader dr = loc.GetLocalByCodigo();
                bool existelocal = false;

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        G_LOCAL_NOMBRE = dr["nombre"].ToString();
                        G_WEBSERVER = dr["ecommerce_ftp"].ToString();
                        G_WEBDATABASE = dr["ecommerce_db"].ToString();
                        G_WEBUSER = dr["ecommerce_user"].ToString();
                        G_WEBPASSWORD = dr["ecommerce_oauth"].ToString();
                        G_LOCAL_CLOUD_ECOMMERCE = dr["ecommerce_cloud"].ToString();
                        G_LOCAL_RUT = dr["rut"].ToString();
                        /****** SI ESTA AUTORIZADO PARA FACTURAR ELECTRÓNICAMENTE  *****/
                        if (dr["emisor_dte_boleta"].ToString() == "0")
                        {
                            MessageBox.Show("Información: El recinto indicado no está autorizado para Subir Información Electrónica.", "Error de Configuración", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Application.Exit();
                        }
                    }
                }

                dr.Close();
                loc.CerrarTransaccion();
              
            }
            catch (Exception ex)
            {
                log.Error("Error:", ex);
                MessageBox.Show("Error: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }



        }
        public string VerificaServicio(string SERVICENAME)
        {
            ServiceController sc = new ServiceController(SERVICENAME);

            switch (sc.Status)
            {
                case ServiceControllerStatus.Running:
                    return "Running";
                case ServiceControllerStatus.Stopped:
                    return "Stopped";
                case ServiceControllerStatus.Paused:
                    return "Paused";
                case ServiceControllerStatus.StopPending:
                    return "Stopping";
                case ServiceControllerStatus.StartPending:
                    return "Starting";
                default:
                    return "Status Changing";
            }
        }

        public void StartService(string serviceName, int timeoutMilliseconds)
        {
            ServiceController service = new ServiceController(serviceName);
            try
            {
                TimeSpan timeout = TimeSpan.FromMilliseconds(timeoutMilliseconds);

                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running, timeout);
                
            }
            catch
            {
                // ...
            }
        }

        public string rut(string numrut)
        {
            double[] mataux = new double[10];
            int i;
            double suma;
            string salida = "";
            string[] guia = new[] { "4", "3", "2", "7", "6", "5", "4", "3", "2" };
            // guia = Array("4", "3", "2", "7", "6", "5", "4", "3", "2")
            suma = 0;
            for (i = 0; i <= 8; i++)
            {
                mataux[i] = Convert.ToInt32(guia[i]) * double.Parse(numrut.Substring(i, 1));
                suma = suma + mataux[i];
            }

            //For i = 0 To 8
            //    mataux(i) = Val(guia(i)) * Val(Mid(numrut, i + 1, 1))
            //    suma = suma + mataux(i)
            //Next

            salida = Convert.ToString(11 - suma % 11);
            switch (salida)
            {
                case "11":
                    {
                        salida = "0";
                        break;
                    }

                case "10":
                    {
                        salida = "K";
                        break;
                    }
            }

            return salida;
        }

        public Boolean EsNumero(KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar))
            {
                return false;
            }
            else if (Char.IsControl(e.KeyChar)) //permitir teclas de control como retroceso 
            {
                return false;
            }
            else
            {
                //el resto de teclas pulsadas se desactivan 
                return true;
            }
                       
        }

        public static void EnviarEmail(string xmailDestino, string xMailCopia, string xNombreVendedor,string xSubject,string xBody)
        {
            try
            {                
            
                SmtpClient MyMail = new SmtpClient();
                MailMessage MyMsg = new MailMessage();
                MyMail.Host = G_SMTPSERVER.ToLower();
                MyMail.Port = 25;
                MyMsg.Priority = MailPriority.Normal;
                MyMsg.To.Add(new MailAddress(xmailDestino));
                if(xMailCopia != "")
                {
                    MyMsg.To.Add(new MailAddress(xMailCopia));
                }
                MyMsg.Subject = xSubject;
                MyMsg.SubjectEncoding = Encoding.UTF8;
                MyMsg.IsBodyHtml = true;
                MyMsg.From = new MailAddress(G_SMTPUSUARIO,xNombreVendedor);
                MyMsg.BodyEncoding = Encoding.UTF8;
                MyMsg.Body = xBody;
                MyMail.UseDefaultCredentials = false;
                NetworkCredential MyCredentials = new NetworkCredential(G_SMTPUSUARIO.ToLower(), G_SMTPCLAVE);
                //MyMail.EnableSsl = true;
                MyMail.Credentials = MyCredentials;
                MyMail.Timeout = 5000000;
                MyMail.Send(MyMsg);
            }
            catch(SmtpException ex)
            {
                MessageBox.Show("Error:" + ex.Message.ToString());
            }
           
        }

        public void ColoreaCeldaYTexto(GridViewCellInfo cell, Color color, Color colorTexto, Font xFont)
        {
            cell.Style.Font = xFont;
            cell.Style.ForeColor = colorTexto;
            cell.Style.CustomizeFill = true;
            cell.Style.GradientStyle = GradientStyles.Solid;
            cell.Style.BackColor = color; // Color.FromArgb(171, 230, 37)
        }

        public  bool PingToHost(string xServer)
        {
            bool salida = false;

            // AGREGAR REGLA A WINDOWS 10 PARA PERMITIR PING
            // netsh advfirewall firewall add rule name="ping" protocol=ICMPV4 dir=in action=allow

            try
            {
                Ping pi = new Ping();
                PingReply res;
                string s2;
                s2 = xServer;
                res = pi.Send(s2);
                if (res.Status == IPStatus.Success)
                {
                    salida = true;
                }
                else
                {
                    salida = false;
                }
            }
            catch (PingException ex)
            {
                log.Error("Error:" + MethodBase.GetCurrentMethod().DeclaringType.Name + "->" + MethodInfo.GetCurrentMethod().ToString(), ex);
                return false;
            }


            return salida;
        }

        public string servidor
        {
            get{
                return G_SERVIDOR;
            }
            set{
                G_SERVIDOR = value;
            }

        } 

        public string caja
        {
            get
            {
                return G_CAJA;
            }
            set
            {
                G_CAJA = value;
            }
        }

        public string vendedor
        {
            get
            {
                return G_VENDEDOR;
            }
            set
            {
                G_VENDEDOR = value;
            }

        }

        public string local
        {
            get
            {
                return G_LOCAL;
            }
            set
            {
                G_LOCAL = value;
            }
        }


    }
}
