using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using MySql.Data;
using MySql.Data.MySqlClient;
using Telerik.WinForms;
using Telerik.WinControls;
using System.Windows.Forms;
using Telerik.WinControls.UI;
using System.Globalization;

using System.Net.NetworkInformation;
using System.Reflection;
using System.Drawing;
using System.Text.RegularExpressions;
using SchoolManagementAdmin.objetos;
using Eltit.clases;
using SamplesDTE;

namespace Eltit.Clases
{
   public class FuncionesClass
    {

       /******* VARIALES GLBALES DEL XML DE CONFIGURACION******/
        public static string G_SERVIDOR = "";
        public static string G_USUARIOSISTEMA;
        public static string G_MYSQL_USER;
        public static string G_MYSQL_PASS;
        public static string G_CLIENTE_PREFIJO;
        /*************** VARIABLES DE EMPRESA *****************/
        public static string G_EMPRESAACTIVA;
        public static string G_EMPRESANOMBRE;
        public static string G_EMPRESADIRECCION;
        public static string G_EMPRESACOMUNA;
        public static string G_EMPRESACIUDAD;
        public static string G_EMPRESARUT;
        public static string G_EMPRESAGIRO;
        public static string G_EMPRESACODIGOCONTA;
        public static string G_EMPRESACRCC;
        public static DateTime G_DTE_FECHARESOLUCION;
        public static  Int32 G_DTE_NUMERO_RESOLUCION;
        public static string G_DTE_NOMBRE_CERTIFICADO;
        public static string G_DTE_RUT_ENVIA;
        public static string G_DTE_CASO_BASICO;
        public static string G_DTE_CASO_EXENTA;
        public static string G_DTE_CASO_COMPRAS;
        public static string[] G_DTE_ACTECO = new string[4];
        public static string G_RUBROACTIVO;
        public static string G_NOMBRERUBROACTIVO;
        public static string G_EMPRESAFONO;
        public static string G_EMPRESACORREO;
        public static string G_EMPRESASITIOWEB;
        /*************** VARIABLES DE LOCALES *****************/
        public static string G_LOCAL = "";
        public static string G_LOCAL_NOMBRE = "";
        public static string G_IMPRESORA_TICKET;
        /*************** VARIABLES DE TABLAS *****************/
        public static string TBL_EMPRESAS;
        public static string TBL_LOCALES;
        public static string TBL_VENTAS;
        public static string TBL_VENTAS_DETALLE;
        public static string TBL_MANTENCION;
        public static string TBL_CLIENTES;

        public static string BASE_VENTAS;
        public static string BASE_MANTENCION;
        public static string BASE_DTE;
        public static string G_PATH_DTE_FILE;
        public static string _BASE_FOLDER_PROD;

        public static string MAIL_GESTION_SMTP;
        public static string MAIL_GESTION_DIRECCION;
        public static string MAIL_GESTION_CLAVE;
        public static string MAIL_SOPORTE;
        public static string MAIL_MASTER_LOCAL;
        public static bool DTE_GENERA_AUTOMATICO;
        public static double G_IVA;
        public static bool DTE_PRODUCCION_BOLETA;
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static string G_PREFIJO_HOST = "eltit_";
        public static string host_direccion = "200.73.113.56";
        public static string host_user = "";
        public static string host_pass = "";

        /// <summary>
        public static string G_SERVIDORMASTER = "192.168.204.9";
        public static string G_USERS_LOG = "";
        //public static string G_SERVIDORMASTER = "192.168.204.9";
        /// </summary>


        public void CargaConfiguracionInicial()
        {
           
            RadMessageBox.SetThemeName("TelerikMetroBlue");

            G_MYSQL_USER = "sistema";
            
            G_MYSQL_PASS = "desarrollo_1990";
            G_CLIENTE_PREFIJO = "eltit_";
            G_USUARIOSISTEMA = "EDUVERGARA";
            this.LeeXMLConfiguracion();
            this.CargaValoresTablas();
            //this.LeeDatosLocal(G_LOCAL);
            /******************* CARGA PREFIJO DE RUT DE EMPRESA *********************/
            G_PATH_DTE_FILE = @"C:\PlaceDTE";
            G_IVA = 19;
            _BASE_FOLDER_PROD = @"C:\PlaceDTE\"; //+ G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + Convert.ToDouble(G_EMPRESARUT.Substring(0, 9)) + @"\Produccion\";
            
        }

        private void LeeDatosEmpresa(string xCodigo)
        {
            //EmpresaClass emp = new EmpresaClass(G_SERVIDOR);
            //MySqlDataReader dr = emp.GetEmpresaByCodigo(G_EMPRESAACTIVA);
            //G_CLIENTE_PREFIJO = Load.GetCliente(G_SERVIDOR);
            //if (dr.HasRows == true)
            //{
            //    if(dr.Read())
            //    {
            //        G_EMPRESANOMBRE = dr["nombre"].ToString();
            //        G_EMPRESARUT = dr["rut"].ToString();
            //        G_EMPRESADIRECCION = dr["direccion"].ToString();
            //        G_EMPRESACOMUNA = dr["comuna"].ToString();
            //        G_EMPRESACIUDAD = dr["ciudad"].ToString();
            //        G_EMPRESAGIRO = dr["dte_giro"].ToString();
            //        G_DTE_FECHARESOLUCION = Convert.ToDateTime(dr["dte_fecharesolucion"]);
            //        G_DTE_NUMERO_RESOLUCION = Convert.ToInt32(dr["dte_numeroresolucion"]);
            //        G_DTE_RUT_ENVIA = dr["dte_rutenvia"].ToString();
            //        G_DTE_NOMBRE_CERTIFICADO = dr["dte_certificado"].ToString();
            //        G_DTE_CASO_BASICO = dr["dte_caso_basico"].ToString();
            //        G_DTE_CASO_EXENTA = dr["dte_caso_exento"].ToString();
            //        G_DTE_CASO_COMPRAS = dr["dte_caso_compras"].ToString();
            //        G_DTE_ACTECO[0] = dr["dte_acteco_1"].ToString();
            //        G_DTE_ACTECO[1] = dr["dte_acteco_2"].ToString();
            //        G_DTE_ACTECO[2] = dr["dte_acteco_3"].ToString();
            //        G_DTE_ACTECO[3] = dr["dte_acteco_4"].ToString();
            //        MAIL_GESTION_SMTP = dr["dte_servermail"].ToString();
            //        MAIL_GESTION_DIRECCION = dr["dte_mail"].ToString();
            //        MAIL_GESTION_CLAVE = dr["dte_clavemail"].ToString();
            //    }
            //}

            //dr.Close();
            //emp.CerrarTransaccion();
        }

        private void LeeDatosLocal(string xCodigo)
        {
            try
            {
                Locales loc = new Locales(G_SERVIDOR, G_MYSQL_USER,G_MYSQL_PASS);
                MySqlDataReader dr = loc.GetLocalByCodigo(xCodigo);
                bool existelocal = false;

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        G_LOCAL_NOMBRE = dr["nombre"].ToString();
                        G_EMPRESAACTIVA = dr["empresacontable"].ToString();
                        G_IVA = Convert.ToDouble(dr["iva"]);
                        existelocal = true;

                        /****** SI ESTA AUTORIZADO PARA FACTURAR ELECTRÓNICAMENTE  *****/
                        if ((int)dr["emisor_dte"] == 0)
                        {
                            RadMessageBox.Show("Información: El recinto indicado no está Autorizado para emitir Documentos Electrónicos Tributarios.", "Error de Configuración", MessageBoxButtons.OK, RadMessageIcon.Info);
                            Application.Exit();
                        }
                    }
                }

                dr.Close();
                loc.CerrarTransaccion();

                if (existelocal == true)
                {
                    this.LeeDatosEmpresa(G_EMPRESAACTIVA);
                }
                else
                {
                    log.Info("No se cargo el Local indicado en la configuracion" + G_LOCAL_NOMBRE);

                    RadMessageBox.Show("No se pudo Cargar el Local seleccionado en la Configuracion", "Error de Configuración", MessageBoxButtons.OK, RadMessageIcon.Info);
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error:", ex);
                RadMessageBox.Show("Error: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, RadMessageIcon.Info);
            }



        }
        private void CargaValoresTablas()
        {
            TBL_EMPRESAS = "mae_empresas_tributarias";
            TBL_LOCALES = "mae_empresas_locales";
            TBL_VENTAS = "sv_documento_cabeza_";
            TBL_VENTAS_DETALLE = "sv_documento_detalle_";
            TBL_CLIENTES = "sv_maestroclientes";
            /****************************************/
            BASE_VENTAS = "ventas";
            BASE_MANTENCION = "mantencion";
          
            
        }
        private void LeeXMLConfiguracion()
        {
            try
            {
                String ruta = "";
                XmlDocument xmlDoc = new XmlDocument();
                ruta = "ConfigDTEControl.xml";
                xmlDoc.Load(ruta);

                XmlNodeList nodeList = xmlDoc.GetElementsByTagName("config");
                XmlNodeList lista = ((XmlElement)nodeList[0]).GetElementsByTagName("datos");

                foreach (XmlElement nodo in lista)
                {
                    XmlNodeList server = nodo.GetElementsByTagName("SERVIDOR");                   
                    XmlNodeList local = nodo.GetElementsByTagName("RECINTO");
                    XmlNodeList prod = nodo.GetElementsByTagName("PRODUCTION");
                    XmlNodeList imp = nodo.GetElementsByTagName("IMPRESORA_TICKET");

                    G_SERVIDOR = server[0].InnerText;
                    G_LOCAL = local[0].InnerText;
                    DTE_PRODUCCION_BOLETA = Convert.ToBoolean(prod[0].InnerText);

                    G_IMPRESORA_TICKET = imp[0].InnerText; 

                }
            }
            catch (Exception ex)
            {               
                log.Error("Error:", ex);
            }


        }

        public static string getTiposDTE(string xPrefix, string xSufijo)
        {
            string[] tipos = { "'FVE'", "'FEE'", "'NDE'","'NCE'" };
            string salida = "";
            int i = 0;
            // ... Loop with the foreach keyword.
            foreach (string value in tipos)
            {
                salida = salida + xPrefix + value + xSufijo;
                i++;
            }
            salida = salida.Substring(0, salida.Length - 3);
            return salida;
        }
        public static string getTiposDTE2(string xPrefix, string xSufijo)
        {
            string[] tipos = { "'FV'", "'FE'", "'ND'", "'NF'" };
            string salida = "";
            int i = 0;
            // ... Loop with the foreach keyword.
            foreach (string value in tipos)
            {
                salida = salida + xPrefix + value + xSufijo;
                i++;
            }
            salida = salida.Substring(0, salida.Length - 3);
            return salida;
        }
        public static string getNombredocumentoByCodigo2(string xCodigo)
        {
            string salida = "";
            switch (xCodigo)
            {
                case "FV":
                    salida = "FACTURA DE VENTA ELECTRÓNICA";
                    break;
                case "FE":
                    salida = "FACTURA NO EFECTA O EXENTA ELECTRÓNICA";
                    break;
                case "ND":
                    salida = "NOTA DE DEBITO ELECTRONICA";
                    break;
                case "NF":
                    salida = "NOTA DE CREDITO ELECTRÓNICA FACTURA";
                    break;
                case "NB":
                    salida = "NOTA DE CREDITO ELECTRÓNICA BOLETA";
                    break;
            }
            return salida;
        }
        public static string getTipoSIIByTipoDoc(string xCodigo)
        {
            string salida = "";
            switch (xCodigo)
            {
                case "FV":
                    salida = "33";
                    break;
                case "FE":
                    salida = "34";
                    break;
                case "NB":
                    salida = "56";
                    break;
                case "NF":
                    salida = "61";
                    break;
                case "BV":
                    salida = "39";
                    break;
            }
            return salida;
        }
        public static string getTipoDocByTipoSII(string xCodigo)
        {
            string salida = "";
            switch (xCodigo)
            {
                case "33":
                    salida = "FV";
                    break;
                case "34":
                    salida = "FE";
                    break;
                case "56":
                    salida = "ND";
                    break;
                case "61":
                    salida = "NF";
                    break;
            }
            return salida;
        }
        public static string getNombredocumentoByCodigo(string xCodigo)
        {
            string salida = "";
            switch (xCodigo)
            {
                case "FVE":
                    salida = "FACTURA DE VENTA ELECTRÓNICA";
                    break;
                case "FEE":
                    salida = "FACTURA NO EFECTA O EXENTA ELECTRÓNICA";
                    break;
                case "NDE":
                    salida = "NOTA DE DEBITO ELECTRONICA";
                    break;
                case "NFE":
                    salida = "NOTA DE CREDITO ELECTRÓNICA";
                    break;
                case "NBE":
                    salida = "NOTA DE CREDITO BOLETA ELECTRÓNICA";
                    break;
                case "BEL":
                    salida = "BOLETA DE VENTA ELECTRÓNICA";
                    break;
                case "BEE":
                    salida = "BOLETA EXENTA ELECTRÓNICA";
                    break;
            }
          return salida;
        }

        public static object RetornaIconoRecursos(string xextencion)
        {
            object salida = null;

            if (xextencion == ".jpg" | xextencion == ".jpeg")
                salida = Eltit.Properties.Resources.icon_jpg_20;
            if (xextencion == ".png")
                salida = Eltit.Properties.Resources.icons8_png_filled_50;
            if (xextencion == ".pdf")
                salida = Eltit.Properties.Resources.icon_pdf_25;
            if (xextencion == ".mp4" | xextencion == ".mov")
                salida = Eltit.Properties.Resources.icons8_video_25;
            if (xextencion == ".doc" | xextencion == ".docx")
                salida = Eltit.Properties.Resources.icons_word_20;
            if (xextencion == ".xlsx")
                salida = Eltit.Properties.Resources.icon_excel_25;
            if (xextencion == ".att")
                salida = Eltit.Properties.Resources.icons8_adjuntar_48;
            if (xextencion == ".ok")
                salida = Eltit.Properties.Resources.OK_48;
            if (xextencion == ".alert")
                salida = Eltit.Properties.Resources.icons8_exclamacion;
            return salida;
        }

       

        public static string GetFechaMysql(string xfecha)
        {
            return xfecha.Substring(6, 4) + "-" + xfecha.Substring(3, 2) + "-" + xfecha.Substring(0, 2);
        }
        public static void CargaAnosEnCbo(ref RadDropDownList cbo, int inicial, int final)
        {
            int i = 0;
            for (i = inicial; (i <= final); i++)
            {
                cbo.Items.Add(i.ToString());
            }

            cbo.SelectedIndex = (cbo.Items.Count - 2);
        }
        public static void CargaMesesEnCbo(ref RadDropDownList cbo, bool abreviado, int seleccionado)
        {
            int i = 0;
            string nombreMes = "";
            cbo.Items.Add("- Seleccione Mes -");
            try
            {
                DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                for (i = 1; (i <= 12); i++)
                {
                    if ((abreviado == true))
                    {
                        nombreMes = formatoFecha.GetAbbreviatedMonthName(i);
                    }
                    else
                    {
                        nombreMes = formatoFecha.GetMonthName(i);
                    }

                    cbo.Items.Add(nombreMes.ToUpper());
                }

                cbo.SelectedIndex = seleccionado;
            }
            catch (System.Exception End)
            {

                log.Error("Error", End);
            }
        }
        public bool PingToHost(string xServer)
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
        public void ColoreaCelda(GridViewCellInfo cell, Color color)
        {
            Font myFont = new Font(new FontFamily("Calibri"), 11.0F, FontStyle.Bold);

            cell.Style.Font = myFont;
            cell.Style.CustomizeFill = true;
            cell.Style.GradientStyle = GradientStyles.Solid;
            cell.Style.BackColor = color; // Color.FromArgb(171, 230, 37)
        }

        public string FechaMysql(string xfecha)
        {
            string salida;

            salida = xfecha.Substring( 6, 4) + "-" + xfecha.Substring( 3, 2) + "-" + xfecha.Substring( 0, 2);

            return salida;
        }

        public static bool IsValidEmail(string email)
        {
            if (email == string.Empty)
                return false;
            // Compruebo si el formato de la dirección es correcto.
            Regex re = new Regex(@"^[\w._%-]+@[\w.-]+\.[a-zA-Z]{2,4}$");
            Match m = re.Match(email);
            return (m.Captures.Count != 0);
        }

        public bool existeUsuario(string xuser,string xpass)
        {
            bool salida = false;
            MySqlDataReader dr;

            try
            {
                Usuarios usr = new Usuarios(G_SERVIDORMASTER, G_MYSQL_USER, G_MYSQL_PASS);
                dr = usr.GetUsuario(xuser, xpass);


                if (dr.HasRows == true)
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




        //public Boolean GetPermisosUsuario(string xuser)
        //{

        //    MySqlDataReader dr = null;
            
        //    bool salida = false;

        //    try
        //    {

        //        Usuarios per = new Usuarios(G_SERVIDORMASTER, G_MYSQL_USER, G_MYSQL_PASS);
        //        dr = per.GetPermisosUsuario(xuser);

        //        if (dr.HasRows == true)
        //        {
        //            salida = true;
        //        }
        //        else
        //        {
        //            salida = false;
                    
        //        }

        //    }
        //    catch (PingException ex)
        //    {
        //        log.Error("Error:" + MethodBase.GetCurrentMethod().DeclaringType.Name + "->" + MethodInfo.GetCurrentMethod().ToString(), ex);

        //    }

        //    return salida;

        //}



    }

}
