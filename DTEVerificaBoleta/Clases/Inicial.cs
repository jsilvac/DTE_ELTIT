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
using PlaceDTE.Clases;
using PlaceDTE;
using System.ServiceProcess;

namespace SamplesDTE.Clases
{
   public class Inicial
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

        public static Int32 G_DTE_CANTIDAD_POR_ENVIO;
        public static bool G_DTE_ENVIO_AUTOMATICO;
        public static bool G_DTE_PRODUCCION;

        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        //public static string G_SERVIDOR_XML_DIRECCION = "192.168.4.200";
        public static string G_SERVIDOR_XML_DIRECCION = "192.168.4.23";
        public static string G_SERVIDOR_XML_ROOT = "sistema";
        public static string G_SERVIDOR_XML_PASS = "desarrollo_1990";

        public void CargaConfiguracionInicial()
        {
            RadMessageBox.SetThemeName("TelerikMetroBlue");
            G_MYSQL_USER = "root";
            G_MYSQL_PASS = "1121";    
            G_USUARIOSISTEMA = "EDUVERGARA";
            this.LeeXMLConfiguracion();
            this.CargaValoresTablas();
            this.LeeDatosLocal(G_LOCAL);
            /******************* CARGA PREFIJO DE RUT DE EMPRESA *********************/
            BASE_DTE = G_EMPRESARUT.Substring(0, 9);
            BASE_DTE = "dte_"+ Convert.ToDouble(BASE_DTE).ToString();
          
        }

        private void LeeDatosEmpresa(string xCodigo)
        {
            EmpresaClass emp = new EmpresaClass(G_SERVIDOR);
            MySqlDataReader dr = emp.GetEmpresaByCodigo(G_EMPRESAACTIVA);
            G_CLIENTE_PREFIJO = Load.GetCliente(G_SERVIDOR);
            if (dr.HasRows == true)
            {
                if(dr.Read())
                {
                    G_EMPRESANOMBRE = dr["nombre"].ToString();
                    G_EMPRESARUT = dr["rut"].ToString();
                    G_EMPRESADIRECCION = dr["direccion"].ToString();
                    G_EMPRESACOMUNA = dr["comuna"].ToString();
                    G_EMPRESACIUDAD = dr["ciudad"].ToString();
                    G_EMPRESAGIRO = dr["dte_giro"].ToString();
                    G_DTE_FECHARESOLUCION = Convert.ToDateTime(dr["dte_fecharesolucion"]);
                    G_DTE_NUMERO_RESOLUCION = Convert.ToInt32(dr["dte_numeroresolucion"]);
                    G_DTE_ACTECO[0] = dr["dte_acteco_1"].ToString();
                    G_DTE_ACTECO[1] = dr["dte_acteco_2"].ToString();
                    G_DTE_ACTECO[2] = dr["dte_acteco_3"].ToString();
                    G_DTE_ACTECO[3] = dr["dte_acteco_4"].ToString();

                    G_DTE_RUT_ENVIA = dr["dte_rutenvia"].ToString();
                    G_DTE_NOMBRE_CERTIFICADO = dr["dte_certificado"].ToString();
                    G_DTE_CASO_BASICO = dr["dte_caso_basico"].ToString();
                    G_DTE_CASO_EXENTA = dr["dte_caso_exento"].ToString();
                    G_DTE_CASO_COMPRAS = dr["dte_caso_compras"].ToString();
                    if(G_DTE_NUMERO_RESOLUCION == 0)
                    {
                        G_DTE_PRODUCCION = false;
                    }
                    else
                    {
                        G_DTE_PRODUCCION = true;
                    }
                        
                }
            }

            dr.Close();
            emp.CerrarTransaccion();
        }
        private void LeeDatosLocal(string xCodigo)
        {
            try
            {
                LocalesClass loc = new LocalesClass(G_SERVIDOR);
                MySqlDataReader dr = loc.GetLocalByCodigo(xCodigo);
                bool existelocal = false;

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        G_LOCAL_NOMBRE = dr["nombre"].ToString();
                        G_EMPRESAACTIVA = dr["empresacontable"].ToString();
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
                    
                    RadMessageBox.Show("No se pudo Cargar el Local seleccionado en la Configuracion","Error de Configuración", MessageBoxButtons.OK, RadMessageIcon.Info);
                    Application.Exit();
                }
            }
            catch(Exception ex)
            {
                log.Error("Error:" + MethodBase.GetCurrentMethod().DeclaringType.Name + "->" + MethodInfo.GetCurrentMethod().ToString(), ex);
                RadMessageBox.Show("Error: " + ex.Message.ToString(), "Error", MessageBoxButtons.OK, RadMessageIcon.Info);
            }
          
            

        }
        private void CargaValoresTablas()
        {
            TBL_EMPRESAS = "mae_empresas_tributarias";
            TBL_LOCALES = "mae_empresas_locales";
            TBL_VENTAS = "local_venta_cabeza_";
            TBL_VENTAS_DETALLE = "local_venta_detalle_";
            TBL_CLIENTES = "mae_clientes";
            /****************************************/
            BASE_VENTAS = "local";
            BASE_MANTENCION = "mantencion";
          
            
        }
        private void LeeXMLConfiguracion()
        {
            try
            {
                String ruta = "";
                XmlDocument xmlDoc = new XmlDocument();
                ruta = "ConfigDTEEnvio.xml";
                xmlDoc.Load(ruta);

                XmlNodeList nodeList = xmlDoc.GetElementsByTagName("config");
                XmlNodeList lista = ((XmlElement)nodeList[0]).GetElementsByTagName("datos");

                foreach (XmlElement nodo in lista)
                {
                    XmlNodeList server = nodo.GetElementsByTagName("SERVIDOR");                   
                    XmlNodeList local = nodo.GetElementsByTagName("RECINTO");
                    XmlNodeList automatico = nodo.GetElementsByTagName("SENDTIMERSET");
                    XmlNodeList cantenvio = nodo.GetElementsByTagName("SENDQTYITEM");
                    XmlNodeList produccion = nodo.GetElementsByTagName("SERVER_DTE");

                    G_SERVIDOR = server[0].InnerText;
                    G_LOCAL = local[0].InnerText;
                    G_DTE_CANTIDAD_POR_ENVIO = Convert.ToInt32(cantenvio[0].InnerText);
                    G_DTE_ENVIO_AUTOMATICO = Convert.ToBoolean(automatico[0].InnerText);
                   // G_DTE_PRODUCCION = Convert.ToBoolean(produccion[0].InnerText);

                }
            }
            catch (Exception ex)
            {
                log.Error("Error:" + MethodBase.GetCurrentMethod().DeclaringType.Name + "->" + MethodInfo.GetCurrentMethod().ToString(), ex);
            }


        }

        public static string getTiposDTE(string xPrefix, string xSufijo)
        {
            string[] tipos = { "'FVE'", "'FEE'", "'NDE'","'NCE'", "'BEL'", "'BEE'" };
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
                case "BEL":
                    salida = "BOLETA DE VENTA AFECTA ELECTRÓNICA";
                    break;
                case "BEE":
                    salida = "BOLETA DE VENTA EXENTA ELECTRÓNICA";
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
                case "ND":
                    salida = "56";
                    break;
                case "NF":
                    salida = "61";
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
                case "NCE":
                    salida = "NOTA DE CREDITO ELECTRÓNICA";
                    break;
                case "BEL":
                    salida = "BOLETA DE VENTA AFECTA ELECTRÓNICA";
                    break;
                case "BEE":
                    salida = "BOLETA DE VENTA EXENTA ELECTRÓNICA";
                    break;
            }
          return salida;
        }

        public static object RetornaIconoRecursos(string xextencion)
        {
            object salida = null;

            if (xextencion == ".jpg" | xextencion == ".jpeg")
                salida = Properties.Resources.icon_jpg_20;
            if (xextencion == ".png")
                salida = Properties.Resources.icons8_png_filled_50;
            if (xextencion == ".pdf")
                salida = Properties.Resources.icon_pdf_25;
            if (xextencion == ".mp4" | xextencion == ".mov")
                salida = Properties.Resources.icons8_video_25;
            if (xextencion == ".doc" | xextencion == ".docx")
                salida = Properties.Resources.icons_word_20;
            if (xextencion == ".xlsx")
                salida = Properties.Resources.icon_excel;
            if (xextencion == ".att")
                salida = Properties.Resources.icons8_adjuntar_48;
            if (xextencion == ".ok")
                salida = Properties.Resources.OK_48;
            if (xextencion == ".alert")
                salida = Properties.Resources.icons8_exclamacion;
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

                log.Error("Error:"+ MethodInfo.GetCurrentMethod().ToString(), End);
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
        public static bool PingToHost(string xServer)
        {
            bool salida = false;


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


    }

    }
