using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using SchoolManagementAdmin.objetos;
using MetroFramework;
using System.IO;
using System.Net.NetworkInformation;
using MySql;
using MySql.Data.MySqlClient;
using OpenPop.Pop3;
using OpenPop.Mime;
using System.Web;

using System.Xml.Linq;
using System.Xml;
using PlaceSoft.Eltit.Class;
using PlaceSoft.Eltit.Class.clases;
using PlaceSoft.Eltit.Functions ;
using PlaceSoft.Eltit.Functions.clases;
using PlaceSoft.Eltit.Handler;

namespace SchoolManagementAdmin
{
    public partial class frmMain : Form
    {
        private static readonly log4net.ILog log =
log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        MysqlConnet cnn = new MysqlConnet();
        string optionsDirection = "down";
        string toastDirection = "down";
        string rightDirection = "right";

        //For animated panels timeout
        int optionsTimeOut = 0;

        int optionsX;
        int optionsY;
        int rightX;
        int rightY;
        Funciones fun = new Funciones("eltit_");


        public frmMain()
        {
            InitializeComponent();
        }


        //method to set fullscreen
        private void setFullScreen()
        {
            int x = Screen.PrimaryScreen.Bounds.Width;
            int y = Screen.PrimaryScreen.Bounds.Height;
            Location = new Point(0, 0);
            Size = new Size(x, y);
        }

        //method to set the position of the main panel that holds the controls to center of the form.
        private void setMainPanelPosition()
        {
            int mX = (Width - pnlMain.Width) / 2;
            int mY = (Height - pnlMain.Height) / 2;
            pnlMain.Location = new Point(mX, mY);
        }

        //method to set the initial location of the options panel.
        private void setOptionsPanelPosition()
        {
            int x = Width;
            optionsX = 0;

        }

        private void setRightOptionsPanelPosition()
        {

        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            int numero = 0;
            string status = "";

            Inicial config = new Inicial();

            //status = config.VerificaServicio("MySQL");
            //if (status != "Running")
            //{                
            //    System.Threading.Thread.Sleep(15000);
            //}

            config.CargaConfiguracion();

            CargaClientes();
            // RUTINA QUE CARGA LA VERIFICACION DE FOLIOS DE FACTURAS GUIAS Y NC
            this.CargaVerificacionDeFolios();
            // RUTINA QUE CARGA LA VERIFICACION DE FOLIOS DE FACTURAS GUIAS Y NC
            //this.CargaCerificacionDeFoliosBoletas();
            lblNombreLocal.Text = "SISTEMAS INFORMÁTICOS PLACESOFT SPA";
            //setFullScreen();
            setOptionsPanelPosition();
            setRightOptionsPanelPosition();
            setMainPanelPosition();


            DateTime buildDate = new FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).LastWriteTime;
            //lblversion.Text = buildDate.ToString();
            /****************** REGIÓN QUE VERIFICA SI EL SISTEMA TIENE REGISTROS PENDIENTES *******/
            timerSync.Interval = Convert.ToInt32(Inicial.G_INTERVAL);
            timerSync.Start();
            txtInfo.Text = "-> Iniciando ciclo de Automatización a las " + DateTime.Now + " ..." + Environment.NewLine;
            txtInfo.Refresh();
            System.Threading.Thread.Sleep(1000);



        }

        private double VerificaFolios(string xTipo, string xLocal, string xCliente, string xBase, string xServer, string xUser, string xPass)
        {
            LocalesClass lo = new LocalesClass();
            MySqlDataReader dr = lo.GetFoliosCriticosFacturas(xTipo, xCliente, xLocal, xServer, xBase, xUser, xPass);
            double hasta = 0;
            double ultimo = 0;
            double disponibles = 0;
            if (dr.HasRows == true)
            {
                if (dr.Read())
                {
                    hasta = Convert.ToDouble(dr["hasta"].ToString());
                    ultimo = Convert.ToDouble(dr["ultimo"].ToString());

                    disponibles = (hasta - ultimo);
                }
            }

            dr.Close();
            lo.CerrarTransaccion();

            return disponibles;
        }

        private double VerificaFoliosBoletas(string xTipo, string xLocal, string xCliente, string xBase, string xServer, string xUser, string xPass, string xCaja)
        {
            LocalesClass lo = new LocalesClass();
            MySqlDataReader dr = lo.GetFoliosCriticosBoletas(xTipo, xCliente, xLocal, xServer, xBase, xUser, xPass, xCaja);
            double hasta = 0;
            double ultimo = 0;
            double disponibles = 0;

            if (dr.HasRows == true)
            {
                if (dr.Read())
                {
                    hasta = Convert.ToDouble(dr["hasta"].ToString());
                    ultimo = Convert.ToDouble(dr["ultimo"].ToString());

                    disponibles = (hasta - ultimo);
                }
            }

            dr.Close();
            lo.CerrarTransaccion();

            return disponibles;
        }

        private void CargaClientes()
        {
            LocalesClass loc = new LocalesClass();
            MySqlDataReader dr = loc.GetLocalByCodigo();
            int i = 0;
            object img = null;
            Inicial fu = new Inicial();

            if (dr.HasRows == true)
            {
                while (dr.Read())
                {

                    if (fu.PingToHost(dr["servidor_ventas"].ToString()) == true)
                    {
                        img = Properties.Resources.green_fw;

                    }
                    else
                    {
                        img = Properties.Resources.yelow;
                    }


                    gvPagos.Rows.Add(dr["codigo"].ToString(), dr["nombrelocal"].ToString(), dr["servidor_ventas"].ToString(), "",
                                    img, 0, 0, 0, 0, dr["rut"].ToString());
                    i++;



                }
            }

            dr.Close();
            loc.CerrarTransaccion();
        }

        private void CargaVerificacionDeFolios()
        {
            int i = 0;
            double caf_33 = 0;
            double caf_61 = 0;
            double caf_39 = 0;
            double caf_52 = 0;
            string base_dte = "";
            string prefijo = "eltit_";
            Locales localClass = new Locales(Inicial.G_SERVIDOR, Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
            Inicial fu = new Inicial();

            for (i = 0; i <= gvPagos.Rows.Count - 1; i++)
            {
                string cod_local = gvPagos.Rows[i].Cells[0].Value.ToString(); //CODIGO LOCAL
                string cliente = gvPagos.Rows[i].Cells[1].Value.ToString();// nombre local
                string server = gvPagos.Rows[i].Cells[2].Value.ToString(); // servidor de ventas
                string rut = gvPagos.Rows[i].Cells[9].Value.ToString();
                base_dte = "eltit_fae" + cod_local;
                caf_39 = 0;
                caf_33 = 0;
                caf_52 = 0;
                caf_61 = 0;

                localClass.getLocalDTE(cod_local);
                Inicial fun = new Inicial();
                if (fun.PingToHost(server) == true)
                {
                    fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[5], Color.White, Color.Black, new Font("Arial", 8, FontStyle.Bold));
                    fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[6], Color.White, Color.Black, new Font("Arial", 8, FontStyle.Bold));
                    fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[7], Color.White, Color.Black, new Font("Arial", 8, FontStyle.Bold));
                    fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[8], Color.White, Color.Black, new Font("Arial", 8, FontStyle.Bold));

                    lblcurrent.Text = "Verificando Folios En " + cliente + "[" + cod_local + "]";
                    lblcurrent.Refresh();
                    gvPagos.Rows[i].Cells[4].Value = Properties.Resources.green_fw;

                    caf_33 = this.VerificaFolios("33", cod_local, prefijo, base_dte,
                                         localClass.IP_servidor, localClass.Mysql_user, localClass.Mysql_pass);
                    caf_39 = 0;

                    caf_52 = this.VerificaFolios("52", cod_local, prefijo, base_dte,
                                         localClass.IP_servidor, localClass.Mysql_user, localClass.Mysql_pass);

                    caf_61 = this.VerificaFolios("61", cod_local, prefijo, base_dte,
                                         localClass.IP_servidor, localClass.Mysql_user, localClass.Mysql_pass);

                    /*********** REGION QUE SETEA LAS CELDAS DE LOS CAF RESPECTIVOS ***********/
                    gvPagos.Rows[i].Cells[5].Value = caf_33;
                    gvPagos.Rows[i].Cells[6].Value = 0; // 39
                    gvPagos.Rows[i].Cells[7].Value = caf_52;
                    gvPagos.Rows[i].Cells[8].Value = caf_61;

                    if (caf_33 < Convert.ToDouble(localClass.Caf_critico_33))
                    {
                        fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[5], Color.Red, Color.Black, new Font("Arial", 8, FontStyle.Bold));

                        this.EnviarEmailFolios(localClass.Rut, cod_local + " " + localClass.Nombrelocal, server,
                            "33", Convert.ToDouble(localClass.Caf_critico_33), caf_33);
                    }
                    if (caf_39 < Convert.ToDouble(localClass.Caf_critico_39))
                    {
                        // fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[6], Color.Red, Color.Black, new Font("Arial", 8, FontStyle.Bold));

                    }
                    if (caf_52 < Convert.ToDouble(150))
                    {
                        fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[7], Color.Red, Color.Black, new Font("Arial", 8, FontStyle.Bold));
                        this.EnviarEmailFolios(localClass.Rut, cod_local + " " + localClass.Nombrelocal, server,
                            "52", Convert.ToDouble(150), caf_52);
                    }
                    if (caf_61 < Convert.ToDouble(localClass.Caf_critico_61))
                    {
                        fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[8], Color.Red, Color.Black, new Font("Arial", 8, FontStyle.Bold));
                        this.EnviarEmailFolios(localClass.Rut, cod_local + " " + localClass.Nombrelocal, server,
                            "61", Convert.ToDouble(localClass.Caf_critico_61), caf_61);
                    }
                    if (cod_local=="42")
                    {
                        string xd;
                        xd = cod_local.ToString();
                    }
                }
                else
                {
                    gvPagos.Rows[i].Cells[4].Value = Properties.Resources.yelow;

                }


            }
        }

        private void CargaCerificacionDeFoliosBoletas()
        {
            int i = 0;
            double caf_33 = 0;
            double caf_61 = 0;

            double caf_52 = 0;
            string base_dte = "";

            Locales localClass = new Locales(Inicial.G_SERVIDOR, Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
            LocalesClass loc = new LocalesClass();

            Inicial fu = new Inicial();
            DataTable dt;

            for (i = 0; i <= gvPagos.Rows.Count - 1; i++)
            {
                string cod_local = gvPagos.Rows[i].Cells[0].Value.ToString(); //CODIGO LOCAL
                string cliente = gvPagos.Rows[i].Cells[1].Value.ToString();// nombre local
                string server = gvPagos.Rows[i].Cells[2].Value.ToString(); // servidor de ventas
                string rut = gvPagos.Rows[i].Cells[9].Value.ToString();
                base_dte = "eltit_fae" + cod_local;


                localClass.getLocalDTE(cod_local);
                Inicial fun = new Inicial();


                if (fun.PingToHost(server) == true)
                {
                    fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[6], Color.White, Color.Black, new Font("Arial", 8, FontStyle.Bold));

                    lblcurrent.Text = "Verificando Folios 39 En " + cliente + "[" + cod_local + "]";
                    lblcurrent.Refresh();
                    gvPagos.Rows[i].Cells[4].Value = Properties.Resources.green_fw;

                    PopBoletas frm = new PopBoletas();
                    frm.LOCAL_ACTIVO = cod_local;
                    frm.lblNombreLocal.Text = "[" + cod_local + "]" + cliente;
                    frm.lblNombreLocal.Refresh();
                    frm.ShowDialog();
                    frm.Dispose();

                }
                else
                {
                    gvPagos.Rows[i].Cells[4].Value = Properties.Resources.yelow;

                }


            }


        }

        private void EnviarEmailFolios(string xRut, string xLocal, string xServidor, string xTipo, double xCritico, double xCurrCaf)
        {

            string htmlString = @"<html>";
            double total = 0;
            htmlString = htmlString + "<body>";
            htmlString = htmlString + "<img src='http://www.placesoft.cl/images/eltit/header_eltit.png' border='0'  />";
            htmlString = htmlString + "<p>Se Han detectado folios insuficientes para el Local " + xLocal + "</p>";
            htmlString = htmlString + "--------------------------------------------<br>";
            htmlString = htmlString + " Rut      " + Convert.ToDouble(xRut.Substring(0, 9)).ToString() + "-" + xRut.Substring(9, 1) + "<br>";
            htmlString = htmlString + " Local    " + xLocal + "<br>";
            htmlString = htmlString + " Servidor " + xServidor + "<br>";
            htmlString = htmlString + " Tipo Doc " + xTipo + "<br>";
            htmlString = htmlString + " Critico[" + xCritico + "] Quedan[" + xCurrCaf + "]  <br>";
            htmlString = htmlString + "--------------------------------------------<br>";
            htmlString = htmlString + "<p>Enviado el " + DateTime.Now.ToString("yyyy-MM-dd") + " a las " + DateTime.Now.ToString("HH:mm:ss tt") + "</p>";
            htmlString = htmlString + "<img src='http://www.placesoft.cl/images/eltit/footer_eltit.png' border='0'  />";
            htmlString = htmlString + "</body>";
            htmlString = htmlString + " </html>";

            Inicial.EnviarEmail(Inicial.G_CORREO_SOPORTE_PRINCIPAL, Inicial.G_CORREO_SOPORTE_COPIA, "CAF " + xTipo, "CAF CRITICOS EN " + xLocal, htmlString);

        }

        private void timerSync_Tick(object sender, EventArgs e)
        {

            if (Inicial.G_ERROR == true)
            {
                timerSync.Stop();
                txtInfo.Text += DateTime.Now + " -> ERROR Encontrado a las el servicio fue detenido " + Environment.NewLine;
                txtInfo.Refresh();
            }

            //try
            //{4

            if (rbCentralizacion.CheckState == CheckState.Checked)
            {
                
                string minuto = "";
                int hora = 0;

                hora = Convert.ToInt32(DateTime.Now.ToString("HH:mm:ss tt").Substring(0, 2));
                minuto = DateTime.Now.ToString("HH:mm:ss tt").Substring(3, 2);


                if (hora >= 9 && hora <= 23 || chForzar.Checked == true )
                {
                   // ProcesaCentralizacion();
                    System.Threading.Thread.Sleep(1000);
                }


                    //TIMER PROGRAMADO PARA VERIFICACION DE FOLIOS CADA UNA HORA '00'              





                    if (minuto == "00" && (hora >= 9 && hora <= 23))
                {
                    timerSync.Stop();
                    CargaVerificacionDeFolios();
                    System.Threading.Thread.Sleep(1000);
                    this.CargaCerificacionDeFoliosBoletas();


                    /**************** REGION QUE REVISA LAS BANDEJAS DE CORREOS DE EMPRESAS ********/
                    txtInfo.Text += DateTime.Now + " -> INICIO DE TAREA IMPORTACION DE CORREOS <- " + Environment.NewLine;
                    txtInfo.SelectionStart = txtInfo.Text.Length;
                    txtInfo.ScrollToCaret();
                    txtInfo.Refresh();


                    //if (Directory.Exists(@"\\192.168.4.6\fae\"))
                    //{
                    //    timerSync.Stop();
                    //    this.GeneraImporteDeCorreos();
                    //    System.Threading.Thread.Sleep(1000);
                    //    timerSync.Start();
                    //    txtInfo.Text += DateTime.Now + " -> IMPORTACION DE CORREOS FINALIZADA <- " + Environment.NewLine;
                    //    txtInfo.SelectionStart = txtInfo.Text.Length;
                    //    txtInfo.ScrollToCaret();
                    //    txtInfo.Refresh();
                    //}
                    //else
                    //{
                    //    //MessageBox.Show("No se encontró el directorio o la ruta para los temporales Recibidos [192.168.4.6\fae]");
                    //    txtInfo.Text += DateTime.Now + " -> NO SE ENCONTRO LA RUTA PARA LA IMPORTACIÓN DE CORREOS <- " + Environment.NewLine;
                    //    txtInfo.SelectionStart = txtInfo.Text.Length;
                    //    txtInfo.ScrollToCaret();
                    //    txtInfo.Refresh();
                    //}


                    timerSync.Start();

                }

                if (minuto.Substring(1,1) == "0" || minuto.Substring(1, 1) == "9" )
                { 
                    txtInfo.Text += DateTime.Now + " -> INICIO DE TAREA IMPORTACION DE CORREOS <- " + Environment.NewLine;
                    txtInfo.SelectionStart = txtInfo.Text.Length;
                    txtInfo.ScrollToCaret();
                    txtInfo.Refresh();

                    timerSync.Stop();
                    if (Directory.Exists(@"\\192.168.4.6\fae\"))
                    {
                        timerSync.Stop();
                        this.GeneraImporteDeCorreos();
                        System.Threading.Thread.Sleep(1000);
                        timerSync.Start();
                    }
                    else
                    {
                        //MessageBox.Show("No se encontró el directorio o la ruta para los temporales Recibidos [192.168.4.6\fae]");
                        txtInfo.Text += DateTime.Now + " -> NO SE ENCONTRO LA RUTA PARA LA IMPORTACIÓN DE CORREOS <- " + Environment.NewLine;
                        txtInfo.SelectionStart = txtInfo.Text.Length;
                        txtInfo.ScrollToCaret();
                        txtInfo.Refresh();

                    }
                    txtInfo.Text += DateTime.Now + " -> FIN DE TAREA IMPORTACION DE CORREOS <- " + Environment.NewLine;
                    txtInfo.SelectionStart = txtInfo.Text.Length;
                    txtInfo.ScrollToCaret();
                    txtInfo.Refresh();

                    timerSync.Start();
                }

            }



            if (rbCorreos.CheckState == CheckState.Checked)
            {
                //ProcesaCorreos(); 29232
            }


            //}
            //catch(Exception ex)
            //{
            //    log.Error(ex);
            //    timerSync.Stop();
            //    MessageBox.Show("Error:" + ex.Message.ToString());
            //}

        }

        private void GeneraImporteDeCorreos()
        {
            PopLeeCorreos frm = new PopLeeCorreos();             
            frm.ShowDialog();
        }

        private void ProcesaCentralizacion()
        {
            txtInfo.SelectionStart = txtInfo.TextLength;
            txtInfo.SelectionLength = 0;
            txtInfo.ScrollToCaret();
            Inicial fu = new Inicial();
            string minuto = "";
            minuto = DateTime.Now.ToString("hh:mm:ss tt").Substring(4, 1);
            if (minuto == txtMin1.Text || minuto == txtMin2.Text || minuto == txtMin3.Text )
            {
                int i = 0;
                for (i = 0; i <= gvPagos.Rows.Count - 1; i++)
                {
                    string local = gvPagos.Rows[i].Cells[0].Value.ToString();
                    string cliente = gvPagos.Rows[i].Cells[1].Value.ToString();                    
                    string server = gvPagos.Rows[i].Cells[2].Value.ToString();
                    string rut = gvPagos.Rows[i].Cells[9].Value.ToString();

                    fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[1], Color.YellowGreen, Color.White, new Font("Arial", 8, FontStyle.Bold));
                    gvPagos.Refresh();
                   
                    if (fun.PingToHost(server) == true)
                    {
                        lblcurrent.Text = "Verificando " + cliente;
                        lblcurrent.Refresh();
                        gvPagos.Rows[i].Cells[4].Value = Properties.Resources.green_fw;
                        this.gvPagos.TableElement.UpdateView();
                        this.IniciaProcesoDatos(cliente, server, local, rut);
                    }
                    else
                    {
                        gvPagos.Rows[i].Cells[4].Value = Properties.Resources.yelow;
                        this.gvPagos.TableElement.UpdateView();
                    }

                    if (Inicial.G_ERROR == true)
                    {
                        timerSync.Stop();
                        txtInfo.Text += DateTime.Now + " -> ERROR Encontrado en " + cliente + " el servicio fue detenido " + Environment.NewLine;
                        txtInfo.Refresh();
                    }

                    fu.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[1], Color.White, Color.Black, new Font("Arial", 8, FontStyle.Bold));
                    gvPagos.Refresh();

                }
            }
        }

        private int IniciaProcesoDatos(string xCliente, string xServer, string xLocal, string xrut)
        {            
            Inicial fun = new Inicial();
            txtInfo.Text += DateTime.Now +" -> Procesando Cliente " + xCliente + " en " + xServer + " ..." + Environment.NewLine;
            txtInfo.Refresh();
            if(fun.PingToHost(xServer) == true )
            {
                txtInfo.AppendText(DateTime.Now +"    - Conexión Success con " + xServer + " " + Environment.NewLine);
                txtInfo.Refresh();
               
            }else
            {
                txtInfo.AppendText(DateTime.Now + "    - No se pudo establecer conexión con " + xCliente  + " " + Environment.NewLine);
                txtInfo.Refresh();
                return 0;
            }
            pictureBox2.Refresh();
            System.Threading.Thread.Sleep(100);
            //string minuto = "";
            //minuto = DateTime.Now.ToString("hh:mm:ss tt").Substring(4, 1);

           
                this.InciciaTraspaso(xCliente,xServer,xLocal, xrut);
                
                txtInfo.AppendText(DateTime.Now + " ========== Ciclo ["+ xCliente +"] Finalizado =========" + Environment.NewLine);
                txtInfo.Refresh();
                System.Threading.Thread.Sleep(2000);
                
            

            return 1;
        }

        private void BuscaCaf(Locales local, string xTipo )
        {
            string filePath = "";
            string xmlString = "";
            string caf_desde = "";
            string caf_hasta = "";
            try
            {
                Caf myCaf = new Caf(local.IP_servidor, local.Mysql_user, local.Mysql_pass);
                MySqlDataReader dr = myCaf.GetCafByCajaLocal(local.Local, "", xTipo);
               

                if (dr.HasRows == true)
                {
                    while (dr.Read())
                    {
                        //filePath = @"C:\PlaceDTE\" + local.Prefijo.Replace("_","") + @"\" + Convert.ToDouble(local.Rut.Substring(0, 9)) + @"\Produccion\Caf\" + local.Local + @"\";
                        //filePath = filePath + string.Format("{0}_{1}_{2}.dat", Convert.ToInt32(xTipo), dr["desde"].ToString(), dr["hasta"].ToString());

                        if (!File.Exists(filePath) || 1 == 1)
                        {
                            FileStream fst;
                            BinaryWriter bw;
                            string tmp_path = @"C:\temp\" + DateTime.Now.Ticks + ".xml";

                            fst = new FileStream(tmp_path, FileMode.OpenOrCreate, FileAccess.Write);
                            bw = new BinaryWriter(fst);

                            string strxml = dr["xml"].ToString().Replace("±	", "");



                            //filePath = @"C:\PlaceDTE\" + local.Prefijo.Replace("_", "") + @"\" + Convert.ToDouble(local.Rut.Substring(0, 9)) + @"\Produccion\Caf\" + local.Local + @"\";
                            //filePath = filePath + string.Format("{0}_{1}_{2}.dat", Convert.ToInt32(xTipo), caf_desde, caf_hasta);

                            Encoding ByteConverter = Encoding.GetEncoding("ISO-8859-1");
                            byte[] textEnBytes = ByteConverter.GetBytes(strxml);

                            bw.Write(textEnBytes);
                            bw.Flush();
                            bw.Close();
                            bw.Dispose();

                            /*************************** SERIALIZAR INICIO Y FIN DE CAF **********************/
                            XmlDocument xmlDoc = xmlDoc = new XmlDocument();
                            string TIPO_XML;

                            xmlDoc.Load(tmp_path);
                            XmlNode node = xmlDoc.DocumentElement.FirstChild;
                            XmlNodeList lstNodos = node.ChildNodes;

                            TIPO_XML = xmlDoc.DocumentElement.Name;
                            string XML_DTE = "";


                            if (TIPO_XML == "AUTORIZACION")
                            {
                                for (int i = 0; i < lstNodos.Count; i++)
                                {

                                    if (lstNodos[i].Name == "DA")
                                    {
                                        XML_DTE = lstNodos[i].OuterXml;
                                        XmlNodeList lstChilds = lstNodos[i].ChildNodes;
                                        for (int j = 0; j < lstChilds.Count; j++)
                                        {
                                            if (lstChilds[j].Name == "RNG")
                                            {
                                                XmlNodeList fist2 = lstChilds[j].ChildNodes;
                                                for (int l = 0; l < fist2.Count; l++)
                                                {
                                                    if (fist2[l].Name == "D")
                                                    {
                                                        caf_desde = fist2[l].InnerText;
                                                    }

                                                    if (fist2[l].Name == "H")
                                                    {
                                                        caf_hasta = fist2[l].InnerText;
                                                    }

                                                }
                                            }
                                        }

                                    }

                                }

                            }

                            /*************************** FIN SERIALIZAR INICIO Y FIN DE CAF **********************/


                            filePath = @"C:\PlaceDTE\" + local.Prefijo.Replace("_", "") + @"\" + Convert.ToDouble(local.Rut.Substring(0, 9)) + @"\Produccion\Caf\" + local.Local + @"\";
                            filePath = filePath + string.Format("{0}_{1}_{2}.dat", Convert.ToInt32(xTipo), caf_desde, caf_hasta);

                            if (!File.Exists(filePath))
                            {
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
                }


                dr.Close();
                myCaf.CerrarTransaccion();
            }
            catch(Exception ex)
            {
                log.Debug("Current: " + filePath);
                log.Error("Err:", ex);
            }
          

        }

        private int InciciaTraspaso(string xCliente, string xServer, string xLocal, string xRut)
        {
            Locales local = new Locales(Inicial.G_SERVIDOR, Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
            local.getLocalDTE(xLocal);
            int registros = 0;

            /*
             *VERIFICA SI LOS CAF ESTAN CREADOS, SI NO, LOS TRAE PARA TIMBRAR CORRECTAMENTE
             */
           
            
                       
            if (local.IP_servidor == "")
            {
                return 0;
            }
            Inicial fun = new Inicial();
            if (fun.PingToHost(local.IP_servidor) == true)
            {
                BuscaCaf(local, "39");
            }
            else
            {
                txtInfo.AppendText(DateTime.Now + "    -> No se pudo conectar al Servidor de Destino " + local.Servidor_destino + Environment.NewLine);
                txtInfo.Refresh();
                return 0;
            }


            txtInfo.AppendText(DateTime.Now + "    -> Verificando Registros Pendientes en " + local.IP_servidor + Environment.NewLine);
            txtInfo.Refresh();

            // Sincroniza SYNC = new Sincroniza(local, "placesof");
            //int nro = SYNC.GetRowsCount("01");
            
            Documentos doc = new Documentos(lblServidorPrincipal.Text, cnn.getPrincipalMysqlRoot(), cnn.getPrincipalMysqlPass());
            int nro = 0;
            DataTable dt = new DataTable();
            MySqlDataReader dr = doc.GetDocumentosBoletaPendientes(xLocal, 200);
             
            if(dr.HasRows == true)
             {
                dt.Load(dr);
                dr.Close();
                nro = dt.Rows.Count;
            }
            else
            {
                nro = 0;
            }

            doc.CerrarTransaccion();                                 
            
            txtInfo.AppendText(DateTime.Now + "    -> Se encontraron " + nro + " Registros en [" + local.Local + "]" + Environment.NewLine);
            txtInfo.Refresh();
            if (nro <= 0)
            {
                txtInfo.AppendText(DateTime.Now + "    ===== Fin Cliclo "+ local.Razon_social +" ===== "  + Environment.NewLine);
                return 0;
            }
      
            txtInfo.AppendText(DateTime.Now + "    -> Verificando conexión con el Host de destino " + lblIpHostDestino.Text + Environment.NewLine);
            txtInfo.Refresh();
            if (fun.PingToHost(lblServidorPrincipal.Text) == true)
            {
                txtInfo.AppendText(DateTime.Now + "    -> Conexión Success Con el Host " + lblIpHostDestino.Text + " " + Environment.NewLine);
                txtInfo.Refresh();
            }
            else
            {
                txtInfo.AppendText(DateTime.Now + "    -> No se pudo establecer conexión con " + lblIpHostDestino.Text + " " + Environment.NewLine);
                txtInfo.Refresh();
                return 0;
            }

            registros = this.GeneraBoletaDocumento(xCliente, local, dt);
            txtInfo.AppendText(DateTime.Now + "    -> Se insertaron " + registros + " " +   Environment.NewLine);
            //txtInfo.AppendText(DateTime.Now + "    ===== Fin Cliclo " + cliente.Prefijo + " ===== " + Environment.NewLine);
            txtInfo.Refresh();
            return registros;
        }
             
       
        private int GeneraBoletaDocumento(string xCliente, Locales xLocal, DataTable xDT)
        {
            Documentos doc = new Documentos(lblServidorPrincipal.Text, cnn.getPrincipalMysqlRoot(), cnn.getPrincipalMysqlPass());
            MySqlDataReader dr;
            int salida = 0;
            DataTable dtDetalle = new DataTable();
            string tipo = "";
            string numero = "";
            string caja = "";
            string fecha = "";
            string rut = "";
            int xFolioSII = 0;
            int cuenta = 0;
            int retorno = 0;

            foreach (DataRow row in xDT.Rows)
            {
                tipo    = row["tipo"].ToString();
                numero  = row["numero"].ToString();
                caja    = row["caja"].ToString();
                fecha   = row["fecha"].ToString();
                rut     = row["rut"].ToString();
                xFolioSII = Convert.ToInt32(row["foliosii"].ToString());

                fecha = fun.FechaMysql(fecha);
                dtDetalle = doc.GetDoucumentosDetalleByTipoCajaNroInternoFechaLocal(xLocal.Local, tipo, numero, caja, fecha);

               retorno = this.GeneraXMLBoleta(xCliente, xLocal, tipo, fecha,caja,xFolioSII,numero,dtDetalle);
                cuenta = cuenta + retorno;

            }
                                          

            return cuenta;
        }

        private int GeneraXMLBoleta(string xCliente, Locales xLocal,string tipo,string xFecha,string xCaja,int xFolioSII,string xNumero, DataTable xDT)
        {
            int salida = 0;

            try
            {
                Handler handler;
                handler = new Handler(xCliente);

                string xtipo = "";
                string base_venta = "eltit_ventas" + xLocal.Local;
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

                if (tipo == "BV")
                {
                    xtipo = "BV";
                    tipoFiscal = "39";
                    typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica;
                }
                if (tipo == "BE")
                {
                    xtipo = "BE";
                    tipoFiscal = "41";
                    typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta;
                }
                log.Debug("Generando Local:" + xLocal.Local + " NRO " + xNumero + " Caja " + xCaja + " FECHA " + xFecha);

                foreach (DataRow row in xDT.Rows)
                {
                    rut_venta = row["rut"].ToString();

                    //formaPago = dr["formapago"].ToString();
                    ItemBoleta item = new ItemBoleta();
                    codigoArticulo = row["codigo"].ToString();
                    precio = Convert.ToInt32(row["precio"].ToString());
                    cantidad = Convert.ToDouble(row["cantidad"].ToString());
                    dcto = Convert.ToDouble(row["descuentopesos"]);
                    DctoLinea = dcto;
                    totalLinea = Convert.ToInt32(Math.Round(cantidad * precio, 4) - DctoLinea);

                    item.Nombre = row["descripcion"].ToString();
                    item.Cantidad = cantidad;
                    item.Codigo = codigoArticulo;
                    if (tipo == "BE")
                    {
                        item.Afecto = false;
                    }
                    else
                    {
                        item.Afecto = true;
                    }

                    item.Precio = Convert.ToInt32(Math.Round(precio, 4));
                    item.Porce_Descuento = Convert.ToInt32(row["descuento"].ToString());
                    item.Monto_Descuento = Convert.ToInt32(dcto);

                    item.Total = totalLinea;
                    item.UnidadMedida = string.Empty;
                    items.Add(item);


                }


                /********** GET ULTIMO FOLIO FISCAL DTE ****************/
                //DTE DTE = new DTE(xLocal.IP_servidor, xLocal.Mysql_user, xLocal.Mysql_pass);
                //int FOLIO_CAF = DTE.GetUltimoFolioDTEByLocalCaja(xLocal.Local, tipoFiscal, xCaja, base_venta);
                //FOLIO_CAF = GetUltimoFolio(tipoFiscal, FOLIO_CAF, xCaja);
                if(rut_venta == "" || rut_venta.Length != 10)
                {
                    rut_venta = "0666666666";
                }
                int FOLIO_CAF = xFolioSII;
                handler.tipo = typeDTE;
                handler.casoPruebas = string.Empty;
                handler.Folio = (int)FOLIO_CAF;
                handler.idDte = "ENVIOFOLIO_" + FOLIO_CAF + "T" + xtipo; // "DTE" + DOC_TIPO + "F" + FOLIO_CAF;
                handler.rutcliente = Convert.ToDouble(rut_venta.Substring(0, 9)) + "-" + rut_venta.Substring(9, 1);
                handler.rutCertificado = xLocal.Rut_certificado;
                handler.nombreCertificado = xLocal.Nombre_certificado;
                handler.fechaResolucion = Convert.ToDateTime(xLocal.Fecha_resolucion);
                handler.numero_resolucion = xLocal.Numero_resolucion;
                handler.emisor_rut = Convert.ToDouble(xLocal.Rut.Substring(0, 9)) + "-" + xLocal.Rut.Substring(9, 1);
                handler.emisor_razon_social = xLocal.Razon_social;
                handler.emisor_giro = xLocal.Giro;
                handler.emisor_comuna = xLocal.Comuna_empresa;
                handler.emisor_ciudad = xLocal.Comuna_empresa;
                handler.emisor_direccion = xLocal.Direccion_empresa;
                handler.cod_sucursal_sii = xLocal.Codigo_sucursal_sii;
                handler.fechaEmision = Convert.ToDateTime(xFecha);
                 var dte = handler.GenerateDTEBoleta();

                handler.GenerateDetails(dte, items);
                handler.ReferenciasBoleta(dte);

                if(dte.Documento.Encabezado.Totales.MontoTotal > 0)
                {
                    //if (xNumero == "0006601487")
                    //{
                    //    xNumero = "0006601487";
                    //}
                    var path = handler.TimbrarYFirmarXMLDTE(dte, @"C:\PlaceDTE\eltit\" + Convert.ToDouble(xLocal.Rut.Substring(0, 9)) + @"\Produccion\Caf\" + xLocal.Local + @"\");

                    if (File.Exists(path))
                    {
                        FileInfo fi = new FileInfo(path);
                        //string destino = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\xml\" + FuncionesClass.G_LOCAL + @"\DTE39F" + FOLIO_CAF + ".xml";
                        string destino = @"C:\PlaceDTE\eltit\" + Convert.ToDouble(xLocal.Rut.Substring(0, 9)) + @"\Produccion\xml\" + xLocal.Local.Substring(0, 2) + @"\DTE" + tipoFiscal + "F" + FOLIO_CAF + ".xml";
                        fi.CopyTo(destino, true);
                        //System.Threading.Thread(1000);
                        string XML = File.ReadAllText(destino, Encoding.GetEncoding("ISO-8859-1"));

                        this.GrabaXMLDesarrollo(xLocal, xNumero, xCaja, tipoFiscal, xtipo, XML, xFecha, FOLIO_CAF, rut_venta, dte.Documento.Encabezado.Totales.MontoTotal);
                        salida = 1;
                        GrabaXMLInternet(xLocal, xNumero, xCaja, tipoFiscal, xtipo, XML, xFecha, FOLIO_CAF, rut_venta, dte.Documento.Encabezado.Totales.MontoTotal);

                        this.MarcaDocumentoSubido(xLocal.Local, xNumero, xCaja, xtipo, xFecha, DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"));
                        
                        System.IO.File.Delete(path);
                    }
                }
                else
                {
                    this.MarcaDocumentoSubido(xLocal.Local, xNumero, xCaja, tipo, xFecha, "INI CAJA");
                }

              
            }
            catch(Exception ex)
            {
                log.Error("Error: Empresa " + xLocal.Codigo_contable + " Local: " + xLocal.Local + " NRO:" + xNumero + " " + xCaja , ex);
              
                if (ex.Message.ToString().Contains("Duplicate"))
                {
                    this.MarcaDocumentoSubido(xLocal.Local, xNumero, xCaja, tipo, xFecha, "Duplicate");
                }
                if (ex.Message.ToString().Contains("CAF"))
                {
                    this.MarcaDocumentoSubido(xLocal.Local, xNumero, xCaja, tipo, xFecha, "NO CAF");
                }
                if (ex.Message.ToString().Contains("Referencia"))
                {
                    this.MarcaDocumentoSubido(xLocal.Local, xNumero, xCaja, tipo, xFecha, "ERR REF");
                }
                salida = 0;
            }

            return salida;

        }

        private void GrabaXMLDesarrollo(Locales local, string xNroInterno, string xCaja, string xTipoFiscal, string xTipoInterno,
                                      string XML, string xFecha, int xFolioSII, string xRutCliente, double xTotal)
        {
            DTEClass dte = new DTEClass(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);
            dte.GrabaXML(local.Codigo_contable, local.Local, Convert.ToDouble(local.Rut.Substring(0, 9)).ToString(), xTipoFiscal, xFolioSII,
                        xFecha, xTipoInterno, xNroInterno, xCaja, XML, xRutCliente, xTotal);
        }

        private void GrabaXMLInternet(Locales local, string xNroInterno, string xCaja, string xTipoFiscal,string xTipoInterno, 
                                       string XML, string xFecha, int xFolioSII,string xRutCliente, double xTotal)
        {
            DTEHost dte = new DTEHost(lblIpHostDestino.Text, cnn.getHostMysqlRoot(), cnn.getHostMysqlPass());
            dte.GrabaXML(local.Codigo_contable, local.Local, Convert.ToDouble(local.Rut.Substring(0, 9)).ToString(), xTipoFiscal, xFolioSII,
                        xFecha, xTipoInterno, xNroInterno, xCaja, XML, xRutCliente, xTotal);


        }
        private void MarcaDocumentoSubido(string xLocal, string xNumeroInterno, string xCaja, string xTipo, string xFecha, string xStatus)
        {
            Documentos dc = new Documentos(lblServidorPrincipal.Text, cnn.getPrincipalMysqlRoot(), cnn.getPrincipalMysqlPass());
            dc.MarcaBoletaSubidaInternet(xLocal, xNumeroInterno, xCaja, xTipo, xFecha, xStatus);

        }

        private void frmMain_MouseMove(object sender, MouseEventArgs e)
        {
         
          
        }

        private void Options_Tick(object sender, EventArgs e)
        {
            if (optionsTimeOut < 1000)
            {
                optionsTimeOut++;
            }
            if (optionsTimeOut == 1000)
            {
                if (optionsDirection == "up")
                {
                    optionsDirection = "down";
                }
            }
          
          
        }

        private void pbHide_Click(object sender, EventArgs e)
        {
            optionsDirection = "down";
        }

        private void pbExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pbHome_Click(object sender, EventArgs e)
        {
            rightDirection = "right";
        }

         private void CargaReloj()
        {

            DateTime thisDay = DateTime.Today;
            CultureInfo ci = CultureInfo.InvariantCulture;

            lbldia.Text = thisDay.ToString("dddd").ToUpper();
            lblnumero.Text = thisDay.ToString("dd");
            lblmes.Text = thisDay.ToString("MMMM").ToUpper() + ", " + thisDay.ToString("yyyy");
            lblhora.Text = DateTime.Now.ToString("H:mm:ss");
        }
  
        private void timerReloj_Tick(object sender, EventArgs e)
        {
            CargaReloj();
        }

        private void pbAccounts_Click(object sender, EventArgs e)
        {
            timerReloj.Stop();
            timerReloj.Enabled = false;
            Application.Exit();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            
            if(timerSync.Enabled == true)
            {
                timerSync.Enabled = false;
                txtInfo.AppendText("[STOP]" + DateTime.Now + " Automatización Pausada." + Environment.NewLine);
                pnlAccounts.Enabled = true;
            }
            else
            {
                txtInfo.AppendText("[START]" + DateTime.Now + " Automatización Iniciada." + Environment.NewLine);
                timerSync.Enabled = true;
                pnlAccounts.Enabled = false;
            }
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
               
                Hide();
                notifyIcon1.Visible = true;
                //notifyIcon1.Icon = SystemIcons.Information;
                notifyIcon1.BalloonTipText = "Esta aplicación se está ejecutando en segundo plano.";
                notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon1.ShowBalloonTip(100);

            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }

        private void ProcesaCorreos()
        {

            string minuto = "";
            minuto = DateTime.Now.ToString("hh:mm:ss tt").Substring(4, 1);
            if (minuto == txtMin1.Text  || minuto == txtMin3.Text )
            {
                //chForzar.Checked = false;P
                //timerSync.Enabled = false;
                int i = 0;
                ClienteDTE cli;
                for (i = 0; i < gvPagos.Rows.Count; i++)
                {
                    string cliente = gvPagos.Rows[i].Cells[1].Value.ToString();
                    string local = gvPagos.Rows[i].Cells[2].Value.ToString();
                    string server = gvPagos.Rows[i].Cells[3].Value.ToString();
                    string rut = gvPagos.Rows[i].Cells[9].Value.ToString();
                    lblcurrent.Text = "Verificando " + cliente;
                    lblcurrent.Refresh();
                    cli = new ClienteDTE(cliente, rut, "00");
                    Read_Emails(cli);
                }
            }  
            
           
        }
    
        protected List<Email> Emails
        {
            get;  set; 
              
        }


        private void Read_Emails(ClienteDTE cliente)
        {
            Pop3Client pop3Client;
            
            if(cliente.Smtp_intercambio == "" || cliente.Smtp_direccion == "" || cliente.Smtp_clave == "")
            {
                return;
            }
            txtInfo.AppendText(DateTime.Now + " ==> Abriendo Bandeja cliente " + cliente.Prefijo + " <==" + Environment.NewLine);
            txtInfo.Refresh();
            pop3Client = new Pop3Client();
            pop3Client.Connect(cliente.Smtp_intercambio, 110, false);
            pop3Client.Authenticate(cliente.Smtp_direccion, cliente.Smtp_clave, AuthenticationMethod.UsernameAndPassword);
            //pop3Client.Connect("correo.eltit.cl", 7110, false);
            //pop3Client.Authenticate("eltit_dte_08@eltit.cl", "Eltit.2020.", AuthenticationMethod.UsernameAndPassword);


            txtInfo.AppendText(DateTime.Now + " ==> Autenticación success con " + cliente.Smtp_direccion + " <==" + Environment.NewLine);
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


                List<MessagePart> attachments = message.FindAllAttachments();
                bool grabado = false;
                foreach (MessagePart attachment in attachments)
                {
                    string extension = Path.GetExtension(attachment.FileName);
                    if (extension == ".xml" || extension == ".XML")
                    {
                        string filename = string.Format(@"{0}{1}_{2}{3}", @"C:\Temp\", Path.GetFileNameWithoutExtension(attachment.FileName), "", Path.GetExtension(attachment.FileName));
                        attachment.Save(new FileInfo(filename));
                        // LeeXML2(message.Headers.From.Address, attachment.FileName, filename);
                      grabado =  LeeXML2(cliente.Rut, cliente.Prefijo, filename, message.Headers.From.Address, attachment.FileName, message.Headers.Date);

                    }
                                           

                }
                

                if(grabado == true || message.Headers.Subject.Contains("Mail Delivery") || message.Headers.Subject.Contains("Mail delivery") || message.Headers.Subject.Contains("Warning") )
                {
                    pop3Client.DeleteMessage(i);
                    counter++;
                    txtInfo.AppendText(DateTime.Now + " => Importando  Correo From [" + message.Headers.From + "] <==" + Environment.NewLine);
                    txtInfo.Refresh();
                }
               
              if (counter > 150)
                {
                    break;
                }
            }
            pop3Client.Disconnect();
            txtInfo.AppendText(DateTime.Now + " ********* Se importaron  [" + counter+ "] a cliente "+ cliente.Prefijo +" *********" + Environment.NewLine);
            txtInfo.Refresh();
        }

 

        private bool LeeXML2(string xRutCliente,string xPRefijo, string xml, string xCorreo, string xNombreArchivo, string xFechaCorreo)
        {
            XmlDocument xmlDoc = xmlDoc = new XmlDocument();
            bool salida = false;
            xmlDoc.Load(xml);
            XmlNode node = xmlDoc.DocumentElement.FirstChild;
            XmlNodeList lstNodos = node.ChildNodes;
            string XML_DTE = "";
            string rutEmisor = "";
            string razonSocialEmisor ="";
            string rutReceptor = "";
            string tipoDoc = "";
            string numeroDoc = "";
            string FechaEmision = "";
            double Totaldoc = 0;
            string TIPO_XML = "";



            try
            {

                TIPO_XML = xmlDoc.DocumentElement.Name;
                string RutFormat = Convert.ToDouble(xRutCliente.Substring(0, 9)) + "-" + xRutCliente.Substring(9, 1);
                if (TIPO_XML == "EnvioDTE" || TIPO_XML == "EnvioBOLETA" )
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
                                this.GrabaDteProveedor(xPRefijo, xRutCliente, "00", rutEmisor, razonSocialEmisor, tipoDoc, numeroDoc, FechaEmision,
                                                                          xCorreo, Totaldoc, XML_DTE, xNombreArchivo);
                                salida = true;
                            }
                            else
                            {
                                salida = false;
                            }

                        }// END IF NODO DTE
                    }
                }
                if (TIPO_XML == "RespuestaDTE"|| TIPO_XML == "EnvioRecibos" )
                {
                    for (int i = 0; i < lstNodos.Count; i++)
                    {
                        if (lstNodos[i].Name == "Caratula")
                        {
                            XmlNodeList nodoresponde = ((XmlElement)lstNodos[0]).GetElementsByTagName("RutResponde");
                            XmlNodeList nodorecibe   = ((XmlElement)lstNodos[0]).GetElementsByTagName("RutRecibe");
                            rutEmisor = nodoresponde[0].InnerXml;
                            rutReceptor = nodorecibe[0].InnerXml;

                            if (RutFormat != rutReceptor)
                            {
                                return false;
                            }
                            else
                            {
                                this.GrabaXMLIntercambio(xPRefijo, xRutCliente, "00", TIPO_XML, xNombreArchivo, xCorreo, xmlDoc.InnerXml, rutEmisor, rutReceptor, xFechaCorreo);
                                salida = true;
                            }

                        }
                    }

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
                            //XmlNodeList nodoemisor = ((XmlElement)lstNodos[0]).GetElementsByTagName("RUTEMISOR");
                            //XmlNodeList nodotrack = ((XmlElement)lstNodos[0]).GetElementsByTagName("TRACKID");
                            rutEmisor = lstNodos[i].InnerText;
                           
                            if (RutFormat != rutEmisor)
                            {
                                return false;
                            }
   
                        }

                        if (lstNodos[i].Name == "TRACKID")
                        {
                            track =  lstNodos[i].InnerText;
                        }
                    }

                    
                    XmlNodeList lstNodosRes = xmlDoc.DocumentElement.ChildNodes;
                    

                    for (int x = 0; x < lstNodosRes.Count; x++)
                    {
                        if (lstNodosRes[x].Name == "REVISIONENVIO")
                        {
                            if(lstNodosRes[x].OuterXml.Contains("Rechazado"))
                            {
                                estado = "ERROR";
                            }
                        }

                    }
                                                         


                    this.GrabaResultadoenvio(xPRefijo, "00", xRutCliente,rutEmisor, track, XML_DTE, estado, xCorreo.ToUpper());
                    salida = true;

                    /*************** AQUI ENVIR CORREO DE AVISO ********************/
                    
                }
                

               
            }
            catch(Exception ex)
            {
                log.Error(ex);
                return false;
            }
            



            return salida;

        }
        private void GrabaResultadoenvio(string xPrefijo,string xLocal, string xRutEmisor,string xRutFormat, string xTrack, string XML, 
                                        string xEstado, string xCorreoOrigen)
        {
            ClienteDTE cliente = new ClienteDTE(xPrefijo, xRutEmisor, xLocal);
            cliente.GrabaResultadoEnvio(xRutFormat, xTrack, XML, xEstado, xCorreoOrigen);
        }
        private void GrabaXMLIntercambio(string xPrefijo, string xRut, string xLocal,string xTipoXML, string xNombreArchivo, 
                                         string xCorreo, string XML, string xResponde, string xRecibe, string xFechaEnvio)
        {
            ClienteDTE cliente = new ClienteDTE(xPrefijo, xRut, xLocal);
            cliente.GrabaDteAcuse(xCorreo, xTipoXML, xNombreArchivo, XML, xResponde, xRecibe, xFechaEnvio);
        }
        private void GrabaDteProveedor(string xPrefijo,string xRut, string  xLocal, 
                                    string xRutEmisor, string xRazonSocial,string xTipoDTE, string xFolioDTE,
                                    string xFechaEmision, string xCorreo, double xtotal,string XML,string xNombreArchivo
                                    )
        {
            ClienteDTE cliente = new ClienteDTE(xPrefijo, xRut, xLocal);
            cliente.GrabaDteProveedor(xRutEmisor, xRazonSocial.ToUpper(), xTipoDTE, xFolioDTE, xFechaEmision, xCorreo.ToUpper(), xtotal, XML, xNombreArchivo);
            
        }
        private bool ExisteDTE(string xPrefijo, string xRut, string xLocal)
        {
            bool salida = false;
            ClienteDTE cliente = new ClienteDTE(xPrefijo, xRut, xLocal);


            return salida;
        }
     




        protected void Download(object sender, EventArgs e)
        {
            //LinkButton lnkAttachment = (sender as LinkButton);
            //GridViewRow row = (lnkAttachment.Parent.Parent.NamingContainer as GridViewRow);
            //List<Attachment> attachments = this.Emails.Where(email => email.MessageNumber == Convert.ToInt32(gvEmails.DataKeys[row.RowIndex].Value)).FirstOrDefault().Attachments;
            //Attachment attachment = attachments.Where(a => a.FileName == lnkAttachment.Text).FirstOrDefault();
            //Response.AddHeader("content-disposition", "attachment;filename=" + attachment.FileName);
            //Response.ContentType = attachment.ContentType;
            //Response.BinaryWrite(attachment.Content);
            //Response.End();
        }

        //private void rbCentralizacion_ToggleStateChanged(object sender, StateChangedEventArgs args)
        //{
        //    if(rbCentralizacion.CheckState == CheckState.Checked)
        //    {
        //        txtInfo.Text += DateTime.Now + " ===> Activando la Centralización de MySQL <=== " + Environment.NewLine;
        //        txtInfo.Refresh();
        //    }
          
        //}

        //private void rbCorreos_ToggleStateChanged(object sender, StateChangedEventArgs args)
        //{
        //    if(rbCorreos.CheckState == CheckState.Checked)
        //    {
        //        txtInfo.Text += DateTime.Now + " ===> Activando la Centralización de Correos <=== " + Environment.NewLine;
        //        txtInfo.Refresh();
        //    }
            
        //}

        private void pnlMain_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
[Serializable]
public class Email
{
    public Email()
    {
        this.Attachments = new List<Attachment>();
    }
    public int MessageNumber { get; set; }
    public string From { get; set; }
    public string Subject { get; set; }
    public string Body { get; set; }
    public DateTime DateSent { get; set; }
    public List<Attachment> Attachments { get; set; }
}
[Serializable]
public class Attachment
{
    public string FileName { get; set; }
    public string ContentType { get; set; }
    public byte[] Content { get; set; }
}