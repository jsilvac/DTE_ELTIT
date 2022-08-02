using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Globalization;
using System.IO;
using PlaceSoft.Eltit.Class;
using PlaceSoft.Eltit.Class.clases;
using PlaceSoft.Eltit.Functions;
using PlaceSoft.Eltit.Functions.clases;
using PlaceSoft.Eltit.Handler;
using SchoolManagementAdmin.objetos;

namespace PlaceSoft
{
    public partial class frmEnviodDocumentosSII : Telerik.WinControls.UI.RadForm
    {
        private Icon[] icons = new Icon[2];
        private int currentIcon = 0;
        Handler handler = null;
        private static readonly log4net.ILog log =
           log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        int cuenta = 0;
        private string[] tiposDTE = { "33","34","52", "56" };
        private bool paso = false;

        public frmEnviodDocumentosSII()
        {
            InitializeComponent();            
        }

        private void frmEnviodDocumentosSII_Load(object sender, EventArgs e)
        {

            // this.InicializaControlesDeEmpresa();
            lblInformacion.Text = "EMPRESA: ";
            gvInforme.RootElement.Font = new Font("Arial",6);
            gvInforme.Font = new Font("Arial", 6);

            //icons[0] = new Icon("sii.ico");
            //icons[1] = new Icon("xml.ico");
         
            
            CargaEmpresas();
            System.Threading.Thread.Sleep(1000);        
                                 
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //notifyIcon1.Icon = icons[currentIcon];
            //currentIcon++;
            //if (currentIcon == 2)
            //    currentIcon = 0;
            string hora = DateTime.Now.TimeOfDay.ToString().Substring(0, 2);

            string minutoBoleta = DateTime.Now.TimeOfDay.ToString().Substring(4, 1);
            lblStatus.Text = "Verificando a las " + DateTime.Now.ToShortTimeString();
            lblStatus.Text = "Envio Programado a las " + txtHoraEnvio.Text + " - " + DateTime.Now.ToString("HH:mm:ss");
            lblStatus.Refresh();
            if(Convert.ToInt32(hora) >= 1 && Convert.ToInt32(hora) <= 4 || chbForzar.Checked == true)
            {
                if (Convert.ToInt32(minutoBoleta) == Convert.ToInt32(txtMinuto1.Text) || Convert.ToInt32(minutoBoleta) == Convert.ToInt32(txtMinuto2.Text) ||
                Convert.ToInt32(minutoBoleta) == Convert.ToInt32(txtMinuto3.Text) || Convert.ToInt32(minutoBoleta) == Convert.ToInt32(txtMinuto4.Text))
                {
                    ProcesaEmpresa();
                }
            }
        
        }

        private void frmEnviodDocumentosSII_Activated(object sender, EventArgs e)
        {
            if(paso == false)
            {
                
                paso = true;
                timer1.Enabled = true;
            }
          
        }

        private void CargaEmpresas()
        {
            Empresas emp = new Empresas(Inicial.G_SERVIDOR, Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
            MySqlDataReader dr = emp.GetEmpresasBoleta();
            int i = 0;

           if(dr.HasRows == true)
            {
                while(dr.Read())

                {
                    gvEmpresas.Rows.Add(dr["codigo_contable"].ToString(), dr["rut"].ToString(), dr["razon_social"].ToString());
                    i++;
                }
            }

            dr.Close();
            emp.CerrarTransaccion();

            lblStatus.Text = "Se Encontraron " + i + " empresas";

        }

        private void ProcesaEmpresa()
        {
            int i = 0;
            Empresas emp;
            MySqlDataReader dr;
            string rut = "";
            string cod_empresa = "";
            Funciones fu = new Funciones("eltit_");

           for(i=0;i<gvEmpresas.Rows.Count ; i++)
            {
                cod_empresa = gvEmpresas.Rows[i].Cells[0].Value.ToString();
                rut = gvEmpresas.Rows[i].Cells[1].Value.ToString();

                emp = new Empresas(Inicial.G_SERVIDOR, Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
                dr = emp.GetDatoEmpresaByRut("eltit_", rut);

                gvInforme.Rows.Clear();
                gvInforme.Refresh();
                if(dr.HasRows == true)
                {
                    if(dr.Read())
                    { 
                        handler = new Handler("eltit_");
                        lblRut.Text = dr["rut"].ToString();
                        lblNombreEmpresa.Text = dr["razon_social"].ToString();
                        lblDireccion.Text = dr["direccion"].ToString();
                        lblComuna.Text = dr["comuna"].ToString();                  
                        lblCodigo.Text = dr["codigo_contable"].ToString();
                        lblCertificadoRut.Text = dr["rut_certificado"].ToString();
                        lblCertificadoNombre.Text = dr["nombre_certificado"].ToString();
                        lblFechaResolucion.Text = dr["fecha_resolucion"].ToString();
                        lblNroResolucion.Text = dr["numero_resolucion"].ToString();

                        lblStatus.Text = "Generando Envios de " + lblNombreEmpresa.Text + " [" + lblCodigo.Text + "]";
                        lblStatus.Refresh();
                        handler.rutEmpresa = Convert.ToDouble(lblRut.Text.Substring(0, 9)) + "-" + lblRut.Text.Substring(9, 1);
                        handler.rutCertificado = lblCertificadoRut.Text;
                        handler.nombreCertificado = lblCertificadoNombre.Text;
                        handler.fechaResolucion = Convert.ToDateTime(lblFechaResolucion.Text);
                        handler.numero_resolucion = Convert.ToInt32(lblNroResolucion.Text);

                        fu.ColoreaCelda(gvEmpresas.Rows[gvEmpresas.RowCount - 1].Cells[0], Color.YellowGreen);

                    }
                }

                dr.Close();
                emp.CerrarTransaccion();

                RadPageView1.Refresh();

                System.Threading.Thread.Sleep(1000);
                
                btnGenerar_Click(null, null);

                //btnGeneraXML_Click(null,null);
            }

        }

        private void InicializaControlesDeEmpresa()
        {
            lblRut.Text = "";
            lblNombreEmpresa.Text = "";
            lblDireccion.Text = "";
            lblComuna.Text = "";
                       
        }

     

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            string hora = DateTime.Now.TimeOfDay.ToString().Substring(0, 2);
            string hora2 = DateTime.Now.TimeOfDay.ToString().Substring(0, 4);
            string minutoBoleta = DateTime.Now.TimeOfDay.ToString().Substring(4, 1);
            /**************** ENVIO PROGRAMADO ENTRE 00:00:00 y 01:59:00 ******************/

            //lblStatus.Text = "Minuto Boleta [" + minutoBoleta + "]";
            //lblStatus.Refresh();
            //if ( Convert.ToInt32(minutoBoleta) == Convert.ToInt32(txtMinuto1.Text) || Convert.ToInt32(minutoBoleta) == Convert.ToInt32(txtMinuto2.Text) || Convert.ToInt32(minutoBoleta) == Convert.ToInt32(txtMinuto3.Text) || Convert.ToInt32(minutoBoleta) == Convert.ToInt32(txtMinuto4.Text))
            //{
            //lblStatus.Text = "Envio Programado a las " + txtHoraEnvio.Text + " - " + DateTime.Now.ToString("HH:mm:ss");
           
                lblStatus.Text = "Buscando Documentos Electrónicos de " + lblNombreEmpresa.Text;
                lblStatus.Refresh();
                int quedan =  this.BuscaVentas("39");

                if (gvInforme.Rows.Count > 0)
                {
                    btnGeneraXML_Click(null, null);
                    gvInforme.Rows.Clear();
                }
            //}
             
        }

        private int BuscaVentas(string xTipo)
        {
            DTEClass dte = new DTEClass(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);
            MySqlDataReader dr = dte.getXMLEmpresa(lblRut.Text, lblCodigo.Text,DateTime.Now.ToString("yyyy-MM-dd"), 50);
            DataTable dt = new DataTable();

            dt.Load(dr);
            dr.Close();
            dte.CerrarTransaccion();

            gvInforme.Rows.Clear();

            foreach(DataRow row in dt.Rows)
            {
                gvInforme.Rows.Add(row["fae_tipo"].ToString(), row["fae_folio"].ToString(), row["fae_cajadocumento"].ToString(), row["fae_fecha"].ToString(), row["fae_cliente_rut"].ToString(), "BOLETA VENTA", row["fae_recinto"].ToString());
            }

            lblStatus.Text = "Se encontraron " + dt.Rows.Count + " Documentos";
            lblStatus.Refresh();
            return 0;
        }

        private void frmGeneraDocumentos_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
                timer1.Enabled = true;
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

        private void btnGeneraXML_Click(object sender, EventArgs e)
        {
            Funciones fu = new Funciones("eltit_");

            if(fu.PingToHost("8.8.8.8") == true)
            {
                btnGeneraXML.Text = "Generando Sobre de Envío...";
                btnGeneraXML.Refresh();
                btnGeneraXML.Enabled = false;
                this.GenerarSobreDeEnvio();
                btnGeneraXML.Enabled = true;
                btnGeneraXML.Text = "Enviar al SII";
                btnGeneraXML.Refresh();
            }
            else
            {
                btnGeneraXML.Text = "Sin Conexión a Internet.";
            }
           
        }

        private void GenerarSobreDeEnvio()
        {
            string xTipo = "";
            string xNro = "";
            string xCaja = "";
            string xLocal = "";
            string xfecha = "";
        
            int i = 0;
            string prefix = "ENVIO_BOL_";
            Funciones fu = new Funciones("eltit_");

            List<PlaceSoft.DTE.Engine.Documento.DTE> dtes = new List<PlaceSoft.DTE.Engine.Documento.DTE>();
            List<string> xmlDtes = new List<string>();

            for (i=0;i <= gvInforme.Rows.Count-1;i++ )
            {
                xTipo  =  gvInforme.Rows[i].Cells[0].Value.ToString();
                xNro   =  gvInforme.Rows[i].Cells[1].Value.ToString();
                xCaja  =  gvInforme.Rows[i].Cells[2].Value.ToString();
                xLocal =  gvInforme.Rows[i].Cells[6].Value.ToString();
                xfecha =  gvInforme.Rows[i].Cells[3].Value.ToString();
                xfecha = fu.FechaMysql(xfecha);
     
                DTEClass dteClass = new DTEClass(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);

                MySqlDataReader dr = dteClass.getXMLEmpresaByNumeroLocalCaja(lblRut.Text,lblCodigo.Text,xLocal,xNro,xCaja,xfecha);// venta.getDocumentoByLocalFolioSIITipo(xLocal, xTipo, xNro);

                if(dr.HasRows == true)
                {
                    if(dr.Read())
                    {
                        //string xml = File.ReadAllText(dr["fae_xml"].ToString(), Encoding.GetEncoding("ISO-8859-1"));

                        //Byte[] byteBLOBData = new Byte[0];
                        //byteBLOBData = (Byte[])(dr["fae_xml"]);
                        //String xml = System.Text.Encoding.UTF8.GetString(byteBLOBData);
                        string xml = dr["fae_xml"].ToString();

                        var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);

                        /*Generar envio para el SII  un envío puede contener 1 o varios DTE. No es necesario que sean del mismo tipo,
                         es decir, en un envío pueden ir facturas electrónicas afectas, notas de crédito, guias de despacho, etc.  */
                        dtes.Add(dte);
                        xmlDtes.Add(xml);

                    }
                }
                
                dr.Close();
                dteClass.CerrarTransaccion();
            }

            var filePath = "";

            lblStatus.Text = "Generando Sobre ...";
            lblStatus.Refresh();
            var EnvioSII = handler.GenerarEnvioBoletaToSII(dtes, xmlDtes);
            filePath = handler.FirmarEnvioDTEBoleta(EnvioSII);

            lblStatus.Text = "Sobre Timbrado Electrónicamente por " + lblCertificadoNombre.Text;
            lblStatus.Refresh();

            if (File.Exists(filePath))
            {
                FileInfo fi = new FileInfo(filePath);
                string destino = @"C:\PlaceDTE\eltit\"+ Convert.ToDouble(lblRut.Text.Substring(0, 9)) + @"\Produccion\envios\EnvioBOLETA\" + prefix + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xml";
                fi.CopyTo(destino, true);
                string xml = File.ReadAllText(destino, Encoding.GetEncoding("ISO-8859-1"));
                lblStatus.Text = "Enviando Sobre a SII " ;
                lblStatus.Refresh();
                System.Threading.Thread.Sleep(1000);
                this.EnviarAlSII(destino,dtes);

                lblStatus.Text = "Sobre Enviado Satisfactoriamente a las " + DateTime.Now.ToShortTimeString();
                lblStatus.Refresh();

                System.IO.File.Delete(filePath);

            }

         

        }

        private void EnviarAlSII(string pathSobre, List<PlaceSoft.DTE.Engine.Documento.DTE> dtes)
        {
            try
            {
                /*Procedemos a enviar el 'Envío' al SII, que no es otra cosa que simular un upload vía browser*/

                string pathFile = pathSobre;
                long trackId =  handler.EnviarEnvioDTEToSII(pathFile, "XH1F-EFZ5-ZH93", Convert.ToBoolean(chbProduccion.CheckState), true);
                
                log.Info("Número de Envío " + trackId + " a las " + DateTime.Now.ToLongTimeString());

                if(trackId != 0)
                {
                    string fecha = DateTime.Now.ToString("yyyy-MM-dd");
                    string hora = DateTime.Now.ToString("HH:mm:ss");
                    GrabaEnvío(pathSobre, trackId.ToString(),fecha,hora);

                }

            }
            catch (Exception ex)
            {
                timer1.Stop();
                MessageBox.Show("Error: " + ex.Message.ToString());
                log.Error("Error:", ex);
            }
        }

        private void GrabaEnvío(string pathSobre,string xTrack,string xFecha, string xHora)
        {
            DTEClass dteClass = new DTEClass(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);
            int i = 0;
            var filenNme = Path.GetFileNameWithoutExtension(pathSobre) + ".xml";

            for (i = 0; i <= gvInforme.Rows.Count - 1; i++)
            {
                //venta.ActualizaTrackEnDTE(lblRecinto.Text.Substring(0, 2), gvInforme.Rows[i].Cells[0].Value.ToString(),
                //                        gvInforme.Rows[i].Cells[2].Value.ToString(), xFecha, xHora, filenNme, xTrack.ToString());

                dteClass.ActualizaTrackEnDTE(gvInforme.Rows[i].Cells[6].Value.ToString(), gvInforme.Rows[i].Cells[0].Value.ToString(), gvInforme.Rows[i].Cells[1].Value.ToString(),
                                             xFecha, DateTime.Now.ToString("HH:mm:ss"), filenNme, xTrack, lblRut.Text, lblCodigo.Text);

            }
            string xml = File.ReadAllText(pathSobre, Encoding.GetEncoding("ISO-8859-1"));
            this.GrabaSobreEnvioBOLETA(xTrack, filenNme, xml);
        }

        private void GrabaSobreEnvioBOLETA(string xTrack, string xNombreSobre,string XMLSobre)
        {
            DTEClass dte = new DTEClass(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);
            dte.GrabaSobreEnvioBOLETA(lblCodigo.Text, lblRut.Text, xNombreSobre, XMLSobre);


        }
        private void GeneraXML(MySqlDataReader xDrDetalles, string tipo,string xFolioCaf, string xRutCliente,string xFechaEmision,
            string xTipoInterno, string xNroInterno,string xCaja)
        {

            PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronica;
            if (tipo == "33")
            {
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronica;
            }
            if (tipo == "34")
            {
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronicaExenta;
            }
            if (tipo == "46")
            {
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaCompraElectronica;
            }
            if (tipo == "56")
            {
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaDebitoElectronica;
            }
            if (tipo == "61")
            {
                typeDTE = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica;
            }

            handler.Folio = Convert.ToInt32(xFolioCaf);
            handler.casoPruebas = " " ;
            handler.rutEmpresa = lblRut.Text;
            handler.idDte = "DTE" + tipo + "F" + xFolioCaf;
            handler.tipo = typeDTE;
            /********************* BUSCAR DATOS DE CLIENTE *******************/
            //ClientesClass cliente = new ClientesClass(FuncionesClass.G_SERVIDOR);
            //MySqlDataReader drCliente = cliente.getClienteByRut(xRutCliente, "0");
            //var dte = handler.GenerateDTE(drCliente,xFechaEmision);
            ///********************** BUSCA LOS DETALLES DE LA VENTA************/            
            //handler.GenerateDetails(dte, xDrDetalles);
            //drCliente.Close();
            //cliente.CerrarTransaccion();
            //if(tipo == "56" || tipo == "61")
            //{
            //    handler.Referencias(dte, xDrDetalles);
            //}
            var path = ""; //handler.TimbrarYFirmarXMLDTE(dte);
            //handler.ValidateDTE(path, ChileSystems.DTE.Security.Firma.Firma.TipoXML.DTE);
           /****************** SI EXISTE EL PATH Y ES VÁLIDO ENTONCES LO GRABAMOS *****************/
            if (File.Exists(path))
            {
               // FileInfo fi = new FileInfo(path);
               // string destino = @"C:\PlaceDTE\PlaceSoft\xml\dte" + tipo + "F" + xFolioCaf + ".xml";
               // fi.CopyTo(destino, true);
               // string xml = File.ReadAllText(destino, Encoding.GetEncoding("ISO-8859-1"));


               // this.GrabaXML(tipo, xFolioCaf, xFechaEmision,xTipoInterno,xNroInterno,xCaja, xml);
               //// MessageBox.Show("Documento generado exitosamente");
               // frmGeneraPDF doPdf = new frmGeneraPDF();
               // doPdf.txtLocalActivo.Text = lblRecinto.Text.Substring(0, 2);
               // doPdf.CargaTimbraje(path);
               // //System.Threading.Thread.Sleep(1000);
               // //doPdf.btnGenerar_Click(null, null);
               // //System.Threading.Thread.Sleep(1000);
               // doPdf.Show();

               // this.ActualizaCaf(lblRecinto.Text.Substring(0, 2), xTipoInterno, xCaja,xNroInterno,xFolioCaf);
               // lblInformacion.Text = "Último Generado " + tipo + " CAF "+xFolioCaf + " el " + DateTime.Now.ToLongDateString() + " a las " + DateTime.Now.ToLongTimeString();
            }
                       
        }

        private void ActualizaCaf(string xLocal, string xTipo, string xCaja, string xNumero,string xFolioSII)
        {
            //VentasClass ve = new VentasClass(FuncionesClass.G_SERVIDOR);
            //ve.setBaseDTE(FuncionesClass.BASE_DTE);
            //ve.ActualizaFolioSII(xLocal, xTipo, xNumero, xCaja, xFolioSII);
        }

        private void GrabaXML(string xSiiTipo, string xSiiFolio, string xFechaEmision, 
                              string xInternoTipo,string xInternNro,string xCaja, string xml)
        {
            //VentasClass venta = new VentasClass(FuncionesClass.G_SERVIDOR);
            //venta.setBaseDTE(FuncionesClass.BASE_DTE);
            //xml = xml.Replace("'", "~");
            //xFechaEmision = FuncionesClass.GetFechaMysql(xFechaEmision);
            //venta.GrabaXML(lblRecinto.Text.Substring(0,2),xSiiTipo,xSiiFolio,xFechaEmision,xInternoTipo,xInternNro,xCaja,xml);

        }

        private string GetUltimoFolio(string xtipo)
        {
            //CafFoliosClass caf = new CafFoliosClass(FuncionesClass.G_SERVIDOR,FuncionesClass.BASE_DTE);
            //string ultimo = caf.GetUltimoFolioByTipo(xtipo, FuncionesClass.G_EMPRESAACTIVA);

            //return ultimo;

            return "";
        }

        private void radToggleSwitch1_ValueChanged(object sender, EventArgs e)
        {
           
            
        }

        private void frmEnviodDocumentosSII_FormClosed(object sender, FormClosedEventArgs e)
        {
            timer1.Enabled = false;
        }

        private void btnEnviar_Click(object sender, EventArgs e)
        {

        }

        private void radGroupBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
