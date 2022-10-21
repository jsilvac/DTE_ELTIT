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
using System.IO;
using PlaceSoftDTE.clases;
using Newtonsoft.Json;
using PlaceSoft.Eltit.Class;
using PlaceSoft.Eltit.Handler;
using PlaceSoft.Eltit.Class.clases;
using static PlaceSoft.Enum.Ambiente;
using Newtonsoft.Json.Linq;
using PlaceSoft.DTE.Engine.Enum;
using PlaceSoft.Eltit.Functions;
using Eltit;

namespace SamplesDTE
{
    public partial class frmVerificaStatusBoletaSII : Telerik.WinControls.UI.RadForm
    {
        private Icon[] icons = new Icon[2];
        private int currentIcon = 0;

        private static readonly log4net.ILog log =
           log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        int cuenta = 0;
        private string[] tiposDTE = { "33", "34", "52", "56" };

        Handler handler;
        public frmVerificaStatusBoletaSII()
        {
            InitializeComponent();
        }

        private void frmVerificaStatusBoletaSII_Load(object sender, EventArgs e)
        {
            //FuncionesClass config = new FuncionesClass();
            //config.CargaConfiguracionInicial();
            // this.InicializaControlesDeEmpresa();


            icons[0] = new Icon("sii.ico");
            icons[1] = new Icon("xml.ico");




            this.CargaEmpresas();
            System.Threading.Thread.Sleep(1000);
            timer1.Start();

        }

        private void CargaEmpresas()
        {
            Empresas emp = new Empresas(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);
            MySqlDataReader dr = emp.GetEmpresasBoleta();
            int i = 0;

            if (dr.HasRows == true)
            {
                while (dr.Read())
                {
                    gvEmpresas.Rows.Add(dr["codigo_contable"].ToString(), dr["rut"].ToString(), dr["razon_social"].ToString());
                }
            }

            dr.Close();
            emp.CerrarTransaccion();
        }

        private void InicializaControlesDeEmpresa()
        {
            Empresas emp = new Empresas(Inicial.G_SERVIDOR, Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
            MySqlDataReader dr;
            string cod_empresa = "";
            int i = 0;
            string rut = "";

            // dr = emp.GetDatoEmpresaByRut("eltit_", "0775877707");

            for (i = 0; i < gvEmpresas.Rows.Count; i++)
            {
                cod_empresa = gvEmpresas.Rows[i].Cells[0].Value.ToString();
                rut = gvEmpresas.Rows[i].Cells[1].Value.ToString();

                emp = new Empresas(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);
                dr = emp.GetDatoEmpresaByRut("eltit_", rut);

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                       
                        lblRut.Text = dr["rut"].ToString();
                        lblNombreEmpresa.Text = dr["razon_social"].ToString();
                        lblDireccion.Text = dr["direccion"].ToString();
                        lblComuna.Text = dr["comuna"].ToString();
                        lblCodigo.Text = dr["codigo_contable"].ToString();
                        lblCertificadoRut.Text = dr["rut_certificado"].ToString();
                        lblCertificadoNombre.Text = dr["nombre_certificado"].ToString();
                        lblFechaResolucion.Text = dr["fecha_resolucion"].ToString();
                        lblNroResolucion.Text = dr["numero_resolucion"].ToString();

                        //handler = new Handler("eltit_");
                        //handler.rutEmpresa = Convert.ToDouble(lblRut.Text.Substring(0, 9)) + "-" + lblRut.Text.Substring(9, 1);
                        //handler.rutCertificado = lblCertificadoRut.Text;
                        //handler.nombreCertificado = lblCertificadoNombre.Text;



                        CargaDTEEmpresa(lblRut.Text, lblCodigo.Text);

                    }
                }

                dr.Close();
                emp.CerrarTransaccion();

            }

        }

        private void CargaDTEEmpresa(string xRut, string xCodigo)
        {
            DTEClass dte = new DTEClass(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);
            MySqlDataReader dr;

            dr = dte.getXMLEmpresaSinEstado(xRut, xCodigo, 50);

            if (dr.HasRows == true)
            {
                while (dr.Read())
                {
                    gvInforme.Rows.Add(dr["fae_tipo"].ToString(), dr["fae_folio"].ToString(), dr["fae_fecha"].ToString(), dr["fae_cliente_rut"].ToString(),
                                         dr["fae_monto_total"].ToString(), "");

                }
            }

            dr.Close();
            dte.CerrarTransaccion();

            this.RecorrerGrillaValidacion();


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //notifyIcon1.Icon = icons[currentIcon];
            //currentIcon++;
            //if (currentIcon == 2)
            //    currentIcon = 0;

            //   btnGenerar_Click(sender, e);

            string minuto = DateTime.Now.ToString("HH:mm:ss");

            if (minuto.Substring(4, 1) == txtMinuto1.Text || minuto.Substring(4, 1) == txtMinuto2.Text || minuto.Substring(4, 1) == txtMinuto3.Text || minuto.Substring(4, 1) == txtMinuto4.Text)
            {
                InicializaControlesDeEmpresa();
            }

            lblStatus.Text = "Ultima Comprobación a as " + DateTime.Now.ToString("HH:mm:ss");
            // timer1.Stop();
        }


        private void btnGenerar_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            this.Refresh();
            long trackId = long.Parse(txtTrack.Text);
            lblStatus.Text = "Conectando con SII";
            lblStatus.Refresh();
            try
            {


                int rut = Convert.ToInt32(lblRut.Text.Substring(0, 9));
                string dv = lblRut.Text.Substring(9, 1);
                handler = new Handler("eltit_");



                //if (chkProduccion.Checked == true)
                //{
                //    var responseEstadoDTE = handler.ConsultarEstadoEnvioBoleta(AmbienteEnum.Produccion, trackId,
                //                        rut, dv, lblCertificadoNombre.Text);
                //    textRespuesta.Text = JsonConvert.SerializeObject(responseEstadoDTE, Formatting.Indented);
                //}
                //else
                //{
                //    var responseEstadoDTE = handler.ConsultarEstadoEnvioBoleta(AmbienteEnum.Certificacion, trackId,
                //                        rut, dv, lblCertificadoNombre.Text);
                //    textRespuesta.Text = JsonConvert.SerializeObject(responseEstadoDTE, Formatting.Indented);
                //}



                lblStatus.Text = "Resultados Desplegados";
                lblStatus.Refresh();


                this.Enabled = true;
                this.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error:" + ex);
            }

        }

        private void RecorrerGrillaValidacion()
        {
            int i = 0;
            string numero = "";
            string rut = "";
            string total = "";
            string fecha = "";
            string tipo = "";

            for (i = 0; i <= gvInforme.Rows.Count - 1; i++)
            {
                tipo = gvInforme.Rows[i].Cells[0].Value.ToString();
                numero = gvInforme.Rows[i].Cells[1].Value.ToString();
                fecha = Inicial.GetFechaMysql(gvInforme.Rows[i].Cells[2].Value.ToString());
                rut = gvInforme.Rows[i].Cells[3].Value.ToString();
                total = gvInforme.Rows[i].Cells[4].Value.ToString();

                lblStatus.Text = "Verificando Documento [" + tipo + "-" + numero + "-" + total + "]";
                lblStatus.Refresh();
                VerificarEstadoDocumento(tipo, numero, fecha, rut, total, i);
                lblStatus.Text = lblStatus.Text + " STATUS:" + gvInforme.Rows[i].Cells[5].Value.ToString();
                lblStatus.Refresh();
                System.Threading.Thread.Sleep(500);
                gvInforme.Refresh();
            }

            gvInforme.Rows.Clear();

        }

        private void VerificarEstadoDocumento(string xTipo, string xNumero, string xFecha, string xRut, string xTotal, int xIndice)
        {
            //var responseEstadoDTE = handler.ConsultarEstadoEnvioBoleta(AmbienteEnum.Produccion, trackId,
            //                    rut, dv, lblCertificadoNombre.Text);
            //textRespuesta.Text = JsonConvert.SerializeObject(responseEstadoDTE, Formatting.Indented);

            DateTime fecha_Emision = Convert.ToDateTime(xFecha);
            int rutReceptor = int.Parse(xRut.Substring(0, 9));
            string dvReceptor = xRut.Substring(9, 1); 
            int folio = int.Parse(xNumero);
            int total = int.Parse(xTotal);
            Funciones fu = new Funciones("eltit_");
            Enum.TryParse("BoletaElectronica", out TipoDTE.DTEType tipoDTE);

            handler = new Handler("eltit_");
            handler.rutEmpresa = Convert.ToDouble(lblRut.Text.Substring(0, 9)) + "-" + lblRut.Text.Substring(9, 1);
            handler.rutCertificado = lblCertificadoRut.Text;
            handler.nombreCertificado = lblCertificadoNombre.Text;


            try
            {
                var responseEstadoDTE = handler.ConsultarEstadoDTE(chkProduccion.Checked ? AmbienteEnum.Produccion : AmbienteEnum.Certificacion, rutReceptor, dvReceptor, tipoDTE, folio, fecha_Emision, total);

                textRespuesta.Text = responseEstadoDTE.ResponseXml;
                textRespuesta.Text = JsonConvert.SerializeObject(responseEstadoDTE, Formatting.Indented);

                dynamic data = JObject.Parse(textRespuesta.Text);

                gvInforme.Rows[xIndice].Cells[5].Value = data.Estado.ToString();

                if (data.Estado.ToString() == "DOK")
                {
                    fu.ColoreaCelda(gvInforme.Rows[xIndice].Cells[5], Color.GreenYellow);
                }

                ActualizaEstadoBoleta(xTipo, xNumero, xFecha, data.Estado.ToString(), data.GlosaEstado.ToString());
                System.Threading.Thread.Sleep(500); 
                
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ha ocurrido un error:" + ex);
                log.Error("Error:" + ex);
            }


        }

        private void ActualizaEstadoBoleta(string xTipo, string xNumero, string xFecha, string xStatus, string xGlosaStatus)
        {
            DTEClass dte = new DTEClass(Inicial.G_SERVIDOR_XML_DIRECCION, Inicial.G_SERVIDOR_XML_ROOT, Inicial.G_SERVIDOR_XML_PASS);

            dte.ActualizaEstadoDTEBoleta(xTipo, xNumero, xFecha, xStatus, xGlosaStatus, lblRut.Text, lblCodigo.Text);

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


        }


        private void radToggleSwitch1_ValueChanged(object sender, EventArgs e)
        {


        }

        private void frmEnviodDocumentosSII_FormClosed(object sender, FormClosedEventArgs e)
        {
            timer1.Enabled = false;
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            this.timer1.Stop();
            frmVerificaTrackEnvioBoletaSII frm = new frmVerificaTrackEnvioBoletaSII();
            frm.ShowDialog();

            this.timer1.Start();
        }

    }


}
