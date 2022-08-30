using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using SamplesDTE.Clases;
using System.IO;
using log4net;
using MySql.Data;
using MySql.Data.MySqlClient;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Globalization;
using System.Diagnostics;
using PlaceSoftDTE.clases;
using PlaceSoft.Eltit.Class;
using PlaceDTE;
using System.Net.Mail;
using PlaceSoft.Eltit.Class.clases;
using System.Net;
using System.Net.Mime;

namespace SamplesDTE
{
    public partial class frmEnviaXML2 : Telerik.WinControls.UI.RadForm
    {
        //Autorizacion aut;
     
        private static readonly ILog log =
          LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public string RUTA_PDF = "";
        public string RUTA_XML = "";
        public string EMISOR_RUT = "";
        public string MYSQL_ROOT = "";
        public string MYSQL_SERVER = "";
        public string MYSQL_PASS = "";
        public string EMISOR_NOMBRE = "";
        public string EMISOR_MAIL = "";
        public string EMISOR_SMTP_MAIL = "";
        public string EMISOR_PASS_MAIL = "";
        public string RECEPTOR_RUT = "";
        public string RECEPTOR_NOMBRE = "";
        public string RECEPTOR_MAILINTERCAMBIO = "";
        public string MAIL_RECEPTORCLIENTE = "";
        public string TIPO_DTE = "";
        public string MONTO = "";
        public string TIPODOCINTERNO = "";
        List<string> listCorreo = new List<string>();

        MySqlDataReader dr = null;
        Correo m;// new Correo(MYSQL_SERVER, MYSQL_ROOT, MYSQL_PASS);

        public frmEnviaXML2()
        {
            InitializeComponent();
        }

        private void frmRegeneraBoleta_Load(object sender, EventArgs e)
        {

            /********* AQUI GENERAR RUTINA PARA CARGAR DATOS DEL CORREO QUE ENVIA ****////////

            this.getDatosEmisor();
            this.getdatosReceptor();



            /**************** AQUI GENERAR RUTINA PARA BUSCAR CLIENTE **********************/
 
            /*******************************************************************************/

            // Capturamos correo del maestro clientes


        }


        public void getDatosEmisor()
        {
            m = new Correo(MYSQL_SERVER, MYSQL_ROOT, MYSQL_PASS);

            dr = m.GetCorreByERutEmpresa(EMISOR_RUT);

            if (dr.HasRows == true)
            {
                while(dr.Read())
                {
                    EMISOR_MAIL = dr["mailsalida"].ToString();
                    EMISOR_PASS_MAIL = dr["clavemail"].ToString();
                    EMISOR_SMTP_MAIL = dr["servermail"].ToString();
                }
            }
            dr.Close();
            m.CerrarTransaccion();
        }

        public void getdatosReceptor()
        {

            m = new Correo(MYSQL_SERVER, MYSQL_ROOT, MYSQL_PASS);

            dr = null;
            dr = m.GetCorreInetcambioCli(RECEPTOR_RUT);
            if (dr.HasRows == true)
            {
                while (dr.Read())
                {
                    RECEPTOR_MAILINTERCAMBIO = dr["mailintercambio"].ToString();
                    RECEPTOR_NOMBRE = dr["razonsocial"].ToString();

                }
            }




            MySqlDataReader datosMaestroCliente = m.GetMaestroClientes(RECEPTOR_RUT);
            dr = null;
           
            string correoCliente = "";
            if (datosMaestroCliente.HasRows == true)
            {
                if (datosMaestroCliente.Read())
                {
                    correoCliente = datosMaestroCliente["email"].ToString();
                }
            }
            txtRemitente.Text = EMISOR_MAIL;
            txtCliente.Text = RECEPTOR_NOMBRE;
            txtEmail.Text = RECEPTOR_MAILINTERCAMBIO;

            if (correoCliente != "")
            {
                txtCC.Text = correoCliente;
            }
            else
            {
                txtCC.Text = "No registra correo.";
            }

            lblNombreDocumento.Text = TIPO_DTE;
            lblTotalDoc.Text = MONTO;
            txtAsunto.Text = "Documento tributario electrónico";
            lblNombreDocumento.Text = TIPODOCINTERNO;


            //dr.Close();
            m.CerrarTransaccion();


        }


        public void EnviarEmail()
        {
            try
            {

                EMISOR_MAIL ="eltit_dte@placesoft.cl";
                EMISOR_PASS_MAIL = "ELTIT_1990";
                EMISOR_SMTP_MAIL = "mail.placesoft.cl";

                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(EMISOR_SMTP_MAIL);
                mail.From = new MailAddress(EMISOR_MAIL);
                mail.To.Add(txtEmail.Text);
                mail.Subject = txtAsunto.Text ;
                mail.Body = txtGlosa.Text;

               // if (txtCC.Text != "")
//mail.CC.Add("jsilva@eltit.cl");

                foreach (string s in listCorreo)
                {
                    System.Net.Mail.Attachment attachment;
                    attachment = new System.Net.Mail.Attachment(s.ToString());
                    mail.Attachments.Add(attachment);
                }

                SmtpServer.Port = 25;
                SmtpServer.Credentials = new System.Net.NetworkCredential(EMISOR_MAIL, EMISOR_PASS_MAIL);
               // SmtpServer.EnableSsl = false;

                SmtpServer.Send(mail);
                RadMessageBox.Show(this, "Correo Enviado Satisfactoriamente.", "OK", MessageBoxButtons.OK, RadMessageIcon.Info);
                lblStatus.Text = "Correo Enviado OK...";
                lblStatus.Refresh();

                btnEnviar.Enabled = false;
            }
            catch (Exception ex)
            {
                RadMessageBox.Show(this, "Error al enviar Email: " + ex.Message.ToString(), "OK", MessageBoxButtons.OK, RadMessageIcon.Info);
            }
        }


        public  void EnviarEmail2()
        {
            try
            {

                EMISOR_MAIL = "eltit_dte@placesoft.cl";
                EMISOR_PASS_MAIL = "eltit_1990";
                EMISOR_SMTP_MAIL = "mail.placesoft.cl";


                SmtpClient MyMail = new SmtpClient();
                MailMessage MyMsg = new MailMessage();
                MyMail.Host = EMISOR_SMTP_MAIL.ToLower();
                MyMail.Port = 25;
                MyMsg.Priority = MailPriority.Normal;
                MyMsg.To.Add(new MailAddress(txtEmail.Text));
                if (txtCC.Text != "")
                    MyMsg.To.Add(new MailAddress(txtCC.Text));

                MyMsg.Subject = txtAsunto.Text;
                MyMsg.SubjectEncoding = Encoding.UTF8;
                MyMsg.IsBodyHtml = true;
                MyMsg.From = new MailAddress(EMISOR_MAIL, EMISOR_NOMBRE);
                MyMsg.BodyEncoding = Encoding.UTF8;
                MyMsg.Body = txtGlosa.Text;
                MyMail.UseDefaultCredentials = false;
                NetworkCredential MyCredentials = new NetworkCredential(EMISOR_MAIL.ToLower(), EMISOR_PASS_MAIL.ToLower());


                if (RUTA_XML != "")
                {
                    System.Net.Mail.Attachment data1 = new System.Net.Mail.Attachment(RUTA_XML, MediaTypeNames.Application.Octet);
                    ContentDisposition disposition2 = data1.ContentDisposition;
                    disposition2.CreationDate = System.IO.File.GetCreationTime(RUTA_XML);
                    disposition2.ModificationDate = System.IO.File.GetLastWriteTime(RUTA_XML);
                    disposition2.ReadDate = System.IO.File.GetLastAccessTime(RUTA_XML);
                    MyMsg.Attachments.Add(data1);
                }


                if (RUTA_PDF != "")
                {
                    System.Net.Mail.Attachment data2 = new System.Net.Mail.Attachment(RUTA_PDF, MediaTypeNames.Application.Octet);
                    ContentDisposition disposition2 = data2.ContentDisposition;
                    disposition2.CreationDate = System.IO.File.GetCreationTime(RUTA_PDF);
                    disposition2.ModificationDate = System.IO.File.GetLastWriteTime(RUTA_PDF);
                    disposition2.ReadDate = System.IO.File.GetLastAccessTime(RUTA_PDF);
                    MyMsg.Attachments.Add(data2);
                }

                MyMail.Credentials = MyCredentials;
                MyMail.Timeout = 5000000;
                MyMail.Send(MyMsg);
            }
            catch (Exception ex)
            {
                log.Error(ex);
                //MessageBox.Show("Error:" + ex.Message.ToString());
                //string linea;
                //linea = ex.StackTrace.Substring(ex.StackTrace.Length - 7, 7);
                //log.WriteLog(Application.ProductName, System.Reflection.MethodInfo.GetCurrentMethod().ToString() + "(" + linea + ")", ex.Message.ToString);
            }
        }

        private void GenerarEnvioXML()
        {

        }



        private void btnBuscar_Click(object sender, EventArgs e)
        {
          
        }

        
        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
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
        
        }

        private void btnEnviar_Click(object sender, EventArgs e)
        {
            if(RUTA_PDF != "")
            {
                listCorreo.Add(RUTA_PDF);
            }
            if (chbXML.Checked == true)
            {
                listCorreo.Add(RUTA_XML);
            }

            this.EnviarEmail2();
        }
    }
}
