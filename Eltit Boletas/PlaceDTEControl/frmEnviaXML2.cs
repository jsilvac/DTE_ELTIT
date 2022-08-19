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



            /********************************************************************************/

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
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(EMISOR_SMTP_MAIL);
                mail.From = new MailAddress(EMISOR_MAIL);
                mail.To.Add(txtEmail.Text);
                mail.Subject = txtAsunto.Text ;
                mail.Body = txtGlosa.Text;

                if (txtCC.Text != "")
                    mail.CC.Add("jsilva@eltit.cl");

                foreach (string s in listCorreo)
                {
                    System.Net.Mail.Attachment attachment;
                    attachment = new System.Net.Mail.Attachment(s.ToString());
                    mail.Attachments.Add(attachment);
                }

                SmtpServer.Port = 465;
                SmtpServer.Credentials = new System.Net.NetworkCredential(EMISOR_MAIL, "estaes");
                SmtpServer.EnableSsl = true;

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


        //    Public Sub EnviarEmail(xCorreo As Correos, xAsunto As String, xMail As String, xCC As String, ByVal list As List(Of String))
        //    Try
        //        Dim mail As New MailMessage()
        //        Dim SmtpServer As New SmtpClient(xCorreo.Smtp1)
        //        mail.From = New MailAddress(xCorreo.Direccion1)
        //        mail.[To].Add(xMail)
        //        mail.Subject = xAsunto & "  :: " & G_EMPRESANOMBRE
        //        mail.Body = txtGlosa.Text

        //        If xCC<> "" Then
        //            mail.CC.Add(xCC)
        //        End If

        //        For Each s As String In list
        //            Dim attachment As System.Net.Mail.Attachment
        //            attachment = New System.Net.Mail.Attachment(s.ToString())
        //            mail.Attachments.Add(attachment)
        //        Next

        //        SmtpServer.Port = 25
        //        SmtpServer.Credentials = New System.Net.NetworkCredential(xCorreo.Direccion1.ToLower, xCorreo.Clave1)
        //        SmtpServer.EnableSsl = False

        //        SmtpServer.Send(mail)
        //        RadMessageBox.Show(Me, "Correo Enviado Satisfactoriamente.", "OK", MessageBoxButtons.OK, RadMessageIcon.Info)
        //        lblStatus.Text = "Correo Enviado OK..."
        //        lblStatus.Refresh()

        //        btnEnviar.Enabled = False
        //    Catch ex As Exception
        //        RadMessageBox.Show(Me, "Error al enviar Email: " & ex.Message.ToString, "OK", MessageBoxButtons.OK, RadMessageIcon.Info)
        //    End Try

        //End Sub

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

            this.EnviarEmail();
        }
    }
}
