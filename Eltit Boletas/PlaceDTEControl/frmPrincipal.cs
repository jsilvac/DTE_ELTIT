using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using Eltit;
using Eltit.clases;
using Eltit.Clases;
using MySql.Data.MySqlClient;

namespace SamplesDTE
{
    public partial class frmPrincipal : Telerik.WinControls.UI.RadForm
    {
        private Icon[] icons = new Icon[2];
        private int currentIcon = 0;
        Handler handler = null;
        private static readonly log4net.ILog log =
           log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        Usuarios usr = new Usuarios(FuncionesClass.G_SERVIDORMASTER, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);


        public frmPrincipal()
        {
            
            InitializeComponent();

           

        }

        private void frmGeneraDocumentos_Load(object sender, EventArgs e)
        {
            FuncionesClass config = new FuncionesClass();
            config.CargaConfiguracionInicial();
            this.InicializaControlesDeEmpresa();
            lblusuario.Text = "USUARIO: " + FuncionesClass.G_USERS_LOG;

            lblInformacion.Text = "EMPRESAS ELTIT - DEPARTAMENTO DE SISTEMAS | " + DateTime.Now.ToLongDateString() + " | SERVIDOR: " + FuncionesClass.G_SERVIDOR ;
   
            icons[0] = new Icon("factura.ico");
            icons[1] = new Icon("xml.ico");        
            handler = new Handler();

            /**********************************************************************/
            //LocalesClass loc = new LocalesClass(FuncionesClass.G_SERVIDOR);
            this.revisaPermisos();

        }

        private void InicializaControlesDeEmpresa()
        {
            // lblRut.Text = Convert.ToDouble(FuncionesClass.G_EMPRESARUT.Substring(0, 9)) + "-" + FuncionesClass.G_EMPRESARUT.Substring(9, 1);


           // lblRecinto.Text = "SISTEMAS INFORMÁTICOS PLACESOFT SPA";
           
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
 
        private void radMenuItem2_Click(object sender, EventArgs e)
        {
            frmIngresaTimbraje frm = new frmIngresaTimbraje();
            frm.Show();
        }

        private void radMenuItem6_Click(object sender, EventArgs e)
        {
            frmActualizaInfoIntercambio frm = new frmActualizaInfoIntercambio();
            frm.ShowDialog();
        }

        public void radMenuItem7_Click(object sender, EventArgs e)
        {
            frmGeneraInformeRCOF form = new frmGeneraInformeRCOF();
            form.Show();
        }

        private void radMenuItem10_Click(object sender, EventArgs e)
        {
            frmGeneraInformeLibroBoletas frm = new frmGeneraInformeLibroBoletas();
            frm.Show();
        }

        private void radMenu1_Click(object sender, EventArgs e)
        {

        }

        private void radMenuItem14_Click(object sender, EventArgs e)
        {
            frmRegeneraMasivoBoletas frm = new frmRegeneraMasivoBoletas();
            frm.ShowDialog();
        }

        private void radMenuItem15_Click(object sender, EventArgs e)
        {
            frmRegeneraMasivoFacturas frm = new frmRegeneraMasivoFacturas();
            frm.Show();
        }

        private void radMenuItem13_Click(object sender, EventArgs e)
        {
            frmIngresaTimbrajeFacturas frm = new frmIngresaTimbrajeFacturas();
            frm.ShowDialog();
        }

        private void radMenuItem16_Click(object sender, EventArgs e)
        {
            frmInformeBoletasGeneradas frm = new frmInformeBoletasGeneradas();
            frm.Show();
        }
        private void radMenuItem12_Click(object sender, EventArgs e)
        {
        //    frmActualizaInfoIntercambio frm = new frmActualizaInfoIntercambio();
        //    frm.Show();
        }



        private void revisaPermisos()
        {
            MySqlDataReader dr =null ;
            dr = usr.GetPermisosUsuario (FuncionesClass.G_USERS_LOG);

            if (dr.HasRows == true)
            {

                while (dr.Read())
                {
                        
                    if (dr["programa"].ToString() == "frmGeneraInformeLibroBoletas")
                    {
                        radMenuItem10.Enabled = true;
                    }
                    if (dr["programa"].ToString() == "frmGeneraInformeRCOF")
                    {
                        radMenuItem7.Enabled = true;
                    }
                    if (dr["programa"].ToString() == "frmInformeBoletasGeneradas")
                    {
                        radMenuItem16.Enabled = true;
                    }
                    if (dr["programa"].ToString() == "frmIngresaTimbraje")
                    {
                        radMenuItem2.Enabled = true;
                    }
                    if (dr["programa"].ToString() == "frmIngresaTimbrajeFacturas")
                    {
                        radMenuItem13.Enabled = true;
                    }
                    //if (dr["programa"].ToString() == "frmPopEnviaRCOF")
                    //{
                    //    radMenuItem7.Enabled = true;
                    //}
                    //if (dr["programa"].ToString() == "frmPopGeneraLibro")
                    //{
                    //    radMenuItem7.Enabled = true;
                    //}
                    if (dr["programa"].ToString() == "frmRegeneraMasivoBoletas")
                    {
                        radMenuItem14.Enabled = true;
                    }
                    if (dr["programa"].ToString() == "frmRegeneraMasivoFacturas")
                    {
                        radMenuItem15.Enabled = true;
                    }
                    if (dr["programa"].ToString() == "frmActualizaInfoIntercambio")
                    {
                        radMenuItem6.Enabled = true;
                    }

                }



            }
            else
            {

                MessageBox.Show(" Su Usuario no tiene privilegios... ");
                
            }
        }

        private void radMenuItem18_Click(object sender, EventArgs e)
        {
            frmReimprimir r = new frmReimprimir();
            r.Show();
        }
    }
}
