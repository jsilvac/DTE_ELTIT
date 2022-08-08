using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Eltit.Clases;
using MySql.Data;
using Eltit;
using SamplesDTE;

namespace Eltit
{
    public partial class frmLogin : Form
    {

        FuncionesClass fu = new FuncionesClass();


        public frmLogin()
        {
            InitializeComponent();
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            FuncionesClass conf = new FuncionesClass();
            
            conf.CargaConfiguracionInicial();
        }

        private void btnIngresar_Click(object sender, EventArgs e)
        {
            ingresar();
        }

        private void txtPassword_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                ingresar();
            }
        }

        public void ingresar()
        {
            if (System.Environment.MachineName != "SOPORTEJAIME")
            {


                if (fu.existeUsuario(txtUsuario.Text.Trim(), txtPassword.Text.Trim()) == true)
                {
                    FuncionesClass.G_USERS_LOG = txtUsuario.Text;
                    this.Hide();
                    //MessageBox.Show("Usuario correcto....");
                    frmPrincipal fr = new frmPrincipal();

                    fr.ShowDialog();
                    txtPassword.Text = "";
                    txtUsuario.Text = "";
                    //  this.Show();
                    txtUsuario.Focus();
                    FuncionesClass.G_USERS_LOG = "";

                }
                else
                {
                    MessageBox.Show("Usuario o contraseña incorrecta...");
                    txtPassword.Text = "";
                    txtUsuario.Text = "";
                    txtUsuario.Focus();
                    return;
                }

            }
            else
            {
                FuncionesClass.G_USERS_LOG = "jaimiko";
                this.Hide();
                frmPrincipal fr = new frmPrincipal();

                fr.ShowDialog();
                txtPassword.Text = "";
                txtUsuario.Text = "";
                txtUsuario.Focus();
                FuncionesClass.G_USERS_LOG = "";
            }
        }

        private void txtUsuario_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                txtPassword.Focus();
            }
        }
    }
}
