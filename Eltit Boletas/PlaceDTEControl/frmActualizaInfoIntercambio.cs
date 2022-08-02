
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
using Eltit.Clases;

namespace SamplesDTE
{
    public partial class frmActualizaInfoIntercambio : Telerik.WinControls.UI.RadForm
    {
    
        private string _CURR_COMPANY;
        private static readonly ILog log =
          LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public frmActualizaInfoIntercambio()
        {
            InitializeComponent();
        }

        private void frmActualizaInfoIntercambio_Load(object sender, EventArgs e)
        {
            FuncionesClass config = new FuncionesClass();
            config.CargaConfiguracionInicial();
            this.InicializaControlesDeEmpresa();

            _CURR_COMPANY = FuncionesClass.G_EMPRESAACTIVA;
        
            lblInformacion.Text = "EMPRESA: " + FuncionesClass.G_EMPRESANOMBRE;
            radGroupBox1.GroupBoxElement.Header.Font = new System.Drawing.Font("Arial", 6);

        }


        private void InicializaControlesDeEmpresa()
        {
            //lblRut.Text = FuncionesClass.G_EMPRESARUT;
            //lblNombreEmpresa.Text = FuncionesClass.G_EMPRESANOMBRE;
            //lblDireccion.Text = FuncionesClass.G_EMPRESADIRECCION;
            //lblComuna.Text = FuncionesClass.G_EMPRESACOMUNA;
            //lblCiudad.Text = FuncionesClass.G_EMPRESACIUDAD;

        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            FuncionesClass fu = new FuncionesClass();

            if(fu.PingToHost(FuncionesClass.G_SERVIDOR) == true)
            {
                this.CargaTimbraje();
            }
            else
            {
                RadMessageBox.Show(this, "NO SE PUEDE ESTABLECER CONEXION CON EL SERVIDOR [" + FuncionesClass.G_SERVIDOR + "]" , "Atencion", MessageBoxButtons.OK);
            }
            
        }


        public void CargaTimbraje()
        {
            try
            {
                openFileDialog1.ShowDialog();
                openFileDialog1.InitialDirectory = @"C:\";
                openFileDialog1.Filter = "txt files (*.csv)|*.txt|All files (*.*)|*.*";
                openFileDialog1.FilterIndex = 2;

                if (File.Exists(openFileDialog1.FileName))
                {
                    double cont = 0;
                    List<string> Lineas = new List<string>();
                    String strLine = String.Empty;
                    txtFilePath.Text = openFileDialog1.FileName;
                    StreamReader sr = null;
                    

                    sr = new StreamReader(txtFilePath.Text);
                    while(sr.Peek() >= 0)
                    {
                        strLine = String.Empty;
                        strLine = sr.ReadLine();
                        if(cont > 0)
                        {
                            Lineas.Add(strLine);
                        }
                        
                        cont++;
                    }

                    this.ProcesaLinea(Lineas);
                   // RadMessageBox.Show(this, "Se ingresaron :" + cont + " Registros.", "Atencion", MessageBoxButtons.OK);
                }
            }
            catch(Exception ex)
            {
                log.Error("Error:", ex);
                RadMessageBox.Show(this, "Error:" + ex.Message.ToString(), "Atencion", MessageBoxButtons.OK);
            }
           
        }
       private void ProcesaLinea(List<string> xLineas)
        {

            //DTEClass DTE = new DTEClass("placesoft_", FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            //DTE.IngresaRegistroIntercambio(xLineas,ref lblInfo,"placesoft_mantencion");

            Intercambio intclass = new Intercambio(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            intclass.IngresaRegistroIntercambio(xLineas,ref lblInfo, "eltit_fae");


        }
        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

     
        private void Return()
        {
            txtFilePath.Text = "";
            
           
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
            txtFilePath.Text = "";
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {

        }
    }
}
