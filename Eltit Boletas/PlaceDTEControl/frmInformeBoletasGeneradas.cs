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
using System.Xml.Serialization;
using System.IO;
using Telerik.WinControls.UI;
using System.Xml;
using Eltit.Clases;
using System.util;
using PlaceSoft.Eltit.Class.clases;

namespace Eltit { 

public partial class frmInformeBoletasGeneradas : Telerik.WinControls.UI.RadForm
    {
        public string mysql_root;
        public string mysql_pass;
        private Icon[] icons = new Icon[2];
        private int currentIcon = 0;
        Handler handler = new Handler();
        List<PlaceSoft.DTE.Engine.Documento.DTE> dtes = new List<PlaceSoft.DTE.Engine.Documento.DTE>();
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        string BASE_DTE = "";
        string BASE_VENTAS = "";
        private string cliente_sistema;
        private string XML_DETALLE ="";
        private int countGenerados = 0;
        List<string> rangoUtilizado39 = new List<string>();
        List<string> rangoAnulado39 = new List<string>();

        List<string> rangoUtilizado41 = new List<string>();
        List<string> rangoAnulado41 = new List<string>();

        List<string> rangoUtilizado61 = new List<string>();
        List<string> rangoAnulado61 = new List<string>();

        public bool muestraDetalles;
        public DataTable dt_locales;

        public frmInformeBoletasGeneradas()
        {
            InitializeComponent();
        }

        private void frmInformeBoletasGeneradas_Load(object sender, EventArgs e)
        {
       
            try
            {
                lblRoot.Text = "adminerp_general";
                lblPassword.Text = "fran061cony252agus203elba214";
                lblServidorVentas.Text = "192.168.4.9";
                GetEmpresasContables();
                gvInforme.Columns[5].ReadOnly = false;

                dtpdesde.Value = DateTime.Now;
                dtpHasta.Value = DateTime.Now;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error:[" + ex.Message.ToString() + "]");
            }
          


        }

    private void ddLlocales_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
    {
        if (ddLlocales.SelectedIndex > 0)
        {
            LeeDatosLocal();
            txtCaja.Focus();

        }
    }
    private void CargaLocales()
    {
        Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);

        MySqlDataReader dr = cli.GetLocalesByEmpresa(ddlEmpresas.Text.Substring(0, 2));

        ddLlocales.Items.Clear();
        ddLlocales.Items.Add("-- SELECCIONE UN LOCAL --");

        if (dr.HasRows == true)
        {
            while (dr.Read())
            {
                ddLlocales.Items.Add(dr["codigo"].ToString() + " " + dr["nombrelocal"].ToString());
            }
        }

        dr.Close();
        cli.CerrarTransaccion();

        ddLlocales.SelectedIndex = 0;




    }
    private void ddlEmpresas_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
    {
        if (ddlEmpresas.SelectedIndex != 0)
        {
                LeeDatosEmpresa();
            CargaLocales();
        }
    }
    private void LeeDatosLocal()
    {
        Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
        MySqlDataReader dr = cli.GetDatosLocalByCodigo(ddLlocales.Text.Substring(0, 2));
        FuncionesClass fu = new FuncionesClass();

        if (dr.HasRows == true)
        {
            if (dr.Read())
            {
                lblDireccion.Text = dr["direccion"].ToString();            
               // lblServidorVentas.Text = dr["servidor_ventas"].ToString();
                lblRutEmpresa.Text = dr["rut"].ToString();
                if (fu.PingToHost(lblServidorVentas.Text) == true)
                {
                        if(lblServidorVentas.Text == "192.168.4.9")
                        {
                            lblRoot.Text = "adminerp_general";
                            lblPassword.Text = "fran061cony252agus203elba214";
                        }
                        else
                        {
                            lblRoot.Text = "conta";
                            lblPassword.Text = "conta";
                        }

                        txxCajaFolios.Enabled = true;

                    pbStatus.Image = global::Eltit.Properties.Resources.OK_48;
                    btnVer.Enabled = true;
                        txxCajaFolios.Focus();
                }
                else
                {
                      btnVer.Enabled = false;
                        pbStatus.Image = global::Eltit.Properties.Resources.icons8_exclamacion;
                        RadMessageBox.Show(this, "No se puede establecer conexión con el Host de Destino [" + lblServidorVentas.Text + "]", "Atencion", MessageBoxButtons.OK);
                }

            }
        }

        dr.Close();
        cli.CerrarTransaccion();

    }
    private void GetEmpresasContables()
    {
        Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
        MySqlDataReader dr = cli.GetClientesDTE();

        ddlEmpresas.Items.Clear();
        ddlEmpresas.Items.Add("-- SELECCIONE EMPRESA --");

        if (dr.HasRows == true)
        {
            while (dr.Read())
            {
                ddlEmpresas.Items.Add(dr["codigo_contable"].ToString() + " " + dr["razon_social"].ToString());
            }
        }

        dr.Close();
        cli.CerrarTransaccion();

        ddlEmpresas.SelectedIndex = 0;
            
    }
        private void LeeDatosEmpresa()
        {
            Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = cli.GetEmpresaByCodigo(ddlEmpresas.Text.Substring(0, 2));
            FuncionesClass fu = new FuncionesClass();

            if (dr.HasRows == true)
            {
                if (dr.Read())
                {
                    lblComuna.Text = dr["comuna"].ToString();
                    lblCiudad.Text = dr["ciudad"].ToString();
                    lblDireccionEmpresa.Text = dr["direccion"].ToString();
                    lblGiro.Text = dr["giro"].ToString();

                    lblRutCertificado.Text    = dr["rut_certificado"].ToString();
                    lblNombreCertificado.Text = dr["nombre_certificado"].ToString();
                    lblFechaResolucion.Text   = dr["fecha_resolucion"].ToString();
                    lblNumeroResolucion.Text  = dr["numero_resolucion"].ToString();
                    lblCodigoSucursal.Text    = dr["codigo_sucursal_sii"].ToString();

                }
            }

            dr.Close();
            cli.CerrarTransaccion();

        }

        private void BuscaBoletas()
        {

        }
        private void btnVer_Click(object sender, EventArgs e)
        {
            if(ddlEmpresas.SelectedIndex > 0 && ddLlocales.SelectedIndex > 0 )
            {
                this.Enabled = false;
                this.Refresh();
                if(txxCajaFolios.Text.Length != 2 )
                {
                    RadMessageBox.Show(this, "Debe Selecionar Una Caja de venta Válida. [" + txxCajaFolios.Text + "]", "Atencion", MessageBoxButtons.OK);
                }
                else
                {
                    GetDocumentosByCajaaLocalDesdeHata();
                }

                this.Enabled = true;
                this.Refresh();

            }
            else
            {
                RadMessageBox.Show(this, "Debe Selecionar Una Empresa y Local Válidos. [" + ddLlocales.Text + "]", "Atencion", MessageBoxButtons.OK);
            }
        }
        private void BuscaCaf(string xCaja, string xTipo)
        {
            Eltit.Clases.Caf myCaf = new Eltit.Clases.Caf(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            MySqlDataReader dr = myCaf.GetCafByCajaLocal(ddLlocales.Text.Substring(0, 2), xCaja, xTipo);
            string filePath = "";
            string xmlString = "";
           

            if (dr.HasRows == true)
            {
                while(dr.Read())
                {
                    filePath = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + Convert.ToDouble(lblRutEmpresa.Text.Substring(0, 9)) + @"\Produccion\Caf\" + ddLlocales.Text.Substring(0, 2) + @"\";
                    filePath =  filePath + string.Format("{0}_{1}_{2}.dat", Convert.ToInt32(xTipo),dr["desde"].ToString(), dr["hasta"].ToString());

                    if(!File.Exists(filePath))
                    {

                        FileStream fst;
                        BinaryWriter bw;
                        string tmp_path = @"C:\temp\" + DateTime.Now.Ticks + ".xml";

                        fst = new FileStream(tmp_path, FileMode.OpenOrCreate, FileAccess.Write);
                        bw = new BinaryWriter(fst);
                        string strxml = dr["xml"].ToString().Replace("±	","");
                        Encoding ByteConverter = Encoding.GetEncoding("ISO-8859-1");
                        byte[] textEnBytes = ByteConverter.GetBytes(strxml);

                        bw.Write(textEnBytes);
                        bw.Flush();
                        bw.Close();
                        bw.Dispose();

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

            dr.Close();
            myCaf.CerrarTransaccion();
            
        }
        private void GetDocumentosByCajaaLocalDesdeHata()
        {
            Documentos dc = new Documentos(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            MySqlDataReader dr = null;
            string base_datos = FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + ddLlocales.Text.Substring(0,2) ;
            object img = null;
            int count = 0;
  
            string existe = "";
            DataTable dt = new DataTable();
            /**************     VERIFICA SI EL CAF EXISTE SI NO LOS TRAE **************/
            // this.BuscaCaf(txxCajaFolios.Text, tipoFiscal);
            string base_venta = FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + ddLlocales.Text.Substring(0, 2);
            dr = dc.GetDocumentosCabezaByLocalNroInternoCaja(ddLlocales.Text.Substring(0, 2), txxCajaFolios.Text,
                                                            FuncionesClass.GetFechaMysql(dtpdesde.Text), FuncionesClass.GetFechaMysql(dtpHasta.Text), base_venta);
            dt.Load(dr);
            gvInforme.Rows.Clear();
            dr.Close();
            dc.CerrarTransaccion();

            ddlEmpresas.Enabled = false;
            ddLlocales.Enabled = false;
            btnGenera.Enabled = true;
            bool paso = false;

            foreach (DataRow row in dt.Rows)
            {
                paso = false;
                existe = this.VerificaBoleta(row["tipo"].ToString(), row["numero"].ToString(), row["fecha"].ToString(), row["caja"].ToString());
                if(existe != "")
                {
                    paso = true;
                }
                if(chbNoGeneradas.CheckState == CheckState.Checked)
                {
                    if(paso == false)
                    {
                        gvInforme.Rows.Add(row["tipo"].ToString(), row["numero"].ToString(), row["fecha"].ToString(), row["caja"].ToString(), String.Format("{0:N0}", row["total"]), existe, paso);
                        count++;
                    }

                }
                else
                {
                    gvInforme.Rows.Add(row["tipo"].ToString(), row["numero"].ToString(), row["fecha"].ToString(), row["caja"].ToString(), String.Format("{0:N0}", row["total"]), existe, paso);
                    count++;
                }
               
            }

            lblInfo.Text = "Total Registros " + count.ToString();

     }

        private string VerificaBoleta(string xtipo, string xNumero, string xFecha, string xCaja)
        {
            string salida = "";

            if(xtipo == "BV")
            {
                xtipo = "39";
            }
            else
            {
                xtipo = "41";
            }

            //DTEClass dte = new DTEClass("192.168.4.200", "sistema", "desarrollo_1990");
            DTEClass dte = new DTEClass("192.168.4.23", "sistema", "desarrollo_1990");
            salida = dte.ExisteDTELocalNrocajaFecha(lblRutEmpresa.Text, ddLlocales.Text.Substring(0, 2), ddlEmpresas.Text.Substring(0, 2), xCaja,
                                           xNumero,FuncionesClass.GetFechaMysql(xFecha), xtipo);


            return salida;
        }

        private void InicializaControlesDeEmpresa()
        {
           
              
        }

  
        private void btnGenera_Click(object sender, EventArgs e)
        {
            if(gvInforme.Rows.Count > 0)
            {
                if(RadMessageBox.Show(this, "Desea Blanquear la Revisión para volver a regenerar los documentos? ", "Atencion",MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    countGenerados = 0;
                    blanqueaRevision();
                }
               
            }
        }


        private void blanqueaRevision()
        {
            int i = 0;
            bool check = false;
            string tipo = "";
            string numero = "";
            string fecha = "";
            string caja = "";
            int count = 0;

            for (i=0; i <= gvInforme.Rows.Count-1;i++ )
            {
                check = true; // Convert.ToBoolean(gvInforme.Rows[i].Cells[6].Value);
                if(check == true)
                {
                    tipo = gvInforme.Rows[i].Cells[0].Value.ToString();
                    numero = gvInforme.Rows[i].Cells[1].Value.ToString();
                    fecha = FuncionesClass.GetFechaMysql(gvInforme.Rows[i].Cells[2].Value.ToString());
                    caja = gvInforme.Rows[i].Cells[3].Value.ToString();
                 
                    BlanqueaRevision(ddLlocales.Text.Substring(0, 2), tipo, numero, caja, fecha , "");
                    count++;
                    countGenerados++;
                }

            }


            if(count == 0)
            {
                RadMessageBox.Show(this, "Debe Seleccionar Al menos un Documento.", "Atencion", MessageBoxButtons.OK);
            }
            else
            {
                RadMessageBox.Show(this, "Se Re Generaron " + count + " Documentos tipo [" + tipo +"]", "Atencion", MessageBoxButtons.OK);
                btnGenera.Enabled = false;
            }



        }

        private void BlanqueaRevision(string xLocal,string xTipo, string xNumero, 
                                    string xcaja, string xfecha, string xGlosa)
        {
            Documentos dc = new Documentos(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            dc.MarcaRevisionBoletaElectronica(xLocal, xNumero, xTipo, xfecha, xcaja, xGlosa);
        }
   
 

        private void MarcarCheck(bool val)
        {
            int i = 0;

            for(i=0; i <= gvInforme.Rows.Count -1; i++)
            {

                gvInforme.Rows[i].Cells[6].Value = val;
            }

        }

    }



}