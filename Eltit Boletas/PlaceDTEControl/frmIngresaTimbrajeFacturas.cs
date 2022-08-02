using PlaceSoft.DTE.Engine.Documento;
using PlaceSoft.DTE.Engine.XML;
using System;
using System.Drawing;
using System.Windows.Forms;
using Telerik.WinControls;
using System.IO;
using log4net;
using MySql.Data;
using MySql.Data.MySqlClient;
using Eltit.Clases;


namespace SamplesDTE
{
    public partial class frmIngresaTimbrajeFacturas : Telerik.WinControls.UI.RadForm
    {
        Autorizacion aut;
        private string _CURR_COMPANY;
        private static readonly ILog log =
          LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private string BASE_RUT;

        private string TipoBoleta = "";


        public frmIngresaTimbrajeFacturas()
        {
            InitializeComponent();
        }

        private void frmIngresaTimbraje_Load(object sender, EventArgs e)
        {
            this.InicializaControlesDeEmpresa();
    
            ddLlocales.SelectedIndex = 0;

            TipoBoleta = "39";
            lblInformacion.Text = "EMPRESAS ELTIT | " + DateTime.Now.ToLongDateString() + " SERVIDOR: " + FuncionesClass.G_SERVIDOR;
            groupCaf.GroupBoxElement.Header.Font = new Font("Arial", 6);
            gvInforme.TableElement.Font = new Font("Arial", 8);
            GetEmpresasContables();
        }

        private void GetEmpresasContables()
        {
            Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = cli.GetClientesDTE();

            ddlEmpresas.Items.Clear();
            ddlEmpresas.Items.Add("-- SELECCIONE EMPRESA --");
            
            if (dr.HasRows == true)
            {
                while(dr.Read())
                {
                    ddlEmpresas.Items.Add(dr["codigo_contable"].ToString() + " " + dr["razon_social"].ToString());
                }
            }

            dr.Close();
            cli.CerrarTransaccion();

            ddlEmpresas.SelectedIndex = 0;


        }
        private void ddlEmpresas_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            if(ddlEmpresas.SelectedIndex != 0)
            {
                CargaLocales();
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


        private void InicializaControlesDeEmpresa()
        {
            //BASE_RUT = Convert.ToDouble(FuncionesClass.G_EMPRESARUT.Substring(0, 9)).ToString();
 
            //lblDireccion.Text = FuncionesClass.G_EMPRESADIRECCION;
            //lblComuna.Text = FuncionesClass.G_EMPRESACOMUNA;
      
            FuncionesClass._BASE_FOLDER_PROD = @"C:\PlaceDTE\" + FuncionesClass.G_CLIENTE_PREFIJO.Replace("_", "") + @"\" + BASE_RUT + @"\Produccion\";
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            this.CargaTimbraje();
        }

        private void CargaTimbraje()
        {
            try
            {
                openFileDialog1.ShowDialog();
                if (File.Exists(openFileDialog1.FileName))
                {
                    string nombreArchivo = "";
                    int id = 0;
                    lblFilePath.Text = openFileDialog1.FileName;
                    aut = XmlHandler.DeserializeRaw<Autorizacion>(openFileDialog1.FileName);
                    //aut.CAF.IdCAF = 1;
                    textFecha.Text = aut.CAF.Datos.FechaAutorizacion.ToShortDateString();
                    //textRango.Text = aut.CAF.Datos.RangoAutorizado.Desde.ToString() + " - " + aut.CAF.Datos.RangoAutorizado.Hasta.ToString();
                    nombreArchivo = Path.GetFileName(openFileDialog1.FileName);
                    lbNombreArchivo.Text = nombreArchivo;
                    id = gvInforme.Rows.Count+1;
                                       
                    aut.CAF.IdCAF = id;
                    lblCafDesde.Text = aut.CAF.Datos.RangoAutorizado.Desde.ToString();
                    lblCafHasta.Text = aut.CAF.Datos.RangoAutorizado.Hasta.ToString();

                    txtDesde.Text = aut.CAF.Datos.RangoAutorizado.Desde.ToString();
                    txtHasta.Text = aut.CAF.Datos.RangoAutorizado.Hasta.ToString();

                    lblRutCaf.Text = aut.CAF.Datos.RutEmisor.ToString();

                    lblRutaDestino.Text = @"c:\fae\"+ ddLlocales.Text.Substring(0,2)  + @"\folios\";



                    string tipo = string.Empty;
                    switch (aut.CAF.Datos.TipoDTE)
                    {
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaCompraElectronica:
                            tipo = "FACTURA DE COMPRA ELECTRÓNICA";
                            lbTipo.Text = "46";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronica:
                            tipo = "FACTURA ELECTRÓNICA";
                            lbTipo.Text = "33";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronicaExenta:
                            tipo = "FACTURA ELECTRÓNICA EXENTA";
                            lbTipo.Text = "34";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica:
                            tipo = "GUIA DE DESPACHO ELECTRÓNICA";
                            lbTipo.Text = "52";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica:
                            tipo = "NOTA DE CRÉDITO ELECTRÓNICA";
                            lbTipo.Text = "61";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaDebitoElectronica:
                            tipo = "NOTA DE DÉBITO ELECTRÓNICA";
                            lbTipo.Text = "56";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica:
                            tipo = "BOLETA ELECTRÓNICA";
                            lbTipo.Text = "39";
                            break;
                        case PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta:
                            tipo = "BOLETA ELECTRÓNICA EXENTA";
                            lbTipo.Text = "41";
                            break;
                    }
                    textTipoCAF.Text = tipo;
                  

                    if(lbTipo.Text == TipoBoleta)
                    {
                        RadMessageBox.Show(this, "El CAF seleccionado no corresponde al tipo de Documento Indicado:. [" + TipoBoleta + " <>  "+ lbTipo.Text +" ]", "Atencion", MessageBoxButtons.OK);
                        btnGenerar.Enabled = false;
                    }
                    else
                    {
                        btnGenerar.Enabled = true;
                    }
                                                                                                              

                }
            }
            catch(Exception ex)
            {
                log.Error("Error:", ex);
                RadMessageBox.Show(this, "Error:" + ex.Message.ToString(), "Atencion", MessageBoxButtons.OK);
            }
           
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            Limpiar();
        }

        private void Limpiar()
        {
            lblRutaDestino.Text = "";
            lblRutCaf.Text = "";
            lblcodigoCajaCaf.Text = "";
            lblDescripcionCajaCaf.Text = "";
            pbStatus.Image = null;
            lblFilePath.Text = "";
            lbTipo.Text = "";
            textTipoCAF.Text = "";
            textFecha.Text = "";
            lbNombreArchivo.Text = "";
            lblCafDesde.Text = "0";
            lblCafHasta.Text = "0";
            txtDesde.Text = "0";
            txtHasta.Text = "0";
            btnGenerar.Enabled = true;
            lblDireccion.Text = "";
            lblComuna.Text = "";
            lblServidorVentas.Text = "";
            txtCaja.Text = "";
            gvInforme.Rows.Clear();
            ddlEmpresas.SelectedIndex = 0;
            ddLlocales.SelectedIndex = 0;
            ddlEmpresas.Enabled = true;
            ddLlocales.Enabled = true;
            ddlEmpresas.Focus();
            lblRutEmpresa.Text = "";
            
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {

            bool val1 = false;
            bool val2 = false;

            if ( Convert.ToDouble(txtDesde.Text)  >= Convert.ToDouble(lblCafDesde.Text) &&
                Convert.ToDouble(txtDesde.Text) <= Convert.ToDouble(lblCafHasta.Text) )
            {
                val1 = true;
            }
            if (Convert.ToDouble(txtHasta.Text) >= Convert.ToDouble(lblCafDesde.Text) &&
               Convert.ToDouble(txtHasta.Text) <= Convert.ToDouble(lblCafHasta.Text))
            {
                val2 = true;
            }
  
            if(val1 == false || val2 == false)
            {
                RadMessageBox.Show(this, "Los Rangos ingresados no Coicinden con los contenidos en el CAF indicado:" , "Atencion", MessageBoxButtons.OK);
                return;
            }

            string RutCaf = lblRutCaf.Text.Replace("-", "").PadLeft(10, Convert.ToChar("0"));
            if (lblRutEmpresa.Text != RutCaf)
            {
                RadMessageBox.Show(this, "El Rut de la Empresa y el del Caf seleccionado no coinciden [" + lblRutEmpresa.Text + " <> " + RutCaf + "]" , "Atencion", MessageBoxButtons.OK);
                return;
            }

            try
            {

                string filePath =  @"C:\temp\";              

                bool exists = System.IO.Directory.Exists(filePath);

                if (!exists)
                    System.IO.Directory.CreateDirectory(filePath);

                string xmlString = null;
                xmlString = File.ReadAllText(openFileDialog1.FileName);

                lblRutaDestino.Text = lblRutaDestino.Text.Replace(@"\",@"\\");


                FuncionesClass fu = new FuncionesClass();
                string fechaMysql = fu.FechaMysql(textFecha.Text);

                Caf myCaf = new Caf(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
                myCaf.GrabaCafLocal(ddLlocales.Text.Substring(0, 2), lbTipo.Text, lblCafDesde.Text,
                                     lblCafHasta.Text, fechaMysql , lblRutaDestino.Text, lbNombreArchivo.Text, xmlString, xmlString);

               

                RadMessageBox.Show(this, "CAF almacenado satisfactoriamente", "Success", MessageBoxButtons.OK);
                this.CargaCajasByLocal();
                btnGenerar.Enabled = true;
            }
            catch(Exception ex)
            {
                log.Error("Error:", ex);
                RadMessageBox.Show(this, "Error:" + ex.Message.ToString(), "Atencion", MessageBoxButtons.OK);
            }
            
        }

        private void GetListadoFolios()
        {
            gvInforme.Rows.Clear();
            Caf caf = new Caf(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = null;
            string base_dte = "eltit_dte_" + Convert.ToDouble(lblRutEmpresa.Text.Substring(0, 9));
            gvInforme.Rows.Clear();
            dr = caf.GetLitadoFolios(ddLlocales.Text.Substring(0,2), base_dte);
            if(dr.HasRows == true)
            {
                while(dr.Read())
                {
                    gvInforme.Rows.Add(dr["caf_tipo"].ToString(), dr["caf_desde"].ToString(), dr["caf_hasta"].ToString(), dr["caf_nombre"].ToString());
                }
            }

            dr.Close();
            caf.CerrarTransaccion();
                
        }

        private void ddLlocales_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            if(ddLlocales.SelectedIndex > 0)
            {
                LeeDatosLocal();
                 txtCaja.Focus();

            }
        }
        private void LeeDatosLocal()
        {
            Clientes cli = new Clientes(FuncionesClass.G_SERVIDOR, FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
            MySqlDataReader dr = cli.GetDatosLocalByCodigo(ddLlocales.Text.Substring(0, 2));
            FuncionesClass fu = new FuncionesClass();
       
            if(dr.HasRows == true)
            {
                if(dr.Read())
                {
                    lblDireccion.Text = dr["direccion"].ToString();
                    lblComuna.Text = dr["comuna"].ToString();
                    lblServidorVentas.Text = dr["servidor_ventas"].ToString();
                    lblRutEmpresa.Text = dr["rut"].ToString();
                    if(fu.PingToHost(lblServidorVentas.Text) == true)
                    {
                        if (lblServidorVentas.Text == "192.168.4.9")
                        {
                            lblRoot.Text = "adminerp_general";
                            lblPassword.Text = "fran061cony252agus203elba214";
                        }
                        else
                        {
                            lblRoot.Text = "conta";
                            lblPassword.Text = "conta";
                        }
                        pbStatus.Image = Eltit.Properties.Resources.OK_48;
                        btnVer.Enabled = true;
                    }
                    else
                    {
                        btnVer.Enabled = false;
                        pbStatus.Image = Eltit.Properties.Resources.icons8_exclamacion;
                        RadMessageBox.Show(this, "No se puede establecer conexión con el Host de Destino [" + lblServidorVentas.Text + "]", "Atencion", MessageBoxButtons.OK);
                    }
                    
                }
            }

            dr.Close();
            cli.CerrarTransaccion();

        }
        private void txtCaja_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtCaja_KeyDown(object sender, KeyEventArgs e)
        {
          
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            if(ddlEmpresas.SelectedIndex == 0 || ddLlocales.SelectedIndex == 0 )
            {
                RadMessageBox.Show("Debe Seleccionar Empresa y  Local", "Local", MessageBoxButtons.OK);           
            }
            else
            {
                this.Enabled = false;
                this.Refresh();
                FuncionesClass fu = new FuncionesClass();
                if(fu.PingToHost(lblServidorVentas.Text) == true)
                {
                    CargaCajasByLocal();
                }
                else
                {
                    RadMessageBox.Show(this, "No se puede establecer conexión con el Host de Destino [" + lblServidorVentas.Text + "]" , "Atencion", MessageBoxButtons.OK);
                }

                this.Enabled = true;
                this.Refresh();             
            }
        }
        
        private void CargaCajasByLocal()
        {
            gvInforme.Rows.Clear();

            gvInforme.Rows.Add("33", "Factura de Venta Electronica", "", "", "");
            gvInforme.Rows.Add("52", "Guia de Despacho Electrónica", "", "", "");
            gvInforme.Rows.Add("61", "Nota de Crédito Electronica", "", "", "");


            gvInforme.Refresh();


            int i = 0;
            string tipo = "";
            double hasta = 0;
            double ultimo = 0;
            double disponibles = 0;

            for (i=0; i <= gvInforme.Rows.Count -1; i++)
            {
                tipo = gvInforme.Rows[i].Cells[0].Value.ToString();
                Caf myCaf = new Caf(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
                MySqlDataReader dr = myCaf.GetCafDisponiblesFacturasByLocal(tipo, ddLlocales.Text.Substring(0, 2));
                
                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        hasta = Convert.ToDouble(dr["hasta"].ToString());
                        ultimo = Convert.ToDouble(dr["ultimo"].ToString());

                        disponibles = (hasta - ultimo);
                                               
                        gvInforme.Rows[i].Cells[2].Value = ultimo;
                        gvInforme.Rows[i].Cells[3].Value = hasta;
                        gvInforme.Rows[i].Cells[4].Value = disponibles;
                    }
                }

                dr.Close();
                myCaf.CerrarTransaccion();


            }



















            //Caf myCaf = new Caf(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            //MySqlDataReader dr = myCaf.GetCafDisponiblesFacturasByLocal(TipoBoleta,ddLlocales.Text.Substring(0, 2));
            //double quedan = 0;

            //gvInforme.Rows.Clear();
            //if(dr.HasRows == true)
            //{

            //    while(dr.Read())
            //    {
            //        quedan = Convert.ToDouble(dr["folios"].ToString()) - Convert.ToDouble(dr["ultimo"].ToString());
            //        gvInforme.Rows.Add(dr["numero"].ToString(), dr["tipo"].ToString(), dr["descripcion"].ToString(), dr["ultimo"].ToString(),
            //                           dr["folios"].ToString(), quedan, dr["nombrearchivo"].ToString());

            //        FuncionesClass fun = new FuncionesClass();
            //        if (quedan <= Convert.ToDouble(txtCriticidad.Text))
            //        {
            //            fun.ColoreaCelda(gvInforme.Rows[gvInforme.CurrentRow.Index].Cells[5], Color.Red);
            //        }
            //        else
            //        {
            //            fun.ColoreaCelda(gvInforme.Rows[gvInforme.CurrentRow.Index].Cells[5], Color.YellowGreen);
            //        }              

            //    }

            //    ddlEmpresas.Enabled = true;
            //    ddLlocales.Enabled = true;
            //    btnVer.Enabled = false; 
            //}

            //if(gvInforme.Rows.Count > 0 )
            //{
            //    gvInforme.Rows[0].IsCurrent = true;
            //    gvInforme.ClearSelection();
            //}


            //dr.Close();
            //myCaf.CerrarTransaccion();

        }
        private void CargaCafCaja(string xlocal, string xCaja)
        {
            Caf myCaf = new Caf(lblServidorVentas.Text, lblRoot.Text, lblPassword.Text);
            MySqlDataReader dr = myCaf.GetCaFByLocalCaja(xlocal, xCaja);

            gvInforme.Rows.Clear();
            if (dr.HasRows == true)
            {
                while (dr.Read())
                {
                    gvInforme.Rows.Add(dr["tipo"].ToString(), dr["fecharecepcion"].ToString(), dr["desde"].ToString(), dr["hasta"].ToString(), dr["nombredelarchivo"].ToString());
                }
            }

            dr.Close();
            myCaf.CerrarTransaccion();
        }
        private void txtCaja_Leave(object sender, EventArgs e)
        {
            ddlEmpresas.Enabled = false;
            ddLlocales.Enabled = false;
        }

        private void txtCaja_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == (char)13)
            {
                txtCaja.Text = txtCaja.Text.PadLeft(2, Convert.ToChar("0"));
                btnVer.Enabled = true;
            }
            
        }

        private void gvInforme_Click(object sender, EventArgs e)
        {

        }

        private void gvInforme_CellClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            int row = 0;
            string caja = "";
            string nombrecaja = "";


            if(gvInforme.Rows.Count > 0 )
            {
                row = gvInforme.CurrentRow.Index;

                caja = gvInforme.Rows[row].Cells[0].Value.ToString();
                nombrecaja = gvInforme.Rows[row].Cells[1].Value.ToString();

                groupCaf.Enabled = true;

                lblcodigoCajaCaf.Text = caja;
                lblDescripcionCajaCaf.Text = nombrecaja;
            }
                      

        }

        private void rbAfectas_ToggleStateChanged(object sender, Telerik.WinControls.UI.StateChangedEventArgs args)
        {
            if(rbAfectas.CheckState == CheckState.Checked)
            {
                TipoBoleta = "39";
            }
            else
            {
                TipoBoleta = "41";
            }
        }

        private void rbExentas_ToggleStateChanged(object sender, Telerik.WinControls.UI.StateChangedEventArgs args)
        {
            if (rbExentas.CheckState == CheckState.Checked)
            {
                TipoBoleta = "41";
            }
            else
            {
                TipoBoleta = "39";
            }
        }
    }
}
