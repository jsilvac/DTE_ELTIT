using MySql.Data.MySqlClient;
using PlaceSoft.Eltit.Class;
using SchoolManagementAdmin.objetos;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SchoolManagementAdmin
{
    public partial class PopBoletas : Form
    {
        public string LOCAL_ACTIVO;
        public DataTable DT_LOCALES;
        int i = 0;
        private bool cargo = false;
        
        public PopBoletas()
        {
            InitializeComponent();
        }

        private void PopBoletas_Load(object sender, EventArgs e)
        {
            gvPagos.TableElement.Font = new Font("Arial", 8);
            gvGrilla2.TableElement.Font = new Font("Arial", 8);

            gvPagos.TableElement.RowHeight = 21;
            gvGrilla2.TableElement.RowHeight = 21;
        }

        private void PopBoletas_Activated(object sender, EventArgs e)
        {
            if(cargo == false)
            {
                CargaVerificacionLocales();
                cargo = true;
                timer1.Enabled = true;
                timer1.Start();
            }
        }

        private void CargaVerificacionLocales()
        {
            LocalesClass loc = new LocalesClass();
            Locales localClass = new Locales(Inicial.G_SERVIDOR, Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
            localClass.getLocalDTE(LOCAL_ACTIVO);


            DataTable dt = loc.GetCajasByLocal("39", LOCAL_ACTIVO, localClass.IP_servidor, localClass.Mysql_user, localClass.Mysql_pass);
            double caf_39 = 0;
            string base_dte = "";
            Inicial fun = new Inicial();


            caf_39 = 0;
            base_dte = "eltit_fae" + LOCAL_ACTIVO;
            // gvPagos.Rows[i].Cells[6].Value = caf_39 ;

            int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                caf_39 = 0;
                caf_39 = this.VerificaFoliosBoletas("39", LOCAL_ACTIVO, "eltit_", base_dte,
                                 localClass.IP_servidor, localClass.Mysql_user, localClass.Mysql_pass, row["caja"].ToString());

                if(i <= 25)
                {
                    gvPagos.Rows.Add(row["caja"].ToString(), "CAJA " + row["caja"].ToString(), localClass.IP_servidor, "", caf_39.ToString());


                    //gvPagos.Rows[i].Cells[3].Value = caf_39;
                    if (caf_39 <= Convert.ToDouble(localClass.Caf_critico_39))
                    {
                        fun.ColoreaCeldaYTexto(gvPagos.Rows[gvPagos.RowCount - 1].Cells[4], Color.Red, Color.Black, new Font("Arial", 8, FontStyle.Bold));
                        if (row["caja"].ToString() != "50")
                        {
                            this.EnviarEmailFolios(localClass.Rut, lblNombreLocal.Text, localClass.IP_servidor,
                            "39", Convert.ToDouble(localClass.Caf_critico_39), caf_39, row["caja"].ToString(), localClass.Correo_soporte);
                        }
                    }
                    i++;
                }
                else
                {
                    gvGrilla2.Rows.Add(row["caja"].ToString(), "CAJA " + row["caja"].ToString(), localClass.IP_servidor, "", caf_39.ToString());


                    //gvPagos.Rows[i].Cells[3].Value = caf_39;
                    if (caf_39 <= Convert.ToDouble(localClass.Caf_critico_39))
                    {
                        fun.ColoreaCeldaYTexto(gvGrilla2.Rows[gvGrilla2.RowCount - 1].Cells[4], Color.Red, Color.Black, new Font("Arial", 8, FontStyle.Bold));
                        if (row["caja"].ToString() != "50")
                        {
                            this.EnviarEmailFolios(localClass.Rut, lblNombreLocal.Text, localClass.IP_servidor,
                            "39", Convert.ToDouble(localClass.Caf_critico_39), caf_39, row["caja"].ToString(), localClass.Correo_soporte);
                        }
                    }
                    i++;
                }
               
            }


          

        }
        private void EnviarEmailFolios(string xRut, string xLocal, string xServidor, string xTipo, double xCritico, double xCurrCaf, 
                                       string xCaja, string xcorreo_soporte)
        {

            string htmlString = @"<html>";
         
            htmlString = htmlString + "<body>";
            htmlString = htmlString + "<img src='http://www.placesoft.cl/images/eltit/header_eltit.png' border='0'  />";
            htmlString = htmlString + "<p>Se Han detectado Folios Críticos</p>";
            htmlString = htmlString + "--------------------------------------------<br>";

            htmlString = htmlString + "<div style='overflow-x:auto; font-family:Arial, Helvetica, sans-serif; color:#666; font-size:12px;'> ";
            htmlString = htmlString + " <table  border='0'> ";
            htmlString = htmlString + " <tr> ";
            htmlString = htmlString + " <td><strong>Rut</strong></td> ";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + Convert.ToDouble(xRut.Substring(0, 9)).ToString() + "-" + xRut.Substring(9, 1) + "</td> ";
            htmlString = htmlString + " </tr> ";
            htmlString = htmlString + " <tr> ";
            htmlString = htmlString + " <td><strong>Local</strong></td> ";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + xLocal + "</td> ";
            htmlString = htmlString + " </tr> ";
            htmlString = htmlString + " <tr> ";
            htmlString = htmlString + " <td><strong>Servidor</strong></td> ";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + xServidor + "</td> ";
            htmlString = htmlString + " </tr>";
            htmlString = htmlString + " <tr>";
            htmlString = htmlString + " <td><strong>Tipo Doc</strong></td>";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + xTipo + "</td>";
            htmlString = htmlString + " </tr>";
            htmlString = htmlString + " <tr>";
            htmlString = htmlString + " <td><strong>Caja Doc</strong></td>";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>" + xCaja + "</td>";
            htmlString = htmlString + " </tr>";
            htmlString = htmlString + " <tr>";
            htmlString = htmlString + " <td><strong>Critico</strong></td>";
            htmlString = htmlString + " <td bgcolor='#F3F3F3'>[" + xCritico + "] Quedan[" + xCurrCaf + "]</td>";
            htmlString = htmlString + " </tr>";
    
            htmlString = htmlString + " </table>";
            htmlString = htmlString + " </div> <br>";



            //htmlString = htmlString + "--------------------------------------------<br>";
            //htmlString = htmlString + " Local &nbsp;&nbsp;&nbsp;    :" + xLocal + "<br>";
            //htmlString = htmlString + " Servidor   :" + xServidor + "<br>";
            //htmlString = htmlString + " Tipo Doc :" + xTipo + "<br>";
            //htmlString = htmlString + " Caja Doc :" + xCaja + "<br>";
            //htmlString = htmlString + " Critico         :[" + xCritico + "] Quedan[" + xCurrCaf + "]  <br>";
            htmlString = htmlString + "--------------------------------------------<br>";
            htmlString = htmlString + "<p>Enviado el " + DateTime.Now.ToString("dd-MM-yyyy") + " a las " + DateTime.Now.ToString("HH:mm:ss tt") + "</p>";
            htmlString = htmlString + "<img src='http://www.placesoft.cl/images/eltit/footer_eltit.png' border='0'  />";
            htmlString = htmlString + "</body>";
            htmlString = htmlString + " </html>";

            Inicial.EnviarEmail(xcorreo_soporte, Inicial.G_CORREO_SOPORTE_COPIA, "CAF " + xTipo, "CAF CRITICOS EN " + xLocal, htmlString);

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

        private void timer1_Tick(object sender, EventArgs e)
        {
           

            if(i > 20)
            {
                timer1.Stop();
                timer1.Enabled = false;
                this.Close();
            }
            i++;
        }
    }
}
