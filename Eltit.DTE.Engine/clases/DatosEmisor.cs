using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Eltit.DTE.clases
{
    public class DatosEmisor
    {
        private string rut;

        private string razon;
        private string direccion;
        private string comuna;
        private string giro_1;
        private string giro_2;
        private string giro_3;
        private string giro_4;
        private string glosa_res;
        private string web_verificacion;
        private string sii;
        private string fono;
        private string email;
        private string cliente;
        private string img;
        private string mysql_server;
        private string mysql_password;
        private string codigo_sucursal_sii;

        public DatosEmisor(string xRut, string xCliente, string xServer, string xLocal)
        {
            this.mysql_server = xServer;
            this.cliente = xCliente;

            mysql_password = "desarrollo_1990";
            this.CargaEmisor(xRut, xLocal);
        }
        private void CargaEmisor(string xRut, string xLocal)
        {
            string query = "";
            MySqlDataReader dr;
            query = " SELECT * from maestroempresas ";
            query += " WHERE  rut = '" + xRut + "'  ";

            try
            {
                Conectar cnn = new Conectar(mysql_server, cliente + "conta", "sistema", this.mysql_password);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        if (dr.Read())
                        {
                            this.Rut = dr["rut"].ToString();
                            this.Razon = dr["nombre"].ToString();
                            this.Direccion = dr["direccion"].ToString();
                            this.Comuna = dr["comuna"].ToString();
                            this.Giro_1 = dr["girodte"].ToString();
                            this.Giro_2 = "";//dr["giro2"].ToString();
                            this.Giro_3 = "";// dr["giro3"].ToString();
                            this.Giro_4 = "";// dr["giro4"].ToString();
                            this.Glosa_res = "Res. " + dr["numeroresolucion"].ToString() + " de " + dr["fecharesolucion"].ToString();
                            this.Web_verificacion = "SII.CL";// dr["web_verificacion"].ToString();
                            this.Sii = dr["oficina_sii"].ToString();
                            this.Fono = "452-379500";// dr["fono"].ToString();
                            //this.Img = dr["img"].ToString();
                            this.Email = dr["mailsalida"].ToString();
                            this.Codigo_sucursal_sii = this.LeeCodigoSucursalSII(xLocal);
                        }


                    }
                    dr.Close();
                }

                cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }


        }
        private string LeeCodigoSucursalSII(string xlocal)
        {
            string salida = "0";

            string query = "";
            MySqlDataReader dr;
            query = " SELECT * from clientes_locales ";
            query += " WHERE  codigo = '" + xlocal + "'  ";

            try
            {
                Conectar cnn = new Conectar(mysql_server, cliente + "conta", "sistema", this.mysql_password);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        if (dr.Read())
                        {
                            salida = dr["codigo_sucursal_sii"].ToString();
                        }


                    }
                    dr.Close();
                }

                cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }

            return salida;
        }
        public string Rut
        {
            get => rut;
            set => rut = value;
        }
        public string Razon { get => razon; set => razon = value; }
        public string Direccion { get => direccion; set => direccion = value; }
        public string Comuna { get => comuna; set => comuna = value; }
        public string Giro_1 { get => giro_1; set => giro_1 = value; }
        public string Giro_2 { get => giro_2; set => giro_2 = value; }
        public string Giro_3 { get => giro_3; set => giro_3 = value; }
        public string Giro_4 { get => giro_4; set => giro_4 = value; }
        public string Glosa_res { get => glosa_res; set => glosa_res = value; }
        public string Web_verificacion { get => web_verificacion; set => web_verificacion = value; }
        public string Sii { get => sii; set => sii = value; }
        public string Fono { get => fono; set => fono = value; }
        public string Email { get => email; set => email = value; }
        public string Img { get => img; set => img = value; }
        public string Codigo_sucursal_sii { get => codigo_sucursal_sii; set => codigo_sucursal_sii = value; }
    }
}
