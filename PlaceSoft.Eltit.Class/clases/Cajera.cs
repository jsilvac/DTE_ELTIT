using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlaceSoft.Eltit.Class.clases
{
    public class Cajera
    {
        private string cliente;
        private string mysql_server;
        private string mysql_password;
        private string rut;
        private string local;


        public Cajera(string xRut, string xCliente, string xServer, string xLocal)
        {
            this.cliente = xCliente;
            this.mysql_server = xServer;
            this.mysql_password = "desarrollo_1990";
            this.rut = xRut;
            this.local = xLocal;
        }

        public string GetCajera(string xRut)
        {
            string salida = "0";

            string query = "";
            MySqlDataReader dr;
            query = " SELECT nombre from sv_maestrocajeras ";
            query += " WHERE rut ='" + xRut + "' limit 0,1 ";

            try
            {
                Conectar cnn = new Conectar(Mysql_server, Cliente + "ventas", "sistema", this.Mysql_password);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        if (dr.Read())
                        {
                            salida = dr["nombre"].ToString();
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
        public string Cliente { get => cliente; set => cliente = value; }
        public string Mysql_server { get => mysql_server; set => mysql_server = value; }
        public string Mysql_password { get => mysql_password; set => mysql_password = value; }
        public string Rut { get => rut; set => rut = value; }

    }
}
