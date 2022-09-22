using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System;
using System.Windows.Forms;

namespace Eltit.DTE.clases
{
    class VentasClass
    {
        private string cliente;
        private string mysql_server;
        private string mysql_password;
        private string rut;
        private string local;


        public VentasClass( string xServer, string xPass, string xRoot)
        {
            this.cliente = "eltit_";
            this.mysql_server = xServer;
            this.mysql_password = "desarrollo_1990";
            this.rut = "";
            this.local = "";
        }

    
        public string GetXMLFacturas(string xlocal, string xTipo_Sii, string xFolio, string xFecha)
        {
            string salida = "0";

            string query = "";
            MySqlDataReader dr;
            query = " SELECT * from sv_dte"+ xlocal;
            query += " WHERE  localdocumento = '" + xlocal + "' and tipo='"+ xTipo_Sii +"' and fecha='"+ xFecha +"' and numero='" + xFolio + "' ";

            try
            {
                Conectar cnn = new Conectar(Mysql_server, Cliente + "fae"+ xlocal, "sistema", this.Mysql_password);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        if (dr.Read())
                        {
                            salida = dr["xml"].ToString();
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
        public string GetXMLGuias(string xlocal, string xTipo_Sii, string xNumeroD, string xFecha)
        {
            string salida = "0";

            xNumeroD = xNumeroD.PadLeft(10, Convert.ToChar("0"));
            string query = "";
            MySqlDataReader dr;
            query = " SELECT * from sv_dte" + xlocal;
            query += " WHERE  localdocumento = '" + xlocal + "' and tipo='" + xTipo_Sii + "' and fecha='" + xFecha + "' and numero='" + xNumeroD + "' ";

            try
            {
                Conectar cnn = new Conectar(Mysql_server, Cliente + "fae" + xlocal, "sistema", this.Mysql_password);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        if (dr.Read())
                        {
                            salida = dr["xml"].ToString();
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

        public string GetCajera(string xRut)
        {
            string salida = "0";

            string query = "";
            MySqlDataReader dr;
            query = " SELECT nombre from sv_maestrocajeras ";
            query += " WHERE rut ='"+ xRut +"' limit 0,1 ";

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

        public string GetCajeraX(string xRut)
        {
            string salida = "0";

            string query = "";
            MySqlDataReader dr;


            query = " SELECT nombre from sv_maestrocajeras ";
            query += " WHERE rut like '" + xRut + "%' limit 0,1 ";

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
        public MySqlDataReader GetDocumentoCabeza(string xLocal, string xTipo, string xFolio, string xCaja, string xFecha)
        {
            string salida = "";

            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT * from sv_documento_cabeza_"+ xLocal;
            query += " WHERE local='" + xLocal + "' AND ";
            query += " tipo='" + xTipo + "' AND ";
            query += " numero=lpad('" +xFolio+ "',10,'0') AND ";
            query += " caja='" + xCaja + "' AND ";
            query += " fecha='" + xFecha + "' LIMIT 1";

            try
            {
                Conectar cnn = new Conectar(Mysql_server, Cliente + "ventas"+ xLocal, "sistema", this.Mysql_password);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        return dr;
                    }
                }

                cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }

            return dr;
        }


        public string Cliente { get => cliente; set => cliente = value; }
        public string Mysql_server { get => mysql_server; set => mysql_server = value; }
        public string Mysql_password { get => mysql_password; set => mysql_password = value; }
        public string Rut { get => rut; set => rut = value; }

    }
}
