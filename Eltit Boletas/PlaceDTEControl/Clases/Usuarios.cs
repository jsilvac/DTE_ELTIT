using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eltit.Clases;
using MySql.Data.MySqlClient;


namespace Eltit.clases
{
    class Usuarios
    {
        Conectar cnn;
        private string MYSQL_SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;

        public Usuarios(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.MYSQL_SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }

        public MySqlDataReader GetUsuario(string xuser, string xpass)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From usuarios ";
            query += "Where usuario = '" + xuser + "'  and password='" + xpass + "'";

            cnn = new Conectar(this.MYSQL_SERVER, "eltit_users", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GetPermisosUsuario(string xuser)
        {
            MySqlDataReader dr = null;
            string query = "";
            DataTable dt = new DataTable();

            query = "Select *  From programas ";
            query += "Where usuario = '" + xuser + "' and acceso='1' ";

            cnn = new Conectar(this.MYSQL_SERVER, "eltit_users", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
               
            }

            return dr;
        }


        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }

    }
}
