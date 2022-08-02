using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eltit.Clases
{
    class Empresas
    {

        Conectar cnn;
        private string MYSQL_SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;

        public Empresas(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.MYSQL_SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }
        public MySqlDataReader GetEmpresasBoleta()
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte  where activo = 1 ";
            query += "order by codigo_contable ";


            cnn = new Conectar(this.MYSQL_SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GetDatoEmpresaByRut(string xCliente, string xRut)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte ";
            query += "Where prefijo = '" + xCliente + "' and rut = '" + xRut + "' ";

            cnn = new Conectar(this.MYSQL_SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        //public MySqlDataReader GetUsuario(string xuser, string xpass)
        //{
        //    string query = "";
        //    MySqlDataReader dr = null;

        //    query = "Select *  From usuarios ";
        //    query += "Where usuario = '" + xuser + "'  and password='"+  xpass +"'";

        //    cnn = new Conectar(this.MYSQL_SERVER, "eltit_users", MYSQL_ROOT, MYSQL_PASS);
        //    if (cnn.OpenConnection() == true)
        //    {
        //        MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
        //        dr = cmd.ExecuteReader();
        //    }

        //    return dr;
        //}


        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }

    }
}
