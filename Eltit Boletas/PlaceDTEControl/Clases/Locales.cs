using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eltit.Clases
{
    class Locales
    {
        Conectar cnn;
        private string MYSQL_SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;

        public Locales(string xServer, string xroot, string xpass)
        {
            this.MYSQL_ROOT = xroot;
            this.MYSQL_PASS = xpass;
            this.MYSQL_SERVER = xServer;
        }

        public MySqlDataReader GetLocalByCodigo(string xCodigolocal)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select * clientes_locales From  ";

            if (xCodigolocal != "")
            {
                query += " Where codigo  = '" + xCodigolocal + "' ";
            }
            query += " Order by codigo ";


            cnn = new Conectar(MYSQL_SERVER, "fae_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader getLocalesByCodigoContable(string xCodEmpresa)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_locales ";

            if (xCodEmpresa != "")
            {
                query += "Where codigo_contable = '" + xCodEmpresa + "' And emite_39 = 1 ";
            }
            query += " Order by codigo ";


            cnn = new Conectar(MYSQL_SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
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
