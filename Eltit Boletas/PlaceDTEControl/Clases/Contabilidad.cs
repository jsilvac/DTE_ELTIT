using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eltit.Clases
{
    class Contabilidad
    {
        Conectar cnn;

        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;

        public Contabilidad(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }

        public MySqlDataReader GetTotalByDiaEmpresa(string xCodigoEmpresa, string xdia)
        {
            string query = "";
            MySqlDataReader dr = null;
         
            query  = "SELECT 'conta',fecha,ifnull(SUM(exento),0) as exento, ifnull(SUM(total),0) as total FROM eltit_conta" + xCodigoEmpresa +".boletasdeventa   ";
            query += " WHERE fecha='"+ xdia +"' GROUP BY fecha ";
            query += "  ";

            cnn = new Conectar(SERVER, "eltit_conta" + xCodigoEmpresa, MYSQL_ROOT, MYSQL_PASS);
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
