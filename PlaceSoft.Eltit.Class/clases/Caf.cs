using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace PlaceSoft.Eltit.Class.clases
{
   public class Caf
    {
        Conectar cnn;
        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;
        private string CLIENTE = "eltit_";

        private static readonly log4net.ILog log =
          log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Caf(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }


        public MySqlDataReader GetCafByCajaLocal(string xLocal, string xCaja, string XtIPO)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT * FROM  sv_caf" + xLocal + "";
            query += " where local = '" + xLocal + "' and tipo = '" + XtIPO + "'  ";
            if (xCaja != "")
            {
                query += " AND caja = '" + xCaja + "' ";
            }

            query += " ORDER BY hasta DESC LIMIT 0,200 ";

            cnn = new Conectar(this.SERVER, CLIENTE + "fae" + xLocal, MYSQL_ROOT, MYSQL_PASS);
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
