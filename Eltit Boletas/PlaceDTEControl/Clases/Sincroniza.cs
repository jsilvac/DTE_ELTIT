using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eltit.Clases
{
    class Sincroniza
    {
        Conectar cnn;
        private string MYSQL_SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;

        public Sincroniza(string xServer, string xRoot, string xPass)
        {

            this.MYSQL_ROOT = xRoot;
            this.MYSQL_PASS = xPass;
            this.MYSQL_SERVER = xServer;
        }

        public void GrabaSincronizador(string xcadena, string xbdatos)
        {

            string query = "";

            if (xbdatos.Contains("_fae"))
            {
                xcadena = xcadena + " ON DUPLICATE KEY UPDATE local = local ";
            }
            
            xcadena = xcadena.Replace("'", "~");
            query = "INSERT INTO sincronizador (";
            query += "servidor,consulta,basedatos,fechacreacion,horacreacion) VALUES (";
            query += " '" + MYSQL_SERVER + "','" + xcadena + "','" + xbdatos + "', NOW(),current_time() )  ";

            Conectar cnn = new Conectar(this.MYSQL_SERVER, FuncionesClass.G_CLIENTE_PREFIJO + "sincroniza", this.MYSQL_ROOT, this.MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }

            cnn.CloseConnection();

        }
        public void GrabaSincronizadorMaster(string xcadena, string xbdatos)
        {

            string query = "";

            if (xbdatos.Contains("_fae") )
            {
                xcadena = xcadena + " ON DUPLICATE KEY UPDATE local = local ";
            }


            xcadena = xcadena.Replace("'", "~");
            query = "INSERT INTO sincronizador_master (";
            query += "servidor,consulta,basedatos,fechacreacion,horacreacion) VALUES (";
            query += " '" + MYSQL_SERVER + "','" + xcadena + "','" + xbdatos + "', NOW(),current_time() )  ";

            Conectar cnn = new Conectar(this.MYSQL_SERVER, FuncionesClass.G_CLIENTE_PREFIJO + "sincroniza", this.MYSQL_ROOT, this.MYSQL_PASS );
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }

            cnn.CloseConnection();
        }
    }
}
