using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eltit.Clases
{
    class Consultas_BD
    {
        Conectar cnn;

        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;

        public Consultas_BD(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }


        public MySqlDataReader Getfecha(string xCodigoEmpresa, string xdia)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "SELECT NOW()";

            cnn = new Conectar(SERVER, "mysql" , MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        

    }
}
