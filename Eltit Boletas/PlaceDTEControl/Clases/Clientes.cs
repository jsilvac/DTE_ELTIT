using MySql.Data.MySqlClient;
using System;

namespace Eltit.Clases
{
    class Clientes
    {
        Conectar cnn;

        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;

        public Clientes(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }

        public MySqlDataReader getClienteByRutSucursal(string xRut, string xSucursal)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From sv_maestroclientes ";
            query += "Where rut = '" + xRut + "' and sucursal = '" + xSucursal + "' ";

            cnn = new Conectar(this.SERVER, "eltit_ventas", this.MYSQL_ROOT , this.MYSQL_PASS );
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GetEmpresaByCodigo(string xCodigoEmpresa)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte ";
            query += "Where codigo_contable = '" + xCodigoEmpresa + "'  ";

            cnn = new Conectar(SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GetClientesDTE()
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte ";
            query += "Where activo  = 1  AND envia_boletas_sii='1' ";

            cnn = new Conectar(SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GetClientesDTE_AD()
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte ";
            query += "Where activo  = 1  AND envia_boletas_sii='0' ";

            cnn = new Conectar(SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        public MySqlDataReader GetLocalesByEmpresa(string xCodigoEmpresa)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_locales ";
            query += "Where emite_39  = 1 and codigo_contable = '"+ xCodigoEmpresa +"'  ";

            cnn = new Conectar(SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GetLocalesByEmpresaAD(string xCodigoEmpresa)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_locales ";
            query += "Where   codigo_contable = '" + xCodigoEmpresa + "'  ";

            cnn = new Conectar(SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GetDatosLocalByCodigo(string xCodigocal)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_locales ";
            //query += "Where emite_39  = 1 and codigo = '" + xCodigocal + "'  ";
            query += "Where  codigo = '" + xCodigocal + "'  ";

            cnn = new Conectar(SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        public MySqlDataReader GetClientesByPrefijoRut(string xCliente, string xRut)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte ";
            query += "Where prefijo = '"+ xCliente +"' and rut = '"+ xRut +"' ";

            cnn = new Conectar(FuncionesClass.G_SERVIDOR, "eltit_dte_manager", FuncionesClass.G_MYSQL_USER, FuncionesClass.G_MYSQL_PASS);
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
