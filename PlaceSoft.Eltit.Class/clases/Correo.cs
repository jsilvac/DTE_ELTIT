using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlaceSoft.Eltit.Class.clases
{
    public class Correo
    {
        Conectar cnn;
        private string CLIENTE_PREFIX = "eltit_";
        private string MYSQL_SERVER = "";
        private string MYSQL_ROOT = "";
        private string MYSQL_PASS = "";
       

        //private string rut;
        //private string local;


        public Correo(string xServer, string xRoot, string xPass)
        {

            this.MYSQL_SERVER = xServer;
            this.MYSQL_PASS = xPass;
            this.MYSQL_ROOT = xRoot;

        }

        public MySqlDataReader GetCorreByERutEmpresa(string xRutEmpresa)
        {
            string query = "";
            MySqlDataReader dr = null;


            query = "SELECT mailsalida,clavemail,servermail  ";
            query += " FROM maestroempresas ";
            query += " WHERE rut='"+ xRutEmpresa +"' ";
          //  query += " ";
            try
            {
                cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "conta" , "sistema", this.MYSQL_PASS);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                   
                }

                //cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }

            return dr;
        }

        public MySqlDataReader GetCorreInetcambioCli(string xRut)
        {
            
            MySqlDataReader dr = null;
            string query = "";

            xRut = xRut.Substring(1, xRut.Length-1); 

            query = " SELECT *FROM sv_fae_proveedores ";
            query += " WHERE rut LIKE '%"+ xRut +"'  ";
            try
            {
                cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "fae", MYSQL_ROOT, MYSQL_PASS);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }

            return dr;
        }

        public MySqlDataReader GetMaestroClientes(string xRut)
        {
            xRut = xRut.Replace("-","");
            MySqlDataReader dr = null;
            string query = "SELECT *FROM sv_maestroclientes WHERE rut='"+xRut+"' ";

            try
            {
                cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas", MYSQL_ROOT, MYSQL_PASS);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }
            return dr;
        }

        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }





        public string CLIENTE_PREFIX1 { get => CLIENTE_PREFIX; set => CLIENTE_PREFIX = value; }
        public string MYSQL_SERVER1 { get => MYSQL_SERVER; set => MYSQL_SERVER = value; }
        public string MYSQL_ROOT1 { get => MYSQL_ROOT; set => MYSQL_ROOT = value; }
        public string MYSQL_PASS1 { get => MYSQL_PASS; set => MYSQL_PASS = value; }
    }
}

namespace PlaceSoft.Eltit.Class.clases
{
    //public class Correo
    //{
    //}
}