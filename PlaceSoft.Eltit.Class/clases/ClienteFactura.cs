using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlaceSoft.Eltit.Class.clases
{
    public class ClienteFactura
    {
        /*************************************************
        * DATOS DEL CLIENTE
        *************************************************/
        private string rut;
        private string sucursal;
        private string nombre;
        private string direccion;
        private string comuna;
        private string ciudad;
        private string fono1;
        private string celular;
        private string giro;
        private string email;
        private string contacto;
        
        /*************************************************
        * DATOS PARA LA CONEXION A LA BASE DE DATOS
        *************************************************/
        private Conectar cnn;
        private string user_root;
        private string mysql_server;
        private string mysql_password;
        private string local;
        /************************************************/
        public ClienteFactura(string xServidor, string xUserRoot)
        {
            this.user_root = xUserRoot;
            this.mysql_server = xServidor;
            this.mysql_password = "desarrollo_1990";
        }

        public MySqlDataReader GetClienteByRutSucursal(string xRut, string xSucursal)
        {
            string salida = "";
            string query = "";

            MySqlDataReader dr = null;

            query  = " SELECT rut, sucursal, nombre, direccion, comuna, ciudad, fono1, celular, giro, IF(email,'', 'SIN DATOS') email, contacto ";
            query += " FROM sv_maestroclientes ";
            query += " WHERE rut = '"+ xRut +"' AND sucursal = '"+ xSucursal +"' ";

            cnn = new Conectar(this.mysql_server, "eltit_ventas", this.user_root, this.mysql_password);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();

            }
            return dr;
        }

        /* BASE DE DATOS */
        public string User_root { get => user_root; set => user_root = value; }
        public string Mysql_server { get => mysql_server; set => mysql_server = value; }
        public string Mysql_password { get => mysql_password; set => mysql_password = value; }
        public string Local { get => local; set => local = value; }
        /* CLIENTE */
        public string Rut { get => rut; set => rut = value; }
        public string Sucursal { get => sucursal; set => sucursal = value; }
        public string Nombre { get => nombre; set => nombre = value; }
        public string Direccion { get => direccion; set => direccion = value; }
        public string Comuna { get => comuna; set => comuna = value; }
        public string Ciudad { get => ciudad; set => ciudad = value; }
        public string Fono1 { get => fono1; set => fono1 = value; }
        public string Celular { get => celular; set => celular = value; }
        public string Giro { get => giro; set => giro = value; }
        public string Email { get => email; set => email = value; }
        public string Contacto { get => contacto; set => contacto = value; }
        
    }
}
