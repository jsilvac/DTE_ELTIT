using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using MetroFramework;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;
using System.Reflection;
using System.Data;

namespace SchoolManagementAdmin.objetos
{
    class Sincroniza
    {
        //Conectar cnnWeb;
        private static readonly log4net.ILog log =
  log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        private ClienteDTE cliente;
        private string Prefijo_host;
        private string clave_destino;

        public  Sincroniza(ClienteDTE xCliente, string xPrefijoHost)
        {
            this.cliente = xCliente;
            this.Prefijo_host = xPrefijoHost;
            this.clave_destino = "lawila_321";
        }
        public void GrabaSincronizador(string xcadena, string xbdatos, string xprioridad)
        {

            string query = "";

            xcadena = xcadena.Replace("'", "~");
            query = "INSERT INTO log_track (";
            query += "server,query_str,basedatos,fecha_creacion,hora_creacion) VALUES (";
            query += " '" + cliente.IP_servidor + "','" + xcadena + "','" + xbdatos + "', NOW(),current_time() )  ";

            Conectar cnn = new Conectar(cliente.IP_servidor, cliente.Prefijo + "_log", cliente.Mysql_user, cliente.Mysql_pass);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }

            cnn.CloseConnection();

        }
        public int GetRowsCount(string cloud)
        {
            string query = "";
            int salida = 0;

            query = "SELECT count(id) FROM log_track WHERE  ";
            query = query + " " + cliente.RetornaWhere() + " ";          
            query = query + " and cloud_" + cliente.Cloud_up + " IS NULL ";

            Conectar cnn = new Conectar(cliente.IP_servidor,  cliente.Prefijo + "_log",cliente.Mysql_user, cliente.Mysql_pass);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();  

                if(dr.HasRows == true)
                {
                    if(dr.Read())
                    {
                        salida = Convert.ToInt32(dr[0]);
                    }
                }             
                dr.Close();
            }

            cnn.CloseConnection();
            return salida;

        }

        public int GetRowsBoletas(string cloud)
        {
            string query = "";
            int salida = 0;

            query = "SELECT count(id) FROM log_track WHERE  ";
            query = query + "" + cliente.RetornaWhere() + " ";
            query = query + "AND " + cliente.Cloud_up + " IS NULL ";

            Conectar cnn = new Conectar(cliente.IP_servidor, cliente.Prefijo + "log", cliente.Mysql_user, cliente.Mysql_pass);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        salida = Convert.ToInt32(dr[0]);
                    }
                }
                dr.Close();
            }

            cnn.CloseConnection();
            return salida;

        }

        public int SincronizaData()
        {
            string query = "";
            string palabras = "";
            string basereal = "";
            long id = 0;
            int conta = 0;
            DataTable dt = new DataTable();
            List<string> list = new List<string>();

            try
            {

                query = "SELECT * FROM log_track WHERE ";
                query = query + " "+ cliente.RetornaWhere() +" ";             
                query = query + " AND cloud_" + cliente.Cloud_up + " IS NULL ORDER BY id Limit " + cliente.Numero_registros +" ";

                Conectar cnn = new Conectar(cliente.IP_servidor, cliente.Prefijo + "_log", cliente.Mysql_user, cliente.Mysql_pass);

                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
             
                    MySqlDataReader dr = cmd.ExecuteReader();
                    dt.Load(dr);

                    dr.Close();
                    cnn.CloseConnection();

                }

                foreach (DataRow row in dt.Rows)
                {
                    System.Threading.Thread.Sleep(200);
                    palabras = row[2].ToString();
                    basereal = this.Prefijo_host + "_" + row[3].ToString();

                    if (palabras.Contains(cliente.Prefijo + "_local00."))
                    {
                        //basereal = this.Prefijo_host + cliente.Prefijo + "_local00";
                        palabras = palabras.Replace(cliente.Prefijo + "_local00.", "");
                    }

                    id = (long)Convert.ToDouble((row[0]));
                    if (id == 410584)
                    {
                        palabras = palabras;
                    }
                    this.TraspasaDatos(palabras, basereal, id);
                    conta = conta + 1;
                }
                log.Debug("SYNC: Se insertaron " + conta + " Registros en " + cliente.Prefijo);

             }
            catch (Exception ex)
            {
                log.Error("Error:", ex);
                //MessageBox.Show( "ERROR: " + MethodBase.GetCurrentMethod() + " - " + ex.Message.ToString(), "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                Inicial.G_ERROR = true;
            }

            return conta;

        }

        private void TraspasaDatos(string Xconsulta, string xbaseDatos, long Xid)
        {
            string consulta = "";
            int salida = 0;
            //string rut_base = Funciones.G_LOCAL_RUT.Substring(0, 9);
            consulta = Xconsulta.Replace("~", "'");

            if (consulta.Contains("UPDATE ") && consulta.Contains("SET ") )
            {                
                consulta = consulta.Replace("UPDATE ", "UPDATE IGNORE ");
            }
            if (consulta.Contains("INSERT ") && consulta.Contains("VALUES"))
            {
                if(!consulta.Contains("INSERT IGNORE"))
                {
                    consulta = consulta.Replace("INSERT ", "INSERT IGNORE ");
                }
               
            }
            Conectar cnn = new Conectar(cliente.Servidor_destino,xbaseDatos , "placesof_dte", this.clave_destino);
            if (cnn.OpenConnection() == true)                 
            {
                MySqlCommand cmd = new MySqlCommand(consulta, cnn.connection);

                salida = cmd.ExecuteNonQuery();
                cnn.CloseConnection();
                this.ActualizaFecha(Xid, cliente.Cloud_up);
            }
            else
            {
                //MessageBox.Show( "NO ES POSIBLE CONECTAR CON EL HOST REMOTO " + Funciones.G_WEBSERVER, "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
           


        }



        private void ActualizaFecha(long xid, string xnombre)
        {
            string query = "";

            try
            {
                query = "UPDATE log_track SET cloud_" + xnombre + "= NOW() ";
                query += "where id =" + xid + "";

                Conectar cnn = new Conectar(cliente.IP_servidor, this.cliente.Prefijo + "_log", cliente.Mysql_user, cliente.Mysql_pass);

                if (cnn.OpenConnection() == true)
                {
                    System.Threading.Thread.Sleep(200);
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    cmd.ExecuteNonQuery();
                }
                else
                {
                    log.Debug("No se pudo establecer conección con " + cliente.Prefijo);
                    return;
                }
                cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                log.Error("Error Al Actualizar:", ex);
                //MessageBox.Show( "ERROR: " + MethodBase.GetCurrentMethod() + " - " + ex.Message.ToString(), "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                Inicial.G_ERROR = true;
            }
           

        }

    }
}
