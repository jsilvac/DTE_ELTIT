using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace Eltit.Clases
{
    class Intercambio
    {
        Conectar cnn;
        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;
        private static readonly log4net.ILog log =
          log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Intercambio(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }

        public int IngresaRegistroIntercambio(List<string> xLineas, ref Label xLabel, string xBase)
        {
            string query = "";

            string[] campos;
            string rut = "";
            string nombre = "";
            string mail = "";
            string fecha_resol = "";
            string num_resol = "";
            int count = 0;

            try
            {
                //cnn = new ConectarClass(this.SERVER,  this._CLIENTE, MYSQL_ROOT, MYSQL_PASS);

                cnn = new Conectar(this.SERVER, xBase, MYSQL_ROOT, MYSQL_PASS, 180);

                if (cnn.OpenConnection() == true)
                {
                    foreach (var a in xLineas)
                    {
                        campos = a.Split(Convert.ToChar(";"));
                        if (campos.Length == 6)
                        {
                            rut = campos[0];
                            nombre = campos[1];
                            mail = campos[4];
                            nombre = nombre.Replace("'", "");

                            num_resol = campos[2];
                            fecha_resol = campos[3];


                            //ALTER TABLE `eltit_fae`.`sv_fae_proveedores` ADD COLUMN `fecha_actualizacion` DATE DEFAULT '0000-00-00' NOT NULL AFTER `mailintercambio`; 
                            query = "REPLACE INTO sv_fae_proveedores(rut, razonsocial, numeroresolucion, ";
                            query += " fecharesolucion, mailintercambio, fecha_actualizacion ) ";
                            query += " Values(";
                            query += " '" + rut + "', '" + nombre + "', '" + num_resol + "' , ";
                            query += " '" + fecha_resol + "', '" + mail + "','" + DateTime.Now.ToString("yyyy-MM-dd") + "' )";
                            


                            MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                            cmd.ExecuteNonQuery();
                            count++;

                            xLabel.Text = "Procesando " + count + " de " + xLineas.Count + " Contribuyentes";
                            xLabel.Refresh();
                        }

                    }

                }

                cnn.CloseConnection();

                xLabel.Text = "SE PROCESARON " + count + " REGISTROS DE " + xLineas.Count + " Contribuyentes";
                MessageBox.Show(xLabel.Text);

            }
            catch (Exception ex)
            {
                log.Error(ex);
                MessageBox.Show(ex.Message);
            }

            return count;
        }


    }
}
