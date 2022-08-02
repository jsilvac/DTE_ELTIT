using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlaceSoft.Eltit.Class
{
    public class Conectar
    {
        public MySqlConnection connection;
        private string server;
        private string database;
        private string uid;
        private string password;


        //Constructor
        public Conectar(string xservidor, string xdatabase, string xusuario, string xpassword)
        {
            Initialize(xservidor, xdatabase, xusuario, xpassword);
        }
        public Conectar(string xservidor, string xdatabase, string xusuario, string xpassword, int timeout)
        {
            Initialize(xservidor, xdatabase, xusuario, xpassword, timeout);
        }

        private void Initialize(string xservidor, string xbase, string xusuarios, string xpassword, [Optional] int timeout)
        {
            server = xservidor;
            database = xbase;
            uid = xusuarios;
            password = xpassword;
            string connectionString;
            connectionString = "SERVER=" + server + ";" + "DATABASE=" +
            database + ";" + "UID=" + uid + ";" + "PWD=" + password + ";Allow Zero Datetime=True;Allow User Variables=True;Respect Binary flags=false;";


            if (timeout > 0)
            {
                connectionString = connectionString + ";Connect Timeout=" + timeout + ";";
            }
            else
            {
                connectionString = connectionString + ";default command timeout=50";
            }


            connection = new MySqlConnection(connectionString);
        }

        public bool OpenConnection()
        {
            try
            {
                connection.Open();
                return true;

            }

            catch (MySqlException ex)
            {
                //When handling errors, you can your application's response based 
                //on the error number.
                //The two most common error numbers when connecting are as follows:
                //0: Cannot connect to server.
                //1045: Invalid user name and/or password.
                // this.CloseConnection();
                switch (ex.Number)
                {
                    case 0:

                        //MessageBox.Show("SE HA PERDIDO CONECCION. PRESIONE ACEPTAR PARA VOLVER A CONECTAR.");
                        //MessageBox.Show(ex.Message.ToString());
                        //Application.Restart();                        
                        //this.ReloadForm();
                        this.CloseConnection();
                        connection.Dispose();
                        MySqlConnection.ClearPool(connection);
                        break;

                    case 1045:
                        MessageBox.Show("USUARIO O CONTRASEÑA DE SERVIDOR INCORRECTA");
                        break;
                }
                return false;
            }
            //finally
            //{
            //    //con.Close();
            //    //con.Dispose();
            //    //SqlConnection.ClearPool(con);
            //    this.CloseConnection();
            //    connection.Dispose();
            //    MySqlConnection.ClearPool(connection);
            //}
        }


        //Close connection
        public bool CloseConnection()
        {
            try
            {
                connection.Close();
                connection.Dispose();
                MySqlConnection.ClearPool(connection);
                return true;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private void ReloadForm()
        {
            //try
            //{
            //    foreach (Form frm in Application.OpenForms)
            //            {

            //                if (frm.GetType() == typeof(frmMasterEntry)  )
            //                {
            //                    // MessageBox.Show("El formulario 2 esta abierto");
            //                    frm.Close();
            //                    frm.Dispose();
            //                    frmMasterEntry f = new frmMasterEntry();
            //                    f.Show();
            //                    break;

            //                }

            //            }
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message.ToString());
            //}

        }

    }
}
