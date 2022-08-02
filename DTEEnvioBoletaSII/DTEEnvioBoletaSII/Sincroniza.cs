using SchoolManagementAdmin.objetos;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using System.Threading;
using MetroFramework;


namespace SchoolManagementAdmin
{
    public partial class frmSincroniza : Form
    {
        public frmSincroniza()
        {
            InitializeComponent();
        }

        public frmMain main;

        string[] month = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };

        //For animated panels direction
        string optionsDirection = "down";
        string toastDirection = "down";
        string rightDirection = "right";

        //For animated panels timeout
        int optionsTimeOut = 0;
        int toastTimeOut = 0;
        int RightTimeOut = 0;

        //For animated panels position
        int optionsX;
        int optionsY;
        int rightX;
        int rightY;

        //method to set fullscreen
        private void setFullScreen()
        {
            int x = Screen.PrimaryScreen.Bounds.Width;
            int y = Screen.PrimaryScreen.Bounds.Height;
            Location = new Point(0, 0);
            Size = new Size(x, y);
        }

        //method to set the position of the main panel that holds the controls to center of the form.
        private void setMainPanelPosition()
        {
            int mX = (Width - pnlMain.Width) / 2;
            int mY = (Height - pnlMain.Height) / 2;
            pnlMain.Location = new Point(mX, mY);
        }

        //method to set the initial location of the options panel.
        private void setOptionsPanelPosition()
        {
            int x = Width;
            optionsX = 0;
            optionsY = Height + pnlOptions.Height;
            pnlOptions.Size = new Size(x, pnlOptions.Height);
            pnlOptions.Location = new Point(optionsX, optionsY);
            pbHide.Location = new Point(Width - 41, pbHide.Location.Y);
            int mX = (pnlOptions.Width - pnlOptionsMain.Width) / 2;
            int mY = pnlOptionsMain.Location.Y;
            pnlOptionsMain.Location = new Point(mX, mY);
            cmbAcademicYear.Location = new Point(Width - 125, cmbAcademicYear.Location.Y);
            lblAcademicYear.Location = new Point(Width - 226, lblAcademicYear.Location.Y);
        }

        private void setRightOptionsPanelPosition()
        {
            int y = Height;
            rightY = 0;
            rightX = Width + pnlRightOptions.Width;
            pnlRightOptions.Size = new Size(pnlRightOptions.Width, y);
            pnlRightOptions.Location = new Point(rightX, rightY);
            int rX = pnlRightMain.Location.X;
            int rY = (pnlRightOptions.Height - pnlRightMain.Height) / 2;
            pnlRightMain.Location = new Point(rX, rY);
        }

        private void setTimeLabelsPosition()
        {
            //int x = (pnlTimeTile.Width - lblDate.Width) / 2;
            //lblDate.Location = new Point(x, lblDate.Location.Y);
            //x = (pnlTimeTile.Width - lblDayOfWeek.Width) / 2;
            //lblDayOfWeek.Location = new Point(x, lblDayOfWeek.Location.Y);
            //x = (pnlTimeTile.Width - lblMonthYear.Width) / 2;
            //lblMonthYear.Location = new Point(x, lblMonthYear.Location.Y);
            //int y = (pnlCurrentStatus.Height - lblCurrentStatus.Height) / 2;
            //lblCurrentStatus.Location = new Point(lblCurrentStatus.Location.X, y);
        }

        private void frmTemplate_Load(object sender, EventArgs e)
        {
            //lblCurrentStatus.Text = "Current\nStatus";
            //lblDayOfWeek.Text = DateTime.Today.DayOfWeek.ToString();
            //lblMonthYear.Text = month[DateTime.Today.Month - 1] + ", " + DateTime.Today.Year.ToString();
            //lblDate.Text = DateTime.Today.Day.ToString();
            setTimeLabelsPosition();
            //lblIntroduction.Text = "This a software for demo purpose only. The final product will contain all the given options which include the \ndatabase connectivity.\nAt present the software does not contain database connectivity.";
            setFullScreen();
            setOptionsPanelPosition();
            setRightOptionsPanelPosition();
            setMainPanelPosition();
            Options.Start();
            RightOptions.Start();

        }

        private void frmTemplate_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Y >= Height - 15 && e.X < (Width - pnlRightOptions.Width))
            {
                optionsDirection = "up";
                rightDirection = "right";
                optionsTimeOut = 0;
            }
            if (e.X >= Width - 15)
            {
                rightDirection = "left";
                RightTimeOut = 0;
                optionsDirection = "down";
            }
            if (e.X < (Width - pnlRightOptions.Width))
            {
                rightDirection = "Left";
            }
        }

        private void Options_Tick(object sender, EventArgs e)
        {
            if (optionsTimeOut < 1000)
            {
                optionsTimeOut++;
            }
            if (optionsTimeOut == 1000)
            {
                if (optionsDirection == "up")
                {
                    optionsDirection = "down";
                }
            }
            if (optionsDirection == "up")
            {
                if (optionsY > Height - pnlOptions.Height + 3)
                {
                    optionsY -= 3;
                    pnlOptions.Location = new Point(optionsX, optionsY);
                }
            }
            else
            {
                if (optionsY < Height)
                {
                    optionsY += 3;
                }
                pnlOptions.Location = new Point(optionsX, optionsY);
            }
        }

        private void RightOptions_Tick(object sender, EventArgs e)
        {
            if (RightTimeOut < 1000)
            {
                RightTimeOut++;
            }
            if (RightTimeOut == 1000)
            {
                if (rightDirection == "left")
                {
                    rightDirection = "right";
                }
            }
            if (rightDirection == "left")
            {
                if (rightX > Width - pnlRightOptions.Width)
                {
                    rightX -= 2;
                    pnlRightOptions.Location = new Point(rightX, rightY);
                }
            }
            else
            {
                if (rightX < Width)
                {
                    rightX += 2;
                }
                pnlRightOptions.Location = new Point(rightX, rightY);
            }
        }

        private void pbHide_Click(object sender, EventArgs e)
        {
            optionsDirection = "down";
        }

        private void pbExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pbHome_Click(object sender, EventArgs e)
        {
            main.Show();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string status = "";
            Inicial f = new Inicial();
            status = f.VerificaServicio("Hamachi2Svc");
            int i = 1;
            if (status == "Running")
            {
                double resultados = this.GetRowsCount();
                int vueltas = 0;
                double decima = 0;

                if (resultados > 0)
                {
                    if (resultados <= 50)
                    {
                        this.SincronizaNotas((int)resultados);
                        MetroMessageBox.Show(this, "SE ENVIARON " + resultados + " Registros.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    {
                        decima = resultados / 50;
                        vueltas = (int)Math.Ceiling(resultados / 50);

                        while (i <= vueltas)
                        {
                            this.SincronizaNotas(50);
                            i = i + 1;
                        }
                        MetroMessageBox.Show(this, "SE ENVIARON " + resultados + " Registros.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "NO EXISTEN REGISTROS PENDIENTES PARA ENVÍO.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                  
            }
            else
            {
                // RUTINA QUE LEVANTARÁ EL SERVICIO
                MetroMessageBox.Show(this, "EL SERVICIO LOGME IN SE ENCUENTRA DETENIDO. INICIANDO SERVICIO...", "OK", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);

                f.StartService("Hamachi2Svc", 10000);
                System.Threading.Thread.Sleep(5000);
                if(f.VerificaServicio("Hamachi2Svc") == "Running")
                {
                    MetroMessageBox.Show(this, "EL SERVICIO HA SIDO INICIADO CORRECTAMENTE.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }

            }
            
        }
 
        public void SincronizaNotas(int limite)
        {
            string query = "";
            string palabras = "";
            string basereal = "";
            double id = 0;
            int conta = 0;
            int total = 0;
            List<string> list = new List<string>();

            try
            {                

                query = "SELECT * FROM log_track WHERE "+ Inicial.G_LOCAL_CLOUD_ECOMMERCE +" IS NULL ";
                query += "ORDER BY id LIMIT 0," + limite + " ";
  
                Conectar cnn = new Conectar(Inicial.G_SERVIDOR, Inicial.G_CLIENTE_SISTEMA + "log", Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);

                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    total = 0;// GetRowsCount();
                    MySqlDataReader dr = cmd.ExecuteReader();

                    lblinfo.Text = "ENVIANDO REGISTROS... ";
                    lblinfo.Refresh();

                    while (dr.Read())
                    {
                        Thread.Sleep(300);
                        lblinfo.Text = "ENVIANDO " + (conta + 1)  ;
                        lblinfo.Refresh();
                        palabras = dr[2].ToString();
                        basereal = dr[3].ToString();
                        id = Convert.ToDouble(dr[0]);

                        this.TraspasaDatos(palabras, basereal, id, Inicial.G_LOCAL_CLOUD_ECOMMERCE, "UNO");
                        int pos = 0;
                        string nro = "";
                        if(palabras.Contains("sv_documento_cabeza"))
                        {
                            pos = palabras.IndexOf("VALUES");
                            nro = palabras.Substring((pos + 28), 10);
                            list.Add(nro);
                        }
                        conta = conta + 1;
                    }
                    cnn.CloseConnection();
                   
                    if(list.Count > 0)
                    {
                        //this.EnviarEmailCentral(list);
                    }                   
                }
                else
                {
                    MetroMessageBox.Show(this, "NO SE PUDO ESTABLECER UNA CONEXION.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }

            }
            catch(Exception ex)
            {
                MetroMessageBox.Show(this, "ERROR: "+ ex.Message.ToString(), "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
           
        }

        double GetRowsCount()
        {
            string query = "";
            double salida = 0;

            query = "SELECT count(id) FROM eltit_sincroniza.sincronizador_master WHERE server_01 ='0000-00-00'  ";
            query += "ORDER BY id LIMIT 0,"+ Inicial.G_NROREGISTROS +" ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_sincroniza", "root", "123");
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    salida = Convert.ToDouble(dr[0]);
                }
                dr.Close();
            }
            
            cnn.CloseConnection();
            return salida;

        }

        private void TraspasaDatos(string Xconsulta,string Xbasereal, double Xid,string xserver,string Xindice )
        {
            string consulta = "";
            int salida = 0;
            consulta = Xconsulta.Replace("~", "'");

            Conectar cnn = new Conectar(Inicial.G_WEBSERVER, Xbasereal, Inicial.G_WEBUSER, Inicial.G_WEBPASSWORD);

            if(cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(consulta, cnn.connection);

                salida = cmd.ExecuteNonQuery();
                this.ActualizaFecha(Xid, Inicial.G_LOCAL_CLOUD_ECOMMERCE);
            }else
            {
                MetroMessageBox.Show(this, "NO ES POSIBLE CONECTAR CON EL SERVIDOR " + Inicial.G_CENTRAL, "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

            cnn.CloseConnection();

        }
        private void ActualizaFecha(double xid, string xserver)
        {
            string query = "";

            query = "UPDATE log_track SET "+ Inicial.G_LOCAL_CLOUD_ECOMMERCE +" = NOW() ";
            query += "where id = '" + xid + "' ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, Inicial.G_CLIENTE_SISTEMA + "log", Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            else
            {
                //sin conexion
            }
            cnn.CloseConnection();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            lblinfo2.Text = "BUSCANDO PRODUCTOS...";
            lblinfo2.Refresh();
            //this.SincronizaProductos();
            lblinfo2.Text = "BUSCANDO PRECIOS...";
            lblinfo2.Refresh();
            this.SincronizaPrecios();
            lblinfo2.Text = "BUSCANDO OFERTAS Y NOVEDADES";
            lblinfo2.Refresh();
            this.SincronizaOfertasYNovedades();

            MetroMessageBox.Show(this, "TAREA REALIZADA SATISFACTORIAMENTE", "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            lblinfo2.Text = "TAREA FINALIZADA";
            button3.Enabled = false;
        }

        private void SincronizaProductos()
        {

            string maximo = "";
            string query = "";
            Conectar cnn;
            cnn = new Conectar(Inicial.G_WEBSERVER, Inicial.G_WEBDATABASE, Inicial.G_WEBUSER, Inicial.G_WEBPASSWORD);
            Productos p = new Productos();
            maximo = p.GetMaximoFechaProductoLocal();

            query = "SELECT codigobarra, descripcion, codigoseccion, codigodepto, ";
            query += "codigolinea, tipoembalaje, cantidadporembalaje, ";
            query += "publicado, stockpositivo, DATE_FORMAT(fechaactualizacion, '%Y-%m-%d %H:%i:%s') AS fechaamodificacion ";
            query += "From web_r_maestroproductos_fijo_00  ";      
            query += "Where fechaactualizacion > '" + maximo + "' ";

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();
                // PROGRESSBAR
                while(dr.Read())
                {
                    p.InsertaProducto(dr["codigobarra"].ToString(), dr["descripcion"].ToString(), dr["codigoseccion"].ToString(),
                                        dr["codigodepto"].ToString(), dr["codigolinea"].ToString(), dr["tipoembalaje"].ToString(),
                                        dr["cantidadporembalaje"].ToString(), dr["publicado"].ToString(), dr["stockpositivo"].ToString(),
                                        dr["fechaamodificacion"].ToString());



                }
            }
            cnn.CloseConnection();
        }

        private void SincronizaPrecios()
        {

            string maximo = "";
            string query = "";
            double precio1 = 0;
            double precio2 = 0;
            double precio3 = 0;
            Conectar cnn;
            cnn = new Conectar(Inicial.G_WEBSERVER, Inicial.G_WEBDATABASE, Inicial.G_WEBUSER, Inicial.G_WEBPASSWORD);
            Productos p = new Productos();
           // maximo = p.GetMaximoFechaPrecioLocal();

            query = "SELECT local, codigo, codigoprecio,  ";
            query += "preciosistema, preciopuntoventa,forzarupdate, ";
            query += "IF(fechavigencia='0000-00-00', DATE_FORMAT(NOW(), '%Y-%m-%d'), fechavigencia) AS fechavigencia, ";
            query += "preciooferta, DATE_FORMAT(fechaactualizacion, '%Y-%m-%d %H:%i:%s') AS fechaamodificacion ";
            query += "From web_r_maestroproductos_precios_00  ";
            //query += "Where fechaactualizacion > '" + maximo + "' ";
            query += "WHERE (DATE_SUB(CURDATE(),INTERVAL 120 DAY) <= fechaactualizacion OR forzarupdate <> '')  ";
            query += "Order by fechaactualizacion";


            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();
                // PROGRESSBAR
                while (dr.Read())
                {
                    precio1 = Convert.ToDouble(dr["preciosistema"].ToString());
                    precio2 = Convert.ToDouble(dr["preciopuntoventa"].ToString());
                    precio3 = Convert.ToDouble(dr["preciooferta"].ToString());

                    p.InsertaPrecio(dr["local"].ToString(), dr["codigo"].ToString(), dr["codigoprecio"].ToString(),
                        precio1,precio2, dr["fechavigencia"].ToString(),precio3, dr["fechaamodificacion"].ToString(),
                        dr["forzarupdate"].ToString());



                }
            }
            cnn.CloseConnection();
        }
        private void SincronizaOfertasYNovedades()
        {

            string maximo = "";
            string query = "";
            string seccion = "";
            string depto = "";
            string barra = "";
            int publicado = 0;
            string fecha = "";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_WEBSERVER, Inicial.G_WEBDATABASE, Inicial.G_WEBUSER, Inicial.G_WEBPASSWORD);
            Productos p = new Productos();
           // maximo = p.GetMaximoFechaPrecioLocal();

            query = "SELECT codigoseccion, codigodepartamento, codigobarra,publicado, DATE_FORMAT(fechanovedad, '%Y-%m-%d') as fechanovedad  ";
            query += "From web_r_maestroproductos_relacionados  ";
            query += "Where (codigoseccion = '00016' OR codigoseccion = '00017') and YEAR(fechanovedad) >= YEAR(CURRENT_DATE()) ";

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();
                // PROGRESSBAR
                while (dr.Read())
                {
                    seccion = dr["codigoseccion"].ToString();
                    depto =   dr["codigodepartamento"].ToString();
                    barra =   dr["codigobarra"].ToString();
                    publicado = Convert.ToInt32(dr["publicado"]);
                    fecha     = dr["fechanovedad"].ToString();

                    p.InsertaRelacionados(dr["codigoseccion"].ToString(), dr["codigodepartamento"].ToString(), 
                           dr["codigobarra"].ToString(), publicado, fecha);



                }
            }
            cnn.CloseConnection();
        }
        private void button9_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void pnlMain_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
