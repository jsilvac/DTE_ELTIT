using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data;
using MySql;
using MySql.Data.MySqlClient;
using System.Windows.Forms;
using System.Data;

namespace SchoolManagementAdmin.objetos
{
   public class Rollo
    {
        
        private int xlinea = 0;
        private double xcantidad = 0;
        private string xcodigo = "";
        private string xdescripcion = "";
        private double xdescuento = 0;
        private double xprecio = 0;
        private double xtotal = 0;
        private string xvendedor = "";
        private string xfecha = "";
        private string xhora = "";
        private string xobservacion = "";
     

        public Rollo(int xlinea, double xcantidad, string xcodigo, string xdescripcion, double xdescuento, double xprecio, double xtotal, string xvendedor, string xfecha, string xhora)
        {
            
            this.xlinea = xlinea;
            this.xcantidad = xcantidad;
            this.xcodigo = xcodigo;
            this.xdescripcion = xdescripcion;
            this.xdescuento = xdescuento;
            this.xprecio = xprecio;
            this.xtotal = xtotal;
            this.xvendedor = xvendedor;
            this.xfecha = xfecha;
            this.xhora = xhora;
        }

        public Rollo()
        {

        }
        public int ExisteItemEnRollo(string xcodigo)
        {
            string query = "";
            int salida = 0;

            query = "Select COUNT(*) from sv_rollo_00 where codigo = LPAD('"+ xcodigo +"','13',0) ";
            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);

                salida = Convert.ToInt32(cmd.ExecuteScalar());


            }
            cnn.CloseConnection();
            return salida;
        }
        public Rollo(string xcodigo)
        {
            string query = "";
 
            query  = "SELECT ro.cantidad,ro.codigo,ro.descripcion,ro.precio,ro.total, ob.observaciones ";
            query += "FROM sv_rollo_00 AS ro ";
            query += "INNER JOIN sv_rollo_observaciones_00 AS ob ";
            query += "ON(ro.codigo = ob.codigo)";
            query += "Where ro.codigo = LPAD('"+ xcodigo + "',13,'0') Limit 0,1 ";
            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    this.cantidad =  Convert.ToDouble(dr[0].ToString());
                    this.codigo = Convert.ToString(dr[1].ToString());
                    this.descripcion = Convert.ToString(dr[2].ToString());
                    this.precio = Convert.ToDouble(dr[3].ToString());
                    this.total = Convert.ToDouble(dr[4].ToString());
                    this.observacion = Convert.ToString(dr[5].ToString());
                }
                cnn.CloseConnection();
            }

           
        }
        public void GrabaItemEnRollo()
        {
            string query = "";
            string desc = "";


            if (this.descripcion.Length > 50)
            {
                desc = this.descripcion.Substring(0, 49);
            }
            else
            {
                desc = this.descripcion;
            
            }
            this.EliminaItemRollo(Inicial.G_CAJA, this.codigo);
            Inicial f = new Inicial();
            f.CargaConfiguracion();
            linea = this.getLinea() + 1;
            query = "REPLACE INTO sv_rollo_00( ";
            query += "local, caja, linea, cantidad, ";
            query += "codigo, descripcion, descuento, precio, ";
            query += "total, vendedor, fecha, hora ";
            query += ")Values(";
            query += "'"+ f.local +"', '"+ f.caja +"', '"+ linea +"', '"+ this.cantidad +"', ";
            query += "'"+ this.codigo +"', '"+ desc +"', '"+ this.descuento +"', '"+ this.precio +"', ";
            query += "'"+ this.total +"', '"+ this.vendedor + "',NOW(), CURRENT_TIME() ) ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }

            cnn.CloseConnection();

        }
        public void GrabaObservacion(string xcodigo, string xcaja, string xobservacion)
        {
            string query = "";

            query = "REPLACE INTO sv_rollo_observaciones_00(codigo, fecha, caja, observaciones)Values( ";
            query += " '"+ xcodigo + "', NOW(), '"+ xcaja +"', '"+ xobservacion +"' )   ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");
            
            if(cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query,cnn.connection);

                cmd.ExecuteNonQuery();
                
            }

            cnn.CloseConnection();

        }
        public void EliminaRolloObservaciones()
        {
            string query = "DELETE FROM sv_rollo_observaciones_00 ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            cnn.CloseConnection();
        }
        private int getLinea()
        {
            string query = "";
            int salida = 0;
            query = "Select COUNT(codigo) from sv_rollo_00 ";
            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    salida = Convert.ToInt32(dr[0]);

                }
            }
            cnn.CloseConnection();
            return salida;


        }
    public DataTable GetRollo()
        {
            DataTable dt = new DataTable();
            string query = "SELECT fo.imagen,MID( ro.codigo,9,5) AS codigo, CAST(concat('(',ro.cantidad,') ',ro.descripcion )AS CHAR),  ro.total ";
            query += "FROM sv_rollo_00 AS ro ";          
            query += "left JOIN eltit_gestion00.r_maestroproductos_fotos_00 AS fo ";
            query += "ON (ro.codigo = fo.codigobarra) ";
            query += "Where ro.atencion = 'ACTIVA' Order by ro.linea DESC ";
            //textBox1.Text = query;
            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);

                MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                da.Fill(dt);
                da.Dispose();

            }
            cnn.CloseConnection();
            return dt;

        }
        //public DataTable GetRollo()
        //{
        //    DataTable dt = new DataTable();
        //    string query = "SELECT fo.imagen,MID( ro.codigo,9,5) AS codigo, CAST(concat('(',ro.cantidad,') ',ro.descripcion )AS CHAR),  ro.total ";
        //    query += "FROM sv_rollo_00 AS ro ";
        //    query += "INNER JOIN eltit_gestion00.r_maestroproductos_precios_00 AS mpp ";
        //    query += "ON(ro.codigo = mpp.codigo) AND mpp.codigoprecio = '02'";
        //    query += "left JOIN eltit_gestion00.r_maestroproductos_fotos_00 AS fo ";
        //    query += "ON (ro.codigo = fo.codigobarra) ";
        //    query += "Where ro.atencion = 'ACTIVA' Order by ro.linea DESC ";
        //    //textBox1.Text = query;
        //    Conectar cnn = new Conectar(Funciones.G_SERVIDOR, "eltit_ventas00", "root", "123");
        //    if (cnn.OpenConnection() == true)
        //    {
        //        MySqlCommand cmd = new MySqlCommand(query, cnn.connection);

        //        MySqlDataAdapter da = new MySqlDataAdapter(cmd);
        //        da.Fill(dt);
        //        da.Dispose();

        //    }
        //    cnn.CloseConnection();
        //    return dt;

        //}
        public void EliminaRollo(string xcaja)
        {
            string query = "DELETE FROM sv_rollo_00 ";
            query += "WHERE caja = '" + xcaja + "' and atencion = 'ACTIVA' ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query,cnn.connection);
                cmd.ExecuteNonQuery();
            }
            cnn.CloseConnection();
        }

        public void EliminaItemRollo(string xcaja, string xcodigo)
        {
            string query = "DELETE FROM sv_rollo_00 ";
            query += "WHERE caja = '" + xcaja + "' and codigo = LPAD('" + xcodigo + "',13,'0') ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            cnn.CloseConnection();
            this.EliminaObservacionRollo(xcaja, xcodigo);

        }
        public void EliminaObservacionRollo(string xcaja, string xcodigo)
        {
            string query = "DELETE FROM sv_rollo_observaciones_00 ";
            query += "WHERE caja = '" + xcaja + "' and codigo = LPAD('" + xcodigo + "',13,'0') ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            cnn.CloseConnection();

        }

        public int linea
        {
            get
            {
                return xlinea;
            }

            set
            {
                xlinea = value;
            }
        }

        public double cantidad
        {
            get
            {
                return xcantidad;
            }

            set
            {
                xcantidad = value;
            }
        }

        public string codigo
        {
            get
            {
                return xcodigo;
            }

            set
            {
                xcodigo = value;
            }
        }

        public string descripcion
        {
            get
            {
                return xdescripcion;
            }

            set
            {
                xdescripcion = value;
            }
        }

        public double descuento
        {
            get
            {
                return xdescuento;
            }

            set
            {
                xdescuento = value;
            }
        }

        public double precio
        {
            get
            {
                return xprecio;
            }

            set
            {
                xprecio = value;
            }
        }

        public double total
        {
            get
            {
                return xtotal;
            }

            set
            {
                xtotal = value;
            }
        }

        public string vendedor
        {
            get
            {
                return xvendedor;
            }

            set
            {
                xvendedor = value;
            }
        }

        public string fecha
        {
            get
            {
                return xfecha;
            }

            set
            {
                xfecha = value;
            }
        }

        public string hora
        {
            get
            {
                return xhora;
            }

            set
            {
                xhora = value;
            }
        }

        public string observacion
        {
            get
            {
                return xobservacion;
            }

            set
            {
                xobservacion = value;
            }
        }


    }
}
