using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;

namespace SchoolManagementAdmin.objetos
{
    class Productos
    {

        public string GetProductosInventario(string xlocal, string xdesde, string xhasta, string xBodega)
        {
            string query = "";
            string salida = "";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "dcoloso_inventario00", "placesoft", "1121");

            query = "SELECT tipo_mov,numero,fecha_emision, art_codigo, art_descripcion FROM local_movimientos_detalle_"+ xlocal +" ";
            query = query + "WHERE fecha_emision BETWEEN '"+ xdesde +"' AND '"+ xhasta +"' GROUP BY art_codigo ORDER BY fecha_emision;";
            query = query + "";
  
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    this.GeneraStock(xlocal, dr["art_codigo"].ToString(), xhasta, xBodega);
                }

                dr.Close();
            }
            
            cnn.CloseConnection();
            return salida;
        }
        private void GeneraStock(string xlocal, string xCodigo, string xhasta, string xBodega)
        {
            double stock = 0;

            stock = GetStock(Inicial.G_RUBRO, "2006", xhasta, xlocal, xBodega, xCodigo);

        }
        public string GetMaximoFechaFotoLocal()
        {
            string query = "";
            string salida = "";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");

            query = "SELECT DATE_FORMAT(fecha2, '%Y-%m-%d %H:%i:%s') FROM r_maestroproductos_fotos_00 ORDER BY fecha2 DESC LIMIT 0,1 ";
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    salida = Convert.ToString(dr[0].ToString());
                }
            }
            cnn.CloseConnection();
            return salida;
        }
        public string GetMaximoFechaProductoLocal()
        {
            string query = "";
            string salida = "";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");

            query = "SELECT DATE_FORMAT(fecha2, '%Y-%m-%d %H:%i:%s') as fecha2 FROM r_maestroproductos_fijo_00 ORDER BY fecha2 DESC LIMIT 0,1 ";
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    salida = Convert.ToString(dr[0].ToString());
                }
            }
            cnn.CloseConnection();
            return salida;
        }
        public string GetMaximoFechaPrecioLocal()
        {
            string query = "";
            string salida = "";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");

            query = "SELECT DATE_FORMAT(fecha2, '%Y-%m-%d %H:%i:%s') as fecha2 FROM r_maestroproductos_precios_00 ORDER BY fecha2 DESC LIMIT 0,1 ";
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    salida = Convert.ToString(dr[0].ToString());
                }
            }
            cnn.CloseConnection();
            return salida;
        }
        public void InsertaImagen(string codigo, PictureBox img, string xfecha )
        {

            MemoryStream ms = new MemoryStream();
            img.Image.Save(ms, ImageFormat.Jpeg);
            byte[] pic_arr = new byte[ms.Length];
            ms.Position = 0;
            ms.Read(pic_arr, 0, pic_arr.Length);
            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");
            string query = "";
            try
            {
                    query = "REPLACE INTO r_maestroproductos_fotos_00(codigobarra, imagen,fechaactualizacion,fecha2) ";
                    query += "VALUES(@codigo, @img, @fecha, @id) ";
                    if (cnn.OpenConnection() == true)
                    {
                        MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                        cmd.Parameters.AddWithValue("@codigo", codigo);
                        cmd.Parameters.AddWithValue("@img", pic_arr);
                        cmd.Parameters.AddWithValue("@fecha", "2015-10-02");
                        cmd.Parameters.AddWithValue("@id", xfecha);

                         cmd.ExecuteNonQuery();

                    }
                cnn.CloseConnection();
            }catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
           
            


        }
        public void InsertaProducto(string xcodigo, string xdescripcion, string xcodigoseccion, string xcodigodepto,
                                    string xcodigolinea, string tipoembalaje, string xcantporembalaje,
                                    string xpublicado, string xstockpositivo, string xfecha)

        {
            string query = "";

            query  = "REPLACE INTO r_maestroproductos_fijo_00(codigobarra, descripcion, codigoseccion, codigodepto, ";
            query += "codigolinea, tipoembalaje, cantidadporembalaje, ";
            query += "publicado, stockpositivo, fecha2) VALUES( ";
            query += " '"+ xcodigo + "','" + xdescripcion + "','" + xcodigoseccion + "','" + xcodigodepto + "', ";
            query += " '" + xcodigolinea + "','" + tipoembalaje + "','" + xcantporembalaje + "', ";
            query += " '" + xpublicado + "','" + xstockpositivo + "','" + xfecha + "' )";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");
            if(cnn.OpenConnection()== true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }


            cnn.CloseConnection();

        }

        public void InsertaPrecio(string xlocal, string xcodigo, string xcodigoprecio,
                                    double xpreciosistema, double xpreciopuntoventa, string xfechavigencia,
                                    double xpreciooferta, string xfecha, string xforzar)

        {
            string query = "";
                        
            query = "REPLACE INTO r_maestroproductos_precios_00(local, codigo, codigoprecio, ";
            query += "preciosistema, preciopuntoventa,fechavigencia, ";
            query += "fecha2,preciooferta) VALUES( ";
            query += " '" + xlocal + "','" + xcodigo + "','" + xcodigoprecio + "', ";
            query += " '" + xpreciosistema + "','" + xpreciopuntoventa + "','" + xfechavigencia + "', ";
           

            if (xforzar != "")
            {
                query += " NOW(),'" + xpreciooferta + "' )";
            }
            else
            {
                query += " '" + xfecha + "','" + xpreciooferta + "' )";
            }

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }


            cnn.CloseConnection();

        }
        public void InsertaRelacionados(string seccion, string depto, string codigo,
                                   double publicado, string fechanovedad)

        {
            string query = "";

            query = "REPLACE INTO r_maestroproductos_relacionados(codigoseccion, codigodepartamento, codigobarra, ";
            query += "publicado, fechanovedad ";
            query += " ) VALUES( ";
            query += " '" + seccion + "','" + depto + "','" + codigo + "', ";
            query += " '" + publicado + "','" + fechanovedad + "' ) ";


            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }


            cnn.CloseConnection();

        }
        public double GetStock(string rubro, string desde, string hasta, string loc, string bodega, string codigo)
        {
            string añode = "2006";
            string sql = "";
            double salida = 0;

            sql = "SELECT  ifnull((SELECT sum(if (mt.operacion ='+',mov.art_unidades,mov.art_unidades * -1)) ";
            sql = sql + "FROM local_movimientos_detalle_" + loc + " as mov," + Inicial.G_CLIENTE_SISTEMA + "mantencion";
            sql = sql + ".mae_tipos_de_documentos as mt WHERE mt.codigo=mov.tipo_mov and mov.art_codigo=mps.codigo_articulo  ";
            sql = sql + "and mov.almacen_destino = '" + bodega + "' AND mov.fecha_emision >= '2006-01-01' AND mov.fecha_emision <='" + hasta + "'  ";
            sql = sql + "AND mov.tipo_mov <> 'NPE' group BY mov.art_codigo ORDER BY mov.art_codigo),'0') as saldo ";
            sql = sql + "FROM mae_articulos_stock_" + rubro + " as mps WHERE mps.local = '" + loc + "' AND mps.codigo_almacen = '" + bodega + "' ";
            sql = sql + "And mps.codigo_articulo ='" + codigo + "' Order by  mps.codigo_articulo asc  limit 0,1 ";
            sql = sql + "";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_CENTRAL, Inicial.G_CLIENTE_SISTEMA + "inventario"+ rubro, "placesoft", "1121");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(sql, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    salida = Convert.ToDouble(dr["saldo"].ToString());
                }
                if(salida < 0)
                {
                    salida = 0;
                }
                this.setStockLocal(codigo,salida);
            }
            
            cnn.CloseConnection();
            return salida;
        }
        public double GetPrecioSistema(string codigo, string rubro)
        {
            
            string sql = "";
            double salida = 0;
            sql = "SELECT mpp.preciopuntoventa FROM r_maestroproductos_precios_" + rubro + " as mpp ";            
            sql = sql + "WHERE codigo = '" + codigo + "' and codigoprecio = '01' ";
            
            Conectar cnn;
            cnn = new Conectar(Inicial.G_CENTRAL, "eltit_gestion00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(sql, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    salida = Convert.ToDouble(dr["preciopuntoventa"].ToString());
                }
               // this.setStockLocal(codigo, salida);
            }

            cnn.CloseConnection();
            return salida;

        }
        public void setPrecioLocal(string xcodigo, double xprecio)
        {

            string query = "";

            query = "UPDATE r_maestroproductos_precios_00 SET ";
            query += "preciopuntoventa = '" + xprecio + "' ";
            query += "WHERE codigo = '" + xcodigo + "' and codigoprecio = '01' ";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }


            cnn.CloseConnection();
        }
        private void setStockLocal(string xcodigo, double xstock)
        {
            
                string query = "";

                query = "UPDATE mae_articulos_00 SET ";
                query += "stock_temporal = '" + xstock + "' ";
                query += " WHERE codigobarra = '"+ xcodigo +"' ";

                Conectar cnn;
                cnn = new Conectar(Inicial.G_SERVIDOR, Inicial.G_CLIENTE_SISTEMA + "inventario00", "placesoft", "1121");
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    cmd.ExecuteNonQuery();

                //Sincroniza snc = new Sincroniza(Funciones.G_SERVIDOR);
                //snc.GrabaSincronizador(query, Funciones.G_CLIENTE_SISTEMA + "inventario00");

            }


            cnn.CloseConnection();
        }
        public double getStockLocal(string xcodigo)
        {
            double salida = 0;
            string query = "";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion00", "root", "123");

            query = "SELECT precio2 FROM r_maestroproductos_fijo_00 WHERE CODIGOBARRA = '" + xcodigo + "' LIMIT 0,1 ";
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    salida = Convert.ToDouble(dr[0].ToString());
                }
            }
            cnn.CloseConnection();
            return salida;

        }
        public string getDescripcionByRubro(string xrubro, string xcodigo)
        {
            string salida = "";
            string query = "";

            Conectar cnn;
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_gestion" + xrubro, "root", "123");

            query = "SELECT descripcion FROM r_maestroproductos_fijo_"+xrubro  +" WHERE codigobarra = '" + xcodigo + "' LIMIT 0,1 ";
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    salida = Convert.ToString(dr[0].ToString());
                }
            }
            cnn.CloseConnection();
            return salida;

        }

    }
}
