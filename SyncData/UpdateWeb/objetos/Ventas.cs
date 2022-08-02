using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.Data;

namespace SchoolManagementAdmin.objetos
{
    class Ventas
    {
        Conectar cnn;


        public DataTable GetDetalleByNumeroCaja(string xlocal, string xnumero, string xfecha)
        {
            string query = "";
            string[] fecha = null;
            string fechafinal = "";
            fecha = xfecha.Split('-');
            fechafinal = fecha[2] + "-" + fecha[1] + "-" + fecha[0];
            DataTable dt = new DataTable();
            query = "Select dd.codigo, dd.descripcion, dd.cantidad, dd.precio, dd.total  From ";
            query += "sv_documento_detalle_00 as dd ";
            query += "WHERE dd.local = '" + xlocal + "' And tipo = 'PV'  And dd.caja = '" + Inicial.G_CAJA + "' ";
            query += "And dd.numero = '" + xnumero + "' ";
            query += "And dd.fecha  = '" + fechafinal + "' ";

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

        public void GrabaVenta(string xlocal, string xrut, string xtipo, string xsucursal, 
                                string xcaja, string xvendedor)
        {
            string query = "";
            double subtotal = 0;
            double neto = 0;
            double iva = 0;
            double total = 0;

            string numero = this.UltimaNumero(xlocal, xtipo, xcaja);
            query  = "SELECT local, linea, now() as fecha, codigo, descripcion, cantidad, precio, descuento, total FROM ";
		    query += " sv_rollo_00 WHERE caja = '" + xcaja + "' ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if(cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr2 = cmd.ExecuteReader();
               
                while (dr2.Read())
                {
                    this.GrabaDetalle(xlocal, xtipo, numero, dr2["linea"].ToString(),
                                     dr2["fecha"].ToString(), xrut, xsucursal,
                                     dr2["codigo"].ToString(), dr2["descripcion"].ToString(),
                                     dr2["cantidad"].ToString(), dr2["precio"].ToString(), 0,
                                     0, dr2["total"].ToString(), xvendedor, "GENERADO POR PREVENTA INTERNET",
                                     xcaja);

                    subtotal = subtotal + Convert.ToDouble(dr2["total"].ToString());
                }
                    neto = Math.Round(subtotal / 1.19);			      
			        iva = subtotal - neto;
			        total = neto + iva;
                dr2.Close();
                           
                this.GrabaCabeza(xlocal, xtipo, numero, xrut, xsucursal, xvendedor, subtotal, neto, iva, total, "Preventa", xcaja);
              
            }
            
            cnn.CloseConnection();

        }
        // GRABA DETALLES DEL DOCUMENTO
        private void GrabaDetalle(string xlocal, string xtipo, string xnumero, string xlinea,
                                  string xfecha, string xrut, string xsucursal,
                                  string xcodigo, string xdescripcion, string xcantidad,
                                  string xprecio, double xdescuento, double xdescuento2,
                                  string xtotal, string xvendedor, string xglosa, string xcaja)
        {
            string query = "";
            query  = "INSERT IGNORE INTO eltit_ventas00.sv_documento_detalle_00  ";
            query += "(local, tipo, numero, linea, fecha, vencimiento, rut, sucursal, codigo, descripcion, cantidad , ";
            query += "unidades, precio, descuento, descuento2, total, vendedor, glosa, horaventas, caja, pcosto,bodega "; 
            query += ") VALUES( ";
            query += " '"+ xlocal + "', '" + xtipo + "', LPAD('"+ xnumero + "',10,'0'), LPAD('"+ xlinea +  "',3,'0'),  ";
            query += " NOW(),NOW(), '"+ xrut + "', '"+ xsucursal + "', '"+ xcodigo +  "',  ";
            query += " '"+ xdescripcion + "', '"+ xcantidad + "', '"+ xcantidad + "', " + xprecio + ",  ";
            query += " " + xdescuento + ", '"+ xdescuento2 + "', " + xtotal + ", '" + xvendedor + "', '"+ xglosa +  "', ";
            query += " CURTIME(), '" + xcaja + "','1', '00' )  ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
                // SINCRONIZADATOS
                //Sincroniza snc = new Sincroniza("");
                ////snc.GrabaSincronizador(query, "eltit_ventas00");
                //this.GrabaObservaciones("00", "PV", xnumero, xlinea, xrut, xcaja,xcodigo);
            }
            cnn.CloseConnection();
        }
        private void GrabaCabeza(string xlocal, string xtipo,string xnumero,string xrut, 
                                string xsucursal,string xvendedor,double xsubtotal, 
                                double xneto,double xiva,double xtotal,string xusuario,string xcaja)
        {

            string query = "";
            query = "INSERT IGNORE INTO eltit_ventas00.sv_documento_cabeza_00 ";
            query += "(local, tipo, numero, fecha, vencimiento, rut, sucursal, ";
            query += "cajera, subtotal, neto, iva, total, vendedor, horaventas, glosaguia, caja, foliosii, formapago,condicionesdepago) VALUES ( ";
            query += " '" + xlocal + "', '" + xtipo + "', LPAD('" + xnumero + "',10,'0'), NOW(), NOW(), '" + xrut + "', ";
            query += " '" + xsucursal + "' ,'" + xvendedor + "', " + xsubtotal + ", " + xneto + ",TRUNCATE(" + xiva + ",0), TRUNCATE(" + xtotal + ",0), ";
            query += " '" + xvendedor + "' , CURTIME() , '" + xusuario + "',  '" + xcaja + "', LPAD('" + xnumero + "',10,'0') , '1','NO') ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            if(cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);

                cmd.ExecuteNonQuery();
                // Sincronizadatostas00
                //Sincroniza snc = new Sincroniza();
                //snc.GrabaSincronizador(query, "eltit_ventas00");
            }

            cnn.CloseConnection();

        }
        private void GrabaObservaciones(string xlocal, string xtipo,string xnumero, string xlinea, string xrut, 
                                        string xcaja, string xcodigo)
        {
            Rollo rollo = new Rollo(xcodigo);
            string query = "";
            query = "INSERT IGNORE INTO eltit_ventas00.sv_documento_observaciones_00( ";
            query += "local, tipo,   numero, ";
            query += "linea, fecha, rut, ";
            query += "caja,  codigo, observaciones) VALUES( ";
            query += " '"+ xlocal + "','" + xtipo + "','" + xnumero + "', ";
            query += " '" + xlinea + "',NOW(), '" + xrut + "', ";
            query += " '"+ xcaja + "','" + xcodigo + "','" + rollo.observacion + "') ";

            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");

            try
            {
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    cmd.ExecuteNonQuery();
                    //Sincroniza snc = new Sincroniza();
                    //snc.GrabaSincronizador(query, "eltit_ventas00");

                    cnn.CloseConnection();
                }
            }
            catch(Exception ex)
            {                
                cnn.CloseConnection();
            }
            

        }
        private string UltimaNumero(string xlocal, string xtipo, string xcaja)
        {
            string query = "";
            string salida = "";
            query  = "Select  MAX(numero) + 1  From ";
            query += " sv_documento_cabeza_00 ";
            query += "WHERE local = '" + xlocal + "' And tipo = '" + xtipo + "'  And caja = '"+ xcaja + "' ";
            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");
            if(cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                MySqlDataReader dr = cmd.ExecuteReader();

                while(dr.Read())
                {
                    salida = Convert.ToString(dr[0]);
                }
               
            }
            if(salida == "")
            {
                salida = "1";
            }
            cnn.CloseConnection();
            return salida.PadLeft(10,'0');


        }

        public MySqlCommand GetNotasByVendedor(string xlocal, string xvendedor, string xcaja,string xdia,string xmes, string xaño)
        {
            string query = "";
            MySqlCommand salida = null;

            query = "Select dc.numero,dc.fecha, cl.nombre, dc.total  From ";
            query += "sv_documento_cabeza_00 as dc ";
            query += "INNER JOIN eltit_ventas.sv_maestroclientes AS cl ";
            query += "ON(dc.rut = cl.rut) And dc.sucursal = cl.sucursal ";
            query += "WHERE dc.local = '" + xlocal + "' And tipo = 'PV'  And dc.caja = '" + xcaja + "' ";
            query += "And dc.vendedor = LPAD('" + xvendedor + "',10,'0') ";
            query += "And MONTH(dc.fecha) = '" + xmes + "' AND YEAR(dc.fecha) = '" + xaño + "' ";
            query += "AND  DAY(dc.fecha) = '" + xdia + "' ";
            Conectar cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);                
                salida = cmd;
            }
            cnn.CloseConnection();
            return salida;


        }
        public MySqlDataReader GetNotasByNumeros(string xlocal, string xvendedor, string xcaja,string xCondicion)
        {
            string query = "";
            MySqlDataReader salida = null;

            query = "Select dc.numero,dc.fecha, cl.nombre, dc.total  From ";
            query += "sv_documento_cabeza_00 as dc ";
            query += "INNER JOIN eltit_ventas.sv_maestroclientes AS cl ";
            query += "ON(dc.rut = cl.rut) And dc.sucursal = cl.sucursal ";
            query += "WHERE dc.local = '" + xlocal + "' And tipo = 'PV'  And dc.caja = '" + xcaja + "' ";
            query += "And dc.vendedor = LPAD('" + xvendedor + "',10,'0') ";
            query += "and "+ xCondicion +" ";
       
            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_ventas00", "root", "123");
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                salida = cmd.ExecuteReader();
            }
         
            return salida;

        }
        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }
    }
}
