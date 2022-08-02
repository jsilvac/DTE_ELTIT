using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlaceSoft.Eltit.Class
{
  public class Empresas
    {

        Conectar cnn;
        private string MYSQL_SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;

        public Empresas(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.MYSQL_SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }
        public void GrabaDteRecibidosProveedores(string xRutEmisor, string xRazonSocial, string xTipoDTE, string xFolioDTE,
                                    string xFechaEmision, string xCorreo, double xtotal, string XML, string xNombreArchivo,
                                    string xcodigoFAE, int xCant, string xFechaRecepcion)
        {
            string query = "";

            query = "REPLACE INTO sv_dte"+ xcodigoFAE +"_recibidos( ";
            query += " tipo, numero, fecha, ";
            query += " rut, nombre, fecharecepcion, ";
            query += " nombrearchivo,  monto, correo_proveedor, ";
            query += " xml, rut2   ) VALUES( ";
            query += " '" + xTipoDTE + "', '" + xFolioDTE + "','" + xFechaEmision + "', ";
            query += " '" + xRutEmisor + "', '" + xRazonSocial + "','" + xFechaRecepcion + "', ";
            query += " '" + xNombreArchivo + "', '" + xtotal + "','" + xCorreo + "' , ";
            query += " '" + XML + "','" + xRutEmisor + "') ";


            Conectar cnn;
            cnn = new Conectar(this.MYSQL_SERVER, "eltit_fae" + xcodigoFAE, this.MYSQL_ROOT, this.MYSQL_SERVER);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                /************** GRABA EN SINCRONIZADOR ************/
                //Sincroniza sync = new Sincroniza(this, "placesof");
                //sync.GrabaSincronizador(query, this.Prefijo + "_dte_" + this.Rut_database, "0");
            }


            cnn.CloseConnection();
        }
        public void GrabaDteRecepcion(string xRutEmisor, string xtipoRecibo,string xCorreo, string XML, string xNombreArchivo,
                                    string xcodigoFAE, int xCant , string xFechaRecepcion)
        {
            string query = "";

            query = "INSERT IGNORE INTO sv_recepcion_dte"+ xcodigoFAE +"( ";
            query += " correo, archivo, fecha_recepcion, ";
            query += " archivo_recepcion, archivo_respuesta, fecha_respuesta, ";
            query += " rut,  tipo, documentos";
            query += ") VALUES( ";
            query += " '" + xCorreo + "', '" + xNombreArchivo + "','" + xFechaRecepcion + "', ";
            query += " '" + XML + "', '','', ";
            query += " '" + xRutEmisor + "', '" + xtipoRecibo + "','" + xCant + "' ) ";
            query += " ";


            Conectar cnn;
            cnn = new Conectar(this.MYSQL_SERVER, "eltit_fae"+ xcodigoFAE,this.MYSQL_ROOT, this.MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                /************** GRABA EN SINCRONIZADOR ************/
                //Sincroniza sync = new Sincroniza(this, "placesof");
                //sync.GrabaSincronizador(query, this.Prefijo + "_dte_" + this.Rut_database, "0");
            }


            cnn.CloseConnection();
        }
        public MySqlDataReader GetEmpresasBoleta()
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte  where activo = 1 and   ";
            //query += " (rut = '0775877707' OR ";  // SUPERMERCADO PUCON ORIENTE LTDA.
            //query += "  rut = '0775753404' OR ";  // SUPERMERCADO ELTIT LTDA.
            //query += "  rut = '0763539105' OR ";  // SUPERMERCADO SANTA VICTORIA LTDA
            //query += "  rut = '0775753005' OR ";  // FERRETERIA ELTIT LTDA.
            //query += "  RUT = '077576550K' OR ";  // VIEJO ALMACEN LTDA
            //query += "  rut = '0775765305' OR ";  // ALMACENES ELTIT
            //query += "  rut = '0775773308' OR ";  // FARMACIA PUCON LTDA
            //query += "  rut = '0775779101' OR ";  // COMERCIAL GALPON VERDE LTDA.
            //query += "  rut = '0775810408' OR ";  // IMPORT. DE VESTUARIO LOS TRAPOS LTDA.
            //query += "  rut = '0775877804' OR ";  // GIMNASIO NEW LIFE LTDA
            //query += "  rut = '0775834803' ";     // TURISMO Y HOTELERA DEL VOLCAN LTDA
            //query += "  ) ";
            query += "  envia_boletas_sii = 1 ";
            query += "Order by codigo_contable ";


            cnn = new Conectar(this.MYSQL_SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GEtEmpresasContaByRut(string xRut)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From maestroempresas  where rut = '"+ xRut +"' and empresafae <> '' and certificado <> ''  ";
            query += " ";


            cnn = new Conectar(this.MYSQL_SERVER, "eltit_conta", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        public MySqlDataReader GEtEmpresasConta()
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From maestroempresas  where empresafae <> '' and certificado <> ''  ";
            query += "and servermail <> '' Order by codigoempresa ";


            cnn = new Conectar(this.MYSQL_SERVER, "eltit_conta", MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }


        public MySqlDataReader GetDatoEmpresaByRut(string xCliente, string xRut)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte ";
            query += "Where prefijo = '" + xCliente + "' and rut = '" + xRut + "' ";

            cnn = new Conectar(this.MYSQL_SERVER, "eltit_dte_manager", MYSQL_ROOT, MYSQL_PASS);
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
