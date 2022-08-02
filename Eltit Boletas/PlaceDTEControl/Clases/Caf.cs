using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace Eltit.Clases
{
    class Caf
    {
        Conectar cnn;
        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;
        private static readonly log4net.ILog log =
          log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Caf(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }






        public MySqlDataReader GetLitadoFolios(string xLocal, string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select * from dte_fae_caf" + xLocal + " where caf_local = '" + xLocal + "' order by caf_tipo,caf_desde";

            cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }


        public MySqlDataReader BuscaRCOF(string xLocal, string xBase, string xFecha)
        {
            string query = "";

            MySqlDataReader dr = null;
            query = " SELECT * from  dte_boe_rcof" + xLocal + " ";
            query += " Where  fae_recinto = '" + xLocal + "' And fae_fecha = '" + xFecha + "' Limit 0,1  ";

            cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public bool VerificaCaf(string xLocal, string xbase, string xTipo, string xFolio)
        {
            string query = "";
            bool salda = false;
            MySqlDataReader dr;
            query = " SELECT caf_hasta from  dte_fae_caf" + xLocal + " ";
            query += " where  caf_tipo = '" + xTipo + "' And caf_desde <= '" + xFolio + "' and caf_hasta >= '" + xFolio + "'  ";

            cnn = new Conectar(SERVER, xbase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
                if (dr.HasRows == true)
                {
                    salda = true;
                }
                dr.Close();
            }

            cnn.CloseConnection();
            return salda;
        }

        /// <summary>
        /// Funcion Que sirve para Traer los Caf
        /// </summary>
        /// <param name="xLocal"></param>
        /// <param name="xCaja"></param>
        /// <param name="XtIPO"></param>
        /// <returns></returns>
        public MySqlDataReader GetCafByCajaLocal(string xLocal, string xCaja, string XtIPO)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT * FROM  sv_caf" + xLocal + "";
            query += " where local = '"+ xLocal +"' and tipo = '"+ XtIPO +"'  ";
            if(xCaja != "")
            {
                query += " AND caja = '" + xCaja + "' ";
            }

            query += " ORDER BY hasta DESC LIMIT 0,3 ";

            cnn = new Conectar(this.SERVER, FuncionesClass.G_CLIENTE_PREFIJO + "fae" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
               dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public void GrabaCafLocal(string xLocal, string XtIPO, string xDesde, string xHasta,
                                  string xFechaRecepcion, string xUbicacion, string xNombreArchivo,
                                  string xml, string xml2)
        {
            string query = "";

            query = " INSERT INTO  sv_caf" + xLocal + "( ";
            query += " tipo, desde, hasta, fecharecepcion, ";
            query += " ubicacion, nombredelarchivo,xml, local, xml2) VALUES( ";
            query += " '" + XtIPO + "','" + xDesde + "','" + xHasta + "','" + xFechaRecepcion + "',  ";
            query += " '" + xUbicacion + "','" + xNombreArchivo + "','" + xml + "','" + xLocal + "','" + xml + "' ";
            query += " )";

            cnn = new Conectar(this.SERVER, FuncionesClass.G_CLIENTE_PREFIJO + "fae" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }


        }

        public void GrabaCafLocalCaja(string xLocal, string xCaja, string XtIPO, string xDesde, string xHasta,
                                      string xFechaRecepcion, string xUbicacion, string xNombreArchivo,
                                      string xml, string xml2 )
        {
            string query = "";

            query  = " INSERT INTO  sv_caf" + xLocal + "( ";
            query += " tipo, desde, hasta, fecharecepcion, ";
            query += " ubicacion, nombredelarchivo,xml, local, xml2, caja) VALUES( ";
            query += " '"+ XtIPO + "','" + xDesde + "','" + xHasta + "','" + xFechaRecepcion + "',  ";
            query += " '" + xUbicacion + "','" + xNombreArchivo + "','" + xml + "','" + xLocal + "','" + xml + "', ";
            query += " '" + xCaja + "')";
            
            cnn = new Conectar(this.SERVER, FuncionesClass.G_CLIENTE_PREFIJO + "fae" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                 cmd.ExecuteNonQuery();
            }

      
        }
        public MySqlDataReader GetCaFByLocalTodos(string xTipo, string xLocal)
        {
            string query = "";
            MySqlDataReader dr = null;

            query  = " SELECT caj.local,'"+ xTipo + "' As tipo, caj.numero, caj.descripcion , ";
            query += " IFNULL( ";
            query += "  (SELECT MAX(dte.numero) FROM "+ FuncionesClass.G_CLIENTE_PREFIJO +"fae"+ xLocal + ".sv_dte" + xLocal + " AS dte WHERE dte.localdocumento = caj.local ";
            query += "  AND dte.cajadocumento = caj.numero AND dte.tipo = '"+ xTipo +"' GROUP BY dte.tipo ) ";
            query += "  ,0) AS ultimo, ";
            query += "  IFNULL( ";
            query += "  (SELECT dte.hasta FROM eltit_fae"+ xLocal +".sv_caf"+xLocal+" AS dte WHERE dte.caja = caj.numero ORDER BY dte.hasta DESC LIMIT 0,1 ) ";
            query += "  ,0) AS folios, ";
            query += "  IFNULL( ";
            query += "  (SELECT dte.nombredelarchivo FROM eltit_fae"+ xLocal +".sv_caf" + xLocal + " AS dte WHERE dte.caja = caj.numero ORDER BY dte.hasta DESC LIMIT 0,1 ) ";
            query += "  ,0) AS nombreArchivo ";
            query += "  FROM " + FuncionesClass.G_CLIENTE_PREFIJO + "ventas.sv_maestrodecajas AS caj ";
            query += "  WHERE caj.local = '"+ xLocal +"'   ";
            query += "  ORDER BY caj.numero ;";
            query += "  ";


            cnn = new Conectar(this.SERVER, FuncionesClass.G_CLIENTE_PREFIJO + "fae" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;

        }
        public MySqlDataReader GetCafDisponiblesFacturasByLocal(string xTipo, string xLocal)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT  caf.tipo,caf.fecharecepcion, ";
            query += " IFNULL(  (  SELECT MAX(dte.numero) FROM sv_dte" + xLocal + " AS dte ";
            query += " WHERE caf.tipo = dte.tipo LIMIT 0,1   ),  0 )AS ultimo, caf.hasta  ";
            query += " FROM sv_caf" + xLocal + " AS caf  WHERE tipo = '" + xTipo + "' ";
            query += " ORDER BY caf.fecharecepcion DESC LIMIT 0,1 ;  ";
            query += "  ";

            cnn = new Conectar(this.SERVER, FuncionesClass.G_CLIENTE_PREFIJO + "fae" + xLocal, MYSQL_ROOT, MYSQL_PASS); ;
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;

        }
        public MySqlDataReader GetCaFByLocalCaja(string xLocal, string xCaja)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " Select * from sv_caf" +xLocal + " ";
            query += " where caja = '" + xCaja +"' ";
            query += " order by fecharecepcion desc ";

            cnn = new Conectar(this.SERVER, FuncionesClass.G_CLIENTE_PREFIJO + "fae"+ xLocal, MYSQL_ROOT, MYSQL_PASS);
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
