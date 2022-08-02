using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eltit.Clases
{
    class Documentos__
    {
        Conectar cnn;
        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;
        private static readonly log4net.ILog log =
          log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public Documentos__(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
         }

        public void ActualizaFolioSII(string xLocal,string xBase, string xnumeroInterno,string xTipoInterno, string xFecha, string xCaja, string xFolioSII)
        {
            string query = "";

            query = "  UPDATE eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc ";
            query += "  SET dc.foliosii = '"+ xFolioSII + "' ";
            query += "  where dc.local = '"+ xLocal + "' AND dc.tipo ='" + xTipoInterno +"' AND dc.numero = '"+ xnumeroInterno  +"' AND dc.fecha = '"+ xFecha +"'  ";
            query += "  and dc.caja = '"+ xCaja +"' ";

            cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }



        }
        public MySqlDataReader GetDocumentosGuasByLocalNroInternoCajaDesdeHasta(string xLocal, string xCaja, string xTipoInterno,
                                                          string xFolioDesde, string xFolioHasta, string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dc.tipo, dc.numero, dc.caja, dc.fecha, dc.total, IFNULL(dte.xml, '0') AS xml, dc.foliosii, ";
            query += " dc.indicador_traslado,dc.rut ";
            query += " FROM eltit_ventas" + xLocal + ".sv_guias_cabeza_" + xLocal + " AS dc ";
            query += " LEFT JOIN eltit_fae" + xLocal + ".sv_dte" + xLocal + " AS dte ON(dc.numero = dte.numerodocumento) AND dc.caja = dte.cajadocumento ";
            query += " AND dc.fecha = dte.fecha AND dc.local = dte.localdocumento ";

            if (xTipoInterno == "NC")
            {
                query += " WHERE dc.caja = '" + xCaja + "' AND (dc.tipo = 'NB' OR dc.tipo = 'NF') AND ";
            }
            else
            {
                query += " WHERE dc.caja = '" + xCaja + "' AND dc.tipo = '" + xTipoInterno + "' AND ";
            }

            query += " (dc.numero >= '" + xFolioDesde + "' AND dc.numero <= '" + xFolioHasta + "') and dc.fecha >= '2020-01-01' ORDER BY dc.numero ASC  ";



            cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }
            return dr;
        }
        public MySqlDataReader GetDocumentosCabezaByLocalNroInternoCajaDesdeHasta(string xLocal, string xCaja,string xTipoInterno,
                                                             string xFolioDesde, string xFolioHasta, string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dc.tipo, dc.numero, dc.caja, dc.fecha, dc.total, IFNULL(dte.xml, '0') AS xml, dc.foliosii,'0' as indicador_traslado,dc.rut ";
            query += " FROM eltit_ventas"+ xLocal +".sv_documento_cabeza_"+ xLocal +" AS dc ";
            query += " LEFT JOIN eltit_fae"+ xLocal +".sv_dte"+ xLocal +" AS dte ON(dc.numero = dte.numerodocumento) AND dc.caja = dte.cajadocumento ";
            query += " AND dc.fecha = dte.fecha AND dc.local = dte.localdocumento ";

            if(xTipoInterno == "NC")
            {
                query += " WHERE dc.caja = '" + xCaja + "' AND (dc.tipo = 'NB' OR dc.tipo = 'NF') AND ";
            }
            else
            {
                query += " WHERE dc.caja = '" + xCaja + "' AND dc.tipo = '" + xTipoInterno + "' AND ";
            }
          
            query += " (dc.numero >= '" + xFolioDesde + "' AND dc.numero <= '"+ xFolioHasta + "') and dc.fecha >= '2020-01-01' ORDER BY dc.numero ASC  ";

            cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }
            return dr;
        }

        public MySqlDataReader GetDoucumentosDetalleByTipoCajaNroOnternoFechaLocal(string xLocal, string xTipo, string xNro, string xCaja, string xFecha, string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dd.tipo, dd.numero, dd.caja, dd.fecha,dc.vencimiento, dd.linea, dd.rut, ";
            query += " dd.codigo, dd.descripcion,  dd.cantidad, dd.precio, dd.descuento, dd.descuento,dd.total, ";
            query += " dd.vendedor,  dd.impuesto, dd.impuesto,dd.porcentajeimpuesto,dd.descuentopesos, ";
            query += " dd.tipo, dd.numero ,dc.neto, dc.iva, dc.exento, ";
            query += " dc.total, dd.tipo, dd.numero, dd.fecha ,dc.indicador_traslado  ";
            query += " FROM " + FuncionesClass.G_CLIENTE_PREFIJO + "ventas" + xLocal + ".sv_documento_detalle_" + xLocal + " AS dd ";
            query += " INNER JOIN  " + FuncionesClass.G_CLIENTE_PREFIJO + "ventas"+ xLocal +".sv_documento_cabeza_"+ xLocal +" AS dc  ON(dd.local = dc.local) AND dd.tipo = dc.tipo ";
            query += " AND dd.numero = dc.numero AND dd.caja = dc.caja AND dd.rut = dc.rut AND dd.fecha = dc.fecha ";
            query += " WHERE dd.local ='"+xLocal+"' AND dd.tipo ='"+xTipo+"' AND dd.numero ='"+ xNro +"' AND dd.caja = '"+ xCaja +"' ";
            query += " AND dd.fecha = '"+ xFecha +"' ORDER BY dd.linea ";
            query += " ";
            
            cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
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
