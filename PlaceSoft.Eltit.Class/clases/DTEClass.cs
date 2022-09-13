using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlaceSoft.Eltit.Class.clases
{
   public class DTEClass
    {
        Conectar cnn;
        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;
        private string _CLIENTE = "eltit_";
        private static readonly log4net.ILog log =
          log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static object Engine { get; set; }

        public DTEClass(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }

        public string GetXMLFacturas(string xlocal, string xTipo_Sii, string xFolio, string xFecha)
        {
            string salida = "0";

            string query = "";
            MySqlDataReader dr;
            query = " SELECT * from sv_dte" + xlocal;
            query += " WHERE  localdocumento = '" + xlocal + "' and tipo='" + xTipo_Sii + "' and fecha='" + xFecha + "' and numero='" + xFolio + "' ";

            try
            {
                Conectar cnn = new Conectar(SERVER, "eltit_fae"+ xlocal, MYSQL_ROOT, MYSQL_PASS);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        if (dr.Read())
                        {
                            salida = dr["xml"].ToString();
                        }
                    }
                    dr.Close();
                }

                cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                log.Error("Exepcion", ex);
                //MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }

            return salida;
        }

        public string GetXMLBoletas(string xlocal, string xTipo_Sii, string xFolio, string xFecha)
        {
            string salida = "0";

            string query = "";
            MySqlDataReader dr;
            query = " SELECT * from sv_dte_boe" + xlocal;
            query += " WHERE  localdocumento = '" + xlocal + "' and tipo='" + xTipo_Sii + "' and fecha='" + xFecha + "' and numero='" + xFolio + "' ";

            try
            {
                Conectar cnn = new Conectar(SERVER, "eltit_fae" + xlocal, MYSQL_ROOT, MYSQL_PASS);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        if (dr.Read())
                        {
                            salida = dr["xml"].ToString();
                        }
                    }
                    dr.Close();
                }

                cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                log.Error("Exepcion", ex);
                //MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }

            return salida;
        }

        public string ExisteDTELocalNrocajaFecha(string xRutEmpresa, string xLocal, string xCodempresa,string xCaja, 
                                            string xNro, string xFecha, string xTipo)
        {
            string query = "";
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();
            MySqlDataReader dr = null;
            string salida = "";
            query = " SELECT fae_status_sii ";
            query += " FROM  dte_boe_local" + xCodempresa + " ";
            query += " WHERE  fae_tipo = '" + xTipo + "' and fae_folio = '" + Convert.ToInt32(xNro) + "' ";
            query += " and fae_cajadocumento = '" + xCaja + "' and fae_fecha = '" + xFecha + "' and fae_recinto = '" + xLocal + "' ";
            query += "  Limit 0,1 ";

            cnn = new Conectar(SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();

                if (dr.HasRows == true)
                {
                    if(dr.Read())
                    {
                        salida = dr[0].ToString();
                    }
                }
            }

            dr.Close();
            cnn.CloseConnection();

            return salida;
        }
        public void EliminaDTELocalTrack(string xRutEmpresa, string xCodempresa, string xTrack)
        {
            string query = "";
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();
            query = "  DELETE FROM dte_boe_local" + xCodempresa + " ";
            query += " WHERE  fae_trackenvio_sii ='" + xTrack + "' and fae_status_sii = 'FAU' ";

            cnn = new Conectar(SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

            }

            cnn.CloseConnection();
        }
        public int BlanqueRevisado(string xTipo, string xFolioSII, string xfecha,string xCaja,  string xLocal)
        {
            string query = "";
            string tabla = "";
            int salida = 0;
     
            query = "UPDATE sv_documento_cabeza_" + xLocal + " ";
            query += "SET revisado = '' ";
            query += " Where local = '" + xLocal + "' and tipo = '" + xTipo + "' and numero = '" + xFolioSII + "' ";
            query += " and caja = '" + xCaja + "' and fecha = '" + xfecha + "' ";
            query += " ";

            cnn = new Conectar(this.SERVER, "eltit_ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS, 180);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                salida = cmd.ExecuteNonQuery();

            }
            cnn.CloseConnection();
            return salida;
        }

        public MySqlDataReader getBoletasByTrack(string xRutEmpresa, string xCodEmpresa, string xTrack, string xStatus)
        {
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();
            string query = "";
            MySqlDataReader dr = null;

            query += " SELECT * ";
            query += " FROM dte_boe_local" + xCodEmpresa + "   ";
            query += " WHERE fae_trackenvio_sii = '"+ xTrack + "'  ";
            
            if(xStatus != "" )
            {
                query += " And fae_status_sii = '" + xStatus + "' ";
            }
            query += " ORDER BY fae_fecha, fae_folio ";

            cnn = new Conectar(this.SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS, 180);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.CommandTimeout = 240;
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public int ActualizaEstadoDTEBoleta(string xTipo, string xFolioSII, string xfecha,string xStatus, string xGlosaStatus,  string xRutEmpresa, string xCodEmpresa)
        {
            string query = "";
            string tabla = "";
            int salida = 0;
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();

            if (xTipo == "39" || xTipo == "41")
            {
                tabla = "dte_boe_local";
            }
            else
            {
                tabla = "dte_fae_local";
            }

            query = "UPDATE " + tabla + xCodEmpresa + " ";
            query += "SET fae_status_sii = '" + xStatus + "', fae_glosa_sii = '" + xGlosaStatus + "' "; 
            query += " Where fae_tipo = '" + xTipo + "' and fae_folio = '" + xFolioSII + "' and fae_fecha = '" + xfecha + "' ";

            cnn = new Conectar(this.SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS, 180);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                salida = cmd.ExecuteNonQuery();

            }
            cnn.CloseConnection();
            return salida;
        }
        public MySqlDataReader getXMLEmpresaSinEstado(string xRutEmpresa, string xCodEmpresa, int xlimite)
        {
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();
            string query = "";
            MySqlDataReader dr = null;

            query += " SELECT * ";
            query += " FROM dte_boe_local" + xCodEmpresa + "   ";
            query += " WHERE fae_status_sii = '' and fae_trackenvio_sii <> '' ORDER BY fae_fecha, fae_folio ";
            query += "  LIMIT 0," + xlimite +" ";

            cnn = new Conectar(this.SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS, 180);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.CommandTimeout = 240;
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public int GrabaSobreEnvioBOLETA(string CodEmpresa,string xRutEmpresa,string xNombreSobre, string XMLSobre)
        {
            string query = "";
            int ok = 0;
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();

            query = "Insert Into dte_fae_sobres_envios" + CodEmpresa + "(fecha_envio,hora_envio, ";
            query += " nombre_archivo, xml, trackid) ";
            query += " Values(";
            query += " NOW(),CURRENT_TIME(), '" + xNombreSobre + "' , '" + XMLSobre + "','" + xRutEmpresa + "') ";

            cnn = new Conectar(this.SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS, 180);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                ok = cmd.ExecuteNonQuery();
            }

            cnn.CloseConnection();
            return ok;
        }

        public int ActualizaTrackEnDTE(string xLocal, string xTipo, string xFolioSII, string xfecha,
                                       string xHora, string xArchivoEnvio, string xTrack, 
                                        string xRutEmpresa, string xCodEmpresa)
        {
            string query = "";
            string tabla = "";
            int salida = 0;
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();

            if (xTipo == "39" || xTipo == "41")
            {
                tabla = "dte_boe_local";
            }
            else
            {
                tabla = "dte_fae_local";
            }

            query = "UPDATE " + tabla + xCodEmpresa + " ";
            query += "SET fae_fechaenvio_sii = '" + xfecha + "', fae_horaenvio_sii = '" + xHora + "', ";
            query += "fae_trackenvio_sii = '" + xTrack + "',fae_sobreenvio_sii = '" + xArchivoEnvio + "' ";
            query += " Where fae_tipo = '" + xTipo + "' and fae_folio = '" + xFolioSII + "' and fae_recinto = '"+ xLocal +"' ";

            cnn = new Conectar(this.SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS, 180);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                salida = cmd.ExecuteNonQuery();

            }
            cnn.CloseConnection();
            return salida;
        }

        public MySqlDataReader getXMLEmpresaByNumeroLocalCaja(string xRutEmpresa,string xCodEmpresa, 
                                                                string xLocal,string xNumero, string xCaja, string xFecha)
        {
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();
            string query = "";
            MySqlDataReader dr = null;

            query += " SELECT fae_xml ";
            query += " FROM dte_boe_local" + xCodEmpresa + "   ";
            query += " WHERE fae_folio = '"+ xNumero +"' and fae_recinto = '"+ xLocal +"' and fae_cajadocumento = '"+ xCaja +"' ";
            query += " AND fae_trackenvio_sii = '' and fae_fecha >= '2020-12-01' LIMIT 0,1 ";

            cnn = new Conectar(this.SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS, 240);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.CommandTimeout = 240;
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        public MySqlDataReader getXMLEmpresa(string xRutEmpresa, string CodEmpresa, string xFechaHasta, int xLimite)
        {
            string rut_base = Convert.ToDouble(xRutEmpresa.Substring(0, 9)).ToString();
            string query = "";
            MySqlDataReader dr = null;

            query += " SELECT fae_folio,fae_tipo, fae_recinto, fae_fecha, fae_tipodocumento, fae_cajadocumento, fae_xml, fae_cliente_rut ";
            query += " FROM dte_boe_local"+ CodEmpresa +" WHERE fae_numerointerno <> '' AND fae_trackenvio_sii = '' ";
            query += " AND (fae_fecha > '2020-12-31' AND fae_fecha < '"+ xFechaHasta +"'  ) ";
            query += " ORDER BY  fae_fecha,fae_recinto, fae_folio LIMIT 0," + xLimite +" ";

            cnn = new Conectar(this.SERVER, "eltit_dte_" + rut_base, MYSQL_ROOT, MYSQL_PASS, 350);

            if (cnn.OpenConnection() == true)
            {  
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.CommandTimeout = 350;
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public int GrabaXML(string CodEmpresa, string xLocal, string xRutDB, string xTipoSII, int xFolioSII, string xFechaEmision, string xTipoInterno,
                       string xNroInterno, string xCajaDocumento, string XML, string xRutCliente, double xMontoTotal)
        {
            string query = "";
            int ok = 0;

            query = "Insert ignore Into dte_boe_local" + CodEmpresa + "(fae_tipo,fae_folio,fae_fecha,fae_tipodocumento, ";
            query += "fae_cajadocumento,fae_numerointerno,fae_xml, fae_cliente_rut,fae_recinto, fae_monto_total) ";
            query += " Values(";
            query += " '" + xTipoSII + "','" + xFolioSII + "', '" + xFechaEmision + "' , '" + xTipoInterno + "', ";
            query += " '" + xCajaDocumento + "','" + xNroInterno.PadLeft(10, Convert.ToChar("0")) + "','" + XML + "','" + xRutCliente + "','" + xLocal + "'," + xMontoTotal + ")  ";
            query += " ";
            cnn = new Conectar(SERVER, _CLIENTE + "dte_" + xRutDB, this.MYSQL_ROOT, this.MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                ok = cmd.ExecuteNonQuery();
            }

            cnn.CloseConnection();
            return ok;
        }

        public MySqlDataReader GetBoletasByLocalDia(string xCodEmpresa, string xBaseDTE, string xBaseVentas, string xDia, string xTipDte, string xTipoInterno, DataTable dt)
        {
            string query = "";

            MySqlDataReader dr = null;
            string local = "";

            try
            {
                foreach (DataRow row in dt.Rows)
                {

                    local = row["codigo"].ToString();
                    //query += " SELECT dte.tipo AS tipo_doc, dte.numero AS foliosii,dte.fecha AS fecha_emision, dte.fechaenviosii AS fae_fechaenvio_sii, ";
                    //query += " '00:00:00' AS fae_horaenvio_sii, dte.cajadocumento AS caja_doc,  IFNULL(0,0) AS monto_exento, IFNULL(dc.total,0) AS monto_total,  ";
                    //query += " dte.xml AS fae_xml , dte.localdocumento FROM eltit_fae" + local + ".sv_dte" + local + " AS dte ";
                    //query += " LEFT JOIN eltit_ventas" + local + ".sv_documento_cabeza_" + local + " AS dc  ON(LPAD(dte.numero,10,'0') = dc.foliosii) ";
                    //query += " AND dc.local = dte.localdocumento AND dc.fecha  = dte.fechadocumento AND dte.cajadocumento = dc.caja ";
                    //query += " AND dte.tipodocumento = dc.tipo ";
                    //query += " WHERE dc.local = '" + local + "' AND dc.fecha = '" + xDia + "' AND dte.tipo ='" + xTipos + "' ";

                    query += " SELECT " + xTipDte + " AS tipo_doc, dc.foliosii AS foliosii,dc.fecha AS fecha_emision, ";
                    query += " '00:00:00' AS fae_horaenvio_sii, dc.caja AS caja_doc,  IFNULL(0,0) AS monto_exento, IFNULL(dc.total,0) AS monto_total, ";
                    query += " IFNULL(dte.xml,'') AS fae_xml ,dte.localdocumento FROM  ";
                    query += " eltit_ventas" + local + ".sv_documento_cabeza_" + local + " AS dc ";
                    query += " LEFT JOIN  eltit_fae" + local + ".sv_dte" + local + " AS dte ";
                    query += " ON(dc.foliosii =LPAD(dte.numero,10,'0')) ";
                    query += " AND dc.local = dte.localdocumento AND dc.fecha  = dte.fechadocumento AND dte.cajadocumento = dc.caja ";
                    query += " AND dte.tipodocumento = dc.tipo ";
                    query += " WHERE dc.local = '" + local + "' AND dc.fecha = '" + xDia + "' AND dc.tipo ='" + xTipoInterno + "' ";
                    query += "";
                    query += " UNION ";
                }

                query = query.Substring(0, query.Length - 7);
                query += "Order by foliosii ASC  ";


                cnn = new Conectar(this.SERVER, "eltit_ventas" + local, MYSQL_ROOT, MYSQL_PASS, 180);

                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    cmd.CommandTimeout = 540;
                    dr = cmd.ExecuteReader();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex);

            }



            return dr;
        }
        public MySqlDataReader GetBoletasByLocalDiaCliente(string xLocal, string xIPHamachi, string xBaseDTE, string xBaseVentas, string xDia)
        {
            string query = "";

            MySqlDataReader dr = null;
            //Conectar cnn;
            query = "SELECT dc.tipo_doc,dc.foliosii,dc.fecha_emision, dte.fae_fechaenvio_sii, dte.fae_horaenvio_sii,dc.caja_doc,dc.monto_exento, dc.monto_total,dte.fae_xml ";
            query += " FROM local_venta_cabeza_" + xLocal + " AS dc ";
            query += " INNER JOIN " + xBaseDTE + ".dte_boe_local" + xLocal + " AS dte ON(dc.foliosii = LPAD(dte.fae_folio,10,'0')) ";
            query += " AND dc.local = dte.fae_recinto AND dc.fecha_emision  = dte.fae_fecha ";
            query += "WHERE dc.local = '" + xLocal + "' and dc.fecha_emision = '" + xDia + "' ";
            query += "And (dc.tipo_doc ='BEL' OR dc.tipo_doc = 'BEE')  Order by dc.foliosii";

            cnn = new Conectar(xIPHamachi, xBaseVentas, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        public void GrabaRCOF(string xLocal, string xFecha, int nroEnvio, string XML, string xFolioDesde, string xFolioHasta,
                           double xTotal, string track, string xBaseDTE)
        {
            string query = "";

            query = "INSERT INTO dte_boe_rcof" + xLocal + " ";
            query += " (fae_recinto, fae_fecha, fae_nro_envio, ";
            query += " fae_folio_desde, fae_folio_hasta ,fae_xml,";
            query += " fae_fechaenvio_sii, fae_horaenvio_sii, fae_monto_total, fae_trackenvio_sii) ";
            query += " Values(";
            query += " '" + xLocal + "','" + xFecha + "', '" + nroEnvio + "'   ";
            query += " ,'" + xFolioDesde + "', '" + xFolioHasta + "', '" + XML + "', ";
            query += " NOW(), CURRENT_TIME() ," + xTotal + ", '" + track + "' )";
            //query += " ";

            cnn = new Conectar(SERVER, xBaseDTE, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                //Sincroniza sync = new Sincroniza(SERVER, MYSQL_ROOT, MYSQL_PASS);
                //sync.GrabaSincronizador(query, xBaseDTE);

            }
        }
        public void GrabaRCOFDetalles(string xLocal, string xFecha, int nroEnvio, string XMLDetalles, string xUsuario,
                                    string xBaseDTE, double xTotal)
        {
            string query = "";

            query = "INSERT INTO dte_boe_rcof" + xLocal + "_detalles ";
            query += " (fecha_contable, nro_secuencia, ";
            query += " fecha_envio, hora_envio, ";
            query += " xml_detalles,usuario_envio, total_resumen) ";
            query += " Values(";
            query += " '" + xFecha + "', '" + nroEnvio + "', ";
            query += " NOW(), CURRENT_TIME(), ";
            query += " '" + XMLDetalles + "', '" + xUsuario + "', " + xTotal + " ) ";


            cnn = new Conectar(SERVER, xBaseDTE, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                //Sincroniza sync = new Sincroniza(SERVER, MYSQL_ROOT, MYSQL_PASS);
                //sync.GrabaSincronizador(query, xBaseDTE);

            }
        }

        public MySqlDataReader BuscaRCOF(string xEmpresa, string xFecha, string xBaseDatos)
        {
            string query = "";

            MySqlDataReader dr = null;
            query = " SELECT * from  dte_boe_rcof" + xEmpresa + " ";
            query += " Where  fae_recinto = '" + xEmpresa + "' And fae_fecha = '" + xFecha + "' Limit 0,1  ";

            cnn = new Conectar(SERVER, xBaseDatos, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        public int GetUltimoFolioDTEByLocalCaja(string xLocal, string xTipo, string xCaja, string xBaseDatos)
        {
            string query = "";
            MySqlDataReader dr = null;
            int salida = 0;

            /// ULTIMO OLIO DISPONIBLE ////
            query = " SELECT MAX(dte.numero+1) FROM " + this._CLIENTE + "fae" + xLocal + ".sv_dte" + xLocal + " AS dte ";
            query += " WHERE dte.localdocumento = '" + xLocal + "' AND dte.tipo = '" + xTipo + "' ";

            if (xCaja != "")
            {
                query += "AND dte.cajadocumento = '" + xCaja + "' ";
            }

            query += " ";


            cnn = new Conectar(this.SERVER, xBaseDatos, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        salida = Convert.ToInt32(dr[0].ToString());
                    }
                }


            }

            return salida;
        }

        public MySqlDataReader GetNotasCreditoBoletaByLocalDia(string xLocal, string xDia, string xBaseDatos, string xBaseDTE)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "SELECT dc.tipo_doc,dc.foliosii,dc.fecha_emision, dte.fae_fechaenvio_sii, dte.fae_horaenvio_sii,dc.caja_doc,dc.monto_exento, dc.monto_total,dte.fae_xml ";
            query += " FROM local_venta_cabeza_" + xLocal + " AS dc ";
            query += " INNER JOIN " + xBaseDTE + ".dte_fae_local" + xLocal + " AS dte ON(dc.foliosii = dte.fae_folio) ";
            query += " AND dc.caja_doc = dte.fae_cajadocumento AND dc.fecha_emision  = dte.fae_fecha ";
            query += "WHERE dc.local = '" + xLocal + "' and dc.fecha_emision = '" + xDia + "' ";
            query += "And dc.tipo_doc ='NFE' and (dc.ref_tipo = 'BEL'  OR dc.ref_tipo='BEE') Order by dc.foliosii";

            cnn = new Conectar(this.SERVER, xBaseDatos, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        public void GrabaXML(string xLocal, string xTipoFiscal, int xFolioFiscal, string xTipoInterno, string xNroInterno, string xFecha,
                           string xBase, string XML, string xCaja, string xRut, string xNombre, double xMonto)
        {
            string query = "";

            /*
             * Region Que elimina el antiguo XML para poder insertar el local
             * 
             * */
            this.EliminaDTELocal(xLocal, xTipoInterno, xNroInterno, xFecha, xCaja);

            /////////////// by jaimiko   2021-08-12 //////////////

            if (XML.Contains("<?xml") == true)
            {
               XML = XML.Substring(45, XML.Length-45);

            }

            query = "  INSERT INTO sv_dte" + xLocal + " (";
            query += "  tipo, numero, fecha, tipodocumento, ";
            query += "  cajadocumento, localdocumento, numerodocumento,fechadocumento, ";
            query += "  xml, rut, nombre, monto) Values(";
            query += "  '" + xTipoFiscal + "','" + xFolioFiscal + "','" + xFecha + "','" + xTipoInterno + "', ";
            query += "  '" + xCaja + "','" + xLocal + "','" + xNroInterno + "','" + xFecha + "', ";
            query += "  '" + XML + "','" + xRut + "','" + xNombre + "','" + xMonto + "' ";
            query += "  )";

            cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                //Sincroniza sync = new Sincroniza(SERVER, MYSQL_ROOT, MYSQL_PASS);
                //sync.GrabaSincronizadorMaster(query, xBase);

            }



        }
        public bool VerificaCaf(string xLocal, string xTipo, string xFolio, string xCaja, string xBase)
        {
            string query = "";
            bool salda = false;
            MySqlDataReader dr;
            string tabla = "sv_caf" + xLocal;


            query = " SELECT hasta from  " + tabla + " ";
            query += " where  tipo = '" + xTipo + "' And desde <= '" + xFolio + "' and hasta >= '" + xFolio + "'  ";
            if (xCaja != "")
            {
                query += "  and caja = '" + xCaja + "' ";
            }


            Conectar cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
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
        public int BuscaRangoSiguiente(string xLocal, string xTipo, string xFolio, string xCaja, string xBase)
        {
            string query = "";
            int salda = 0;
            MySqlDataReader dr;
            query = " SELECT desde from  sv_caf" + xLocal + " ";
            query += " where  tipo = '" + xTipo + "' And desde >= '" + xFolio + "' and caja = '" + xCaja + "' Order by desde Limit 0,1  ";

            Conectar cnn = new Conectar(SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        salda = Convert.ToInt32(dr["desde"].ToString());
                    }

                }

                dr.Close();
            }
            cnn.CloseConnection();

            return salda;
        }
        public void EliminaDTELocal(string xLocal, string xTipoInterno, string xNroInterno, string xFechaEmision, string xCaja)
        {
            string query = "";

            query = "  DELETE FROM sv_dte" + xLocal + " ";
            query += " WHERE tipodocumento ='" + xTipoInterno + "' AND numerodocumento ='" + xNroInterno + "' AND fechadocumento = '" + xFechaEmision + "' ";
            query += " AND cajadocumento = '" + xCaja + "' AND localdocumento = '" + xLocal + "' ";

            cnn = new Conectar(SERVER, "eltit_fae" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                //Sincroniza sync = new Sincroniza(SERVER, MYSQL_ROOT, MYSQL_PASS);
                //sync.GrabaSincronizadorMaster(query, xBase);

            }

        }
        public int ActualizaTrackEnDTE(string xLocal, string xTipo, string xFolioSII, string xfechaEnvio,
                                     string xTrack)
        {
            string query = "";
            string tabla = "";
            int salida = 0;

            tabla = "sv_dte";
            query = "UPDATE " + tabla + xLocal + " ";
            query += "SET fechaenviosii = '" + xfechaEnvio + "',  ";
            query += "track = '" + xTrack + "'  ";
            query += " Where tipodocumento = '" + xTipo + "' and numero = '" + xFolioSII + "'  ";

            cnn = new Conectar(this.SERVER, "eltit_fae" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                salida = cmd.ExecuteNonQuery();

            }
            cnn.CloseConnection();
            return salida;
        }

        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }
        //public int GetUltimoFolioDTEByLocalCaja(string xLocal, string xTipo, string xCaja, string xBaseDatos)
        //{
        //    string query = "";
        //    MySqlDataReader dr = null;
        //    int salida = 0;

        //    /// ULTIMO OLIO DISPONIBLE ////
        //    query = " SELECT MAX(dte.numero+1) FROM " + _CLIENTE + "fae" + xLocal + ".sv_dte" + xLocal + " AS dte ";
        //    query += " WHERE dte.localdocumento = '" + xLocal + "' AND dte.tipo = '" + xTipo + "' ";

        //    if (xCaja != "")
        //    {
        //        query += "AND dte.cajadocumento = '" + xCaja + "' ";
        //    }

        //    query += " ";


        //    cnn = new Conectar(this.SERVER, xBaseDatos, MYSQL_ROOT, MYSQL_PASS);
        //    if (cnn.OpenConnection() == true)
        //    {
        //        MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
        //        dr = cmd.ExecuteReader();

        //        if (dr.HasRows == true)
        //        {
        //            if (dr.Read())
        //            {
        //                salida = Convert.ToInt32(dr[0].ToString());
        //            }
        //        }


        //    }

        //    return salida;
        //}










    }
}
