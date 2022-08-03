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
    public class Documentos
    {
        Conectar cnn;
        private string CLIENTE_PREFIX = "eltit_";
        private string MYSQL_SERVER = "";
        private string MYSQL_ROOT = "";
        private string MYSQL_PASS = "";
        //private string rut;
        //private string local;


        public Documentos(  string xServer,  string xRoot, string xPass)
        {
           // this.CLIENTE_PREFIX = xCliente;
            this.MYSQL_SERVER = xServer;
            this.MYSQL_PASS = xPass;
            this.MYSQL_ROOT = xRoot;
            //this.rut = xRut;
            //this.local = xLocal;
        }

        public void MarcaRevisionBoletaElectronica(string xLocal, string xnumeroInterno, string xTipoInterno, 
                                                    string xFecha, string xCaja, string xGlosa)
        {
            string query = "";

            query = "  UPDATE eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + "  ";
            query += " SET revisado = '"+ xGlosa +"' ";
            query += "  WHERE local = '" + xLocal + "' AND tipo ='" + xTipoInterno + "' AND numero = '" + xnumeroInterno + "' AND fecha = '" + xFecha + "'  ";
            query += "  and caja = '" + xCaja + "' ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            cnn.CloseConnection();
        }

        public MySqlDataReader GetDocumentosCabezaByLocalNroInternoCaja(string xLocal, string xCaja, 
                                                     string xDesde, string xHasta,    string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dc.tipo, dc.numero, dc.caja, dc.fecha, dc.total, dc.foliosii,'0' as indicador_traslado,dc.rut, dc.revisado ";
            query += " FROM eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc ";
            query += " ";
            query += " WHERE (dc.tipo = 'BV' or dc.tipo = 'BE') and dc.fecha BETWEEN '"+ xDesde + "' AND '" + xHasta + "' ";

            if (xCaja != "")
            {
                query += " and dc.caja = '" + xCaja + "' ";
            }           
            query += " and dc.fecha >= '2021-01-01'and local = '"+ xLocal +"' ORDER By dc.fecha, dc.numero ASC  ";

            cnn = new Conectar(MYSQL_SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }
            return dr;
        }
        
        public void EliminaDocumentoByNumeroCaja(string xLocal, string xnumeroInterno, string xTipoInterno, string xFecha, string xCaja)
        {
            this.EliminaDocumentoCabezaByNumeroCaja(xLocal, xnumeroInterno, xTipoInterno, xFecha, xCaja);
            this.EliminaDocumentoDetalleByNumeroCaja(xLocal, xnumeroInterno, xTipoInterno, xFecha, xCaja);
            this.EliminaDocumentoPagosByNumeroCaja(xLocal, xnumeroInterno, xTipoInterno, xFecha, xCaja);

        }

        private void EliminaDocumentoCabezaByNumeroCaja(string xLocal, string xnumeroInterno, string xTipoInterno, string xFecha, string xCaja)
        {
            string query = "";

            query = "  DELETE FROM  eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + "  ";
            query += "  WHERE local = '" + xLocal + "' AND tipo ='" + xTipoInterno + "' AND numero = '" + xnumeroInterno + "' AND fecha = '" + xFecha + "'  ";
            query += "  and caja = '" + xCaja + "' ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            cnn.CloseConnection();
        }
        private void EliminaDocumentoDetalleByNumeroCaja(string xLocal, string xnumeroInterno, string xTipoInterno, string xFecha, string xCaja)
        {
            string query = "";

            query = "  DELETE FROM  eltit_ventas" + xLocal + ".sv_documento_detalle_" + xLocal + "  ";
            query += "  WHERE local = '" + xLocal + "' AND tipo ='" + xTipoInterno + "' AND numero = '" + xnumeroInterno + "' AND fecha = '" + xFecha + "'  ";
            query += "  and caja = '" + xCaja + "' ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            cnn.CloseConnection();
        }
        private void EliminaDocumentoPagosByNumeroCaja(string xLocal, string xnumeroInterno, string xTipoInterno, string xFecha, string xCaja)
        {
            string query = "";

            query = "  DELETE FROM  eltit_ventas" + xLocal + ".sv_documento_pagos_" + xLocal + "  ";
            query += "  WHERE local = '" + xLocal + "' AND tipo ='" + xTipoInterno + "' AND numero = '" + xnumeroInterno + "' AND fecha = '" + xFecha + "'  ";
            query += "  and caja = '" + xCaja + "' ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            cnn.CloseConnection();
        }

        public string GeneraDocAPartirDeOtro(string xLocal, string xNumero, string xCaja, string xFecha, string xTipo, double xTotal)
        {
            string ultimo = this.GetUltimoNumeroByCajaLocal(xLocal, xTipo, xCaja);

            this.GeneraNuevoDetalle(xLocal, xNumero, xCaja, xFecha, xTipo, ultimo);
            this.GeneraNuevoCabeza(xLocal, xNumero, xCaja, xFecha, xTipo, xTotal, ultimo);
            this.GeneraNuevoPagos(xLocal, xNumero, xCaja, xFecha, xTipo,ultimo);

            return ultimo;
        }
        private void GeneraNuevoCabeza(string xLocal, string xNumero, string xCaja, string xFecha, string xTipo, double xTotal,
                                  string xNuevoNumero )
        {
            string query = "";

            query = "  INSERT INTO eltit_ventas"+ xLocal + ".sv_documento_cabeza_" + xLocal + " (";
            query += " LOCAL,tipo, numero,caja, ";
            query += " fecha, plazo, vencimiento, rut, ";
            query += " sucursal, cajera, notapedido, notaventa, ";
            query += " ordencompra, subtotal, neto, iva, ";
            query += " impuestoharina, impuestocarne, impuestoilarefrescos, impuestoilalicores, ";
            query += " impuestoilavinos, impuestoespecifico, exento, retencionparcial, retenciontotal, ";
            query += " total, abono, descuento, contabilizado, ";
            query += " pagado, comision, fechapagocomision, numeroliquidacion, ";
            query += " porcentajecomision, nula, lugarretiro, cuotas, ";
            query += " monto_cuota, intereses, pie, horaventas, ";
            query += " solocredito, adicional, donacion, boletadesde, ";
            query += " boletahasta, impuestoila, vendedor, descuento2, ";
            query += " transporte,  condicionesdepago, revisado, bultos, ";
            query += " abono2, foliosii, checkeado, fechacreacion, ";
            query += "  puntos, montoredondeo, impresion ";
            query += " ) ";
            query += " SELECT       LOCAL, tipo, '"+ xNuevoNumero +"', caja, ";
            query += " fecha, plazo, vencimiento, rut, ";
            query += " sucursal, cajera, notapedido, notaventa, ";
            query += " ordencompra, subtotal, neto, iva, ";
            query += " impuestoharina, impuestocarne, impuestoilarefrescos, ";
            query += " impuestoilalicores, impuestoilavinos, impuestoespecifico, ";
            query += " exento, retencionparcial, retenciontotal, total, ";
            query += " abono,  descuento, contabilizado, pagado, ";
            query += " comision, fechapagocomision, numeroliquidacion, porcentajecomision, ";
            query += " nula, lugarretiro,  cuotas, monto_cuota, intereses, ";
            query += " pie, horaventas, solocredito, adicional, donacion, ";
            query += " boletadesde, boletahasta, impuestoila, vendedor, ";
            query += " descuento2, transporte, condicionesdepago, 'REGENERADO[" + xNumero + "]', ";
            query += " bultos, abono2, '" + xNuevoNumero + "', checkeado, fechacreacion, ";
            query += " puntos, montoredondeo, impresion ";
            query += " FROM sv_documento_cabeza_"+ xLocal +" ";
            query += " WHERE tipo = '"+ xTipo +"' AND numero = '"+ xNumero +"' ";
            query += " AND caja = '"+ xCaja +"' AND fecha = '"+ xFecha +"' AND total = " + xTotal +" ";
            query += "  ";



            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
            



        }

        private void GeneraNuevoDetalle (string xLocal, string xNumero, string xCaja, string xFecha, string xTipo, string xNuevoNumero)
        {
            string query = "";
      
            query = "  INSERT INTO eltit_ventas"+ xLocal +".sv_documento_detalle_"+ xLocal +" ";
            query += " ( ";
            query += " LOCAL, tipo, numero, linea, fecha, ";
            query += " vencimiento, rut,  sucursal, codigo, ";
            query += " descripcion, cantidad, unidades, precio, ";
            query += " descuento, total, vendedor, pcosto, bodega, ";
            query += " caja, numerofactura, impuesto, porcentajeimpuesto, ";
            query += " descuento2, glosa, horaventas, descuentopesos, ";
            query += " fechacreacion, tipodespacho, tipodocumento, numerodocumento, ";
            query += " despachado, nula, foliosii, cajera, contabilizado) ";
            query += " SELECT LOCAL, tipo, ";
            query += " '"+ xNuevoNumero +"', linea, fecha, vencimiento, ";
            query += " rut, sucursal, codigo, descripcion, ";
            query += " cantidad, unidades, precio, descuento, ";
            query += " total, vendedor, pcosto, bodega, caja, ";
            query += " numerofactura, impuesto, porcentajeimpuesto, descuento2, ";
            query += " 'REGENERADO', horaventas, descuentopesos, fechacreacion, ";
            query += " tipodespacho, tipodocumento, numerodocumento, despachado, ";
            query += " nula, '" + xNuevoNumero + "', cajera, contabilizado ";
            query += " FROM sv_documento_detalle_" + xLocal + " ";
            query += " WHERE tipo = '"+ xTipo +"' AND numero = '"+ xNumero +"' AND caja = '"+ xCaja +"' AND fecha = '"+ xFecha +"' ";
            query += "  ";
            query += "  ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
         
        }

        private void GeneraNuevoPagos(string xLocal, string xNumero, string xCaja, string xFecha, string xTipo, string xNuevoNumero)
        {
            string query = "";
        

            query  = " INSERT INTO eltit_ventas" + xLocal + ".sv_documento_pagos_" + xLocal + " ";
            query += " ( ";
            query += " LOCAL, tipo, numero, lineapago, ";
            query += " fecha, tipopago, cuentacorriente, banco, ";
            query += " plaza, numerodocumento, monto, vencimiento, ";
            query += " rut, glosa, pagoenlazado, localdocumento, ";
            query += " foliofiscal, cuotas, montocuotas, rutcredito, ";
            query += " primervencimiento, caja,  nula, cajera, contabilizado ";
            query += " ) ";
            query += " SELECT local, tipo, '"+ xNuevoNumero +"', lineapago, fecha, ";
            query += " tipopago, cuentacorriente, banco, Plaza, ";
            query += " numerodocumento, monto, vencimiento, rut, ";
            query += " 'REGENERADO', pagoenlazado, localdocumento, ";
            query += " '" + xNuevoNumero + "', cuotas, montocuotas, rutcredito, ";
            query += " primervencimiento, caja, nula, ";
            query += " cajera, contabilizado ";
            query += " FROM sv_documento_pagos_"+ xLocal +" ";
            query += " WHERE tipo = '"+ xTipo +"' AND numero = '"+ xNumero +"' AND caja = '"+ xCaja +"' AND fecha = '"+ xFecha +"'  ";
            query += "  ";
            query += "  ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }
           
        }


        private string GetUltimoNumeroByCajaLocal(string xLocal, string xTipo, string xCaja)
        {
            string query = "";
            MySqlDataReader dr = null;
            string salida = "";

            /// ULTIMO OLIO DISPONIBLE ////
            query = " SELECT MAX(numero+1) FROM sv_documento_cabeza_"+ xLocal +" ";
            query += " WHERE local = '" + xLocal + "' AND tipo = '" + xTipo + "' and caja = '"+ xCaja +"' and contabilizado ='B' ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        salida = dr[0].ToString();
                    }
                }


            }

            return salida.PadLeft(10,Convert.ToChar("0"));
        }

        //public Documentos(string xMysqlServer, string xMysqlRoot, string xMysqlPass)
        //{
        //    this.MYSQL_SERVER = xMysqlServer;
        //    this.MYSQL_PASS = xMysqlPass;
        //    this.MYSQL_ROOT = xMysqlRoot;
        //}

        public MySqlDataReader GetDoucumentosDetalleByTipoCajaNroOnternoFechaLocal(string xLocal, string xTipo, string xNro, string xCaja, string xFecha, string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dd.tipo, dd.numero, dd.caja, dd.fecha,dc.vencimiento, dd.linea, dd.rut, ";
            query += " dd.codigo, dd.descripcion,  dd.cantidad, dd.precio, dd.descuento, dd.descuento,dd.total, ";
            query += " dd.vendedor,  dd.impuesto, dd.impuesto,dd.porcentajeimpuesto,dd.descuentopesos, ";
            query += " dd.tipo, dd.numero ,dc.neto, dc.iva, dc.exento, ";
            //   query += " dc.total, dd.tipo, dd.numero, dd.fecha ,dc.indicador_traslado  ";
            query += " dc.total, dd.tipo, dd.numero, dd.fecha  ";
            query += " FROM " + CLIENTE_PREFIX + "ventas" + xLocal + ".sv_documento_detalle_" + xLocal + " AS dd ";
            query += " INNER JOIN  " + CLIENTE_PREFIX + "ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc  ON(dd.local = dc.local) AND dd.tipo = dc.tipo ";
            query += " AND dd.numero = dc.numero AND dd.caja = dc.caja AND dd.rut = dc.rut AND dd.fecha = dc.fecha ";
            query += " WHERE dd.local ='" + xLocal + "' AND dd.tipo ='" + xTipo + "' AND dd.numero ='" + xNro + "' AND dd.caja = '" + xCaja + "' ";
            query += " AND dd.fecha = '" + xFecha + "' ORDER BY dd.linea ";
            query += " ";

            cnn = new Conectar(MYSQL_SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }
            return dr;
        }

        public MySqlDataReader GetDocumentosCabezaByLocalNroInternoCajaDesdeHasta(string xLocal, string xCaja, string xTipoInterno,
                                                          string xFolioDesde, string xFolioHasta, string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dc.tipo, dc.numero, dc.caja, dc.fecha, dc.total, IFNULL(dte.xml, '0') AS xml, dc.foliosii,'0' as indicador_traslado,dc.rut ";
            query += " FROM eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc ";
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

            cnn = new Conectar(MYSQL_SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }
            return dr;
        }
        public MySqlDataReader GetDocumentosCabezaByLocalNroInternoCajaDesdeHastaFecha(string xLocal, string xCaja, string xTipoInterno,
                                                        string xNumero, string xFecha, string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dc.tipo, dc.numero, dc.caja, dc.fecha, dc.total, IFNULL(dte.xml, '0') AS xml, dc.foliosii,'0' as indicador_traslado,dc.rut ";
            query += " FROM eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc ";
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

            query += " dc.numero = '" + xNumero + "' And dc.fecha = '"+ xFecha +"' ORDER BY dc.numero ASC  ";

            cnn = new Conectar(MYSQL_SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }
            return dr;
        }

        public MySqlDataReader GetDocumentosCabezaByLocalNroInternoCajaFechadesdeFechahasta(string xLocal, string xCaja, string xTipoInterno,
                                                string xNumero, string xFechaDesde,string xFechaHasta, string xBase)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dc.tipo, dc.numero, dc.caja, dc.fecha, dc.total, IFNULL(dte.xml, '0') AS xml, dc.foliosii,'0' as indicador_traslado,dc.rut,dc.cajera ";
            query += " FROM eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc ";
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

            query += "  dc.fecha >= '" + xFechaDesde + "' And dc.fecha <= '" + xFechaHasta + "' and dc.numero = '" + xNumero + "'  ";
           // query += " ORDER BY dc.numero ASC ";

            cnn = new Conectar(MYSQL_SERVER, xBase, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }
            return dr;
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

            query += " (dc.numero >= '" + xFolioDesde + "' AND dc.numero <= '" + xFolioHasta + "') and dc.fecha >= '2021-01-01' ORDER BY dc.numero ASC  ";



            cnn = new Conectar(MYSQL_SERVER, "eltit_ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }
            return dr;
        }

        public void ActualizaFolioSII(string xLocal, string xBase, string xnumeroInterno, string xTipoInterno, string xFecha, string xCaja, string xFolioSII)
        {
            string query = "";

            query = "  UPDATE eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc ";
            query += "  SET dc.foliosii = '" + xFolioSII + "' ";
            query += "  where dc.local = '" + xLocal + "' AND dc.tipo ='" + xTipoInterno + "' AND dc.numero = '" + xnumeroInterno + "' AND dc.fecha = '" + xFecha + "'  ";
            query += "  and dc.caja = '" + xCaja + "' ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }



        }
        public void MarcaBoletaSubidaInternet(string xLocal,string xNroInterno, string xCaja, string xTipo, string xFecha, string xStatus)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " UPDATE  ";
            query += "  eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc ";
            //if(xStatus == "OK")
            //{
            //    query += " SET dc.revisado = now()   ";
            //}
            //else
            //{
            //    query += " SET dc.revisado = 'NO'   ";
            //}
            query += " SET dc.revisado = '"+ xStatus +"'   ";
            query += " WHERE revisado = ''  and tipo = '"+ xTipo +"' and local = '"+ xLocal +"' ";
            query += " and caja = '" + xCaja + "' and fecha = '" + xFecha + "' and numero = '"+ xNroInterno +"' ";


            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();
            }

            cnn.CloseConnection();
        
        }

        public MySqlDataReader GetDocumentosBoletaPendientes(string xLocal, int xLimit)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT dc.tipo, dc.numero, dc.caja, dc.fecha, dc.total, dc.foliosii, dc.rut ";
            query += " FROM eltit_ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc ";
            query += " where local= '"+ xLocal + "' and tipo = 'BV' and caja <> '50' AND fecha >= '2021-01-01' and fecha <= NOW() ";
            query += " and revisado = ''   "; 
           // query += " AND local = '00' ";// borrrar
            query += " Order by fecha, numero Limit 0," + xLimit + " ";

            cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas"+ xLocal, MYSQL_ROOT, MYSQL_PASS);

            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.CommandTimeout = 240;
                dr = cmd.ExecuteReader();
            }
            return dr;
        }

        public DataTable GetDoucumentosDetalleByTipoCajaNroInternoFechaLocal(string xLocal, string xTipo, string xNro, string xCaja, 
                                                                                    string xFecha)
        {
            string query = "";
            MySqlDataReader dr = null;
            DataTable dt = new DataTable();

            query = " SELECT dd.tipo, dd.numero, dd.caja, dd.fecha,dc.vencimiento, dd.linea, dd.rut, ";
            query += " dd.codigo, dd.descripcion,  dd.cantidad, dd.precio, dd.descuento, dd.descuento,dd.total, ";
            query += " dd.vendedor,  dd.impuesto, dd.impuesto,dd.porcentajeimpuesto,dd.descuentopesos, ";
            query += " dd.tipo, dd.numero ,dc.neto, dc.iva, dc.exento, ";
            query += " dc.total, dd.tipo, dd.numero, dd.fecha   ";
            query += " FROM " + this.CLIENTE_PREFIX + "ventas" + xLocal + ".sv_documento_detalle_" + xLocal + " AS dd ";
            query += " INNER JOIN  " + this.CLIENTE_PREFIX + "ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " AS dc  ";
            query += " ON(dd.local = dc.local) AND dd.tipo = dc.tipo ";
            query += " AND dd.numero = dc.numero AND dd.caja = dc.caja AND dd.fecha = dc.fecha ";
            query += " ";
          //  query += " AND dd.rut = dc.rut ";
            query += " WHERE dd.local ='" + xLocal + "' AND dd.tipo ='" + xTipo + "' AND dd.numero ='" + xNro + "' AND dd.caja = '" + xCaja + "' ";
            query += " AND dd.fecha = '" + xFecha + "' ORDER BY dd.linea ";
            query += " ";

            cnn = new Conectar(MYSQL_SERVER,  CLIENTE_PREFIX + "ventas" + xLocal , MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
                dt.Load(dr);

                dr.Close();
                cnn.CloseConnection();
            }
          

            return dt;
        }

        public MySqlDataReader GetDocumentoCabeza(string xLocal, string xTipo, string xFolio, string xCaja, string xFecha)
        {
            string salida = "";

            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT * from sv_documento_cabeza_" + xLocal;
            query += " WHERE local='" + xLocal + "' AND ";
            query += " tipo='" + xTipo + "' AND ";
            query += " numero=lpad('" + xFolio + "',10,'0') AND ";
            query += " caja='" + xCaja + "' AND ";
            query += " fecha='" + xFecha + "' LIMIT 1";

            try
            {
                Conectar cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, "sistema", this.MYSQL_PASS);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        return dr;
                    }
                }

                cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }

            return dr;
        }

        public List<string> GetPagosByDocumento(string xLocal, string xTipo, string xFolio, string xCaja, string xFecha)
        {
            List<string> salida = new List<string>();

            string query = "";
            MySqlDataReader dr = null;

            //SELECT tp.nombre FROM sv_documento_pagos_00 AS dp
            //INNER JOIN eltit_ventas.sv_tiposdepagoclientes AS tp
            //ON(LPAD(dp.tipopago, 2, "0") = tp.codigo)
            //WHERE dp.numero = '0000691575' AND dp.tipo = 'FV';

            query =  " SELECT tp.codigom, tp.nombre FROM sv_documento_pagos_" + xLocal + "AS dp";
            query += " INNER JOIN eltit_ventas.sv_tiposdepagoclientes AS tp ";
            query += " ON(LPAD(dp.tipopago,2,'0') = tp.codigo) ";
            query += " WHERE dp.numero = " + xFolio + ""; 
            query += " AND dp.tipo = "+ xTipo +" ";
            query += " AND dp.fecha = " + xFecha + " ";
            query += " AND dp.caja = "+ xCaja +" ";
            //query += " WHERE local='" + xLocal + "' AND ";
            //query += " tipo='" + xTipo + "' AND ";
            //query += " numero=lpad('" + xFolio + "',10,'0') AND ";
            //query += " caja='" + xCaja + "' AND ";
            //query += " fecha='" + xFecha + "' LIMIT 1";

            try
            {
                Conectar cnn = new Conectar(MYSQL_SERVER, CLIENTE_PREFIX + "ventas" + xLocal, "sistema", this.MYSQL_PASS);
                if (cnn.OpenConnection() == true)
                {
                    MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {                                                 
                         while (dr.Read())
                        {
                            salida.Add(dr["codigo"].ToString() + " "+ dr["nombre"].ToString());
                        }
                        
                    }
                }

                cnn.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exepcion no controlada:" + ex.Message.ToString());
            }

            return salida.ToList<string>() ;
        }

        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }



    }
}
