using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Eltit.Clases
{
    class Ventas
    {
        Conectar cnn;
        private string MYSQL_SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;
        private string BASE_DTE;
        private string _CLIENTE;

        public Ventas()
        {

        }
        public Ventas(string xCliente, string xServer, string xmysqlRoot, string xmysqlPass)
        {

            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
            this.MYSQL_SERVER = xServer;
            _CLIENTE = xCliente;
        }

        public MySqlDataReader getDetalleGuiasByTipoNroCaja(string xLocal, string xTipo, string xNro, string xCaja, string xFecha)
        {
            string query = "";
            string base_dte = _CLIENTE + "fae" + xLocal;
            string tabla = "";
            MySqlDataReader dr = null;

            /*
             * descuento: Dato que corresponde al descuento en porcentaje
             * descuentopesos : Datos que corresponde al descuento en pesos
             * impuesto: Dato que corresponde al codigo impuesto de tabla impuestos[string de 5 digitos]
             * porcentajeimpuesto: Dato que corresponde al impuesto en porcentaje de la tabla impuesto[decimales]
             * 
             * */
            query = "SELECT dd.tipo, dd.numero, dd.caja, dd.fecha,dc.vencimiento, dd.linea, dd.rut, dd.codigo, dd.descripcion,  ";
            query += "dd.cantidad, dd.precio, dd.descuento, dd.descuentopesos,dd.total, dd.vendedor,  ";
            query += " dd.impuesto, dd.porcentajeimpuesto, imp.codigofae, imp.nombrecorto, imp.porcentaje AS taza, ";
            query += " dd.tipo, dd.numero ,dc.neto, dc.iva, dc.exento, dc.total, dd.tipodocumento as ref_tipo, ";
            query += " dd.numerodocumento as ref_numero, '' as observacion, '0' as tipo_traslado, '' as glosa_guia  ";
            query += " FROM " + _CLIENTE + "ventas" + xLocal + ".sv_guias_detalle_" + xLocal + " as dd ";
            query += " INNER JOIN  " + _CLIENTE + "ventas" + xLocal + ".sv_guias_cabeza_" + xLocal + " as dc  ";
            query += " ON(dd.local = dc.local) AND dd.tipo = dc.tipo AND dd.numero = dc.numero ";
            query += " AND dd.caja = dc.caja AND dd.rut = dc.rut ";
            query += " INNER JOIN  " + _CLIENTE + "gestion.g_maestroimpuestos AS imp  ";
            query += " ON(IF(dd.impuesto = '0','00000',dd.impuesto) = imp.codigo)  ";
            query += " WHERE dd.local ='" + xLocal + "' AND dd.tipo = '" + xTipo + "' AND dd.numero ='" + xNro + "' and dd.caja = '" + xCaja + "' ";
            query += " and dd.fecha = '" + xFecha + "' ";
            query += " Order by dd.linea ";

            cnn = new Conectar(MYSQL_SERVER, base_dte, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public string LeeformaPago(string xLocal, string xNumero, string xTipo, string xCaja, string xFecha)
        {
            string query = "";
            string base_dte = this.BASE_DTE;
            MySqlDataReader dr = null;
            string salida = "";

            query = "SELECT dp.tipopago, tp.nombre  ";
            query += "FROM sv_documento_pagos_" + xLocal + " as  dp  ";
            query += "INNER JOIN eltit_ventas.sv_tiposdepagoclientes as tp ";
            query += " ON(LPAD(dp.tipopago,2,'0') = tp.codigo ) ";
            query += "WHERE dp.numero = '" + xNumero + "' And dp.tipo = '" + xTipo + "' and dp.caja = '" + xCaja + "' and dp.fecha = '" + xFecha + "'  ";

            cnn = new Conectar(MYSQL_SERVER, "eltit_ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        salida = dr["tipopago"].ToString() + "-" + dr["nombre"].ToString();
                    }
                }

            }

            return salida;
        }

        public string GetFechaReferencia(string xLocal, string xFolioSII, string xTipo)
        {
            string query = "";
            string base_dte = this.BASE_DTE;
            string salida = "";
            MySqlDataReader dr = null;

            query = "SELECT fecha  ";
            query += "FROM " + _CLIENTE + "ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + "   ";
            query += "WHERE local = '" + xLocal + "' AND foliosii = '" + xFolioSII + "' AND tipo = '" + xTipo + "' ";
            query += " AND fecha > '2020-01-01' ";

            cnn = new Conectar(MYSQL_SERVER, "eltit_ventas" + xLocal, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();

                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        salida = dr["fecha"].ToString();
                    }
                }
            }

            return salida;
        }


        public MySqlDataReader getXmlByNroAtencion(string xLocal, string xnumero)
        {
            string query = "";
            string base_dte = this.BASE_DTE;
            MySqlDataReader dr = null;

            query = "SELECT *  ";
            query += "FROM plcerti_fae_local00   ";
            query += "WHERE sii_caso  Like '%" + xnumero + "%'  order by sii_caso ";


            cnn = new Conectar(MYSQL_SERVER, base_dte, MYSQL_ROOT, MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }


        public MySqlDataReader getVentaDetalleDocumentosByTipoNroCaja(string xLocal, string xTipo, string xNro, string xCaja, string xFecha)
        {
            string query = "";
            string base_dte = _CLIENTE + "fae" + xLocal;
            string tabla = "";
            MySqlDataReader dr = null;

            /*
             * descuento: Dato que corresponde al descuento en porcentaje
             * descuentopesos : Datos que corresponde al descuento en pesos
             * impuesto: Dato que corresponde al codigo impuesto de tabla impuestos[string de 5 digitos]
             * porcentajeimpuesto: Dato que corresponde al impuesto en porcentaje de la tabla impuesto[decimales]
             * 
             * */
            query = "SELECT dd.tipo, dd.numero, dd.caja, dd.fecha,dc.vencimiento, dd.linea, dd.rut, dd.codigo, dd.descripcion,  ";
            query += "dd.cantidad, dd.precio, dd.descuento, dd.descuentopesos,dd.total, dd.vendedor,  ";
            query += " dd.impuesto, dd.porcentajeimpuesto, imp.codigofae, imp.nombrecorto, imp.porcentaje AS taza, ";
            query += " dd.tipo, dd.numero ,dc.neto, dc.iva, dc.exento, dc.total, dd.tipodocumento as ref_tipo, ";
            query += " dd.numerodocumento as ref_numero, '' as observacion, '0' as tipo_traslado, '' as glosa_guia  ";
            query += " FROM " + _CLIENTE + "ventas" + xLocal + ".sv_documento_detalle_" + xLocal + " as dd ";
            query += " INNER JOIN  " + _CLIENTE + "ventas" + xLocal + ".sv_documento_cabeza_" + xLocal + " as dc  ";
            query += " ON(dd.local = dc.local) AND dd.tipo = dc.tipo AND dd.numero = dc.numero ";
            query += " AND dd.caja = dc.caja AND dd.rut = dc.rut ";
            query += " INNER JOIN  " + _CLIENTE + "gestion.g_maestroimpuestos AS imp  ";
            query += " ON(IF(dd.impuesto = '0','00000',dd.impuesto) = imp.codigo)  ";
            query += " WHERE dd.local ='" + xLocal + "' AND dd.tipo ='" + xTipo + "' AND dd.numero ='" + xNro + "' and dd.caja = '" + xCaja + "' ";
            query += " and dd.fecha = '" + xFecha + "' ";
            query += " Order by dd.linea ";

            cnn = new Conectar(MYSQL_SERVER, base_dte, MYSQL_ROOT, MYSQL_PASS);
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

        public MySqlDataReader GetDocumentoCabeza(string xLocal, string xTipo, string xFolio, string xCaja, string xFecha)
        {
            string salida = "";

            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT * from sv_documento_cabeza_00 ";
            query += " WHERE local='"+ xLocal +"' AND ";
            query += " tipo='"+ xTipo +"' AND ";
            query += " numero='"+ xFolio +"' AND ";
            query += " caja='"+ xCaja +"' AND ";
            query += " fecha='"+ xFecha +"' LIMIT 1";

            try
            {
                Conectar cnn = new Conectar(MYSQL_SERVER, BASE_DTE, MYSQL_ROOT, MYSQL_PASS);
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


    }
}
