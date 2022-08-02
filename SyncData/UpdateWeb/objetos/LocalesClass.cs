using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace SchoolManagementAdmin.objetos
{
    class LocalesClass
    {
        Conectar cnn;

        public MySqlDataReader GetLocalByCodigo()
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT cli.prefijo, loc.codigo, loc.nombrelocal,loc.servidor_ventas,loc.servidorprincipal, ";
            query += " loc.critico_33,loc.critico_39, loc.critico_52, loc.critico_61, loc.rut ,cli.mysql_user, cli.mysql_pass  ";
            query += " FROM clientes_locales AS loc INNER JOIN clientes_dte AS cli ON(loc.rut = cli.rut) ";
            query += " WHERE loc.emite_39 = 1 And loc.codigo <> '25'  ";
           // query += " and (loc.codigo = '17' OR loc.codigo = '00') ";
            query += "  ";
            query += " ORDER BY loc.codigo ";
            query += "  ";

            cnn = new Conectar(Inicial.G_SERVIDOR, "eltit_dte_manager" , Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }

        public MySqlDataReader GetFoliosCriticosFacturas(string xTipo,string xCliente, string xLocal, string xServidor,string xBase , 
                                            string xUsuario, string xClave)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT  caf.tipo,caf.fecharecepcion, ";
            query += " IFNULL(  (  SELECT MAX(dte.numero) FROM sv_dte"+ xLocal +" AS dte ";
            query += " WHERE caf.tipo = dte.tipo and dte.administracion = '' LIMIT 0,1   ),  0 )AS ultimo, caf.hasta  ";
            query += " FROM sv_caf"+ xLocal +" AS caf  WHERE tipo = '"+ xTipo +"'  ";
            query += " ORDER BY caf.fecharecepcion DESC LIMIT 0,1 ;  ";
            query += "  ";

            cnn = new Conectar(xServidor,xBase, xUsuario, xClave);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
            }

            return dr;
        }
        public MySqlDataReader GetFoliosCriticosBoletas(string xTipo, string xCliente, string xLocal, string xServidor, string xBase,
                                           string xUsuario, string xClave, string xCaja)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = " SELECT  caf.tipo,caf.fecharecepcion, ";
            query += " IFNULL(  (  SELECT MAX(dte.numero) FROM sv_dte" + xLocal + " AS dte ";
            query += " WHERE caf.tipo = dte.tipo AND dte.cajadocumento = '"+ xCaja + "' LIMIT 0,1   ),  0 )AS ultimo, caf.hasta  ";
            query += " FROM sv_caf" + xLocal + " AS caf  WHERE tipo = '" + xTipo + "' and caja = '"+ xCaja +"' ";
            query += " ORDER BY caf.fecharecepcion DESC LIMIT 0,1 ;  ";
            query += "  ";

            cnn = new Conectar(xServidor, xBase, xUsuario, xClave);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
              
            }

            return dr;
        }


        public DataTable GetCajasByLocal(string xCliente, string xLocal, string xServidor,
                                           string xUsuario, string xClave)
        {
            string query = "";
            MySqlDataReader dr = null;
            DataTable dt = new DataTable();

            query  = " SELECT caja FROM sv_caf" + xLocal +" AS caf WHERE tipo = '39' AND caja <> ''  ";
            query += " GROUP BY caja ORDER BY caja ";

            cnn = new Conectar(xServidor, "eltit_fae" + xLocal , xUsuario, xClave);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
                dt.Load(dr);
            }

            return dt;
        }



        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }

    }
}