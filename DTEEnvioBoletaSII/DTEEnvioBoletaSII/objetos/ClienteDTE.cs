using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SchoolManagementAdmin.objetos
{
    class ClienteDTE
    {
        Conectar cnn;
        private string prefijo;
        private string rut;
        private string local;
        private string razon_social;
        private string fecha_activacion;
        private string ip_servidor;
        private string fecha_finalizacion;
        private bool activo;
        private string servidor_destino;
        private string fecha_resolucion;
        private int numero_resolucion;
        private string rut_certificado;
        private string nombre_certificado;
        private int caf_critico_33;
        private int caf_critico_39;
        private int caf_critico_61;
        private string cloud_up;
        private string mysql_user;
        private string mysql_pass;
        private bool sube_dte_boletas;
        private bool sube_dte_facturas;
        private bool sube_ventas;
        private DateTime inicio_sincronizacion;
        private string where;
        private int numero_registros;
        private string smtp_intercambio;
        private string smtp_direccion;
        private string smtp_clave;
        private string rut_database;

        public string Prefijo { get => prefijo; set => prefijo = value; }
        public string Rut { get => rut; set => rut = value; }
        public string Local { get => local; set => local = value; }
        public string Razon_social { get => razon_social; set => razon_social = value; }
        public string Fecha_activacion { get => fecha_activacion; set => fecha_activacion = value; }
        public string IP_servidor { get => ip_servidor; set => ip_servidor = value; }
        public string Fecha_finalizacion { get => fecha_finalizacion; set => fecha_finalizacion = value; }
        public bool Activo { get => activo; set => activo = value; }
        public string Servidor_destino { get => servidor_destino; set => servidor_destino = value; }
        public string Fecha_resolucion { get => fecha_resolucion; set => fecha_resolucion = value; }
        public int Numero_resolucion { get => numero_resolucion; set => numero_resolucion = value; }
        public string Rut_certificado { get => rut_certificado; set => rut_certificado = value; }
        public string Nombre_certificado { get => nombre_certificado; set => nombre_certificado = value; }
        public int Caf_critico_33 { get => caf_critico_33; set => caf_critico_33 = value; }
        public int Caf_critico_39 { get => caf_critico_39; set => caf_critico_39 = value; }
        public int Caf_critico_61 { get => caf_critico_61; set => caf_critico_61 = value; }
        public string Cloud_up { get => cloud_up; set => cloud_up = value; }
        public string Mysql_user { get => mysql_user; set => mysql_user = value; }
        public string Mysql_pass { get => mysql_pass; set => mysql_pass = value; }
        public bool Sube_dte_boletas { get => sube_dte_boletas; set => sube_dte_boletas = value; }
        public bool Sube_dte_facturas { get => sube_dte_facturas; set => sube_dte_facturas = value; }
        public bool Sube_ventas { get => sube_ventas; set => sube_ventas = value; }
        public DateTime Inicio_sincronizacion { get => inicio_sincronizacion; set => inicio_sincronizacion = value; }
        public string Where { get => where; set => where = value; }
        public int Numero_registros { get => numero_registros; set => numero_registros = value; }
        public string Smtp_intercambio { get => smtp_intercambio; set => smtp_intercambio = value; }
        public string Smtp_direccion { get => smtp_direccion; set => smtp_direccion = value; }
        public string Smtp_clave { get => smtp_clave; set => smtp_clave = value; }
        public string Rut_database { get => rut_database; set => rut_database = value; }

        public ClienteDTE(string prefijo, string rut, string local)
        {
            string query = "";
            MySqlDataReader dr = null;

            query = "Select *  From clientes_dte ";
            query += "Where activo  = 1 and prefijo = '"+ prefijo +"' and local = '"+ local +"'";

            cnn = new Conectar(Inicial.G_SERVIDOR, "aliupos_manager", Inicial.G_MYSQL_USER, Inicial.G_MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        this.Prefijo = dr["prefijo"].ToString();
                        this.Rut = dr["rut"].ToString();
                        this.Local = dr["local"].ToString();
                        this.Razon_social = dr["razon_social"].ToString();
                        this.Fecha_activacion = dr["fecha_activacion"].ToString();
                        this.IP_servidor = dr["ip_servidor"].ToString();
                        this.Fecha_finalizacion = dr["fecha_finalizacion"].ToString();
                        this.activo = Convert.ToBoolean(dr["activo"]);
                        this.Servidor_destino = dr["servidor_destino"].ToString();
                        this.Fecha_resolucion = dr["fecha_resolucion"].ToString();
                        this.Numero_resolucion = Convert.ToInt32(dr["numero_resolucion"]);
                        this.Rut_certificado = dr["rut_certificado"].ToString();
                        this.Nombre_certificado = dr["nombre_certificado"].ToString();
                        this.caf_critico_33 = Convert.ToInt32(dr["critico_33"]);
                        this.caf_critico_39 = Convert.ToInt32(dr["critico_39"]);
                        this.caf_critico_61 = Convert.ToInt32(dr["critico_61"]);
                        this.Cloud_up =  dr["cloud"].ToString() ;
                        this.Mysql_user = dr["mysql_user"].ToString();
                        this.Mysql_pass = dr["mysql_pass"].ToString() ;
                        this.Sube_dte_boletas = Convert.ToBoolean(dr["sube_dte_boletas"]);
                        this.Sube_dte_facturas = Convert.ToBoolean(dr["sube_dte_facturas"]); 
                        this.Sube_ventas = Convert.ToBoolean(dr["sube_ventas"]);
                        this.Inicio_sincronizacion = Convert.ToDateTime(dr["inicio_sincroniza"]); 
                        this.Numero_registros = Convert.ToInt32(dr["numero_registros"]);
                        this.smtp_intercambio = dr["mail_intercambio_smtp"].ToString();
                        this.smtp_direccion = dr["mail_intercambio_direccion"].ToString();
                        this.smtp_clave = dr["mail_intercambio_clave"].ToString();
                        this.Rut_database = Convert.ToDouble(dr["rut"].ToString().Substring(0, 9)).ToString();
                    }
                }

                dr.Close();
            }

            cnn.CloseConnection();
            
        }

        public string RetornaWhere()
        {
            string salida = "(query_str LIKE '%dte_boe_rcof" + this.local + "%'  OR ";

            if( this.Sube_dte_boletas == true)
            {
                salida = salida + "query_str LIKE '%dte_boe_local"+ this.local +"%'  OR ";
                salida = salida + "";

              
            }
            if (this.Sube_dte_facturas == true)
            {
                salida = salida + "query_str LIKE '%dte_fae_local" + this.local + "%' OR ";
                salida = salida + "";
                salida = salida + "";
            }
            if (this.Sube_ventas == true)
            {
                salida = salida + " query_str LIKE '%local_venta_detalle_" + this.local + "%' OR ";
                salida = salida + "query_str LIKE '%local_venta_cabeza_" + this.local + "%'  OR ";
                salida = salida + "query_str LIKE '%local_venta_mediodepago_" + this.local + "%'  OR ";
                salida = salida + "query_str LIKE '%local_venta_cobranza_" + this.local + "%'  OR ";
                salida = salida + "query_str LIKE '%local_credito_pago_cabeza_" + this.local + "%'  OR ";
                salida = salida + "query_str LIKE '%local_credito_pago_detalle_" + this.local + "%'  OR ";
                salida = salida + "query_str LIKE '%local_venta_observaciones_" + this.local + "%'  OR ";
                salida = salida + "";

            }
            salida = salida + " query_str LIKE '%dte_fae_sobres_envios%' OR  ";
            salida = salida + " query_str LIKE '%dte_fae_acusesdte%' OR ";
            salida = salida + " query_str LIKE '%dte_fae_envio_proveedores%') ";
            salida = salida + "AND fecha_creacion >= '"+ this.Inicio_sincronizacion.ToString("yyyy/MM/dd") + "' ";
            salida = salida + "AND  query_str NOT LIKE '%local_movimientos_detalle%' ";
            
            //G_WHERE_CONSULTAS = " ( query_str LIKE '%local_venta_detalle_%' OR ";
            //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%dte_fae_sobres_envios%'  OR ";
            //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%local_venta_cabeza_%'  OR ";
            //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%local_venta_mediodepago_%'  OR ";
            //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%dte_boe_local%'  OR ";
            //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str LIKE '%dte_fae_local%' ) And ";
            //G_WHERE_CONSULTAS = G_WHERE_CONSULTAS + " query_str NOT LIKE '%local_movimientos_detalle%' AND ";





            return salida;
        }

        public bool ExisteDTE(string xRutFormat, string xTipoDTE, string xFolio, string xFechaEmision)
        {
            string query = "";
            bool salida = false;
            MySqlDataReader dr = null;

            query = "SELECT * FROM dte_fae_envio_proveedores ";
            query += " WHERE rut_emisor = '"+ xRutFormat +"' and tipo_dte  = '"+ xTipoDTE  +"' ";
            query += " AND folio_dte = '" + xFolio + "' AND fecha_emision = '"+ xFechaEmision +"' ";
            query += " ";


            Conectar cnn;
            cnn = new Conectar(this.IP_servidor, this.Prefijo + "_dte_" + this.Rut_database, this.Mysql_user, this.Mysql_pass);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
                if(dr.HasRows == true)
                {
                    salida = true;
                }
               
            }

            dr.Close();
            cnn.CloseConnection();


            return salida;
        }

        public void GrabaResultadoEnvio(string xRutEmisor, string xTrackSII, string XML, string xEstado,  string xCorreoOrigen)
        {
            string query = "";

            query = "REPLACE INTO dte_fae_envio_resultado_sii( ";
            query += " rut_emisor, track_sii, fecha, ";
            query += " hora, xml, estado, correo_origen ";
            query += ") VALUES( ";
            query += " '" + xRutEmisor + "', '" + xTrackSII + "', NOW(), ";
            query += " CURRENT_TIME(), '" + XML + "','" + xEstado + "','" + xCorreoOrigen + "' ) ";
        
            Conectar cnn;
            cnn = new Conectar(this.IP_servidor, this.Prefijo + "_dte_" + this.Rut_database, this.Mysql_user, this.Mysql_pass);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                /************** GRABA EN SINCRONIZADOR ************/
                Sincroniza sync = new Sincroniza(this, "placesof");
                sync.GrabaSincronizador(query, this.Prefijo + "_dte_" + this.Rut_database, "0");
            }


            cnn.CloseConnection();
        }

        public void GrabaDteAcuse(string xCorreo, string xTipoXML, string xNombreArchivo,
                                     string XML, string xRutResponde,  string xrutRecibe, string xFechaEnvio)
        {
            string query = "";

            query = "REPLACE INTO dte_fae_acusesdte( ";
            query += " correo, tipo_xml, nombre_archivo, ";
            query += " fecha_recepcion, XML, ";
            query += " rut_responde, rut_recibe, fecha_envio";
            query += ") VALUES( ";
            query += " '" + xCorreo + "', '" + xTipoXML + "','" + xNombreArchivo + "', ";
            query += "  NOW(), '" + XML + "',";
            query += " '" + xRutResponde + "','" + xrutRecibe + "','" + xFechaEnvio + "' ) ";
            query += " ";


            Conectar cnn;
            cnn = new Conectar(this.IP_servidor, this.Prefijo + "_dte_" + this.Rut_database, this.Mysql_user, this.Mysql_pass);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                /************** GRABA EN SINCRONIZADOR ************/
                Sincroniza sync = new Sincroniza(this, "placesof");
                sync.GrabaSincronizador(query, this.Prefijo + "_dte_" + this.Rut_database, "0");
            }


            cnn.CloseConnection();
        }
        public void GrabaDteProveedor(string xRutEmisor, string xRazonSocial, string xTipoDTE,  string xFolioDTE,
                                      string xFechaEmision, string xCorreo, double xtotal, string XML, string xNombreArchivo )
        {
            string query = "";
  
            query = "REPLACE INTO dte_fae_envio_proveedores( ";
            query += " rut_emisor, razon_social, tipo_dte, ";
            query += " folio_dte, fecha_emision, correo, ";
            query += " total, fecha_hora, xml, nombre_archivo";
            query += ") VALUES( ";
            query += " '" + xRutEmisor + "', '" + xRazonSocial + "','" + xTipoDTE + "', ";
            query += " '" + xFolioDTE + "', '" + xFechaEmision + "','" + xCorreo + "', ";
            query += " '" + xtotal + "', NOW(),'" + XML + "','" + xNombreArchivo + "' ) ";
            query += " ";
   

            Conectar cnn;
            cnn = new Conectar(this.IP_servidor, this.Prefijo + "_dte_" + this.Rut_database  ,this.Mysql_user,this.Mysql_pass);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                cmd.ExecuteNonQuery();

                /************** GRABA EN SINCRONIZADOR ************/
                Sincroniza sync = new Sincroniza(this, "placesof");
                sync.GrabaSincronizador(query, this.Prefijo + "_dte_" + this.Rut_database, "0");
            }


            cnn.CloseConnection();
        }
        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }


    }
}
