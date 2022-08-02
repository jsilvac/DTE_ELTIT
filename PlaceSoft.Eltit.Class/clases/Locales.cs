using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlaceSoft.Eltit.Class
{
    public class Locales
    {
        Conectar cnn;
        private string prefijo;
        private string rut;
        private string local;
        private string giro;
        private string razon_social;
        private string direccion_empresa;
        private string comuna_empresa;
        private string codigo_sucursal_sii;
        private string codigo_contable;
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
        private string correo_soporte;
        private string nombrelocal;

        private string MYSQL_SERVER = "";
        private string MYSQL_ROOT = "";
        private string MYSQL_PASS = "";
        private string CLIENTE_PREFIX = "eltit_";


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
        public string Giro { get => giro; set => giro = value; }
        public string Direccion_empresa { get => direccion_empresa; set => direccion_empresa = value; }
        public string Comuna_empresa { get => comuna_empresa; set => comuna_empresa = value; }
        public string Codigo_sucursal_sii { get => codigo_sucursal_sii; set => codigo_sucursal_sii = value; }
        public string Codigo_contable { get => codigo_contable; set => codigo_contable = value; }
        public string Correo_soporte { get => correo_soporte; set => correo_soporte = value; }
        public string Nombrelocal { get => nombrelocal; set => nombrelocal = value; }


        public Locales(string xMysqlServer,string xMysqlRoot, string xMysqlPass)
        {
            this.MYSQL_SERVER = xMysqlServer;
            this.MYSQL_ROOT = xMysqlRoot;
            this.MYSQL_PASS = xMysqlPass;
        }


        public void getLocalDTE(string local)
        {
            string query = "";
            MySqlDataReader dr = null;

            query  = " SELECT loc.codigo, loc.nombrelocal, emp.razon_social, loc.codigo_contable, loc.servidor_ventas, emp.fecha_resolucion, ";
            query += " emp.numero_resolucion,emp.fecha_activacion,emp.nombre_certificado, emp.rut_certificado, ";
            query += " emp.rut, emp.fecha_finalizacion,emp.activo, loc.critico_33,loc.critico_39, loc.critico_61, ";
            query += " emp.giro, emp.direccion as direccionempresa, emp.comuna as comunaempresa, loc.codigo_sucursal_sii, ";
            query += " loc.ventas_mysql_root, loc.ventas_mysql_pass,emp.codigo_contable,correo_soporte ";
            query += " FROM clientes_locales AS loc ";
            query += " INNER JOIN clientes_dte AS emp ON(loc.codigo_contable = emp.codigo_contable ) AND loc.rut = emp.rut ";
            query += " Where loc.codigo = '" + local + "'";

            cnn = new Conectar(this.MYSQL_SERVER, "eltit_dte_manager", this.MYSQL_ROOT, this.MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                dr = cmd.ExecuteReader();
                if (dr.HasRows == true)
                {
                    if (dr.Read())
                    {
                        this.Prefijo = "eltit_";// dr["prefijo"].ToString();
                        this.Rut = dr["rut"].ToString();
                        this.Local = dr["codigo"].ToString();
                        this.Razon_social = dr["razon_social"].ToString();
                        this.Giro = dr["giro"].ToString();
                        this.Direccion_empresa = dr["direccionempresa"].ToString();
                        this.Comuna_empresa = dr["comunaempresa"].ToString();
                        this.Fecha_activacion = dr["fecha_activacion"].ToString();
                        this.IP_servidor = dr["servidor_ventas"].ToString();
                        this.Fecha_finalizacion = dr["fecha_finalizacion"].ToString();
                        this.activo = Convert.ToBoolean(dr["activo"]);
                        this.Codigo_sucursal_sii = dr["codigo_sucursal_sii"].ToString();
                        this.Codigo_contable = dr["codigo_contable"].ToString(); ;
                        //this.Servidor_destino = dr["servidor_destino"].ToString();
                        this.Fecha_resolucion = dr["fecha_resolucion"].ToString();
                        this.Numero_resolucion = Convert.ToInt32(dr["numero_resolucion"]);
                        this.Rut_certificado = dr["rut_certificado"].ToString();
                        this.Nombre_certificado = dr["nombre_certificado"].ToString();
                        this.caf_critico_33 = Convert.ToInt32(dr["critico_33"]);
                        this.caf_critico_39 = Convert.ToInt32(dr["critico_39"]);
                        this.caf_critico_61 = Convert.ToInt32(dr["critico_61"]);
                        //this.Cloud_up = dr["cloud"].ToString();
                        this.Mysql_user = dr["ventas_mysql_root"].ToString();
                        this.Mysql_pass = dr["ventas_mysql_pass"].ToString();
                        //this.Sube_dte_boletas = Convert.ToBoolean(dr["sube_dte_boletas"]);
                        //this.Sube_dte_facturas = Convert.ToBoolean(dr["sube_dte_facturas"]);
                        //this.Sube_ventas = Convert.ToBoolean(dr["sube_ventas"]);
                        //this.Inicio_sincronizacion = Convert.ToDateTime(dr["inicio_sincroniza"]);
                        //this.Numero_registros = Convert.ToInt32(dr["numero_registros"]);
                        //this.smtp_intercambio = dr["mail_intercambio_smtp"].ToString();
                        //this.smtp_direccion = dr["mail_intercambio_direccion"].ToString();
                        //this.smtp_clave = dr["mail_intercambio_clave"].ToString();
                        this.Rut_database = Convert.ToDouble(dr["rut"].ToString().Substring(0, 9)).ToString();

                        this.Correo_soporte = dr["correo_soporte"].ToString();
                        this.Nombrelocal = dr["nombrelocal"].ToString();
                    }
                }

                dr.Close();
            }

            cnn.CloseConnection();

        }


        public void CerrarTransaccion()
        {
            cnn.CloseConnection();
        }


    }
}
