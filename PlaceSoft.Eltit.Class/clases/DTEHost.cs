using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlaceSoft.Eltit.Class.clases
{
    public class DTEHost
    {
        Conectar cnn;
        private string SERVER;
        private string MYSQL_ROOT;
        private string MYSQL_PASS;
        private string _CLIENTE = "empresas_eltit_";
        
        private static readonly log4net.ILog log =
          log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public DTEHost(string xServer, string xmysqlRoot, string xmysqlPass)
        {
            this.SERVER = xServer;
            this.MYSQL_ROOT = xmysqlRoot;
            this.MYSQL_PASS = xmysqlPass;
        }


        public int GrabaXML(string CodEmpresa,string xLocal,string xRutDB, string xTipoSII, int xFolioSII, string xFechaEmision, string xTipoInterno,
                         string xNroInterno, string xCajaDocumento, string XML , string xRutCliente, double xMontoTotal)
        {
            string query = "";
            int ok = 0;

            query = "Insert Into dte_boe_local" + CodEmpresa + "(fae_tipo,fae_folio,fae_fecha,fae_tipodocumento, ";
            query += "fae_cajadocumento,fae_numerointerno,fae_xml, fae_cliente_rut,fae_recinto, fae_monto_total) ";
            query += " Values(";
            query += " '" + xTipoSII + "','" + xFolioSII + "', '" + xFechaEmision + "' , '" + xTipoInterno + "', ";
            query += " '" + xCajaDocumento + "','" + xNroInterno.PadLeft(10, Convert.ToChar("0")) + "','" + XML + "','" + xRutCliente + "','" + xLocal + "'," + xMontoTotal + ")  ";
            query += " ";
            cnn = new Conectar(SERVER, _CLIENTE + "dte_" + xRutDB , this.MYSQL_ROOT, this.MYSQL_PASS);
            if (cnn.OpenConnection() == true)
            {
                MySqlCommand cmd = new MySqlCommand(query, cnn.connection);
                ok = cmd.ExecuteNonQuery();                
            }

            cnn.CloseConnection();
            return ok;
        }
    }
}
