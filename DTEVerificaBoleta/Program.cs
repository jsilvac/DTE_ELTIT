using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SamplesDTE
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool activo;
            System.Threading.Mutex m = new System.Threading.Mutex(true, "frmVerificaTrackEnvioBoleta",
                                       out activo);

            if (!activo)
            {
                MessageBox.Show("Ya Existe una instancia abierta de este Programa.[frmVerificaTrackEnvioBoleta]");
                Application.Exit();
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
               // Application.Run(new frmVerificaTrackEnvioBoletaSII());
                Application.Run(new frmVerificaStatusBoletaSII());
            }
            //Liberamos la exclusión mutua
            m.ReleaseMutex();

            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new frmEnviodDocumentosSII());

        }

    }
}
