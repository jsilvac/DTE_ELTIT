using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SamplesDTE
{
    public partial class MuestraTimbre : Form
    {
        public MuestraTimbre()
        {
            InitializeComponent();
        }

        private void MuestraTimbre_Load(object sender, EventArgs e)
        {

        }

        private void botonCargarDTE_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string pathFile = openFileDialog1.FileName;
            string xml = File.ReadAllText(pathFile, Encoding.GetEncoding("ISO-8859-1"));
            var dte = PlaceSoft.DTE.Engine.XML.XmlHandler.DeserializeFromString<PlaceSoft.DTE.Engine.Documento.DTE>(xml);
            pictureBoxTimbre.Image = dte.Documento.TimbrePDF417;
        }
    }
}
