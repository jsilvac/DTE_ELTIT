using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eltit.Clases
{    
        public class ItemBoleta
        {
            public string Codigo { get; set; }
            public string Nombre { get; set; }
            public double Cantidad { get; set; }
            public int Precio { get; set; }
            public int Porce_Descuento { get; set; }
            public int Monto_Descuento { get; set; }
            public int Total { get; set; }
            public bool Afecto { get; set; }
            public string UnidadMedida { get; set; }
            public double Porce_impuesto { get; set; }

    }
}