using PlaceSoft.DTE.WS.EstadoDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SamplesDTE.Clases;
using MySql.Data.MySqlClient;

using Eltit.Clases;

namespace Eltit
{
    public class Handler
    {
        public string casoPruebas;
        public int Folio;
        public string idDte;
        public string rutEmpresa = "";

        double neto = 0, netoExento = 0;
        double iva = 0, total = 0;

        public string secuenciaEnvio = "";
        
        public string rutCertificado = "";
        public string nombreCertificado = "";
        public DateTime fechaResolucion ;
        public int numero_resolucion = 0;
        public PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType tipo; //= PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica;
        public bool usaReferencia = false;
        public DateTime fechaEmision;
        public string rutcliente = "";
        public string nombrecliente = "";
        public string cod_sucursal_sii = "";

        public string emisor_rut = "";
        public string emisor_razon_social = "";
        public string emisor_giro = "";
        public string emisor_comuna = "";
        public string emisor_ciudad = "";
        public string emisor_direccion = "";
        public string emisor_acteco = "";
        public string REF_TIPO = "";
        public string REF_NUMERO = "";

        public PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum emisor_tipo_traslado;

        #region Generar Documento Boletas
        public PlaceSoft.DTE.Engine.Documento.DTE GenerateDTEBoleta()
        {
            // DOCUMENTO
            var dte = new PlaceSoft.DTE.Engine.Documento.DTE();


            dte.Documento.Id = idDte;

            // DOCUMENTO - ENCABEZADO - IDENTIFICADOR DEL DOCUMENTO - CAMPOS OBLIGATORIOS
            //TipoDTE = Se indica el tipo de documento. Esta API soporta:
            //Factura Electrónica, Factura Electrónica Exenta, Boleta Electrónica, Nota de Crédito Electrónica, Nota de Débito Electrónica
            dte.Documento.Encabezado.IdentificacionDTE.TipoDTE = tipo;
            dte.Documento.Encabezado.IdentificacionDTE.FechaEmision = fechaEmision;
            dte.Documento.Encabezado.IdentificacionDTE.Folio = Folio;
            dte.Documento.Encabezado.IdentificacionDTE.IndicadorServicio = PlaceSoft.DTE.Engine.Enum.IndicadorServicio.IndicadorServicioEnum.BoletaVentasYServicios;


            //DOCUMENTO - ENCABEZADO - EMISOR - CAMPOS OBLIGATORIOS          
            dte.Documento.Encabezado.Emisor.Rut = emisor_rut.Replace(".", "");
            dte.Documento.Encabezado.Emisor.RazonSocialBoleta = emisor_razon_social;
            dte.Documento.Encabezado.Emisor.GiroBoleta = emisor_giro;
            dte.Documento.Encabezado.Emisor.ComunaOrigen = emisor_comuna;
            dte.Documento.Encabezado.Emisor.CiudadOrigen = emisor_ciudad;
            dte.Documento.Encabezado.Emisor.DireccionOrigen = emisor_direccion;

            if(cod_sucursal_sii != "")
            {
                dte.Documento.Encabezado.Emisor.CodigoSucursal = Convert.ToInt32(cod_sucursal_sii);
            }
            


            //DOCUMENTO - ENCABEZADO - RECEPTOR - CAMPOS OBLIGATORIOS
            dte.Documento.Encabezado.Receptor.Rut = "66666666-6"; //RUT CLIENTE
            dte.Documento.Encabezado.Receptor.RazonSocial = "CLIENTE BOLETA VENTA"; //nombrecliente


            return dte;
        }

        public void GenerateDetails(PlaceSoft.DTE.Engine.Documento.DTE dte, List<ItemBoleta> detalles)
        {
            //DOCUMENTO - DETALLES
            dte.Documento.Detalles = new List<PlaceSoft.DTE.Engine.Documento.Detalle>();

            int contador = 1;
            double totalLinea = 0;
            foreach (var det in detalles)
            {
                var detalle = new PlaceSoft.DTE.Engine.Documento.Detalle();
                /********************** INSTANCIA UN OBJETO PARA EL CODIGO *********************/
                detalle.CodigosItem = new List<PlaceSoft.DTE.Engine.Documento.CodigoItem>();
                var objCodigo = new PlaceSoft.DTE.Engine.Documento.CodigoItem();
                objCodigo.TipoCodigo = "EAN13";
                objCodigo.ValorCodigo = det.Codigo;
                detalle.CodigosItem.Add(objCodigo);
                /*******************************************************************************/

                detalle.NumeroLinea = contador;
                /*IndicadorExento = Sólo aplica si el producto es exento de IVA*/
                detalle.IndicadorExento = det.Afecto ? PlaceSoft.DTE.Engine.Enum.IndicadorFacturacionExencion.IndicadorFacturacionExencionEnum.NotSet : PlaceSoft.DTE.Engine.Enum.IndicadorFacturacionExencion.IndicadorFacturacionExencionEnum.NoAfectoOExento;

                detalle.Nombre = det.Nombre;
                detalle.Cantidad = (double)det.Cantidad;
                detalle.Precio = det.Precio;
                if (det.Porce_Descuento > 0)
                {
                    //detalle.Descuento = det.Monto_Descuento;
                    //detalle.DescuentoPorcentaje = det.Porce_Descuento;
                    detalle.DescuentoPorcentaje = det.Porce_Descuento;
                    detalle.Descuento = Convert.ToInt32(det.Monto_Descuento);
                    detalle.SubDescuentos = new List<PlaceSoft.DTE.Engine.Documento.SubDescuento>();
                    var sbdcto = new PlaceSoft.DTE.Engine.Documento.SubDescuento();
                    sbdcto.TipoDescuento = PlaceSoft.DTE.Engine.Enum.ExpresionDinero.ExpresionDineroEnum.Porcentaje;
                    sbdcto.ValorDescuento = det.Porce_Descuento;
                    detalle.SubDescuentos.Add(sbdcto);
                }
                totalLinea = Math.Round(detalle.Cantidad * detalle.Precio) - detalle.Descuento;
                if (!string.IsNullOrEmpty(det.UnidadMedida))
                {
                    detalle.UnidadMedida = det.UnidadMedida;
                }
                /*Monto del item*/
                /*Recordar que debe restarse el descuento del detalle y sumarse el recargo*/
                detalle.MontoItem = (int)totalLinea;
                dte.Documento.Detalles.Add(detalle);
                contador++;
            }
            GenerateTotals(dte);
        }

        private void GenerateTotals(PlaceSoft.DTE.Engine.Documento.DTE dte)
        {
            calculosTotales(dte);

            if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE ==
               PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica)
            {

            }
            else if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE ==
                PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta)
            {

            }
            else
            {
                dte.Documento.Encabezado.Totales.MontoNeto = (int)Math.Round(neto, 0);
                dte.Documento.Encabezado.Totales.MontoExento = (int)Math.Round(netoExento, 0);
                if (neto != 0)
                {

                    dte.Documento.Encabezado.Totales.IVA = (int)Math.Round(iva, 0); ;
                }
            }
            dte.Documento.Encabezado.Totales.MontoExento = (int)Math.Round(netoExento, 0);
            dte.Documento.Encabezado.Totales.MontoTotal = (int)Math.Round(total, 0);
        }

        public void ReferenciasBoleta(PlaceSoft.DTE.Engine.Documento.DTE dte)
        {
            dte.Documento.Referencias = new List<PlaceSoft.DTE.Engine.Documento.Referencia>();
            var c = 1;
            /*Si estás en modo certificación, necesitas agregar esta referencia*/
            // REFERENCIA A SET DE PRUEBAS
            dte.Documento.Referencias.Add(new PlaceSoft.DTE.Engine.Documento.Referencia()
            {
                CodigoReferencia = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.NotSet,
                Numero = c,
                RazonReferencia = casoPruebas
            });
        }
        #endregion

        #region Region que Genera Facturas
        public PlaceSoft.DTE.Engine.Documento.DTE GenerateDTEFacturas(MySqlDataReader drCliente, string xfecha, string xActeco)
        {
            // DOCUMENTO
            var dte = new PlaceSoft.DTE.Engine.Documento.DTE();
            double montoExento = 0;


            if (drCliente.HasRows == true)
            {
                if (drCliente.Read())
                {
                    //
                    // DOCUMENTO - ENCABEZADO - CAMPO OBLIGATORIO
                    //Id = puede ser compuesto según tus propios requerimientos pero debe ser único         
                    dte.Documento.Id = idDte;

                    // DOCUMENTO - ENCABEZADO - IDENTIFICADOR DEL DOCUMENTO - CAMPOS OBLIGATORIOS
                    //TipoDTE = Se indica el tipo de documento. Esta API soporta:
                    //Factura Electrónica, Factura Electrónica Exenta, Boleta Electrónica, Nota de Crédito Electrónica, Nota de Débito Electrónica
                    dte.Documento.Encabezado.IdentificacionDTE.TipoDTE = tipo;
                    dte.Documento.Encabezado.IdentificacionDTE.FechaEmision = Convert.ToDateTime(xfecha);
                    dte.Documento.Encabezado.IdentificacionDTE.Folio = Folio;
                    //dte.Documento.Encabezado.IdentificacionDTE.FormaPago = PlaceSoft.DTE.Engine.Enum.FormaPago.FormaPagoEnum.Contado;
                    //dte.Documento.Encabezado.IdentificacionDTE.TipoDespacho = PlaceSoft.DTE.Engine.Enum.TipoDespacho.TipoDespachoEnum.EmisorACliente;
                    // PlaceSoft.DTE.Engine.Enum.TipoTraslado.TipoTrasladoEnum.OperacionConstituyeVenta;
                    if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.GuiaDespachoElectronica)
                    {
                        dte.Documento.Encabezado.IdentificacionDTE.TipoTraslado = emisor_tipo_traslado;
                    }

                    //dte.Documento.Encabezado.IdentificacionDTE.IndicadorServicioBoleta = PlaceSoft.DTE.Engine.Enum.IndicadorServicio.IndicadorServicioBoletaEnum.BoletaVentasYServicios;
                    //dte.Documento.Encabezado.IdentificacionDTE.IndicadorServicio = PlaceSoft.DTE.Engine.Enum.IndicadorServicio.IndicadorServicioEnum.BoletaVentasYServicios;

                    //DOCUMENTO - ENCABEZADO - EMISOR - CAMPOS OBLIGATORIOS          
                    dte.Documento.Encabezado.Emisor.Rut = emisor_rut;
                    dte.Documento.Encabezado.Emisor.RazonSocial = emisor_razon_social;
                    dte.Documento.Encabezado.Emisor.Giro = emisor_giro;
                    dte.Documento.Encabezado.Emisor.ComunaOrigen = emisor_comuna;
                    dte.Documento.Encabezado.Emisor.CiudadOrigen = emisor_ciudad;
                    dte.Documento.Encabezado.Emisor.DireccionOrigen = emisor_direccion;
                    dte.Documento.Encabezado.Emisor.ActividadEconomica.Add(emisor_acteco);


                    //if (xActeco.Length != 6)
                    //{
                    //    dte.Documento.Encabezado.Emisor.ActividadEconomica.Add(FuncionesClass.G_DTE_ACTECO[0]);
                    //}
                    //else
                    //{
                    //    dte.Documento.Encabezado.Emisor.ActividadEconomica.Add(xActeco);
                    //}
                    //DOCUMENTO - ENCABEZADO - RECEPTOR - CAMPOS OBLIGATORIOS
                    dte.Documento.Encabezado.Receptor.Rut = Convert.ToInt32(drCliente["rut"].ToString().Substring(0, 9)) + "-" + drCliente["rut"].ToString().Substring(9, 1);
                    dte.Documento.Encabezado.Receptor.Giro = drCliente["giro"].ToString();
                    dte.Documento.Encabezado.Receptor.RazonSocial = drCliente["nombre"].ToString();

                    dte.Documento.Encabezado.Receptor.Direccion = drCliente["direccion"].ToString();
                    dte.Documento.Encabezado.Receptor.Comuna = drCliente["comuna"].ToString();
                    dte.Documento.Encabezado.Receptor.Ciudad = drCliente["comuna"].ToString();
                    dte.Documento.Encabezado.Receptor.Contacto = drCliente["contacto"].ToString();

                    if (drCliente["email"].ToString() != "")
                    {
                        if (FuncionesClass.IsValidEmail(drCliente["email"].ToString()))
                        {
                            dte.Documento.Encabezado.Receptor.CorreoElectronico = drCliente["email"].ToString();
                        }

                    }

                }
            }
            return dte;
        }

        public void GenerateDetailsFacturas(PlaceSoft.DTE.Engine.Documento.DTE dte, MySqlDataReader xdrDetalles, string xFormaPago)
        {

            /************************ INICIO DE BUCLE QUE RECORRE LA GRILLA **********************/
            double dcto = 0;
            double DctoLinea = 0;
            double totalLinea = 0;
            double totalExento = 0;
            double tasaIVa = FuncionesClass.G_IVA;
            double total_bebidas = 0;
            double total_vinos = 0;
            double total_licores = 0;
            double total_cervezas = 0;
            double total_noazucar = 0;
            double total_harina = 0;
            double total_carne = 0;
            double BASE_IMPONIBLE = 0;
            double TOTAL_ADICIONALES = 0;
            double porce_impuesto = 0;
            double sub_neto = 0;
            double TOTAL_IMPUESTO = 0;

            double taza_bebidas = 0;
            double taza_vinos = 0;
            double taza_licores = 0;
            double taza_cervezas = 0;
            double taza_no_azucar = 0;
            double taza_harina = 0;
            double taza__carne = 0;

            double cantidad = 0;
            double descuento = 0;
            double precio = 0;
            double precioNeto = 0;
            double montoIva = 0;
            double montoTotal = 0;
            double totalNeto = 0;

            if (xdrDetalles.HasRows == true)
            {
                while (xdrDetalles.Read())
                {
                    REF_TIPO = xdrDetalles["ref_tipo"].ToString();
                    REF_NUMERO = xdrDetalles["ref_numero"].ToString();

                    var detalle = new PlaceSoft.DTE.Engine.Documento.Detalle();
                    /********************** INSTANCIA UN OBJETO PARA EL CODIGO *********************/
                    detalle.CodigosItem = new List<PlaceSoft.DTE.Engine.Documento.CodigoItem>();
                    var objCodigo = new PlaceSoft.DTE.Engine.Documento.CodigoItem();
                    objCodigo.TipoCodigo = "EAN13";
                    objCodigo.ValorCodigo = xdrDetalles["codigo"].ToString();
                    detalle.CodigosItem.Add(objCodigo);
                    /*******************************************************************************/
                    detalle.NumeroLinea = Convert.ToInt32(xdrDetalles["linea"].ToString());
                    detalle.Nombre = xdrDetalles["descripcion"].ToString();

                    if (xdrDetalles["descripcion"].ToString().Length > 60)
                    {
                        detalle.Nombre = xdrDetalles["descripcion"].ToString().Substring(0, 60);
                    }

                    precio = Convert.ToDouble(xdrDetalles["precio"]);
                    dcto = Convert.ToDouble(xdrDetalles["descuento"]);
                    cantidad = Convert.ToDouble(xdrDetalles["cantidad"]);

                    //detalle.Cantidad = Convert.ToDouble(xdrDetalles["art_cantidad"]);
                    // detalle.Precio = Convert.ToDouble(xdrDetalles["art_precio"]);
                    if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE != PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronicaExenta)
                    {
                        precioNeto = Math.Round(precio / (FuncionesClass.G_IVA / 100 + 1), 4);
                    }
                    else
                    {
                        //detalle.Precio = Convert.ToDouble(xdrDetalles["art_precio"]);
                        detalle.IndicadorExento = PlaceSoft.DTE.Engine.Enum.IndicadorFacturacionExencion.IndicadorFacturacionExencionEnum.NoAfectoOExento;
                    }

                    /*Monto del item*/
                    /*Recordar que debe restarse el descuento del detalle y sumarse el recargo*/


                    //DctoLinea = Math.Round((detalle.Precio * detalle.Cantidad) * (dcto / 100),5);
                    //totalLinea = Math.Round(detalle.Cantidad * detalle.Precio, 5) - DctoLinea;




                    /********************** REGION DE LOS IMPUESTOS ADICIONALES ********************/
                    if (xdrDetalles["impuesto"].ToString() != "00000" && dte.Documento.Encabezado.IdentificacionDTE.TipoDTE.ToString() != "FacturaElectronicaExenta")
                    {
                        //porce_impuesto = Convert.ToDouble(xdrDetalles["porce_impuesto"].ToString()) ;
                        //BASE_IMPONIBLE = Convert.ToDouble(Convert.ToDouble(xdrDetalles["art_precio"]) - DctoLinea) / ((IVA / 100 + 1) + (porce_impuesto));
                        //BASE_IMPONIBLE = Math.Round(BASE_IMPONIBLE, 4);
                        //sub_neto = sub_neto + totalLinea; 
                        //TOTAL_IMPUESTO = Math.Round((BASE_IMPONIBLE * porce_impuesto) * detalle.Cantidad, 4);


                        porce_impuesto = ((FuncionesClass.G_IVA / 100) + Convert.ToDouble(xdrDetalles["porcentajeimpuesto"].ToString()));
                        porce_impuesto = (porce_impuesto + 1);
                        precioNeto = Math.Round(precio / porce_impuesto, 4);
                        descuento = Math.Round(precioNeto * (dcto / 100), 4);
                        //detalle.Precio = precioNeto;
                        BASE_IMPONIBLE = Math.Round(precioNeto - descuento, 4);
                        BASE_IMPONIBLE = Math.Round(BASE_IMPONIBLE, 4);
                        sub_neto = sub_neto + Math.Round(BASE_IMPONIBLE * cantidad, 4);
                        TOTAL_IMPUESTO = Math.Round((BASE_IMPONIBLE * Convert.ToDouble(xdrDetalles["porcentajeimpuesto"].ToString())) * cantidad, 4);




                        /*********** BEBIDAS COD FAE 271 ****************/
                        if (xdrDetalles["impuesto"].ToString() == "00001")
                        {
                            taza_bebidas = Convert.ToDouble(xdrDetalles["taza"]);
                            total_bebidas = total_bebidas + TOTAL_IMPUESTO;

                            detalle.CodigoImpuestoAdicional = new List<PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum>();
                            var obj = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();
                            obj.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.BebidasAnalcoholicasYMineralesAltaAzucar;

                            detalle.CodigoImpuestoAdicional.Add(obj.TipoImpuesto);

                        }
                        /************* VINOS Y CHAMPAGNE 25 ************/
                        if (xdrDetalles["impuesto"].ToString() == "00002")
                        {
                            taza_vinos = Convert.ToDouble(xdrDetalles["taza"]);
                            total_vinos = total_vinos + TOTAL_IMPUESTO;

                            detalle.CodigoImpuestoAdicional = new List<PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum>();
                            var obj = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();
                            obj.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.Vinos;

                            detalle.CodigoImpuestoAdicional.Add(obj.TipoImpuesto);
                        }
                        /******************** LICORES 24 ***************/
                        if (xdrDetalles["impuesto"].ToString() == "00003")
                        {
                            taza_licores = Convert.ToDouble(xdrDetalles["taza"]);
                            total_licores = total_licores + TOTAL_IMPUESTO;
                            detalle.CodigoImpuestoAdicional = new List<PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum>();
                            var obj = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();
                            obj.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.Licores;

                            detalle.CodigoImpuestoAdicional.Add(obj.TipoImpuesto);
                        }
                        /******************** IMPUESTO HARINA 19 ***************/
                        if (xdrDetalles["impuesto"].ToString() == "00004")
                        {
                            taza_harina = Convert.ToDouble(xdrDetalles["taza"]);
                            total_harina = total_harina + TOTAL_IMPUESTO;
                            detalle.CodigoImpuestoAdicional = new List<PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum>();
                            var obj = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();
                            obj.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.IVAAnticipadoHarina;

                            detalle.CodigoImpuestoAdicional.Add(obj.TipoImpuesto);
                        }
                        /******************** IMPUESTO CARNE 18 ***************/
                        if (xdrDetalles["impuesto"].ToString() == "00005")
                        {
                            taza__carne = Convert.ToDouble(xdrDetalles["taza"]);
                            total_carne = total_carne + TOTAL_IMPUESTO;
                            detalle.CodigoImpuestoAdicional = new List<PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum>();
                            var obj = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();
                            obj.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.IVAAnticipadoCarne;

                            detalle.CodigoImpuestoAdicional.Add(obj.TipoImpuesto);
                        }
                        /******************** CERVEZAS 26 ***************/
                        if (xdrDetalles["impuesto"].ToString() == "00006")
                        {
                            taza_cervezas = Convert.ToDouble(xdrDetalles["taza"]);
                            total_cervezas = total_cervezas + TOTAL_IMPUESTO;

                            detalle.CodigoImpuestoAdicional = new List<PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum>();
                            var obj = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();
                            obj.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.Cervezas;

                            detalle.CodigoImpuestoAdicional.Add(obj.TipoImpuesto);

                        }
                        /******************** NO AZUCARADAS 27 ***************/
                        if (xdrDetalles["impuesto"].ToString() == "00007")
                        {
                            taza_no_azucar = Convert.ToDouble(xdrDetalles["taza"]);
                            total_noazucar = total_noazucar + TOTAL_IMPUESTO;

                            detalle.CodigoImpuestoAdicional = new List<PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum>();
                            var obj = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();
                            obj.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.BebidasAnalcoholicasYMinerales;

                            detalle.CodigoImpuestoAdicional.Add(obj.TipoImpuesto);
                        }




                    }// FIN IF IMPUESTOS ADICIONALES

                    else
                    {
                        if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE.ToString() != "FacturaElectronicaExenta")
                        {
                            descuento = Math.Round(precioNeto * (dcto / 100), 4);
                            BASE_IMPONIBLE = Math.Round(precioNeto - descuento, 4);
                            sub_neto = sub_neto + (BASE_IMPONIBLE * cantidad);
                        }
                        else
                        {
                            precioNeto = (precio);
                            descuento = Math.Round(precioNeto * (dcto / 100), 4);
                            BASE_IMPONIBLE = Math.Round(precioNeto - descuento, 4);
                            sub_neto = sub_neto + (BASE_IMPONIBLE * cantidad);
                        }

                    }


                    if (dcto > 0)
                    {
                        detalle.DescuentoPorcentaje = dcto;
                        detalle.Descuento = Convert.ToInt32(descuento * cantidad);
                        detalle.SubDescuentos = new List<PlaceSoft.DTE.Engine.Documento.SubDescuento>();
                        var sbdcto = new PlaceSoft.DTE.Engine.Documento.SubDescuento();
                        sbdcto.TipoDescuento = PlaceSoft.DTE.Engine.Enum.ExpresionDinero.ExpresionDineroEnum.Porcentaje;
                        sbdcto.ValorDescuento = dcto;
                        detalle.SubDescuentos.Add(sbdcto);
                    }

                    detalle.Precio = Math.Round(precioNeto, 4);
                    detalle.Cantidad = cantidad;


                    detalle.MontoItem = (int)Math.Round(((detalle.Precio - descuento) * cantidad), 0);
                    // detalle.MontoItem = (int)Math.Round(BASE_IMPONIBLE * detalle.Cantidad, 0, MidpointRounding.AwayFromZero);



                    dte.Documento.Detalles.Add(detalle);
                } // FIN WHILE DETALLES
            }// FIN IF HASROW

            string []pago = xFormaPago.Split(Convert.ToChar("-"));

            if (pago[0].ToString() != "20")
            {
                dte.Documento.Encabezado.IdentificacionDTE.FormaPago = PlaceSoft.DTE.Engine.Enum.FormaPago.FormaPagoEnum.Contado;
            }
            else
            {
                dte.Documento.Encabezado.IdentificacionDTE.FormaPago = PlaceSoft.DTE.Engine.Enum.FormaPago.FormaPagoEnum.Credito;
            }

            //dte.Documento.Encabezado.IdentificacionDTE.FechaVencimiento = Convert.ToDateTime(xdrDetalles["vencimiento"].ToString());



            /************************ SI EL DESCUENTO GLOBAL ES AFECTO **************************/

            //if (Convert.ToDouble(xdrDetalles["dctoglobalafecto"]) > 0)
            //{
            //    dte.Documento.DescuentosRecargos = new List<PlaceSoft.DTE.Engine.Documento.DescuentosRecargos>();
            //    var dctos = new PlaceSoft.DTE.Engine.Documento.DescuentosRecargos();
            //    dctos.Numero = 1;
            //    dctos.TipoMovimiento = PlaceSoft.DTE.Engine.Enum.TipoMovimiento.TipoMovimientoEnum.Descuento;
            //    dctos.TipoValor = PlaceSoft.DTE.Engine.Enum.ExpresionDinero.ExpresionDineroEnum.Porcentaje;
            //    //dctos.IndicadorExento = PlaceSoft.DTE.Engine.Enum.IndicadorExento.IndicadorExentoEnum.;
            //    dctos.Valor = Convert.ToInt32(Convert.ToDouble(xdrDetalles["dctoglobalafecto"]));
            //    dte.Documento.DescuentosRecargos.Add(dctos);
            //}
            /************************ SI EL DESCUENTO GLOBAL ES EXENTO **************************/



            /********************  REGION DE CODIGO PARA LOS TOTALES DEL DOCUMENTO *********************/
            total_bebidas = Math.Round(total_bebidas, 2);
            total_vinos = Math.Round(total_vinos, 2);
            total_licores = Math.Round(total_licores, 2);
            total_cervezas = Math.Round(total_cervezas, 2);
            total_noazucar = Math.Round(total_noazucar, 2);
            total_harina = Math.Round(total_harina, 2);
            total_carne = Math.Round(total_carne, 2);

            TOTAL_ADICIONALES = Math.Round(total_bebidas + total_vinos + total_licores + total_cervezas + total_noazucar + total_harina + total_carne, 5);
            TOTAL_ADICIONALES = Math.Round(TOTAL_ADICIONALES);

            if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.FacturaElectronicaExenta)
            {
                totalExento = sub_neto;
                dte.Documento.Encabezado.Totales.MontoExento = Convert.ToInt32(totalExento);
                iva = 0;
                tasaIVa = 0;
                sub_neto = 0;


            }
            else// SI LA FACTURA ES CON IVA
            {                
                iva = Math.Round((sub_neto * (FuncionesClass.G_IVA / 100 + 1)) - sub_neto, 4);

                /************* INSTANCIO UN LIST DE IMPUESTOS ************************/

                dte.Documento.Encabezado.Totales.ImpuestosRetenciones = new List<PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones>();
                if (total_bebidas > 0)
                {
                    //dte.Documento.Encabezado.Totales.ImpuestosRetenciones = new List<PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones>();
                    var impuesto = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();


                    impuesto.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.BebidasAnalcoholicasYMineralesAltaAzucar;
                    impuesto.TasaImpuesto = taza_bebidas;
                    total_bebidas = Math.Round(total_bebidas, 0, MidpointRounding.AwayFromZero);
                    impuesto.MontoImpuesto = Convert.ToInt32(total_bebidas);

                    dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Add(impuesto);
                }
                if (total_vinos > 0)
                {
                    // dte.Documento.Encabezado.Totales.ImpuestosRetenciones = new List<PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones>();
                    var impuesto = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();

                    impuesto.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.Vinos;
                    impuesto.TasaImpuesto = taza_vinos;
                    total_vinos = Math.Round(total_vinos, 0, MidpointRounding.AwayFromZero);
                    impuesto.MontoImpuesto = Convert.ToInt32(total_vinos);

                    dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Add(impuesto);
                }
                if (total_licores > 0)
                {
                    //dte.Documento.Encabezado.Totales.ImpuestosRetenciones = new List<PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones>();
                    var impuesto = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();

                    impuesto.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.Licores;
                    impuesto.TasaImpuesto = taza_licores;
                    total_licores = Math.Round(total_licores, 0, MidpointRounding.AwayFromZero);
                    impuesto.MontoImpuesto = Convert.ToInt32(total_licores);

                    dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Add(impuesto);
                }
                if (total_cervezas > 0)
                {
                    //dte.Documento.Encabezado.Totales.ImpuestosRetenciones = new List<PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones>();
                    var impuesto = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();

                    impuesto.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.Cervezas;
                    impuesto.TasaImpuesto = taza_cervezas;
                    total_cervezas = Math.Round(total_cervezas, 0, MidpointRounding.AwayFromZero);
                    impuesto.MontoImpuesto = Convert.ToInt32(total_cervezas);

                    dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Add(impuesto);
                }
                if (total_noazucar > 0)
                {
                    //dte.Documento.Encabezado.Totales.ImpuestosRetenciones = new List<PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones>();
                    var impuesto = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();

                    impuesto.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.BebidasAnalcoholicasYMinerales;
                    impuesto.TasaImpuesto = taza_no_azucar;
                    total_noazucar = Math.Round(total_noazucar, 0, MidpointRounding.AwayFromZero);
                    impuesto.MontoImpuesto = Convert.ToInt32(total_noazucar);

                    dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Add(impuesto);
                }
                if (total_harina > 0)
                {
                    //dte.Documento.Encabezado.Totales.ImpuestosRetenciones = new List<PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones>();
                    var impuesto = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();

                    impuesto.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.IVAAnticipadoHarina;
                    impuesto.TasaImpuesto = taza_harina;
                    total_harina = Math.Round(total_harina, 0, MidpointRounding.AwayFromZero);
                    impuesto.MontoImpuesto = Convert.ToInt32(total_harina);

                    dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Add(impuesto);
                }
                if (total_carne > 0)
                {
                    //dte.Documento.Encabezado.Totales.ImpuestosRetenciones = new List<PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones>();
                    var impuesto = new PlaceSoft.DTE.Engine.Documento.ImpuestosRetenciones();

                    impuesto.TipoImpuesto = PlaceSoft.DTE.Engine.Enum.TipoImpuesto.TipoImpuestoEnum.IVAAnticipadoCarne;
                    impuesto.TasaImpuesto = taza__carne;
                    total_carne = Math.Round(total_carne, 0, MidpointRounding.AwayFromZero);
                    impuesto.MontoImpuesto = Convert.ToInt32(total_carne);

                    dte.Documento.Encabezado.Totales.ImpuestosRetenciones.Add(impuesto);
                }

            }
            sub_neto = Math.Round(sub_neto, 4);
            iva = Math.Round(iva, 4);
            dte.Documento.Encabezado.Totales.MontoNeto = Convert.ToInt32(sub_neto);
            if (iva >= 1)
            {
                dte.Documento.Encabezado.Totales.IVA = Convert.ToInt32(iva);
                dte.Documento.Encabezado.Totales.TasaIVA = Convert.ToDouble(tasaIVa);
            }


            montoTotal = Math.Round(totalExento + sub_neto + iva + TOTAL_ADICIONALES, 5);

            dte.Documento.Encabezado.Totales.MontoTotal = Convert.ToInt32(montoTotal);


        }

        public void ReferenciaFacturas(PlaceSoft.DTE.Engine.Documento.DTE dte,string xLocal, string SERVER , string ROOT , string PASS )
        {
            dte.Documento.Referencias = new List<PlaceSoft.DTE.Engine.Documento.Referencia>();
            var c = 1;
            string TipoRef = "";
            string CODREF = "";
            string FechaRef = "";
            string tipoSII = "";
            string textoRef = "";
            string ref_fecha = "";
            string ref_tipo = this.REF_TIPO;
            string ref_numero = this.REF_NUMERO;
            Ventas venta;
            FechaRef = "";

            /*Ejemplo de referencia a una orden de compra*/
            if (ref_tipo != "" && ref_numero != "")
            {
                venta = new Ventas(FuncionesClass.G_CLIENTE_PREFIJO, SERVER, ROOT, PASS);
                FechaRef = venta.GetFechaReferencia(xLocal, REF_NUMERO, REF_TIPO);
                TipoRef = FuncionesClass.getTipoSIIByTipoDoc(ref_tipo);
                CODREF = "3";
                PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum refe =
                PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.AnulaDocumentoReferencia;
                if (CODREF == "1")
                {
                    refe = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.AnulaDocumentoReferencia;
                }
                if (CODREF == "2")
                {
                    refe = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.CorrigeTextoDocumentoReferencia;
                }
                if (CODREF == "3")
                {
                    refe = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.CorrigeMontos;
                }

                if (TipoRef == "33")
                {
                    dte.Documento.Referencias.Add(new PlaceSoft.DTE.Engine.Documento.Referencia()
                    {
                        CodigoReferencia = refe,
                        FechaDocumentoReferencia = Convert.ToDateTime(FechaRef),
                        //Folio de Referencia = Debe ir el folio de la factura o documento que estás refenciando                    
                        FolioReferencia = Convert.ToInt32(ref_numero).ToString(),
                        IndicadorGlobal = 0,
                        Numero = c,
                        RazonReferencia = "Devolucion de Mercaderia Nro " + ref_numero + " ",
                        TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.FacturaElectronica
                    });
                }
                if (TipoRef == "61")
                {

                    if (dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaDebitoElectronica)
                    {
                        refe = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.AnulaDocumentoReferencia;
                        textoRef = "Anula Documento de Referencia ";
                    }
                    else
                    {
                        refe = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.CorrigeMontos;
                        textoRef = "Corrige monto Referencia";
                    }

                    dte.Documento.Referencias.Add(new PlaceSoft.DTE.Engine.Documento.Referencia()
                    {
                        CodigoReferencia = refe,
                        FechaDocumentoReferencia = Convert.ToDateTime(FechaRef),
                        //Folio de Referencia = Debe ir el folio de la factura o documento que estás refenciando                    
                        FolioReferencia = Convert.ToInt32(ref_numero).ToString(),
                        IndicadorGlobal = 0,
                        Numero = c,
                        RazonReferencia = textoRef + " " + ref_numero + " ",
                        TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.NotaCreditoElectronica
                    });

                }
                if (TipoRef == "34")
                {

                    dte.Documento.Referencias.Add(new PlaceSoft.DTE.Engine.Documento.Referencia()
                    {
                        CodigoReferencia = refe,
                        FechaDocumentoReferencia = Convert.ToDateTime(FechaRef),
                        //Folio de Referencia = Debe ir el folio de la factura o documento que estás refenciando                    
                        FolioReferencia = Convert.ToInt32(ref_numero).ToString(),
                        IndicadorGlobal = 0,
                        Numero = c,
                        RazonReferencia = "Devolucion de Mercaderia Nro " + ref_numero.ToString() + " ",
                        TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.FacturaExentaElectronica
                    });
                }
                if (TipoRef == "56")
                {
                    dte.Documento.Referencias.Add(new PlaceSoft.DTE.Engine.Documento.Referencia()
                    {
                        CodigoReferencia = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.AnulaDocumentoReferencia,
                        FechaDocumentoReferencia = Convert.ToDateTime(FechaRef),
                        //Folio de Referencia = Debe ir el folio de la factura o documento que estás refenciando                    
                        FolioReferencia = Convert.ToInt32(ref_numero).ToString(),
                        IndicadorGlobal = 0,
                        Numero = c,
                        RazonReferencia = "Devolucion de Mercaderia Nro " + ref_numero.ToString() + " ",
                        TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.FacturaExentaElectronica
                    });
                }
                if (TipoRef == "39")
                {
                    dte.Documento.Referencias.Add(new PlaceSoft.DTE.Engine.Documento.Referencia()
                    {
                        CodigoReferencia = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.CorrigeMontos,
                        FechaDocumentoReferencia = Convert.ToDateTime(FechaRef),
                        //Folio de Referencia = Debe ir el folio de la factura o documento que estás refenciando                    
                        FolioReferencia = Convert.ToInt32(ref_numero).ToString(),
                        IndicadorGlobal = 0,
                        Numero = c,
                        RazonReferencia = "Devolucion de Mercaderia Nro " + ref_numero.ToString() + " ",
                        TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.BoletaElectronica
                    });
                }
                if (TipoRef == "41")
                {
                    dte.Documento.Referencias.Add(new PlaceSoft.DTE.Engine.Documento.Referencia()
                    {
                        CodigoReferencia = PlaceSoft.DTE.Engine.Enum.TipoReferencia.TipoReferenciaEnum.AnulaDocumentoReferencia,
                        FechaDocumentoReferencia = Convert.ToDateTime(FechaRef),
                        //Folio de Referencia = Debe ir el folio de la factura o documento que estás refenciando                    
                        FolioReferencia = Convert.ToInt32(ref_numero).ToString(),
                        IndicadorGlobal = 0,
                        Numero = c,
                        RazonReferencia = "Devolucion de Mercaderia Nro " + ref_numero.ToString() + " ",
                        TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoReferencia.BoletaExentaElectronica
                    });
                }

            }

        }


        #endregion

                      

        public string TimbrarYFirmarXMLDTE(PlaceSoft.DTE.Engine.Documento.DTE dte, string xPathCaf)
        {
            /*En primer lugar, el documento debe timbrarse con el CAF que descargas desde el SII, es simular
             * cuando antes debías ir con las facturas en papel para que te las timbraran */
            dte.Documento.Timbrar(
                EnsureExists
                ((int)dte.Documento.Encabezado.IdentificacionDTE.TipoDTE, dte.Documento.Encabezado.IdentificacionDTE.Folio, xPathCaf));
            /*Finalmente, el documento timbrado debe firmarse con el certificado digital*/
            /*Se debe entregar en el argumento del método Firmar, el "FriendlyName" o Nombre descriptivo del certificado*/
            /*Retorna el filePath donde estará el archivo XML timbrado y firmado, listo para ser enviado al SII*/
            return dte.Firmar(nombreCertificado);
        }

        public bool ValidateEnvio(string filePath, PlaceSoft.DTE.Security.Firma.Firma.TipoXML tipo)
        {
            string messageResult = string.Empty;
            if (PlaceSoft.DTE.Engine.XML.XmlHandler.ValidateWithSchema(filePath, out messageResult, PlaceSoft.DTE.Engine.XML.Schemas.EnvioDTE))
                if (PlaceSoft.DTE.Security.Firma.Firma.VerificarFirma(filePath, tipo))
                    return true;
                else
                    throw new Exception("NO SE HA PODIDO VERIFICAR LA FIRMA DEL ENVÍO");
            throw new Exception(messageResult);
        }

        public bool ValidateDTE(string filePath, PlaceSoft.DTE.Security.Firma.Firma.TipoXML tipo)
        {
            string messageResult = string.Empty;
            bool validaFirma = false ;
            if (PlaceSoft.DTE.Engine.XML.XmlHandler.ValidateWithSchema(filePath, out messageResult, PlaceSoft.DTE.Engine.XML.Schemas.DTE))
                validaFirma = PlaceSoft.DTE.Security.Firma.Firma.VerificarFirma(filePath, tipo);
                if (validaFirma == true)
                {
                    return true;
                }
                else
                {
                    throw new Exception("NO SE HA PODIDO VERIFICAR LA FIRMA DEL ENVÍO");
                    throw new Exception(messageResult);
                }
                    
        }

        private string EnsureExists(int tipoDTE, int folio, string xPathCaf)
        {
            var cafFile = string.Empty;
           
            foreach (var file in System.IO.Directory.GetFiles(xPathCaf))

                if (ParseName((new FileInfo(file)).Name, tipoDTE, folio))
                    cafFile = file;
            if (string.IsNullOrEmpty(cafFile))
                throw new Exception("NO HAY UN CÓDIGO DE AUTORIZACIÓN DE FOLIOS (CAF) ASIGNADO PARA ESTE TIPO DE DOCUMENTO (" + tipoDTE + ") QUE INCLUYA EL FOLIO REQUERIDO (" + folio + ").");
            return cafFile;
        }

        private static bool ParseName(string name, int tipoDTE, int folio)
        {
            try
            {
                var values = name.Substring(0, name.IndexOf('.')).Split('_');
                int tipo = Convert.ToInt32(values[0]);
                int desde = Convert.ToInt32(values[1]);
                int hasta = Convert.ToInt32(values[2]);
                return tipoDTE == tipo && desde <= folio && folio <= hasta;
            }
            catch { return false; }
        }
        
        #region Envio

 
        private PlaceSoft.DTE.Engine.Envio.EnvioDTE GenerarEnvioCliente(PlaceSoft.DTE.Engine.Documento.DTE dte, string dteXML)
        {
            var EnvioCustomer = new PlaceSoft.DTE.Engine.Envio.EnvioDTE();
            EnvioCustomer.SetDTE = new PlaceSoft.DTE.Engine.Envio.SetDTE();
            EnvioCustomer.SetDTE.DTEs.Add(dte);
            EnvioCustomer.SetDTE.dteXmls.Add(dteXML);
            EnvioCustomer.SetDTE.Caratula = new PlaceSoft.DTE.Engine.Envio.Caratula();
            EnvioCustomer.SetDTE.Caratula.FechaEnvio = DateTime.Now;
            /*Fecha de Resolución y Número de Resolución se averiguan en el sitio del SII según ambiente de producción o certificación*/
            EnvioCustomer.SetDTE.Caratula.FechaResolucion = DateTime.Now;
            EnvioCustomer.SetDTE.Caratula.NumeroResolucion = 80;

            EnvioCustomer.SetDTE.Caratula.RutEmisor = rutEmpresa;
            EnvioCustomer.SetDTE.Caratula.RutEnvia = rutCertificado;
            EnvioCustomer.SetDTE.Caratula.RutReceptor = dte.Documento.Encabezado.Receptor.Rut;
            /*Generalmente al cliente se le envía una sola factura, sin embargo si no es el caso, 
             se pueden agregar varias tal cual como está el método GenerarEnvioDTEToSII()*/
            EnvioCustomer.SetDTE.Caratula.SubTotalesDTE = new List<PlaceSoft.DTE.Engine.Envio.SubTotalesDTE>()
            {
                new PlaceSoft.DTE.Engine.Envio.SubTotalesDTE()
                {
                    Cantidad = 1,
                    TipoDTE = dte.Documento.Encabezado.IdentificacionDTE.TipoDTE
                }
            };

            return EnvioCustomer;
        }

      

        public long EnviarEnvioDTEToSII(string filePathEnvio, string serialKey, bool produccion)
        {
            string messageResult = string.Empty;
            long trackID = -1;
            int i;
            try
            {
                for (i = 1; i <= 5; i++)
                {
                    string rutEmisorNumero = rutEmpresa.Substring(0, rutEmpresa.Length - 2);
                    string rutEmisorDigito = rutEmpresa.Substring(rutEmpresa.Length - 1);
                    string rutEmpresaNumero = rutEmpresa.Substring(0, rutEmpresa.Length - 2);
                    string rutEmpresaDigito = rutEmpresa.Substring(rutEmpresa.Length - 1);
                    var responseEnvio = PlaceSoft.DTE.WS.EnvioDTE.EnvioDTE.Enviar(rutEmisorNumero, rutEmisorDigito, rutEmpresaNumero, rutEmpresaDigito, filePathEnvio, filePathEnvio, nombreCertificado, produccion, serialKey, out messageResult);

                    if (responseEnvio != null && string.IsNullOrEmpty(messageResult))
                    {
                        trackID = responseEnvio.TrackId;

                        /*Aquí pueden obtener todos los datos de la respuesta, tal como:
                         * Estado
                         * Fecha
                         * Archivo
                         * Glosa
                         * XML
                         * Entre otros*/
                        return trackID;
                    }
                }

                if (i == 5)
                    throw new Exception("SE HA ALCANZADO EL MÁXIMO NÚMERO DE INTENTOS: " + messageResult);
            }
            catch (Exception ex)
            {
                messageResult = ex.Message;
                return 0;
            }
            return 0;
        }
        
        private void calculosTotales(PlaceSoft.DTE.Engine.Documento.DTE dte)
        {
            try
            {
                foreach (var det in dte.Documento.Detalles)
                {
                    double div = 1.19;
                    var NetoUnitario = (det.Precio / div);
                    var Neto = (NetoUnitario * det.Cantidad);
                    double iva_aux = Neto * 0.19;
                    var IVA = Convert.ToInt32(Math.Round(iva_aux, 0));
                    var Total = (int)Math.Round(Neto + iva_aux, 0);
                    if (Total != det.MontoItem)
                    {
                        throw new Exception("Los totales no cuadran");
                    }
                    if (!(det.IndicadorExento == PlaceSoft.DTE.Engine.Enum.IndicadorFacturacionExencion.IndicadorFacturacionExencionEnum.NoAfectoOExento))
                    {
                        neto += Neto;
                    }
                    else
                    {
                        netoExento += det.Precio;
                    }
                }

                neto = Math.Round(neto, 0);
                iva = Math.Round(neto * 0.19, 0);
                total = netoExento + neto + iva;

                int nuevoNeto = (int)Math.Round(neto, 0);
                int nuevoExento = (int)Math.Round(netoExento, 0);
                int nuevoIVA = (int)Math.Round(iva, 0);
                int nuevoTotal = (int)Math.Round(total, 0);
            }
            catch { /*MessageBox.Show("Error. Hay una línea que debe ser borrada", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);*/ }
        }


        #endregion

        #region Boletas Electrónicas

          public PlaceSoft.DTE.Engine.RCOF.ConsumoFolios GenerarRCOF(DateTime xFecha,
                                                                      int xEmitidos39, int xAnulados39, int xUtilizados39, double xTotal39,
                                                                      List<string> xrangoUtilizado39, List<string> xrangoAnulado39,
                                                                      int xEmitidos41, int xAnulados41, int xUtilizados41, double xTotal41,
                                                                      List<string> xrangoUtilizado41, List<string> xrangoAnulado41,
                                                                      int xEmitidos61, int xAnulados61, int xUtilizados61, double xTotal61,
                                                                      List<string> xrangoUtilizado61, List<string> xrangoAnulado61
                                                                      )
        {
            var rcof = new PlaceSoft.DTE.Engine.RCOF.ConsumoFolios();

            rcof.DocumentoConsumoFolios.Caratula.FechaFinal = xFecha; //fechaInicio;
            rcof.DocumentoConsumoFolios.Caratula.FechaInicio = xFecha; // fechaFinal;
            rcof.DocumentoConsumoFolios.Caratula.FechaResolucion = this.fechaResolucion;
            rcof.DocumentoConsumoFolios.Caratula.NroResol = this.numero_resolucion;
            rcof.DocumentoConsumoFolios.Caratula.RutEmisor = rutEmpresa;
            rcof.DocumentoConsumoFolios.Caratula.RutEnvia = rutCertificado;
            rcof.DocumentoConsumoFolios.Caratula.SecEnvio = secuenciaEnvio;
            rcof.DocumentoConsumoFolios.Caratula.FechaEnvio = DateTime.Now;
            List<PlaceSoft.DTE.Engine.RCOF.Resumen> resumenes = new List<PlaceSoft.DTE.Engine.RCOF.Resumen>();


            if (xUtilizados39 > 0 || xUtilizados41 > 0)
            {

                /************  RANGOS DE FOLIOS BOLETAS 39 *******************/

                List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados> ListUtilizados39 = null;
                List<PlaceSoft.DTE.Engine.RCOF.RangoAnulados> ListAnulados39 = null;
                PlaceSoft.DTE.Engine.RCOF.RangoUtilizados RangoUsado39;
                PlaceSoft.DTE.Engine.RCOF.RangoAnulados RangoAnulado39;
                

                if (xrangoUtilizado39.Count > 0)
                {
                    ListUtilizados39 = new List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados>();
                }

                if (xrangoAnulado39.Count > 0)
                {
                    ListAnulados39 = new List<PlaceSoft.DTE.Engine.RCOF.RangoAnulados>();
                }
                                                             

                foreach (string var in xrangoUtilizado39)
                {
                    string[] cadena = var.ToString().Split(Convert.ToChar(","));
                    RangoUsado39 = new PlaceSoft.DTE.Engine.RCOF.RangoUtilizados();
                    RangoUsado39.Inicial = Convert.ToInt32(cadena[0]);
                    RangoUsado39.Final = Convert.ToInt32(cadena[1]);
                    ListUtilizados39.Add(RangoUsado39);

                }
                foreach (string var in xrangoAnulado39)
                {
                    string[] cadena = var.ToString().Split(Convert.ToChar(","));
                    RangoAnulado39 = new PlaceSoft.DTE.Engine.RCOF.RangoAnulados();
                    RangoAnulado39.Inicial = Convert.ToInt32(cadena[0]);
                    RangoAnulado39.Final = Convert.ToInt32(cadena[1]);
                    ListAnulados39.Add(RangoAnulado39);
                }


                /****************** GENERA LOS RANGOS DE BOLETAS EXENTAS 41 *******************/
                List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados> ListUtilizados41 = null;
                List<PlaceSoft.DTE.Engine.RCOF.RangoAnulados> ListAnulados41 = null;
                PlaceSoft.DTE.Engine.RCOF.RangoUtilizados RangoUsado41;
                PlaceSoft.DTE.Engine.RCOF.RangoAnulados RangoAnulado41;

                if (xrangoUtilizado41.Count > 0)
                {
                    ListUtilizados41 = new List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados>();
                }

                if (xrangoAnulado41.Count > 0)
                {
                    ListAnulados41 = new List<PlaceSoft.DTE.Engine.RCOF.RangoAnulados>();
                }

                foreach (string var in xrangoUtilizado41)
                {
                    string[] cadena = var.ToString().Split(Convert.ToChar(","));
                    RangoUsado41 = new PlaceSoft.DTE.Engine.RCOF.RangoUtilizados();
                    RangoUsado41.Inicial = Convert.ToInt32(cadena[0]);
                    RangoUsado41.Final = Convert.ToInt32(cadena[1]);
                    ListUtilizados41.Add(RangoUsado41);

                }
                foreach (string var in xrangoAnulado41)
                {
                    string[] cadena = var.ToString().Split(Convert.ToChar(","));
                    RangoAnulado41 = new PlaceSoft.DTE.Engine.RCOF.RangoAnulados();
                    RangoAnulado41.Inicial = Convert.ToInt32(cadena[0]);
                    RangoAnulado41.Final = Convert.ToInt32(cadena[1]);
                    ListAnulados41.Add(RangoAnulado41);
                }

                /****************** GENERA LOS RANGOS DE BOLETAS EXENTAS 41 *******************/
                List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados> ListUtilizados61 = null;
                List<PlaceSoft.DTE.Engine.RCOF.RangoAnulados> ListAnulados61 = null;
                PlaceSoft.DTE.Engine.RCOF.RangoUtilizados RangoUsado61;
                PlaceSoft.DTE.Engine.RCOF.RangoAnulados RangoAnulado61;

                if (xrangoUtilizado61.Count > 0)
                {
                    ListUtilizados61 = new List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados>();
                }

                if (xrangoAnulado61.Count > 0)
                {
                    ListAnulados61 = new List<PlaceSoft.DTE.Engine.RCOF.RangoAnulados>();
                }

                foreach (string var in xrangoUtilizado61)
                {
                    string[] cadena = var.ToString().Split(Convert.ToChar(","));
                    RangoUsado61 = new PlaceSoft.DTE.Engine.RCOF.RangoUtilizados();
                    RangoUsado61.Inicial = Convert.ToInt32(cadena[0]);
                    RangoUsado61.Final = Convert.ToInt32(cadena[1]);
                    ListUtilizados61.Add(RangoUsado61);

                }
                foreach (string var in xrangoAnulado61)
                {
                    string[] cadena = var.ToString().Split(Convert.ToChar(","));
                    RangoAnulado61 = new PlaceSoft.DTE.Engine.RCOF.RangoAnulados();
                    RangoAnulado61.Inicial = Convert.ToInt32(cadena[0]);
                    RangoAnulado61.Final = Convert.ToInt32(cadena[1]);
                    ListAnulados61.Add(RangoAnulado61);
                }
                                                                                           


                double thisNeto39 = Math.Round((xTotal39) / 1.19, 4);
                double thisIVA39 = Math.Round((xTotal39 ) - thisNeto39, 4);
                double thisTotal39 = Math.Round(thisNeto39 + thisIVA39 + xTotal41, 4);
                thisNeto39 = Math.Round(thisNeto39, 0, MidpointRounding.AwayFromZero);
                thisIVA39 = Math.Round(thisIVA39, 0, MidpointRounding.AwayFromZero);
                double utilizados39 = (xEmitidos39 + xAnulados39);
                resumenes.Add(new PlaceSoft.DTE.Engine.RCOF.Resumen
                {
                    FoliosAnulados = xAnulados39.ToString(),
                    FoliosEmitidos = xEmitidos39.ToString(),
                    FoliosUtilizados = utilizados39.ToString(),
                    MntExento = (int)0,
                    MntIva = (int)thisIVA39,
                    MntNeto = (int)thisNeto39,
                    MntTotal = (int)xTotal39,
                    TasaIVA = 19,
                    TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica,
                    RangoUtilizados = ListUtilizados39,
                    RangoAnulados = ListAnulados39
                });
                /*************************** FIN RESUMEN BOLETAS EFECTAS *******************************/

                /*************************** INICIO BOLETAS ELECTRÓNICAS EXENTAS(41) *******************/

                int Total41 = (int)xTotal41; // dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta).Sum(x => x.Documento.Encabezado.Totales.MontoTotal);
                double thisNeto41 = Math.Round((xTotal41) / 1.19, 4);
                double thisIVA41 = Math.Round((xTotal41) - thisNeto41, 4);
                double thisTotal41 = Math.Round(thisNeto41 + thisIVA41 + xTotal41, 4);
                thisNeto41 = Math.Round(thisNeto41, 0, MidpointRounding.AwayFromZero);
                thisIVA41 = Math.Round(thisIVA41, 0, MidpointRounding.AwayFromZero);
                double utilizados41 = (xEmitidos41 + xAnulados41);
                if (Total41 > 0)
                {

                    resumenes.Add(new PlaceSoft.DTE.Engine.RCOF.Resumen
                    {
                        FoliosAnulados = xAnulados41.ToString(),
                        FoliosEmitidos = xEmitidos41.ToString(),// dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta).Count().ToString(),
                        FoliosUtilizados = xUtilizados41.ToString(), // dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta).Count().ToString(),
                        MntExento = (int)xTotal41,
                        MntTotal = (int)Total41,
                        TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta,
                        RangoUtilizados = ListUtilizados41,
                        RangoAnulados = ListAnulados41
                    });
                }
                /*************************** FIN RESUMEN BOLETAS ELECTRÓNICAS EXENTAS(41) ***************************/


                /***************** TOTAL NOTAS ELECTRÓNICAS CON REFERENCIA BOLETAS ELECTRÓNICAS(61) *****************/
                int desde = 0;
                int hasta = 0;

                int totalesNC = (int)xTotal61; //dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica).Sum(x => x.Documento.Encabezado.Totales.MontoTotal);
                if (totalesNC > 0)
                {
                    int Total61 = (int)xTotal61; // dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronicaExenta).Sum(x => x.Documento.Encabezado.Totales.MontoTotal);
                    double thisNeto61 = Math.Round((xTotal61) / 1.19, 4);
                    double thisIVA61 = Math.Round((xTotal61) - thisNeto61, 4);
                    double thisTotal61 = Math.Round(thisNeto61 + thisIVA61 + xTotal61, 4);
                    thisNeto61 = Math.Round(thisNeto61, 0, MidpointRounding.AwayFromZero);
                    thisIVA61 = Math.Round(thisIVA61, 0, MidpointRounding.AwayFromZero);
                    double utilizados61 = (xEmitidos61 + xAnulados61);


                    resumenes.Add(new PlaceSoft.DTE.Engine.RCOF.Resumen
                    {
                        FoliosAnulados = "0",
                        FoliosEmitidos = xEmitidos61.ToString(),// dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica).Count().ToString(),
                        FoliosUtilizados = xUtilizados61.ToString(), // dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica).Count().ToString(),
                        MntExento = 0,
                        MntIva = (int) thisIVA61,
                        MntNeto = (int) thisNeto61,
                        MntTotal = (int) thisNeto61,
                        TasaIVA = 19,
                        TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica,
                        RangoUtilizados = ListUtilizados61,
                        RangoAnulados = ListAnulados61
                    });

                }


            }
            else // SI NO HUBO MOVIMIENTO SE MANDA UN RCOF VACIO List<string> Telefono 
            {
                resumenes.Add(new PlaceSoft.DTE.Engine.RCOF.Resumen
                {
                    TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica,
                    MntTotal = 0,
                    TasaIVA = 19,
                    FoliosEmitidos = "0",
                    FoliosAnulados = "0",
                    FoliosUtilizados = "0",
                    RangoUtilizados = new List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados>()
                     {
                        null
                     },
                    RangoAnulados = new List<PlaceSoft.DTE.Engine.RCOF.RangoAnulados>()
                     {
                            null
                      }
                });
            }



            rcof.DocumentoConsumoFolios.Resumen = resumenes;

            return rcof;
        }
        private List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados> GetRangosDocumentos(List<PlaceSoft.DTE.Engine.Documento.DTE> dtes, PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType tipo)
        {
            // List<PlaceSoft.DTE.Engine.RCOF.Resumen> resumenes = new List<PlaceSoft.DTE.Engine.RCOF.Resumen>();
            List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados> rangos =  new List<PlaceSoft.DTE.Engine.RCOF.RangoUtilizados>();
            PlaceSoft.DTE.Engine.RCOF.RangoUtilizados rango;
            List<int> folios  = new List<int>();
            int inicial = 0;
            int final = 0;
            int anterior = 0;
            
            bool corta = false;
            foreach (PlaceSoft.DTE.Engine.Documento.DTE dte in dtes )
            {
                if(dte.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.NotaCreditoElectronica)
                {
                    //rango = new PlaceSoft.DTE.Engine.RCOF.RangoUtilizados();
                    //rango.Inicial = dte.Documento.Encabezado.IdentificacionDTE.Folio;
                    //rango.Final = dte.Documento.Encabezado.IdentificacionDTE.Folio;
                    //rangos.Add(rango);
                    folios.Add(dte.Documento.Encabezado.IdentificacionDTE.Folio);
               
                }
            }
                                                  

            int i = 0;
            int count = 0;
            for (i=0; i <= folios.Count-1 ; i++)
            {
                if(count == 0)
                {
                    inicial = folios[i];
                }                
                final = folios[i];
                if( i == folios.Count-1)
                {
                    rango = new PlaceSoft.DTE.Engine.RCOF.RangoUtilizados();
                    rango.Inicial = inicial;
                    rango.Final = final;
                    rangos.Add(rango);
                    corta = false;
                    count = 0;
                }
                else
                {
                    if ((folios[i + 1] - final) > 1)
                    {
                        final = folios[i];
                        corta = true;
                    }
                }            

                //anterior = folios[i];
                count++;
                if (corta == true )
                {
                    rango = new PlaceSoft.DTE.Engine.RCOF.RangoUtilizados();
                    rango.Inicial = inicial;
                    rango.Final = final;
                    rangos.Add(rango);
                    corta = false;
                    count = 0;
                }
               
            }
            return rangos;
        }


        public PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.LibroBoletas GenerateLibroBoletas(List<PlaceSoft.DTE.Engine.Documento.DTE> dtes, string xPeriodo)
        {
            var libro = new PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.LibroBoletas();

            libro.EnvioLibro = new PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.EnvioLibro();

            /*Datos para confeccion de caratula*/
            string periodoTributario =xPeriodo;
            DateTime fechaResolucion = FuncionesClass.G_DTE_FECHARESOLUCION;
            int nResolucion = FuncionesClass.G_DTE_NUMERO_RESOLUCION;
            /*Fecha de Resolución y Número de Resolución se averiguan en el sitio del SII según ambiente de producción o certificación*/
            /*El tipo de libro debe ser "Especial" cuando se trata del set de pruebas*/
            /*El folio de notificacion lo entrega el SII al momento de solicitar el libro, para el set de pruebas no es necesario agregarlo*/
            libro.EnvioLibro.Caratula = new PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.Caratula
            {
                RutEmisor = rutEmpresa,
                RutEnvia = rutCertificado,
                PeriodoTributario = periodoTributario,
                FechaResolucion = fechaResolucion,
                NumeroResolucion = nResolucion,
                TipoLibro = PlaceSoft.DTE.Engine.Enum.TipoLibro.TipoLibroEnum.Especial,
                TipoEnvio = PlaceSoft.DTE.Engine.Enum.TipoEnvioLibro.TipoEnvioLibroEnum.Total
            };

            libro.EnvioLibro.ResumenPeriodo = new PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.ResumenPeriodo();
            libro.EnvioLibro.ResumenPeriodo.TotalesPeriodo = new List<PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.TotalPeriodo>();

            /*Se agregar un Total Periodo por cada tipo de documento. Boletas electrónicas exentas y afectas*/
            /*Boletas electronicas*/
            int totalNeto = dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica).Sum(x => x.Documento.Encabezado.Totales.MontoNeto);
            int totalIVA = dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica).Sum(x => x.Documento.Encabezado.Totales.IVA);
            int totalExento = dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica).Sum(x => x.Documento.Encabezado.Totales.MontoExento);
            int totalTotal = dtes.Where(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica).Sum(x => x.Documento.Encabezado.Totales.MontoTotal);

            libro.EnvioLibro.ResumenPeriodo.TotalesPeriodo.Add(new PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.TotalPeriodo()
            {
                TipoDocumento = PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro.BoletaElectronica,
                CantidadDocumentosAnulados = 0,
                TotalesServicio = new List<PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.TotalServicio>()
                {
                    new PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.TotalServicio()
                    {
                        CantidadDocumentos = dtes.Count(x=>x.Documento.Encabezado.IdentificacionDTE.TipoDTE == PlaceSoft.DTE.Engine.Enum.TipoDTE.DTEType.BoletaElectronica),
                        TasaIVA = 19,
                        TotalIVA = totalIVA,
                        TotalNeto = totalNeto,
                        TotalExento = totalExento,
                        TotalTotal = totalTotal,
                        TipoServicio = (int)PlaceSoft.DTE.Engine.Enum.IndicadorServicio.IndicadorServicioEnum.BoletaVentasYServicios
                    }
                }
            });

            /*Se agregan los dtes del libro*/
            libro.EnvioLibro.Detalles = new List<PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.Detalle>();
            foreach (var dte in dtes)
                libro.EnvioLibro.Detalles.Add(new PlaceSoft.DTE.Engine.InformacionElectronica.LBoletas.Detalle()
                {
                    TipoDocumento = (PlaceSoft.DTE.Engine.Enum.TipoDTE.TipoDocumentoLibro)dte.Documento.Encabezado.IdentificacionDTE.TipoDTE,
                    FolioDocumento = dte.Documento.Encabezado.IdentificacionDTE.Folio,
                    FechaEmision = dte.Documento.Encabezado.IdentificacionDTE.FechaEmision,
                    MontoExento = dte.Documento.Encabezado.Totales.MontoExento,
                    MontoTotal = dte.Documento.Encabezado.Totales.MontoTotal,
                    RutCliente = dte.Documento.Encabezado.Receptor.Rut
                });

            return libro;
        }


        public PlaceSoft.DTE.Engine.Envio.EnvioDTE GenerarEnvioCliente(List<PlaceSoft.DTE.Engine.Documento.DTE> dtes, List<string> xmlDtes,
                                                                  string xRutEmpresa, string xRutEnvia, string xRutRecibe, string xFechaResol, int NroResol)
        {
            var EnvioSII = new PlaceSoft.DTE.Engine.Envio.EnvioDTE();
            EnvioSII.SetDTE = new PlaceSoft.DTE.Engine.Envio.SetDTE();
            EnvioSII.SetDTE.Id = "ENVIO_CLIENTE_" + DateTime.Now.ToString("ddMMyyyyHHmmss");
            /*Es necesario agregar en el envío, los objetos DTE como sus respectivos XML en strings*/
            foreach (var a in dtes)
                EnvioSII.SetDTE.DTEs.Add(a);
            foreach (var a in xmlDtes)
                EnvioSII.SetDTE.dteXmls.Add(a);

            EnvioSII.SetDTE.Caratula = new PlaceSoft.DTE.Engine.Envio.Caratula();
            EnvioSII.SetDTE.Caratula.FechaEnvio = DateTime.Now;
            /*Fecha de Resolución y Número de Resolución se averiguan en el sitio del SII según ambiente de producción o certificación*/
            EnvioSII.SetDTE.Caratula.FechaResolucion = Convert.ToDateTime( xFechaResol);
            EnvioSII.SetDTE.Caratula.NumeroResolucion = NroResol;

            EnvioSII.SetDTE.Caratula.RutEmisor = xRutEmpresa;
            EnvioSII.SetDTE.Caratula.RutEnvia = xRutEnvia;
            EnvioSII.SetDTE.Caratula.RutReceptor = xRutRecibe; //Este es el RUT del SII
            EnvioSII.SetDTE.Caratula.SubTotalesDTE = new List<PlaceSoft.DTE.Engine.Envio.SubTotalesDTE>();

            /*En la carátula del envío, se debe indicar cuantos documentos de cada tipo se están enviando*/
            var tipos = EnvioSII.SetDTE.DTEs.GroupBy(x => x.Documento.Encabezado.IdentificacionDTE.TipoDTE);
            foreach (var a in tipos)
            {
                EnvioSII.SetDTE.Caratula.SubTotalesDTE.Add(new PlaceSoft.DTE.Engine.Envio.SubTotalesDTE()
                {
                    Cantidad = a.Count(),
                    TipoDTE = a.ElementAt(0).Documento.Encabezado.IdentificacionDTE.TipoDTE
                });
            }

            return EnvioSII;
        }
        public string FirmarEnvioDTE(PlaceSoft.DTE.Engine.Envio.EnvioDTE env)
        {
            var filePathEnvio = string.Empty;
            string messageResult = string.Empty;
            filePathEnvio = env.Firmar(nombreCertificado, true);
            if (!ValidateEnvio(filePathEnvio, PlaceSoft.DTE.Security.Firma.Firma.TipoXML.Envio))
            {
                throw new Exception("NO SE PUDO VALIDAR EL SCHEMA DEL ENVIO GENERADO.");
            }
            return filePathEnvio;
        }
        //public string FirmarEnvioDTE(PlaceSoft.DTE.Engine.Envio.EnvioDTE env)
        //{
        //    var filePathEnvio = string.Empty;
        //    string messageResult = string.Empty;
        //    filePathEnvio = env.Firmar(nombreCertificado, true);
        //    if (!ValidateEnvio(filePathEnvio, PlaceSoft.DTE.Security.Firma.Firma.TipoXML.Envio))
        //    {
        //        throw new Exception("NO SE PUDO VALIDAR EL SCHEMA DEL ENVIO GENERADO.");
        //    }
        //    return filePathEnvio;
        //}

        #endregion

    }
}
