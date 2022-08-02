Attribute VB_Name = "Ayudas"
Option Explicit

Public cabezas As Variant
Private help As CAyuda

    Public Sub ayudaCliente(ByRef txt As TextBox, ByRef suc As TextBox, ByRef lbldv As Label)
        Dim cad As String
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestroclientes"
        mensajeAyuda = "Ayuda de Clientes"
        camposAyuda = Array("CONCAT(LEFT(rut,3), '.', MID(rut,4,3), '.', MID(rut,7,3), '-', RIGHT(rut,1), '  ', sucursal)", "nombre", "giro", "cupodirecto")
        cabezasAyuda = Array("rut", "nombre", "giro", "cupodirecto")
        largoAyuda = Array("15n", "50s", "15s", "10n")
        condicionAyuda = "no"
        cantidadAyuda = 4
        txt.MaxLength = 16
        Call Mayuda.cargaAyuda(txt)
        cad = txt.text
        suc.text = Right(cad, 1)
        cad = Replace(cad, ".", "")
        cad = Replace(cad, "-", "")
        txt.MaxLength = 10
        txt.text = cad
        lbldv.Caption = Right(txt.text, 1)
        txt.MaxLength = 9
        txt.text = txt.text
        txt.SetFocus
    End Sub
    
    Public Sub ayudaClienteSin(ByRef txt As TextBox, ByRef lbldv As Label)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestroclientes"
        mensajeAyuda = "Ayuda de Clientes"
        camposAyuda = Array("rut", "nombre", "giro")
        cabezasAyuda = Array("rut", "nombre", "giro")
        largoAyuda = Array("10n", "50s", "15s")
        condicionAyuda = "no"
        cantidadAyuda = 3
        txt.MaxLength = 10
        Call Mayuda.cargaAyuda(txt)
        lbldv.Caption = Right(txt.text, 1)
        txt.MaxLength = 9
        txt.text = txt.text
    End Sub
     Public Sub ayudaTecnico(ByRef txt As TextBox, ByRef lbldv As Label)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestroserviciotecnico"
        mensajeAyuda = "Ayuda de Tecnico"
        camposAyuda = Array("rut", "nombre")
        cabezasAyuda = Array("rut", "nombre")
        largoAyuda = Array("10n", "50s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        txt.MaxLength = 10
        Call Mayuda.cargaAyuda(txt)
        lbldv.Caption = Right(txt.text, 1)
        txt.MaxLength = 9
        txt.text = txt.text
    End Sub
    
        Public Sub ayudagastoscobranza(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_cobranza_gastos"
        mensajeAyuda = "Ayuda de GASTOS"
        camposAyuda = Array("codigo", "descripcion")
        cabezasAyuda = Array("Codigo", "Descripción")
        largoAyuda = Array("15s", "45s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        txt.MaxLength = 2
        Call Mayuda.cargaAyuda(txt)
        txt.text = txt.text
    End Sub
    Public Sub ayudaUsuarios(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = clientesistema & "auditoria.segu_usuarios"
        mensajeAyuda = "Ayuda de Usuarios"
        camposAyuda = Array("usuario", "nombre")
        cabezasAyuda = Array("usuario", "nombre")
        largoAyuda = Array("30s", "30s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        txt.MaxLength = 50
        Call Mayuda.cargaAyuda(txt)
        txt.text = txt.text
    End Sub
    
    Public Sub ayudaClienteDeuda(ByRef txt As TextBox, ByRef lbldv As Label)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestroclientes AS mc INNER JOIN " & baseVentas & empresaActiva & ".sv_documentos_cobranza_" & empresaActiva & " AS dc ON mc.rut = dc.rut"
        mensajeAyuda = "Ayuda de Clientes"
        camposAyuda = Array("mc.rut", "mc.nombre", "mc.giro")
        cabezasAyuda = Array("rut", "nombre", "giro")
        largoAyuda = Array("10n", "50s", "15s")
        condicionAyuda = "no"
        cantidadAyuda = 3
        txt.MaxLength = 10
        Call Mayuda.cargaAyuda(txt)
        lbldv.Caption = Right(txt.text, 1)
        txt.MaxLength = 9
        txt.text = txt.text
    End Sub
    
    Public Sub ayudaCaja(ByRef txt As TextBox, ByVal emp As String)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrodecajas"
        mensajeAyuda = "Ayuda de Cajas"
        camposAyuda = Array("numero", "descripcion", "local")
        cabezasAyuda = Array("codigo", "descripcion", "local")
        largoAyuda = Array("7n", "50s", "5n")
        condicionAyuda = "local = '" & emp & "'"
        cantidadAyuda = 3
        Call Mayuda.cargaAyuda(txt)
    End Sub
    
    Public Sub ayudaBodega(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = basedatos & rubro
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "r_maestrobodegas_" & rubro
        mensajeAyuda = "Ayuda de Bodegas"
        camposAyuda = Array("codigobodega", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("7n", "50s")
        condicionAyuda = "local = '" & empresaActiva & "'"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
    
    
    Public Sub ayudaClienteCredito(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas & empresaActiva
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = baseVentas & ".sv_maestroclientes AS mc INNER JOIN sv_credito_cabeza AS cc ON mc.rut = cc.rut INNER JOIN sv_credito_detalle AS cd ON cc.local = cd.local AND cc.folio = cd.folio"
        mensajeAyuda = "Ayuda de Clientes"
        camposAyuda = Array("mc.rut", "mc.nombre", "cc.folio")
        cabezasAyuda = Array("rut", "nombre", "folio")
        largoAyuda = Array("13n", "50s", "13n")
        condicionAyuda = "cc.local = '" & empresaActiva & "' AND cd.montocuota <> abonocuota"
        cantidadAyuda = 3
        Call Mayuda.cargaAyuda(txt)
    End Sub
    
    Public Sub ayudaDocumentoCliente(ByRef txt As TextBox, ByVal TIPO As String, ByVal rut As String)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas & empresaActiva
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_documento_pagos_" + empresaActiva + " AS dp INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dp.local = dc.local AND dp.tipo = dc.tipo AND dp.numero = dc.numero"
        mensajeAyuda = "Ayuda de Clientes"
        camposAyuda = Array("dp.numero", "dp.monto", "DATE_FORMAT(dc.fecha,'%d-%m-%Y')")
        cabezasAyuda = Array("numero", "monto credito", "fecha")
        largoAyuda = Array("13n", "13n", "13n")
        condicionAyuda = "dp.local = '" & empresaActiva & "' AND dp.tipo = '" & TIPO & "' AND dp.rut = '" & rut & "' AND tipopago = '6' AND dc.nula = 'N' AND dc.total <> dc.abono"
        cantidadAyuda = 3
        Call Mayuda.cargaAyuda(txt)
    End Sub
    
    Public Sub ayudaFolio(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas & empresaActiva
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = baseVentas & ".sv_maestroclientes AS mc INNER JOIN sv_credito_cabeza AS cc ON mc.rut = cc.rut INNER JOIN sv_credito_detalle AS cd ON cc.local = cd.local AND cc.folio = cd.folio"
        mensajeAyuda = "Ayuda de Clientes"
        camposAyuda = Array("cc.folio", "mc.nombre", "mc.rut")
        cabezasAyuda = Array("folio", "nombre", "rut")
        largoAyuda = Array("13n", "50s", "13n")
        condicionAyuda = "cc.local = '" & empresaActiva & "' AND cd.montocuota <> abonocuota"
        cantidadAyuda = 3
        Call Mayuda.cargaAyuda(txt)
    End Sub

    Public Sub ayudaVendedores(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrovendedores"
        mensajeAyuda = "Ayuda de Vendedores"
        camposAyuda = Array("rut", "nombre")
        cabezasAyuda = Array("Rut", "nombre")
        largoAyuda = Array("7n", "50s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub

    Public Sub ayudaZonas(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrozonas"
        mensajeAyuda = "Ayuda de Zonas"
        camposAyuda = Array("codigozona", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("10n", "50s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub

    Public Sub ayudaBancos(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrobancos"
        mensajeAyuda = "Ayuda de Bancos"
        camposAyuda = Array("codigobanco", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("10n", "50s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
     Public Sub ayudavivienda(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrotipovivienda"
        mensajeAyuda = "Ayuda Tipo Vivienda"
        camposAyuda = Array("codigo", "nombre")
        cabezasAyuda = Array("Codigo", "Nombre")
        largoAyuda = Array("10n", "50s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
    
    Public Sub ayudaEmpresa(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = basedatos
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "g_maestroempresas"
        mensajeAyuda = "Ayuda de Locales"
        camposAyuda = Array("codigo", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("10n", "50s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
        Public Sub ayudaEmpresaremu(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = basedatos
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = clientesistema & "remu" & ".maestroempresas"
        mensajeAyuda = "Ayuda de Empresas"
        camposAyuda = Array("codigo", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("10n", "50s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub

    Public Sub ayudaLocalesRubro(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = basedatos
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "g_maestroempresas"
        mensajeAyuda = "Ayuda de Locales"
        camposAyuda = Array("codigo", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("10n", "50s")
        condicionAyuda = "rubro = '" & rubro & "'"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
    
    Public Sub ayudaComprobante(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas & empresaActiva
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_prorroga AS p LEFT JOIN " & baseVentas & ".sv_maestroclientes AS mc ON p.rut = mc.rut"
        mensajeAyuda = "Ayuda de Comprobantes de Prorroga"
        camposAyuda = Array("p.numero", "mc.nombre", "p.rut")
        cabezasAyuda = Array("numero", "nombre", "rut")
        largoAyuda = Array("10n", "50s", "10n")
        condicionAyuda = "no"
        cantidadAyuda = 3
        Call Mayuda.cargaAyuda(txt)
    End Sub

    Public Sub ayudaEgreso(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas & empresaActiva
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_egresoscaja_" + rubro
        mensajeAyuda = "Ayuda de Comprobantes de Egreso"
        camposAyuda = Array("numero", "tipo", "recibido")
        cabezasAyuda = Array("numero", "tipo", "recibido por")
        largoAyuda = Array("10n", "10s", "30s")
        condicionAyuda = "no"
        cantidadAyuda = 3
        Call Mayuda.cargaAyuda(txt)
    End Sub
    
    Public Sub ayudaTipoEgreso(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrotipoegresoscaja"
        mensajeAyuda = "Ayuda de Tipos de Egreso"
        camposAyuda = Array("codigo", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("8n", "25s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub

    Public Sub ayudaTipoPago(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_tiposdepagoclientes"
        mensajeAyuda = "Ayuda de Tipos de Pago"
        camposAyuda = Array("codigo", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("8n", "25s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub

 Public Sub ayudaTipoPagopago(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrotipopago_pago"
        mensajeAyuda = "Ayuda de Tipos de Pago"
        camposAyuda = Array("codigo", "nombre")
        cabezasAyuda = Array("codigo", "nombre")
        largoAyuda = Array("8n", "25s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
    Public Sub ayudacajera(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrocajeras"
        mensajeAyuda = "Ayuda de Cajeros"
        camposAyuda = Array("rut", "nombre")
        cabezasAyuda = Array("Rut", "Nombre")
        largoAyuda = Array("8n", "25s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
    Public Sub ayudaServicio(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestroserviciotecnico"
        mensajeAyuda = "Ayuda de Tecnicos"
        camposAyuda = Array("rut", "nombre")
        cabezasAyuda = Array("Rut", "Nombre")
        largoAyuda = Array("8n", "25s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
    
  Public Sub ayudaProducto2(lista As Grid, ByRef txt As TextBox)
        Dim tipre As String
        servidorAyuda = servidor
        basedatosAyuda = basedatos & rubro
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "r_maestroproductos_fijo_" & rubro & " AS mpf LEFT JOIN r_maestroproductos_precios_" & rubro & " AS mpp ON mpf.codigobarra = mpp.codigo LEFT JOIN r_maestroproductos_stock_" & rubro & " AS mps ON mpf.codigobarra = mps.codigo"
        mensajeAyuda = "Ayuda de Productos"
        camposAyuda = Array("mpf.codigobarra", "mpf.descripcion", "mpf.referenciaproveedor", "CONCAT('$ ', IFNULL(FORMAT(mpp.preciopuntoventa,0),0))", "IFNULL(mps.stockactual,0)")
        cabezasAyuda = Array("codigo", "descripcion", "referencia", "precio", "stock")
        largoAyuda = Array("13n", "30s", "20s", "15n", "5n")
        tipre = "01"
        
        condicionAyuda = "mpp.local = '" + empresaActiva + "' AND mps.local = '" & empresaActiva & "' AND mps.bodega = '" & "00" & "' AND mpp.codigoprecio = '" & tipre & "' AND mpf.descontinuado='0'"
        
        Rem condicionAyuda = "no"
        cantidadAyuda = 5
        Call Mayuda.cargaAyuda(txt)
        lista.Cell(lista.ActiveCell.row, lista.ActiveCell.col).text = txt.text
    End Sub
    Public Sub ayudaProducto(lista As Grid, ByRef txt As TextBox)
        Dim tipre As String
        servidorAyuda = servidor
        basedatosAyuda = basedatos & rubro
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "r_maestroproductos_fijo_" & rubro & " AS mpf LEFT JOIN r_maestroproductos_precios_" & rubro & " AS mpp ON mpf.codigobarra = mpp.codigo LEFT JOIN r_maestroproductos_stock_" & rubro & " AS mps ON mpf.codigobarra = mps.codigo"
        mensajeAyuda = "Ayuda de Productos"
        camposAyuda = Array("mpf.codigobarra", "mpf.descripcion", "mpf.referenciaproveedor", "CONCAT('$ ', IFNULL(FORMAT(mpp.preciopuntoventa,0),0))", "IFNULL(mps.stockactual,0)")
        cabezasAyuda = Array("codigo", "descripcion", "referencia", "precio", "stock")
        largoAyuda = Array("13n", "40s", "20s", "10n")
        tipre = "01"
        
        condicionAyuda = "mpp.local = '" + empresaActiva + "' AND mps.local = '" & empresaActiva & "' AND mps.bodega = '" & bodega & "' AND mpp.codigoprecio = '" & tipre & "' and mpf.descontinuado='0'"
        
        Rem condicionAyuda = "no"
        cantidadAyuda = 4
        Call Mayuda.cargaAyuda(txt)
        lista.Cell(lista.ActiveCell.row, lista.ActiveCell.col).text = txt.text
    End Sub
    
    Public Sub ayudaProductotxt(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = basedatos & rubro
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "r_maestroproductos_fijo_" & rubro & " AS mpf"
        mensajeAyuda = "Ayuda de Productos"
        camposAyuda = Array("mpf.codigobarra", "mpf.descripcion", "mpf.referenciaproveedor", "mpf.")
        cabezasAyuda = Array("codigo", "descripcion", "referencia")
        largoAyuda = Array("13n", "50s", "20s")
        condicionAyuda = "mpf.descontinuado='0'"
        cantidadAyuda = 3
        Call Mayuda.cargaAyuda(txt)
    End Sub












