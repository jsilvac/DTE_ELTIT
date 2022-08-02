Attribute VB_Name = "FPDetalleDocumentos"
Option Explicit
    Private CAMPOS(37, 3) As String
    Private Type ventaCabeza
        loc As String
        TIPO As String
        NUMERO As String
        fecha As String
        plazo As String
        vencimiento As String
        rut As String
        sucursal As String
        cajera As String
        notapedido As String
        notaventas As String
        ordencompra As String
        subtotal As String
        neto As String
        iva As String
        impuestoHarina As String
        impuestoila As String
        impuestoespecifico As String
        exento As String
        retencionparcial As String
        retenciontotal As String
        total As String
        abono As String
        Descuento As String
        contabilizado As String
        PAGADO As String
        comision As String
        fechapagocomision As String
        nula As String
        boletadesde As String
        boletahasta As String
        condicionesdepago As String
        transporte As String
        revisado As String
        bultos As String
        caja As String
        
        
        
    End Type
        
    Private Type ventaDetalle
        loc As String
        TIPO As String
        NUMERO As String
        linea As String
        fecha As String
        rut As String
        sucursal As String
        CODIGO As String
        descripcion As String
        cantidad As String
        unidades As String
        PRECIO As String
        Descuento As String
        total As String
        vendedor As String
        pcosto As String
        bodega As String
        vencimiento As String
    End Type
        
    Public Type ventaDocumento
        cabeza As ventaCabeza
        detalle As ventaDetalle
    End Type
    
'=============================================================================
'LEER VENTA
'=============================================================================
    Public Function leerVentaDocumento(ByRef v As ventaDocumento, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String, ByRef data As Adodc, ByRef lista As Grid, Optional caja As String) As Boolean
        leerVentaDocumento = True
        If leerVentaCabeza(v.cabeza, codigo1, codigo2, operador, caja) = False Then
            leerVentaDocumento = False
        Else
            Call leerVentaDetalle(data, codigo1, v.cabeza.NUMERO, "=", lista, v.cabeza.caja, Format(v.cabeza.fecha, "yyyy-mm-dd"))
        End If
    End Function
    
    Private Function leerVentaCabeza(ByRef vc As ventaCabeza, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String, caja) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "numero"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = "plazo"
        CAMPOS(5, 0) = "vencimiento"
        CAMPOS(6, 0) = "rut"
        CAMPOS(7, 0) = "sucursal"
        CAMPOS(8, 0) = "cajera"
        CAMPOS(9, 0) = "notapedido"
        CAMPOS(10, 0) = "notaventa"
        CAMPOS(11, 0) = "ordencompra"
        CAMPOS(12, 0) = "subtotal"
        CAMPOS(13, 0) = "neto"
        CAMPOS(14, 0) = "iva"
        CAMPOS(15, 0) = "impuestoharina"
        CAMPOS(16, 0) = "impuestoila"
        CAMPOS(17, 0) = "impuestoespecifico"
        CAMPOS(18, 0) = "exento"
        CAMPOS(19, 0) = "retencionparcial"
        CAMPOS(20, 0) = "retenciontotal"
        CAMPOS(21, 0) = "total"
        CAMPOS(22, 0) = "abono"
        CAMPOS(23, 0) = "descuento"
        CAMPOS(24, 0) = "contabilizado"
        CAMPOS(25, 0) = "pagado"
        CAMPOS(26, 0) = "comision"
        CAMPOS(27, 0) = "IFNULL(fechapagocomision,'1900-01-01')"
        CAMPOS(28, 0) = "nula"
        CAMPOS(29, 0) = "boletadesde"
        CAMPOS(30, 0) = "boletahasta"
        CAMPOS(31, 0) = "descuento2"
        CAMPOS(32, 0) = "condicionesdepago"
        CAMPOS(33, 0) = "transporte"
        CAMPOS(34, 0) = "revisado"
        CAMPOS(35, 0) = "bultos"
        CAMPOS(36, 0) = "caja"
        CAMPOS(37, 0) = ""
        CAMPOS(0, 2) = "sv_documento_cabeza_" + localAuditoria
        
        op = 5
        sql.response = CAMPOS
        Select Case listaDocumentos.formulario
            Case "auditoria"
                condicion = "caja='" & caja & "' and local = '" & localAuditoria & "' AND tipo = '" & codigo1 & "' AND foliosii " & operador & " '" & codigo2 & "' AND fecha BETWEEN '" & fechaAuditIni & "' AND '" & fechaAuditFin & "' "
                Set sql.conexion = ventasAuditoria
            Case Else
                condicion = "caja='" & caja & "' and local = '" & empresaActiva & "' AND tipo = '" & codigo1 & "' AND foliosii " & operador & " '" & codigo2 & "' AND rut = '" & rut_cliente & "' "
                Set sql.conexion = ventasAuditoria
        End Select
        If operador = "<" Then
            condicion = condicion & "ORDER BY numero DESC"
        Else
            condicion = condicion & "ORDER BY numero ASC"
        End If
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerVentaCabeza = True
            Call asignaCabeza(vc, sql)
        Else
            leerVentaCabeza = False
        End If
    End Function
    
    Private Sub leerVentaDetalle(ByRef data As Adodc, ByVal codigo1 As String, ByVal codigo2 As String, ByVal operador As String, ByRef lista As Grid, ByRef caja As String, fecha)
        Dim tabla As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
       Set csql.ActiveConnection = ventasRubro
       
'        tabla = "SELECT CONCAT(codigo, '" & vbTab & "', descripcion, '" & vbTab & "', cantidad, '" & vbTab & "', unidades, '" & vbTab & "', precio, '" & vbTab & "', descuento, '" & vbTab & "', total, '" & vbTab & "', pcosto) AS item "
'        tabla = tabla & "FROM sv_documento_detalle_" + localAuditoria + " "
'        Select Case listaDocumentos.formulario
'            Case "auditoria"
'                tabla = tabla & "WHERE fecha='" & fecha & "' and local = '" & localAuditoria & "' AND tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' and caja='" & caja & "' and local='" & localAuditoria & "' ORDER BY linea ASC"
'                Call ConectarControlData(data, servidor, baseVentas & localAuditoria, usuario, password, tabla)
'            Case Else
'                tabla = tabla & "WHERE fecha='" & fecha & "' and local = '" & empresaActiva & "' AND tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' and caja='" & caja & "' and local='" & localAuditoria & "'  ORDER BY linea ASC"
'                Call ConectarControlData(data, servidor, baseVentas & localAuditoria, usuario, password, tabla)
'        End Select

        tabla = "SELECT CONCAT(codigo, '" & vbTab & "', descripcion, '" & vbTab & "', cantidad, '" & vbTab & "', unidades, '" & vbTab & "', precio, '" & vbTab & "', descuento, '" & vbTab & "', total, '" & vbTab & "', pcosto) AS item "
        tabla = tabla & "FROM " & baseVentas & localAuditoria & " .sv_documento_detalle_" + localAuditoria + " "
        Select Case listaDocumentos.formulario
            Case "auditoria"
                tabla = tabla & "WHERE fecha='" & fecha & "' and local = '" & localAuditoria & "' AND tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' and caja='" & caja & "' and local='" & localAuditoria & "' ORDER BY linea ASC"
                 
            Case Else
                tabla = tabla & "WHERE fecha='" & fecha & "' and local = '" & empresaActiva & "' AND tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' and caja='" & caja & "' and local='" & localAuditoria & "'  ORDER BY linea ASC"
        End Select

       csql.sql = tabla
       csql.Execute
        
        lista.Rows = 1
        lista.AutoRedraw = False
        If csql.RowsAffected > 0 Then
             Set resultados = csql.OpenResultset
             
            
            While Not resultados.EOF
                lista.AddItem Replace(resultados("item"), ".", ","), True
                resultados.MoveNext
            Wend
        End If
        csql.Close
        Set resultados = Nothing
        Set csql = Nothing
        
        
        lista.AutoRedraw = True
        lista.Refresh
    End Sub
'=============================================================================
'LEER VENTA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asignaCabeza(ByRef vc As ventaCabeza, ByRef sql As sqlventas.sqlventa)
        vc.loc = sql.response(0, 3)
        vc.TIPO = sql.response(1, 3)
        vc.NUMERO = sql.response(2, 3)
        vc.fecha = sql.response(3, 3)
        vc.plazo = sql.response(4, 3)
        vc.vencimiento = sql.response(5, 3)
        vc.rut = sql.response(6, 3)
        vc.sucursal = sql.response(7, 3)
        vc.cajera = sql.response(8, 3)
        vc.notapedido = sql.response(9, 3)
        vc.notaventas = sql.response(10, 3)
        vc.ordencompra = sql.response(11, 3)
        vc.subtotal = sql.response(12, 3)
        vc.neto = sql.response(13, 3)
        vc.iva = sql.response(14, 3)
        vc.impuestoHarina = sql.response(15, 3)
        vc.impuestoila = sql.response(16, 3)
        vc.impuestoespecifico = sql.response(17, 3)
        vc.exento = sql.response(18, 3)
        vc.retencionparcial = sql.response(19, 3)
        vc.retenciontotal = sql.response(20, 3)
        vc.total = sql.response(21, 3)
        vc.abono = sql.response(22, 3)
        vc.Descuento = sql.response(23, 3)
        vc.contabilizado = sql.response(24, 3)
        vc.PAGADO = sql.response(25, 3)
        vc.comision = sql.response(26, 3)
        vc.fechapagocomision = sql.response(27, 3)
        vc.nula = sql.response(28, 3)
        vc.boletadesde = sql.response(29, 3)
        vc.boletahasta = sql.response(30, 3)
        vc.condicionesdepago = sql.response(31, 3)
        vc.transporte = sql.response(32, 3)
        vc.revisado = sql.response(33, 3)
        vc.bultos = sql.response(34, 3)
        vc.caja = sql.response(36, 3)
        
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================




