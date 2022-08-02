Attribute VB_Name = "FPClientes"
Option Explicit
    Private CAMPOS(30, 3) As String
    
    Private Type pagoClienteCabeza
        FOLIO As String
        rut As String
        fecha As String
        TIPO As String
        MONTO As String
        GLOSA As String
        fechadeposito As String
    End Type
    
    Private Type pagoClienteDetalle
        FOLIO As String
        rut As String
        fecha As String
        linea As String
        tipodocumento As String
        numeroDocumento As String
        montoabonado As String
        tipopago As String
        montopago As String
    End Type
    
    Private Type pagoClienteCheque
        rut As String
        numerocheque As String
        Banco As String
        MONTO As String
        vencimiento As String
        tipodocumento As String
        FOLIO As String
        fecha As String
        codigolocal As String
        cajera As String
        numerodecuenta As String
    End Type
    
    Public Type pagoCliente
        c As pagoClienteCabeza
        D As pagoClienteDetalle
        ch As pagoClienteCheque
    End Type

'=============================================================================
'LEER PAGO CLIENTE
'=============================================================================
    Public Function leerPagoCliente(ByRef pc As pagoCliente, ByVal CODIGO As String, ByRef operador As String, ByRef data As Adodc, ByRef lista1 As Grid, ByRef lista2 As Grid, ByRef lbl1 As Label, ByRef lbl2 As Label, ByRef lbl3 As Label) As Boolean
        If leerPagoClienteCabeza(pc.c, CODIGO, operador) = True Then
            Call leerPagoClienteDetalle(data, pc.c.FOLIO, "=", lista1, lista2, lbl1, lbl2, lbl3)
            leerPagoCliente = True
        Else
            leerPagoCliente = False
        End If
    End Function
    
    Private Function leerPagoClienteCabeza(ByRef pcc As pagoClienteCabeza, ByVal CODIGO As String, ByRef operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "numero"
        CAMPOS(1, 0) = "rut"
        CAMPOS(2, 0) = "fecha"
        CAMPOS(3, 0) = "tipopago"
        CAMPOS(4, 0) = "monto"
        CAMPOS(5, 0) = "glosa"
'        campos(6, 0) = "IFNULL(fechadeposito,'')"
        CAMPOS(6, 0) = ""
        
        CAMPOS(0, 2) = "sv_pagos_cabeza_" & empresaActiva
        
        
        condicion = "local = '" & empresaActiva & "' AND numero " & operador & " '" & CODIGO & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY numero DESC"
        Else
            condicion = condicion & "ORDER BY numero ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPagoClienteCabeza = True
            Call asignaCabeza(pcc, sql)
        Else
            leerPagoClienteCabeza = False
        End If
    End Function
    
    Private Sub leerPagoClienteDetalle(ByRef data As Adodc, ByVal CODIGO As String, ByVal operador As String, ByRef lista1 As Grid, ByRef lista2 As Grid, ByRef lblTotal1 As Label, ByRef lblTotal2 As Label, ByRef lblTotal3 As Label)
        Dim tabla As String
        Dim total As Double
        Dim notacredito As Double
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Dim cSql1 As New rdoQuery
        Dim resultados1 As rdoResultset
        'Cheques
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT CONCAT(c.banco, '" & vbTab & "', mb.nombre, '" & vbTab & "', c.numerocheque, '" & vbTab & "', c.numerodecuenta, '" & vbTab & "', c.monto, '" & vbTab & "', DATE_FORMAT(c.fechavencimiento,'%d'), '" & vbTab & "', DATE_FORMAT(c.fechavencimiento,'%m'), '" & vbTab & "', DATE_FORMAT(c.fechavencimiento,'%Y')) AS item, c.monto "
        csql.sql = csql.sql & "FROM sv_pagos_cabeza_" & empresaActiva & " AS pc INNER JOIN sv_carteracheques AS c ON pc.local = c.local AND pc.numero = c.numero AND pc.rut = c.rut LEFT JOIN " & baseVentas & ".sv_maestrobancos AS mb ON c.banco = mb.codigobanco "
        csql.sql = csql.sql & "WHERE pc.local = '" & empresaActiva & "' AND pc.numero = '" & CODIGO & "' AND c.tipodocumento = 'PA'"
        csql.Execute
        'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
        total = 0
        lista1.Rows = 1
        lista1.AutoRedraw = False
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                If IsNull(resultados("item")) = False Then
                    lista1.AddItem resultados("item"), True
                    total = total + CDbl(resultados("monto"))
                End If
                resultados.MoveNext
            Wend
        End If
        Set csql = Nothing
        csql.Close
        Set resultados = Nothing
        
        lista1.AutoRedraw = True
        lista1.Refresh
        lblTotal1.Caption = Format(total, "$ ###,###,##0")
        
        'Pagados
        Set cSql1.ActiveConnection = ventasRubro
        cSql1.sql = "SELECT CONCAT(tipo, '" & vbTab & "', documento, '" & vbTab & "', montototal, '" & vbTab & "', monto) AS item, monto, tipo "
        cSql1.sql = cSql1.sql & "FROM sv_pagos_detalle_" & empresaActiva & " "
        cSql1.sql = cSql1.sql & "WHERE local = '" & empresaActiva & "' AND numero = '" & CODIGO & "' ORDER BY linea ASC"
        cSql1.Execute
        'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
        total = 0
        notacredito = 0
        lista2.Rows = 1
        lista2.AutoRedraw = False
        If cSql1.RowsAffected > 0 Then
          Set resultados1 = cSql1.OpenResultset
            While Not resultados1.EOF
                lista2.AddItem resultados1("item"), True
                If resultados1("tipo") = "NV" Then
                    notacredito = notacredito + CDbl(resultados1("monto"))
                Else
                    total = total + CDbl(resultados1("monto"))
                End If
                resultados1.MoveNext
            Wend
        End If
        Set cSql1 = Nothing
        cSql1.Close
        Set resultados1 = Nothing
        
        lista2.AutoRedraw = True
        lista2.Refresh
        lblTotal2.Caption = Format(total - notacredito, "$ ###,###,##0")
        lblTotal3.Caption = Format(total, "$ ###,###,##0")
    End Sub
'=============================================================================
'LEER PAGO CLIENTE
'=============================================================================

'=============================================================================
'LEER DOCUMENTOS CLIENTE
'=============================================================================
    Public Sub leerDocumentos(ByRef data As Adodc, ByVal rut As String, ByRef lblTotal As Label, ByRef lista As Grid)
        Dim tabla As String
        Dim total As Double
        Dim TIPO As String
        Dim cadena As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        
        total = 0
        lista.Rows = 1
        lista.AutoRedraw = False
        'COBRANZA
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT CONCAT(tipo, '" & vbTab & "', numero, '" & vbTab & "', monto, '" & vbTab & "', monto-abono) AS item, monto - abono AS saldo, tipo, numero "
        csql.sql = csql.sql & "FROM sv_documentos_cobranza_" & empresaActiva & " "
        csql.sql = csql.sql & "WHERE local = '" & empresaActiva & "' AND rut = '" & rut & "' AND monto <> abono "
        csql.sql = csql.sql & "UNION "
        csql.sql = csql.sql & "SELECT CONCAT('SA', '" & vbTab & "', numeropago, '" & vbTab & "', monto, '" & vbTab & "', monto-abono) AS item, monto - abono AS saldo, 'SA' AS tipo, numeropago AS numero "
        csql.sql = csql.sql & "FROM " & baseVentas & ".sv_maestroclientes_saldos "
        csql.sql = csql.sql & "WHERE rut = '" & rut & "' AND monto <> abono "
        csql.sql = csql.sql & "ORDER BY tipo, numero ASC "
        csql.Execute
        
        'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
        If csql.RowsAffected > 0 Then
           Set resultados = csql.OpenResultset
            While Not resultados.EOF
                lista.AddItem resultados("item"), True
                lista.Cell(lista.Rows - 1, 0).text = resultados("tipo")
                total = total + CDbl(resultados("saldo"))
                resultados.MoveNext
            Wend
        End If
        Set csql = Nothing
        csql.Close
        Set resultados = Nothing
        
        lista.AutoRedraw = True
        lista.Refresh
        lblTotal.Caption = Format(total, "$ ###,###,##0")
    End Sub
    
    Public Sub agregarNotaCredito(ByRef data As Adodc, ByRef Pagar As Grid, ByRef lbl As Label)
        Dim TIPO As String
        Dim NUMERO As String
        Dim total As Double
        Dim tabla As String
        Dim fila As Long
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        
        fila = Pagar.Rows - 1
        TIPO = Pagar.Cell(fila, 1).text
        NUMERO = Pagar.Cell(fila, 2).text
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT CONCAT(dd.tipo, '" & vbTab & "', dd.numero, '" & vbTab & "', ROUND(dd.total + (dd.total * 19 / 100) + (dd.total * 12 / 100), 0), '" & vbTab & "', CONCAT(REPEAT('0', 10 - LENGTH(dd.cantidad)), dd.cantidad)) AS item, ROUND(dd.total + (dd.total * " & iva & " / 100) + (dd.total * " & iha & " / 100), 0) AS total "
        csql.sql = csql.sql & "FROM sv_documento_detalle_" + empresaActiva + " AS dd INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dd.local = dc.local AND dd.tipo = dc.tipo AND dd.numero = dc.numero "
        csql.sql = csql.sql & "WHERE dd.local = '" & empresaActiva & "' AND dd.tipo = 'NV' AND dd.cantidad = " & Val(NUMERO) & " AND dd.codigo = '0000000000100' AND dc.total < dc.abono AND dc.nula = 'N' "
        csql.sql = csql.sql & "ORDER BY dd.linea ASC"
        csql.Execute
        
        'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
        total = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Pagar.AddItem resultados("item"), True
                total = total + CDbl(resultados("total"))
                resultados.MoveNext
            Wend
            lbl.Caption = Format(CDbl(lbl.Caption) - total, "$ ###,###,##0")
        End If
        Set csql = Nothing
        csql.Close
        Set resultados = Nothing
        
        Pagar.Cell(fila, 4).text = Format(CDbl(Pagar.Cell(fila, 4).text) - total, "$ ###,###,##0")
    End Sub
'=============================================================================
'LEER DOCUMENTOS CLIENTE
'=============================================================================


'=============================================================================
'GRABAR PAGO CLIENTE
'=============================================================================
    Public Sub grabarPagoCliente(ByRef pc As pagoCliente, ByVal modifica As Boolean, ByRef lista1 As Grid, ByRef lista2 As Grid)
        Call grabarPagoClienteCabeza(pc.c, modifica)
        Call grabarPagoClienteDetalle(lista2, pc.D, modifica, pc.c)
        Call grabarPagoClienteCheque(lista1, pc, modifica)
        If modifica = False Then
            Call modificaAbono(lista2, pc.c.rut)
        End If
    End Sub
    
    Private Sub grabarPagoClienteCabeza(ByRef pcc As pagoClienteCabeza, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designaCabeza(pcc, sql)
        condicion = ""
        If modifica = True Then
            condicion = "numero = '" & pcc.FOLIO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = PClientes.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
        op = sql.Status
    End Sub
    
    Private Sub grabarPagoClienteDetalle(ByRef lista As Grid, ByRef pcd As pagoClienteDetalle, ByVal modifica As Boolean, ByRef p As pagoClienteCabeza)
        
        Dim op As Integer
        Dim i As Long
        Dim lin As String
        Set sql = New sqlventas.sqlventa
        
        condicion = ""
        If modifica = True Then
            Call eliminarPagoDetalle(p)
        End If
        op = 2
        
        For i = 1 To lista.Rows - 1
            lin = Str(i)
            lin = Mid(lin, 2, Len(lin))
            If lista.Cell(i, 1).text <> "" And lista.Cell(i, 2).text <> "" And lista.Cell(i, 3).text <> "" And lista.Cell(i, 4).text <> "" Then
                pcd.linea = String(3 - Len(lin), "0") & lin
                pcd.tipodocumento = lista.Cell(i, 1).text
                pcd.numeroDocumento = lista.Cell(i, 2).text
                pcd.montoabonado = Format(lista.Cell(i, 4).text, "########0")
                pcd.montopago = lista.Cell(i, 3).text
                
                Call designaDetalle(pcd, sql)
                
                Set sql.conexion = ventasRubro
                sql.audit = True: sql.programaactivo = PClientes.Caption
                Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
                Call sql.sqlventas(op, condicion)
            End If
        Next i
    End Sub
    
    Private Sub grabarPagoClienteCheque(ByRef lista As Grid, ByRef pc As pagoCliente, ByVal modifica As Boolean)
        
        Dim op As Integer
        Dim i As Long
        Set sql = New sqlventas.sqlventa
        
        condicion = ""
        If modifica = True Then
            Call eliminarPagoCheque(pc.c)
        End If
        op = 2
        
        For i = 1 To lista.Rows - 1
            If lista.Cell(i, 1).text <> "" And lista.Cell(i, 2).text <> "" And lista.Cell(i, 3).text <> "" And lista.Cell(i, 4).text <> "" And lista.Cell(i, 5).text <> "" And lista.Cell(i, 6).text <> "" Then
                pc.ch.Banco = lista.Cell(i, 1).text
                pc.ch.numerocheque = lista.Cell(i, 3).text
                pc.ch.numerodecuenta = lista.Cell(i, 4).text
                pc.ch.MONTO = lista.Cell(i, 5).text
                pc.ch.vencimiento = lista.Cell(i, 8).text & "-" & lista.Cell(i, 7).text & "-" & lista.Cell(i, 6).text
                
                Call designaCheques(pc.ch, sql)
                
                Set sql.conexion = ventasRubro
                 sql.audit = True: sql.programaactivo = PClientes.Caption
                Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
                Call sql.sqlventas(op, condicion)
            End If
        Next i
    End Sub
    
    Private Sub modificaAbono(ByRef lista As Grid, ByVal rut As String)
        
        Dim i As Long
        Dim TIPO As String
        Dim NUMERO As String
        Dim abono As String
        Dim tabla As String
        Dim csql As rdoQuery
        
        For i = 1 To lista.Rows - 1
            TIPO = lista.Cell(i, 1).text
            NUMERO = lista.Cell(i, 2).text
            abono = Format(lista.Cell(i, 4).text, "########0")
            
            If TIPO <> "SA" Then
                Set csql = New rdoQuery
                Set csql.ActiveConnection = ventasRubro
                csql.sql = "UPDATE sv_documentos_cobranza_" & empresaActiva & " AS dc SET dc.abono = dc.abono + '" & abono & "' "
                condicion = "WHERE dc.local = '" & empresaActiva & "' AND dc.tipo = '" & TIPO & "' AND dc.numero = '" & NUMERO & "'"
                csql.sql = csql.sql & condicion
                csql.Execute
                    Call sincronizadatos(csql.sql, ventasRubro)
                csql.Close
                Set csql = Nothing
                
                Set csql = New rdoQuery
                Set csql.ActiveConnection = ventasRubro
                csql.sql = "UPDATE sv_documento_cabeza_" + empresaActiva + " AS c SET c.abono = c.abono + '" & abono & "' "
                condicion = "WHERE c.local = '" & empresaActiva & "' AND c.tipo = '" & TIPO & "' AND c.numero = '" & NUMERO & "'"
                csql.sql = csql.sql & condicion
                csql.Execute
                    Call sincronizadatos(csql.sql, ventasRubro)
                csql.Close
                Set csql = Nothing
            Else
                Set csql = New rdoQuery
                Set csql.ActiveConnection = ventas
                csql.sql = "UPDATE sv_maestroclientes_saldos AS ms SET ms.abono = ms.abono + '" & abono & "' "
                condicion = "WHERE ms.rut = '" & rut & "' AND numeropago = '" & NUMERO & "' "
                csql.sql = csql.sql & condicion
                csql.Execute
                    Call sincronizadatos(csql.sql, ventas)
                csql.Close
                Set csql = Nothing
            End If
        Next i
    End Sub
    
    Private Sub modificaAbonoAnterior(ByRef lista As Grid, ByVal rut As String)
        
        Dim i As Long
        Dim TIPO As String
        Dim NUMERO As String
        Dim abono As String
        Dim tabla As String
        Dim csql As rdoQuery
        
        For i = 1 To lista.Rows - 1
            TIPO = lista.Cell(i, 1).text
            NUMERO = lista.Cell(i, 2).text
            abono = Format(lista.Cell(i, 4).text, "########0")
            If lista.Cell(i, 1).text <> "" And lista.Cell(i, 2).text <> "" And lista.Cell(i, 3).text <> "" And lista.Cell(i, 4).text <> "" Then
                If TIPO <> "SA" Then
                    Set csql = New rdoQuery
                    Set csql.ActiveConnection = ventasRubro
                    csql.sql = "UPDATE sv_documentos_cobranza_" & empresaActiva & " AS dc SET dc.abono = dc.abono - '" & abono & "' "
                    condicion = "WHERE dc.local = '" & empresaActiva & "' AND dc.tipo = '" & TIPO & "' AND dc.numero = '" & NUMERO & "'"
                    csql.sql = csql.sql & condicion
                    csql.Execute
                    Call sincronizadatos(csql.sql, ventasRubro)
                    csql.Close
                    Set csql = Nothing
                    
                    Set csql = New rdoQuery
                    Set csql.ActiveConnection = ventasRubro
                    csql.sql = "UPDATE sv_documento_cabeza_" + empresaActiva + " AS c SET c.abono = c.abono - '" & abono & "' "
                    condicion = "WHERE c.local = '" & empresaActiva & "' AND c.tipo = '" & TIPO & "' AND c.numero = '" & NUMERO & "'"
                    csql.sql = csql.sql & condicion
                    csql.Execute
                        Call sincronizadatos(csql.sql, ventasRubro)
                    csql.Close
                    Set csql = Nothing
                Else
                    Set csql = New rdoQuery
                    Set csql.ActiveConnection = ventas
                    csql.sql = "UPDATE sv_maestroclientes_saldos AS ms SET ms.abono = ms.abono - '" & abono & "' "
                    condicion = "WHERE ms.rut = '" & rut & "' AND numeropago = '" & NUMERO & "' "
                    csql.sql = csql.sql & condicion
                    csql.Execute
                        Call sincronizadatos(csql.sql, ventas)
                    csql.Close
                    Set csql = Nothing
                End If
            End If
        Next i
    End Sub
'=============================================================================
'GRABAR PAGO CLIENTE
'=============================================================================

'=============================================================================
'ELIMINAR PAGO CLIENTE
'=============================================================================
    Public Sub eliminarPagoCliente(ByRef pc As pagoCliente, ByRef lista As Grid)
        Call eliminarPagoCabeza(pc.c)
        Call eliminarPagoDetalle(pc.c)
        Call eliminarPagoCheque(pc.c)
        Call modificaAbonoAnterior(lista, pc.c.rut)
    End Sub
    
    Private Sub eliminarPagoCabeza(ByRef c As pagoClienteCabeza)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND numero = '" & c.FOLIO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_pagos_cabeza_" & empresaActiva
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = PClientes.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        op = sql.Status
    End Sub
    
    Private Sub eliminarPagoDetalle(ByRef c As pagoClienteCabeza)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND numero = '" & c.FOLIO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_pagos_detalle_" & empresaActiva
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = PClientes.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        op = sql.Status
    End Sub
    
    Private Sub eliminarPagoCheque(ByRef c As pagoClienteCabeza)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND tipodocumento = 'PA' AND numero = '" & c.FOLIO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_carteracheques"
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = PClientes.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        op = sql.Status
    End Sub
'=============================================================================
'ELIMINAR PAGO CLIENTE
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asignaCabeza(ByRef pcc As pagoClienteCabeza, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        pcc.FOLIO = sql.response(0, 3)
        pcc.rut = sql.response(1, 3)
        pcc.fecha = sql.response(2, 3)
        pcc.TIPO = sql.response(3, 3)
        pcc.MONTO = sql.response(4, 3)
        pcc.GLOSA = sql.response(5, 3)
        pcc.fechadeposito = sql.response(6, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designaCabeza(ByRef pcc As pagoClienteCabeza, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "rut"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = "tipopago"
        CAMPOS(5, 0) = "monto"
        CAMPOS(6, 0) = "glosa"
        CAMPOS(7, 0) = "fechadeposito"
        CAMPOS(8, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = pcc.FOLIO
        CAMPOS(2, 1) = pcc.rut
        CAMPOS(3, 1) = pcc.fecha
        CAMPOS(4, 1) = pcc.TIPO
        CAMPOS(5, 1) = pcc.MONTO
        CAMPOS(6, 1) = pcc.GLOSA
        CAMPOS(7, 1) = pcc.fechadeposito
        CAMPOS(8, 1) = ""
        
        CAMPOS(0, 2) = "sv_pagos_cabeza_" & empresaActiva
        
        sql.response = CAMPOS
    End Sub
    
    Private Sub designaDetalle(ByRef pcd As pagoClienteDetalle, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "rut"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = "linea"
        CAMPOS(5, 0) = "tipo"
        CAMPOS(6, 0) = "documento"
        CAMPOS(7, 0) = "monto"
        CAMPOS(8, 0) = "montototal"
        CAMPOS(9, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = pcd.FOLIO
        CAMPOS(2, 1) = pcd.rut
        CAMPOS(3, 1) = pcd.fecha
        CAMPOS(4, 1) = pcd.linea
        CAMPOS(5, 1) = pcd.tipodocumento
        CAMPOS(6, 1) = pcd.numeroDocumento
        CAMPOS(7, 1) = pcd.montoabonado
        CAMPOS(8, 1) = pcd.montopago
        CAMPOS(9, 1) = ""
        
        CAMPOS(0, 2) = "sv_pagos_detalle_" & empresaActiva
        sql.response = CAMPOS
    End Sub
    
    Private Sub designaCheques(ByRef pcch As pagoClienteCheque, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "rut"
        CAMPOS(2, 0) = "sucursal"
        CAMPOS(3, 0) = "numerocheque"
        CAMPOS(4, 0) = "banco"
        CAMPOS(5, 0) = "monto"
        CAMPOS(6, 0) = "fechavencimiento"
        CAMPOS(7, 0) = "tipodocumento"
        CAMPOS(8, 0) = "numero"
        CAMPOS(9, 0) = "fecharecepcion"
        CAMPOS(10, 0) = "codigolocal"
        CAMPOS(11, 0) = "cajera"
        CAMPOS(12, 0) = "numerodecuenta"
        CAMPOS(13, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = pcch.rut
        CAMPOS(2, 1) = "0"
        CAMPOS(3, 1) = pcch.numerocheque
        CAMPOS(4, 1) = pcch.Banco
        CAMPOS(5, 1) = pcch.MONTO
        CAMPOS(6, 1) = pcch.vencimiento
        CAMPOS(7, 1) = pcch.tipodocumento
        CAMPOS(8, 1) = pcch.FOLIO
        CAMPOS(9, 1) = pcch.fecha
        CAMPOS(10, 1) = pcch.codigolocal
        CAMPOS(11, 1) = pcch.cajera
        CAMPOS(12, 1) = pcch.numerodecuenta
        CAMPOS(13, 1) = ""
        
        CAMPOS(0, 2) = "sv_carteracheques"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================





