Attribute VB_Name = "FdetallePagos"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type pagos
        tipodocumento As String
        numeroDocumento As String
        rut As String
        sucursal As String
        linea As String
        tipopago As String
        MONTO As String
        CREDITO As String
        NUMERO As String
        Banco As String
        fecha As String
        vencimiento As String
        cuenta As String
        vencDoc As String
        foliosii As String
    End Type
    
'=============================================================================
'LEER PAGO
'=============================================================================
'    Public Function leerVenta(ByRef v As venta, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String, ByRef data As Adodc, ByRef lista As Grid) As Boolean
'        If leerVentaCabeza(v.cabeza, codigo1, codigo2, operador) = True Then
'            If leerVentaImpuestos(v.impuestos, codigo1, v.cabeza.numero, "=") = True Then
'                Call leerVentaDetalle(data, codigo1, v.cabeza.numero, "=", lista)
'                leerVenta = True
'            Else
'                leerVenta = False
'            End If
'        Else
'            leerVenta = False
'        End If
'    End Function
'
'    Private Function leerVentaCabeza(ByRef vc As ventaCabeza, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String) As Boolean
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        campos(0, 0) = "tipo"
'        campos(1, 0) = "numero"
'        campos(2, 0) = "fecha"
'        campos(3, 0) = "rut"
'        campos(4, 0) = "sucursal"
'        campos(5, 0) = ""
'
'        campos(0, 2) = "movimientos_cabeza_" & empresaactiva
'
'        condicion = "tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "'"
'        If operador = "<" Then
'            condicion = condicion & "ORDER BY numero DESC"
'        Else
'            condicion = condicion & "ORDER BY numero ASC"
'        End If
'        op = 5
'        sql.response = campos
'        Set sql.Conexion = gestion
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerVentaCabeza = True
'            Call asignaCabeza(vc, sql)
'        Else
'            leerVentaCabeza = False
'        End If
'    End Function
'
'    Private Sub leerVentaDetalle(ByRef data As Adodc, ByVal codigo1 As String, ByVal codigo2 As String, ByVal operador As String, ByRef lista As Grid)
'        Dim tabla As String
'        tabla = "SELECT codigo, CONCAT(cantidad, '" & vbTab & "', precio, '" & vbTab & "', total) AS item "
'        tabla = tabla & "FROM movimientos_detalle_" & empresaactiva & " "
'        tabla = tabla & "WHERE tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' ORDER BY linea ASC"
'        Call ConectarControlData(data, servidor, basedatos, usuario, password, tabla)
'        If data.Recordset.RecordCount > 0 Then
'            lista.Rows = 1
'            lista.AutoRedraw = False
'            data.Recordset.MoveFirst
'
'            While Not data.Recordset.EOF
'                lista.AddItem data.Recordset.Fields("codigo") & vbTab & leerNombreProducto(data.Recordset.Fields("codigo")) & vbTab & data.Recordset.Fields("item"), True
'                data.Recordset.MoveNext
'            Wend
'
'            lista.AutoRedraw = True
'            lista.Refresh
'        End If
'    End Sub
'
'    Private Function leerVentaImpuestos(ByRef vi As ventaImpuestos, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String) As Boolean
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        campos(0, 0) = "tipo"
'        campos(1, 0) = "numero"
'        campos(2, 0) = "vencimiento"
'        campos(3, 0) = "notapedido"
'        campos(4, 0) = "cajera"
'        campos(5, 0) = "subtotal"
'        campos(6, 0) = "neto"
'        campos(7, 0) = "iva"
'        campos(8, 0) = "descuentoporcentaje"
'        campos(9, 0) = "descuentopesos"
'        campos(10, 0) = "total"
'        campos(11, 0) = ""
'
'        campos(0, 2) = "movimientos_impuestos_" & empresaactiva
'
'        condicion = "tipo = '" & codigo1 & "' AND numero " & operador & " '" & codigo2 & "' "
'        If operador = "<" Then
'            condicion = condicion & "ORDER BY numero DESC"
'        Else
'            condicion = condicion & "ORDER BY numero ASC"
'        End If
'        op = 5
'        sql.response = campos
'        Set sql.Conexion = gestion
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerVentaImpuestos = True
'            Call asignaImpuestos(vi, sql)
'        Else
'            leerVentaImpuestos = False
'        End If
'    End Function
'=============================================================================
'LEER PAGOS
'=============================================================================

'=============================================================================
'GRABAR PAGOS
'=============================================================================
    Public Sub grabarPagos(ByRef lista As Grid, ByRef p As pagos, ByVal modifica As Boolean, ByVal sucursal As String)
        
        Dim op As Integer
        Dim i As Long
        Dim lin As String
        Dim abono As Double
        Set sql = New sqlventas.sqlventa
        
        condicion = ""
        If modifica = True Then
            Call eliminarPagos(p.tipodocumento, p.numeroDocumento, p.fecha, PVentas.dato30.text)
        End If
        op = 2
        abono = 0
        For i = 1 To lista.Rows - 1
            lin = Str(i)
            lin = Mid(lin, 2, Len(lin))
            If lista.Cell(i, 1).text <> "" And lista.Cell(i, 2).text <> "" Then
                p.linea = String(3 - Len(lin), "0") & lin
                p.tipopago = Left(lista.Cell(i, 1).text, 1)
                p.MONTO = lista.Cell(i, 2).text
                If p.tipopago = 7 Then
                    p.tipopago = 1
                End If
                If p.tipopago = 9 Then
                    p.CREDITO = lista.Cell(i, 2).text
                    Call grabarCobranza(p, sucursal, lista)
                Else
                    abono = abono + CDbl(lista.Cell(i, 2).text)
                End If
                p.NUMERO = lista.Cell(i, 3).text
                p.Banco = lista.Cell(i, 4).text
                p.cuenta = lista.Cell(i, 5).text
                p.vencimiento = lista.Cell(i, 8).text & "-" & lista.Cell(i, 7).text & "-" & lista.Cell(i, 6).text
                
                Rem Call designa(p, sql)
                
                Rem Set sql.conexion = gestionRubro
                Rem Call sql.sqlventas(op, condicion)
                
'                If Left(lista.Cell(i, 1).text, 1) = "2" Or Left(lista.Cell(i, 1).text, 1) = "3" Then
'                    Call designaCartera(p, sql)
'
'                    Set sql.conexion = ventasRubro
'                    Call sql.sqlventas(op, condicion)
'                End If
'
                Call designaPago(p, sql)
                
                Set sql.conexion = ventasRubro
                Call sql.sqlventas(op, condicion)
            End If
        Next i
        Call modificaAbono(Format(abono, "########0"), p.tipodocumento, p.foliosii)
    End Sub
    
    Private Sub grabarCobranza(ByRef p As pagos, ByVal sucursal As String, ByRef lista As Grid)
        
        Dim tabla As String
        Dim csql As New rdoQuery
        Dim abono As Double
        Dim i As Long
        
        abono = 0
        For i = 1 To lista.Rows - 1
            If InStr(1, lista.Cell(i, 1).text, "9", vbBinaryCompare) = 0 Then
                abono = abono + CDbl(lista.Cell(i, 2).text)
            End If
        Next i
        
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "INSERT INTO sv_documentos_cobranza_" & empresaActiva & " "
        csql.sql = csql.sql & "(local, tipo, numero, fechaemision, vencimiento, rut, sucursal, cajera, monto, abono, observaciones, vendedor) "
        csql.sql = csql.sql & "VALUES('" & empresaActiva & "', '" & p.tipodocumento & "', '" & p.numeroDocumento & "', '" & p.fecha & "', '" & p.vencDoc & "', '" & p.rut & "', '" & sucursal & "', '" & cajera & "', '" & p.MONTO & "', '" & abono & "', 'GENERADO AUTOMATICAMENTE POR VENTA A CREDITO', '" & vend & "') "
        'condicion = "WHERE local = '" & empresaActiva & "' AND tipo = '" & tipo & "' AND numero = '" & numero & "'"
        'cSql.sql = "UPDATE sv_documento_cabeza SET abono = abono + '" & abono & "' "
        'cSql.sql = cSql.sql & condicion
        csql.Execute
        
        'Set cSql.ActiveConnection = gestionRubro
        'condicion = "WHERE tipo = '" & tipo & "' AND numero = '" & numero & "'"
        'cSql.sql = "UPDATE l_movimientos_cabeza_" & empresaActiva & " SET montocancelado = montocancelado + '" & abono & "' "
        'cSql.sql = cSql.sql & condicion
        'cSql.Execute
    End Sub
    
    Private Sub modificaAbono(ByVal abono As String, ByVal TIPO As String, ByVal NUMERO As String)
        
        Dim tabla As String
        Dim csql As New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        condicion = "WHERE local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND foliosii = '" & NUMERO & "'"
        csql.sql = "UPDATE sv_documento_cabeza_" + empresaActiva + " SET abono = abono + '" & abono & "' "
        csql.sql = csql.sql & condicion
        csql.Execute
            Call sincronizadatos(csql.sql, ventasRubro)
        
    End Sub
'=============================================================================
'GRABAR PAGOS
'=============================================================================

'=============================================================================
'ELIMINAR PAGOS
'=============================================================================
    Public Sub eliminarPagos(ByVal TIPO As String, ByVal NUMERO As String, fecha, caja)
    
        Call eliminarPagosSV(TIPO, NUMERO, fecha, caja)
    End Sub
    
    Public Sub eliminarPagosSV(ByVal TIPO As String, ByVal NUMERO As String, fecha, caja)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND numero = '" & NUMERO & "' and fecha='" & fecha & "' and caja='" + caja + "' "
        op = 4
        CAMPOS(0, 2) = "sv_documento_pagos_" + empresaActiva
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        
        condicion = "local = '" & empresaActiva & "' AND tipodocumento = '" & TIPO & "' AND numero = '" & NUMERO & "'"
'        op = 4
'        campos(0, 2) = "sv_carteracheques"
'        sql.response = campos
'        Set sql.conexion = ventasRubro
'        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR PAGOS
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
'    Private Sub asignaCabeza(ByRef vc As ventaCabeza, ByRef sql as sqlventas.sqlventa)
'        vc.tipo = sql.response(0, 3)
'        vc.numero = sql.response(1, 3)
'        vc.fecha = sql.response(2, 3)
'        vc.rut = sql.response(3, 3)
'        vc.sucursal = sql.response(4, 3)
'    End Sub
'
'    Private Sub asignaImpuestos(ByRef vi As ventaImpuestos, ByRef sql as sqlventas.sqlventa)
'        vi.tipo = sql.response(0, 3)
'        vi.numero = sql.response(1, 3)
'        vi.vencimiento = sql.response(2, 3)
'        vi.notapedido = sql.response(3, 3)
'        vi.cajera = sql.response(4, 3)
'        vi.subtotal = sql.response(5, 3)
'        vi.neto = sql.response(6, 3)
'        vi.iva = sql.response(7, 3)
'        vi.descuentoporcentaje = sql.response(8, 3)
'        vi.descuentopesos = sql.response(9, 3)
'        vi.total = sql.response(10, 3)
'    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef p As pagos, ByRef sql As sqlventas.sqlventa)
     
    End Sub
    
    Private Sub designaCartera(ByRef p As pagos, ByRef sql As sqlventas.sqlventa)
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "rut"
        CAMPOS(2, 0) = "sucursal"
        CAMPOS(3, 0) = "numerocheque"
        CAMPOS(4, 0) = "banco"
        CAMPOS(5, 0) = "plaza"
        CAMPOS(6, 0) = "monto"
        CAMPOS(7, 0) = "fechavencimiento"
        CAMPOS(8, 0) = "tipodocumento"
        CAMPOS(9, 0) = "numero"
        CAMPOS(10, 0) = "fecharecepcion"
        CAMPOS(11, 0) = "codigolocal"
        CAMPOS(12, 0) = "cajera"
        CAMPOS(13, 0) = "numerodecuenta"
        CAMPOS(14, 0) = ""
    
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = p.rut
        CAMPOS(2, 1) = p.sucursal
        CAMPOS(3, 1) = p.NUMERO
        CAMPOS(4, 1) = p.Banco
        CAMPOS(5, 1) = ""
        CAMPOS(6, 1) = p.MONTO
        CAMPOS(7, 1) = p.vencimiento
        CAMPOS(8, 1) = p.tipodocumento
        CAMPOS(9, 1) = p.numeroDocumento
        CAMPOS(10, 1) = p.fecha
        CAMPOS(11, 1) = empresaActiva
        CAMPOS(12, 1) = ""
        CAMPOS(13, 1) = p.cuenta
        CAMPOS(14, 1) = ""
        
        CAMPOS(0, 2) = "sv_carteracheques"
        sql.response = CAMPOS
    End Sub
    
    Private Sub designaPago(ByRef p As pagos, ByRef sql As sqlventas.sqlventa)
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "numero"
        CAMPOS(3, 0) = "lineapago"
        CAMPOS(4, 0) = "fecha"
        CAMPOS(5, 0) = "tipopago"
        CAMPOS(6, 0) = "cuentacorriente"
        CAMPOS(7, 0) = "banco"
        CAMPOS(8, 0) = "plaza"
        CAMPOS(9, 0) = "numerodocumento"
        CAMPOS(10, 0) = "monto"
        CAMPOS(11, 0) = "vencimiento"
        CAMPOS(12, 0) = "rut"
        CAMPOS(13, 0) = "foliofiscal"
        CAMPOS(14, 0) = "caja"
        CAMPOS(15, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = p.tipodocumento
        CAMPOS(2, 1) = p.numeroDocumento
        CAMPOS(3, 1) = p.linea
        CAMPOS(4, 1) = p.fecha
        CAMPOS(5, 1) = p.tipopago
        CAMPOS(6, 1) = p.cuenta
        CAMPOS(7, 1) = p.Banco
        CAMPOS(8, 1) = ""
        CAMPOS(9, 1) = p.NUMERO
        CAMPOS(10, 1) = p.MONTO
        CAMPOS(11, 1) = p.vencimiento
        CAMPOS(12, 1) = p.rut
        CAMPOS(13, 1) = p.foliosii
        CAMPOS(14, 1) = PVentas.dato30.text
        CAMPOS(0, 2) = "sv_documento_pagos_" + empresaActiva
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



Public Function generacadena(response, Opcion) As String
Dim cadena As String
Dim response1 As String
Dim i As Double


Select Case Opcion

                Case 2:    '<<<<<   INSERTA   >>>>>>


                    cadena = "INSERT INTO " & response(0, 2) & " ("
                    response1 = ""
                    'NOMBRE DE response.
                    i = 0
                    While response(i, 0) <> ""
                        cadena = cadena & response(i, 0) + ","
                        response1 = response1 & "[" & response(i, 0) & "]"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & ") VALUES ("
                    
                    'VALORES ASIGNADOS A CADA CAMPO.
                    i = 0
                    While response(i, 0) <> ""
                        cadena = cadena & "'" & response(i, 1) & "',"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & ") "
                    cadena = cadena & "ON DUPLICATE KEY UPDATE " & response(0, 0) & " = " & response(0, 0)
                    generacadena = cadena
                    
                Case 3:    '<<<<<   ACTUALIZA   >>>>>>
                    
        
                    cadena = "UPDATE " & response(0, 2) & " SET "
                    i = 0
                    response1 = ""
                    response2 = ""
                    While response(i, 0) <> ""
                        cadena = cadena & response(i, 0) & "= '" & response(i, 1) & "',"
                        response1 = response1 & "[" & response(i, 0) & "]"
                        response2 = response2 & "[" & response(i, 1) & "]"
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " WHERE " & condicion
                    generacadena = cadena
                    
                    
                Case 4:    '<<<<<   ELIMINA   >>>>>>
                    If audit = True Then
                        Call auditoria(Opcion, condicion)
                    End If
                    cadena = "DELETE FROM " & response(0, 2) & " WHERE " & condicion
                    generacadena = cadena
                     
        
                Case 5:    '<<<<<   LEE   >>>>>>
                    cadena = "SELECT "
                    i = 0
                    While response(i, 0) <> ""
                        cadena = cadena & response(i, 0) & ","
                        i = i + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " FROM " & response(0, 2) & " WHERE " & condicion
                    generacadena = cadena
 End Select
                    
                    
 End Function
