Attribute VB_Name = "FLibroVentas"
Option Explicit
    Private deptos(1 To 5, 0 To 9) As String
    Private totales(10) As Double
    Private totales2(10) As Double
    Private listaempresa As String
    
Public Sub generaInformeLV(ByRef data As Adodc, ByRef impresion As Grid, ByVal tipo As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Dim documento As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    If tipo = "FV" Then documento = "FACTURAS"
    If tipo = "BV" Then documento = "BOLETAS "
    If tipo = "ZE" Then documento = "ZETAS   "
    
    Call cargaCabeza("LISTADO LIBRO DE VENTAS " + documento + " DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
    Call resumenVentas(data, impresion, tipo, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Function resumenVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal tipo As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim tabla As String
    Dim rubAux As String
    Dim harinas As Double
    Dim subproductos As Double
    Dim envases As Double
    Dim trigo As Double
    Dim maquila As Double
    Dim otros As Double
    Dim cadena As String
    Dim tipoDoc As String
    Dim numeroDoc As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim resultados As rdoResultset
        
    Dim i As Integer

    rubAux = rubro
    Call Conectarventas(servidor, baseVentas + empresaActiva, usuario, password)
    
    
    Set csql.ActiveConnection = ventasRubro
    
    csql.sql = "SELECT dc.tipo, dc.numero , dc.fecha, dc.rut, IFNULL(mc.nombre,'') as nombre, dc.neto, dc.iva, dc.exento, dc.total,dc.impuestoilarefrescos,dc.impuestoilavinos,dc.impuestoilalicores,dc.impuestoharina,dc.impuestocarne,dc.foliosii,dc.caja "
    csql.sql = csql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc left JOIN " & baseVentas & ".sv_maestroclientes AS mc ON dc.rut = mc.rut AND mc.sucursal = '0'"
    csql.sql = csql.sql & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND " & tipo & " "
    csql.sql = csql.sql & "ORDER BY dc.foliosii "
    'Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    csql.Execute
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    If csql.RowsAffected > 0 Then
       impresion.Rows = csql.RowsAffected + 5
       Set resultados = csql.OpenResultset
        While Not resultados.EOF
           If existedte(empresaActiva, resultados("tipo"), resultados("numero"), resultados("fecha"), resultados("caja"), "0") = False Then
           linea = linea + 1
            impresion.Cell(linea, 0).text = resultados("caja") + resultados("numero")
            impresion.Cell(linea, 1).text = resultados("tipo")
            impresion.Cell(linea, 2).text = resultados("numero")
            impresion.Cell(linea, 3).text = resultados("fecha")
            impresion.Cell(linea, 4).text = resultados("rut")
            impresion.Cell(linea, 5).text = resultados("nombre")
            impresion.Cell(linea, 6).text = resultados("neto")
            impresion.Cell(linea, 7).text = resultados("iva")
            impresion.Cell(linea, 8).text = resultados("impuestoilarefrescos")
            impresion.Cell(linea, 9).text = resultados("impuestoilavinos")
            impresion.Cell(linea, 10).text = resultados("impuestoilalicores")
            impresion.Cell(linea, 11).text = resultados("impuestoharina")
            impresion.Cell(linea, 12).text = resultados("impuestocarne")
            impresion.Cell(linea, 13).text = resultados("exento")
            impresion.Cell(linea, 14).text = resultados("total")
            
            
            totales(1) = totales(1) + CDbl(resultados("neto"))
            totales(2) = totales(2) + CDbl(resultados("iva"))
            totales(3) = totales(3) + CDbl(resultados("exento"))
            totales(4) = totales(4) + CDbl(resultados("total"))
            totales(5) = totales(5) + CDbl(resultados("impuestoilarefrescos"))
            totales(6) = totales(6) + CDbl(resultados("impuestoilavinos"))
            totales(7) = totales(7) + CDbl(resultados("impuestoilalicores"))
            totales(8) = totales(8) + CDbl(resultados("impuestoharina"))
            totales(9) = totales(9) + CDbl(resultados("impuestocarne"))
            
            totales2(1) = totales2(1) + CDbl(resultados("neto"))
            totales2(2) = totales2(2) + CDbl(resultados("iva"))
            totales2(3) = totales2(3) + CDbl(resultados("exento"))
            totales2(4) = totales2(4) + CDbl(resultados("total"))
            totales2(5) = totales2(5) + CDbl(resultados("impuestoilarefrescos"))
            totales2(6) = totales2(6) + CDbl(resultados("impuestoilavinos"))
            totales2(7) = totales2(7) + CDbl(resultados("impuestoilalicores"))
            totales2(8) = totales2(8) + CDbl(resultados("impuestoharina"))
            totales2(9) = totales2(9) + CDbl(resultados("impuestocarne"))
            End If
            resultados.MoveNext
        Wend
    

    linea = linea + 1
            impresion.Range(linea, 5, linea, 14).Borders(cellEdgeTop) = cellThick
            impresion.Cell(linea, 5).text = "TOTALES GENERALES"
            impresion.Cell(linea, 6).text = totales2(1)
            impresion.Cell(linea, 7).text = totales2(2)
            impresion.Cell(linea, 8).text = totales2(5)
            impresion.Cell(linea, 9).text = totales2(6)
            impresion.Cell(linea, 10).text = totales2(7)
            impresion.Cell(linea, 11).text = totales2(8)
            impresion.Cell(linea, 12).text = totales2(9)
            impresion.Cell(linea, 13).text = totales2(3)
            impresion.Cell(linea, 14).text = totales2(4)
        
    End If
Set csql = Nothing
csql.Close
Set resultados = Nothing

    'Call sumaGrilla(impresion)
End Function

Private Sub sumaGrilla(ByRef impresion As Grid)
    Dim i As Long
    Dim j As Long
    Dim totalNeto As Double
    Dim totalIVA As Double
    Dim totalIHA As Double
    Dim total As Double
    Dim escribir As Boolean
    
    For i = 1 To impresion.Rows - 1
        totalNeto = 0
        totalIVA = 0
        totalIHA = 0
        total = 0
        For j = 4 To 9
            If IsNumeric(impresion.Cell(i, j).text) = True Then
                totalNeto = totalNeto + CDbl(impresion.Cell(i, j).text)
                escribir = True
            Else
                escribir = False
            End If
        Next j
        Select Case impresion.Cell(i, 0).text
            Case "1", "2"
                totalIVA = Round(totalNeto * iva / 100 + 0.1, 0)
                totalIHA = Round(totalNeto * iha / 100 + 0.1, 0)
            Case "3"
                totalIVA = Round(totalNeto / (1 + iva / 100) + 0.1, 0)
                totalIHA = 0
            Case "4"
                totalIVA = 0
                totalIHA = 0
        End Select
        total = totalNeto + totalIVA + totalIHA
        If escribir = True Then
            impresion.Cell(i, 10).text = totalNeto
            impresion.Cell(i, 11).text = totalIVA
            impresion.Cell(i, 12).text = totalIHA
            impresion.Cell(i, 13).text = total
        End If
    Next i
    totalIVA = 0
    totalNeto = 0
    total = 0
    For i = 1 To impresion.Rows - 1
        If IsNumeric(impresion.Cell(i, 11).text) = True And impresion.Cell(i, 0).text <> "" Then
            If impresion.Cell(i, 11).Font.Bold = True Then
                totalIVA = totalIVA + CDbl(impresion.Cell(i, 11).text)
                totalIHA = totalIHA + CDbl(impresion.Cell(i, 12).text)
                total = total + CDbl(impresion.Cell(i, 13).text)
            End If
        End If
    Next i
    impresion.Cell(i - 1, 11).text = totalIVA
    impresion.Cell(i - 1, 12).text = totalIHA
    impresion.Cell(i - 1, 13).text = total
End Sub

Private Function leerDocumentosLocal(ByVal numloc As String, ByVal rubAux As String, ByVal orden As String) As String
    
    Dim campos(2, 3) As String
    Dim op As Integer
    Dim tipo As String
    Set sql = New sqlventas.sqlventa
    campos(0, 0) = "IFNULL(COUNT(*),0)"
    campos(1, 0) = ""
    
    campos(0, 2) = baseVentas & rubAux & ".sv_documento_cabeza_" + empresaActiva + " AS dd"
    
    Select Case orden
        Case "1"
            tipo = "tipo = 'FV'"
        Case "2"
            tipo = "tipo = 'NV'"
        Case "3"
            tipo = "tipo = 'BV' OR tipo = 'ZE'"
        Case "4"
            tipo = "tipo = 'FE'"
    End Select
    condicion = "dd.local = '" & numloc & "' AND dd.nula = 'N' AND " & tipo & " "
    op = 5
    sql.response = campos
    Set sql.conexion = ventasRubro
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        leerDocumentosLocal = sql.response(0, 3)
    Else
        leerDocumentosLocal = ""
    End If
End Function






