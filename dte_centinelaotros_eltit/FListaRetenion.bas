Attribute VB_Name = "FListaRetenion"
Option Explicit

Public Sub generaInformeLR(ByRef data As Adodc, ByRef Impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    
    Call cargaCabezaRetencion("ANEXO INFORME MENSUAL VENDEDORES DE HARINA", empresaActiva, Impresion, fecha1)
    Call resumenVentas(data, Impresion, fecha1, fecha2)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Sub resumenVentas(ByRef data As Adodc, ByRef Impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim cadena As String
    Dim total As Double
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    
    cSql.sql = "SELECT dc.rut, dc.sucursal, @iha := IFNULL(@iha + SUM(dc.impuestoharina),0) AS iha, @iha:=0 "
    cSql.sql = cSql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc "
    cSql.sql = cSql.sql & "WHERE dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.nula = 'N' AND dc.impuestoharina <> 0 "
    cSql.sql = cSql.sql & "GROUP BY rut, sucursal "
    cSql.sql = cSql.sql & "ORDER BY rut, sucursal ASC"
    cSql.Execute
   
   ' Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
    'data.Recordset.Requery
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        total = 0
        While Not resultados.EOF
            cadena = resultados("rut")
            cadena = Left(cadena, 9) & "-" & Right(cadena, 1) & vbTab
            cadena = cadena & leerNombreClienteSucursal(resultados("rut"), resultados("sucursal")) & vbTab
            cadena = cadena & leerDireccionCliente(resultados("rut"), resultados("sucursal")) & ", " & leerComunaCliente(resultados("rut"), resultados("sucursal")) & vbTab
            cadena = cadena & Format(resultados("iha"), "$ ###,###,##0")
            Impresion.AddItem cadena, True
            total = total + CDbl(resultados("iha"))
            resultados.MoveNext
        Wend
        Impresion.AddItem "", True
        Impresion.AddItem vbTab & vbTab & "TOTAL RETENCION IVA (" & iha & " %)" & vbTab & Format(total, "$ ###,###,##0"), True
        Impresion.Range(Impresion.Rows - 1, 1, Impresion.Rows - 1, Impresion.Cols - 1).FontBold = True
        Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        Impresion.Range(Impresion.Rows - 1, 3, Impresion.Rows - 1, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        Impresion.Cell(Impresion.Rows - 1, 3).Alignment = cellCenterCenter
    End If
    Set cSql = Nothing
    cSql.Close
    Set resultados = Nothing
    
End Sub



