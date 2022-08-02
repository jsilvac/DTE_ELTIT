Attribute VB_Name = "FListadoStock"
Option Explicit

Public Sub generaInformeLS(ByRef data As Adodc, ByRef impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("LIBRO DE EXISTENCIAS DE HARINA Y SUBPRODUCTOS DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
    Call ListaStock(data, impresion, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Sub ListaStock(ByRef data As Adodc, ByRef impresion As Grid, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim tabla As String
    Dim cadena As String
    Dim fecha As String
    Dim producto As String
    Dim saldo As Double
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set cSql.ActiveConnection = ventasRubro
    cSql.sql = "SELECT tipoproducto, CONCAT(DATE_FORMAT(fecha,'%d-%m-%Y'), '" & vbTab & "', CASE tipo WHEN '1' THEN 'PRODUCCION' WHEN '2' THEN 'DEVOLUCIONES' WHEN '3' THEN 'COMPRAS' WHEN '4' THEN 'VENTAS' WHEN '5' THEN 'TRASLADO LOCAL' WHEN '6' THEN 'OTROS EGRESOS' END, '" & vbTab & "', monto, '" & vbTab & "', 0) AS item, DATE_FORMAT(fecha,'%d-%m-%Y') AS fecha, tipo, monto "
    cSql.sql = cSql.sql & "FROM produccion "
    cSql.sql = cSql.sql & "WHERE local = '" & empresaActiva & "' AND (tipo = '1' OR tipo = '2' OR tipo = '3') AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
    cSql.sql = cSql.sql & "UNION "
    cSql.sql = cSql.sql & "SELECT tipoproducto, CONCAT(DATE_FORMAT(fecha,'%d-%m-%Y'), '" & vbTab & "', CASE tipo WHEN '1' THEN 'PRODUCCION' WHEN '2' THEN 'DEVOLUCIONES' WHEN '3' THEN 'COMPRAS' WHEN '4' THEN 'VENTAS' WHEN '5' THEN 'TRASLADO LOCAL' WHEN '6' THEN 'OTROS EGRESOS' END, '" & vbTab & "', 0, '" & vbTab & "', monto) AS item, DATE_FORMAT(fecha,'%d-%m-%Y') AS fecha, tipo, -1 * monto AS monto "
    cSql.sql = cSql.sql & "FROM produccion "
    cSql.sql = cSql.sql & "WHERE local = '" & empresaActiva & "' AND (tipo = '4' OR tipo = '5' OR tipo = '6') AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
    cSql.sql = cSql.sql & "ORDER BY tipoproducto, fecha, tipo ASC "
    cSql.Execute
    
    
'    Call ConectarControlData(data, servidor, baseDatos & rubro, usuario, password, tabla)
  '  data.Recordset.Requery
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        producto = resultados("tipoproducto")
        
        impresion.AddItem "MES: " & UCase(Format(fecha1, "mmmm")) & " DE " & Format(fecha1, "yyyy") & "     PRODUCTO: HARINA", True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        
        fecha = DateAdd("m", -1, fecha1)
        saldo = leerSaldoStock(fecha1, "1")
        impresion.AddItem Format(fecha, "mm") & " DE " & Format(fecha, "yyyy") & vbTab & "SALDO ANTERIOR" & vbTab & vbTab & vbTab & saldo
        While Not resultados.EOF
            If producto <> resultados("tipoproducto") Then
                impresion.AddItem "", True
                impresion.RowHeight(impresion.Rows - 1) = 0
                Call impresion.HPageBreaks.Add(impresion.Rows - 1)
                impresion.AddItem "MES: " & UCase(Format(fecha1, "mmmm")) & " DE " & Format(fecha1, "yyyy") & "     PRODUCTO: SUBPRODUCTO", True
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                producto = resultados("tipoproducto")
                saldo = leerSaldoStock(fecha1, "2")
                impresion.AddItem Format(fecha, "mm") & " DE " & Format(fecha, "yyyy") & vbTab & "SALDO ANTERIOR" & vbTab & vbTab & vbTab & saldo
            End If
            cadena = resultados("item") & vbTab
            saldo = saldo + CDbl(resultados("monto"))
            cadena = cadena & saldo
            impresion.AddItem cadena, True
            resultados.MoveNext
        Wend
    End If
    Set cSql = Nothing
    cSql.Close
    Set resultados = Nothing
    
End Sub

Private Function leerSaldoStock(ByVal fecha As String, ByVal TIPO As String) As Double
    Dim monto1 As Double
    Dim monto2 As Double
    Dim cSql As rdoQuery
    Dim resultado As rdoResultset
    
    Set cSql = New rdoQuery
    Set cSql.ActiveConnection = gestionRubro
    cSql.sql = "SELECT IFNULL(SUM(monto), 0) "
    cSql.sql = cSql.sql & "FROM produccion "
    cSql.sql = cSql.sql & "WHERE local = '" & empresaActiva & "' AND fecha < '" & fecha & "' AND tipoproducto = '" & TIPO & "' AND (tipo = '1' OR tipo = '2' OR tipo = '3') "
    cSql.sql = cSql.sql & "GROUP BY tipoproducto "
    cSql.Execute
    If cSql.RowsAffected > 0 Then
        Set resultado = cSql.OpenResultset
        monto1 = resultado(0)
    Else
        monto1 = 0
    End If
    Set resultado = Nothing
    cSql.Close
    Set cSql = Nothing
    
    Set cSql = New rdoQuery
    Set cSql.ActiveConnection = gestionRubro
    cSql.sql = "SELECT IFNULL(SUM(monto), 0) "
    cSql.sql = cSql.sql & "FROM produccion "
    cSql.sql = cSql.sql & "WHERE local = '" & empresaActiva & "' AND fecha < '" & fecha & "' AND tipoproducto = '" & TIPO & "' AND (tipo = '4' OR tipo = '5' OR tipo = '6') "
    cSql.sql = cSql.sql & "GROUP BY tipoproducto "
    cSql.Execute
    If cSql.RowsAffected > 0 Then
        Set resultado = cSql.OpenResultset
        monto2 = resultado(0)
    Else
        monto2 = 0
    End If
    Set resultado = Nothing
    cSql.Close
    Set cSql = Nothing
    leerSaldoStock = monto1 - monto2
End Function


