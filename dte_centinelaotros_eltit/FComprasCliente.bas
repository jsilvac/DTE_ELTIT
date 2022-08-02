Attribute VB_Name = "FComprasCliente"
Option Explicit
    Public nombreCliente As String

Public Sub generaInformeHCK(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal rut As String, ByVal sucursal As String)
    Dim i As Long
    Impresion.Rows = 1
    Impresion.AutoRedraw = False
    
    Call cargaCabeza("HISTORICO DE COMPRA DE CLIENTES POR KILOS", empresaActiva, Impresion)
    Call resumenCompras(data, Impresion, TIPO, rut, sucursal)
    
    Impresion.AutoRedraw = True
    Impresion.Refresh
End Sub

Private Sub resumenCompras(ByRef data As Adodc, ByRef Impresion As Grid, ByVal TIPO As String, ByVal rutCliente As String, ByVal sucCliente As String)
    Dim tabla As String
    Dim cadena As String
    Dim año As String
    Dim pos As Integer
    Dim mes As Integer
    Dim rut As String
    Dim sucursal As String
    Dim sumaMeses(2, 13) As String
    Dim i As Integer
    Dim j As Integer
    Dim fecha As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim fecha3 As String
    Dim fecha4 As String
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Set cSql.ActiveConnection = ventasRubro
    
    fecha = Format(DateSerial(Year(fechasistema) - 1, Month(fechasistema), 1), "yyyy-mm-dd")
    primerAño = Val(Year(fecha))
    fecha1 = Format(DateSerial(Year(fecha), 1, 1), "yyyy-mm-dd")
    fecha2 = Format(DateSerial(Year(fecha), Month(fecha) + 1, 0), "yyyy-mm-dd")
    
    fecha = Format(DateSerial(Year(fechasistema), Month(fechasistema), 1), "yyyy-mm-dd")
    segundoAño = Val(Year(fecha))
    fecha3 = Format(DateSerial(Year(fecha), 1, 1), "yyyy-mm-dd")
    fecha4 = Format(DateSerial(Year(fecha), Month(fecha) + 1, 0), "yyyy-mm-dd")
    
    cSql.sql = "SELECT DATE_FORMAT(dd.fecha,'%Y') AS año, DATE_FORMAT(dd.fecha,'%m') AS mes, dd.rut, dd.sucursal, SUM(dd.unidades) AS unidades "
    cSql.sql = cSql.sql & "FROM sv_documento_detalle_" + empresaActiva + " AS dd "
    cSql.sql = cSql.sql & "WHERE (dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "') OR (dd.fecha BETWEEN '" & fecha3 & "' AND '" & fecha4 & "') "
    If rutCliente <> "" Then
        cSql.sql = cSql.sql & "AND dd.rut = '" & rutCliente & "' AND dd.sucursal = '" & sucCliente & "' "
    End If
    cSql.sql = cSql.sql & "GROUP BY rut, sucursal, año, mes "
    cSql.sql = cSql.sql & "ORDER BY rut, sucursal, año, mes ASC "
    cSql.Execute
    
    'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        
        
        For j = 0 To cantMeses + 1
            sumaMeses(0, j) = "0"
            sumaMeses(1, j) = "0"
            sumaMeses(2, j) = "0"
        Next j
        i = 1
        rut = resultados("rut")
        sucursal = resultados("sucursal")
        While Not resultados.EOF
            If rut = resultados("rut") And sucursal = resultados("sucursal") Then
                If resultados("año") = primerAño Then
                    pos = 0
                Else
                    pos = 1
                End If
                mes = Val(resultados("mes"))
                sumaMeses(pos, mes) = Format(resultados("unidades"), "###,###,##0")
                sumaMeses(pos, cantMeses + 1) = CDbl(sumaMeses(pos, cantMeses + 1)) + CDbl(resultados("unidades"))
            Else
                cadena = rut & sucursal & vbTab
                cadena = cadena & primerAño & vbTab
                For j = 1 To cantMeses + 1
                    cadena = cadena & sumaMeses(0, j) & vbTab
                Next j
                Impresion.AddItem cadena, True
                cadena = leerNombreClienteSucursal(rut, sucursal) & vbTab
                cadena = cadena & segundoAño & vbTab
                For j = 1 To cantMeses + 1
                    cadena = cadena & sumaMeses(1, j) & vbTab
                Next j
                Impresion.AddItem cadena, True
                Impresion.AddItem "", True
                
                For j = 0 To cantMeses + 1
                    sumaMeses(0, j) = "0"
                    sumaMeses(1, j) = "0"
                Next j
                rut = resultados("rut")
                sucursal = resultados("sucursal")
'                data.Recordset.MovePrevious
            End If
            resultados.MoveNext
        Wend
        cadena = rut & sucursal & vbTab
        cadena = cadena & primerAño & vbTab
        For j = 1 To cantMeses + 1
            cadena = cadena & sumaMeses(0, j) & vbTab
        Next j
        Impresion.AddItem cadena, True
        cadena = leerNombreClienteSucursal(rut, sucursal) & vbTab
        cadena = cadena & segundoAño & vbTab
        For j = 1 To cantMeses + 1
            cadena = cadena & sumaMeses(1, j) & vbTab
        Next j
        Impresion.AddItem cadena, True
    End If
End Sub
