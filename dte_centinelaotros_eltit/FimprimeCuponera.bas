Attribute VB_Name = "FimprimeCuponera"
Option Explicit
    Private CAMPOS(30, 3) As String
    Private Type CuponeraCabeza
        local As String
        FOLIO As String
        rut As String
        tipodocumento As String
        numeroDocumento As String
        fechacompra As String
        total As String
        abono As String
        CUOTAS As String
    End Type
    
    Private Type CuponeraDetalle
        local As String
        FOLIO As String
        cuota As String
        vencimiento As String
        montocuota As String
        abonoCuota As String
        interesmora As String
        fechaPago As String
    End Type

    Public Type Cuponera
        cabeza As CuponeraCabeza
        detalle As CuponeraDetalle
    End Type
    
'=============================================================================
'LEER CUPONERA
'=============================================================================
    Public Function leerCuponera(ByRef c As Cuponera, ByVal codigo1 As String, ByVal codigo2 As String, ByVal codigo3 As String, ByVal operador As String, ByRef lista As Grid, data As Adodc) As Boolean
        leerCuponera = False
        If leerCuponeraCabeza(c.cabeza, codigo1, codigo2, codigo3, operador) = True Then
            Call leerCuponeraDetalle(c.cabeza, lista, data)
            leerCuponera = True
        End If
    End Function
    
    Private Function leerCuponeraCabeza(ByRef cc As CuponeraCabeza, ByVal codigo1 As String, ByVal codigo2 As String, ByVal codigo3 As String, ByRef operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "folio"
        CAMPOS(1, 0) = "rut"
        CAMPOS(2, 0) = "tipodoc"
        CAMPOS(3, 0) = "numerodoc"
        CAMPOS(4, 0) = "fechacompra"
        CAMPOS(5, 0) = "total"
        CAMPOS(6, 0) = "abono"
        CAMPOS(7, 0) = "cuotas"
        CAMPOS(8, 0) = ""

        CAMPOS(0, 2) = "sv_credito_cabeza"

        condicion = "rut = '" & codigo1 & "' AND tipodoc = '" & codigo2 & "' AND numerodoc " & operador & " '" & codigo3 & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY numerodoc DESC"
        Else
            condicion = condicion & "ORDER BY numerodoc ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCuponeraCabeza = True
            Call asignaCabeza(cc, sql)
        Else
            leerCuponeraCabeza = False
        End If
    End Function
    
    Private Sub leerCuponeraDetalle(ByRef cc As CuponeraCabeza, ByRef lista As Grid, ByRef data As Adodc)
        Dim tabla As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        
        Set csql.ActiveConnection = ventasRubro
        
        csql.sql = "SELECT CONCAT('1', '" & vbTab & "', folio, '" & vbTab & "', cuota, '" & vbTab & "', DATE_FORMAT(vencimiento,'%d-%m-%Y'), '" & vbTab & "', montocuota, '" & vbTab & "', abonocuota) AS item "
        csql.sql = csql.sql & "FROM sv_credito_detalle "
        csql.sql = csql.sql & "WHERE folio = '" & cc.FOLIO & "' "
        csql.sql = csql.sql & "ORDER BY cuota ASC"
        csql.Execute
        'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
        lista.Rows = 1
        lista.AutoRedraw = False
        If csql.RowsAffected > 0 Then
           Set resultados = csql.OpenResultset
            While Not resultados.EOF
                lista.AddItem resultados("item"), True
            resultados.MoveNext
            Wend
        End If
        Set csql = Nothing
        csql.Close
        Set resultados = Nothing
        
        lista.AutoRedraw = True
        lista.Refresh
    End Sub
    
    Public Function LEERDOCUMENTO(ByRef cc As CuponeraCabeza, ByVal rut As String, ByVal TIPO As String, ByVal NUMERO As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "total"
        CAMPOS(1, 0) = "abono"
        CAMPOS(2, 0) = "fecha"
        CAMPOS(3, 0) = ""

        CAMPOS(0, 2) = "sv_documento_cabeza_" + empresaActiva

        condicion = "local = '" & empresaActiva & "' AND rut = '" & rut & "' AND tipo = '" & TIPO & "' AND numero = '" & NUMERO & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            LEERDOCUMENTO = True
            Call asignaDocumento(cc, sql)
        Else
            LEERDOCUMENTO = False
        End If
    End Function
'=============================================================================
'LEER CUPONERA
'=============================================================================

'=============================================================================
'GRABAR CUPONERA
'=============================================================================
    Public Sub grabarCuponera(ByRef c As Cuponera, ByVal modifica As Boolean, ByRef lista As Grid)
        Call grabarCuponeraCabeza(c.cabeza, modifica)
        Call grabarCuponeraDetalle(lista, c.detalle, modifica)
    End Sub
    
    Private Sub grabarCuponeraCabeza(ByRef cc As CuponeraCabeza, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designaCabeza(cc, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & cc.local & "' AND folio = '" & cc.FOLIO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        op = sql.Status
    End Sub
    
    Private Sub grabarCuponeraDetalle(ByRef lista As Grid, ByRef cd As CuponeraDetalle, ByVal modifica As Boolean)
        
        Dim op As Integer
        Dim i As Long
        Dim cuota As String
        Set sql = New sqlventas.sqlventa
        
        condicion = ""
        If modifica = True Then
            'Call eliminarCuponeraDetalle(cd, lista)
        End If
        op = 2
        
        cd.local = empresaActiva
        For i = 1 To lista.Rows - 1
            cuota = Str(i)
            cuota = Mid(cuota, 2, Len(cuota))
            cuota = String(3 - Len(cuota), "0") & cuota
            If lista.Cell(i, 2).text <> "" And lista.Cell(i, 3).text <> "" And lista.Cell(i, 4).text <> "" And lista.Cell(i, 5).text <> "" And lista.Cell(i, 6).text <> "" Then
                cd.FOLIO = lista.Cell(i, 2).text
                cd.cuota = cuota
                cd.vencimiento = Format(lista.Cell(i, 4).text, "yyyy-mm-dd")
                cd.montocuota = lista.Cell(i, 5).text
                cd.abonoCuota = lista.Cell(i, 6).text
                cd.interesmora = ""
                
                Call designaDetalle(cd, sql)
                
                Set sql.conexion = ventasRubro
                Call sql.sqlventas(op, condicion)
            End If
        Next i
    End Sub
'=============================================================================
'GRABAR CUPONERA
'=============================================================================

''=============================================================================
''ELIMINAR CLIENTE
''=============================================================================
'    Public Sub eliminarCliente(ByRef c As Cliente)
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        condicion = "rut = '" & c.rut & "' AND sucursal = '" & c.sucursal & "'"
'        op = 4
'        campos(0, 2) = "sv_maestroclientes"
'        sql.response = campos
'        Set sql.conexion = ventasRubro
'        Call sql.sqlventas(op, condicion)
'    End Sub
''=============================================================================
''ELIMINAR CLIENTE
''=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asignaCabeza(ByRef cc As CuponeraCabeza, ByRef sql As sqlventas.sqlventa)
        cc.FOLIO = sql.response(0, 3)
        cc.rut = sql.response(1, 3)
        cc.tipodocumento = sql.response(2, 3)
        cc.numeroDocumento = sql.response(3, 3)
        cc.fechacompra = sql.response(4, 3)
        cc.total = sql.response(5, 3)
        cc.abono = sql.response(6, 3)
        cc.CUOTAS = sql.response(7, 3)
    End Sub
    
    Private Sub asignaDetalle(ByRef cd As CuponeraDetalle, ByRef sql As sqlventas.sqlventa)
        'cc.folio = sql.response(0, 3)
        'cc.rut = sql.response(1, 3)
        'cc.tipoDocumento = sql.response(2, 3)
        'cc.numeroDocumento = sql.response(3, 3)
        'cc.fechaCompra = sql.response(4, 3)
        'cc.total = sql.response(5, 3)
        'cc.cuotas = sql.response(6, 3)
    End Sub
    
    Private Sub asignaDocumento(ByRef cc As CuponeraCabeza, ByRef sql As sqlventas.sqlventa)
        cc.total = sql.response(0, 3)
        cc.abono = sql.response(1, 3)
        cc.fechacompra = sql.response(2, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designaCabeza(ByRef cc As CuponeraCabeza, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "folio"
        CAMPOS(2, 0) = "rut"
        CAMPOS(3, 0) = "tipodoc"
        CAMPOS(4, 0) = "numerodoc"
        CAMPOS(5, 0) = "fechacompra"
        CAMPOS(6, 0) = "total"
        CAMPOS(7, 0) = "abono"
        CAMPOS(8, 0) = "cuotas"
        CAMPOS(9, 0) = ""
        
        CAMPOS(0, 1) = cc.local
        CAMPOS(1, 1) = cc.FOLIO
        CAMPOS(2, 1) = cc.rut
        CAMPOS(3, 1) = cc.tipodocumento
        CAMPOS(4, 1) = cc.numeroDocumento
        CAMPOS(5, 1) = cc.fechacompra
        CAMPOS(6, 1) = cc.total
        CAMPOS(7, 1) = cc.abono
        CAMPOS(8, 1) = cc.CUOTAS
        CAMPOS(9, 1) = ""
        
        CAMPOS(0, 2) = "sv_credito_cabeza"
        sql.response = CAMPOS
    End Sub
    
    Private Sub designaDetalle(ByRef cd As CuponeraDetalle, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "folio"
        CAMPOS(2, 0) = "cuota"
        CAMPOS(3, 0) = "vencimiento"
        CAMPOS(4, 0) = "montocuota"
        CAMPOS(5, 0) = "abonocuota"
        CAMPOS(6, 0) = "interesmora"
        CAMPOS(7, 0) = ""
        
        CAMPOS(0, 1) = cd.local
        CAMPOS(1, 1) = cd.FOLIO
        CAMPOS(2, 1) = cd.cuota
        CAMPOS(3, 1) = cd.vencimiento
        CAMPOS(4, 1) = cd.montocuota
        CAMPOS(5, 1) = cd.abonoCuota
        CAMPOS(6, 1) = cd.interesmora
        CAMPOS(7, 1) = ""
        
        CAMPOS(0, 2) = "sv_credito_detalle"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================

