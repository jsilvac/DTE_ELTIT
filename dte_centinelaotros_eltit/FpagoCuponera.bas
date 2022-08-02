Attribute VB_Name = "FpagoCuponera"
Option Explicit
    Public saldo As Double
    Private CAMPOS(30, 3) As String
    Private Type PagoCuponeraCabeza
        local As String
        FOLIO As String
        rut As String
        tipodocumento As String
        numeroDocumento As String
        MONTO As String
    End Type

    Private Type PagoCuponeraCuota
        local As String
        FOLIO As String
        cuota As String
        MONTO As String
        venciminto As String
    End Type

    Public Type PagoCupon
        cabeza As PagoCuponeraCabeza
        cuota As PagoCuponeraCuota
    End Type

'=============================================================================
'LEER CUPONERA
'=============================================================================
    Public Function leerPagoCuponera(ByRef pc As PagoCupon, ByVal FOLIO As String, ByVal rut As String, ByRef lista As Grid, ByRef data As Adodc) As Boolean
        leerPagoCuponera = False
        If leerPagoCuponeraCabeza(pc.cabeza, FOLIO, rut) = True Then
            If leerPagoCuponeraCuota(pc.cuota, pc.cabeza.FOLIO) = True Then
                Call leerPagoCuponeraDetalle(pc.cabeza.FOLIO, lista, data)
                leerPagoCuponera = True
            Else
                lista.Rows = 1
            End If
        End If
    End Function

    Private Function leerPagoCuponeraCabeza(ByRef pcc As PagoCuponeraCabeza, ByVal FOLIO As String, ByVal rut As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "folio"
        CAMPOS(1, 0) = "rut"
        CAMPOS(2, 0) = "tipodoc"
        CAMPOS(3, 0) = "numerodoc"
        CAMPOS(4, 0) = "total - abono"
        CAMPOS(5, 0) = ""

        CAMPOS(0, 2) = "sv_credito_cabeza"
        If FOLIO <> "" And FOLIO <> "0000000000" Then
            condicion = "local = '" & empresaActiva & "' AND folio = '" & FOLIO & "'"
        Else
            condicion = "local = '" & empresaActiva & "' AND rut = '" & rut & "'"
        End If
        condicion = condicion & "ORDER BY folio ASC"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPagoCuponeraCabeza = True
            Call asignaCabeza(pcc, sql)
        Else
            leerPagoCuponeraCabeza = False
        End If
    End Function

    Private Function leerPagoCuponeraCuota(ByRef pcu As PagoCuponeraCuota, ByVal FOLIO As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "folio"
        CAMPOS(1, 0) = "cuota"
        CAMPOS(2, 0) = "montocuota"
        CAMPOS(3, 0) = "vencimiento"
        CAMPOS(4, 0) = ""

        CAMPOS(0, 2) = "sv_credito_detalle"

        condicion = "local = '" & empresaActiva & "' AND folio = '" & FOLIO & "' AND abonocuota = '0' "
        condicion = condicion & "ORDER BY vencimiento ASC"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPagoCuponeraCuota = True
            Call asignaCuota(pcu, sql)
        Else
            leerPagoCuponeraCuota = False
        End If
    End Function

    Private Sub leerPagoCuponeraDetalle(ByVal FOLIO As String, ByRef lista As Grid, ByRef data As Adodc)
        Dim tabla As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT CONCAT(cuota, '" & vbTab & "', DATE_FORMAT(vencimiento,'%d-%m-%Y'), '" & vbTab & "', DATE_FORMAT(fechapago,'%d-%m-%Y'), '" & vbTab & "', montocuota, '" & vbTab & "', montocuota - abonocuota) AS item, montocuota - abonocuota AS saldo "
        csql.sql = csql.sql & "FROM sv_credito_detalle "
        csql.sql = csql.sql & "WHERE local = '" & empresaActiva & "' AND folio = '" & FOLIO & "' "
        csql.sql = csql.sql & "ORDER BY cuota ASC"
        csql.Execute
        
        'Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
        lista.Rows = 1
        lista.AutoRedraw = False
        saldo = 0
        If csql.RowsAffected > 0 Then
           Set resultados = csql.OpenResultset
            While Not resultados.EOF
                saldo = saldo + resultados("saldo")
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

'    Public Function leerDocumento(ByRef cc As PagoCuponeraCabeza, ByVal rut As String, ByVal tipo As String, ByVal numero As String) As Boolean
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        campos(0, 0) = "total"
'        campos(1, 0) = "abono"
'        campos(2, 0) = "fecha"
'        campos(3, 0) = ""''''
'
'        campos(0, 2) = "sv_documento_cabeza"''

'        condicion = "local = '" & empresaactiva & "' AND rut = '" & rut & "' AND tipo = '" & tipo & "' AND numero = '" & numero & "'"
'        op = 5
'        sql.response = campos
'        Set sql.conexion = ventasRubro
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerDocumento = True
'            Call asignaDocumento(cc, sql)
'        Else
'            leerDocumento = False
'        End If
'    End Function
'=============================================================================
'LEER CUPONERA
'=============================================================================

''=============================================================================
''GRABAR CUPONERA
''=============================================================================
'    Public Sub grabarCuponera(ByRef c As Cuponera, ByVal modifica As Boolean, ByRef lista As Grid)
'        Call grabarCuponeraCabeza(c.cabeza, modifica)
'        Call grabarCuponeraDetalle(lista, c.detalle, modifica)
'    End Sub
'
'    Private Sub grabarCuponeraCabeza(ByRef cc As CuponeraCabeza, ByVal modifica As Boolean)
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        Call designaCabeza(cc, sql)
'        condicion = ""
'        If modifica = True Then
'            condicion = "local = '" & cc.local & "' AND folio = '" & cc.folio & "'"
'            op = 3
'        Else
'            op = 2
'        End If
'        Set sql.conexion = ventasRubro
'        Call sql.sqlventas(op, condicion)
'        op = sql.status
'    End Sub
'
'    Private Sub grabarCuponeraDetalle(ByRef lista As Grid, ByRef cd As CuponeraDetalle, ByVal modifica As Boolean)
'
'        Dim op As Integer
'        Dim i As Long
'        Dim cuota As String
'        Set sql =new sqlventas.sqlventa
'
'        condicion = ""
'        If modifica = True Then
'            'Call eliminarCuponeraDetalle(cd, lista)
'        End If
'        op = 2
'
'        cd.local = empresaactiva
'        For i = 1 To lista.Rows - 1
'            cuota = Str(i)
'            cuota = Mid(cuota, 2, Len(cuota))
'            cuota = String(3 - Len(cuota), "0") & cuota
'            If lista.Cell(i, 2).text <> "" And lista.Cell(i, 3).text <> "" And lista.Cell(i, 4).text <> "" And lista.Cell(i, 5).text <> "" And lista.Cell(i, 6).text <> "" Then
'                cd.folio = lista.Cell(i, 2).text
'                cd.cuota = cuota
'                cd.vencimiento = Format(lista.Cell(i, 4).text, "yyyy-mm-dd")
'                cd.montocuota = lista.Cell(i, 5).text
'                cd.abonoCuota = lista.Cell(i, 6).text
'                cd.interesmora = ""
'
'                Call designaDetalle(cd, sql)
'
'                Set sql.conexion = ventasRubro
'                Call sql.sqlventas(op, condicion)
'            End If
'        Next i
'    End Sub
''=============================================================================
''GRABAR CUPONERA
''=============================================================================
'
'''=============================================================================
'''ELIMINAR CLIENTE
'''=============================================================================
''    Public Sub eliminarCliente(ByRef c As Cliente)
''
''        Dim op As Integer
''        Set sql =new sqlventas.sqlventa
''        condicion = "rut = '" & c.rut & "' AND sucursal = '" & c.sucursal & "'"
''        op = 4
''        campos(0, 2) = "sv_maestroclientes"
''        sql.response = campos
''        Set sql.conexion = ventasRubro
''        Call sql.sqlventas(op, condicion)
''    End Sub
'''=============================================================================
'''ELIMINAR CLIENTE
'''=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asignaCabeza(ByRef pcc As PagoCuponeraCabeza, ByRef sql As sqlventas.sqlventa)
        pcc.local = empresaActiva
        pcc.FOLIO = sql.response(0, 3)
        pcc.rut = sql.response(1, 3)
        pcc.tipodocumento = sql.response(2, 3)
        pcc.numeroDocumento = sql.response(3, 3)
        pcc.MONTO = sql.response(4, 3)
    End Sub
    
    Private Sub asignaCuota(ByRef pcu As PagoCuponeraCuota, ByRef sql As sqlventas.sqlventa)
        pcu.local = empresaActiva
        pcu.FOLIO = sql.response(0, 3)
        pcu.cuota = sql.response(1, 3)
        pcu.MONTO = sql.response(2, 3)
        pcu.venciminto = sql.response(3, 3)
    End Sub

'    Private Sub asignaDetalle(ByRef cd As CuponeraDetalle, ByRef sql as sqlventas.sqlventa)
'        'cc.folio = sql.response(0, 3)
'        'cc.rut = sql.response(1, 3)
'        'cc.tipoDocumento = sql.response(2, 3)
'        'cc.numeroDocumento = sql.response(3, 3)
'        'cc.fechaCompra = sql.response(4, 3)
'        'cc.total = sql.response(5, 3)
'        'cc.cuotas = sql.response(6, 3)
'    End Sub
'
'    Private Sub asignaDocumento(ByRef cc As CuponeraCabeza, ByRef sql as sqlventas.sqlventa)
'        cc.total = sql.response(0, 3)
'        cc.abono = sql.response(1, 3)
'        cc.fechaCompra = sql.response(2, 3)
'    End Sub
''=============================================================================
''PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
''=============================================================================
'
''=============================================================================
''PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
''=============================================================================
'    Private Sub designaCabeza(ByRef cc As CuponeraCabeza, ByRef sql as sqlventas.sqlventa)
'        Dim cad As String
'        campos(0, 0) = "local"
'        campos(1, 0) = "folio"
'        campos(2, 0) = "rut"
'        campos(3, 0) = "tipodoc"
'        campos(4, 0) = "numerodoc"
'        campos(5, 0) = "fechacompra"
'        campos(6, 0) = "total"
'        campos(7, 0) = "abono"
'        campos(8, 0) = "cuotas"
'        campos(9, 0) = ""
'
'        campos(0, 1) = cc.local
'        campos(1, 1) = cc.folio
'        campos(2, 1) = cc.rut
'        campos(3, 1) = cc.tipoDocumento
'        campos(4, 1) = cc.numeroDocumento
'        campos(5, 1) = cc.fechaCompra
'        campos(6, 1) = cc.total
'        campos(7, 1) = cc.abono
'        campos(8, 1) = cc.cuotas
'        campos(9, 1) = ""
'
'        campos(0, 2) = "sv_credito_cabeza"
'        sql.response = campos
'    End Sub
'
'    Private Sub designaDetalle(ByRef cd As CuponeraDetalle, ByRef sql as sqlventas.sqlventa)
'        Dim cad As String
'        campos(0, 0) = "local"
'        campos(1, 0) = "folio"
'        campos(2, 0) = "cuota"
'        campos(3, 0) = "vencimiento"
'        campos(4, 0) = "montocuota"
'        campos(5, 0) = "abonocuota"
'        campos(6, 0) = "interesmora"
'        campos(7, 0) = ""
'
'        campos(0, 1) = cd.local
'        campos(1, 1) = cd.folio
'        campos(2, 1) = cd.cuota
'        campos(3, 1) = cd.vencimiento
'        campos(4, 1) = cd.montocuota
'        campos(5, 1) = cd.abonoCuota
'        campos(6, 1) = cd.interesmora
'        campos(7, 1) = ""
'
'        campos(0, 2) = "sv_credito_detalle"
'        sql.response = campos
'    End Sub
''=============================================================================
''PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
''=============================================================================
'
'
