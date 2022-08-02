Attribute VB_Name = "FGuiaMolienda"
Option Explicit
    Private CAMPOS(30, 3) As String
    
    Public Type GuiaMoliendas
        codLoc As String
        TIPO As String
        NUMERO As String
        FOLIO As String
        rut As String
        sucursal As String
        fecha As String
        trigo As String
        harina As String
        afrecho As String
        harinilla As String
        impurezas As String
        valor As String
        tipodocumento As String
        numeroDocumento As String
    End Type
    
    Public Type Moliendas
        FOLIO As String
        fecha As String
        rut As String
        sucursal As String
        trigo As String
        Descuento As String
        subproductos As String
        harina As String
        afrecho As String
        harinilla As String
        impurezas As String
        valor As String
    End Type

'=============================================================================
'LEER GUIA MOLIENDA
'=============================================================================
    Public Function leerGuiaMolienda(ByRef gm As GuiaMoliendas, ByVal NUMERO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "tipo"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "folio"
        CAMPOS(3, 0) = "rut"
        CAMPOS(4, 0) = "sucursal"
        CAMPOS(5, 0) = "fecha"
        CAMPOS(6, 0) = "trigo"
        CAMPOS(7, 0) = "harina"
        CAMPOS(8, 0) = "afrecho"
        CAMPOS(9, 0) = "harinilla"
        CAMPOS(10, 0) = "impurezas"
        CAMPOS(11, 0) = "valor"
        CAMPOS(12, 0) = "tipodocumento"
        CAMPOS(13, 0) = "numerodocumento"
        CAMPOS(14, 0) = ""

        CAMPOS(0, 2) = "sv_guiasmolienda"

        condicion = "local = '" & empresaActiva & "' AND tipo = '" & gm.TIPO & "' AND numero " & operador & " '" & NUMERO & "'"
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
            leerGuiaMolienda = True
            Call asignaGuia(gm, sql)
        Else
            leerGuiaMolienda = False
        End If
    End Function
'=============================================================================
'LEER GUIA MOLIENDA
'=============================================================================
    
'=============================================================================
'GRABAR GUIA MOLIENDA
'=============================================================================
    Public Sub grabarGuiaMolienda(ByRef gm As GuiaMoliendas, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designaGuia(gm, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & empresaActiva & "' AND tipo = 'GM' AND numero = '" & gm.NUMERO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR GUIA MOLIENDA
'=============================================================================

'=============================================================================
'ELIMINAR GUIA MOLIENDA
'=============================================================================
    Public Sub eliminarGuiaMolienda(ByRef gm As GuiaMoliendas)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & gm.TIPO & "' AND numero = '" & gm.NUMERO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_guiasmolienda"
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR GUIA MOLIENDA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asignaGuia(ByRef gm As GuiaMoliendas, ByRef sql As sqlventas.sqlventa)
        gm.TIPO = sql.response(0, 3)
        gm.NUMERO = sql.response(1, 3)
        gm.FOLIO = sql.response(2, 3)
        gm.rut = sql.response(3, 3)
        gm.sucursal = sql.response(4, 3)
        gm.fecha = sql.response(5, 3)
        gm.trigo = sql.response(6, 3)
        gm.harina = sql.response(7, 3)
        gm.afrecho = sql.response(8, 3)
        gm.harinilla = sql.response(9, 3)
        gm.impurezas = sql.response(10, 3)
        gm.valor = sql.response(11, 3)
        gm.tipodocumento = sql.response(12, 3)
        gm.numeroDocumento = sql.response(13, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designaGuia(ByRef gm As GuiaMoliendas, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "numero"
        CAMPOS(3, 0) = "folio"
        CAMPOS(4, 0) = "rut"
        CAMPOS(5, 0) = "sucursal"
        CAMPOS(6, 0) = "fecha"
        CAMPOS(7, 0) = "trigo"
        CAMPOS(8, 0) = "harina"
        CAMPOS(9, 0) = "afrecho"
        CAMPOS(10, 0) = "harinilla"
        CAMPOS(11, 0) = "impurezas"
        CAMPOS(12, 0) = "valor"
        CAMPOS(13, 0) = "tipodocumento"
        CAMPOS(14, 0) = "numerodocumento"
        CAMPOS(15, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = gm.TIPO
        CAMPOS(2, 1) = gm.NUMERO
        CAMPOS(3, 1) = gm.FOLIO
        CAMPOS(4, 1) = gm.rut
        CAMPOS(5, 1) = gm.sucursal
        CAMPOS(6, 1) = gm.fecha
        CAMPOS(7, 1) = gm.trigo
        CAMPOS(8, 1) = gm.harina
        CAMPOS(9, 1) = gm.afrecho
        CAMPOS(10, 1) = gm.harinilla
        CAMPOS(11, 1) = gm.impurezas
        CAMPOS(12, 1) = gm.valor
        CAMPOS(13, 1) = gm.tipodocumento
        CAMPOS(14, 1) = gm.numeroDocumento
        CAMPOS(15, 1) = ""

        CAMPOS(0, 2) = "sv_guiasmolienda"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    
'*************************************************************************************
'*************************************************************************************

'=============================================================================
'LEER MOLIENDA
'=============================================================================
    Public Function leerMolienda(ByRef m As Moliendas, ByVal FOLIO As String, ByRef operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "folio"
        CAMPOS(1, 0) = "fecha"
        CAMPOS(2, 0) = "rut"
        CAMPOS(3, 0) = "sucursal"
        CAMPOS(4, 0) = "trigo"
        CAMPOS(5, 0) = "harina"
        CAMPOS(6, 0) = "afrecho"
        CAMPOS(7, 0) = "harinilla"
        CAMPOS(8, 0) = "impurezas"
        CAMPOS(9, 0) = "valor"
        CAMPOS(10, 0) = ""

        CAMPOS(0, 2) = baseTrigo & ".moliendas"

        condicion = "folio " & operador & " '" & FOLIO & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY folio DESC"
        Else
            condicion = condicion & "ORDER BY folio ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerMolienda = True
            Call asigna(m, sql)
        Else
            leerMolienda = False
        End If
    End Function
'=============================================================================
'LEER MOLIENDA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef m As Moliendas, ByRef sql As sqlventas.sqlventa)
        m.FOLIO = sql.response(0, 3)
        m.fecha = sql.response(1, 3)
        m.rut = sql.response(2, 3)
        m.sucursal = sql.response(3, 3)
        m.trigo = sql.response(4, 3)
        m.harina = sql.response(5, 3)
        m.afrecho = sql.response(6, 3)
        m.harinilla = sql.response(7, 3)
        m.impurezas = sql.response(8, 3)
        m.valor = sql.response(9, 3)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================




