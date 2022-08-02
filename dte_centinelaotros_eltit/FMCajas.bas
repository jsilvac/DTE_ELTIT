Attribute VB_Name = "FMCajas"
Option Explicit
    Private CAMPOS(30, 3) As String
    Public Type caja
        codLoc As String
        NUMERO As String
        descripcion As String
        folioboletas As String
        foliofacturas As String
        folionotadebito As String
        folionotacredito As String
        folioboletafiscal As String
        folioboletaelectronica As String
        foliofacturaelectronica As String
        folionotadebitoelectronica As String
        folionotacreditoelectronica As String
        foliocomprobantepagos As String
        
        
    
    End Type

'=============================================================================
'LEER CAJA
'=============================================================================
    Public Function leerCaja(ByRef c As caja, ByVal codLoc As String, ByVal CODIGO As String, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "descripcion"
        CAMPOS(3, 0) = "folioboletas"
        CAMPOS(4, 0) = "foliofacturas"
        CAMPOS(5, 0) = "folionotadebito"
        CAMPOS(6, 0) = "folionotacredito"
        CAMPOS(7, 0) = "folioboletafiscal"
        CAMPOS(8, 0) = "folioboletaelectronica"
        CAMPOS(9, 0) = "foliofacturaelectronica"
        CAMPOS(10, 0) = "folionotadebitoelectronica"
        CAMPOS(11, 0) = "folionotacreditoelectronica"
        CAMPOS(12, 0) = "foliocomprobantepagos"
        
        CAMPOS(13, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrodecajas"
        
        condicion = "local = '" & codLoc & "' AND numero " & operador & " '" & CODIGO & "' "
        If operador = "<" Then
            condicion = condicion & "ORDER BY numero DESC"
        Else
            condicion = condicion & "ORDER BY numero ASC"
        End If
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCaja = True
            Call asigna(c, sql)
        Else
            leerCaja = False
        End If
    End Function
'=============================================================================
'LEER CAJA
'=============================================================================

'=============================================================================
'GRABAR CAJA
'=============================================================================
    Public Sub grabarCaja(ByRef c As caja, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(c, sql)
        condicion = ""
        If modifica = True Then
            condicion = "local = '" & c.codLoc & "' AND numero = '" & c.NUMERO & "'"
            op = 3
        Else
            op = 2
        End If
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = MCajas.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR CAJA
'=============================================================================

'=============================================================================
'ELIMINAR CAJA
'=============================================================================
    Public Sub eliminarCaja(ByRef c As caja)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & c.codLoc & "' AND numero = '" & c.NUMERO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_maestrodecajas"
        sql.response = CAMPOS
        Set sql.conexion = ventas
        sql.audit = True:    sql.programaactivo = MCajas.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion

        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR CAJA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
    Private Sub asigna(ByRef c As caja, ByRef sql As sqlventas.sqlventa)
        c.codLoc = sql.response(0, 3)
        c.NUMERO = sql.response(1, 3)
        c.descripcion = sql.response(2, 3)
        c.folioboletas = sql.response(3, 3)
        c.foliofacturas = sql.response(4, 3)
        c.folionotacredito = sql.response(5, 3)
        c.folionotadebito = sql.response(6, 3)
        c.folioboletafiscal = sql.response(7, 3)
        c.folioboletaelectronica = sql.response(8, 3)
        c.foliofacturaelectronica = sql.response(9, 3)
        c.folionotadebitoelectronica = sql.response(10, 3)
        c.folionotacreditoelectronica = sql.response(11, 3)
        c.foliocomprobantepagos = sql.response(12, 3)
        
        
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef c As caja, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "descripcion"
        CAMPOS(3, 0) = "folioboletas"
        CAMPOS(4, 0) = "foliofacturas"
        CAMPOS(5, 0) = "folionotadebito"
        CAMPOS(6, 0) = "folionotacredito"
        CAMPOS(7, 0) = "folioboletafiscal"
        CAMPOS(8, 0) = "folioboletaelectronica"
        CAMPOS(9, 0) = "foliofacturaelectronica"
        CAMPOS(10, 0) = "folionotadebitoelectronica"
        CAMPOS(11, 0) = "folionotacreditoelectronica"
        CAMPOS(12, 0) = "foliocomprobantepagos"
        
        CAMPOS(13, 0) = ""
        
        
        CAMPOS(0, 1) = c.codLoc
        CAMPOS(1, 1) = c.NUMERO
        CAMPOS(2, 1) = c.descripcion
        CAMPOS(3, 1) = c.folioboletas
        CAMPOS(4, 1) = c.foliofacturas
        CAMPOS(5, 1) = c.folionotadebito
        CAMPOS(6, 1) = c.folionotacredito
        CAMPOS(7, 1) = c.folioboletafiscal
        CAMPOS(8, 1) = c.folioboletaelectronica
        CAMPOS(9, 1) = c.foliofacturaelectronica
        CAMPOS(10, 1) = c.folionotadebitoelectronica
        CAMPOS(11, 1) = c.folionotacreditoelectronica
        CAMPOS(12, 1) = c.foliocomprobantepagos
        
        CAMPOS(0, 2) = "sv_maestrodecajas"
        sql.response = CAMPOS
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================



