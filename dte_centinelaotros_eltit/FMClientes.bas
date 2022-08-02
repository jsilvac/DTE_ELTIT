Attribute VB_Name = "FMClientes"
Option Explicit
    Private campos(45, 3) As String
'    Public Type cliente
'        rut As String
'        sucursal As String
'        nombre As String
'        direccion As String
'        comuna As String
'        ciudad As String
'        fono1 As String
'        fono2 As String
'        fax As String
'        celular As String
'        giro As String
'        email As String
'        contacto As String
'        diapago As String
'        plazo As String
'        credito As String
'        descuento As String
'        cupoDirecto As String
'        cupoutilizado As String
'        protestos As String
'        prorrogas As String
'        boletas As String
'        facturas As String
'        Cheques As String
'        listaprecios As String
'        observaciones As String
'        vendedor As String
'      End Type
      
       Public Type Cliente
        rut As String
        sucursal As String
        nombre As String
        direccion As String
        comuna As String
        ciudad As String
        fono1 As String
        fono2 As String
        fax As String
        celular As String
        giro As String
        email As String
        casilla As String
        contacto As String
        DIAPAGO As String
        plazo As String
        CREDITO As String
        creditoTMP As String
        creditoDirecto As String
        Descuento As String
        cupotmp As String
        cupodirecto As String
        Zona As String
        cupoutilizadoTMP As String
        cupoutilizadoDirecto As String
        CUPOUTILIZADO As String
        protestos As String
        boletas As String
        facturas As String
        zetas As String
        Cheques As String
        prorrogas As String
        garantias As String
        morosidad As String
        cantprotestos As String
        cantboletas As String
        cantfacturas As String
        cantzetas As String
        cantcheques As String
        cantprorrogas As String
        cantgarantias As String
        cantmorosidad As String
        terceraedad As String
        TIPOCLIENTE As String
        bloqueoTMP As String
        bloqueoDirecto As String
        listaprecios As String
        observaciones As String
        vendedor As String
        rebajainterespp As String
        rebajainteresmora As String
        repactacion As String
        CUOTASREPACTACION As String
        CARTA As String
        constructora As String
        
    End Type
    
         
    Public Type personales
        rut As String
        sucursal As String
        fechanacimiento As String
        sexo As String
        nacionalidad As String
        estadocivil As String
        rutconyuge As String
        nombreconyuge As String
    End Type
    
    Public Type laborales
        rut As String
        sucursal As String
        labor As String
        rutempleador As String
        nombre As String
        direccion As String
        comuna As String
        ciudad As String
        fono As String
        antiguedad As String
        codeudor As String
    End Type
    
    Public Type adicionales
        rut As String
        sucursal As String
        rutadicional As String
        nombre As String
        porcentajecupo As String
    End Type
    
    Public Type financieros
        rut As String
        sucursal As String
        ingresomensual As String
        pagoscasascomerciales As String
        arriendo As String
        tipovivienda As String
        tasacionvivienda As String
        vehiculos As String
        tasacionvehiculos As String
        cuentacorriente As String
        Banco As String
        numerocuenta As String
        antuguedad As String
        otrastarjetas As String
        otratarjeta1 As String
        otratarjeta2 As String
        otratarjeta3 As String
        otratarjetacupo1 As String
        otratarjetacupo2 As String
        otratarjetacupo3 As String
        fechaimpresionpagare As String
        fechaautorizacioncredito As String
        autorizador As String
        fechaentregatarjeta As String
    End Type
    
    Public Type Clientes
        cc As Cliente
        cp As personales
        cl As laborales
        ca As adicionales
        cf As financieros
    End Type

    
'=============================================================================
'LEER CLIENTE
'=============================================================================
    Public Function leerCliente(ByRef c As Cliente, ByVal codigo1 As String, ByVal codigo2 As String, ByRef operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "nombre"
        campos(3, 0) = "direccion"
        campos(4, 0) = "comuna"
        campos(5, 0) = "ciudad"
        campos(6, 0) = "fono1"
        campos(7, 0) = "fono2"
        campos(8, 0) = "fax"
        campos(9, 0) = "celular"
        campos(10, 0) = "giro"
        campos(11, 0) = "email"
        campos(12, 0) = "contacto"
        campos(13, 0) = "plazo"
        campos(14, 0) = "credito"
        campos(15, 0) = "descuento"
        campos(16, 0) = "cupodirecto"
        campos(17, 0) = "protestos"
        campos(18, 0) = "prorrogas"
        campos(19, 0) = "boletas"
        campos(20, 0) = "facturas"
        campos(21, 0) = "cheques"
        campos(22, 0) = "diapago"
        campos(23, 0) = "cupoutilizadodirecto"
        campos(24, 0) = "comisionvendedor"
        campos(25, 0) = "observaciones"
        campos(26, 0) = "listaprecios"
        campos(27, 0) = "vendedor"
        campos(28, 0) = "tipocliente"
        campos(29, 0) = "bloqueotmp"
        campos(30, 0) = "rebajainterespp"
        campos(31, 0) = "rebajainteresmora"
        campos(32, 0) = "repactacion"
        campos(33, 0) = "cuotasrepactacion"
        campos(34, 0) = "terceraedad"
        campos(35, 0) = "carta"
        campos(36, 0) = "constructora"
        campos(37, 0) = ""
        
        campos(0, 2) = "sv_maestroclientes"

        condicion = "rut " & operador & " '" & codigo1 & "' AND sucursal = '" & codigo2 & "' "
        If operador = "<" Then
            condicion = condicion & "ORDER BY rut DESC"
        Else
            condicion = condicion & "ORDER BY rut ASC"
        End If
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCliente = True
            If sql.response(26, 3) = "3" Then lista4 = "S" Else lista4 = ""
            Call asigna(c, sql)
             diadepago = 0
             montocredito = 0
             
             If sql.response(14, 3) = "T" Then
             diadepago = sql.response(22, 3)
             montocredito = sql.response(16, 3)
             
             End If
             
        Else
            leerCliente = False
        End If
    End Function
'=============================================================================
'LEER CLIENTE
'=============================================================================

'=============================================================================
'GRABAR CLIENTE
'=============================================================================
    Public Sub grabarCliente(ByRef c As Cliente, ByVal modifica As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        Call designa(c, sql)
        condicion = ""
        If modifica = True Then
            condicion = "rut = '" & c.rut & "' AND sucursal = '" & c.sucursal & "'"
            op = 3
        Else
            op = 2
        End If

        sql.audit = True:    sql.programaactivo = MClientes.Caption
        
        
        Set sql.conexion = ventas
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        
        
        
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR CLIENTE
'=============================================================================

'=============================================================================
'ELIMINAR CLIENTE
'=============================================================================
    Public Sub eliminarCliente(ByRef c As Cliente)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & c.rut & "' AND sucursal = '" & c.sucursal & "'"
        op = 4
        campos(0, 2) = "sv_maestroclientes"
        sql.response = campos
        
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = MClientes.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        
        
        Call sql.sqlventas(op, condicion)
        Call eliminarClientePersonales(c.rut, c.sucursal)
        Call eliminarClienteLaborales(c.rut, c.sucursal)
        Call eliminarClienteAdicionales(c.rut, c.sucursal)
        Call eliminarClienteFinanciero(c.rut, c.sucursal)
        
    End Sub
    
'=============================================================================
'ELIMINAR CLIENTE
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================
Private Sub asigna(ByRef c As Cliente, ByRef sql As sqlventas.sqlventa)
        c.rut = sql.response(0, 3)
        c.sucursal = sql.response(1, 3)
        c.nombre = sql.response(2, 3)
        c.direccion = sql.response(3, 3)
        c.comuna = sql.response(4, 3)
        c.ciudad = sql.response(5, 3)
        c.fono1 = sql.response(6, 3)
        c.fono2 = sql.response(7, 3)
        c.fax = sql.response(8, 3)
        c.celular = sql.response(9, 3)
        c.giro = sql.response(10, 3)
        c.email = sql.response(11, 3)
        c.contacto = sql.response(12, 3)
        c.plazo = sql.response(13, 3)
        c.CREDITO = sql.response(14, 3)
        c.Descuento = sql.response(15, 3)
        c.cupodirecto = sql.response(16, 3)
        c.protestos = sql.response(17, 3)
        c.prorrogas = sql.response(18, 3)
        c.boletas = sql.response(19, 3)
        c.facturas = sql.response(20, 3)
        c.Cheques = sql.response(21, 3)
        c.DIAPAGO = sql.response(22, 3)
        c.CUPOUTILIZADO = sql.response(23, 3)
        c.listaprecios = sql.response(26, 3)
        c.observaciones = sql.response(25, 3)
        c.TIPOCLIENTE = sql.response(28, 3)
        c.bloqueoTMP = sql.response(29, 3)
        c.rebajainterespp = sql.response(30, 3)
        c.rebajainteresmora = sql.response(31, 3)
        c.repactacion = sql.response(32, 3)
        c.CUOTASREPACTACION = sql.response(33, 3)
        c.terceraedad = sql.response(34, 3)
        c.CARTA = sql.response(35, 3)
        c.constructora = sql.response(36, 3)
        
        
        
        
        
'        c.vendedor = sql.response(27, 3)
        
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA BASE DE DATOS A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================
    Private Sub designa(ByRef c As Cliente, ByRef sql As sqlventas.sqlventa)
        Dim cad As String
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "nombre"
        campos(3, 0) = "direccion"
        campos(4, 0) = "comuna"
        campos(5, 0) = "ciudad"
        campos(6, 0) = "fono1"
        campos(7, 0) = "fono2"
        campos(8, 0) = "fax"
        campos(9, 0) = "celular"
        campos(10, 0) = "giro"
        campos(11, 0) = "email"
        campos(12, 0) = "contacto"
        campos(13, 0) = "plazo"
        campos(14, 0) = "credito"
        campos(15, 0) = "descuento"
        campos(16, 0) = "cupodirecto"
        campos(17, 0) = "listaprecios"
        campos(18, 0) = "observaciones"
        campos(19, 0) = "diapago"
        campos(20, 0) = "terceraedad"
         If c.CREDITO <> "A" Then
            campos(21, 0) = "estadoanterior"
        Else
            campos(21, 0) = ""
        End If
        campos(22, 0) = "vendedor"
        
        campos(23, 0) = "tipocliente"
        campos(24, 0) = "bloqueotmp"
        campos(25, 0) = "rebajainterespp"
        campos(26, 0) = "rebajainteresmora"
        campos(27, 0) = "repactacion"
        campos(28, 0) = "cuotasrepactacion"
        
        
        campos(29, 0) = ""
        campos(0, 1) = c.rut
        campos(1, 1) = c.sucursal
        campos(2, 1) = c.nombre
        campos(3, 1) = c.direccion
        campos(4, 1) = c.comuna
        campos(5, 1) = c.ciudad
        campos(6, 1) = c.fono1
        campos(7, 1) = c.fono2
        campos(8, 1) = c.fax
        campos(9, 1) = c.celular
        campos(10, 1) = c.giro
        campos(11, 1) = c.email
        campos(12, 1) = c.contacto
        campos(13, 1) = c.plazo
        campos(14, 1) = c.CREDITO
        campos(15, 1) = c.Descuento
        campos(16, 1) = c.cupodirecto
        campos(17, 1) = c.listaprecios
        campos(18, 1) = c.observaciones
        campos(19, 1) = c.DIAPAGO
        campos(20, 1) = c.terceraedad
              
        If c.CREDITO <> "A" Then
            campos(21, 1) = c.CREDITO
        Else
            campos(21, 1) = ""
        End If
        campos(22, 1) = c.vendedor
        campos(23, 1) = c.TIPOCLIENTE
        campos(24, 1) = c.bloqueoTMP
        campos(25, 1) = c.rebajainterespp
        campos(26, 1) = c.rebajainteresmora
        campos(27, 1) = c.repactacion
        campos(28, 1) = c.CUOTASREPACTACION
        campos(0, 2) = "sv_maestroclientes"
        sql.response = campos
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LA BASE DE DATOS
'=============================================================================

    Public Function leerCupoutilizado(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "IFNULL(SUM(total - abono),0)"
        campos(1, 0) = ""
        
        campos(0, 2) = "sv_documento_cabeza_" + empresaActiva
        
        condicion = "rut = '" & rut & "' AND sucursal  = '" & sucursal & "' AND (tipo = 'FV' OR tipo = 'BV' OR tipo = 'ZE' OR tipo = 'FE' OR tipo = 'NV') AND nula = 'N'"
        op = 5
        sql.response = campos
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCupoutilizado = sql.response(0, 3)
        Else
            leerCupoutilizado = "0"
        End If
    End Function

    Public Function leerPesosProtesto(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "IFNULL(SUM(monto),0)"
        campos(1, 0) = "IFNULL(COUNT(monto),0)"
        campos(2, 0) = ""
        
        campos(0, 2) = "sv_protesto_" & empresaActiva
        
        condicion = "rut = '" & rut & "' AND sucursal  = '" & sucursal & "' AND cancelado = fechacheque"
        op = 5
        sql.response = campos
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPesosProtesto = sql.response(0, 3) & "/" & sql.response(1, 3)
        Else
            leerPesosProtesto = "0"
        End If
    End Function
    
    Public Function leerPesosProrroga(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "IFNULL(SUM(monto),0)"
        campos(1, 0) = "IFNULL(COUNT(monto),0)"
        campos(2, 0) = ""
        
        campos(0, 2) = "sv_prorroga_" & empresaActiva
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "' AND fprorroga > '" & fechasistema & "'"
        op = 5
        sql.response = campos
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPesosProrroga = sql.response(0, 3) & "/" & sql.response(1, 3)
        Else
            leerPesosProrroga = "0"
        End If
    End Function
    
    Public Function leerPesosBoleta(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "IFNULL(SUM(total),0)"
        campos(1, 0) = "IFNULL(COUNT(total),0)"
        campos(2, 0) = ""
        
        campos(0, 2) = "sv_documento_cabeza_" + empresaActiva
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "' AND tipo = 'BV' AND nula = 'N' AND total <> abono "
        op = 5
        sql.response = campos
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPesosBoleta = sql.response(0, 3) & "/" & sql.response(1, 3)
        Else
            leerPesosBoleta = "0"
        End If
    End Function
    
    Public Function leerPesosFactura(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "IFNULL(SUM(total),0)"
        campos(1, 0) = "IFNULL(COUNT(total),0)"
        campos(2, 0) = ""
        
        campos(0, 2) = "sv_documento_cabeza_" + empresaActiva
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "' AND tipo = 'FV' AND nula = 'N' AND total <> abono "
        op = 5
        sql.response = campos
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPesosFactura = sql.response(0, 3) & "/" & sql.response(1, 3)
        Else
            leerPesosFactura = "0"
        End If
    End Function
' revisar
    Public Function leerPesosCheque(ByVal rut As String, ByVal sucursal As String) As String
'
'        Dim op As Integer
'        Set sql =new sqlventas.sqlventa
'        campos(0, 0) = "IFNULL(SUM(monto),0)"
'        campos(1, 0) = "IFNULL(COUNT(monto),0)"
'        campos(2, 0) = ""
'
'        campos(0, 2) = "sv_carteracheques"
'
'        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "' AND fechavencimiento >= '" & fechasistema & "'"
'        op = 5
'        sql.response = campos
'        Set sql.conexion = ventasRubro
'        Call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerPesosCheque = sql.response(0, 3) & "/" & sql.response(1, 3)
'        Else
'            leerPesosCheque = "0"
'        End If
    End Function
'revisar
    Public Sub actualizarDatosCliente(ByVal rut As String, ByVal sucursal As String)
        Dim csql As rdoQuery
        Dim CUPOUTILIZADO As String
        Dim protestos As String
        Dim boletas As String
        Dim facturas As String
        Dim Cheques As String
        Dim prorrogas As String
        Dim compras As String
        Dim cadena As String
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        
        
        CUPOUTILIZADO = leerCupoutilizado(rut, sucursal)
        
        cadena = leerPesosProtesto(rut, sucursal)
        protestos = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        
        cadena = leerPesosProrroga(rut, sucursal)
        prorrogas = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        
        cadena = leerPesosBoleta(rut, sucursal)
        boletas = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        
        cadena = leerPesosFactura(rut, sucursal)
        facturas = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        
        cadena = leerPesosCheque(rut, sucursal)
        
'        Cheques = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        Cheques = 0
        compras = Format(CDbl(boletas) + CDbl(facturas), "########0")
        
        csql.sql = "UPDATE sv_maestroclientes "
        csql.sql = csql.sql & "SET cupoutilizadodirecto = " & CUPOUTILIZADO & ", protestos = " & protestos & ", boletas = " & boletas & ", facturas = " & facturas & ", cheques = " & Cheques & ", prorrogas = " & prorrogas & ", compras = " & compras & " "
        csql.sql = csql.sql & "WHERE rut = '" & rut & "' AND sucursal = '" & sucursal & "' "
        csql.Execute
            Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub

Public Function leerClienteadicional(ByRef c As Clientes, ByVal operador As String) As Boolean
       
        If leerClienteCliente(c, operador) = True Then
           
            Call leerClientePersonales(c)
            Call leerClienteLaborales(c)
            Call leerClienteAdicionales(c)
            Call leerClienteFinancieros(c)
        End If
    End Function
 Public Function leerClienteCliente(ByRef c As Clientes, ByVal operador As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "nombre"
        campos(3, 0) = "direccion"
        campos(4, 0) = "comuna"
        campos(5, 0) = "ciudad"
        campos(6, 0) = "fono1"
        campos(7, 0) = "fono2"
        campos(8, 0) = "fax"
        campos(9, 0) = "celular"
        campos(10, 0) = "giro"
        campos(11, 0) = "email"
        campos(12, 0) = "casilla"
        campos(13, 0) = "contacto"
        campos(14, 0) = "diapago"
        campos(15, 0) = "plazo"
        campos(16, 0) = "creditotmp"
        campos(17, 0) = "creditodirecto"
        campos(18, 0) = "descuento"
        campos(19, 0) = "IFNULL(cupodirecto,0)"
        campos(20, 0) = "IFNULL(cupodirecto,0)"
        campos(21, 0) = "zona"
        campos(22, 0) = "IFNULL(cupoutilizadotmp,0)"
        campos(23, 0) = "IFNULL(cupoutilizadodirecto,0)"
        campos(24, 0) = "protestos"
        campos(25, 0) = "boletas"
        campos(26, 0) = "facturas"
        campos(27, 0) = "zetas"
        campos(28, 0) = "cheques"
        campos(29, 0) = "prorrogas"
        campos(30, 0) = "garantias"
        campos(31, 0) = "morosidad"
        campos(32, 0) = "cantprotestos"
        campos(33, 0) = "cantboletas"
        campos(34, 0) = "cantfacturas"
        campos(35, 0) = "cantzetas"
        campos(36, 0) = "cantcheques"
        campos(37, 0) = "cantprorrogas"
        campos(38, 0) = "cantgarantias"
        campos(39, 0) = "cantmorosidad"
        campos(40, 0) = "terceraedad"
        campos(41, 0) = "tipocliente"
        campos(42, 0) = "bloqueotmp"
        campos(43, 0) = "bloqueodirecto"
        
        campos(44, 0) = ""

        campos(0, 2) = "sv_maestroclientes"

        Select Case operador
            Case "<"
                condicion = "rut " & operador & " '" & c.cc.rut & "' ORDER BY rut DESC"
            Case "="
                condicion = "rut " & operador & " '" & c.cc.rut & "' AND sucursal = '" & c.cc.sucursal & "'"
            Case ">"
                condicion = "rut " & operador & " '" & c.cc.rut & "' ORDER BY rut ASC"
            Case Else
                condicion = "1"
        End Select
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerClienteCliente = True
'            Call asignaCliente(c, sql)
        Else
            leerClienteCliente = False
        End If
    End Function
    
    Private Sub leerClientePersonales(ByRef c As Clientes)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "fechanacimiento"
        campos(3, 0) = "sexo"
        campos(4, 0) = "nacionalidad"
        campos(5, 0) = "estadocivil"
        campos(6, 0) = "rutconyuge"
        campos(7, 0) = "nombreconyuge"
        campos(8, 0) = ""

        campos(0, 2) = "sv_maestroclientes_personales"

        condicion = "rut = '" & c.cc.rut & "' AND sucursal = '" & c.cc.sucursal & "'"
        
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        Call asignaPersonales(c, sql)
    End Sub
    
    Private Sub leerClienteLaborales(ByRef c As Clientes)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "labor"
        campos(3, 0) = "rutempleador"
        campos(4, 0) = "nombre"
        campos(5, 0) = "direccion"
        campos(6, 0) = "comuna"
        campos(7, 0) = "ciudad"
        campos(8, 0) = "fono"
        campos(9, 0) = "antiguedad"
        campos(10, 0) = "codeudor"
        campos(11, 0) = ""

        campos(0, 2) = "sv_maestroclientes_laborales"

        condicion = "rut = '" & c.cc.rut & "' AND sucursal = '" & c.cc.sucursal & "'"
        
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        Call asignaLaborales(c, sql)
    End Sub
    
    Public Sub leerClienteAdicionales(ByRef c As Clientes)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "rutadicional"
        campos(3, 0) = "nombre"
        campos(4, 0) = "porcentajecupo"
        campos(5, 0) = ""

        campos(0, 2) = "sv_maestroclientes_adicionales"

        condicion = "rut = '" & c.cc.rut & "' AND sucursal = '" & c.cc.sucursal & "' AND rutadicional = '" & c.ca.rutadicional & "'"
        
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        Call asignaAdicionales(c, sql)
    End Sub
    
    Private Sub leerClienteFinancieros(ByRef c As Clientes)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "IFNULL(ingresomensual,0)"
        campos(3, 0) = "IFNULL(pagoscasascomerciales,0)"
        campos(4, 0) = "IFNULL(arriendo,0)"
        campos(5, 0) = "tipovivienda"
        campos(6, 0) = "IFNULL(tasacionvivienda,0)"
        campos(7, 0) = "vehiculos"
        campos(8, 0) = "IFNULL(tasacionvehiculos,0)"
        campos(9, 0) = "cuentacorriente"
        campos(10, 0) = "banco"
        campos(11, 0) = "numerocuenta"
        campos(12, 0) = "antiguedad"
        campos(13, 0) = "otrastarjetas"
        campos(14, 0) = "otratarjeta1"
        campos(15, 0) = "otratarjeta2"
        campos(16, 0) = "otratarjeta3"
        campos(17, 0) = "IFNULL(otratarjetacupo1,0)"
        campos(18, 0) = "IFNULL(otratarjetacupo2,0)"
        campos(19, 0) = "IFNULL(otratarjetacupo3,0)"
        campos(20, 0) = "IFNULL(fechaimpresionpagare,'')"
        campos(21, 0) = "IFNULL(fechaautorizacioncredito,'')"
        campos(22, 0) = "autorizador"
        campos(23, 0) = "IFNULL(fechaentregatarjeta,'')"
        campos(24, 0) = ""

        campos(0, 2) = "sv_maestroclientes_financiero"

        condicion = "rut = '" & c.cc.rut & "' AND sucursal = '" & c.cc.sucursal & "'"
        
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        Call asignaFinancieros(c, sql)
    End Sub
    
     Private Sub asignaAdicionales(ByRef c As Clientes, ByRef sql As sqlventas.sqlventa)
        If sql.Status = 0 Then
            c.ca.rut = sql.response(0, 3)
            c.ca.sucursal = sql.response(1, 3)
            c.ca.rutadicional = sql.response(2, 3)
            c.ca.nombre = sql.response(3, 3)
            c.ca.porcentajecupo = sql.response(4, 3)
        Else
            c.ca.rut = ""
            c.ca.sucursal = ""
            c.ca.rutadicional = ""
            c.ca.nombre = ""
            c.ca.porcentajecupo = ""
        End If
    End Sub
    
    Private Sub asignaFinancieros(ByRef c As Clientes, ByRef sql As sqlventas.sqlventa)
        If sql.Status = 0 Then
            c.cf.rut = sql.response(0, 3)
            c.cf.sucursal = sql.response(1, 3)
            c.cf.ingresomensual = sql.response(2, 3)
            c.cf.pagoscasascomerciales = sql.response(3, 3)
            c.cf.arriendo = sql.response(4, 3)
            c.cf.tipovivienda = sql.response(5, 3)
            c.cf.tasacionvivienda = sql.response(6, 3)
            c.cf.vehiculos = sql.response(7, 3)
            c.cf.tasacionvehiculos = sql.response(8, 3)
            c.cf.cuentacorriente = sql.response(9, 3)
            c.cf.Banco = sql.response(10, 3)
            c.cf.numerocuenta = sql.response(11, 3)
            c.cf.antuguedad = sql.response(12, 3)
            c.cf.otrastarjetas = sql.response(13, 3)
            c.cf.otratarjeta1 = sql.response(14, 3)
            c.cf.otratarjeta2 = sql.response(15, 3)
            c.cf.otratarjeta3 = sql.response(16, 3)
            c.cf.otratarjetacupo1 = sql.response(17, 3)
            c.cf.otratarjetacupo2 = sql.response(18, 3)
            c.cf.otratarjetacupo3 = sql.response(19, 3)
            c.cf.fechaimpresionpagare = sql.response(20, 3)
            c.cf.fechaautorizacioncredito = sql.response(21, 3)
            c.cf.autorizador = sql.response(22, 3)
            c.cf.fechaentregatarjeta = sql.response(23, 3)
        Else
            c.cf.rut = ""
            c.cf.sucursal = ""
            c.cf.ingresomensual = "0"
            c.cf.pagoscasascomerciales = "0"
            c.cf.arriendo = "0"
            c.cf.tipovivienda = ""
            c.cf.tasacionvivienda = "0"
            c.cf.vehiculos = ""
            c.cf.tasacionvehiculos = "0"
            c.cf.cuentacorriente = ""
            c.cf.Banco = ""
            c.cf.numerocuenta = ""
            c.cf.antuguedad = ""
            c.cf.otrastarjetas = ""
            c.cf.otratarjeta1 = ""
            c.cf.otratarjeta2 = ""
            c.cf.otratarjeta3 = ""
            c.cf.otratarjetacupo1 = "0"
            c.cf.otratarjetacupo2 = "0"
            c.cf.otratarjetacupo3 = "0"
            c.cf.fechaimpresionpagare = ""
            c.cf.fechaautorizacioncredito = ""
            c.cf.autorizador = ""
            c.cf.fechaentregatarjeta = ""
        End If
    End Sub
      
    Private Sub asignaPersonales(ByRef c As Clientes, ByRef sql As sqlventas.sqlventa)
        If sql.Status = 0 Then
            c.cp.rut = sql.response(0, 3)
            c.cp.sucursal = sql.response(1, 3)
            c.cp.fechanacimiento = sql.response(2, 3)
            c.cp.sexo = sql.response(3, 3)
            c.cp.nacionalidad = sql.response(4, 3)
            c.cp.estadocivil = sql.response(5, 3)
            c.cp.rutconyuge = sql.response(6, 3)
            c.cp.nombreconyuge = sql.response(7, 3)
        Else
            c.cp.rut = ""
            c.cp.sucursal = ""
            c.cp.fechanacimiento = ""
            c.cp.sexo = ""
            c.cp.nacionalidad = ""
            c.cp.estadocivil = ""
            c.cp.rutconyuge = ""
            c.cp.nombreconyuge = ""
        End If
    End Sub
  Private Sub asignaLaborales(ByRef c As Clientes, ByRef sql As sqlventas.sqlventa)
        If sql.Status = 0 Then
            c.cl.rut = sql.response(0, 3)
            c.cl.sucursal = sql.response(1, 3)
            c.cl.labor = sql.response(2, 3)
            c.cl.rutempleador = sql.response(3, 3)
            c.cl.nombre = sql.response(4, 3)
            c.cl.direccion = sql.response(5, 3)
            c.cl.comuna = sql.response(6, 3)
            c.cl.ciudad = sql.response(7, 3)
            c.cl.fono = sql.response(8, 3)
            c.cl.antiguedad = sql.response(9, 3)
            c.cl.codeudor = sql.response(10, 3)
        Else
            c.cl.rut = ""
            c.cl.sucursal = ""
            c.cl.labor = ""
            c.cl.rutempleador = ""
            c.cl.nombre = ""
            c.cl.direccion = ""
            c.cl.comuna = ""
            c.cl.ciudad = ""
            c.cl.fono = ""
            c.cl.antiguedad = ""
            c.cl.codeudor = "0"
        End If
    End Sub
    
    Private Sub designaPersonales(ByRef cp As personales, ByRef sql As sqlventas.sqlventa)
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "fechanacimiento"
        campos(3, 0) = "sexo"
        campos(4, 0) = "nacionalidad"
        campos(5, 0) = "estadocivil"
        campos(6, 0) = "rutconyuge"
        campos(7, 0) = "nombreconyuge"
        campos(8, 0) = ""
        
        campos(0, 1) = cp.rut
        campos(1, 1) = cp.sucursal
        campos(2, 1) = cp.fechanacimiento
        campos(3, 1) = cp.sexo
        campos(4, 1) = cp.nacionalidad
        campos(5, 1) = cp.estadocivil
        campos(6, 1) = cp.rutconyuge
        campos(7, 1) = cp.nombreconyuge
        campos(8, 1) = ""
        
        campos(0, 2) = "sv_maestroclientes_personales"
        sql.response = campos
    End Sub
    
    Private Sub designaLaborales(ByRef cl As laborales, ByRef sql As sqlventas.sqlventa)
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "labor"
        campos(3, 0) = "rutempleador"
        campos(4, 0) = "nombre"
        campos(5, 0) = "direccion"
        campos(6, 0) = "comuna"
        campos(7, 0) = "ciudad"
        campos(8, 0) = "fono"
        campos(9, 0) = "antiguedad"
        campos(10, 0) = "codeudor"
        campos(11, 0) = ""
        
        campos(0, 1) = cl.rut
        campos(1, 1) = cl.sucursal
        campos(2, 1) = cl.labor
        campos(3, 1) = cl.rutempleador
        campos(4, 1) = cl.nombre
        campos(5, 1) = cl.direccion
        campos(6, 1) = cl.comuna
        campos(7, 1) = cl.ciudad
        campos(8, 1) = cl.fono
        campos(9, 1) = cl.antiguedad
        campos(10, 1) = cl.codeudor
        campos(11, 1) = ""
        
        campos(0, 2) = "sv_maestroclientes_laborales"
        sql.response = campos
    End Sub
    
    Private Sub designaAdicionales(ByRef ca As adicionales, ByRef sql As sqlventas.sqlventa)
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "rutadicional"
        campos(3, 0) = "nombre"
        campos(4, 0) = "porcentajecupo"
        campos(5, 0) = ""
        
        campos(0, 1) = ca.rut
        campos(1, 1) = ca.sucursal
        campos(2, 1) = ca.rutadicional
        campos(3, 1) = ca.nombre
        campos(4, 1) = ca.porcentajecupo
        campos(5, 1) = ""
        
        campos(0, 2) = "sv_maestroclientes_adicionales"
        sql.response = campos
    End Sub
    
    Private Sub designaFinancieros(ByRef cf As financieros, ByRef sql As sqlventas.sqlventa)
        campos(0, 0) = "rut"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "ingresomensual"
        campos(3, 0) = "pagoscasascomerciales"
        campos(4, 0) = "arriendo"
        campos(5, 0) = "tipovivienda"
        campos(6, 0) = "tasacionvivienda"
        campos(7, 0) = "vehiculos"
        campos(8, 0) = "tasacionvehiculos"
        campos(9, 0) = "cuentacorriente"
        campos(10, 0) = "banco"
        campos(11, 0) = "numerocuenta"
        campos(12, 0) = "antiguedad"
        campos(13, 0) = "otrastarjetas"
        campos(14, 0) = "otratarjeta1"
        campos(15, 0) = "otratarjeta2"
        campos(16, 0) = "otratarjeta3"
        campos(17, 0) = "otratarjetacupo1"
        campos(18, 0) = "otratarjetacupo2"
        campos(19, 0) = "otratarjetacupo3"
        campos(20, 0) = ""
        
        campos(0, 1) = cf.rut
        campos(1, 1) = cf.sucursal
        campos(2, 1) = cf.ingresomensual
        campos(3, 1) = cf.pagoscasascomerciales
        campos(4, 1) = cf.arriendo
        campos(5, 1) = cf.tipovivienda
        campos(6, 1) = cf.tasacionvivienda
        campos(7, 1) = cf.vehiculos
        campos(8, 1) = cf.tasacionvehiculos
        campos(9, 1) = cf.cuentacorriente
        campos(10, 1) = cf.Banco
        campos(11, 1) = cf.numerocuenta
        campos(12, 1) = cf.antuguedad
        campos(13, 1) = cf.otrastarjetas
        campos(14, 1) = cf.otratarjeta1
        campos(15, 1) = cf.otratarjeta2
        campos(16, 1) = cf.otratarjeta3
        campos(17, 1) = cf.otratarjetacupo1
        campos(18, 1) = cf.otratarjetacupo2
        campos(19, 1) = cf.otratarjetacupo3
        campos(20, 1) = ""
        
        campos(0, 2) = "sv_maestroclientes_financiero"
        sql.response = campos
    End Sub
    
    Public Sub grabarClientePersonales(ByRef cp As personales, ByVal modificar As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        If modificar = True Then
            condicion = "rut = '" & cp.rut & "' AND sucursal = '" & cp.sucursal & "'"
            op = 3
        Else
            condicion = ""
            op = 2
        End If
        Set sql.conexion = ventas
        Call designaPersonales(cp, sql)
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Public Sub grabarClienteLaborales(ByRef cl As laborales, ByVal modificar As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        If modificar = True Then
            condicion = "rut = '" & cl.rut & "' AND sucursal = '" & cl.sucursal & "'"
            op = 3
        Else
            condicion = ""
            op = 2
        End If
        Set sql.conexion = ventas
        Call designaLaborales(cl, sql)
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Public Sub grabarClienteAdicionales(ByRef ca As adicionales, ByVal modificar As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        If modificar = True Then
            condicion = "rut = '" & ca.rut & "' AND sucursal = '" & ca.sucursal & "' AND rutadicional = '" & ca.rutadicional & "'"
            op = 3
        Else
            condicion = ""
            op = 2
        End If
        Set sql.conexion = ventas
        Call designaAdicionales(ca, sql)
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Public Sub grabarClienteFinancieros(ByRef cf As financieros, ByVal modificar As Boolean)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        If modificar = True Then
            condicion = "rut = '" & cf.rut & "' AND sucursal = '" & cf.sucursal & "'"
            op = 3
        Else
            condicion = ""
            op = 2
        End If
        Set sql.conexion = ventas
        Call designaFinancieros(cf, sql)
        Call sql.sqlventas(op, condicion)
    End Sub
    
  Private Sub eliminarClientePersonales(ByVal rut As String, ByVal sucursal As String)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 4
        Set sql.conexion = ventas
        campos(0, 2) = "sv_maestroclientes_personales"
        sql.response = campos
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Private Sub eliminarClienteLaborales(ByVal rut As String, ByVal sucursal As String)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 4
        Set sql.conexion = ventas
        campos(0, 2) = "sv_maestroclientes_laborales"
        sql.response = campos
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Private Sub eliminarClienteAdicionales(ByVal rut As String, ByVal sucursal As String)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 4
        Set sql.conexion = ventas
        campos(0, 2) = "sv_maestroclientes_adicionales"
        sql.response = campos
        Call sql.sqlventas(op, condicion)
    End Sub
    
    Private Sub eliminarClienteFinanciero(ByVal rut As String, ByVal sucursal As String)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 4
        Set sql.conexion = ventas
        campos(0, 2) = "sv_maestroclientes_financiero"
        sql.response = campos
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR CLIENTE
'=============================================================================

