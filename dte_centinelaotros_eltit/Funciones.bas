Attribute VB_Name = "Funciones"
Option Explicit
    Public CAJA99 As String
    Public CAJA98 As String
    Public desconectado As Boolean
    Public fechaultimacompra1 As Date
    Public fechaultimacompra2 As Date
    Public cedible As Boolean
    Public MONTOMAXIMOCHEQUE As Double
    Public clientesistema As String
    Public FECHAMORA As String
    Public ASEGURADORA As String
    Public toleranciacredito As Double
    Public impresoracredito As String
    Public AUTORIZASISTEMA As Boolean
    Public filaprecio As Integer
    Public diasgracia As Double
    Public IMPRESORAPAGO As String
    Public repactacionimprimir As String
    Public NUEVONUMEROGUIA As String
    Public diascierre As Double
    Public BODEGARETIRO As String
    Public PUERTOCREDITO As String
    Public codigoempresa As String
    Public codigoCONTABLE As String
    Public NOMBREEMPRESA As String
    Public DIRECCIONEMPRESA As String
    Public glosaeliminacionsistema As String
    Public solicitaeliminacion As String
    Public entidaddonacion As String
    Public COMUNAEMPRESA As String
    Public CIUDADEMPRESA As String
    Public rutempresa As String
    Public GIROEMPRESA As String
    Public DATOSEMPRESA(20) As String
    Public glosa1flete As String
    Public glosa2flete As String
    Public nombrecajero As String
    Public codigocajero As String
    Public cantidadcuotas As String
    Public rutcredito As String
    Public primervencimiento As String
    Public montocredito2 As String
    Public montocuotas As String
    Public ultimaFila As Long
    Public altoFila As Long
    Public resX As Integer
    Public resY As Integer
    Private CAMPOS(10, 10) As String
    Public cabezaLocales() As String
    Public cantLocales As Integer
    Public fila1 As Long
    Public fila2 As Long
    Public col1 As Long
    Public col2 As Long
    Public vend As String
    Public tipoGrafico As Integer
    Public primerAño As String
    Public segundoAño As String
    Public meses(12) As String
    Public cantMeses As Integer
    Public ProcesaNC As Boolean
    Public Declare Sub Sleep Lib "kernel32" (ByVal Milisegundos As Long)
    Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Type POINTAPI
        x As Long
        Y As Long
    End Type
    Public mouse As POINTAPI
    Public Const pass = "abcdefghijklmnñopqrstuvwxyz0123456789"
    Public Const passUsuario = "semilla"

Public Sub Centrar(ByRef frm As Form)
    frm.Left = (Screen.Width - frm.Width) / 2 - 25
    frm.Top = (Screen.Height - frm.Height) / 2 - 700
End Sub

Public Function txtBarra(ByVal cadena As String) As String
    Dim I As Integer
    I = InStr(1, cadena, "&", vbBinaryCompare)
    txtBarra = Left(cadena, I - 1)
    txtBarra = UCase(txtBarra & Right(cadena, Len(cadena) - I))
End Function

Public Sub limpiaBarra(ByVal I As Integer)

End Sub

Public Function esNumero(ByVal NUM As Long) As Long
    If NUM <> vbKeyBack Then
        If NUM <> vbKeyReturn Then
            'If num <> vbKeyTab Then
                If NUM < 48 Or NUM > 57 Then
                    esNumero = 0
                Else
                    esNumero = NUM
                End If
            'Else
            '    esNumero = 13
            'End If
        Else
            esNumero = NUM
        End If
    Else
        esNumero = NUM
    End If
End Function

Public Function esNumeroDecimal(ByRef txt As TextBox, ByVal NUM As Long) As Long
    Dim numdec As Long
    numdec = NUM
    NUM = esNumero(NUM)
    If NUM = 0 Then
        If numdec = 46 Then '.
            If InStr(1, txt.text, ",", vbBinaryCompare) <> 0 Then
                esNumeroDecimal = 0
            Else
                esNumeroDecimal = 44    ',
            End If
        Else
            esNumeroDecimal = 0
        End If
    Else
        esNumeroDecimal = numdec
    End If
End Function

Public Sub LimpiarCajas(ByRef frm As Form)
    Dim ctlControl As Object
    On Error Resume Next
    For Each ctlControl In frm.Controls
        ctlControl.text = ""
        DoEvents
    Next ctlControl
End Sub

Public Sub LimpiarLabels(ByRef frm As Form)
    Dim ctlControl As Object
    Dim cad As String
    On Error Resume Next
    For Each ctlControl In frm.Controls
        If InStr(1, ctlControl.Name, "frm", vbBinaryCompare) = 0 Then
            cad = Mid(ctlControl.Name, 4, Len(ctlControl.Name))
            If IsNumeric(cad) = False Then
                ctlControl.Caption = ""
            End If
        End If
        DoEvents
    Next ctlControl
End Sub

Public Sub DeshabilitarCajas(ByRef frm As Form)
    Dim ctlControl As Object
    On Error Resume Next
    For Each ctlControl In frm.Controls
        If InStr(1, ctlControl.Name, "dato", vbBinaryCompare) <> 0 Then
            
            ctlControl.Enabled = False
        End If
        DoEvents
    Next ctlControl
    If frm.Name = "MClientes" Then
        frm.dato2.Enabled = True
    End If
    If frm.Name = "CProtestados" Then
        frm.dato2.Enabled = True
        frm.dato3.Enabled = True
    End If
    If frm.Name = "PVentas" Then
        frm.dato2.Enabled = True
    End If
    frm.dato1.Enabled = True
End Sub

Public Sub HabilitarCajas(ByRef frm As Form, ByVal modifica As Boolean)
    Dim ctlControl As Object
    On Error Resume Next
    For Each ctlControl In frm.Controls
        If InStr(1, ctlControl.Name, "dato", vbBinaryCompare) <> 0 Then
            ctlControl.Enabled = True
        End If
       DoEvents
    Next ctlControl
    If modifica = True Then
        If frm.Name = "MClientes" Then
            frm.dato2.Enabled = False
        End If
        If frm.Name = "CProtestados" Then
            frm.dato2.Enabled = False
        End If
        If frm.Name = "PVentas" Then
        frm.dato2.Enabled = False
    End If
        frm.dato1.Enabled = False
    End If
End Sub

Public Sub Flechas(ByVal codigo As Integer, ByRef anterior As Object)
    If codigo = 38 Then
        If anterior.Enabled = True Then
            anterior.SetFocus
        End If
    End If
    If codigo = 40 Then
        SendKeys "{Tab}"
    End If
End Sub

Public Function rut(ByVal numrut As String) As String
    Dim guia
    Dim mataux(9) As Integer
    Dim I, suma As Integer
    guia = Array("4", "3", "2", "7", "6", "5", "4", "3", "2")
    suma = 0
    For I = 0 To 8
        mataux(I) = Val(guia(I)) * Val(Mid(numrut, I + 1, 1))
        suma = suma + mataux(I)
    Next
    rut = 11 - suma Mod 11
    Select Case rut
        Case "11"
            rut = "0"
        Case "10"
            rut = "K"
    End Select
End Function

Public Function ceros(ByRef txt As TextBox) As String
    ceros = String(txt.MaxLength - Len(txt.text), "0") & txt.text
End Function

Public Sub selecciona(ByRef txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt.text)
End Sub

Public Sub VerificarCajas(ByRef frm As Form, ByRef txt As Object)
    Dim ctlControl As Object
    On Error Resume Next
    For Each ctlControl In frm.Controls
        If InStr(1, ctlControl.Name, "dato", vbBinaryCompare) <> 0 Then
            If ctlControl.text = "" Then
                If Val(Mid(ctlControl.Name, 5, Len(ctlControl.Name))) < Val(Mid(txt.Name, 5, Len(txt.Name))) Then
                    ctlControl.SetFocus
                    Exit For
                Else
                    If InStr(1, txt.Name, "dato", vbBinaryCompare) = 0 Then
                        ctlControl.SetFocus
                        Exit For
                    End If
                End If
            End If
        End If
        DoEvents
    Next ctlControl
End Sub

'=============================================================================
'CARGAR INFORME
'=============================================================================
    Public Sub cargaInforme(ByRef data As Adodc, ByRef lista As Grid)
        lista.Rows = 1
        lista.AutoRedraw = False
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                lista.AddItem data.Recordset.Fields("item"), True
                data.Recordset.MoveNext
            Wend
        End If
        lista.AutoRedraw = True
        lista.Refresh
    End Sub
'=============================================================================
'CARGAR INFORME
'=============================================================================

'=============================================================================
'IMPRIMIR
'=============================================================================
    Public Sub imprimir(ByRef lista As Grid, ByVal titulo As String, ByVal orientacion As Integer)
        Dim I As Integer
        Dim cabeza As FlexCell.ReportTitle
        Set cabeza = New FlexCell.ReportTitle
        cabeza.Align = cellCenter
        cabeza.PrintOnAllPages = True
        cabeza.text = titulo
        
        For I = lista.ReportTitles.Count To 0 Step -1
            lista.ReportTitles.Remove I
        Next I
        lista.ReportTitles.Add cabeza
        lista.ReportTitles.AddBlankReportTitle
        
        lista.PageSetup.Orientation = orientacion
        lista.PageSetup.PrintFixedRow = True
        lista.PageSetup.BlackAndWhite = True
        
        lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeTop) = cellThin
        lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeBottom) = cellThin
        lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeLeft) = cellThin
        lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeRight) = cellThin
        lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellInsideVertical) = cellThin
        
        lista.PrintPreview
    End Sub
'=============================================================================
'IMPRIMIR
'=============================================================================

'=============================================================================
'LEER VENCIMIENTO DEL PAGO
'=============================================================================
    Public Function leerVencimientoPago(ByVal numero As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "IFNULL(DATE_FORMAT(fechavencimiento,'%d-%m-%Y'),'')"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_carteracheques"
        
        condicion = "local = '" & localAuditoria & "' AND tipodocumento = 'PA' AND numero = '" & numero & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerVencimientoPago = sql.response(0, 3)
        Else
            leerVencimientoPago = Format(fechasistema, "dd-mm-yyyy")
        End If
    End Function
'=============================================================================
'LEER VENCIMIENTO DEL PAGO
'=============================================================================

'=============================================================================
'LEER FECHA DEL DEPOSITO
'=============================================================================
    Public Function leerDepositoPago(ByVal numero As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "IFNULL(DATE_FORMAT(fechadeposito,'%d-%m-%Y'),'')"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_pagos_cabeza_" & empresaActiva
        
        condicion = "local = '" & localAuditoria & "' AND numero = '" & numero & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerDepositoPago = sql.response(0, 3)
        Else
            leerDepositoPago = ""
        End If
    End Function
'=============================================================================
'LEER FECHA DEL DEPOSITO
'=============================================================================

'=============================================================================
'LEER CANTIDAD Y MONTO DE LOS CHEQUES
'=============================================================================
    Public Function LEERCHEQUES(ByVal rut As String, ByVal fecha As String, ByVal tabulador As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "IFNULL(COUNT(numerocheque),'0')"
        CAMPOS(1, 0) = "IFNULL(SUM(monto),'0')"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_carteracheques"
        
        condicion = "local = '" & localAuditoria & "' AND fechavencimiento >= '" & fecha & "' AND rut = '" & rut & "' GROUP BY rut"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasAuditoria
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            LEERCHEQUES = sql.response(0, 3) & tabulador & Format(sql.response(1, 3), "$ ###,###,##0")
        Else
            LEERCHEQUES = "0" & tabulador & "$ 0"
        End If
    End Function
'=============================================================================
'LEER CANTIDAD Y MONTO DE LOS CHEQUES
'=============================================================================

'=============================================================================
'LEER CANTIDAD Y MONTO DE LAS FACTURAS
'=============================================================================
    Public Function leerFacturas(ByVal rut As String, ByVal tabulador As String, ByVal data As Adodc) As String
        Dim tabla As String
        Dim contador As Integer
        Dim monto As Double
        Dim saldo As Double
        
        contador = 0
        monto = 0
        tabla = "SELECT numero, monto - abono as saldo FROM sv_documentos_cobranza_" & empresaActiva & " "
        tabla = tabla & "WHERE local = '" & localAuditoria & "' AND tipo = 'FV' AND rut = '" & rut & "'"
        Call ConectarControlData(data, Servidor, baseVentas & rubroAuditoria, usuario, password, tabla)
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                saldo = CDbl(data.Recordset.Fields("saldo"))
                If saldo + leerNotaCreditoFactura(data.Recordset.Fields("numero")) > 0 Then
                    contador = contador + 1
                    monto = monto + saldo
                End If
                data.Recordset.MoveNext
            Wend
        End If
        leerFacturas = Format(contador, "########0") & tabulador & Format(monto, "$ ###,###,##0")
    End Function
'=============================================================================
'LEER CANTIDAD Y MONTO DE LAS FACTURAS
'=============================================================================

'=============================================================================
'LEER NOMBRE DEPARTAMENTO
'=============================================================================
    Public Function leerNombreDepto(ByVal codigoSeccion As String, ByVal codigoDepto As String, ByVal rubAux As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = basedatos & rubAux & ".r_maestrodepartamentos_" & rubAux
        
        condicion = "codigodepto = '" & codigoDepto & "' "
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreDepto = sql.response(0, 3)
        Else
            leerNombreDepto = ""
        End If
    End Function
'=============================================================================
'LEER NOMBRE DEPARTAMENTO
'=============================================================================

'=============================================================================
'LEER NOMBRE LINEA
'=============================================================================
    Public Function leerNombreLinea(ByVal codigoSeccion As String, ByVal codigoDepto As String, ByVal codigoLinea As String, ByVal rubAux As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = basedatos & rubAux & ".r_maestrolineas_" & rubAux
        
        condicion = "codigodepto = '" & codigoDepto & "' AND codigolinea = '" & codigoLinea & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreLinea = sql.response(0, 3)
        Else
            leerNombreLinea = ""
        End If
    End Function
'=============================================================================
'LEER NOMBRE LINEA
'=============================================================================

'=============================================================================
'CUENTA LOCALES
'=============================================================================
    Public Function leerCantidadLocales() As Integer
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "MAX(codigo)"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_maestroempresas"
        
        condicion = "1"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCantidadLocales = Val(sql.response(0, 3))
        Else
            leerCantidadLocales = 0
        End If
    End Function
    
'=============================================================================
'CUENTA LOCALES
'=============================================================================

'=============================================================================
'LEER NOMBRE CLIENTE
'=============================================================================
    Public Function leerNombreCliente(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreCliente = sql.response(0, 3)
        Else
            leerNombreCliente = ""
        End If
    End Function
    
'=============================================================================
'LEER NOMBRE CLIENTE
'=============================================================================

'=============================================================================
'LEER DIRECCION CLIENTE
'=============================================================================
    Public Function leerDireccionCliente(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "direccion"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerDireccionCliente = sql.response(0, 3)
        Else
            leerDireccionCliente = ""
        End If
    End Function
     Public Function leerFonoCliente(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "fono1"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerFonoCliente = sql.response(0, 3)
        Else
            leerFonoCliente = ""
        End If
    End Function
    
    
     Public Function leerGiroCliente(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "giro"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerGiroCliente = sql.response(0, 3)
        Else
            leerGiroCliente = ""
        End If
    End Function
    
    
'=============================================================================
'LEER DIRECCION CLIENTE
'=============================================================================

'=============================================================================
'LEER COMUNA CLIENTE
'=============================================================================
    Public Function leerComunaCliente(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "comuna"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerComunaCliente = sql.response(0, 3)
        Else
            leerComunaCliente = ""
        End If
    End Function
'=============================================================================
'LEER COMUNA CLIENTE
'=============================================================================

'=============================================================================
'LEER CIUDAD CLIENTE
'=============================================================================
    Public Function leerCiudadCliente(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "ciudad"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCiudadCliente = "     " & sql.response(0, 3)
        Else
            leerCiudadCliente = ""
        End If
    End Function
'=============================================================================
'LEER CIUDAD CLIENTE
'=============================================================================

'=============================================================================
'LEER CUPO CLIENTE
'=============================================================================
    Public Function leerCupoCliente(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "cupodirecto"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & codigo & "' "
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCupoCliente = sql.response(0, 3) * (1 + toleranciacredito / 100)
            ' leerCupoCliente = sql.response(0, 3)
        Else
            leerCupoCliente = "0"
        End If
    End Function
    
    Public Function leerCupoClienteSucursal(ByVal rut As String, ByVal sucursal As String) As String
        
        Dim op As Integer
        Dim bloquear As Boolean
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "cupo - cupoutilizadodirecto"
        CAMPOS(1, 0) = "credito"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
       
        If sql.Status = 0 Then
            If sql.response(1, 3) = "" Then sql.response(1, 3) = "N"
            Select Case sql.response(1, 3)
                Case "S"
                    If sql.response(0, 3) < 0 Then
                        bloquear = True
                    Else
                        bloquear = False
                    End If
                    leerCupoClienteSucursal = sql.response(0, 3)
                    estadoAnterior = False
                Case "A"
                    bloquear = False
                    leerCupoClienteSucursal = "1"
                    estadoAnterior = True
                Case "F"
                    bloquear = False
                    If sql.response(1, 3) < 0 Then
                        leerCupoClienteSucursal = "1"
                    Else
                        leerCupoClienteSucursal = sql.response(0, 3)
                    End If
                    estadoAnterior = False
                Case "N", "B"
                    bloquear = False
                    leerCupoClienteSucursal = "0"
                    estadoAnterior = False
            End Select
        Else
            bloquear = False
            leerCupoClienteSucursal = "0"
        End If
        If bloquear = True Then
            Call bloquearCliente(rut, sucursal)
        End If
    End Function
'=============================================================================
'LEER CUPO CLIENTE
'=============================================================================

'=============================================================================
'LEER NOMBRE CLIENTE SUCURSAL
'=============================================================================
    Public Function leerNombreClienteSucursal(ByVal codigo1 As String, ByVal codigo2 As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & codigo1 & "' AND sucursal = '" & codigo2 & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreClienteSucursal = sql.response(0, 3)
        Else
            leerNombreClienteSucursal = ""
        End If
    End Function
'=============================================================================
'LEER NOMBRE CLIENTE SUCURSAL
'=============================================================================

'=============================================================================
'LEER COMISION CLIENTE VENDEDOR
'=============================================================================
    Public Function leerComisionCliente(ByVal rut As String, ByVal sucursal As String) As Double
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "comisionvendedor"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerComisionCliente = CDbl(sql.response(0, 3))
        Else
            leerComisionCliente = 0
        End If
    End Function
'=============================================================================
'LEER COMISION CLIENTE VENDEDOR
'=============================================================================

'=============================================================================
'LEER CHEQUE
'=============================================================================
    Public Function leerCheque(ByVal rut As String, ByVal sucursal As String, ByVal numerocheque As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "CONCAT(fechavencimiento, '*', monto)"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_carteracheques"
        
        condicion = "rut = '" & rut & "' AND sucursal = '" & sucursal & "' AND numerocheque = '" & numerocheque & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCheque = sql.response(0, 3)
        Else
            leerCheque = "0000-00-00*0"
        End If
    End Function
'=============================================================================
'LEER CHEQUE
'=============================================================================

'=============================================================================
'LEER NOMBRE EMPLEADO
'=============================================================================
    Public Function leerNombreVendedor(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrovendedores"
        
        condicion = "rut = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreVendedor = sql.response(0, 3)
        Else
            leerNombreVendedor = "SIN VENDEDOR"
        End If
    End Function
'=============================================================================
'LEER NOMBRE EMPLEADO
'=============================================================================

'=============================================================================
'LEER NOMBRE BANCO
'=============================================================================
    Public Function leerNombreBanco(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrobancos"
        
        condicion = "codigobanco = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreBanco = sql.response(0, 3)
        Else
            leerNombreBanco = ""
        End If
    End Function
'=============================================================================
'LEER NOMBRE BANCO
'=============================================================================

'=============================================================================
'LEER NOMBRE PRODUCTO
'=============================================================================
    
    Public Function leerNombreProducto(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "descripcion"
        CAMPOS(1, 0) = "preciolibre"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_fijo_" & rubro
        
        condicion = "codigobarra = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreProducto = sql.response(0, 3)
            If sql.response(1, 3) = 1 Then
            PRECIOLIBRE = True
            Else
            PRECIOLIBRE = False
            End If
            
        Else
            leerNombreProducto = ""
        
        End If
    End Function
'=============================================================================
'LEER NOMBRE PRODUCTO
'=============================================================================

'=============================================================================
'LEER COSTO PRODUCTO
'=============================================================================
    Public Function leerCostoProducto(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "pcosto"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_fijo_" & rubro
        
        condicion = "codigobarra = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCostoProducto = sql.response(0, 3)
        Else
            leerCostoProducto = "0"
        End If
    End Function
'=============================================================================
'LEER COSTO PRODUCTO
'=============================================================================

'=============================================================================
'LEER PRECIO PRODUCTO
'=============================================================================
    Public Function leerPrecioProducto(ByVal codigo As String, ByVal tipoprecio As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
      
        CAMPOS(0, 0) = "preciopuntoventa"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_precios_" & rubro
        
        condicion = "codigo = '" & codigo & "' AND codigoprecio = '" & tipoprecio & "' and local='" + empresaActiva + "' "
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        
        If sql.Status = 0 And sql.response(0, 3) <> "0" Then
            leerPrecioProducto = sql.response(0, 3)
           
            
            
            
        Else
            leerPrecioProducto = "0"
        End If
    End Function
'=============================================================================
'LEER PRECIO PRODUCTO
'=============================================================================

'=============================================================================
'LEER PRECIO PRODUCTO
'=============================================================================
    Public Function leerUnidadesProducto(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "cantidadporembalaje"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_fijo_" & rubro
        
        condicion = "codigobarra = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerUnidadesProducto = sql.response(0, 3)
        Else
            leerUnidadesProducto = "0"
        End If
    End Function
'=============================================================================
'LEER PRECIO PRODUCTO
'=============================================================================

'=============================================================================
'LEER NOMBRE DOCUMENTO
'=============================================================================
    Public Function leerNombreDocumento(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombredocumento"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_maestrotipodedocumentos"
        
        condicion = "tipos = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreDocumento = sql.response(0, 3)
        Else
            leerNombreDocumento = ""
        End If
    End Function
'=============================================================================
'LEER NOMBRE DOCUMENTO
'=============================================================================

'=============================================================================
'LEER TIPO EGRESO
'=============================================================================
    Public Function leerTipoEgreso(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrotipoegresoscaja"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerTipoEgreso = sql.response(0, 3)
        Else
            leerTipoEgreso = ""
        End If
    End Function
'=============================================================================
'LEER TIPO EGRESO
'=============================================================================

'=============================================================================
'LEER FORMA PAGO
'=============================================================================
    Public Function leerFormaPago(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_tiposdepagoclientes"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerFormaPago = sql.response(0, 3)
        Else
            leerFormaPago = ""
        End If
    End Function
 Public Function leerFormaPago2(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrotipopago_pago"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerFormaPago2 = sql.response(0, 3)
        Else
            leerFormaPago2 = ""
        End If
    End Function
'==
'=============================================================================
'LEER FORMA PAGO
'=============================================================================

'=============================================================================
'LEER NOMBRE EMPRESA
'=============================================================================
    Public Function leerNombreEmpresa(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = "toleranciacredito"
        CAMPOS(2, 0) = "servidorprincipal"
        CAMPOS(3, 0) = "instituciondonacion"
        CAMPOS(4, 0) = "montomaximocheque"
        CAMPOS(5, 0) = ""
        CAMPOS(0, 2) = "g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreEmpresa = sql.response(0, 3)
            toleranciacredito = sql.response(1, 3)
            
            sqlventas.servidor_principal = sql.response(2, 3)
            sqlventas.servidor_ventas = Servidor
            entidaddonacion = sql.response(3, 3)
            MONTOMAXIMOCHEQUE = sql.response(4, 3)
            
            sqlventas.cliente_sql = clientesistema
            
        Else
            leerNombreEmpresa = ""
        End If
    End Function
'=============================================================================
'LEER NOMBRE EMPRESA
'=============================================================================

'=============================================================================
'LEER DIRECCION EMPRESA
'=============================================================================
    Public Function leerDireccionEmpresa(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "CONCAT(direccion, ' - ', ciudad)"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerDireccionEmpresa = sql.response(0, 3)
        Else
            leerDireccionEmpresa = ""
        End If
    End Function
        Public Function leernombreempresaremu(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = clientesistema & "remu" & ".maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leernombreempresaremu = sql.response(0, 3)
        Else
            leernombreempresaremu = ""
        End If
    End Function
    Public Function leerrutempresaremu(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "rut"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = clientesistema & "remu" & ".maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerrutempresaremu = sql.response(0, 3)
        Else
            leerrutempresaremu = ""
        End If
    End Function
     Public Function leerDireccionEmpresa2(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "direccion"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerDireccionEmpresa2 = sql.response(0, 3)
        Else
            leerDireccionEmpresa2 = ""
        End If
    End Function
    Public Function leerCiudadEmpresa(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "ciudad"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCiudadEmpresa = sql.response(0, 3)
        Else
            leerCiudadEmpresa = ""
        End If
    End Function
'=============================================================================
'LEER DIRECCION EMPRESA
'=============================================================================

'=============================================================================
'LEER RUT EMPRESA
'=============================================================================
    Public Function leerRutEmpresa(ByVal codigo As String) As String
        
        Dim op As Integer
        Dim rut As String
        Dim dv As String
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "rut"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            rut = Format(Left(sql.response(0, 3), 9), "###,###,##0")
            dv = Right(sql.response(0, 3), 1)
            leerRutEmpresa = rut & "-" & dv
        Else
            leerRutEmpresa = ""
        End If
    End Function
'=============================================================================
'LEER RUT EMPRESA
'=============================================================================

'=============================================================================
'LEER GIRO EMPRESA
'=============================================================================
    Public Function leerGiroEmpresa(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "giro"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerGiroEmpresa = sql.response(0, 3)
        Else
            leerGiroEmpresa = ""
        End If
    End Function
'=============================================================================
'LEER GIRO EMPRESA
'=============================================================================

'=============================================================================
'LEER FONO EMPRESA
'=============================================================================
    Public Function leerFonoEmpresa(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "fono"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerFonoEmpresa = sql.response(0, 3)
        Else
            leerFonoEmpresa = ""
        End If
    End Function
'=============================================================================
'LEER FONO EMPRESA
'=============================================================================

'=============================================================================
'LEER IMPUESTO
'=============================================================================
    Public Function leerImpuesto(ByVal codigo As String) As Double
        
        Dim op As Integer
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = gestion
        
        csql.sql = "select porcentaje from g_maestroimpuestos where nombrecorto='" & codigo & "' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerImpuesto = resultados(0)
        Else
        leerImpuesto = 0
        End If
        
'        Set sql =new sqlventas.sqlventa
'        campos(0, 0) = "porcentaje"
'        campos(1, 0) = ""
'
'        campos(0, 2) = "g_maestroimpuestos"
'
'        condicion = "nombre = '" & codigo & "'"
'        op = 5
'        sql.response = campos
'        Set sql.conexion = gestion
'        call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerImpuesto = CDbl(sql.response(0, 3))
'        Else
'            leerImpuesto = "0"
'        End If
    End Function
'=============================================================================
'LEER IMPUESTO
'=============================================================================

'=============================================================================
'LEER TIPO IMPUESTO
'=============================================================================
    Public Function leerImpuestoProducto(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "i.nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_fijo_" & rubro & " AS mpf LEFT JOIN " & basedatos & ".g_maestroimpuestos AS i ON mpf.codigoimpuesto = i.codigo"
        
        condicion = "mpf.codigobarra = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerImpuestoProducto = sql.response(0, 3)
        Else
            leerImpuestoProducto = ""
        End If
    End Function
'=============================================================================
'LEER TIPO IMPUESTO
'=============================================================================

'=============================================================================
'LEER RUBRO
'=============================================================================
    Public Function leerRubro(ByVal codigo As String) As String
        Dim op As Integer
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "rubro"
        CAMPOS(1, 0) = ""

        CAMPOS(0, 2) = "g_maestroempresas"

        condicion = "codigo = '" & codigo & "'"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerRubro = sql.response(0, 3)
        Else
            leerRubro = ""
        End If
    End Function
'=============================================================================
'LEER RUBRO
'=============================================================================

'=============================================================================
'LEER RUBRO
'=============================================================================
    Public Function leerNombreRubro(ByVal codigo As String) As String
        Dim op As Integer
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""

        CAMPOS(0, 2) = "g_maestrorubros"

        condicion = "codigo = '" & codigo & "'"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNombreRubro = sql.response(0, 3)
        Else
            leerNombreRubro = ""
        End If
    End Function
'=============================================================================
'LEER RUBRO
'=============================================================================

'=============================================================================
'LEER ULTIMO FOLIO
'=============================================================================
    Public Function leer_Ultimo_Folio(ByVal campo As String, ByVal tabla As String, ByVal largo As Integer, ByRef con As rdoConnection, ByVal condicion As String) As String
        Dim op As Integer
        Dim cad As String
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "IFNULL(MAX(" & campo & ") + 1,'1')"
        CAMPOS(1, 0) = ""

        CAMPOS(0, 2) = tabla

        op = 5
        sql.response = CAMPOS
        Set sql.conexion = con
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            cad = String(largo - Len(sql.response(0, 3)), "0") & sql.response(0, 3)
        Else
            cad = ""
        End If
        leer_Ultimo_Folio = cad
    End Function
'=============================================================================
'LEER ULTIMO FOLIO
'=============================================================================

'=============================================================================
'LEER DOCUMENTO NULO
'=============================================================================
    Public Function leerDocumentoNulo(ByVal TIPO As String, ByVal numero As String) As Boolean
        Dim op As Integer
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nula"
        CAMPOS(1, 0) = ""

        CAMPOS(0, 2) = "sv_documento_cabeza_" + empresaActiva

        condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND numero = '" & numero & "'"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            If sql.response(0, 3) = "N" Then
                leerDocumentoNulo = False
            Else
                leerDocumentoNulo = True
            End If
        Else
            leerDocumentoNulo = False
        End If
    End Function
'=============================================================================
'LEER DOCUMENTO NULO
'=============================================================================

'=============================================================================
'LEER SALDO
'=============================================================================
    Public Function leerSaldo(ByVal folio As String) As String
        Dim op As Integer
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "montocuota - abonocuota"
        CAMPOS(1, 0) = ""

        CAMPOS(0, 2) = "sv_credito_detalle"

        condicion = "local = '" & empresaActiva & "' AND folio = '" & folio & "' AND abonocuota < montocuota AND abonocuota <> '0' ORDER BY cuota"
        
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerSaldo = sql.response(0, 3)
        Else
            leerSaldo = "0"
        End If
    End Function
'=============================================================================
'LEER SALDO
'=============================================================================

'=============================================================================
'LEER TIPO PAGO
'=============================================================================
    
'LEER TIPO PAGO
'=============================================================================

'=============================================================================
'LEER INTERES NORMAL
'=============================================================================
'    Public Function leerInteresNormal(ByVal codigo As String) As Double
'        Dim op As Integer
'
'        Set sql =new sqlventas.sqlventa
'        campos(0, 0) = "interesnormal"
'        campos(1, 0) = ""
'
'        campos(0, 2) = "g_maestroempresas"
'
'        condicion = "codigo = '" & empresaactiva & "'"
'
'        op = 5
'        sql.response = campos
'        Set sql.conexion = gestion
'        call sql.sqlventas(op, condicion)
'        If sql.status = 0 Then
'            leerInteresNormal = CDbl(Replace(sql.response(0, 3), ".", ","))
'        Else
'            leerInteresNormal = 0
'        End If
'    End Function
'=============================================================================
'LEER INTERES NORMAL
'=============================================================================

'=============================================================================
'LEER INTERES MORA
'=============================================================================
    Public Function leerInteresMora(ByVal codigo As String) As String
        Dim op As Integer
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "interesmora"
        CAMPOS(1, 0) = "diasgracia"
        CAMPOS(2, 0) = ""
        CAMPOS(0, 2) = "g_maestroempresas"

        condicion = "codigo = '" & empresaActiva & "'"

        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerInteresMora = CDbl(Replace(sql.response(0, 3), ".", ","))
        diasgracia = sql.response(1, 3)
        Else
            leerInteresMora = 0
        End If
    End Function
'=============================================================================
'LEER INTERES MORA
'=============================================================================

'=============================================================================
'LEER TIPOS DE PAGO
'=============================================================================
    Public Function leerTiposDePago(ByVal numero As String, ByVal TIPO As String, ByRef rollo As Adodc) As String
        Dim tabla As String
        Dim codigo As String
        tabla = "SELECT dp.tipopago, SUM(dp.monto) AS monto "
        tabla = tabla & "FROM sv_documento_pagos_" + empresaActiva + " AS dp INNER JOIN sv_documento_cabeza_" + empresaActiva + " AS dc ON dp.local = dc.local AND dp.tipo = dc.tipo AND dp.numero = dc.numero "
        tabla = tabla & "WHERE dp.local = '" & empresaActiva & "' AND dp.tipo = '" & TIPO & "' AND dp.numero = '" & numero & "' "
        tabla = tabla & "GROUP BY dp.tipopago ORDER BY dp.tipopago ASC"
    
        Call ConectarControlData(rollo, Servidor, baseVentas & empresaActiva, usuario, password, tabla)
        
        If rollo.Recordset.RecordCount > 0 Then
            rollo.Recordset.MoveFirst
            leerTiposDePago = ""
            While Not rollo.Recordset.EOF
                codigo = rollo.Recordset.Fields("tipopago")
                Select Case codigo
                    Case "1"    'EFECTIVO
                        leerTiposDePago = leerTiposDePago & "EFECTIVO " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
                    Case "2"    'CHEQUE
                        leerTiposDePago = leerTiposDePago & "CHEQUE " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
                    Case "6"    'CREDITO DIRECTO
                        leerTiposDePago = leerTiposDePago & "CREDITO DIRECTO " & Format(rollo.Recordset.Fields("monto"), "$ ###,###,##0") & " / "
                End Select
                rollo.Recordset.MoveNext
            Wend
        End If
    End Function
'=============================================================================
'LEER TIPOS DE PAGO
'=============================================================================

'=============================================================================
'LEER VENCIMIENTO
'=============================================================================
    Public Function leerVencimiento(ByVal TIPO As String, ByVal numero As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "DATE_FORMAT(vencimiento,'%d-%m-%Y')"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_cabeza_" + empresaActiva
        
        condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND numero = '" & numero & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerVencimiento = sql.response(0, 3)
        Else
            leerVencimiento = ""
        End If
    End Function
'=============================================================================
'LEER VENCIMIENTO
'=============================================================================


'=============================================================================
'ACTUALIZACION DE STOCK
'=============================================================================
   
   Public Sub actualiza_stock(operacion, producto, venta, compra, bodega, año, cantidad, precio, fecha, rut, loc)
    
        Dim mesproceso As String
        Dim saldo As Double
        Dim suma As Double
        Dim mespro As String
        Dim sumamontos As Double
        Dim sumaunidades As Double
        Dim compraventam As String
        Dim compraventau As String
        
        Dim op As Integer
        
        CAMPOS(0, 2) = "r_maestroproductos_stock_" & rubro
        CAMPOS(0, 0) = "codigo"
        CAMPOS(1, 0) = "stockactual"
        If operacion = "+" Then mesproceso = "ingreso" & Format(fecha, "mm")
        If operacion = "-" Then mesproceso = "egreso" & Format(fecha, "mm")
        
        CAMPOS(2, 0) = mesproceso
        CAMPOS(3, 0) = ""
        condicion = "codigo='" + producto + "' and local='" + loc + "' and bodega='" + bodega + "' and año='" + año + "' limit 0,1"
    
        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status <> 0 Then Exit Sub 'Call crear: GoTo inicio:
        If operacion = "+" Then saldo = CDbl(sqlventas.response(1, 3)) + CDbl(cantidad)
        If operacion = "-" Then saldo = CDbl(sqlventas.response(1, 3)) - CDbl(cantidad)
        suma = CDbl(sqlventas.response(2, 3)) + CDbl(cantidad)
            
        CAMPOS(0, 1) = producto
        CAMPOS(1, 1) = Replace(Format(saldo, "###########0.0000"), ",", ".")
        CAMPOS(2, 1) = Replace(Format(suma, "###########0.0000"), ",", ".")
        CAMPOS(3, 1) = ""
        condicion = "codigo='" + producto + "' and local='" + loc + "' and bodega='" + bodega + "' and año='" + año + "'"
        op = 3
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
        
        Rem comneta para que no actualice y no se caiga
        'If venta = "S" Then Call ACTUALIZAESTADISTICA(producto, cantidad, PRECIO, rut, "V", fecha, año, loc)
        'If compra = "S" Then Call ACTUALIZAESTADISTICA(producto, cantidad, PRECIO, rut, "C", fecha, año, loc)
        
    End Sub
    
    Public Sub desactualiza_stock(operacion, producto, venta, compra, bodega, año, cantidad, precio, fecha, rut, loc)
    
        Dim mesproceso As String
        Dim saldo As Double
        Dim suma As Double
        Dim mespro As String
        Dim sumamontos As Double
        Dim sumaunidades As Double
        Dim compraventam As String
        Dim compraventau As String
        
        Dim op As Integer
        
        CAMPOS(0, 2) = "r_maestroproductos_stock_" & rubro
        CAMPOS(0, 0) = "codigo"
        CAMPOS(1, 0) = "stockactual"
        If operacion = "+" Then mesproceso = "ingreso" & Format(fecha, "mm")
        If operacion = "-" Then mesproceso = "egreso" & Format(fecha, "mm")
        CAMPOS(2, 0) = mesproceso
        CAMPOS(3, 0) = ""
        condicion = "codigo ='" + producto + "' and local='" + loc + "' and bodega='" + bodega + "' and año='" + año + "' limit 0,1"
    
        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
        
        If sqlventas.Status <> 0 Then Exit Sub  'Stop 'Call crear: GoTo inicio:
        If operacion = "+" Then saldo = CDbl(sqlventas.response(1, 3)) - CDbl(cantidad)
        If operacion = "-" Then saldo = CDbl(sqlventas.response(1, 3)) + CDbl(cantidad)
    
        suma = CDbl(sqlventas.response(2, 3)) - CDbl(cantidad)
            
        CAMPOS(0, 1) = producto
        CAMPOS(1, 1) = Replace(Format(saldo, "###########0.0000"), ",", ".")
        CAMPOS(2, 1) = Replace(Format(suma, "###########0.0000"), ",", ".")
        condicion = "codigo='" + producto + "' and local='" + loc + "' and bodega='" + bodega + "' and año='" + año + "'"
        op = 3
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
        
    End Sub
    
    Sub ACTUALIZAESTADISTICA(codigo, unidad, precio, rut, operacion, fecha, año, loc)
        
        Dim op As Integer
        
        CAMPOS(0, 2) = "r_maestroproductos_estadistica_" & rubro
        
        If operacion = "V" Then
        CAMPOS(0, 0) = "uventas" & Format(fecha, "mm")
        CAMPOS(1, 0) = "ventas" & Format(fecha, "mm")
        CAMPOS(2, 0) = ""
        End If
        
        If operacion = "C" Then
        CAMPOS(0, 0) = "ucompras" & Format(fecha, "mm")
        CAMPOS(1, 0) = "compras" & Format(fecha, "mm")
        CAMPOS(2, 0) = ""
        End If
        
        condicion = "codigo='" + codigo + "' and local='" + loc + "' and año='" + año + "' limit 0,1 "
    
        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status <> 0 Then Stop 'Call crear: GoTo inicio:
    
    If operacion = "V" Then
        CAMPOS(0, 0) = "uventas" & Format(fecha, "mm")
        CAMPOS(1, 0) = "ventas" & Format(fecha, "mm")
        CAMPOS(2, 0) = "fechaultimaventa"
        CAMPOS(3, 0) = "cantidadultimaventa"
        CAMPOS(4, 0) = "precioultimaventa"
        CAMPOS(5, 0) = "clienteultimaventa"
        CAMPOS(6, 0) = ""
        End If
        
        If operacion = "C" Then
        CAMPOS(0, 0) = "ucompras" & Format(fecha, "mm")
        CAMPOS(1, 0) = "compras" & Format(fecha, "mm")
        CAMPOS(2, 0) = "fechaultimacompra"
        CAMPOS(3, 0) = "cantidadultimacompra"
        CAMPOS(4, 0) = "precioultimacompra"
        CAMPOS(5, 0) = "proveedorultimacompra"
        CAMPOS(6, 0) = ""
        End If
        
        
        
        CAMPOS(0, 1) = sqlventas.response(0, 3) + unidad
        CAMPOS(1, 1) = sqlventas.response(1, 3) + unidad
        
        CAMPOS(2, 1) = fecha
        CAMPOS(3, 1) = unidad
        CAMPOS(4, 1) = unidad * precio
        CAMPOS(5, 1) = rut
        condicion = "codigo='" + codigo + "' and local='" + loc + "' and año='" + año + "'"
        op = 3
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
        
    End Sub
'=============================================================================
'ACTUALIZACION DE STOCK
'=============================================================================

'====================================================================================
'Rutina para saber si un archivo existe
'====================================================================================
    Function ExisteArchivo(sNombreArchivo As String) As Boolean
        Dim AttrDev%
        On Error Resume Next
        AttrDev = GetAttr(sNombreArchivo)
        If Err.Number Then
            Err.Clear
            ExisteArchivo = False
        Else
            ExisteArchivo = True
        End If
    End Function

'====================================================================================
'Rutina para obtener la configuracion del servidor y financiera
'====================================================================================
    Function leeArchivo(ByVal TIPO As String, ByVal ARCHIVO As String) As String
        Dim numfic As Integer
        Dim cad As String
        Dim cadaux As String
        numfic = FreeFile
        If ExisteArchivo(ARCHIVO) = True Then
            Open ARCHIVO For Input As #numfic
            Do While Not EOF(numfic)
                Line Input #numfic, cad
                cadaux = Left(cad, InStr(1, cad, "=") - 1)
                If cadaux = TIPO Then
                    leeArchivo = Right(cad, Len(cad) - InStr(1, cad, "="))
                    Exit Do
                End If
            Loop
            Close #numfic
        Else
            'Call mensaje.mostrarMensaje("ERROR", "SU CAJA NO TIENE EL ARCHIVO DE CONFIGURACION", "SOLICITELO CON EL ADMINISTADOR")
        End If
    End Function
    
'====================================================================================
'Rutina para guardar la configuracion del servidor y financiera
'====================================================================================
    Sub escribeArchivo(ByVal TIPO As String, ByVal cadena As String, ByVal ARCHIVO As String)
        Dim numfic, numficaux As Integer
        Dim cad, cadaux As String
        numfic = FreeFile
        If TIPO = "SERVIDOR" Or TIPO = "SISTEMA" Then
            Open App.Path & "\" & ARCHIVO For Output As #numfic
            Close #numfic
        End If
        numfic = FreeFile
        Open App.Path & "\" & ARCHIVO For Append As #numfic
        'numficaux = FreeFile
        'Open App.Path & "\" & archivo & ".tmp" For Output As #numficaux
        Do While Not EOF(numfic)
            Line Input #numfic, cad
            cadaux = Left(cad, InStr(1, cad, "=") - 1)
            If cadaux = TIPO Then
                cadaux = TIPO & "=" & cadena
                Print #numficaux, cadaux
            Else
                Print #numficaux, cad
            End If
        Loop
        'Close #numfic
        'Close #numficaux
        'Open App.Path & "\" & archivo For Output As #numfic
        'Open App.Path & "\" & archivo & ".tmp" For Input As #numficaux
        'Do While Not EOF(numficaux)
        '    Line Input #numficaux, cad
        Print #numfic, TIPO & "=" & cadena '& vbCrLf
        'Loop
        Close #numfic
        'Close #numficaux
        'Kill App.Path & "\" & archivo & ".tmp"
    End Sub
    
    Sub escribeArchivoRuta(ByVal TIPO As String, ByVal cadena As String, ByVal ARCHIVO As String)
        Dim numfic As Integer
        numfic = FreeFile
        If TIPO = "SISTEMA" Then
            Open ARCHIVO For Output As #numfic
            Close #numfic
        End If
        numfic = FreeFile
        Open ARCHIVO For Append As #numfic
        Print #numfic, TIPO & "=" & cadena
        Close #numfic
    End Sub


    Public Sub cabezaInforme(ByVal codigovendedor As String, ByRef impresion As Grid, ByVal titulo As String, ByVal orientacion As Integer)
        Dim cabeza As FlexCell.ReportTitle
        Dim I As Integer
        
        For I = impresion.ReportTitles.Count To 0 Step -1
            impresion.ReportTitles.Remove (I)
        Next I
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = titulo
        cabeza.Align = cellCenter
        cabeza.Font.Bold = True
        cabeza.Font.Underline = True
        cabeza.PrintOnAllPages = True
        impresion.ReportTitles.Add cabeza
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = leerNombreEmpresa(empresaActiva)
        cabeza.Align = cellLeft
        cabeza.PrintOnAllPages = True
        impresion.ReportTitles.Add cabeza
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = leerDireccionEmpresa(empresaActiva)
        cabeza.Align = cellLeft
        cabeza.PrintOnAllPages = True
        impresion.ReportTitles.Add cabeza
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = leerRutEmpresa(empresaActiva)
        cabeza.Align = cellLeft
        cabeza.PrintOnAllPages = True
        impresion.ReportTitles.Add cabeza
        
        If codigovendedor <> "" Then
            Set cabeza = New FlexCell.ReportTitle
            cabeza.text = codigovendedor & " " & leerNombreVendedor(codigovendedor)
            cabeza.Align = cellCenter
            cabeza.PrintOnAllPages = True
            impresion.ReportTitles.Add cabeza
        End If
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = " "
        cabeza.PrintOnAllPages = True
        
        impresion.PageSetup.Orientation = orientacion
        
        impresion.ReportTitles.Add cabeza
        
         'PIE DE PAGINA
    impresion.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & usuarioSistema
    impresion.PageSetup.FooterAlignment = cellRight
    impresion.PageSetup.FooterFont.Name = "Verdana"
    impresion.PageSetup.FooterFont.Size = 7
    End Sub

    Public Sub cambiaColor(ByRef frmXP As FrameXp)
        Dim aux
        aux = frmXP.ColorBarraAbajo
        frmXP.ColorBarraAbajo = frmXP.ColorBarraArriba
        frmXP.ColorBarraArriba = aux
    End Sub

    Public Sub cargaCabeza(ByVal titulo As String, ByVal codLoc As String, ByRef impresion As Grid)
        Dim cabeza As FlexCell.ReportTitle
        Dim I As Integer
        
        impresion.ReportTitles.Clear
        'For i = impresion.ReportTitles.Count To 0 Step -1
        '    impresion.ReportTitles.Remove (i)
        'Next i
        
        Set cabeza = New FlexCell.ReportTitle
                
        cabeza.text = titulo
        cabeza.Align = cellCenter
        cabeza.Font.Bold = True
        cabeza.Font.Underline = True
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = leerNombreEmpresa(codLoc)
        cabeza.Align = cellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = leerDireccionEmpresa(codLoc)
        cabeza.Align = cellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = "RUT: " & leerRutEmpresa(codLoc)
        cabeza.Align = cellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = " "
        cabeza.Align = cellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
    End Sub

    Public Sub cargaCabezaRetencion(ByVal titulo As String, ByVal codLoc As String, ByRef impresion As Grid, ByVal fecha As String)
        Dim cabeza As FlexCell.ReportTitle
        Dim I As Integer
        
        impresion.ReportTitles.Clear
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = titulo
        cabeza.Align = cellCenter
        cabeza.Font.Bold = True
        cabeza.Font.Underline = True
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = leerNombreEmpresa(codLoc)
        cabeza.Align = cellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = "RUT: " & leerRutEmpresa(codLoc)
        cabeza.Align = cellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = "INFORMACION MES: " & Format(fecha, "mmmm") & " DE " & Format(fecha, "yyyy")
        cabeza.Align = cellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = "&P"
        cabeza.Align = cellRight
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = " "
        cabeza.Align = cellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
    End Sub

    Public Function calculaSemana(ByVal fecha As String) As String
        Dim fecha1 As String
        fecha1 = Year(fecha) & "-01-01"
        calculaSemana = Format(DateDiff("ww", fecha1, fecha, vbUseSystemDayOfWeek, vbUseSystem) + 1, "00")
    End Function
    
    Public Sub verificaImpresora(ByVal TIPO As Integer, ByRef impresion As Grid)
        Select Case TIPO
            Case 1 'FACTURAS
                If impFacturas(1) = "" Then
                    impresion.PrintPreview
                Else
                    Set Printer = Printers(impFacturas(1))
                    If Printer.DeviceName = impFacturas(2) Then
                        impresion.PageSetup.PrinterName = impFacturas(2)
                        If Val(impFacturas(3)) <> 0 And Val(impFacturas(4)) <> 0 Then
                            impresion.PageSetup.PaperWidth = impFacturas(3)
                            impresion.PageSetup.PaperHeight = impFacturas(4)
                        End If
                        If imprimeDirecto = True Then
                            impresion.DirectPrint
                        Else
                            impresion.PrintPreview
                        End If
                    Else
                        impresion.PrintPreview
                    End If
                End If
            Case 2 'BOLETAS
                If impBoletas(1) = "" Then
                    impresion.PrintPreview
                Else
                    Set Printer = Printers(impBoletas(1))
                    If Printer.DeviceName = impBoletas(2) Then
                        impresion.PageSetup.PrinterName = impBoletas(2)
                        If Val(impBoletas(3)) <> 0 And Val(impBoletas(4)) <> 0 Then
                            impresion.PageSetup.PaperWidth = impBoletas(3)
                            impresion.PageSetup.PaperHeight = impBoletas(4)
                        End If
                        If imprimeDirecto = True Then
                            impresion.DirectPrint
                        Else
                            impresion.PrintPreview
                        End If
                    Else
                        impresion.PrintPreview
                    End If
                End If
            Case 3 'GUIAS
                If impGuias(1) = "" Then
                    impresion.PrintPreview
                Else
                    Set Printer = Printers(impGuias(1))
                    If Printer.DeviceName = impGuias(2) Then
                        impresion.PageSetup.PrinterName = impGuias(2)
                        If Val(impGuias(3)) <> 0 And Val(impGuias(4)) <> 0 Then
                            impresion.PageSetup.PaperWidth = impGuias(3)
                            impresion.PageSetup.PaperHeight = impGuias(4)
                        End If
                        If imprimeDirecto = True Then
                            impresion.DirectPrint
                        Else
                            impresion.PrintPreview
                        End If
                    Else
                        impresion.PrintPreview
                    End If
                End If
            Case 4 'NOTAS DE CREDITO
                If impNCredito(1) = "" Then
                    impresion.PrintPreview
                Else
                    Set Printer = Printers(impNCredito(1))
                    If Printer.DeviceName = impNCredito(2) Then
                        impresion.PageSetup.PrinterName = impNCredito(2)
                        If Val(impNCredito(3)) <> 0 And Val(impNCredito(4)) <> 0 Then
                            impresion.PageSetup.PaperWidth = impNCredito(3)
                            impresion.PageSetup.PaperHeight = impNCredito(4)
                        End If
                        If imprimeDirecto = True Then
                            impresion.DirectPrint
                        Else
                            impresion.PrintPreview
                        End If
                    Else
                        impresion.PrintPreview
                    End If
                End If
            Case 5 'OTROS
                If impOtros(1) = "" Then
                    impresion.PrintPreview
                Else
                    Set Printer = Printers(impOtros(1))
                    If Printer.DeviceName = impOtros(2) Then
                        impresion.PageSetup.PrinterName = impOtros(2)
                        If Val(impOtros(3)) <> 0 And Val(impOtros(4)) <> 0 Then
                            impresion.PageSetup.PaperWidth = impOtros(3)
                            impresion.PageSetup.PaperHeight = impOtros(4)
                        End If
                        If imprimeDirecto = True Then
                            impresion.DirectPrint
                        Else
                            impresion.PrintPreview
                        End If
                    Else
                        impresion.PrintPreview
                    End If
                End If
        End Select
        
    End Sub

'Encripta una cadena de caracteres.
'S = Cadena a encriptar
'P = Password
Function EncryptStr(ByVal s As String, ByVal p As String) As String
    Dim I As Integer, R As String
    Dim C1 As Integer, C2 As Integer
    R = ""
    If Len(p) > 0 Then
        For I = 1 To Len(s)
            C1 = Asc(Mid(s, I, 1))
            If I > Len(p) Then
                C2 = Asc(Mid(p, I Mod Len(p) + 1, 1))
            Else
                C2 = Asc(Mid(p, I, 1))
            End If
            C1 = C1 + C2 + 64
            If C1 > 255 Then C1 = C1 - 256
            If C1 = 13 Then
                R = R & "FLAG"
            Else
                R = R & Chr(C1)
            End If
        Next I
    Else
        R = s
    End If
    EncryptStr = R
End Function


'Desencripta una cadena de caracteres.
'S = Cadena a desencriptar
'P = Password

Function UnEncryptStr(ByVal s As String, ByVal p As String) As String
    Dim I As Integer, R As String
    Dim C1 As Integer, C2 As Integer
    R = ""
    If Len(p) > 0 Then
        If InStr(1, s, "FLAG", vbBinaryCompare) <> 0 Then
            s = Replace(s, "FLAG", Chr(13))
        End If
        For I = 1 To Len(s)
            C1 = Asc(Mid(s, I, 1))
            If I > Len(p) Then
                C2 = Asc(Mid(p, I Mod Len(p) + 1, 1))
            Else
                C2 = Asc(Mid(p, I, 1))
            End If
            C1 = C1 - C2 - 64
            If Sgn(C1) = -1 Then C1 = 256 + C1
                R = R + Chr(C1)
        Next I
    Else
        R = s
    End If
    UnEncryptStr = R
End Function

Public Sub leerDatosConectar()
    Dim ss As String
    Dim K As Integer
    
    Close 20
        Open App.Path + "\confiventas2.txt" For Input As #20
    While EOF(20) = False
    Input #20, ss
    
    If Mid(ss, 1, 8) = "SERVIDOR" Then
        Servidor = Mid(ss, 10, Len(ss) - 9)
    End If
    If Mid(ss, 1, 7) = "USUARIO" Then
        usuario = Mid(ss, 9, Len(ss) - 8)
    End If
    If Mid(ss, 1, 8) = "PASSWORD" Then
        password = Mid(ss, 10, Len(ss) - 9)
    End If
    If Mid(ss, 1, 9) = "BASEDATOS" Then
        basedatos = Mid(ss, 11, Len(ss) - 10)
    End If
    If Mid(ss, 1, 10) = "BASEVENTAS" Then
        baseVentas = Mid(ss, 12, Len(ss) - 11)
    End If
    If Mid(ss, 1, 7) = "EMPRESA" Then
        empresaActiva = Mid(ss, 9, Len(ss) - 8)
    End If
    If Mid(ss, 1, 6) = "BODEGA" Then
        bodega = Mid(ss, 8, Len(ss) - 7)
    End If
    If Mid(ss, 1, 4) = "CAJA" Then
        idCaja = Mid(ss, 6, Len(ss) - 5)
    End If
    If Mid(ss, 1, 4) = "RUTA" Then
        rutaUpdate = Mid(ss, 6, Len(ss) - 5)
    End If
    If Mid(ss, 1, 13) = "IMPRESORAPAGO" Then
        IMPRESORAPAGO = Mid(ss, 15, Len(ss) - 5)
    End If
    If Mid(ss, 1, 12) = "BODEGARETIRO" Then
        BODEGARETIRO = Mid(ss, 14, Len(ss) - 5)
    End If
    If Mid(ss, 1, 16) = "IMPRESORACREDITO" Then
        impresoracredito = Mid(ss, 18, Len(ss) - 5)
    End If
    If Mid(ss, 1, 11) = "ASEGURADORA" Then
        ASEGURADORA = Mid(ss, 13, Len(ss) - 12)
    End If
    If Mid(ss, 1, 14) = "IMPRIMEDIRECTO" Then
    If Mid(ss, 16, Len(ss) - 14) = "S" Then
        imprimeDirecto = True
        Else
        imprimeDirecto = False
    End If
    End If
    
    
    Wend
        Close 20
    
    
        For K = 1 To Len(baseVentas)
        If Mid(baseVentas, K, 1) = "_" Then clientesistema = Mid(baseVentas, 1, K)
        Next K
        
        baseteso = clientesistema & "teso"
        baseauditoria = clientesistema
        segundosespera = "2"
        
        Call Conectar(Servidor, basedatos, usuario, password)
        
        rubro = leerRubro(empresaActiva)
        Call ConectarRubro(Servidor, basedatos, usuario, password)
        Call Conectarventas(Servidor, baseVentas & empresaActiva, usuario, password)
        iva = leerImpuesto("IVA")
        iha = leerImpuesto("IHA")
        fechasistema = Format(Now, "yyyy-mm-dd")
        
    
    envia = False
End Sub


    Public Sub seleccionaUno(ByVal KeyCode As Integer, ByRef txt As TextBox)
        Select Case KeyCode
            Case 37
                If txt.SelStart > 0 And txt.SelStart < txt.MaxLength Then
                    txt.SelStart = txt.SelStart - 1
                End If
                txt.SelLength = 1
            Case 39
                If txt.SelStart > 0 And txt.SelStart <= txt.MaxLength Then
                    txt.SelStart = txt.SelStart - 1
                End If
                txt.SelLength = 1
            Case 96, 97, 98, 99, 100, 101, 102, 103, 104, 105
                txt.SelLength = 1
        End Select
    End Sub

Public Function comparaArchivos(File1 As String, file2 As String) As Boolean
    Dim issame As Boolean
    Dim whole As Double
    Dim part As Long
    Dim buffer1 As String
    Dim buffer2 As String
    Dim start As Long
    Dim x As Long
    Dim nf1 As Integer
    Dim nf2 As Integer
    
    nf1 = FreeFile
    Open File1 For Binary As #nf1
    nf2 = FreeFile
    Open file2 For Binary As #nf2
    issame = True
    If LOF(nf1) <> LOF(nf2) Then
        issame = False
        comparaArchivos = False
    Else
        whole = LOF(nf1) \ 10000
        part = LOF(nf1) Mod 10000
        buffer1 = String(10000, 0)
        buffer2 = String(10000, 0)
        start = 1
        For x = 1 To whole
            Get #nf1, start, buffer1
            Get #nf2, start, buffer2
            If buffer1 <> buffer2 Then
                issame = False
                Exit For
            End If
            start = start + 10000
        Next x
        buffer1 = String(part, 0)
        buffer2 = String(part, 0)
        Get #nf1, start, buffer1
        Get #nf2, start, buffer2
        If buffer1 <> buffer2 Then
            issame = False
        End If
        If issame = True Then
            comparaArchivos = True
        Else
            comparaArchivos = False
        End If
    End If
    Close #nf1
    Close #nf2
End Function

Public Sub VisualFileCopy(ByVal SourceFileName As String, ByVal TargetFileName As String)
       Dim I As Integer
       Dim SourceFileNo As Integer
       Dim TargetFileNo As Integer
       Dim SourceFileSize As Long
       Dim CopyBuffer As String
    
       On Error GoTo FileCopyErrorRoutine
       SourceFileSize = FileLen(SourceFileName)
       CopyBuffer = Space$(25000)             'AS LARGE AS POSSIBLE UNDER 65,000
    
    '--KILL THE CURRENT TARGET FILE IF IT EXISTS
       If Len(Dir$(TargetFileName)) Then
          Kill TargetFileName
       End If
    
    '--OPEN FILES
       SourceFileNo = FreeFile
       Open SourceFileName For Binary Access Read As SourceFileNo
       TargetFileNo = FreeFile
       Open TargetFileName For Binary Access Write As TargetFileNo
    
    '--COPY SOURCE FILE TO TARGET FILE
       For I = 1 To SourceFileSize \ Len(CopyBuffer)
          Get #SourceFileNo, , CopyBuffer
    '--PROGRESS GUAGE
          Put #TargetFileNo, , CopyBuffer
          DoEvents
       Next I
    
    '--COPY ANY ODD PORTION OF THE SOURCE FILE REMAINING
       CopyBuffer = Space$(SourceFileSize - loc(TargetFileNo))
       If Len(CopyBuffer) Then
          Get #SourceFileNo, , CopyBuffer
          Put #TargetFileNo, , CopyBuffer
       End If
       Close SourceFileNo
       Close TargetFileNo
    
    Exit Sub
    
FileCopyErrorRoutine:
       MsgBox error$
       Exit Sub
    End Sub

Public Function GetSystemDir() As String
    Dim Buffer As String, Size As Long
    Const MAX_PATH = 260
    ' Inicializamos la cadena
    Buffer = String(MAX_PATH, 0)
    ' Recuperamos la trayectoria
    Size = GetSystemDirectory(Buffer, Len(Buffer))
    ' Si el resultado es distinto de cero, devolvemos la trayectoria
    '
    If Size <> 0 Then
        GetSystemDir = Left(Buffer, Size)
    End If
End Function

    Public Sub eliminarCheque(ByVal TIPO As String, ByVal numero As String)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "local = '" & empresaActiva & "' AND tipodocumento = '" & TIPO & "' AND numero = '" & numero & "'"
        op = 4
        CAMPOS(0, 2) = "sv_carteracheques"
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
    End Sub

    Public Function leerNotaCreditoFactura(ByVal NUMEROFACTURA As String) As Double
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "IFNULL(SUM(monto),0)"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_documento_notas"
        
        condicion = "local = '" & empresaActiva & "' AND numerofactura = '" & NUMEROFACTURA & "' "
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerNotaCreditoFactura = CDbl(sql.response(0, 3))
        Else
            leerNotaCreditoFactura = 0
        End If
    End Function

    Public Sub UnloadHijos(ByRef frm As MDIForm)
        With frm
            Do While Not (.ActiveForm Is Nothing)
                Unload .ActiveForm
            Loop
        End With
    End Sub
    
    Public Sub UnloadHijo(ByRef frm As MDIForm, ByVal nombre As String)
        With frm
            Do While Not (frm.ActiveForm Is Nothing)
                If .ActiveForm.Caption = nombre Then
                    Unload .ActiveForm
                    Exit Sub
                End If
            Loop
        End With
    End Sub

    Public Sub modificarPrecio(ByVal codigo As String, ByVal precio As String, ByVal tipoprecio As String, precioanterior As String)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "preciopuntoventa"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 1) = Replace(precio, ",", ".")
        CAMPOS(1, 1) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_precios_" & rubro
        
        condicion = "codigo = '" & codigo & "' AND codigoprecio = '" & tipoprecio & "'"
        op = 3
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
      If precio <> precioanterior Then Call HISTORICODEPRECIOS(codigo, precio, tipoprecio, precioanterior)
    End Sub
Public Sub HISTORICODEPRECIOS(ByVal codigo As String, ByVal precio As String, ByVal tipoprecio As String, precioanterior As String)
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "codigo"
        CAMPOS(2, 0) = "codigoprecio"
        CAMPOS(3, 0) = "horamodificacion"
        CAMPOS(4, 0) = "precioanterior"
        CAMPOS(5, 0) = "precionuevo"
        CAMPOS(6, 0) = "nuevo"
        CAMPOS(7, 0) = "preciocosto"
        CAMPOS(8, 0) = "fecha"
        CAMPOS(9, 0) = "usuario"
        CAMPOS(10, 0) = ""
        
        CAMPOS(0, 1) = "00"
        CAMPOS(1, 1) = codigo
        CAMPOS(2, 1) = tipoprecio
        CAMPOS(3, 1) = Time
        CAMPOS(4, 1) = precioanterior
        CAMPOS(5, 1) = precio
        CAMPOS(6, 1) = "N"
        CAMPOS(7, 1) = "0"
        CAMPOS(8, 1) = fechasistema
        CAMPOS(9, 1) = usuarioSistema
        
        CAMPOS(0, 2) = "r_maestroproductos_cambiosdeprecio_" & rubro
        
        
        op = 2
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
    End Sub
    Public Sub bloquearCliente(ByVal rut As String, ByVal sucursal As String)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "UPDATE sv_maestroclientes SET credito = 'B', estadoanterior = 'B' "
        csql.sql = csql.sql & "WHERE rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        csql.Execute
        Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub
    
    Public Sub estadoAnteriorCliente(ByVal rut As String, ByVal sucursal As String)
        Dim csql As rdoQuery
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "UPDATE sv_maestroclientes SET credito = estadoanterior "
        csql.sql = csql.sql & "WHERE rut = '" & rut & "' AND sucursal = '" & sucursal & "'"
        csql.Execute
            Call sincronizadatos(csql.sql, ventas)
        csql.Close
        Set csql = Nothing
    End Sub

    Public Function Redondeo(numero As Double, Optional Decima As Long) As Double
        Dim NumC As Double
        Dim NUM As String
        Dim deci As String
        NumC = numero - Fix(numero)
        If Len(CStr(NumC)) - 2 <= Decima Then
            Redondeo = Round(numero, Decima)
            Exit Function
        Else
            'el 3 es el 0, de Left(CStr(NumC) (=2) + 1, que es el ultimo decimal, el q redondeamos
            NUM = CStr(Fix(numero) + CDbl(Left(CStr(NumC), Decima + 3)))
        End If
        NUM = CStr(numero)
        deci = Right$(NUM, 1)
        'Si es 5 hacemos el redondeo hacia arriba, por eso el 6
        If deci = "5" Then
            deci = "6"
        End If
        NUM = Left$(NUM, Len(NUM) - 1) & deci
        If IsMissing(Decima) Then
            Redondeo = Round(CDbl(NUM))
        Else
            Redondeo = Round(CDbl(NUM), Decima)
        End If
    End Function


 Public Function leerPrecioProducto2(ByVal codigo As String, ByVal tipoprecio As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        
        CAMPOS(0, 0) = "preciopuntoventa"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_precios_" & rubro
        
        condicion = "codigo = '" & codigo & "' AND codigoprecio = '" & tipoprecio & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPrecioProducto2 = sql.response(0, 3)
           
        Else
           CAMPOS(0, 0) = "local"
           CAMPOS(1, 0) = "codigo"
           CAMPOS(2, 0) = "codigoprecio"
           CAMPOS(3, 0) = "preciopuntoventa"
           CAMPOS(4, 0) = ""
           CAMPOS(0, 1) = "00"
           CAMPOS(1, 1) = codigo
           CAMPOS(2, 1) = tipoprecio
           CAMPOS(3, 1) = "0"
        op = 2
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
           
            leerPrecioProducto2 = "0"
        End If
    End Function
'numToLet(Me.txtUserName.text, "PESO", "PESOS", "CENTAVO", "CENTAVOS", 0)
Function WORDNUM(ByVal numero As Variant, Optional TipoCambioSingular As String = "", Optional TipoCambioPlural As String, Optional subTipoCambioSingular As String, Optional subTipoCambioPlural As String, Optional xInternal As Long = 0) As String
     Dim snum As String, vNum() As String, x As Long, Y As Long, Z As Long, sTmp As String
     Dim D1 As String, D2 As String, D3 As String, D4 As String, DFinal As String
     Dim tNum As String, B1 As Boolean, B2 As Boolean, B3 As Boolean
     Dim wNum() As String, xNums As String, xWords As String, Nombres() As String
     
     '***********************************************************************************************
     '* Esta función convierte números en palabras, sin importar el contexto donde se encuentren    *
     '* La presición (por limitancia del lenguaje) es de 28B, Ej: 9999999999999999999999999999 max. *
     '***********************************************************************************************
     
     'Convierte el valor en un string
     snum = Trim(CStr(numero))
            
     'Procesa cada número que exista en la variable por separado
     If xInternal = 0 Then
        'Separa los números limpios de las palabras y los procesa por separado (no incluye números con letras)
        wNum = Split(snum, " ")
        For x = 0 To UBound(wNum)
            'Concatena los strings o números según corresponda
            If IsNumeric(wNum(x)) Then
               'Separa los enteros de los decimales para procesarlos por separado
               If Int(Val(wNum(x))) < wNum(x) Then
                  D1 = Int(Val(wNum(x)))
                  D2 = Mid(CStr(wNum(x)), Len(D1) + 2)
                  DFinal = DFinal & IIf(D1 < 0, "menos ", "") & WORDNUM(D1, TipoCambioSingular, TipoCambioPlural, 1) & " con "
                  DFinal = DFinal & WORDNUM(D2, subTipoCambioSingular, subTipoCambioPlural, , , 1) & " "
               Else
                  DFinal = DFinal & IIf(wNum(x) < 0, "menos ", "") & WORDNUM(wNum(x), TipoCambioSingular, TipoCambioPlural, subTipoCambioSingular, subTipoCambioPlural, 1) & " "
               End If
            Else
               DFinal = DFinal & wNum(x) & " "
            End If
        Next
     Else
        
        'ELimina el signo
        If Not IsNumeric(Left(snum, 1)) Then
           snum = Mid(snum, 2)
        End If
     
        'Elimina cualquier formato posible (incluye valores científicos)
        snum = Format(snum, "0")
        
        'Completa con ceros a la izquierda hasta obtener una longitud múltiplo de 3
        Do While Len(snum) Mod 3 <> 0
           snum = "0" & snum
        Loop
     
        'Dimenciona un arreglo con espacio para cada una de las centenas
        ReDim vNum(Len(snum) / 3 - 1)
        
        'Carga el arreglo con las centenas que corresponda
        For x = 0 To UBound(vNum, 1)
            vNum(x) = Mid(snum, (x + 1) * 3 - 2, 3)
        Next
         
        'Si el arreglo contiene una sola centena, la convierte en palabras
        If UBound(vNum, 1) = 0 Then
            'Asigna los dígitos de la centena y recuerda si son mayores que cero
            D3 = Left(snum, 1): B3 = Val(D3) > 0
            D2 = Mid(snum, 2, 1): B2 = Val(D2) > 0
            D1 = Right(snum, 1): B1 = Val(D1) > 0
            
            'Procesa las unidades
            Select Case D1
                   Case "1": DFinal = "un"
                   Case "2": DFinal = "dos"
                   Case "3": DFinal = "tres"
                   Case "4": DFinal = "cuatro"
                   Case "5": DFinal = "cinco"
                   Case "6": DFinal = "seis"
                   Case "7": DFinal = "siete"
                   Case "8": DFinal = "ocho"
                   Case "9": DFinal = "nueve"
            End Select
            
            'Procesa las decenas
            Select Case D2
                   Case "1"
                        'Maneja lógica del retrasado mental que puso nombres ilógicos a algunos números.
                        Select Case D1
                               Case "0": DFinal = "diez"
                               Case "1": DFinal = "once"
                               Case "2": DFinal = "doce"
                               Case "3": DFinal = "trece"
                               Case "4": DFinal = "catorce"
                               Case "5": DFinal = "quince"
                               Case "6": DFinal = "dieciséis"
                               Case Else
                                    DFinal = "dieci" & DFinal
                        End Select
                   Case "2"
                        If B1 Then
                           If D1 = "2" Then DFinal = "dós"
                           If D1 = "3" Then DFinal = "trés"
                           DFinal = "veinti" & DFinal
                        Else
                           DFinal = "veinte"
                        End If
                   Case "3": If B1 Then DFinal = "treinta y " & DFinal Else DFinal = "treinta"
                   Case "4": If B1 Then DFinal = "cuarenta y " & DFinal Else DFinal = "cuarenta"
                   Case "5": If B1 Then DFinal = "cincuenta y " & DFinal Else DFinal = "cincuenta"
                   Case "6": If B1 Then DFinal = "sesenta y " & DFinal Else DFinal = "sesenta"
                   Case "7": If B1 Then DFinal = "setenta y " & DFinal Else DFinal = "setenta"
                   Case "8": If B1 Then DFinal = "ochenta y " & DFinal Else DFinal = "ochenta"
                   Case "9": If B1 Then DFinal = "noventa y " & DFinal Else DFinal = "noventa"
            End Select
            
            'Procesa las centenas
            Select Case D3
                   Case "1": If B1 Or B2 Then DFinal = "ciento " & DFinal Else DFinal = "cien"
                   Case "2": If B1 Or B2 Then DFinal = "doscientos " & DFinal Else DFinal = "doscientos"
                   Case "3": If B1 Or B2 Then DFinal = "trescientos " & DFinal Else DFinal = "trescientos"
                   Case "4": If B1 Or B2 Then DFinal = "cuatrocientos " & DFinal Else DFinal = "cuatrocientos"
                   Case "5": If B1 Or B2 Then DFinal = "quinientos " & DFinal Else DFinal = "quinientos"
                   Case "6": If B1 Or B2 Then DFinal = "seiscientos " & DFinal Else DFinal = "seiscientos"
                   Case "7": If B1 Or B2 Then DFinal = "setecientos " & DFinal Else DFinal = "setecientos"
                   Case "8": If B1 Or B2 Then DFinal = "ochocientos " & DFinal Else DFinal = "ochocientos"
                   Case "9": If B1 Or B2 Then DFinal = "novecientos " & DFinal Else DFinal = "novecientos"
            End Select
            
            'Si es la ejecución principal efectua algunos arreglines
            If xInternal = 1 Then
               'Validación del cero
               If DFinal = "" Then DFinal = "cero"
               'Validación de terminados en "un"
               If Right(DFinal, 2) = "un" And TipoCambioSingular = "" Then DFinal = DFinal & "o"
            End If
            
        Else 'Si es más de una centena, las separa y procesa independientemente
            Y = -1
            Z = 1
            For x = UBound(vNum) To 0 Step -1
                Y = Y + 1
                
                'Convierte la centena en palabras
                tNum = WORDNUM(vNum(x), xInternal:=2)
                
                'Arregla la terminación "uno" cuando corresponde
                If Y = 0 And Right(tNum, 2) = "un" And TipoCambioSingular & TipoCambioPlural = "" Then tNum = tNum + "o"
                
                'Genera un valor temporal para poder modificar
                sTmp = tNum
                
                'Asigna los nombres genéricos principales
                Nombres = Split(" mil , millón , millones , billón , billones , trillón , trillones , cuatrillón , cuatrillones , quintillón , quintillones , sextillón , sextillones , septillón , septillones , octillón , octillones, nonillón , nonillones , decillón , decillones , undecillón , undecillones , duodecillón , duodecillones , tredecillón , tredecillones , cuatordecillón , cuatordecillones , quindecillón , quindecillones , sexdecillón , sexdecillones , septendecillón , septendecillones , octodecillón , octodecillones , novendecillón , novendecillones , vigintillón , vigintillones ", ",")
                
                'Controla que el índice de nombres no salga de los límites
                If Y > UBound(Nombres) Then
                   WORDNUM = "?"
                   Exit Function
                End If
                
                'Asigna los nombres correspondientes
                If Y Mod 2 > 0 Then
                   D1 = Nombres(0)
                   D2 = Nombres(Y - 1)
                ElseIf Y > 0 Then
                   D1 = Nombres(Y - 1)
                   D2 = Nombres(Y)
                Else
                   D1 = "": D2 = ""
                End If
                
                'Actualiza el nombre del número
                Select Case Y Mod 2
                       Case 0: If sTmp = "un" Then sTmp = sTmp & D1 Else sTmp = sTmp & IIf(tNum = "", "", D2)
                       Case Else
                            If sTmp = "un" Then sTmp = ""
                            sTmp = sTmp & IIf(tNum = "", "", D1)
                            If x = 0 And Y > 1 Then
                               If InStr(1, DFinal, D2, vbTextCompare) = 0 Then sTmp = sTmp & Mid(D2, 2)
                            End If
                End Select
                DFinal = sTmp & DFinal
            Next
        End If
     End If
     
     'Aplica el tipo de moneda cuando corresponda
     If xInternal = 1 Then DFinal = DFinal & " " & IIf(Format(snum, "#0") = "1", TipoCambioSingular, TipoCambioPlural)

     'Asigna el número en palabras
      WORDNUM = UCase(Trim(DFinal))
End Function


Sub cargatexto(ByRef caja As TextBox)
    caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub
Sub empresa()
   
   Dim op As Integer
   Dim K As Integer
    CAMPOS(0, 0) = "codigo"
    CAMPOS(1, 0) = "nombre"
    CAMPOS(2, 0) = "direccion"
    CAMPOS(3, 0) = "comuna"
    CAMPOS(4, 0) = "ciudad"
    CAMPOS(5, 0) = "rut"
    CAMPOS(6, 0) = "giro"
    CAMPOS(7, 0) = "rubro"
    CAMPOS(8, 0) = "codigocontable"
    CAMPOS(9, 0) = ""
    CAMPOS(0, 2) = "g_maestroempresas"
  
    condicion = "codigo = '" & LOCAL_PROCESO & "' ORDER BY codigo"
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = gestion
    
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 4 Then GoTo no:
    codigoempresa = sqlventas.response(0, 3)
    NOMBREEMPRESA = sqlventas.response(1, 3)
    DIRECCIONEMPRESA = sqlventas.response(2, 3)
    COMUNAEMPRESA = sqlventas.response(3, 3)
    CIUDADEMPRESA = sqlventas.response(4, 3)
    rutempresa = sqlventas.response(5, 3)
    GIROEMPRESA = sqlventas.response(6, 3)
    codigoCONTABLE = sqlventas.response(8, 3)
'    sqlventas.rubro_trabajo = rubro
   
    For K = 0 To 6
        DATOSEMPRESA(K) = sqlventas.response(K, 3)
    Next K
    
'    Principal.Caption = "GESTION COMERCIAL             Usuario : " + usuarioSistema + "     Empresa : " + nombreempresa + "                 Fecha : " & Format(fechasistema, "dd-mm-yyyy")
'    Call ConectarGestionRubro(servidor, baseDatos & rubro, usuario, password)
'    'Unload segurity
'    Unload Me
'    segurity.Hide
'
'    Principal.Show
no:
End Sub
  Function Verifica_Permiso(programa As String, operacion As String) As Boolean
    Dim I As Integer
    Dim columna As Integer
    'agrega modifica elimina
    
    
    
    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "SELECT todas," + operacion + " "
        cSql2.sql = cSql2.sql + "FROM segu_permisos "
        cSql2.sql = cSql2.sql + "where usuario='" + usuarioSistema + "' and programa='" + programa + "'"
        cSql2.Execute
        Verifica_Permiso = False
        
        If cSql2.RowsAffected > 0 Then
           Set resultados2 = cSql2.OpenResultset
        If resultados2(1) = 1 Then
        Verifica_Permiso = True
        
        Else
        Verifica_Permiso = False
        
        End If
        If resultados2(0) = 1 Then
        Verifica_Permiso = True
        
        
        
        End If
        End If

End Function
 Public Function leerCodigoProducto(ByVal codigo As String) As Boolean
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "codigobarra"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_fijo_" & rubro
        
        condicion = "codigobarra = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerCodigoProducto = True
            
        Else
            leerCodigoProducto = False
        
        End If
    End Function
Public Function leerUltimoFolioPAGO() As String
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    Dim CAMPOS(3, 3) As String
    
    CAMPOS(0, 0) = "IFNULL(MAX(numero) + 1,'0000000001')"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_cuotas_pago_cabeza"
    condicion = "local<>'99' "
    op = 5
    sql.response = CAMPOS
    Set sql.conexion = ventas
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        If sql.response(0, 3) <> "" And sql.response(0, 3) <> "0" Then
            leerUltimoFolioPAGO = Format(sql.response(0, 3), "0000000000")
        Else
            leerUltimoFolioPAGO = "0000000001"
        End If
    End If
End Function

Public Function LEErcreditoutilizado(rut, fecha) As Double


        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT sum(montocuota-abono)  "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle "
        csql.sql = csql.sql & "WHERE rut='" + rut + "'  and vencimientoactual<='" + Format(fecha, "yyyy-mm-dd") + "' "
        csql.sql = csql.sql & "group by rut "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
           Set resultado = csql.OpenResultset
        If resultado(0) < 0 Then
        LEErcreditoutilizado = 0
        Else
        
        LEErcreditoutilizado = resultado(0)
        End If
        
        Else
        LEErcreditoutilizado = 0
        End If
        
    End Function

'=============================================================================
'LEER TIPO VIVIENDA
'=============================================================================
    Public Function leerTipoVivienda(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestrotipovivienda"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerTipoVivienda = sql.response(0, 3)
        Else
            leerTipoVivienda = ""
        End If
    End Function
'=============================================================================
'LEER TIPO VIVIENDA
'=============================================================================
Public Function leerUltimofoliocaja(TIPO, caja) As String
    
    Dim op As Integer
   
    Dim numCaja As String
    Dim numero As String
    
    
    
    CAMPOS(0, 0) = "numero"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_documento_cabeza_" + empresaActiva
    condicion = "local = '" & empresaActiva & "' AND caja ='" & caja & "' and numero < '9999999999' and tipo = '" + TIPO + "' " & "order by tipo,numero desc limit 0,1"
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventasRubro
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
            leerUltimofoliocaja = Format(CDbl(sqlventas.response(0, 3)) + 1, "0000000000")
       Else
           leerUltimofoliocaja = "0000000001"
     End If
End Function
Public Function leerUltimofoliocajasii(TIPO, caja) As String
    
    Dim op As Integer
   
    Dim numCaja As String
    Dim numero As String
    
    
    CAMPOS(0, 0) = "foliosii"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_documento_cabeza_" + empresaActiva
    condicion = "local = '" & empresaActiva & "' AND caja ='" & caja & "' and numero < '9999999999' and tipo = '" + TIPO + "' " & "order by tipo,numero desc"
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventasRubro
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
            leerUltimofoliocajasii = Format(CDbl(sqlventas.response(0, 3)) + 1, "0000000000")
       Else
           leerUltimofoliocajasii = "0000000001"
     End If
End Function
Public Function leerfoliocaja(TIPO, caja) As String
    
    Dim op As Integer
    Dim numCaja As String
    Dim numero As String
    If TIPO = "BV" Then
    CAMPOS(0, 0) = "folioboletas"
    End If
    If TIPO = "FV" Then
    CAMPOS(0, 0) = "foliofacturas"
    End If
    If TIPO = "NF" Or TIPO = "NB" Then
    CAMPOS(0, 0) = "folionotacredito"
    End If
    
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_maestrodecajas"
    
    condicion = "local = '" & empresaActiva & "' AND numero ='" & caja & "' "
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
            leerfoliocaja = Format(CDbl(sqlventas.response(0, 3)) + 1, "0000000000")
            
    Else
           MsgBox ("CAJA NO SE ENCUENTRA GENERADA")
           
    End If
End Function
Public Function leerdiapagoCliente(ByVal codigo As String) As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "diapago"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes"
        
        condicion = "rut = '" & codigo & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerdiapagoCliente = sql.response(0, 3)
        Else
            leerdiapagoCliente = ""
        End If
    End Function
    Public Function Numero_Texto(nvalor As Long) As String
    
    Dim Mon_Esc, QueES As String
    Dim K As String
    ReDim uni(15) As String
    ReDim Dec(9) As String
    Dim Z, NUM, var As Variant
    Dim c, D, u, v, I As Integer
    Dim textnum As Long
    If Len(nvalor) = 0 Then                        'Si no se ingresa Valor se Devuelve Vacío
        textnum = "": Exit Function
    End If
    If nvalor = 0 Or nvalor > 1E+17 Then
       Mon_Esc = IIf(nvalor = 0, "CERO", "*")
    End If
    ' ------------ UNIDADES ----------------------------------
    uni(1) = "UN"
    uni(2) = "DOS"
    uni(3) = "TRES"
    uni(4) = "CUATRO"
    uni(5) = "CINCO"
    uni(6) = "SEIS"
    uni(7) = "SIETE"
    uni(8) = "OCHO"
    uni(9) = "NUEVE"
    uni(10) = "DIEZ"
    uni(11) = "ONCE"
    uni(12) = "DOCE"
    uni(13) = "TRECE"
    uni(14) = "CATORCE"
    uni(15) = "QUINCE"
    ' ------------ DECENAS ----------------------------------
    Dec(3) = "TREINTA"
    Dec(4) = "CUARENTA"
    Dec(5) = "CINCUENTA"
    Dec(6) = "SESENTA"
    Dec(7) = "SETENTA"
    Dec(8) = "OCHENTA"
    Dec(9) = "NOVENTA"
    
    NUM = String$(19 - Len(Str(Trim(nvalor))), Space(1))
    NUM = NUM + Trim(Str(nvalor))
    I = 1
    Z = ""
    
    Do While True
       K = Mid(NUM, 18 - (I * 3 - 1), 3)
    
       If K = Space(3) Then
          Exit Do
       End If
    
       c = Val(Mid(K, 1, 1))
       D = Val(Mid(K, 2, 1))
       u = Val(Mid(K, 3, 1))
       v = Val(Mid(K, 2, 2))
    
       If I > 1 Then
          If (I = 2 Or I = 4) And Val(K) > 0 Then
             Z = " MIL " + Z
          End If
          If I = 3 And Val(Mid(NUM, 7, 6)) > 0 Then
             If Val(K) = 1 Then
                Z = " MILLON " + Z
             Else
                Z = " MILLONES " + Z
             End If
          End If
          If I = 5 And Val(K) > 0 Then
             If Val(K) = 1 Then
                Z = " BILLON " + Z
             Else
                Z = " BILLONES " + Z
             End If
          End If
       End If
    
       If v > 0 Then
          Select Case v
                 Case 0 To 15
                      Z = uni(v) + Z
                 Case 0 To 19
                      Z = " DIECI" + uni(u) + Z
                 Case 20
                      Z = " VEINTE " + Z
                 Case 0 To 29
                      Z = " VEINTI" + uni(u) + Z
                 Case Else
                      If u = 0 Then
                         Z = Dec(D) + Z
                      Else
                         Z = Dec(D) + " Y " + uni(u) + Z
                      End If
          End Select
       End If
    
       If c > 0 Then
          If c = 1 Then
             If v = 0 Then
                Z = " CIEN " + Z
             Else
                Z = " CIENTO " + Z
             End If
          End If
          If c = 2 Or c = 3 Or c = 4 Or c = 6 Or c = 8 Then
             Z = uni(c) + "CIENTOS " + Z
          End If
          If c = 5 Then
             Z = " QUINIENTOS " + Z
          End If
          If c = 7 Then
             Z = " SETECIENTOS " + Z
          End If
          If c = 9 Then
             Z = " NOVECIENTOS " + Z
          End If
       End If
    
       I = I + 1
    Loop
    
    Mon_Esc = Trim(Z)
    ' CAMBIA "UNO MIL ..." POR "MIL..."
    If Mid(Mon_Esc, 1, 7) = "UN MIL " Then
        Mon_Esc = "MIL " + Trim(Mid(Mon_Esc, 7, Len(Mon_Esc)))
    End If
    Numero_Texto = Mon_Esc + " PESOS" + QueES
End Function
Sub leertiposdeclientes(ByRef frm As Form, ByRef grilla As Grid)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Double
    
        Set csql.ActiveConnection = ventas
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM sv_tiposdeclientes "
        csql.sql = csql.sql + "order by codigo "
        csql.Execute
        linea = 0
        grilla.Rows = csql.RowsAffected + 1
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
               linea = linea + 1
                grilla.Cell(linea, 1).text = resultados(0)
                grilla.Cell(linea, 2).text = resultados(1)
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
           
        End If
        
End Sub

Public Function leertipocli(ByVal codigo As String) As String
        Dim op As Integer
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 2) = "sv_tiposdeclientes"

        condicion = "codigo = '" & codigo & "'"

        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
        leertipocli = sql.response(0, 3)
        
        Else
        leertipocli = ""
        End If
End Function
Public Function escredito(ByVal codigo As String) As Boolean
        Dim op As Integer
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "credito"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 2) = "sv_tiposdeclientes"

        condicion = "codigo = '" & codigo & "' and credito='1' "

        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
        
        escredito = True
        
        Else
        escredito = False
        End If
End Function

    
     Public Function LEErultimoeventoFECHA(rut, FECHAMORA) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT fecha  "
        csql.sql = csql.sql & "FROM sv_cobranza_gestion "
        csql.sql = csql.sql & "WHERE rut='" & rut & "' and fecha >='" & Format(FECHAMORA, "yyyy-mm-dd") & "' "
        csql.sql = csql.sql & "order by fecha desc  limit 0,1 "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
        LEErultimoeventoFECHA = resultado(0)
        Else
        LEErultimoeventoFECHA = ""
        End If
    End Function
 Public Function leernombreevento(codigo) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
      
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT descripcion  "
        csql.sql = csql.sql & "FROM sv_cobranza_gastos "
        csql.sql = csql.sql & "WHERE codigo='" & codigo & "' "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
        leernombreevento = "  " & resultado(0)
        Else
        leernombreevento = ""
        End If
End Function
 Public Function leernombreempresacontable(codigo) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
      
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT nombre  "
        csql.sql = csql.sql & "FROM " & clientesistema & "conta.maestroempresas "
        csql.sql = csql.sql & "WHERE codigoempresa='" & codigo & "' "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
        leernombreempresacontable = "  " & resultado(0)
        Else
        leernombreempresacontable = ""
        End If
End Function


    
    
Public Function LEErBODEGADESPACHO(codigo) As String


        Dim csql As rdoQuery
        Dim resultado As rdoResultset
      
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT nombre  "
        csql.sql = csql.sql & "FROM sv_tipo_despacho "
        csql.sql = csql.sql & "WHERE codigo='" + codigo + "' and local='" + empresaActiva + "' "
        
        csql.Execute
        
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
        
        LEErBODEGADESPACHO = resultado(0)
        Else
        LEErBODEGADESPACHO = ""
        End If
    End Function
Public Function motoractivo(motor) As Boolean


        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "SELECT motor "
        csql.sql = csql.sql & "FROM sv_motoresactivos "
        csql.sql = csql.sql & "WHERE motor='" + motor + "' and local='" + empresaActiva + "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
        
        motoractivo = resultado(0)
        Else
        motoractivo = ""
        End If
    End Function

Sub sincronizadatos(ByVal cadena2 As String, ByRef coneccion As rdoConnection)

        Dim cadena As String
        Dim resultados2 As rdoResultset
        Dim cSql2 As New rdoQuery
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        Dim columnas As String
        Dim basededatos As String
        
        Dim registros As String
        Dim I As Integer
        Dim j As Integer
       If servidor_principal <> servidor_ventas Then
        For I = 1 To 40
            If Mid(coneccion.Connect, I, 1) = ";" Then
                j = I
                basededatos = Mid(coneccion.Connect, 10, j - 10)
            End If
        Next I
        cadena2 = Replace(cadena2, "'", Chr(126))
        Set cSql2.ActiveConnection = conexion
                        cadena = "INSERT INTO " + cliente_sql + "sincroniza.sincronizador ("
                        cadena = cadena + "servidor,consulta,basedatos) VALUES ("
                        cadena = cadena & "'" & servidor_principal & "','" + cadena2 + "','" + basededatos + "')"
                        'VALORES ASIGNADOS A CADA CAMPO.
                        cSql2.sql = cadena
                        
                        cSql2.Execute
                    
    End If
    End Sub

Public Function leebanco(ByVal codigo As String) As String

        Dim op As Integer
        
        
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 2) = baseteso + ".maestrobancos"

        condicion = "codigobanco = '" & codigo & "' "

        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
        
        leebanco = sqlventas.response(0, 3)
        
        Else
        leebanco = ""
        End If
End Function

Public Function leernombrecuenta(CUENTA, Banco) As String


        Dim op As Integer
        
        
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 2) = baseteso + ".maestrocuenta"
    
        condicion = "cuenta = '" & CUENTA & "' and banco='" + Banco + "' "

        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
        
        leernombrecuenta = sqlventas.response(0, 3)
        Else
        leernombrecuenta = ""
        End If
End Function
Public Function leeplaza(ByVal codigo As String) As String
        Dim op As Integer
        
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 2) = baseteso + ".maestroplazas"
        condicion = "codigo = '" & codigo & "' "
        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
        leeplaza = sqlventas.response(0, 3)
        Else
        leeplaza = ""
        End If
End Function

Public Sub grabarcuenta(CUENTA, Banco, PLAZA, nombre, rut, fono, cajera)


        Dim op As Integer
        
        
        CAMPOS(0, 0) = "cuenta"
        CAMPOS(1, 0) = "banco"
        CAMPOS(2, 0) = "plaza"
        CAMPOS(3, 0) = "nombre"
        CAMPOS(4, 0) = "rut"
        CAMPOS(5, 0) = "fono"
        CAMPOS(6, 0) = "cajera"
        CAMPOS(7, 0) = ""
        CAMPOS(0, 1) = CUENTA
        CAMPOS(1, 1) = Banco
        CAMPOS(2, 1) = PLAZA
        CAMPOS(3, 1) = nombre
        CAMPOS(4, 1) = rut
        CAMPOS(5, 1) = fono
        CAMPOS(6, 1) = cajera
        
        
        
        CAMPOS(0, 2) = baseteso + ".maestrocuenta"
    
        condicion = ""

        op = 2
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        
        
End Sub

Public Function leeautorizacionterceraedad(codigo) As Boolean
       
    Dim CAMPOS(3, 3) As String
       
    
    Dim op As Integer
    
    CAMPOS(0, 0) = "terceraedad"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_maestroclientes"
    condicion = "rut= '" + codigo + "' "
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
            leeautorizacionterceraedad = sqlventas.response(0, 3)
    End If
End Function
Public Function leeautorizacion(codigo) As Boolean
       
    Dim CAMPOS(3, 3) As String
       
    
    Dim op As Integer
    
    CAMPOS(0, 0) = "codigo"
    CAMPOS(1, 0) = "nombre"
    CAMPOS(2, 0) = "gerencia"
    CAMPOS(3, 0) = ""
    CAMPOS(0, 2) = "sv_permisos_caja"
    condicion = "codigo = '" + codigo + "' and local='" & empresaActiva & "' "
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    descuentogerencia = False
    
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
            leeautorizacion = True
            claveautorizador = sqlventas.response(1, 3)
            tarjetaautorizador = sqlventas.response(0, 3)
            descuentogerencia = sqlventas.response(2, 3)
        Else
            leeautorizacion = False
        
    End If
End Function
Public Function leecupocuenta(CUENTA, Banco) As Double


        Dim op As Integer
        
        
        CAMPOS(0, 0) = "cupo"
        CAMPOS(1, 0) = "bloqueo"
        CAMPOS(2, 0) = ""
        CAMPOS(0, 2) = baseteso + ".maestrocuenta"
    
        condicion = "cuenta = '" & CUENTA & "'and banco = '" + Banco + "'"

        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
        
        leecupocuenta = sqlventas.response(0, 3)
        If sqlventas.response(1, 3) = "S" Then
        leecupocuenta = -1
        End If
        
        Else
        leecupocuenta = 0
        End If
End Function
Public Function leerjefaautorizacion(codigo) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventas
csql.sql = "select nombre from sv_permisos_caja where "
csql.sql = csql.sql & "local='" & empresaActiva & "' and codigo='" & codigo & "'"
csql.Execute

If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset

leerjefaautorizacion = resultados(0)

End If
csql.Close
Set csql = Nothing
Set resultados = Nothing

End Function
Public Function leerNombreTipoDespacho(ByVal codigo As String) As String
        
        Dim op As Integer
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 2) = "sv_tipo_despacho"
        condicion = "codigo = '" & codigo & "' and local='" & empresaActiva & "' "
        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
            leerNombreTipoDespacho = sqlventas.response(0, 3)
        Else
            leerNombreTipoDespacho = ""
        End If
    End Function
'Public Sub leercredito(TIPO, caja, Boleta, rut)
'
'    Dim op As Integer
'    CAMPOS(0, 0) = "cantidadcuotas"
'    CAMPOS(1, 0) = "rut"
'    CAMPOS(2, 0) = "vencimientoactual"
'    CAMPOS(3, 0) = "montocredito"
'    CAMPOS(4, 0) = "montocuota"
'    CAMPOS(5, 0) = ""
'    CAMPOS(0, 2) = "sv_cuotas_detalle"
'    condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND numero = '" & Boleta & "' and rut='" + rut + "' order by vencimientoactual "
'    op = 5
'    sqlventas.response = CAMPOS
'    Set sqlventas.conexion = CRETMP
'    Call sqlventas.sqlventas(op, condicion)
'    cantidadCUOTAS = ""
'    rutcredito = ""
'    primervencimiento = ""
'    montocredito = ""
'    montocuotas = ""
'    If sqlventas.Status = 0 Then
'    cantidadCUOTAS = sqlventas.response(0, 3)
'    rutcredito = sqlventas.response(1, 3)
'    primervencimiento = Format(sqlventas.response(2, 3), "dd-mm-yyyy")
'    montocredito = sqlventas.response(3, 3)
'    montocuotas = sqlventas.response(4, 3)
'    Else
'
'    End If
'End Sub

Public Sub leercredito(TIPO, caja, boleta, rut)
    
    Dim op As Integer
    CAMPOS(0, 0) = "cantidadcuotas"
    CAMPOS(1, 0) = "rut"
    CAMPOS(2, 0) = "vencimientoactual"
    CAMPOS(3, 0) = "montocredito"
    CAMPOS(4, 0) = "montocuota"
    CAMPOS(5, 0) = ""
    CAMPOS(0, 2) = "sv_cuotas_detalle"
    condicion = "local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND numero = '" & boleta & "' and rut='" + rut + "' order by vencimientoactual "
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    cantidadcuotas = ""
    rutcredito = ""
    primervencimiento = ""
    montocredito2 = ""
    montocuotas = ""
    If sqlventas.Status = 0 Then
    cantidadcuotas = sqlventas.response(0, 3)
    rutcredito = sqlventas.response(1, 3)
    primervencimiento = Format(sqlventas.response(2, 3), "dd-mm-yyyy")
    montocredito2 = sqlventas.response(3, 3)
    montocuotas = sqlventas.response(4, 3)
    Else
        
    End If
End Sub

Public Sub grabarcuotacobranza(rut, fecha, evento, glosaevento)
    Dim csql As New rdoQuery
    
    Set csql.ActiveConnection = ventas
    csql.sql = "insert into sv_cuotas_detalle set local='07',tipo='GC' "
    csql.sql = csql.sql & ",numero='" & Format(leerultimogastocobranza, "0000000000") & "',rut='" & rut & "',numerocuota='1',vencimientooriginal='" & fecha & "' "
    csql.sql = csql.sql & ",vencimientoactual='" & fecha & "',montocuota='" & leercargo(evento) & "',abono='0',fechacompra='" & fecha & "' "
    csql.sql = csql.sql & ",cantidadcuotas='1',glosacompra='" & glosaevento & "',capitalcuota='" & leercargo(evento) & "', montocredito='" & leercargo(evento) & "' "
    csql.Execute
    
    csql.Close
    Set csql = Nothing
    
    
End Sub
Public Function leerultimogastocobranza() As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
      
    Set csql.ActiveConnection = ventas
    csql.sql = "select ifnull(max(numero)+1,'0000000001') from "
    csql.sql = csql.sql & "sv_cuotas_detalle where tipo='GC' "
    csql.Execute
    leerultimogastocobranza = "0000000001"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerultimogastocobranza = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
End Function
Public Function leercargo(evento) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventas
csql.sql = "select monto from sv_cobranza_gastos "
csql.sql = csql.sql & "where codigo='" & evento & "' "
csql.Execute
leercargo = "0"
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leercargo = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function
Public Function leerdescuento(rut) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventas
csql.sql = "select descuento from sv_maestroclientes "
csql.sql = csql.sql & "where rut='" & rut & "' "
csql.Execute
leerdescuento = "0"
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerdescuento = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function
 
Public Sub revisarmenus(ByRef frm As Form)
    Dim ctlControl As Object
    Dim cad As String
    Dim cadindex As String
    Dim tipovariable As String
    
    On Error Resume Next
    For Each ctlControl In frm.Controls
       
           cad = ctlControl.Name
           cadindex = ctlControl.Index
           tipovariable = TypeName(ctlControl)
'           List1.AddItem (cad + " " + cadindex + " " + tipovariable)
           
            If tipovariable = "Menu" And cadindex <> "99" Then
            ctlControl.Caption = Replace(ctlControl.Caption, "&", "")
            If existepermiso(usuarioSistema, ctlControl.Caption) = True Then
            ctlControl.Enabled = True
            Else
            ctlControl.Enabled = False
            
            End If
            End If
       cadindex = "0"
       ' DoEvents
    Next ctlControl
End Sub
Private Function existepermiso(usuario, programa) As Boolean

    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "SELECT todas,ingresa "
        cSql2.sql = cSql2.sql + "FROM segu_permisos "
        cSql2.sql = cSql2.sql + "where usuario='" + usuario + "' and programa='" + programa + "'"
        cSql2.Execute
        existepermiso = False
      
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        If resultados2(0) = 1 Or resultados2(1) = 1 Then
            existepermiso = True
        End If
End If
End Function
Public Function RetornaValor(ByVal valretorno As String, _
                                tabla As String, _
                                condicion As String, _
                                conactiva As rdoConnection) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Set csql.ActiveConnection = ventas
    csql.sql = "SELECT " & valretorno
    csql.sql = csql.sql & " FROM " & tabla
    csql.sql = csql.sql & " WHERE " & condicion
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        If Not resultados.EOF Then
            RetornaValor = resultados(0).Value
        End If
        resultados.Close
        Set resultados = Nothing
    Else
        RetornaValor = "0"
    End If
    csql.Close
End Function

Public Function numerodeprestamo(rut) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventas
csql.sql = "select ifnull(max(numeroprestamo),0) from sv_prestamo "

csql.sql = csql.sql & "where rut='" & rut & "' "
csql.Execute
numerodeprestamo = 1
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    numerodeprestamo = resultados(0) + 1
End If
csql.Close
Set csql = Nothing
End Function

Public Function leerfoliofiscal(loc, TIPO, numero, fecha, caja) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select foliosii from " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc
csql.sql = csql.sql & " where tipo='" + TIPO + "' and numero='" + numero + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "' and caja='" + caja + "' "
csql.Execute
leerfoliofiscal = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerfoliofiscal = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function


Public Function leerUsuario(basedatos, tabla, evento, CONSULTA) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "select usuario from " + clientesistema + "auditoria.auditoriaventas "
csql.sql = csql.sql & " where basedatos='" + basedatos + "' and tabla='" + tabla + "' and evento='" + evento + "' and datosoriginales like '" + CONSULTA + "' limit 0,1 "
csql.Execute
leerUsuario = ""
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerUsuario = resultados(0)
End If
csql.Close
Set csql = Nothing
End Function


Public Sub consultaReplicas(CONSULTA, base)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim cSql2 As New rdoQuery


    Set csql.ActiveConnection = gestion
    
    csql.sql = "select servidorventas,servidorprincipal,codigo from g_maestroempresas "
    csql.sql = csql.sql & "where rubro='" & rubro & "' and servidorventas<>servidorprincipal "
    csql.Execute
        
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
            While Not resultados.EOF
            
        'If resultados(2) <> empresaActiva Then
            Call sincronizadatos2(CONSULTA, gestionRubro, resultados(0), base)
        'End If
        
        resultados.MoveNext
        Wend
        
    End If
    
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
End Sub
Sub sincronizadatos2(ByVal cadena2 As String, ByRef coneccion As rdoConnection, Servidor, base)

        Dim cadena As String
        Dim resultados2 As rdoResultset
        Dim cSql2 As New rdoQuery
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        Dim columnas As String
        Dim basededatos As String
        
        Dim registros As String
        Dim I As Integer
        Dim j As Integer
       
        
        basededatos = base
         
        cadena2 = Replace(cadena2, "'", Chr(126))
        Set cSql2.ActiveConnection = conexion
                        cadena = "INSERT INTO " + cliente_sql + "sincroniza.sincronizador ("
                        cadena = cadena + "servidor,consulta,basedatos,fechacreacion,horacreacion) VALUES ("
                        cadena = cadena & "'" & Servidor & "','" + cadena2 + "','" + basededatos + "','" & Format(Date, "yyyy-mm-dd") & "','" & Time & "')"
                        'VALORES ASIGNADOS A CADA CAMPO.
                        cSql2.sql = cadena
                        cSql2.Execute
   
    End Sub
Public Function leerultimalineaharina(numero) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select ifnull(max(linea+1),'1') from sv_valesharinas "
    csql.sql = csql.sql & "where numero='" & numero & "' "
    csql.Execute
    leerultimalineaharina = "1"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerultimalineaharina = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
End Function
Public Function productomarcado(codigo) As Boolean
    Dim csql As New rdoQuery
    
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "select harina from " & clientesistema & "gestion" & empresaActiva & ".r_maestroproductos_fijo_" & rubro & " "
    csql.sql = csql.sql & "where codigobarra='" & codigo & "' and harina='1' "
    csql.Execute
    productomarcado = False
    If csql.RowsAffected > 0 Then
    productomarcado = True
    End If
    
End Function

Public Function Despachosharinas(fecha, fechadespacho, op) As Long
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "select count(numero) from " & clientesistema & "ventas" & empresaActiva & ".sv_valesharinas"

If op = 1 Then csql.sql = csql.sql & " where fecha = '" & fecha & "' and fechaentrega = '" & fechadespacho & "'"
If op = 2 Then csql.sql = csql.sql & " where fechaentrega = '" & fechadespacho & "'"
If op = 3 Then csql.sql = csql.sql & " where fechaentrega = '" & fecha & "' and fecha <> '" & fecha & "'"
    
    csql.Execute

    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    Despachosharinas = resultados(0)
    End If
    
End Function
Public Function esvip(rut) As Boolean
 Dim csql As New rdoQuery
 
 Set csql.ActiveConnection = ventas
 csql.sql = "select * from " & baseteso & ".clientes_vip "
 csql.sql = csql.sql & "where rut='" & rut & "' "
 csql.Execute
 esvip = False
 
 If csql.RowsAffected > 0 Then
    esvip = True
 End If
 csql.Close
 Set csql = Nothing
 
End Function

Public Function leeultimacompra(codigo, loca) As Double
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestionRubro
        csql.sql = "SELECT precio,fecha FROM l_movimientos_detalle_" & loca & " WHERE  tipo ='OC' AND codigo='" & codigo & "' ORDER BY fecha DESC LIMIT 1"
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            leeultimacompra = resultados(0)
            If loca = "20" Then
            fechaultimacompra2 = resultados(1)
            Else
            fechaultimacompra1 = resultados(1)
            
            End If
            
            resultados.Close
            Set resultados = Nothing
        Else
        leeultimacompra = 0
        fechaultimacompra1 = "2000-01-01"
        End If
 If leeultimacompra = 0 Then leeultimacompra = 1
 
End Function
    
Public Sub actualizafoliosii(TIPO, numero, caja, fecha, nuevofolio)
    Dim csql As rdoQuery
    Dim folio As String
    Dim op As Integer
    Dim folioantiguo As String
    Dim condicion As String
    Dim resul As rdoResultset
    Dim rutcre As String
    nuevofolio = Format(nuevofolio, "0000000000")
    Set csql = New rdoQuery
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "SELECT foliosii,rut from " & cliente_sql & "ventas" & LOCAL_PROCESO & ".sv_otros_documento_cabeza_" & LOCAL_PROCESO
    csql.sql = csql.sql & " where tipo='" + TIPO + "' AND local='" + LOCAL_PROCESO + "' and  numero='" & numero & "' and caja='" & caja & "'  and  fecha='" & Format(fecha, "yyyy-mm-dd") & "' limit 0,1 "
    csql.Execute
    Set resul = csql.OpenResultset
    folioantiguo = resul(0)
    rutcre = resul(1)
    'Set cSql = Nothing
    Set csql = New rdoQuery
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "UPDATE " & cliente_sql & "ventas" & LOCAL_PROCESO & ".sv_otros_documento_cabeza_" & LOCAL_PROCESO & " SET foliosii ='" & nuevofolio & "' where numero='" & numero & "' and caja='" & caja & "' and fecha='" & fecha & "' and local='" + LOCAL_PROCESO + "' and tipo='" + TIPO + "' "
    csql.Execute
    Call sincronizadatos(csql.sql, ventasRubro)
    
    Set csql = New rdoQuery
    Set csql.ActiveConnection = ventasRubro
   ' csql.sql = "UPDATE sv_valesharinas SET numerodocumento ='" & nuevofolio & "' where numerodocumento='" & folioantiguo & "' and caja='" & caja & "' and fecha='" & fechasistema & "'  and tipodocumento='" + tipo + "' "
   ' csql.Execute
    'Call sincronizadatos(csql.sql, ventasRubro)
    
    
        
'   Call Conectarservercredito(serverprincipal, baseVentas, usuario, password)
'
'   If creditoactivo = True Then
'    Set csql.ActiveConnection = CRETMP
'    csql.sql = "UPDATE sv_cuotas_detalle SET numero ='" & nuevofolio & "' where tipo = '" + TIPO + "' AND numero = '" & folioantiguo & "' AND local= '" & empresaActiva & "' and rut='" + rutcre + "'"
'    csql.Execute
'    Call sincronizadatos(csql.sql, ventas)
'    csql.Close
'    Set csql = Nothing
'    End If
'
End Sub


Public Sub sincronizarFechaHora()
        Dim fecha As String
        Dim HORA As String
        Dim dia41 As Double
        Dim mes41 As Double
        Dim ano41 As Double
        Dim time41 As String
        Dim hora41 As Double
        Dim fecha41 As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
    If VERIFICAPING(Servidor) = True Then
        Set csql = New rdoQuery
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT DATE_FORMAT(CURRENT_TIMESTAMP(),'%d-%m-%Y') AS fecha, TIME_FORMAT(CURRENT_TIMESTAMP(),'%T') AS hora "
        'csql.sql = "select '05-02-2013' as fecha,'00:00:01' as hora"
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
            fecha = resultado("fecha")
            HORA = resultado("hora")
            'Para la fecha:
            Date = DateValue(fecha)
            'Para la hora:
            Time = TimeValue(HORA)
            fechasistema = Format(Date, "yyyy-mm-dd")
            Rem fechasistema = "2008-10-01"
             Rem Time = "01:00:00"
        End If
        csql.Close
        Set csql = Nothing
dia41 = CDbl(Format(fechasistema, "dd"))
   mes41 = CDbl(Format(fechasistema, "mm"))
   ano41 = CDbl(Format(fechasistema, "yyyy"))
   
If (empresaActiva = "41") Or (empresaActiva = "00") Or (empresaActiva = "17") Then
        hora41 = Mid(Time$, 1, 2)
    hora41 = hora41 - 5
    If hora41 < 0 Then
        hora41 = 24 + hora41
        dia41 = dia41 - 1
        
   
   
           If dia41 = 0 Then
                    mes41 = mes41 - 1
                    If mes41 = 1 Then dia41 = 31
                    If mes41 = 2 Then dia41 = 28
                    If mes41 = 3 Then dia41 = 31
                    If mes41 = 4 Then dia41 = 30
                    If mes41 = 5 Then dia41 = 31
                    If mes41 = 6 Then dia41 = 30
                    If mes41 = 7 Then dia41 = 31
                    If mes41 = 8 Then dia41 = 31
                    If mes41 = 9 Then dia41 = 30
                    If mes41 = 10 Then dia41 = 31
                    If mes41 = 11 Then dia41 = 30
                    If mes41 = 12 Then dia41 = 31
                    If mes41 = 0 Then
                        ano41 = ano41 - 1
                        mes41 = 12
                        dia41 = 31
                    End If
            End If
    End If
   fecha41 = Format(ano41, "0000") + "-" + Format(mes41, "00") + "-" + Format(dia41, "00")
   time41 = Format(hora41, "00") & Mid(Time$, 3, 6)
   Time$ = time41
   fechasistema = fecha41
   
   
End If
           
           'MsgBox fechasistema
        
  
  End If
End Sub
Public Function generacadena(response, Opcion, condicion) As String
Dim cadena As String
Dim response1 As String
Dim I As Double


Select Case Opcion

                Case 2:    '<<<<<   INSERTA   >>>>>>


                    cadena = "INSERT INTO " & response(0, 2) & " ("
                    response1 = ""
                    'NOMBRE DE response.
                    I = 0
                    While response(I, 0) <> ""
                        cadena = cadena & response(I, 0) + ","
                        response1 = response1 & "[" & response(I, 0) & "]"
                        I = I + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & ") VALUES ("
                    
                    'VALORES ASIGNADOS A CADA CAMPO.
                    I = 0
                    While response(I, 0) <> ""
                        cadena = cadena & "'" & response(I, 1) & "',"
                        I = I + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & ") "
                    cadena = cadena & "ON DUPLICATE KEY UPDATE " & response(0, 0) & " = " & response(0, 0)
                    generacadena = cadena
                    
                Case 3:    '<<<<<   ACTUALIZA   >>>>>>
                    
        
                    cadena = "UPDATE " & response(0, 2) & " SET "
                    I = 0
                    response1 = ""
                    response2 = ""
                    While response(I, 0) <> ""
                        cadena = cadena & response(I, 0) & "= '" & response(I, 1) & "',"
                        response1 = response1 & "[" & response(I, 0) & "]"
                        response2 = response2 & "[" & response(I, 1) & "]"
                        I = I + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " WHERE " & condicion
                    generacadena = cadena
                    
                    
                Case 4:    '<<<<<   ELIMINA   >>>>>>
                    If audit = True Then
                        Call Auditoria(Opcion, condicion)
                    End If
                    cadena = "DELETE FROM " & response(0, 2) & " WHERE " & condicion
                    generacadena = cadena
                     
        
                Case 5:    '<<<<<   LEE   >>>>>>
                    cadena = "SELECT "
                    I = 0
                    While response(I, 0) <> ""
                        cadena = cadena & response(I, 0) & ","
                        I = I + 1
                    Wend
                    cadena = Left(cadena, Len(cadena) - 1)
                    cadena = cadena & " FROM " & response(0, 2) & " WHERE " & condicion
                    generacadena = cadena
 End Select
                    
                    
 End Function
 Public Sub leerdatos_Certificado(ByRef usuario As String, ByRef password As String)
     Dim csql As New rdoQuery
     Dim resultados As rdoResultset
     
     Set csql.ActiveConnection = temporal
     csql.sql = "select licencia,certificado from adminerp_inicio.licencias "
     csql.Execute
     If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         usuario = resultados(0)
         password = resultados(1)
         usuario = leer_certificado_digital(usuario, "leyendo_validez_certificado_firma_sii")
         password = leer_certificado_digital(password, "leyendo_validez_certificado_firma_sii")
         
         Call Conectar(Servidor, "mysql", usuario, password)
'         Call Conectartemporal(servidor, "mysql", usuario, password)
         
     Else
        MsgBox ("NO EXISTE CONFIGURACION NI LICENCIA PARA ESTE SOFTWARE")
        End
     End If
no:
End Sub
Public Function leer_certificado_digital(ByVal s As String, ByVal p As String) As String
     
    Dim I As Integer, R As String
    Dim C1 As Integer, C2 As Integer
    R = ""
    If Len(p) > 0 Then
        If InStr(1, s, "FLAG", vbBinaryCompare) <> 0 Then
            s = Replace(s, "FLAG", Chr(13))
        End If
        For I = 1 To Len(s)
            C1 = Asc(Mid(s, I, 1))
            If I > Len(p) Then
                C2 = Asc(Mid(p, I Mod Len(p) + 1, 1))
            Else
                C2 = Asc(Mid(p, I, 1))
            End If
            C1 = C1 - C2 - 64
            If Sgn(C1) = -1 Then C1 = 256 + C1
                R = R + Chr(C1)
        Next I
    Else
        R = s
    End If
    leer_certificado_digital = R
End Function
