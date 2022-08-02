Attribute VB_Name = "Funciones_Total"
'Sub Conectar_BD()
'RUTINA PARA CONECTAR A LA BASE DE DATOS
'    servidor = "servidor"
'    usuario = "root"
'    Password = "123"
'    basedatos = "conta01"
'    On Error GoTo controlerror
'        Call Conectar(servidor, basedatos, usuario, Password)
'    Exit Sub
'controlerror:

'    Resume Next
'End Sub

Sub Conecta_Cuentas_Corrientes_2()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With ingre02
        .cuentascorrientes.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=conta01"
    End With
End Sub

Sub Desconecta_Cuentas_Corrientes()
'DESCONECTA EL CONTROL DATA DE LA BASE DE DATOS.
    With ingre02
        If .cuentascorrientes.Recordset.State = adStateClosed Then

        Else
            .cuentascorrientes.Recordset.Close
        End If
    End With
End Sub

Sub Limpia_Formulario_2()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With ingre02
        For i = 0 To 11
            .dato(i).Text = ""
        Next i
        'For i = 0 To 1
        '    .lbl(i).Caption = ""
        'Next i
        .dato(0).Enabled = True
    End With
End Sub

Sub Activa_Controles(estado As Boolean, accion As String)
'ACTIVA O DESACTIVA LOS CONTROLES DEL FORMULARIO.
    Dim i As Integer
    
    Select Case accion
    
        Case "CARGAR"
                With ingre02
                    For i = 0 To 12
                        .dato(i).Enabled = estado
                    Next i
                    .flash_opciones.Visible = True
                End With
                
        Case "MODIFICA"
                With ingre02
                    For i = 2 To 12
                        .dato(i).Enabled = estado
                    Next i
                    .flash_opciones.Visible = True
                    .dato(2).SetFocus
            End With
        
    End Select
End Sub

Sub Lista_Cuentas_Corrientes()
'MUESTRA LAS CUENTAS CORRIENTES EN LA GRILLA.
    With ingre02
        .Grilla.Clear
        .cuentascorrientes.RecordSource = "SELECT rut,nombre FROM cuentascorrientes ORDER BY rut"
        .cuentascorrientes.Refresh
    End With
End Sub
'
Sub Carga_Cuenta(codigo_cuenta As String)
'RUTINA QUE CARGA UNA CUENTA INGRESADA SI ESTA EXISTE.
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset

    With maestro01
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT descrpcion,unitariomayor,codigoseccion,codigodepto,codigolinea,impuesto,pcostoiva,pvmayorista1,pventadetalle,stockcritico,descuento,datoextra,ubicacion"
        cSql.SQL = cSql.SQL + "FROM maestroproductos"
        cSql.SQL = cSql.SQL + "WHERE codigoproducto='" & codigo_cuenta & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            .dato(0).Text = codigo_cuenta
            .dato(1).Text = resultados(0)
            .dato(2).Text = resultados(1)
            .dato(3).Text = resultados(2)
            .dato(4).Text = resultados(3)
            .dato(5).Text = resultados(4)
            .dato(6).Text = resultados(5)
            .dato(7).Text = resultados(6)
            .dato(8).Text = resultados(7)
            .dato(9).Text = resultados(8)
            .dato(10).Text = resultados(9)
            .dato(11).Text = resultados(10)
            .dato(12).Text = resultados(11)
            .dato(13).Text = resultados(13)
            resultados.Close
            Set resultados = Nothing
        End If
    End With
End Sub

Sub Carga_Cuenta_Siguiente(codigo_cuenta As String)
'RUTINA QUE CARGA UNA CUENTA INGRESADA SI ESTA EXISTE.
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset

    With ingre01
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor "
        cSql.SQL = cSql.SQL + "WHERE codigo > '" & codigo_cuenta & "' "
        cSql.SQL = cSql.SQL + "ORDER BY codigo"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            Call Carga_Cuenta(resultados(0))
'            .codigo.Text = Mid(resultados(0), 1, 2) & "-" & Mid(resultados(0), 3, 2) & "-" & Mid(resultados(0), 5, 4)
'            .dato(0).Text = resultados(1)
'            .dato(1).Text = resultados(2)
'            .dato(2).Text = resultados(3)
'            .dato(3).Text = resultados(4)
'            .dato(4).Text = resultados(5)
            resultados.Close
            Set resultados = Nothing
        End If
    End With
End Sub

Sub Carga_Cuenta_Anterior(codigo_cuenta As String)
'RUTINA QUE CARGA UNA CUENTA INGRESADA SI ESTA EXISTE.
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset

    With ingre01
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor "
        cSql.SQL = cSql.SQL + "WHERE codigo < '" & codigo_cuenta & "' "
        cSql.SQL = cSql.SQL + "ORDER BY codigo DESC"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            Call Carga_Cuenta(resultados(0))
'            .codigo.Text = Mid(resultados(0), 1, 2) & "-" & Mid(resultados(0), 3, 2) & "-" & Mid(resultados(0), 5, 4)
'            .dato(0).Text = resultados(1)
'            .dato(1).Text = resultados(2)
'            .dato(2).Text = resultados(3)
'            .dato(3).Text = resultados(4)
'            .dato(4).Text = resultados(5)
            resultados.Close
            Set resultados = Nothing
        End If
    End With
End Sub

Sub Manejo_Datos()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    Dim op As Integer
    Dim condicion As String
    
    With maestro01
        codigo_cuenta = .dato(0).Text
        .mp.RecordSource = "SELECT codigoproducto FROM maestroproductos WHERE codigoproducto='" & codigo_cuenta & "'"
        .mp.Refresh
        If .mp.Recordset.RecordCount > 0 Then 'YA EXISTE LA CUENTA
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = .dato(0).Tag 'CODIGO
            campos(1, 0) = .dato(1).Tag 'DESCRIPCION
            campos(2, 0) = .dato(2).Tag 'UNITARIOMAYOR
            campos(3, 0) = .dato(3).Tag 'CODIGOLOCAL
            campos(4, 0) = .dato(4).Tag 'CODIGOSECION
            campos(5, 0) = .dato(5).Tag 'CIUDAD
            campos(6, 0) = .dato(6).Tag 'GIRO
            campos(7, 0) = .dato(7).Tag 'FONO
            campos(8, 0) = .dato(8).Tag 'FAX
            campos(9, 0) = .dato(9).Tag 'CELULAR
            campos(10, 0) = .dato(10).Tag 'EMAIL
            campos(11, 0) = .dato(11).Tag 'CONTACTO
            campos(12, 0) = .dato(12).Tag 'DEST. CHEQUE
            campos(13, 0) = .dato(13).Tag 'DEST. CHEQUE
            
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).Text 'TIPO
            campos(1, 1) = .dato(1).Text 'RUT
            campos(2, 1) = .dato(2).Text 'NOMBRE
            campos(3, 1) = .dato(3).Text 'DIRECCION
            campos(4, 1) = .dato(4).Text 'COMUNA
            campos(5, 1) = .dato(5).Text 'CIUDAD
            campos(6, 1) = .dato(6).Text 'GIRO
            campos(7, 1) = .dato(7).Text 'FONO
            campos(8, 1) = .dato(8).Text 'FAX
            campos(9, 1) = .dato(9).Text 'CELULAR
            campos(10, 1) = .dato(10).Text 'EMAIL
            campos(11, 1) = .dato(11).Text 'CONTACTO
            campos(12, 1) = .dato(12).Text 'DEST. CHEQUE
            campos(13, 0) = .dato(13).Text 'DEST. CHEQUE
            
            'NOMBRE DE TABLA.
            campos(0, 2) = "maestroproductos"
            'CONDICION DE LA CONSULTA.
            condicion = "codigoproducto=" + "'" + codigo_cuenta + "'"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 2
            Call SQLUTIL(op, campos, condicion)
            
        Else                                            'CREA UNA CUENTA NUEVA
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = .dato(0).Tag 'TIPO
            campos(1, 0) = .dato(1).Tag 'RUT
            campos(2, 0) = .dato(2).Tag 'NOMBRE
            campos(3, 0) = .dato(3).Tag 'DIRECCION
            campos(4, 0) = .dato(4).Tag 'COMUNA
            campos(5, 0) = .dato(5).Tag 'CIUDAD
            campos(6, 0) = .dato(6).Tag 'GIRO
            campos(7, 0) = .dato(7).Tag 'FONO
            campos(8, 0) = .dato(8).Tag 'FAX
            campos(9, 0) = .dato(9).Tag 'CELULAR
            campos(10, 0) = .dato(10).Tag 'EMAIL
            campos(11, 0) = .dato(11).Tag 'CONTACTO
            campos(12, 0) = .dato(12).Tag 'DEST. CHEQUE
            campos(13, 0) = .dato(13).Tag 'DEST. CHEQUE
            
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).Text 'TIPO
            campos(1, 1) = .dato(1).Text 'RUT
            campos(2, 1) = .dato(2).Text 'NOMBRE
            campos(3, 1) = .dato(3).Text 'DIRECCION
            campos(4, 1) = .dato(4).Text 'COMUNA
            campos(5, 1) = .dato(5).Text 'CIUDAD
            campos(6, 1) = .dato(6).Text 'GIRO
            campos(7, 1) = .dato(7).Text 'FONO
            campos(8, 1) = .dato(8).Text 'FAX
            campos(9, 1) = .dato(9).Text 'CELULAR
            campos(10, 1) = .dato(10).Text 'EMAIL
            campos(11, 1) = .dato(11).Text 'CONTACTO
            campos(12, 1) = .dato(12).Text 'DEST. CHEQUE
            campos(13, 1) = .dato(13).Text 'DEST. CHEQUE
            
            'NOMBRE DE TABLA
            campos(0, 2) = "maestroproductos"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 1
            Call SQLUTIL(op, campos, "0")
        End If
    End With
End Sub


''==================================================================================
''          RUTINAS DE PRUEBA
''==================================================================================
'
' '   Sub SQLUTIL(Opcion As Integer, campos, condicion As String)
' '
'        Dim cSql As New rdoQuery
'        Dim cadena As String
'        Dim i As Integer
'        Dim resultados As rdoResultset
'
'        Set cSql.ActiveConnection = db
'
'        Select Case Opcion
'
'        Case 1:    '<<<<<   INSERTA   >>>>>>
'
'                    cadena = "INSERT INTO " + campos(0, 2) + " ("
'
'                    'NOMBRE DE CAMPOS.
'                    i = 0
'                    While campos(i, 0) <> ""
'                        cadena = cadena + campos(i, 0) + ","
'                        i = i + 1
'                    Wend
'                    cadena = Left(cadena, Len(cadena) - 1)
'                    cadena = cadena + ") VALUES ("
'
'                    'VALORES ASIGNADOS A CADA CAMPO.
'                    i = 0
'                    While campos(i, 0) <> ""
'                        cadena = cadena + "'" + campos(i, 1) + "',"
'                        i = i + 1
'                    Wend
'                    cadena = Left(cadena, Len(cadena) - 1)
'                    cadena = cadena + ")"
'                    cSql.SQL = cadena
'                    cSql.Execute
'
'        Case 2:    '<<<<<   ACTUALIZA   >>>>>>
'
'                    cadena = "UPDATE " + campos(0, 2) + " SET "
'                    i = 0
'                    While campos(i, 0) <> ""
'                        cadena = cadena + campos(i, 0) + "= '" + campos(i, 1) + "',"
'                        i = i + 1
'                    Wend
'                    cadena = Left(cadena, Len(cadena) - 1)
'                    cadena = cadena + " WHERE " + condicion
'                    cSql.SQL = cadena
'                    cSql.Execute
'
'        Case 3:    '<<<<<   ELIMINA   >>>>>>
'                    cadena = "DELETE FROM " + campos(0, 2) + " WHERE " + condicion
'                    cSql.SQL = cadena
'                    cSql.Execute
'
'        Case 4:    '<<<<<   LEE   >>>>>>
'                    cadena = "SELECT "
'                    i = 0
'                    While campos(i, 0) <> ""
'                        cadena = cadena + campos(i, 0) + ","
'                        i = i + 1
'                    Wend
'                    cadena = Left(cadena, Len(cadena) - 1)
'                    cadena = cadena + " FROM " + campos(0, 2) + " WHERE " + condicion
'                    cSql.SQL = cadena
'                    cSql.Execute
'                    If cSql.RowsAffected > 0 Then
'                        Set resultados = cSql.OpenResultset
'                        estado = 0 'SI ENCONTRO DATOS.
'                        i = 0
'                        'TRASPASA LOS DATOS A LA COLUMNA 3 DE LA MATRIZ GLOBAL "campos".
'                        While campos(i, 0) <> ""
'                            campos(i, 3) = resultados(i)
'                            i = i + 1
'                        Wend
'                        resultados.Close
'                        Set resultados = Nothing
'                        With maestro01
'                            '.codigo.Enabled = False
'                        End With
'                    Else
'                        estado = 4 'NO ENCONTRO DATOS.
'                    End If
'
'
'
'
'        End Select
'    End Sub
'
'
'
'
''==================================================================================
'
'
'Sub Crea_Cuenta(codigo_cuenta As String)
''CREA UNA NUEVA CUENTA DEL MAYOR CON LOS DATOS INGRESADOS.
'    Dim cSql As New rdoQuery
'
'    With maestro01
'        Set cSql.ActiveConnection = db
'        cSql.SQL = "INSERT INTO maestroproductos "
'        cSql.SQL = cSql.SQL + "(codigoproducto,descrpcion,unitariomayor,codigolocal,codigoseccion,codigodepto,codigolinea,impuesto,pcostoiva,pvmayorista1,pventadetalle,stockcritico,descuento,datoextra,ubicacion) "
'        cSql.SQL = cSql.SQL + "VALUES ('" & .dato(0).Text & "','" & codigo_cuenta & "','" & .dato(1) & "','" & .dato(2).Text & "','" & .dato(3).Text & "','" & .dato(4).Text & "','" & .dato(5).Text & "','" & .dato(6).Text & "','" & .dato(7).Text & "','" & .dato(8).Text & "','" & .dato(9).Text & "','" & .dato(10).Text & "','" & .dato(11).Text & "')"
'        cSql.Execute
'    End With
'End Sub
'
'Sub Actualiza_Cuenta(codigo_cuenta As String)
''ACTUALIZA UNA CUENTA DEL MAYOR CON LOS DATOS INGRESADOS.
'    Dim cSql As New rdoQuery
'
'    With maestro01
'        Set cSql.ActiveConnection = db
'        cSql.SQL = "UPDATE maestroproductos "
'        cSql.SQL = cSql.SQL = "SET codigoproducto='" & .codigo_cuenta & "',descripcion='" & dato(1).Text & "',unitariomayor='" & .dato(2).Text & "',codigoseccion='" & .dato(3).Text & "',codigodepto='" & .dato(4).Text & "',codigolinea='" & .dato(5).Text & "',impuesto='" & .dato(6).Text & "',pcostoiva='" & .dato(7).Text & "',pvmayoristal='" & .dato(8).Text & "',pventadetalle='" & .dato(9).Text & "',stockcritico='" & .dato(10).Text & "',descuento='" & .dato(11).Text & "',datoextra='" & .dato(12).Text & "',ubicacion='" & .dato(12).Text & "' "
'        cSql.SQL = cSql.SQL + "WHERE codigoproducto='" & codigo_cuenta & "'"
'        cSql.Execute
'    End With
'End Sub
'
'
