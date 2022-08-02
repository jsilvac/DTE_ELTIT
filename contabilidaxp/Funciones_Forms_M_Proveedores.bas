Attribute VB_Name = "Funciones_Forms_M_Proveedores"

Sub Limpia_Formulario_Proveedores()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With maestro09
        For i = 0 To 9
            .mskproveedores(i).Text = ""
        Next i
        'For i = 0 To 1
        '    .mskproveedores(i).Text = ""
        'Next i
        .rutproveedores.Enabled = True
        .rutproveedores.Text = "  .   .   "
        .rutproveedores.Mask = "99.999.999"
    End With
End Sub



Sub Conecta_Maestro_Proveedores()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro09
        .mprov.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=gestioncomercial"
    End With
End Sub



Sub Carga_Cuenta_Proveedores(codigo_cuenta As String)
'RUTINA QUE CARGA UNA CUENTA INGRESADA SI ESTA EXISTE.
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    With maestro09
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT nombre,direccion,comuna,ciudad,fono1,fono2,fax,contacto,convenio,visitames FROM maestroproveedores WHERE rutproveedor='" & codigo_cuenta & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                .rutproveedores.Text = Mid(codigo_cuenta, 1, 2) & "." & Mid(codigo_cuenta, 3, 3) & "." & Mid(codigo_cuenta, 5, 3)
                .mskproveedores(0).Text = resultados(0)
                .mskproveedores(1).Text = resultados(1)
                .mskproveedores(2).Text = resultados(2)
                .mskproveedores(3).Text = resultados(3)
                .mskproveedores(4).Text = resultados(4)
                .mskproveedores(5).Text = resultados(5)
                .mskproveedores(6).Text = resultados(6)
                .mskproveedores(7).Text = resultados(7)
                .mskproveedores(8).Text = resultados(8)
                .mskproveedores(9).Text = resultados(9)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            .rutproveedores.Enabled = False
        End If
    End With

End Sub

Sub Manejo_Datos_Proveedores()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    
    With maestro09
        codigo_cuenta = Mid(.rutproveedores.Text, 1, 2) & Mid(.rutproveedores.Text, 4, 3) & Mid(.rutproveedores.Text, 8, 3)
        .mprov.RecordSource = "SELECT rutproveedor FROM maestroproveedores WHERE rutproveedor='" & codigo_cuenta & "'"
        .mprov.Refresh
        If .mprov.Recordset.RecordCount > 0 Then 'YA EXISTE LA CUENTA
            Call Actualiza_Maestro_Proveedores(codigo_cuenta)
        Else                                            'CREA UNA CUENTA NUEVA
            Call Guarda_Maestro_Proveedores(codigo_cuenta)
        End If
    End With
End Sub


Sub Actualiza_Maestro_Proveedores(codigo_cuenta As String)
'ACTUALIZA UNA CUENTA DEL MAYOR CON LOS DATOS INGRESADOS.
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset

    With maestro09
        Set cSql.ActiveConnection = db
        cSql.SQL = "UPDATE maestroproveedores SET nombre='" & .mskproveedores(0).Text & "',direccion='" & .mskproveedores(1) & "',direccion='" & .mskproveedores(2).Text & "',comuna='" & .mskproveedores(3).Text & "',ciudad='" & .mskproveedores(4).Text & "',fono1='" & .mskproveedores(4).Text & "',fono2='" & .mskproveedores(5).Text & "',fax='" & .mskproveedores(6).Text & "',contacto='" & .mskproveedores(7).Text & "',convenio='" & .mskproveedores(8).Text & "',visitames='" & .mskproveedores(9).Text & "' WHERE rutproveedor='" & codigo_cuenta & "'"
        cSql.Execute
    End With
End Sub


Sub Guarda_Maestro_Proveedores(codigo_cuenta As String)
'GUARDA UNA NUEVA CUENTA DEL MAYOR CON LOS DATOS INGRESADOS
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset

    With maestro09
        Set cSql.ActiveConnection = db
        cSql.SQL = "INSERT INTO maestroproveedores (rutproveedor,nombre,direccion,comuna,ciudad,fono1,fono2,fax,contacto,convenio,visitames) VALUES ('" & codigo_cuenta & "','" & .mskproveedores(0).Text & "','" & .mskproveedores(1).Text & "','" & .mskproveedores(2).Text & "','" & .mskproveedores(3).Text & "','" & .mskproveedores(4).Text & "','" & .mskproveedores(5).Text & "','" & .mskproveedores(6).Text & "','" & .mskproveedores(7).Text & "','" & .mskproveedores(8).Text & "','" & .mskproveedores(9).Text & "')"
        cSql.Execute
    End With
End Sub
