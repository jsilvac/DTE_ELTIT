Attribute VB_Name = "Funciones_Forms_M_Bodegas"
Sub Carga_Cuenta_Bodegas(codigo_cuenta As String)
'RUTINA QUE CARGA UNA CUENTA INGRESADA SI ESTA EXISTE.
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    With maestro07
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT bodega,nombre,direccion,ciudad,otros FROM maestrobodegas WHERE codigobodega='" & codigo_cuenta & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                .codbodegas.Text = codigo_cuenta
                .txtbodegas(0).Text = resultados(0)
                .txtbodegas(1).Text = resultados(1)
                .txtbodegas(2).Text = resultados(2)
                .txtbodegas(3).Text = resultados(3)
                .txtbodegas(4).Text = resultados(4)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            .codbodegas.Enabled = False
        End If
    End With

End Sub


Sub Conecta_Maestro_Bodegas()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro07
        .mb.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=gestioncomercial"
    End With
End Sub

Sub Manejo_Datos_Bodegas()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    
    With maestro07
        codigo_cuenta = Mid(.codbodegas.Text, 1, 3)
        .mb.RecordSource = "SELECT codigobodega FROM maestrobodegas WHERE codigobodega='" & codigo_cuenta & "'"
        .mb.Refresh
        If .mb.Recordset.RecordCount > 0 Then 'YA EXISTE LA CUENTA
            Call Actualiza_Maestro_Bodegas(codigo_cuenta)
        Else                                            'CREA UNA CUENTA NUEVA
            Call Guarda_Maestro_Bodegas(codigo_cuenta)
        End If
    End With
End Sub


Sub Actualiza_Maestro_Bodegas(codigo_cuenta As String)
'ACTUALIZA UNA CUENTA DEL MAYOR CON LOS DATOS INGRESADOS.
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset

    With maestro07
        Set cSql.ActiveConnection = db
        cSql.SQL = "UPDATE maestrobodegas SET bodega='" & .txtbodegas(0).Text & "',nombre='" & .txtbodegas(1) & "',direccion='" & .txtbodegas(2).Text & "',ciudad='" & .txtbodegas(3).Text & "',otros='" & .txtbodegas(4).Text & "' WHERE codigobodega='" & codigo_cuenta & "'"
        cSql.Execute
    End With
End Sub


Sub Guarda_Maestro_Bodegas(codigo_cuenta As String)
'GUARDA UNA NUEVA CUENTA DEL MAYOR CON LOS DATOS INGRESADOS
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset

    With maestro07
        Set cSql.ActiveConnection = db
        cSql.SQL = "INSERT INTO maestrobodegas (codigobodega,bodega,nombre,direccion,ciudad,otros) VALUES ('" & codigo_cuenta & "','" & .txtbodegas(0).Text & "','" & .txtbodegas(1).Text & "','" & .txtbodegas(2).Text & "','" & .txtbodegas(3).Text & "','" & .txtbodegas(4).Text & "')"
        cSql.Execute
    End With
End Sub


Sub Limpia_Formulario_Bodegas()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With maestro07
        For i = 0 To 4
            .txtbodegas(i).Text = ""
        Next i
        For i = 0 To 1
            .txtbodegas(i).Text = ""
        Next i
        .codbodegas.Text = ""
        '.codigo.Mask = "99-99-9999"
    End With
End Sub

