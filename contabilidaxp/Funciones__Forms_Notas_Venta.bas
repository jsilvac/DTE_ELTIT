Attribute VB_Name = "Funciones_Forms_Notas_Venta"
Sub Cabecera_Notas_Venta() 'OK
    With ventas13
        t$ = "<CODIGO             |<DESCRIPCION                                                                                                        |>PRECIO COSTO+IVA     |>PRECIO MAYORISTA     |>PRECIO DETALLE     "
        .Productos.FormatString = t$
        t$ = "<Nº NOTA VENTA|<PRODUCTO                                                                                                                                  |>PRECIO VENTA UNIT.    |>CANTIDAD     |>TOTAL               "
        .NotaVenta.FormatString = t$
    End With
End Sub

Sub Carga_Numero_Nota_Venta()
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    With ventas13
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT MAX(numero) "
        cSql.SQL = cSql.SQL + "FROM notaventa"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            If Not IsNull(resultados(0).Value) Then
                .NumeroNota.Caption = resultados(0) + 1
            Else
                .NumeroNota.Caption = 1
            End If
            resultados.Close
            Set resultados = Nothing
        End If
    End With

End Sub

Sub Carga_Locales()
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    With ventas13
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigolocal,nombre "
        cSql.SQL = cSql.SQL + "FROM maestrolocales"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                .Locales.AddItem resultados(0) + vbTab + resultados(1)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
    End With

End Sub

Sub Carga_Secciones(codigo_local As String)

    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    With ventas13
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigoseccion,nombre "
        cSql.SQL = cSql.SQL + "FROM maestrosecciones "
        cSql.SQL = cSql.SQL + "WHERE codigolocal='" & codigo_local & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                .Secciones.AddItem resultados(0) + vbTab + resultados(1)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
    End With

End Sub

Sub Carga_Departamentos(codigo_local As String, codigo_seccion As String)

    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    With ventas13
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigodepto,nombre "
        cSql.SQL = cSql.SQL + "FROM maestrodepartamentos "
        cSql.SQL = cSql.SQL + "WHERE codigoseccion='" & codigo_seccion & "' "
        cSql.SQL = cSql.SQL + "AND codigolocal='" & codigo_local & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                .Departamentos.AddItem resultados(0) + vbTab + resultados(1)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
    End With

End Sub

Sub Carga_Lineas(codigo_local As String, codigo_seccion As String, codigo_depto As String)

    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    With ventas13
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigolinea,nombre "
        cSql.SQL = cSql.SQL + "FROM maestrolineas "
        cSql.SQL = cSql.SQL + "WHERE codigoseccion='" & codigo_seccion & "' "
        cSql.SQL = cSql.SQL + "AND codigodepto='" & codigo_depto & "' "
        cSql.SQL = cSql.SQL + "AND codigolocal='" & codigo_local & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                .Lineas.AddItem resultados(0) + vbTab + resultados(1)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
    End With

End Sub

Sub Carga_Productos(codigo_local As String, codigo_seccion As String, codigo_depto As String, codigo_linea As String)

    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    With ventas13
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigoproducto,descripcion,pcostoiva,pvmayorista1,pventadetalle "
        cSql.SQL = cSql.SQL + "FROM maestroproductos "
        cSql.SQL = cSql.SQL + "WHERE codigoseccion='" & codigo_seccion & "' "
        cSql.SQL = cSql.SQL + "AND codigodepto='" & codigo_depto & "' "
        cSql.SQL = cSql.SQL + "AND codigolinea='" & codigo_linea & "' "
        cSql.SQL = cSql.SQL + "AND codigolocal='" & codigo_local & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                .Productos.Rows = 2
                Cabecera_Notas_Venta
                .Productos.AddItem resultados(0) & vbTab & resultados(1) & vbTab & Format(resultados(2), "$ ###,###,###.00") & vbTab & Format(resultados(3), "$ ###,###,###.00") & vbTab & Format(resultados(4), "$ ###,###,###.00"), .Productos.Rows - 1
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        Else
            .Productos.Rows = 2
            Cabecera_Notas_Venta
        End If
    End With

End Sub

'====================================================================================================
'                      MANEJO CONTROL DATA
'====================================================================================================

Sub Conecta_Control_Data()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.

    With ventas13
        .ControlData.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=" & servidor & ";uid=" & usuario & ";pwd=" & password & ";database=" & basedatos & ""
    End With
End Sub


Sub nota_venta()
    
    Dim numero_nota As Integer
    
    With ventas13
        numero_nota = Val(.NumeroNota.Caption)
        .ControlData.RecordSource = "SELECT * FROM detallenotaventa WHERE notaventa='" & numero_nota & "'"
        .ControlData.Refresh
        Cabecera_Notas_Venta
    End With
    
End Sub

