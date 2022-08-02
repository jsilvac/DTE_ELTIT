Attribute VB_Name = "Funciones_Forms_M_Productos"
Sub Conectar_BD()
'RUTINA PARA CONECTAR A LA BASE DE DATOS
    servidor = "eltitxp"
    usuario = "root"
    password = "123"
    basedatos = "conta01"
    On Error GoTo controlerror
        Call Conectar(servidor, basedatos, usuario, password)
    Exit Sub
controlerror:
    
    Resume Next
End Sub


Sub Limpia_Formulario_Productos()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With maestro01
        For i = 0 To 13
            .dato(i).text = ""
        Next i
        For i = 0 To 1
            .dato(i).text = ""
        Next i
        .dato(0).text = ""
        '.codigo.Mask = "99-99-9999"
    End With
End Sub


Sub Conecta_Maestro_Productos()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro01
        .mcm.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=conta01"
    End With
End Sub


Sub Manejo_Datos_Productos()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    Dim op As Integer
    Dim condicion As String
    
    With maestro01
        codigo_cuenta = .dato(0).text
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
            campos(0, 1) = .dato(0).text 'TIPO
            campos(1, 1) = .dato(1).text 'RUT
            campos(2, 1) = .dato(2).text 'NOMBRE
            campos(3, 1) = .dato(3).text 'DIRECCION
            campos(4, 1) = .dato(4).text 'COMUNA
            campos(5, 1) = .dato(5).text 'CIUDAD
            campos(6, 1) = .dato(6).text 'GIRO
            campos(7, 1) = .dato(7).text 'FONO
            campos(8, 1) = .dato(8).text 'FAX
            campos(9, 1) = .dato(9).text 'CELULAR
            campos(10, 1) = .dato(10).text 'EMAIL
            campos(11, 1) = .dato(11).text 'CONTACTO
            campos(12, 1) = .dato(12).text 'DEST. CHEQUE
            campos(13, 0) = .dato(13).text 'DEST. CHEQUE
            
            'NOMBRE DE TABLA.
            campos(0, 2) = "maestroproductos"
            'CONDICION DE LA CONSULTA.
            condicion = "codigoproducto=" + "'" + codigo_cuenta + "'"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 3
            'op = 2
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, condicion)
            'Call SQLUTIL(op, campos, condicion)
            
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
            campos(0, 1) = .dato(0).text 'TIPO
            campos(1, 1) = .dato(1).text 'RUT
            campos(2, 1) = .dato(2).text 'NOMBRE
            campos(3, 1) = .dato(3).text 'DIRECCION
            campos(4, 1) = .dato(4).text 'COMUNA
            campos(5, 1) = .dato(5).text 'CIUDAD
            campos(6, 1) = .dato(6).text 'GIRO
            campos(7, 1) = .dato(7).text 'FONO
            campos(8, 1) = .dato(8).text 'FAX
            campos(9, 1) = .dato(9).text 'CELULAR
            campos(10, 1) = .dato(10).text 'EMAIL
            campos(11, 1) = .dato(11).text 'CONTACTO
            campos(12, 1) = .dato(12).text 'DEST. CHEQUE
            campos(13, 1) = .dato(13).text 'DEST. CHEQUE
            
            'NOMBRE DE TABLA
            campos(0, 2) = "maestroproductos"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 2
            'op = 1
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, condicion)
            'Call SQLUTIL(op, campos, "0")
        End If
    End With
End Sub


'
'
'Sub Carga_Cuenta(codigo_cuenta As String)
''RUTINA QUE CARGA UNA CUENTA INGRESADA SI ESTA EXISTE.
'    Dim cSql As New rdoQuery
'    Dim resultados As rdoResultset
'
'    With maestro01
'        Set cSql.ActiveConnection = db
'        cSql.SQL = "SELECT descripcion,unitariomayor,seccion,departamento,linea FROM maestroproductos WHERE codigoproducto='" & codigo_cuenta & "'"
'        cSql.Execute
'        If cSql.RowsAffected > 0 Then
'            Set resultados = cSql.OpenResultset
'            While Not resultados.EOF
'                .dato1.text = codigo_cuenta
'                .dato(0).text = resultados(0)
'                .dato(1).text = resultados(1)
'                .dato(2).text = resultados(2)
'                .dato(3).text = resultados(3)
'                .dato(4).text = resultados(4)
'                resultados.MoveNext
'            Wend
'            resultados.Close
'            Set resultados = Nothing
'            .dato1.Enabled = False
'        End If
'    End With
'
'End Sub


Sub busca_seccion()

    End
    
End Sub
