Attribute VB_Name = "Funciones_Forms_M_Locales"

Sub Limpia_Formulario_Locales()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With maestro06
        For i = 0 To 7
            .dato(i).text = ""
        Next i
        For i = 0 To 1
            .dato(i).text = ""
        Next i
        .dato(0).text = ""
        
    End With
End Sub


Sub Conecta_Maestro_Locales()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro06
        .ml.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=gestioncomercial"
    End With
End Sub


Sub Manejo_Datos_Locales()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    Dim op As Integer
    Dim condicion As String
    
    With maestro06
        codigo_cuenta = .dato(0).text
        .ml.RecordSource = "SELECT codigolocal FROM maestrolocales WHERE codigolocal='" & codigo_cuenta & "'"
        .ml.Refresh
        If .ml.Recordset.RecordCount > 0 Then 'YA EXISTE LA CUENTA
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = .dato(0).Tag 'CODIGO
            campos(1, 0) = .dato(1).Tag 'NOMBRE
            campos(2, 0) = .dato(2).Tag 'DIRECCION
            campos(3, 0) = .dato(3).Tag 'COMUNA
            campos(4, 0) = .dato(4).Tag 'CIUDAD
            campos(5, 0) = .dato(5).Tag 'TLOCAL
            campos(6, 0) = .dato(6).Tag 'RUT
            campos(7, 0) = .dato(7).Tag 'AUDITORIA
'            campos(8, 0) = .dato(8).Tag 'FAX
'            campos(9, 0) = .dato(9).Tag 'CELULAR
'            campos(10, 0) = .dato(10).Tag 'EMAIL
                        
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).text 'CODIGO
            campos(1, 1) = .dato(1).text 'NOMBRE
            campos(2, 1) = .dato(2).text 'DIRECCION
            campos(3, 1) = .dato(3).text 'COMUNA
            campos(4, 1) = .dato(4).text 'CIUDAD
            campos(5, 1) = .dato(5).text 'TLOCAL
            campos(6, 1) = .dato(6).text 'RUT
            campos(7, 1) = .dato(7).text 'AUDITORIA
'            campos(8, 1) = .dato(8).text 'FAX
'            campos(9, 1) = .dato(9).text 'CELULAR
'            campos(10, 1) = .dato(10).text 'EMAIL
                        
            'NOMBRE DE TABLA.
            campos(0, 2) = "maestrolocales"
            'CONDICION DE LA CONSULTA.
            condicion = "codigolocal=" + "'" + codigo_cuenta + "'"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 3
            'op = 2
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, condicion)
            'Call SQLUTIL(op, campos, condicion)
            
        Else                                            'CREA UNA CUENTA NUEVA
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = .dato(0).Tag 'CODIGO
            campos(1, 0) = .dato(1).Tag 'NOMBRE
            campos(2, 0) = .dato(2).Tag 'DIRECCION
            campos(3, 0) = .dato(3).Tag 'COMUNA
            campos(4, 0) = .dato(4).Tag 'CIUDAD
            campos(5, 0) = .dato(5).Tag 'TLOCAL
            campos(6, 0) = .dato(6).Tag 'RUT
            campos(7, 0) = .dato(7).Tag 'AUDITORIA
'            campos(8, 0) = .dato(8).Tag 'FAX
'            campos(9, 0) = .dato(9).Tag 'CELULAR
'            campos(10, 0) = .dato(10).Tag 'EMAIL

            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).text 'CODIGO
            campos(1, 1) = .dato(1).text 'NOMBRE
            campos(2, 1) = .dato(2).text 'DIRECCION
            campos(3, 1) = .dato(3).text 'COMUNA
            campos(4, 1) = .dato(4).text 'CIUDAD
            campos(5, 1) = .dato(5).text 'TLOCAL
            campos(6, 1) = .dato(6).text 'RUT
            campos(7, 1) = .dato(7).text 'AUDITORIA
'            campos(8, 1) = .dato(8).text 'FAX
'            campos(9, 1) = .dato(9).text 'CELULAR
'            campos(10, 1) = .dato(10).text 'EMAIL
            
            'NOMBRE DE TABLA
            campos(0, 2) = "maestrolocales"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 2
            'op = 1
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, condicion)
            'Call SQLUTIL(op, campos, "0")
        End If
    End With
End Sub
