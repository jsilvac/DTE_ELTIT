Attribute VB_Name = "Funciones_Forms_M_Impuestos"

Sub Limpia_Formulario_Impuestos()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With maestro05
        For i = 0 To 2
            .dato(i).text = ""
        Next i
        For i = 0 To 1
            .dato(i).text = ""
        Next i
        .dato(0).text = ""
       
    End With
End Sub


Sub Conecta_Maestro_Impuestos()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro05
        .mi.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=gestioncomercial"
    End With
End Sub

Sub Manejo_Datos_Impuestos()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    Dim op As Integer
    Dim condicion As String
    
    With maestro05
        codigo_cuenta = .dato(0).text
        .mi.RecordSource = "SELECT codigoimpuesto FROM maestroimpuestos WHERE codigoimpuesto='" & codigo_cuenta & "'"
        .mi.Refresh
        If .mi.Recordset.RecordCount > 0 Then 'YA EXISTE LA CUENTA
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = .dato(0).Tag 'CODIGO
            campos(1, 0) = .dato(1).Tag 'NOMBRE
            campos(2, 0) = .dato(2).Tag 'PORCENTAJE
'            campos(3, 0) = .dato(3).Tag 'PORCENTAJE DESCUENTO
'            campos(4, 0) = .dato(4).Tag 'PORCENTAJE UTILIDAD
'            campos(5, 0) = .dato(5).Tag 'CIUDAD
                        
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).text 'CODIGO
            campos(1, 1) = .dato(1).text 'NOMBRE
            campos(2, 1) = .dato(2).text 'PORCENTAJE
'            campos(3, 1) = .dato(3).text 'PORCENTAJE DESCUENTO
'            campos(4, 1) = .dato(4).text 'PORCENTAJE UTILIDAD
'            campos(5, 1) = .dato(5).Text 'CIUDAD
                        
            'NOMBRE DE TABLA.
            campos(0, 2) = "maestroimpuestos"
            'CONDICION DE LA CONSULTA.
            condicion = "codigoimpuesto=" + "'" + codigo_cuenta + "'"
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
            campos(2, 0) = .dato(2).Tag 'PORCENTAJE
'            campos(3, 0) = .dato(3).Tag 'DESCUENTO VENTA
'            campos(4, 0) = .dato(4).Tag 'MARGEN TEORICO
            
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).text 'CODIGO
            campos(1, 1) = .dato(1).text 'NOMBRE
            campos(2, 1) = .dato(2).text 'PORCENTAJE
'            campos(3, 1) = .dato(3).text 'DESCUENTO VENTA
'            campos(4, 1) = .dato(4).text 'MARGEN TEORICO
'            campos(5, 1) = .dato(5).Text 'CIUDAD
            
            'NOMBRE DE TABLA
            campos(0, 2) = "maestroimpuestos"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 2
            'op = 1
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, condicion)
            'Call SQLUTIL(op, campos, "0")
        End If
    End With
End Sub
