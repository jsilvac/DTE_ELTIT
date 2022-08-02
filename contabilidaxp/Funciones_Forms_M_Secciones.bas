Attribute VB_Name = "Funciones_Forms_M_Secciones"

Sub Conecta_Maestro_Secciones()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro02
        .ms.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=gestioncomercial"
    End With
End Sub


Sub Limpia_Formulario_Secciones()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With maestro02
        For i = 0 To 2
            .dato(i).Text = ""
        Next i
        For i = 0 To 0
            .dato(i).Text = ""
        Next i
        .dato(0).Text = ""
       
    End With
End Sub


Sub Manejo_Datos_Secciones()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    Dim op As Integer
    Dim condicion As String
    
    With maestro02
        codigo_cuenta = .dato(0).Text
        .ms.RecordSource = "SELECT codigoseccion FROM maestrosecciones WHERE codigoseccion='" & codigo_cuenta & "'"
        .ms.Refresh
        If .ms.Recordset.RecordCount > 0 Then 'YA EXISTE LA CUENTA
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = .dato(0).Tag 'CODIGO
            campos(1, 0) = .dato(1).Tag 'NOMBRE
            campos(2, 0) = .dato(2).Tag 'PORCENTAJE
            'campos(3, 0) = .dato(3).Tag '.....
            'campos(4, 0) = .dato(4).Tag '.....
            'campos(5, 0) = .dato(5).Tag '.....
                        
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).Text 'CODIGO
            campos(1, 1) = .dato(1).Text 'NOMBRE
            campos(2, 1) = .dato(2).Text 'PORCENTAJE
            'campos(3, 1) = .dato(3).Text '.....
            'campos(4, 1) = .dato(4).Text '.....
            'campos(5, 1) = .dato(5).Text '.....
            
            'NOMBRE DE TABLA.
            campos(0, 2) = "maestrosecciones"
            'CONDICION DE LA CONSULTA.
            condicion = "codigoseccion=" + "'" + codigo_cuenta + "'"
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
            'campos(3, 0) = .dato(3).Tag '.....
            'campos(4, 0) = .dato(4).Tag '.....
            'campos(5, 0) = .dato(5).Tag '.....
            
            
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).Text 'CODIGO
            campos(1, 1) = .dato(1).Text 'NOMBRE
            campos(2, 1) = .dato(2).Text 'PORCENTAJE
            'campos(3, 1) = .dato(3).Text '.....
            'campos(4, 1) = .dato(4).Text '.....
            'campos(5, 1) = .dato(5).Text '.....
            
            'NOMBRE DE TABLA
            campos(0, 2) = "maestrosecciones"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 2
            'op = 1
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, condicion)
            'Call SQLUTIL(op, campos, "0")
        End If
    End With
End Sub
