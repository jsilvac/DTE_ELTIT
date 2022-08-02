Attribute VB_Name = "Funciones_Forms_M_Linea_Departametos"

Sub Limpia_Formulario_Linea_Departamentos()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With maestro04
        For i = 0 To 4
            .dato(i).Text = ""
        Next i
        For i = 0 To 1
            .dato(i).Text = ""
        Next i
        .dato(0).Text = ""
        '.codigo.Mask = "99-99-9999"
    End With
End Sub


Sub Conecta_Maestro_Linea_Departamentos()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro04
        .mld.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=gestioncomercial"
    End With
End Sub



Sub Manejo_Datos_Linea_Departamentos()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    Dim op As Integer
    Dim condicion As String
    
    With maestro03
        codigo_cuenta = .dato(0).Text
        .mld.RecordSource = "SELECT codigolinea FROM maestrolineas WHERE codigolinea='" & codigo_cuenta & "'"
        .mld.Refresh
        If .mld.Recordset.RecordCount > 0 Then 'YA EXISTE LA CUENTA
            'NOMBRES DE CAMPOS EN TABLA
            campos(0, 0) = .dato(0).Tag 'DEPARTAMENTO
            campos(1, 0) = .dato(1).Tag 'LINEA
            campos(2, 0) = .dato(2).Tag 'NOMBRE
            campos(3, 0) = .dato(3).Tag 'PORCENTAJE DESCUENTO
            campos(4, 0) = .dato(4).Tag 'PORCENTAJE UTILIDAD
'            campos(5, 0) = .dato(5).Tag 'CIUDAD
'            campos(6, 0) = .dato(6).Tag 'GIRO
'            campos(7, 0) = .dato(7).Tag 'FONO
                        
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).Text 'DEPARTAMENTO
            campos(1, 1) = .dato(1).Text 'LINEA
            campos(2, 1) = .dato(2).Text 'NOMBRE
            campos(3, 1) = .dato(3).Text 'PORCENTAJE DESCUENTO
            campos(4, 1) = .dato(4).Text 'PORCENTAJE UTILIDAD
'            campos(5, 1) = .dato(5).Text 'CIUDAD
'            campos(6, 1) = .dato(6).Text 'GIRO
'            campos(7, 1) = .dato(7).Text 'FONO
                        
            'NOMBRE DE TABLA.
            campos(0, 2) = "maestrolineas"
            'CONDICION DE LA CONSULTA.
            condicion = "codigolinea=" + "'" + codigo_cuenta + "'"
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
            campos(2, 0) = .dato(2).Tag 'CODIGO SECCION
            campos(3, 0) = .dato(3).Tag 'DESCUENTO VENTA
            campos(4, 0) = .dato(4).Tag 'MARGEN TEORICO
'            campos(5, 0) = .dato(5).Tag 'CIUDAD
'            campos(6, 0) = .dato(6).Tag 'GIRO
'            campos(7, 0) = .dato(7).Tag 'FONO
            
            'NOMBRE DE DATOS EN PROGRAMA
            campos(0, 1) = .dato(0).Text 'CODIGO
            campos(1, 1) = .dato(1).Text 'NOMBRE
            campos(2, 1) = .dato(2).Text 'CODIGO SECCION
            campos(3, 1) = .dato(3).Text 'DESCUENTO VENTA
            campos(4, 1) = .dato(4).Text 'MARGEN TEORICO
'            campos(5, 1) = .dato(5).Text 'CIUDAD
'            campos(6, 1) = .dato(6).Text 'GIRO
'            campos(7, 1) = .dato(7).Text 'FONO
            
            'NOMBRE DE TABLA
            campos(0, 2) = "maestrolineas"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 2
            'op = 1
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, condicion)
            'Call SQLUTIL(op, campos, "0")
        End If
    End With
End Sub

