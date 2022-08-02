Attribute VB_Name = "Funciones_Forms_M_Departamentos"

Sub Limpia_Formulario_Departamentos()
'LIMPIA LOS CONTROLES DEL FORMULARIO
    Dim i As Integer
    With maestro03
        For i = 0 To 4
            .dato(i).Text = ""
        Next i
        For i = 0 To 1
            .dato(i).Text = ""
        Next i
        .dato(0).Text = ""
        
    End With
End Sub


Sub Conecta_Maestro_Departamentos()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro03
        .md.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=gestioncomercial"
    End With
End Sub



Sub Manejo_Datos_Departamentos()
'RUTINA QUE DETERMINA SI DEBE GUARDAR O ACTUALIZAR LOS DATOS.
    Dim codigo_cuenta As String
    Dim op As Integer
    Dim condicion As String
    
    With maestro03
        codigo_cuenta = .dato(0).Text
        .md.RecordSource = "SELECT codigodepto FROM maestrodepartamentos WHERE codigodepto='" & codigo_cuenta & "'"
        .md.Refresh
        If .md.Recordset.RecordCount > 0 Then 'YA EXISTE LA CUENTA
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
                        
            'NOMBRE DE TABLA.
            campos(0, 2) = "maestrodepartamentos"
            'CONDICION DE LA CONSULTA.
            condicion = "codigodepto=" + "'" + codigo_cuenta + "'"
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
            campos(0, 2) = "maestrodepartamentos"
            'OPCION CON QUE SE LLAMA LA FUNCION.
            op = 2
            'op = 1
            Set SQLUTIL.conexion = db
            Call SQLUTIL.SQLUTIL(op, condicion)
            'Call SQLUTIL(op, campos, "0")
        End If
    End With
End Sub
