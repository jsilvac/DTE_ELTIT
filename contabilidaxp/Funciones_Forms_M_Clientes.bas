Attribute VB_Name = "Funciones_Forms_M_Clientes"
Sub Conecta_Maestro_Clientes()
'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With ventas01
        .mc.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=gestioncomercial"
    End With
End Sub



