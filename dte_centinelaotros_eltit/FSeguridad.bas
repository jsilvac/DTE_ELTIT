Attribute VB_Name = "FSeguridad"
Option Explicit
    Public titCaption As String
    
    Public Function Verificar(ByVal usuario As String, ByVal password As String) As Boolean
        Dim CAMPOS(5, 5) As String
        
        Dim op As Integer
        On Error GoTo error:
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "usuario"
        CAMPOS(1, 0) = "clave"
        CAMPOS(2, 0) = ""
        
        CAMPOS(0, 2) = clientesistema & "auditoria.segu_usuarios"
        
        condicion = "usuario = '" & usuario & "' AND clave ='" & password & "' "
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            Verificar = True
'            If UCase(titCaption) <> "PRINCIPAL" Then
'                Verificar = verificarNivel(usuario)
'            End If
        Else
            Verificar = False
        End If
      Exit Function
error:
      MsgBox "NO DEBE INGRESAR  '  EN LOS CAMPOS ", vbCritical, "ATENCION "
    End Function

    Private Function verificarNivel(ByVal usuario As String) As Boolean
        Dim CAMPOS(5, 5) As String
       
        Dim op As Integer
       Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "usuario"
        CAMPOS(1, 0) = "programa"
        CAMPOS(2, 0) = ""

        CAMPOS(0, 2) = "segu_permisos"

        condicion = "usuario = '" & usuario & "' AND programa = '" & titCaption & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            verificarNivel = True
        Else
            verificarNivel = False
        End If

   
    End Function
    
