Attribute VB_Name = "Mayuda"
Option Explicit
    Private help As CAyuda
    Public mensajeAyuda As String
    Public cabezas As Variant

Public Sub cargaAyudaT(ByVal server As String, ByVal database As String, ByVal user As String, ByVal password As String, ByVal tabledatabase, ByRef text As TextBox, ByVal cmps As Variant, ByVal cfijo As String, ByVal largo As Variant, ByVal cant As Long)
    Set help = New CAyuda
    help.servidor = server
    help.basedatos = database
    help.USUARIO = user
    help.pass = password
    help.tabla = tabledatabase
    Set help.txt = text
    help.campos = cmps
    help.campofijo = cfijo
    help.cantidad = cant
    help.tamaño = largo
    help.MENSAJE = mensajeAyuda
    help.Cabeceras = cabezas
    help.abreForm
End Sub


'Ejemplo de implementacion
'Sub ayudatipocuenta(ByRef caja As TextBox)
'    Dim campos As Variant
'    Dim cfijo As Variant
'    Dim largo As Variant
'    campos = Array("ctacte", "glosa")
'    largo = Array("8s", "40s")
'    cfijo = "CTACTE > '00'"
'    cabezas = Array("codigo", "nombre")
'    mensajeAyuda = "Ayuda tipo de Cuentas Corrientes"        
'    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentasdelmayor", dato1, campos, cfijo, largo, 2)
'    caja.Enabled = True
'    caja.SetFocus
'End Sub


