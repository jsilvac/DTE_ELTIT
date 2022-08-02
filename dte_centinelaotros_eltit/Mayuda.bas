Attribute VB_Name = "Mayuda"
Option Explicit
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'para registrar la dll:                                 '
    '   1.- copiar la dll en c:\windows\system              '
    '   2.- ejecutar el comando regsvr32 nombredeladll.dll  '
    'listo                                                  '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Private help As CAyuda
    Public servidorAyuda As String
    Public basedatosAyuda As String
    Public usuarioAyuda As String
    Public passAyuda As String
    Public tablaAyuda As String
    Public mensajeAyuda As String
    Public cabezasAyuda As Variant
    Public camposAyuda As Variant
    Public largoAyuda As Variant
    Public condicionAyuda As String
    Public cantidadAyuda As Long
    Public sql As sqlventas.sqlventa
    
    
    
Public Sub cargaAyuda(ByRef text As TextBox)
    Set help = New CAyuda
    help.servidor = servidorAyuda
    help.basedatos = basedatosAyuda
    help.usuario = usuarioAyuda
    help.pass = passAyuda
    help.tabla = tablaAyuda
    Set help.txt = text
    help.mensaje = mensajeAyuda
    help.CAMPOS = camposAyuda
    help.cabeceras = cabezasAyuda
    help.campofijo = condicionAyuda
    help.tamaño = largoAyuda
    help.cantidad = cantidadAyuda
    help.abreForm
End Sub

Public Sub cargaAyudaT(ByVal server As String, ByVal database As String, ByVal user As String, ByVal password As String, ByVal tabledatabase, ByRef text As TextBox, ByVal cmps As Variant, ByVal cfijo As String, ByVal largo As Variant, ByVal cant As Long)
    Set help = New CAyuda
    help.servidor = server
    help.basedatos = database
    help.usuario = user
    help.pass = password
    help.tabla = tabledatabase
    Set help.txt = text
    help.CAMPOS = cmps
    help.campofijo = cfijo
    help.cantidad = cant
    help.tamaño = largo
    help.mensaje = mensajeAyuda
    help.cabeceras = cabezas
    help.abreForm
End Sub




