Attribute VB_Name = "imprimearchivo"
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" _
(ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Private Declare Function FormatMessage Lib "kernel32" _
Alias "FormatMessageA" _
(ByVal dwFlags As Long, _
lpSource As Any, _
ByVal dwMessageId As Long, _
ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, _
ByVal nSize As Long, _
Arguments As Long) As Long

Private Const SW_HIDE = 0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

' función que imprime un documento de cualquier aplicación
Public Function PrintFile(FileName As String) As Variant
Dim RetVal As Long
Dim sError As String
Dim LenMsg As Long

' se manda imprimir el documento
RetVal = ShellExecute(0&, "print", FileName, 0&, vbNullString, SW_HIDE)
Rem RetVal = ShellExecute(0&, "print", FileName, 0&, vbNullString, vbHide)

' si se ha producido algún error
If RetVal < 33 Then
sError = Space(1024)
' obtenemos el mensaje de error que manda el sistema
LenMsg = FormatMessage( _
FORMAT_MESSAGE_FROM_SYSTEM, _
ByVal 0&, _
RetVal, _
0&, _
sError, _
Len(sError), _
0&)
' devolvemos el mensaje de error
PrintFile = Left(sError, LenMsg - 1)
Else
' la función tuvo éxito
PrintFile = True
End If

End Function


