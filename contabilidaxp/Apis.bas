Attribute VB_Name = "apis"
Global MODI As Integer
Global SS As String
Option Explicit
Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, _
            ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, _
            ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, _
            ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, _
            ByVal BLENDFUNCT As Long) As Long
    'API PARA DEJAR UNA VENTANA SIEMPRE VISIBLE
 Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
            ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2
    'CODIGO PARA PONERLO EN EL LOAD DEL FORM
    'SetWindowPos Me.hwnd, HWND_NOTOPMOST, Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight, 0
    
Function Degradado(ByVal frmDestino As Object, ByVal frmOrigen As Object, _
            ByVal posx1 As Long, ByVal posx2 As Long, ByVal posy1 As Long, _
            ByVal posy2 As Long, ByVal transparencia As Long)
    transparencia = vbBlue - CLng(transparencia) * (vbYellow + 1)
    AlphaBlend frmDestino.hdc, 0, 0, posx2, posy2, frmOrigen.hdc, posx1, posy1, posx2, posy2, transparencia
End Function


