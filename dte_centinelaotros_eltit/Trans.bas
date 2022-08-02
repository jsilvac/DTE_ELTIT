Attribute VB_Name = "Trans"
Option Explicit
    Public Const GWL_EXSTYLE = (-20)
    Public Const WS_EX_LAYERED = &H80000
    Public Const LWA_ALPHA = &H2

Public Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Public Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Function TranslucentForm(frm As Form, TranslucenceLevel As Byte) As Boolean
    SetWindowLong frm.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hwnd, 0, TranslucenceLevel, LWA_ALPHA
    TranslucentForm = Err.LastDllError = 0
End Function

