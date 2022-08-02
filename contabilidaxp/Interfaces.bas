Attribute VB_Name = "Interfaces"
Option Explicit
Public path

Function equiAncho(ByVal ancho As Long) As Long
    equiAncho = (ancho * 1024) \ 15360
End Function

Function equiAlto(ByVal alto As Long) As Long
    equiAlto = (alto * 700) \ 11520
End Function

