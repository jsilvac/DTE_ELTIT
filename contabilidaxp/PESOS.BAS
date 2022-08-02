Attribute VB_Name = "pesos"
Public u(15) As String

Public K(15) As String

Public Q(15) As String
Sub cargapalabras()
Q(0) = "": K(0) = "": u(0) = ""
Q(10) = "CIEN "
Q(1) = "CIENTO "
Q(2) = "DOSCIENTOS "
 Q(3) = "TRESCIENTOS "
Q(4) = "CUATROCIENTOS "
Q(5) = "QUINIENTOS "
Q(6) = "SEISCIENTOS "
Q(7) = "SETECIENTOS "
Q(8) = "OCHOCIENTOS "
Q(9) = "NOVECIENTOS "
 K(1) = "DIEZ "
 K(11) = "ONCE "
K(12) = "DOCE "
K(13) = "TRECE "
K(14) = "CATORCE "
K(15) = "QUINCE "
 K(2) = "VEINTE "
K(3) = "TREINTA "
K(4) = "CUARENTA "
K(5) = "CINCUENTA "
K(6) = "SESENTA "
K(7) = "SETENTA "
K(8) = "OCHENTA "
K(9) = "NOVENTA "
u(1) = "UN "
u(2) = "DOS "
u(3) = "TRES "
u(4) = "CUATRO "
u(5) = "CINCO "
u(6) = "SEIS "
 u(7) = "SIETE "
u(8) = "OCHO "
u(9) = "NUEVE "


End Sub
Sub palabras(totg)
Dim nume As String
Dim t As Double
totg = Int(totg)
t = Len(Str$(totg)) - 1
nume = String(9 - t, "0") + Mid(Str(totg), 2, t)
millon = ""
If Val(Mid(nume, 1, 3)) <> 0 Then
        If Val(Mid$(nume$, 1, 3)) = 1 Then
        millon$ = "UN MILLON "
        GoTo 58530
        End If
        
58420: mil$ = Q(Val(Mid$(nume$, 1, 1)))
58430: If Val(Mid$(nume$, 2, 2)) < 16 Then GL$ = "" Else GL$ = "Y "
58440: If Val(Mid$(nume$, 3, 1)) = 0 Then GL$ = ""
58450: If Val(Mid$(nume$, 2, 2)) > 10 And Val(Mid$(nume$, 2, 2)) < 16 Then MI1$ = K(Val(Mid$(nume$, 2, 2))) + " MILLONES ": MI2$ = "": GoTo 58520
58500: MI1$ = K(Val(Mid$(nume$, 2, 1))) + GL$
58510: MI2$ = u(Val(Mid$(nume$, 3, 1))) + "MILLONES "
58520: millon$ = mil$ + MI1$ + MI2$
End If

58530 If Val(Mid$(nume$, 4, 3)) = 0 Then MILES$ = "": GoTo 58660
58540 If Val(Mid$(nume$, 4, 3)) = 100 Then MILES$ = "CIEN MIL ": GoTo 58660
58550 MXL$ = Q(Val(Mid$(nume$, 4, 1)))
58560 If Val(Mid$(nume$, 5, 2)) < 16 Then GL$ = " " Else GL$ = "Y "
58570 If Val(Mid$(nume$, 6, 1)) = 0 Then GL$ = ""
58580 If Val(Mid$(nume$, 5, 2)) > 10 And Val(Mid$(nume$, 5, 2)) < 16 Then MX1$ = K(Val(Mid$(nume$, 5, 2))) + "MIL ": MX2$ = "": GoTo 58650
58630 MX1$ = K(Val(Mid$(nume$, 5, 1))) + GL$
58640 MX2$ = u(Val(Mid$(nume$, 6, 1))) + "MIL "
58650 MILES$ = MXL$ + MX1$ + MX2$
58660 If Val(Mid$(nume$, 7, 3)) = 0 Then pesos$ = "": GoTo 58780
58670 MWL$ = Q(Val(Mid$(nume$, 7, 1)))
58680 If Val(Mid$(nume$, 8, 2)) > 10 And Val(Mid$(nume$, 8, 2)) < 16 Then MW1$ = K(Val(Mid$(nume$, 8, 2))): MW2$ = "": GoTo 58770
58730 If Val(Mid$(nume$, 8, 2)) < 16 Then GL$ = " " Else GL$ = "Y "
58740 If Val(Mid$(nume$, 9, 1)) = 0 Then GL$ = ""
58750 MW1$ = K(Val(Mid$(nume$, 8, 1))) + GL$
58760 MW2$ = u(Val(Mid$(nume$, 9, 1)))
58770 pesos$ = MWL$ + MW1$ + MW2$
      If Mid$(nume$, 7, 3) = "100" Then pesos$ = "CIEN "
58780 MONTOCH$ = millon$ + MILES$ + pesos$
      largo = Len(MONTOCH$)
      monto$ = millon$ + MILES$
      MONTO2$ = pesos$ + " PESOS"
      Rem IF LARGO > 72 THEN TOT72 = LARGO ELSE TOT72 = 72
      MONTOCH$ = MONTOCH$ + String$(105 - largo, "*")

End Sub

