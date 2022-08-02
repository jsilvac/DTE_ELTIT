Attribute VB_Name = "Validaciones"
Option Explicit
Public Num, letr As Long
Public coma As Integer
Public largocero As Integer
Public xmlcompra As Boolean
Public xmlventa As Boolean
Public Function leer_certificado_digital(ByVal s As String, ByVal p As String) As String
    ' s=rutacertificado, p=certificado
    Dim i As Integer, R As String
    Dim C1 As Integer, c2 As Integer
    R = ""
    If Len(p) > 0 Then
        If InStr(1, s, "FLAG", vbBinaryCompare) <> 0 Then
            s = Replace(s, "FLAG", Chr(13))
        End If
        For i = 1 To Len(s)
            C1 = Asc(Mid(s, i, 1))
            If i > Len(p) Then
                c2 = Asc(Mid(p, i Mod Len(p) + 1, 1))
            Else
                c2 = Asc(Mid(p, i, 1))
            End If
            C1 = C1 - c2 - 64
            If Sgn(C1) = -1 Then C1 = 256 + C1
                R = R + Chr(C1)
        Next i
    Else
        R = s
    End If
    leer_certificado_digital = R
End Function

Function esNumero(ByVal Num As Long) As Long
    If Num = vbKeyBack Then GoTo fin:
    If Num = vbKeyBack Then GoTo fin:
    If Num = vbKeyReturn Then GoTo fin:
    If Num > 47 And Num < 58 Then GoTo fin:
    If Num = 46 And snum = 1 Then Num = 44: GoTo fin:
    esNumero = 0: GoTo no:
fin:
    esNumero = Num
no:

End Function

Function rut(ByVal numrut As String) As String
    Dim guia
    Dim mataux(9) As Integer
    Dim i, suma As Integer
    guia = Array("4", "3", "2", "7", "6", "5", "4", "3", "2")
    suma = 0
    For i = 0 To 8
        mataux(i) = Val(guia(i)) * Val(Mid(numrut, i + 1, 1))
        suma = suma + mataux(i)
    Next
    rut = 11 - suma Mod 11
    Select Case rut
        Case "11"
            rut = "0"
        Case "10"
            rut = "K"
    End Select
End Function

Sub ceros(ByRef caja As TextBox)
lar = Len(caja.text)
largocero = caja.MaxLength - lar
caja.text = String(largocero, "0") & caja.text

End Sub
Sub ESPACIOS(ByRef caja As TextBox)
lar = Len(caja.text)
largocero = caja.MaxLength - lar
caja.text = String(largocero, " ") & caja.text

End Sub

Sub formato(ByRef caja As TextBox, ByRef deci As Integer)
If caja.text = "" Then caja.text = "0"
valor = CDbl(caja.text)
If deci = 1 Then caja.text = Format(valor, "###,###,###,##0.00")
If deci = 0 Then caja.text = Format(valor, "###,###,###,##0")
End Sub
Sub RECUPERAFECHA()
dia = Mid(Date, 1, 2)
MES = Mid(Date, 4, 2)
año = Mid(Date, 7, 4)
End Sub

'numToLet(Me.txtUserName.text, "PESO", "PESOS", "CENTAVO", "CENTAVOS", 0)
Function WORDNUM(ByVal numero As Variant, Optional TipoCambioSingular As String = "", Optional TipoCambioPlural As String, Optional subTipoCambioSingular As String, Optional subTipoCambioPlural As String, Optional xInternal As Long = 0) As String
     Dim snum As String, vNum() As String, X As Long, Y As Long, Z As Long, sTmp As String
     Dim D1 As String, D2 As String, D3 As String, D4 As String, DFinal As String
     Dim tNum As String, B1 As Boolean, B2 As Boolean, B3 As Boolean
     Dim wNum() As String, xNums As String, xWords As String, Nombres() As String
     
     '***********************************************************************************************
     '* Esta función convierte números en palabras, sin importar el contexto donde se encuentren    *
     '* La presición (por limitancia del lenguaje) es de 28B, Ej: 9999999999999999999999999999 max. *
     '***********************************************************************************************
     
     'Convierte el valor en un string
     snum = Trim(CStr(numero))
            
     'Procesa cada número que exista en la variable por separado
     If xInternal = 0 Then
        'Separa los números limpios de las palabras y los procesa por separado (no incluye números con letras)
        wNum = Split(snum, " ")
        For X = 0 To UBound(wNum)
            'Concatena los strings o números según corresponda
            If IsNumeric(wNum(X)) Then
               'Separa los enteros de los decimales para procesarlos por separado
               If Int(Val(wNum(X))) < wNum(X) Then
                  D1 = Int(Val(wNum(X)))
                  D2 = Mid(CStr(wNum(X)), Len(D1) + 2)
                  DFinal = DFinal & IIf(D1 < 0, "menos ", "") & WORDNUM(D1, TipoCambioSingular, TipoCambioPlural, 1) & " con "
                  DFinal = DFinal & WORDNUM(D2, subTipoCambioSingular, subTipoCambioPlural, , , 1) & " "
               Else
                  DFinal = DFinal & IIf(wNum(X) < 0, "menos ", "") & WORDNUM(wNum(X), TipoCambioSingular, TipoCambioPlural, subTipoCambioSingular, subTipoCambioPlural, 1) & " "
               End If
            Else
               DFinal = DFinal & wNum(X) & " "
            End If
        Next
     Else
        
        'ELimina el signo
        If Not IsNumeric(Left(snum, 1)) Then
           snum = Mid(snum, 2)
        End If
     
        'Elimina cualquier formato posible (incluye valores científicos)
        snum = Format(snum, "0")
        
        'Completa con ceros a la izquierda hasta obtener una longitud múltiplo de 3
        Do While Len(snum) Mod 3 <> 0
           snum = "0" & snum
        Loop
     
        'Dimenciona un arreglo con espacio para cada una de las centenas
        ReDim vNum(Len(snum) / 3 - 1)
        
        'Carga el arreglo con las centenas que corresponda
        For X = 0 To UBound(vNum, 1)
            vNum(X) = Mid(snum, (X + 1) * 3 - 2, 3)
        Next
         
        'Si el arreglo contiene una sola centena, la convierte en palabras
        If UBound(vNum, 1) = 0 Then
            'Asigna los dígitos de la centena y recuerda si son mayores que cero
            D3 = Left(snum, 1): B3 = Val(D3) > 0
            D2 = Mid(snum, 2, 1): B2 = Val(D2) > 0
            D1 = Right(snum, 1): B1 = Val(D1) > 0
            
            'Procesa las unidades
            Select Case D1
                   Case "1": DFinal = "un"
                   Case "2": DFinal = "dos"
                   Case "3": DFinal = "tres"
                   Case "4": DFinal = "cuatro"
                   Case "5": DFinal = "cinco"
                   Case "6": DFinal = "seis"
                   Case "7": DFinal = "siete"
                   Case "8": DFinal = "ocho"
                   Case "9": DFinal = "nueve"
            End Select
            
            'Procesa las decenas
            Select Case D2
                   Case "1"
                        'Maneja lógica del retrasado mental que puso nombres ilógicos a algunos números.
                        Select Case D1
                               Case "0": DFinal = "diez"
                               Case "1": DFinal = "once"
                               Case "2": DFinal = "doce"
                               Case "3": DFinal = "trece"
                               Case "4": DFinal = "catorce"
                               Case "5": DFinal = "quince"
                               Case "6": DFinal = "dieciséis"
                               Case Else
                                    DFinal = "dieci" & DFinal
                        End Select
                   Case "2"
                        If B1 Then
                           If D1 = "2" Then DFinal = "dós"
                           If D1 = "3" Then DFinal = "trés"
                           DFinal = "veinti" & DFinal
                        Else
                           DFinal = "veinte"
                        End If
                   Case "3": If B1 Then DFinal = "treinta y " & DFinal Else DFinal = "treinta"
                   Case "4": If B1 Then DFinal = "cuarenta y " & DFinal Else DFinal = "cuarenta"
                   Case "5": If B1 Then DFinal = "cincuenta y " & DFinal Else DFinal = "cincuenta"
                   Case "6": If B1 Then DFinal = "sesenta y " & DFinal Else DFinal = "sesenta"
                   Case "7": If B1 Then DFinal = "setenta y " & DFinal Else DFinal = "setenta"
                   Case "8": If B1 Then DFinal = "ochenta y " & DFinal Else DFinal = "ochenta"
                   Case "9": If B1 Then DFinal = "noventa y " & DFinal Else DFinal = "noventa"
            End Select
            
            'Procesa las centenas
            Select Case D3
                   Case "1": If B1 Or B2 Then DFinal = "ciento " & DFinal Else DFinal = "cien"
                   Case "2": If B1 Or B2 Then DFinal = "doscientos " & DFinal Else DFinal = "doscientos"
                   Case "3": If B1 Or B2 Then DFinal = "trescientos " & DFinal Else DFinal = "trescientos"
                   Case "4": If B1 Or B2 Then DFinal = "cuatrocientos " & DFinal Else DFinal = "cuatrocientos"
                   Case "5": If B1 Or B2 Then DFinal = "quinientos " & DFinal Else DFinal = "quinientos"
                   Case "6": If B1 Or B2 Then DFinal = "seiscientos " & DFinal Else DFinal = "seiscientos"
                   Case "7": If B1 Or B2 Then DFinal = "setecientos " & DFinal Else DFinal = "setecientos"
                   Case "8": If B1 Or B2 Then DFinal = "ochocientos " & DFinal Else DFinal = "ochocientos"
                   Case "9": If B1 Or B2 Then DFinal = "novecientos " & DFinal Else DFinal = "novecientos"
            End Select
            
            'Si es la ejecución principal efectua algunos arreglines
            If xInternal = 1 Then
               'Validación del cero
               If DFinal = "" Then DFinal = "cero"
               'Validación de terminados en "un"
               If Right(DFinal, 2) = "un" And TipoCambioSingular = "" Then DFinal = DFinal & "o"
            End If
            
        Else 'Si es más de una centena, las separa y procesa independientemente
            Y = -1
            Z = 1
            For X = UBound(vNum) To 0 Step -1
                Y = Y + 1
                
                'Convierte la centena en palabras
                tNum = WORDNUM(vNum(X), xInternal:=2)
                
                'Arregla la terminación "uno" cuando corresponde
                If Y = 0 And Right(tNum, 2) = "un" And TipoCambioSingular & TipoCambioPlural = "" Then tNum = tNum + "o"
                
                'Genera un valor temporal para poder modificar
                sTmp = tNum
                
                'Asigna los nombres genéricos principales
                Nombres = Split(" mil , millón , millones , billón , billones , trillón , trillones , cuatrillón , cuatrillones , quintillón , quintillones , sextillón , sextillones , septillón , septillones , octillón , octillones, nonillón , nonillones , decillón , decillones , undecillón , undecillones , duodecillón , duodecillones , tredecillón , tredecillones , cuatordecillón , cuatordecillones , quindecillón , quindecillones , sexdecillón , sexdecillones , septendecillón , septendecillones , octodecillón , octodecillones , novendecillón , novendecillones , vigintillón , vigintillones ", ",")
                
                'Controla que el índice de nombres no salga de los límites
                If Y > UBound(Nombres) Then
                   WORDNUM = "?"
                   Exit Function
                End If
                
                'Asigna los nombres correspondientes
                If Y Mod 2 > 0 Then
                   D1 = Nombres(0)
                   D2 = Nombres(Y - 1)
                ElseIf Y > 0 Then
                   D1 = Nombres(Y - 1)
                   D2 = Nombres(Y)
                Else
                   D1 = "": D2 = ""
                End If
                
                'Actualiza el nombre del número
                Select Case Y Mod 2
                       Case 0: If sTmp = "un" Then sTmp = sTmp & D1 Else sTmp = sTmp & IIf(tNum = "", "", D2)
                       Case Else
                            If sTmp = "un" Then sTmp = ""
                            sTmp = sTmp & IIf(tNum = "", "", D1)
                            If X = 0 And Y > 1 Then
                               If InStr(1, DFinal, D2, vbTextCompare) = 0 Then sTmp = sTmp & Mid(D2, 2)
                            End If
                End Select
                DFinal = sTmp & DFinal
            Next
        End If
     End If
     
     'Aplica el tipo de moneda cuando corresponda
     If xInternal = 1 Then DFinal = DFinal & " " & IIf(Format(snum, "#0") = "1", TipoCambioSingular, TipoCambioPlural)

     'Asigna el número en palabras
      WORDNUM = Trim(DFinal)
End Function
Public Function esNumeroDecimal(ByRef txt As String, ByVal Num As Long) As Long
    Dim numdec As Long
    numdec = Num
    Num = esNumero(Num)
    If Num = 0 Then
        If numdec = 46 Then '.
            If InStr(1, txt, ",", vbBinaryCompare) <> 0 Then
                esNumeroDecimal = 0
            Else
                esNumeroDecimal = 44    ',
            End If
        Else
            esNumeroDecimal = 0
        End If
    Else
        esNumeroDecimal = numdec
    End If
End Function
