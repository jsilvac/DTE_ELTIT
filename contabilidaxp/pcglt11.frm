VERSION 5.00
Begin VB.Form pcglt11 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Balance Tributario"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "BALANCE TRIBUTARIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "pcglt11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
IMPRIMIR
End Sub

Private Sub Form_Load()
    
    Call Conectar_BD
    Call Conectarconta00("eltitxp", "conta00", "root", "123")
fechasistema = Date
dia = Mid(Date, 1, 2)
mes = Mid(Date, 4, 2)
año = Mid(Date, 7, 4)
End Sub

Sub IMPRIMIR()
    informes.info.FontSize = 8
    informes.info.Clear
    largopagina = 65
    tituloinforme = "BALANCE TRIBUTARIO   " + dia + "/" + mes + "/" + año
    
    titu1(1) = "             "
    titu1(2) = "              "
    titu1(3) = "    T  O  T   "
    titu1(4) = " A  L  E  S   "
    titu1(5) = "    S  A  L   "
    titu1(6) = " D  O  S      "
    titu1(7) = "  I N V E N T "
    titu1(8) = " T A R I O    "
    titu1(9) = "  R E S U L T "
    titu1(10) = " A D O       "
    titu(1) = "CUENTA"
    titu(2) = " NOMBRE DE CUENTA  "
    titu(3) = "  DEBITOS     "
    titu(4) = "  CREDITOS    "
    titu(5) = "  DEUDOR      "
    titu(6) = "  ACREEDOR    "
    titu(7) = "  ACTIVO      "
    titu(8) = "  PASIVO      "
    titu(9) = "  PERDIDA     "
    titu(10) = "  GANANCIA    "
    
    lineas = 70
    Consulta_Informe
    informes.Show
    
End Sub
Sub grilla()
    palabra = ""
    For K = 1 To cancolu
    If tipodato(K) = "s" Or tipodato(K) = "S" Then dato(K) = dato(K) & String(colu(K) - Len(dato(K)), 32)
    If tipodato(K) = "n" Or tipodato(K) = "N" Then dato(K) = String(colu(K) - Len(dato(K)), 32) & dato(K)
    palabra = palabra & dato(K)
    Next K
    If lineas > largopagina Then Call cabeza
    lineas = lineas + 1
    informes.info.AddItem palabr + palabra
    tipocue(0) = ""
End Sub
Sub cabeza()
    If lineas = largopagina Then Call finalpag
    informes.info.AddItem ("")
    informes.info.AddItem ("")
    pagina = pagina + 1
    informes.info.AddItem ("NOMBRE EMPRESA          " + tituloinforme + "                                   PAGINA " + Str$(pagina))
    informes.info.AddItem ("DIRECCION EMPRESA                                                                              " + Mid(Date$, 4, 2) + "-" + Mid(Date$, 1, 2) + "-" + Mid(Date$, 7, 4))
    informes.info.AddItem ("RUT EMPRESA                                                                                    " + Time$)
    informes.info.AddItem ("                                " + tituloinforme)
    informes.info.AddItem String(132, "_")
    TITULOS = ""
    For K = 1 To cancolu
    TITULOS1 = TITULOS1 & titu1(K)
    Next K
    informes.info.AddItem ("    " + TITULOS1)
    For K = 1 To cancolu
    TITULOS = TITULOS & titu(K)
    Next K
    informes.info.AddItem ("    " + TITULOS)
    informes.info.AddItem String(132, "_")

lineas = 8
End Sub
Sub finalpag()
    colu(1) = 12: tipodato(1) = "s"
    colu(2) = 10: tipodato(2) = "s"
    colu(3) = 15: tipodato(3) = "n"
    colu(4) = 15: tipodato(4) = "n"
    colu(5) = 15: tipodato(5) = "n"
    colu(6) = 15: tipodato(6) = "n"
    colu(7) = 15: tipodato(7) = "n"
    colu(8) = 15: tipodato(8) = "n"
    colu(9) = 15: tipodato(9) = "n"
    colu(10) = 15: tipodato(10) = "n"
    cancolu = 10
    dato(1) = "SUB TOTALES "
    dato(2) = ""
    For K = 1 To 8
    Call comas(sumas(K - 1), 12)
    dato(K + 2) = VARIPASO
    Next K
    tipocue(0) = "F"
    grilla
    End Sub
    Sub diferencia()
    colu(1) = 12: tipodato(1) = "s"
    colu(2) = 10: tipodato(2) = "s"
    colu(3) = 15: tipodato(3) = "n"
    colu(4) = 15: tipodato(4) = "n"
    colu(5) = 15: tipodato(5) = "n"
    colu(6) = 15: tipodato(6) = "n"
    colu(7) = 15: tipodato(7) = "n"
    colu(8) = 15: tipodato(8) = "n"
    colu(9) = 15: tipodato(9) = "n"
    colu(10) = 15: tipodato(10) = "n"
    cancolu = 10
    dato(1) = "RESULTADO   "
    dato(2) = ""
    For K = 1 To 8
    Call comas(difer(K - 1), 12)
    dato(K + 2) = VARIPASO
    Next K
    tipocue(0) = "F"
    grilla
    End Sub
    Sub totalfinal()
    colu(1) = 12: tipodato(1) = "s"
    colu(2) = 10: tipodato(2) = "s"
    colu(3) = 15: tipodato(3) = "n"
    colu(4) = 15: tipodato(4) = "n"
    colu(5) = 15: tipodato(5) = "n"
    colu(6) = 15: tipodato(6) = "n"
    colu(7) = 15: tipodato(7) = "n"
    colu(8) = 15: tipodato(8) = "n"
    colu(9) = 15: tipodato(9) = "n"
    colu(10) = 15: tipodato(10) = "n"
    cancolu = 10
    dato(1) = "TOTALES     "
    dato(2) = ""
    For K = 1 To 8
    Call comas(sumast(K - 1), 12)
    dato(K + 2) = VARIPASO
    Next K
    grilla
    informes.info.AddItem String(132, "_")
    End Sub
    




Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
             While Not resultados.EOF
                If Mid(resultados(0), 5, 4) <> "0000" Then
                tipocue(0) = resultados(2)
                LEERSALDOS (resultados(0))
                dato(1) = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4)
                dato(2) = Mid(resultados(1), 1, 10)
                If suma(0) + suma(1) <> 0 Then
                For K = 1 To 8
                Call comas(suma(K - 1), 12)
                dato(K + 2) = VARIPASO
                Next K
                colu(1) = 12: tipodato(1) = "s"
                colu(2) = 10: tipodato(2) = "s"
                colu(3) = 15: tipodato(3) = "n"
                colu(4) = 15: tipodato(4) = "n"
                colu(5) = 15: tipodato(5) = "n"
                colu(6) = 15: tipodato(6) = "n"
                colu(7) = 15: tipodato(7) = "n"
                colu(8) = 15: tipodato(8) = "n"
                colu(9) = 15: tipodato(9) = "n"
                colu(10) = 15: tipodato(10) = "n"
                cancolu = 10
                grilla
                Call total
                End If
                End If
            resultados.MoveNext
            Wend
            Call total1
            informes.info.AddItem String(132, "_")
            
            Call finalpag
            Call diferencia
            Call totalfinal
            resultados.Close
            
            Set resultados = Nothing

        End If
    End With

End Sub
Sub total()
    sumas(0) = sumas(0) + suma(0)
    sumas(1) = sumas(1) + suma(1)
    sumas(2) = sumas(2) + suma(2)
    sumas(3) = sumas(3) + suma(3)
    sumas(4) = sumas(4) + suma(4)
    sumas(5) = sumas(5) + suma(5)
    sumas(6) = sumas(6) + suma(6)
    sumas(7) = sumas(7) + suma(7)
    
End Sub
Sub total1()
    difer(0) = 0: difer(1) = 0: difer(2) = 0: difer(3) = 0
    If sumas(5) > sumas(4) Then difer(4) = sumas(5) - sumas(4): difer(5) = 0
    If sumas(4) > sumas(5) Then difer(5) = sumas(4) - sumas(5): difer(4) = 0
    
    If sumas(7) > sumas(6) Then difer(6) = sumas(7) - sumas(6): difer(7) = 0
    If sumas(6) > sumas(7) Then difer(7) = sumas(6) - sumas(7): difer(6) = 0
    
    
    sumast(0) = sumas(0) + difer(0)
    sumast(1) = sumas(1) + difer(1)
    sumast(2) = sumas(2) + difer(2)
    sumast(3) = sumas(3) + difer(3)
    sumast(4) = sumas(4) + difer(4)
    sumast(5) = sumas(5) + difer(5)
    sumast(6) = sumas(6) + difer(6)
    sumast(7) = sumas(7) + difer(7)

    suma(0) = 0: suma(1) = 0: suma(2) = 0: suma(3) = 0: suma(4) = 0: suma(5) = 0: suma(6) = 0: suma(7) = 0
    
                
End Sub
Sub LEERSALDOS(LLAVE)
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = "debe01"
    campos(5, 0) = "debe02"
    campos(6, 0) = "debe03"
    campos(7, 0) = "debe04"
    campos(8, 0) = "debe05"
    campos(9, 0) = "debe06"
    campos(10, 0) = "debe07"
    campos(11, 0) = "debe08"
    campos(12, 0) = "debe09"
    campos(13, 0) = "debe10"
    campos(14, 0) = "debe11"
    campos(15, 0) = "debe12"
    campos(16, 0) = "haber01"
    campos(17, 0) = "haber02"
    campos(18, 0) = "haber03"
    campos(19, 0) = "haber04"
    campos(20, 0) = "haber05"
    campos(21, 0) = "haber06"
    campos(22, 0) = "haber07"
    campos(23, 0) = "haber08"
    campos(24, 0) = "haber09"
    campos(25, 0) = "haber10"
    campos(26, 0) = "haber11"
    campos(27, 0) = "haber12"
    campos(28, 0) = ""
    
    condicion = "codigo=" + "'" + LLAVE + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
anted = SQLUTIL.datos(2, 3)
anteh = SQLUTIL.datos(3, 3)
sumd = 0: sumh = 0
For K = 1 To 12
sumd = sumd + SQLUTIL.datos(K + 4, 3)
sumh = sumh + SQLUTIL.datos(K + 15, 3)
Next
suma(0) = anted + sumd
suma(1) = anteh + sumh
tipo = tipocue(0)
suma(2) = 0: suma(3) = 0: suma(4) = 0: suma(5) = 0: suma(6) = 0: suma(7) = 0
If suma(0) > suma(1) Then suma(2) = suma(0) - suma(1): suma(3) = 0
If suma(0) < suma(1) Then suma(3) = suma(1) - suma(0): suma(2) = 0

If suma(0) > suma(1) And tipo = "1" Or tipo = "2" Then suma(4) = suma(2): suma(5) = 0
If suma(0) < suma(1) And tipo = "1" Or tipo = "2" Then suma(5) = suma(3): suma(4) = 0

If suma(0) > suma(1) And tipo = "3" Then suma(6) = suma(2): suma(7) = 0
If suma(0) < suma(1) And tipo = "3" Then suma(7) = suma(3): suma(6) = 0
End Sub



