VERSION 5.00
Begin VB.Form auxiliar03 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Mayor  Analitico"
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
      Caption         =   "MAYOR  ANALITICO"
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
Attribute VB_Name = "auxiliar03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
IMPRIMIR
End Sub

Private Sub Form_Load()
    
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)
fechasistema = Date
dia = Mid(Date, 1, 2)
mes = Mid(Date, 4, 2)
año = Mid(Date, 7, 4)
End Sub

Sub IMPRIMIR()
    informes.info.FontSize = 8
    informes.info.Clear
    largopagina = 65
    ' tituloinforme = "MAYOR ANALITICO   " + dia + "/" + mes + "/" + año
    titu(1) = "FECHA"
    titu(2) = "TP"
    titu(3) = "NUMERO"
    titu(4) = "NL"
    titu(5) = "TP"
    titu(6) = "NUMERO"
    titu(7) = "EMISION"
    titu(8) = "VENCTO"
    titu(9) = "GLOSA            "
    titu(10) = "DEBE     "
    titu(11) = "HABER    "
    titu(12) = "SALDO    "
    lineas = 70
    Consulta_Informe
    informes.Show
    
End Sub
Sub grilla()
    palabra = ""
    For K = 0 To cancolu
    If tipodato(K) = "s" Or tipodato(K) = "S" Then dato(K) = dato(K) & String(colu(K) - Len(dato(K)), 32)
    If tipodato(K) = "n" Or tipodato(K) = "N" Then dato(K) = String(colu(K) - Len(dato(K)), 32) & dato(K)
    palabra = palabra & dato(K)
    Next K
    If lineas > largopagina Then Call cabeza
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (" ")
    If Mid(dato(1), 7, 4) <> "0000" Then informes.info.AddItem ("    " + palabra)
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (Mid(palabra, 1, 40))
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (" ")
    lineas = lineas + 1
'    informes.info.AddItem palabr + palabra
End Sub
Sub cabeza()
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
    TITULOS = TITULOS & titu(K)

    Next K
    informes.info.AddItem ("    " + TITULOS)
    informes.info.AddItem String(132, "_")

lineas = 8
End Sub
Sub finalpag()
    colu(1) = 12: tipodato(1) = "s"
    colu(2) = 42: tipodato(2) = "s"
    colu(3) = 15: tipodato(3) = "n"
    colu(4) = 15: tipodato(4) = "n"
    colu(5) = 15: tipodato(5) = "n"
    colu(6) = 15: tipodato(6) = "n"
    colu(7) = 15: tipodato(7) = "n"
    cancolu = 7
    dato(1) = "SUB TOTALES "
    dato(2) = ""
    For K = 1 To 7
    Call comas(sumas(K - 1), 12)
    dato(K + 2) = VARIPASO
    Next K
    tipocue(0) = "F"
    grilla
    End Sub
    Sub diferencia()
    colu(1) = 12: tipodato(1) = "s"
    colu(2) = 42: tipodato(2) = "s"
    colu(3) = 15: tipodato(3) = "n"
    colu(4) = 15: tipodato(4) = "n"
    colu(5) = 15: tipodato(5) = "n"
    colu(6) = 15: tipodato(6) = "n"
    colu(7) = 15: tipodato(7) = "n"
    cancolu = 7
    dato(1) = "RESULTADO   "
    dato(2) = ""
    For K = 1 To 7
    Call comas(difer(K - 1), 12)
    dato(K + 2) = VARIPASO
    Next K
    tipocue(0) = "F"
    grilla
    End Sub
    Sub totalfinal()
    colu(1) = 12: tipodato(1) = "s"
    colu(2) = 42: tipodato(2) = "s"
    colu(3) = 15: tipodato(3) = "n"
    colu(4) = 15: tipodato(4) = "n"
    colu(5) = 15: tipodato(5) = "n"
    colu(6) = 15: tipodato(6) = "n"
    colu(7) = 15: tipodato(7) = "n"
    cancolu = 7
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
         
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            lineas = 70
             While Not resultados.EOF
               If Mid(resultados(0), 5, 4) <> "0000" Then
                tipocue(0) = resultados(0)
                tipocue(1) = resultados(1)
                LEERSALDOS (resultados(0))
                sw = 1
                LEEMOVIMIENTOS
              'If Mid(resultados(0), 5, 4) <> "0000" And sw = 1 Then FINCUENTA
              Call FINCUENTA
              End If
              
              resultados.MoveNext
              Wend
            resultados.Close
            Set resultados = Nothing
        End If


End Sub
Sub LEEMOVIMIENTOS()
Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT tipo,numero,linea,fecha,tipoctacte,rutctacte,codigocuenta,glosacontable,monto,dh,tipodocumento,numerodocumento,fechavencimiento,centrocosto,creadopor,mescontable,añocontable "
        cSql.SQL = cSql.SQL + "FROM movimientoscontables "
        cSql.SQL = cSql.SQL + "order by codigocuenta"
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
         While Not resultados.EOF
             
             If resultados(6) = tipocue(0) Then
             'If Mid(resultados(3), 4, 2) <> mes And Mid(resultados(3), 7, 4) <> año Then
             LEERSALDOS (resultados(6))
             If sw = 1 Then Call SALDOANTERIOR

             dato(0) = resultados(6)
             dato(1) = resultados(3)
             dato(2) = resultados(0)
             dato(3) = resultados(1)
             dato(4) = resultados(2)
             dato(5) = resultados(10)
             dato(6) = resultados(11)
             dato(7) = resultados(12)
             dato(8) = Mid(resultados(7), 1, 40)
             valor = resultados(8)

             If resultados(9) = "d" Or resultados(9) = "D" Then dato(9) = valor: dato(10) = 0
             If resultados(9) = "h" Or resultados(9) = "H" Then dato(10) = valor: dato(9) = 0
             dato(11) = anterior(0) + dato(9) - dato(10)

             Call total1
             For K = 1 To 3
             Call comas(dato(K + 8), 12)
             dato(K + 8) = VARIPASO
             Next K

             colu(1) = 12: tipodato(1) = "s"
             colu(2) = 2:  tipodato(2) = "s"
             colu(3) = 10: tipodato(3) = "s"
             colu(4) = 3: tipodato(4) = "s"
             colu(5) = 2: tipodato(5) = "s"
             colu(6) = 10: tipodato(6) = "s"
             colu(7) = 12: tipodato(7) = "s"
             colu(8) = 40: tipodato(8) = "s"
             colu(9) = 15: tipodato(9) = "n"
             colu(10) = 15: tipodato(10) = "n"
             colu(11) = 15: tipodato(11) = "n"
             cancolu = 11
             grilla
             End If
             'End If
             resultados.MoveNext
             Wend
             resultados.Close
            Set resultados = Nothing

        End If


End Sub
Sub SALDOANTERIOR()
    If lineas > 65 Then Call antecabeza: Call cabeza
    informes.info.AddItem (tipocue(0) + " " + tipocue(1) + "    SALDO ANTERIOR                    ----------------->" + Format(anterior(0), "###,###,###,###0"))
    lineas = lineas + 1
    sw = 2
    End Sub
    Sub antecabeza()
    titu(1) = "FECHA"
    titu(2) = "TP"
    titu(3) = "NUMERO"
    titu(4) = "NL"
    titu(5) = "TP"
    titu(6) = "NUMERO"
    titu(7) = "EMISION"
    titu(8) = "VENCTO"
    titu(9) = "GLOSA            "
    titu(10) = "DEBE     "
    titu(11) = "HABER    "
    titu(12) = "SALDO    "
    End Sub
    
Sub total1()
    sumas(0) = sumas(0) + dato(9)
    sumas(1) = sumas(1) + dato(10)
    sumas(2) = sumas(0) - sumas(1)

    End Sub
Sub FINCUENTA()
    informes.info.AddItem ("  SALDO  CUENTA                              " + Format(sumas(0), "##,###,###,###,###0") + "  " + Format(sumas(1), "##,###,###,###,###0") + "  " + Format(sumas(2), "##,###,###,###,###0"))
    informes.info.AddItem ("     ")
    lineas = lineas + 1: sw = 1
    sumas(0) = 0: sumas(1) = 0: sumas(2) = 0
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
    campos(25, 0) = "HABER10"
    campos(26, 0) = "HABER11"
    campos(27, 0) = "HABER12"
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
For K = 1 To mes - 1
anted = anted + SQLUTIL.datos(K + 3, 3)
anteh = anteh + SQLUTIL.datos(K + 15, 3)
Next K
anterior(0) = anted - anteh
End Sub



