VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form auxiliar01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Balance Tributario"
   ClientHeight    =   6345
   ClientLeft      =   240
   ClientTop       =   1290
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   855
      Left            =   1320
      TabIndex        =   10
      Top             =   4680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      BackColor       =   16761024
      Caption         =   "TIPO DE IMPRESION"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox FOLIO 
         Height          =   285
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton timbrado 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprime Timbrado"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton original 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprime Original"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin CoolButtons.cool_Button GENERA 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   5760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "GENERA INFORME"
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   4215
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7435
      BackColor       =   16761024
      Caption         =   "Configuracion"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   855
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "EMPRESA"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox DATO1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   3
            Text            =   "01"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label empresanombre 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   4
            Top             =   360
            Width           =   3255
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   855
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "MES"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox COMBOMES 
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   3855
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   1095
         Left            =   1320
         TabIndex        =   7
         Top             =   2280
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1931
         BackColor       =   16744576
         Caption         =   "AÑO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox COMBOAÑO 
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   3855
         End
      End
   End
End
Attribute VB_Name = "auxiliar01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private formatogrilla(10, 20) As String
Private sumas(10) As Double
Private suma(10) As Double
Private sumas2(10) As Double
Private sumas3(10) As Double

Private Sub command1_Click()
INFOGRILLA.Grid1.DefaultFont.Size = 8
For K = 1 To 11 - 1
INFOGRILLA.Grid1.Column(K).Width = Val(formatogrilla(2, K)) * INFOGRILLA.Grid1.DefaultFont.Size
Next K
INFOGRILLA.Grid1.PageSetup.Orientation = cellLandscape

INFOGRILLA.Grid1.PageSetup.PrintFixedRow = True


'INFOGRILLA.GRID1.PageSetup.BlackAndWhite = True
INFOGRILLA.Grid1.PageSetup.BottomMargin = 1
INFOGRILLA.Grid1.PageSetup.TopMargin = 1
INFOGRILLA.Grid1.PageSetup.LeftMargin = 1
INFOGRILLA.Grid1.PageSetup.RightMargin = 0






INFOGRILLA.Grid1.PrintPreview 75
End Sub

Private Sub Command2_Click()


INFOGRILLA.Grid1.DefaultFont.Size = 6

INFOGRILLA.Grid1.Column(1).Width = 0

For K = 2 To 11 - 1
        
        
        INFOGRILLA.Grid1.Column(K).Width = Val(formatogrilla(2, K)) * INFOGRILLA.Grid1.DefaultFont.Size
        
        
    Next K


'INFOGRILLA.GRID1.PageSetup.Orientation = cellLandscape
INFOGRILLA.Grid1.PageSetup.Orientation = cellPortrait



INFOGRILLA.Grid1.PageSetup.PrintFixedRow = True


'INFOGRILLA.GRID1.PageSetup.BlackAndWhite = True
INFOGRILLA.Grid1.PageSetup.BottomMargin = 1
INFOGRILLA.Grid1.PageSetup.TopMargin = 1
INFOGRILLA.Grid1.PageSetup.LeftMargin = 1
INFOGRILLA.Grid1.PageSetup.RightMargin = 0


CABEZA




INFOGRILLA.Grid1.PrintPreview 75

End Sub
Sub CABEZA()
Dim objReportTitle As FlexCell.ReportTitle
INFOGRILLA.Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Balance tributario"
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 20
    objReportTitle.PrintOnAllPages = True
    INFOGRILLA.Grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
    For K = 1 To 5
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = DATOSEMPRESA(K)
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Italic = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Color = RGB(128, 0, 0)
    objReportTitle.Align = CellLeft
    INFOGRILLA.Grid1.ReportTitles.Add objReportTitle
    Next K
With INFOGRILLA.Grid1.PageSetup
        
        .Footer = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        .FooterAlignment = cellRight
        .FooterFont.Name = "Verdana"
        .FooterFont.Size = 7
        .FooterMargin = 0.1
        
End With

End Sub
Private Sub Command3_Click()
INFOGRILLA.Grid1.ExportToExcel ("")
End Sub

Private Sub Command4_Click()
INFOGRILLA.Grid1.ExportToHTML ("")


End Sub


Private Sub Command5_Click()
INFOGRILLA.Grid1.DefaultFont.Size = INFOGRILLA.Grid1.DefaultFont.Size + 0.5
For K = 1 To INFOGRILLA.Grid1.Cols - 1
        INFOGRILLA.Grid1.Column(K).Width = Val(formatogrilla(2, K)) * INFOGRILLA.Grid1.DefaultFont.Size
        

Next K


INFOGRILLA.Grid1.Refresh



End Sub
Private Sub Command6_Click()
INFOGRILLA.Grid1.DefaultFont.Size = INFOGRILLA.Grid1.DefaultFont.Size - 0.5
For K = 1 To INFOGRILLA.Grid1.Cols - 1
        INFOGRILLA.Grid1.Column(K).Width = Val(formatogrilla(2, K)) * INFOGRILLA.Grid1.DefaultFont.Size
        

Next K


INFOGRILLA.Grid1.Refresh



End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)


End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(DATO1)
    
End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = "conta"
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestroempresas", DATO1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub


Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + DATO1.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then DATO1.SetFocus: GoTo no:
    COMBOMES.SetFocus
    empresanombre.Caption = SQLUTIL.datos(1, 3)
no:
End Sub

Private Sub Form_Load()
 Call Conectar_BD
 Call Conectarconta(servidor, "conta", USUARIO, password)

For K = 1 To 12
COMBOMES.AddItem MonthName(K)
Next K
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For K = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem K
Next K
COMBOAÑO.ListIndex = K - 2001
DATO1.text = empresaactiva
empresanombre.Caption = nombreempresa
original.Value = True



End Sub
Sub ACEPTA(OPCION)
Dim TIMBRA As String

Dim INFOGRILLA As grillainformes
Set INFOGRILLA = New grillainformes
If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"
If OPCION = 1 Then INFOGRILLA.Caption = "BALANCE TRIBUTARIO": grillainformes.Tag = "auxiliar01" & TIMBRA & FOLIO.text


Call CARGAGRILLA(INFOGRILLA)

Call CARGABALANCE(INFOGRILLA)

INFOGRILLA.Visible = True

INFOGRILLA.Show

End Sub




    Sub diferencia(INFOGRILLA As grillainformes, Row)
    INFOGRILLA.Grid1.Rows = Row + 1
     With INFOGRILLA.Grid1.Range(Row, 1, Row, 10)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
    
    INFOGRILLA.Grid1.Cell(Row, 2).text = "RESULTADOS"
   
    For K = 1 To 8
    INFOGRILLA.Grid1.Cell(Row, K + 2).text = difer(K - 1)
  
    Next K
    End Sub
    Sub totales(INFOGRILLA As grillainformes, Row)
    Dim dife As Double
    
    INFOGRILLA.Grid1.Rows = Row + 1
    
     With INFOGRILLA.Grid1.Range(Row, 1, Row, 10)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
    
    INFOGRILLA.Grid1.Cell(Row, 1).text = ""
    INFOGRILLA.Grid1.Cell(Row, 2).text = "TOTALES"
                 
    For K = 1 To 8
    INFOGRILLA.Grid1.Cell(Row, K + 2).text = sumas(K)
    sumas2(K) = 0
    Next K
    INFOGRILLA.Grid1.Rows = INFOGRILLA.Grid1.Rows + 1
    
     With INFOGRILLA.Grid1.Range(Row + 1, 1, Row + 1, 10)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
    
    dife = sumas(1) - sumas(2)
    If dife < 0 Then sumas2(1) = dife * -1 Else sumas2(2) = dife
    dife = sumas(3) - sumas(4)
    If dife < 0 Then sumas2(3) = dife * -1 Else sumas2(4) = dife
    dife = sumas(5) - sumas(6)
    If dife < 0 Then sumas2(5) = dife * -1 Else sumas2(6) = dife
    dife = sumas(7) - sumas(8)
    If dife < 0 Then sumas2(7) = dife * -1 Else sumas2(8) = dife
    
    INFOGRILLA.Grid1.Cell(Row + 1, 1).text = ""
    INFOGRILLA.Grid1.Cell(Row + 1, 2).text = "RESULTADOS EJERCICIOS"
     INFOGRILLA.Grid1.Rows = INFOGRILLA.Grid1.Rows + 1
     With INFOGRILLA.Grid1.Range(Row + 2, 1, Row + 2, 10)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
                 
                 
    For K = 1 To 8
    INFOGRILLA.Grid1.Cell(Row + 1, K + 2).text = sumas2(K)
    sumas3(K) = sumas2(K) + sumas(K)
    Next K
    
    INFOGRILLA.Grid1.Cell(Row + 2, 1).text = ""
    INFOGRILLA.Grid1.Cell(Row + 2, 2).text = "SUMAS IGUALES"
                 
    For K = 1 To 8
    INFOGRILLA.Grid1.Cell(Row + 2, K + 2).text = sumas3(K)
    
    Next K
    
    
    
    End Sub
    




Sub total()
    
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
Sub LEERSALDOS(LLAVE, TIPO)
Dim SUMD As Double
Dim SUMH As Double
Dim anted As Double
Dim anteh As Double
Dim dife As Double

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
    
    condicion = "codigo=" + "'" + LLAVE + "' and año ='" + Mid(fechasistema, 7, 4) + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = temporal
    
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
    anted = SQLUTIL.datos(2, 3)
    anteh = SQLUTIL.datos(3, 3)
    SUMD = 0: SUMH = 0
For K = 1 To Val(mes)
SUMD = SUMD + SQLUTIL.datos(K + 3, 3)
SUMH = SUMH + SQLUTIL.datos(K + 15, 3)
Next

For K = 1 To 8
suma(K) = 0
sumas2(K) = 0
sumas3(K) = 0
Next K

suma(1) = anted + SUMD
suma(2) = anteh + SUMH
dife = suma(1) - suma(2)

If dife > 0 Then suma(3) = dife
If dife < 0 Then suma(4) = dife * -1


If TIPO = "1" Or TIPO = "2" Then suma(5) = suma(3): suma(6) = suma(4)

If TIPO <> "1" And TIPO <> "2" Then suma(7) = suma(3): suma(8) = suma(4)
For K = 1 To 8
sumas(K) = sumas(K) + suma(K)
Next K

End Sub




Sub CARGAGRILLA(INFOGRILLA As grillainformes)
Rem DATOS DE LA COLUMNA
    
    
    formatogrilla(1, 1) = " CODIGO "
    formatogrilla(1, 2) = " CUENTA         "
    formatogrilla(1, 3) = "DEBITOS"
    formatogrilla(1, 4) = "CREDITOS"
    formatogrilla(1, 5) = "DEUDOR"
    formatogrilla(1, 6) = "ACREEDOR"
    formatogrilla(1, 7) = " ACTIVO"
    formatogrilla(1, 8) = "PASIVO"
    formatogrilla(1, 9) = "PERDIDA"
    formatogrilla(1, 10) = "GANANCIA"
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "9"
    formatogrilla(2, 2) = "30"
    formatogrilla(2, 3) = "12"
    formatogrilla(2, 4) = "12"
    formatogrilla(2, 5) = "11"
    formatogrilla(2, 6) = "11"
    formatogrilla(2, 7) = "11"
    formatogrilla(2, 8) = "11"
    formatogrilla(2, 9) = "11"
    formatogrilla(2, 10) = "11"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "N"
    formatogrilla(3, 4) = "N"
    formatogrilla(3, 5) = "N"
    formatogrilla(3, 6) = "N"
    formatogrilla(3, 7) = "N"
    formatogrilla(3, 8) = "N"
    formatogrilla(3, 9) = "N"
    formatogrilla(3, 10) = "N"
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = "###,###,###,###"
    formatogrilla(4, 4) = "###,###,###,###"
    formatogrilla(4, 5) = "###,###,###,###"
    formatogrilla(4, 6) = "###,###,###,###"
    formatogrilla(4, 7) = "###,###,###,###"
    formatogrilla(4, 8) = "###,###,###,###"
    formatogrilla(4, 9) = "###,###,###,###"
    formatogrilla(4, 10) = "###,###,###,###"
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    formatogrilla(5, 6) = "TRUE"
    formatogrilla(5, 7) = "TRUE"
    formatogrilla(5, 8) = "TRUE"
    formatogrilla(5, 9) = "TRUE"
    formatogrilla(5, 10) = "TRUE"
    
    
    INFOGRILLA.Grid1.Cols = 11
    INFOGRILLA.Grid1.Rows = 2
    
     'INFOGRILLA.GRID1.AllowUserResizing = False
    INFOGRILLA.Grid1.DisplayFocusRect = False
    'INFOGRILLA.GRID1.ExtendLastCol = True
    INFOGRILLA.Grid1.BoldFixedCell = False
    
    INFOGRILLA.Grid1.DrawMode = cellOwnerDraw
    
    INFOGRILLA.Grid1.Appearance = Flat
    INFOGRILLA.Grid1.ScrollBarStyle = Flat
    INFOGRILLA.Grid1.FixedRowColStyle = Flat
    
   'INFOGRILLA.GRID1.BackColorFixed = RGB(90, 158, 214)
   ' INFOGRILLA.GRID1.BackColorFixedSel = RGB(110, 180, 230)
   ' INFOGRILLA.GRID1.BackColorBkg = RGB(90, 158, 214)
   ' INFOGRILLA.GRID1.BackColorScrollBar = RGB(231, 235, 247)
   ' INFOGRILLA.GRID1.BackColor1 = RGB(231, 235, 247)
   ' INFOGRILLA.GRID1.BackColor2 = RGB(239, 243, 255)
   ' INFOGRILLA.GRID1.GridColor = RGB(148, 190, 231)
    INFOGRILLA.Grid1.Column(0).Width = 0
    
    For K = 1 To INFOGRILLA.Grid1.Cols - 1
        
        INFOGRILLA.Grid1.Cell(0, K).text = formatogrilla(1, K)
        INFOGRILLA.Grid1.Column(K).Width = Val(formatogrilla(2, K)) * INFOGRILLA.Grid1.DefaultFont.Size
        
        
        
        INFOGRILLA.Grid1.Column(K).MaxLength = Val(formatogrilla(2, K))
        INFOGRILLA.Grid1.Column(K).FormatString = formatogrilla(4, K)
        INFOGRILLA.Grid1.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then INFOGRILLA.Grid1.Column(K).Alignment = cellRightCenter
        If formatogrilla(3, K) = "D" Then INFOGRILLA.Grid1.Column(K).CellType = cellCalendar
        
    Next K
End Sub
   

Sub CARGABALANCE(INFOGRILLA As grillainformes)
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim LIN As Double
    With informes
        Set cSql.ActiveConnection = temporal
        cSql.SQL = "SELECT codigo,nombre,tipo "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor "
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        LIN = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
             barra.Max = cSql.RowsAffected + 1
             barra.Value = 0
                While Not resultados.EOF
                    If Mid(resultados(0), 5, 4) <> "0000" Then
                    Call LEERSALDOS(resultados(0), resultados(2))
                            If suma(0) + suma(1) <> 0 Then
                            LIN = LIN + 1
                            INFOGRILLA.Grid1.Rows = LIN + 1
                            
                            barra.Value = barra.Value + 1
                            INFOGRILLA.Grid1.Cell(LIN, 1).text = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4)
                            INFOGRILLA.Grid1.Cell(LIN, 2).text = resultados(1)
                            For K = 1 To 8
                            INFOGRILLA.Grid1.Cell(LIN, K + 2).text = suma(K)
                            Next K
                            End If
                    End If
                resultados.MoveNext
                Wend
            Call totales(INFOGRILLA, INFOGRILLA.Grid1.Rows)
            resultados.Close
            
            Set resultados = Nothing

        End If
    End With


End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Private Sub NOIMPRIME_Click()
Unload Me
End Sub

Private Sub GENERA_Click()
Call Conectartemporal(servidor, "conta" + DATO1.text, USUARIO, password)
año = COMBOAÑO.text
mes = COMBOMES.ListIndex + 1
If Val(mes) < 10 Then mes = "0" + Mid(Str(mes), 2, 1) Else mes = Mid(Str(mes), 2, 2)

Call ACEPTA(1)
Unload Me

End Sub
