VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form form29 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTA CERTIFICADOS DE RENTA"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16230
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1082
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   12840
      TabIndex        =   17
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      BackColor       =   16744576
      Caption         =   " Mis Datos"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8865
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   -90
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8925
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   16185
      _ExtentX        =   28549
      _ExtentY        =   15743
      BackColor       =   16744576
      Caption         =   "INFORME CERTIFICADOS DE RENTA"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA SII"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8370
         Width           =   1365
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF8080&
         Caption         =   "Vista Previa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   14
         Top             =   8370
         Width           =   1365
      End
      Begin VB.TextBox FIRMA 
         Height          =   330
         Left            =   5355
         TabIndex        =   10
         Top             =   8460
         Width           =   5460
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "TODOS"
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
         Left            =   12015
         TabIndex        =   9
         Top             =   8235
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8370
         Width           =   1365
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   4
         Top             =   360
         Width           =   15960
         _ExtentX        =   28152
         _ExtentY        =   1852
         BackColor       =   16744576
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin VB.OptionButton opt2 
            BackColor       =   &H00FF8080&
            Caption         =   "SEGUNDA HOJA"
            Height          =   375
            Left            =   5400
            TabIndex        =   21
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00FF8080&
            Caption         =   "PRIMERA HOJA"
            Height          =   375
            Left            =   3600
            TabIndex        =   20
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton Command4 
            Caption         =   "IMPRIME PLANILLA"
            Height          =   495
            Left            =   9360
            TabIndex        =   16
            Top             =   360
            Width           =   3975
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "GENERAR INFORME"
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
            Left            =   13680
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   2055
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   135
            TabIndex        =   7
            Top             =   315
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1191
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
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   270
               Width           =   2865
            End
         End
      End
      Begin XPFrame.FrameXp frameprimerahoja 
         Height          =   6675
         Left            =   0
         TabIndex        =   3
         Top             =   1560
         Width           =   16125
         _ExtentX        =   28443
         _ExtentY        =   11774
         BackColor       =   16744576
         Caption         =   "LISTADO DE FACTURAS DE VENTA EMITIDAS"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin FlexCell.Grid GRID1 
            Height          =   6330
            Left            =   0
            TabIndex        =   12
            Top             =   360
            Width           =   16095
            _ExtentX        =   28390
            _ExtentY        =   11165
            Cols            =   5
            DefaultFontName =   "Arial"
            DefaultFontSize =   8.25
            FixedRowColStyle=   0
            Rows            =   30
         End
      End
      Begin FlexCell.Grid Grid2 
         Height          =   165
         Left            =   10395
         TabIndex        =   13
         Top             =   8730
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   291
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         FixedRowColStyle=   0
         Rows            =   30
         SelectionMode   =   1
      End
      Begin XPFrame.FrameXp frmsegundahoja 
         Height          =   6675
         Left            =   0
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   16125
         _ExtentX        =   28443
         _ExtentY        =   11774
         BackColor       =   16744576
         Caption         =   "LISTADO DE FACTURAS DE VENTA EMITIDAS"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin FlexCell.Grid Grid3 
            Height          =   6330
            Left            =   0
            TabIndex        =   23
            Top             =   360
            Width           =   16095
            _ExtentX        =   28390
            _ExtentY        =   11165
            Cols            =   5
            DefaultFontName =   "Arial"
            DefaultFontSize =   8.25
            FixedRowColStyle=   0
            Rows            =   30
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUT Y NOMBRE REPRESENTANTE LEGAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5355
         TabIndex        =   11
         Top             =   8190
         Width           =   5505
      End
   End
End
Attribute VB_Name = "form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String
Private COSTO1 As Double
Private COSTO2 As Double
Private COSTO3 As Double
Private COSTO10 As Double
Private COSTO20 As Double
Private COSTO30 As Double
Private rea1 As Double
Private rea2 As Double
Private salud As Double
Private previ As Double
Private sincorre1 As Double
Private sincorre2 As Double
Private sincorre3 As Double
Private totalsincorre1 As Double
Private totalsincorre2 As Double
Private totalsincorre3 As Double
Private totalsincorre4 As Double
Private totalsincorre5 As Double
Private totalsincorre6 As Double
Private totalsincorre7 As Double
Private posy As Long
Private posx As Long
Private hoja As Double






Private Sub Check1_Click()
For k = 1 To Grid1.Rows - 2
If Check1.Value = "0" Then
Grid1.Cell(k, 22).text = "0"
Else
Grid1.Cell(k, 22).text = "1"
End If

Next k

End Sub

Private Sub Command1_Click()
Dim s As Integer

CARGAGRILLA2

For s = 1 To Grid1.Rows - 2
If Grid1.Cell(s, 22).text = "1" Then
    Call leercertificado(Grid1.Cell(s, 1).text, Grid1.Cell(s, 2).text, Grid1.Cell(s, 21).text)
    Call IMPRIMIR2(Grid1.Cell(s, 1).text, Grid1.Cell(s, 2).text, Grid1.Cell(s, 21).text)
End If

Next s


End Sub






Private Sub Command3_Click()
Dim D1 As String
Dim D2 As String
Dim D3 As String
Dim D4 As String
Dim D5 As String
Dim D6 As String
Dim D7 As String
Dim D8 As String
Dim D9 As String
Dim D10 As String
Dim D11 As String
Dim D12 As String
Dim D13 As String
Dim D14 As String
Dim D15 As String
Dim D16 As String
Dim D17 As String
Dim D18 As String
Dim D19 As String
Dim D20 As String

Close 10

Open "F1887_" + empresaactiva + ".TXT" For Output As #10
For k = 2 To Grid1.Rows - 8
    D1 = CDbl(Mid(Grid1.Cell(k, 1).text, 2, 8)) & Mid(Grid1.Cell(k, 1).text, 10, 1)
    D2 = Val(Format(Grid1.Cell(k, 3).text, "000000000000"))
    D3 = Val(Format(Grid1.Cell(k, 4).text, "000000000000"))
    D4 = Val(Format(Grid1.Cell(k, 5).text, "000000000000"))
    D5 = Val(Format(Grid1.Cell(k, 6).text, "000000000000"))
    D6 = Val(Format(Grid1.Cell(k, 7).text, "000000000000"))
    D7 = Val(Format(Grid1.Cell(k, 8).text, "000000000000"))
    D8 = Format(Grid1.Cell(k, 9).text, "")
    D9 = Format(Grid1.Cell(k, 10).text, "")
    D10 = Format(Grid1.Cell(k, 11).text, "")
    D11 = Format(Grid1.Cell(k, 12).text, "")
    D12 = Format(Grid1.Cell(k, 13).text, "")
    D13 = Format(Grid1.Cell(k, 14).text, "")
    D14 = Format(Grid1.Cell(k, 15).text, "")
    D15 = Format(Grid1.Cell(k, 16).text, "")
    D16 = Format(Grid1.Cell(k, 17).text, "")
    D17 = Format(Grid1.Cell(k, 18).text, "")
    D18 = Format(Grid1.Cell(k, 19).text, "")
    D19 = Format(Grid1.Cell(k, 20).text, "")
    D20 = Format(Grid1.Cell(k, 21).text, "0000000")
    
    
    Print #10, D1 + ";" + D2 + ";" + D3 + ";" + D4 + ";" + D5 + ";" + D6 + ";" + D7 + ";" + D8 + ";" + D9 + ";" + D10 + ";" + D11 + ";" + D12 + ";" + D13 + ";" + D14 + ";" + D15 + ";" + D16 + ";" + D17 + ";" + D18 + ";" + D19 + ";" + D20
Next k
Close #10
Shell "NOTEPAD " + "F1887_" + empresaactiva + ".TXT"




End Sub


Private Sub Command4_Click()
'Call cabezas4("INFORME CERTIFICADO 1887", "N", 0)
 If opt1.Value = True Then
    Grid1.PageSetup.BottomMargin = 1
    Grid1.PageSetup.TopMargin = 1
    Grid1.PageSetup.LeftMargin = 1
    Grid1.PageSetup.RightMargin = 1
    Grid1.PageSetup.PrintGridlines = True
    
    Grid1.PageSetup.Orientation = cellLandscape
    Grid1.PrintPreview
 End If
 
 If opt2.Value = True Then
    Grid3.PageSetup.BottomMargin = 1
    Grid3.PageSetup.TopMargin = 1
    Grid3.PageSetup.LeftMargin = 0.5
    Grid3.PageSetup.RightMargin = 1
    Grid3.PageSetup.PrintGridlines = True
    Grid3.PageSetup.Orientation = cellLandscape
    Grid3.PrintPreview
 End If

End Sub

Public Sub COMMAND2_Click()
    leer
End Sub
Private Sub FIRMA_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
CENTRAR Me
'    Call Conectar_BD

    sc = 0
CARGAGRILLA
CARGAGRILLA2
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001


End Sub








Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub




Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub


Private Sub lblhistorico_Click(Index As Integer)

End Sub




Private Sub Label16_Click()
End Sub

Sub limpia()
    
    
End Sub

Sub imprimir()
Dim titulo As String
Call CABEZAS2(titulo, "N", "000000000")
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub
Sub IMPRIMIR2(rut, NOMBRE, numero)
Dim titulo As String
Call cabezas3(rut, NOMBRE, numero)
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThin
Grid2.DefaultFont.Size = 7
Grid2.PageSetup.Orientation = cellLandscape

Grid2.PageSetup.PrintFixedRow = True
Grid2.PageSetup.BottomMargin = 1
Grid2.PageSetup.TopMargin = 0.5
Grid2.PageSetup.LeftMargin = 0.5
Grid2.PageSetup.RightMargin = 0.5
Grid2.PageSetup.BlackAndWhite = True
Grid2.PageSetup.PrintGridlines = False

Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThin

If Check2.Value = "1" Then
    Grid2.PrintPreview 100
Else
    Grid2.DirectPrint
End If

   
End Sub


Sub grilla()
    
End Sub




Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
    Dim FORMATOGRILLA(10, 30)
    Grid1.DefaultFont.Size = 8
     Grid1.Rows = 89
     Grid1.Cols = 26
       
    FORMATOGRILLA(1, 1) = ""
    FORMATOGRILLA(1, 2) = ""
    FORMATOGRILLA(1, 3) = ""
    FORMATOGRILLA(1, 4) = ""
    FORMATOGRILLA(1, 5) = ""
    FORMATOGRILLA(1, 6) = ""
    FORMATOGRILLA(1, 7) = ""
    FORMATOGRILLA(1, 8) = ""
    FORMATOGRILLA(1, 9) = ""
    FORMATOGRILLA(1, 10) = ""
    FORMATOGRILLA(1, 11) = ""
    FORMATOGRILLA(1, 12) = ""
    FORMATOGRILLA(1, 13) = ""
    FORMATOGRILLA(1, 14) = ""
    FORMATOGRILLA(1, 15) = ""
    FORMATOGRILLA(1, 16) = ""
    FORMATOGRILLA(1, 17) = ""
    FORMATOGRILLA(1, 18) = ""
    FORMATOGRILLA(1, 20) = ""
    FORMATOGRILLA(1, 21) = ""
    FORMATOGRILLA(1, 22) = ""
    FORMATOGRILLA(1, 23) = ""
    FORMATOGRILLA(1, 24) = ""
    FORMATOGRILLA(1, 25) = ""
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "2"
    FORMATOGRILLA(2, 2) = "3"
    FORMATOGRILLA(2, 3) = "2"
    FORMATOGRILLA(2, 4) = "2"
    FORMATOGRILLA(2, 5) = "4"
    FORMATOGRILLA(2, 6) = "12"
    FORMATOGRILLA(2, 7) = "4"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "4"
    FORMATOGRILLA(2, 10) = "5"
    FORMATOGRILLA(2, 11) = "4"
    FORMATOGRILLA(2, 12) = "4"
    FORMATOGRILLA(2, 13) = "10"
    FORMATOGRILLA(2, 14) = "4"
    FORMATOGRILLA(2, 15) = "4"
    FORMATOGRILLA(2, 16) = "4"
    FORMATOGRILLA(2, 17) = "4"
    FORMATOGRILLA(2, 18) = "4"
    FORMATOGRILLA(2, 19) = "4"
    FORMATOGRILLA(2, 20) = "4"
    FORMATOGRILLA(2, 21) = "4"
    FORMATOGRILLA(2, 22) = "9"
    FORMATOGRILLA(2, 23) = "3"
    FORMATOGRILLA(2, 24) = "15"
    FORMATOGRILLA(2, 25) = "2"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "N"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "N"
    FORMATOGRILLA(3, 15) = "N"
    FORMATOGRILLA(3, 16) = "N"
    FORMATOGRILLA(3, 17) = "N"
    FORMATOGRILLA(3, 18) = "N"
    FORMATOGRILLA(3, 19) = "N"
    FORMATOGRILLA(3, 20) = "N"
    FORMATOGRILLA(3, 21) = "N"
    FORMATOGRILLA(3, 22) = "N"
    FORMATOGRILLA(3, 23) = "N"
    FORMATOGRILLA(3, 24) = "N"
    FORMATOGRILLA(3, 25) = "C"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = ""
    FORMATOGRILLA(4, 17) = ""
    
    
    Rem LOCCKED
    For k = 1 To 25
        FORMATOGRILLA(5, k) = "FALSE"
    Next k
    
    
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "C" Then Grid1.Column(k).Alignment = cellCenterGeneral
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
    
    Call primercuadro
   
    ' SEGUNDO CUADRO
    Call segundocuadro
    Call tercercuadro
      
   
   
   Grid1.Refresh
    
    
End Sub
Sub primercuadro()
    Grid1.Range(1, 2, 1, 15).Merge
    Grid1.Range(1, 2, 1, 15).Alignment = cellCenterGeneral
    Grid1.Cell(1, 2).text = "IMPUESTO AL VALOR AGREGADO D.L. 825/74"
    Grid1.Cell(1, 2).Font.Bold = True
    Grid1.Range(1, 16, 1, 22).Merge
    Grid1.Cell(1, 16).Alignment = cellCenterGeneral
    Grid1.Cell(1, 16).text = " Cantidad de documentos"
    Grid1.Cell(1, 16).Font.Bold = True
    Grid1.Range(1, 23, 1, 25).Merge
    Grid1.Cell(1, 23).Alignment = cellCenterGeneral
    Grid1.Cell(1, 23).text = "Monto Neto"
    Grid1.Cell(1, 23).Font.Bold = True
    
    Grid1.Cell(2, 2).text = "1"
    Grid1.Range(2, 5, 2, 15).Merge
    Grid1.Cell(2, 5).text = "Exportaciones"
    Grid1.Cell(2, 16).text = "585"
    Grid1.Cell(2, 23).text = "20"
    
    Grid1.Range(2, 17, 2, 22).Merge
    Grid1.Range(2, 24, 2, 25).Merge
    
    Grid1.Cell(3, 2).text = "2"
    Grid1.Range(3, 5, 3, 15).Merge
    Grid1.Cell(3, 5).text = " Ventas y/o Servicios orestadis Exentos o No Gravados del giro"
    Grid1.Cell(3, 16).text = "586"
    Grid1.Cell(3, 23).text = "142"
    Grid1.Range(3, 17, 3, 22).Merge
    Grid1.Range(3, 24, 3, 25).Merge
    
    
    Grid1.Cell(4, 2).text = "3"
    Grid1.Range(4, 5, 4, 15).Merge
    Grid1.Cell(4, 5).text = " Ventas con retención sobre el margen de comercialización (contribuyentes retenidos)"
    Grid1.Cell(4, 16).text = "731"
    Grid1.Cell(4, 23).text = "732"
    Grid1.Range(4, 17, 4, 22).Merge
    Grid1.Range(4, 24, 4, 25).Merge
    
    
    Grid1.Cell(5, 2).text = "4"
    Grid1.Range(5, 5, 5, 15).Merge
    Grid1.Cell(5, 5).text = " Ventas y/o Servicios prestados exentos o No Gravados que no son del giro"
    Grid1.Cell(5, 16).text = "714"
    Grid1.Cell(5, 23).text = "715"
    Grid1.Range(5, 17, 5, 22).Merge
    Grid1.Range(5, 24, 5, 25).Merge
    
    
    Grid1.Cell(6, 2).text = "5"
    Grid1.Range(6, 5, 6, 15).Merge
    Grid1.Cell(6, 5).text = " Facturas de Compra recibidas con retención total (contribuyentes retenidos) y Factura de Inicio emitida"
    Grid1.Cell(6, 16).text = "515"
    Grid1.Cell(6, 23).text = "587"
    Grid1.Range(6, 17, 6, 22).Merge
    Grid1.Range(6, 24, 6, 25).Merge
    
    
    Grid1.Cell(7, 2).text = "6"
    Grid1.Range(7, 5, 7, 22).Merge
    Grid1.Cell(7, 5).text = " Facturas de compra recibidas con retención parcial (Total neto según línea N° 14)"
    Grid1.Cell(7, 23).text = "720"
    Grid1.Range(7, 24, 7, 25).Merge
    
     
    Grid1.Range(8, 5, 8, 15).Merge
    Grid1.Range(8, 16, 8, 22).Merge
    Grid1.Cell(8, 16).text = "Cantidad de documentos"
    Grid1.Cell(8, 16).Font.Bold = True
    Grid1.Cell(8, 16).Alignment = cellCenterGeneral
    Grid1.Range(8, 23, 8, 25).Merge
    Grid1.Cell(8, 23).text = "Débito"
    Grid1.Cell(8, 23).Font.Bold = True
    Grid1.Cell(8, 23).Alignment = cellCenterGeneral
    
    
    Grid1.Cell(9, 2).text = "7"
    Grid1.Range(9, 5, 9, 15).Merge
    Grid1.Cell(9, 5).text = " Facturas emitidas por ventas y servicios del giro"
    Grid1.Cell(9, 16).text = "503"
    Grid1.Cell(9, 23).text = "502"
    Grid1.Cell(9, 25).text = "+"
    Grid1.Range(9, 17, 9, 22).Merge
   
    
    Grid1.Cell(10, 2).text = "8"
    Grid1.Range(10, 5, 10, 15).Merge
    Grid1.Cell(10, 5).text = " Facturas y Notas de Débitos por ventas y servicios que no son del giro (activo fijo y otros)"
    Grid1.Cell(10, 16).text = "716"
    Grid1.Cell(10, 23).text = "717"
    Grid1.Cell(10, 25).text = "+"
    Grid1.Range(10, 17, 10, 22).Merge
    
    Grid1.Cell(11, 2).text = "9"
    Grid1.Range(11, 5, 11, 15).Merge
    Grid1.Cell(11, 5).text = " Boletas"
    Grid1.Cell(11, 16).text = "110"
    Grid1.Cell(11, 23).text = "111"
    Grid1.Cell(11, 25).text = "+"
    Grid1.Range(11, 17, 11, 22).Merge
    
    Grid1.Cell(12, 2).text = "10"
    Grid1.Range(12, 5, 12, 15).Merge
    Grid1.Cell(12, 5).text = " Notas de Débito emitidas del giro"
    Grid1.Cell(12, 16).text = "512"
    Grid1.Cell(12, 23).text = "513"
    Grid1.Cell(12, 25).text = "+"
    Grid1.Range(12, 17, 12, 22).Merge
    
    Grid1.Cell(13, 2).text = "11"
    Grid1.Range(13, 5, 13, 15).Merge
    Grid1.Cell(13, 5).text = " Notas de Crédito emitidas por Facturas asociadas al giro"
    Grid1.Cell(13, 16).text = "509"
    Grid1.Cell(13, 23).text = "510"
    Grid1.Cell(13, 25).text = "-"
    Grid1.Range(13, 17, 13, 22).Merge
    
    Grid1.Cell(14, 2).text = "12"
    Grid1.Range(14, 5, 14, 15).Merge
    Grid1.Cell(14, 5).text = " Notas de Crédito emitidas por Vales de máquinas autorizadas por el Servicio"
    Grid1.Cell(14, 16).text = "708"
    Grid1.Cell(14, 23).text = "709"
    Grid1.Cell(14, 25).text = "-"
    Grid1.Range(14, 17, 14, 22).Merge
    
    Grid1.Cell(15, 2).text = "13"
    Grid1.Range(15, 5, 15, 15).Merge
    Grid1.Cell(15, 5).text = " Notas de Crédito emitidas por ventas y servicios que no son del giro (activo fijo y otros)"
    Grid1.Cell(15, 16).text = "733"
    Grid1.Cell(15, 23).text = "734"
    Grid1.Cell(15, 25).text = "-"
    Grid1.Range(15, 17, 15, 22).Merge
    
    Grid1.Cell(16, 2).text = "14"
    Grid1.Range(16, 5, 16, 15).Merge
    Grid1.Cell(16, 5).text = " Facturas de Compra recibidas con retención parcial (contribuyentes retenidos)"
    Grid1.Cell(16, 16).text = "516"
    Grid1.Cell(16, 23).text = "517"
    Grid1.Cell(16, 25).text = "+"
    Grid1.Range(16, 17, 16, 22).Merge
    
    Grid1.Cell(17, 2).text = "15"
    Grid1.Range(17, 5, 17, 15).Merge
    Grid1.Cell(17, 5).text = " Liquidación y Liquidación Factura"
    Grid1.Cell(17, 16).text = "500"
    Grid1.Cell(17, 23).text = "501"
    Grid1.Cell(17, 25).text = "+"
    Grid1.Range(17, 17, 17, 22).Merge
    
    Grid1.Cell(18, 2).text = "16"
    Grid1.Range(18, 3, 18, 22).Merge
    Grid1.Cell(18, 3).Alignment = cellLeftGeneral
    Grid1.Cell(18, 3).text = " Adiciones al Débito Fiscal del mes, originadas en devoluciones excesivas registradas en otros períodos por Art. 27 bis"
    Grid1.Cell(18, 23).text = "154"
    Grid1.Cell(18, 25).text = "+"
    
    
    Grid1.Cell(19, 2).text = "17"
    Grid1.Range(19, 3, 19, 22).Merge
    Grid1.Cell(19, 3).Alignment = cellLeftGeneral
    Grid1.Cell(19, 3).text = " Restitución Adicional por proporción de operaciones exentas y/o no gravadas por concepto Art. 27 bis, inc. 2° (Ley Nº 19.738)"
    Grid1.Cell(19, 23).text = "518"
    Grid1.Cell(19, 25).text = "+"
    
    Grid1.Cell(20, 2).text = "18"
    Grid1.Range(20, 3, 20, 22).Merge
    Grid1.Cell(20, 3).Alignment = cellLeftGeneral
    Grid1.Cell(20, 3).text = " Reintegro del Impuesto de Timbres y Estampillas, Art. 3° Ley N° 20.259"
    Grid1.Cell(20, 23).text = "713"
    Grid1.Cell(20, 25).text = "+"
    
    Grid1.Cell(21, 2).text = "19"
    Grid1.Range(21, 3, 21, 8).Merge
    Grid1.Cell(21, 3).Alignment = cellLeftGeneral
    Grid1.Cell(21, 3).text = " Adiciones al Débito por IEPD, Ley 20.493"
    Grid1.Cell(21, 9).text = "M3"
    Grid1.Cell(21, 10).text = "738"
    Grid1.Range(21, 11, 21, 12).Merge
    Grid1.Cell(21, 13).text = "Base"
    Grid1.Cell(21, 13).Alignment = cellLeftGeneral
    Grid1.Cell(21, 14).text = "739"
    Grid1.Range(21, 15, 21, 17).Merge
    Grid1.Range(21, 18, 21, 19).Merge
    Grid1.Cell(21, 18).text = "Variable"
    Grid1.Cell(21, 18).Alignment = cellLeftGeneral
    Grid1.Cell(21, 20).text = "740"
    Grid1.Range(21, 21, 21, 22).Merge
    Grid1.Cell(21, 23).text = "741"
    Grid1.Cell(21, 25).text = "+"
    
    Grid1.Cell(22, 2).text = "20"
    Grid1.Range(22, 3, 22, 22).Merge
    Grid1.Cell(22, 3).Alignment = cellLeftGeneral
    Grid1.Cell(22, 3).Font.Bold = True
    Grid1.Cell(22, 3).text = " TOTAL DEBITOS"
    Grid1.Cell(22, 23).text = "538"
    Grid1.Cell(22, 25).text = "-"
    
    Grid1.Range(1, 1, 22, 1).Merge
    Grid1.Range(1, 1, 22, 1).Alignment = cellCenterCenter
    Grid1.Range(1, 1, 22, 1).WrapText = True
    Grid1.Cell(1, 1).text = "D E B   I  T O S    Y    V E N T A S "
    Grid1.Range(23, 3, 23, 24).Merge
    
End Sub
    Sub segundocuadro()
    
    Grid1.Range(24, 2, 24, 15).Merge
    Grid1.Range(24, 2, 24, 15).Alignment = cellCenterGeneral
    Grid1.Cell(24, 2).text = "IMPUESTO AL VALOR AGREGADO D.L. 825/74"
    Grid1.Cell(24, 2).Font.Bold = True
    Grid1.Range(24, 16, 24, 22).Merge
    Grid1.Cell(24, 16).Alignment = cellCenterGeneral
    Grid1.Cell(24, 16).text = " Con derecho a Crédito"
    Grid1.Cell(24, 16).Font.Bold = True
    Grid1.Range(24, 23, 24, 25).Merge
    Grid1.Cell(24, 23).Alignment = cellCenterGeneral
    Grid1.Cell(24, 23).text = "Sin derecho a Crédito"
    Grid1.Cell(24, 23).Font.Bold = True
    
    
    Grid1.Cell(25, 2).text = "21"
    Grid1.Range(25, 4, 25, 15).Merge
    Grid1.Cell(25, 4).text = " IVA por documentos electrónicos recibidos"
    Grid1.Cell(25, 4).Alignment = cellLeftGeneral
    Grid1.Cell(25, 16).text = "511"
    Grid1.Cell(25, 23).text = "514"
    Grid1.Range(25, 17, 25, 22).Merge
    Grid1.Range(25, 24, 25, 25).Merge
    
    Grid1.Range(26, 4, 26, 15).Merge
    Grid1.Range(26, 4, 26, 15).Alignment = cellCenterGeneral
    Grid1.Range(26, 16, 26, 22).Merge
    Grid1.Cell(26, 16).Alignment = cellCenterGeneral
    Grid1.Cell(26, 16).text = "Cantidad de documentos"
    Grid1.Cell(26, 16).Font.Bold = True
    Grid1.Range(26, 23, 26, 25).Merge
    Grid1.Cell(26, 23).Alignment = cellCenterGeneral
    Grid1.Cell(26, 23).text = "Monto Neto"
    Grid1.Cell(26, 23).Font.Bold = True
    
    
    
    
    Grid1.Cell(27, 2).text = "22"
    Grid1.Range(27, 6, 27, 15).Merge
    Grid1.Cell(27, 6).text = " Internas afectas"
    Grid1.Cell(27, 6).Alignment = cellLeftGeneral
    Grid1.Cell(27, 16).text = "564"
    Grid1.Cell(27, 23).text = "521"
    Grid1.Range(27, 17, 27, 22).Merge
    Grid1.Range(27, 24, 27, 25).Merge
    
    Grid1.Cell(28, 2).text = "23"
    Grid1.Range(28, 6, 28, 15).Merge
    Grid1.Cell(28, 6).text = " Importaciones"
    Grid1.Cell(28, 6).Alignment = cellLeftGeneral
    Grid1.Cell(28, 16).text = "566"
    Grid1.Cell(28, 23).text = "560"
    Grid1.Range(28, 17, 28, 22).Merge
    Grid1.Range(28, 24, 28, 25).Merge
    
    Grid1.Cell(29, 2).text = "24"
    Grid1.Range(29, 6, 29, 15).Merge
    Grid1.Cell(29, 6).text = " Internas exentas, o no gravadas"
    Grid1.Cell(29, 6).Alignment = cellLeftGeneral
    Grid1.Cell(29, 16).text = "584"
    Grid1.Cell(29, 23).text = "562"
    Grid1.Range(29, 17, 29, 22).Merge
    Grid1.Range(29, 24, 29, 25).Merge
    
    
    Grid1.Range(30, 3, 30, 15).Merge
    Grid1.Range(30, 3, 30, 15).Alignment = cellCenterGeneral
    Grid1.Range(30, 16, 30, 22).Merge
    Grid1.Cell(30, 16).Alignment = cellCenterGeneral
    Grid1.Cell(30, 16).text = "Cantidad de documentos"
    Grid1.Cell(30, 16).Font.Bold = True
    Grid1.Range(30, 23, 30, 25).Merge
    Grid1.Cell(30, 23).Alignment = cellCenterGeneral
    Grid1.Cell(30, 23).text = "Crédito, Recuperación y Reintegro"
    Grid1.Cell(30, 23).Font.Bold = True
    
    Grid1.Cell(31, 2).text = "25"
    Grid1.Range(31, 6, 31, 15).Merge
    Grid1.Cell(31, 6).text = " Facturas recibidas del giro y Facturas de compra emitidas"
    Grid1.Cell(31, 6).Alignment = cellLeftGeneral
    Grid1.Cell(31, 16).text = "519"
    Grid1.Cell(31, 23).text = "520"
    Grid1.Range(31, 17, 31, 22).Merge
    Grid1.Cell(31, 25).text = "+"
    
    Grid1.Cell(32, 2).text = "26"
    Grid1.Range(32, 6, 32, 15).Merge
    Grid1.Cell(32, 6).text = " Facturas activo fijo"
    Grid1.Cell(32, 6).Alignment = cellLeftGeneral
    Grid1.Cell(32, 16).text = "524"
    Grid1.Cell(32, 23).text = "525"
    Grid1.Range(32, 17, 32, 22).Merge
    Grid1.Cell(32, 25).text = "+"
    
    Grid1.Cell(33, 2).text = "27"
    Grid1.Range(33, 6, 33, 15).Merge
    Grid1.Cell(33, 6).text = " Notas de Crédito recibidas"
    Grid1.Cell(33, 6).Alignment = cellLeftGeneral
    Grid1.Cell(33, 16).text = "527"
    Grid1.Cell(33, 23).text = "528"
    Grid1.Range(33, 17, 33, 22).Merge
    Grid1.Cell(33, 25).text = "-"
    
    Grid1.Cell(34, 2).text = "28"
    Grid1.Range(34, 6, 34, 15).Merge
    Grid1.Cell(34, 6).text = " Notas de Débito recibidas"
    Grid1.Cell(34, 6).Alignment = cellLeftGeneral
    Grid1.Cell(34, 16).text = "531"
    Grid1.Cell(34, 23).text = "532"
    Grid1.Range(34, 17, 34, 22).Merge
    Grid1.Cell(34, 25).text = "+"
    
    Grid1.Cell(35, 2).text = "29"
    Grid1.Range(35, 6, 35, 15).Merge
    Grid1.Cell(35, 6).text = " Declaraciones de Ingreso (DIN) importaciones del giro"
    Grid1.Cell(35, 6).Alignment = cellLeftGeneral
    Grid1.Cell(35, 16).text = "534"
    Grid1.Cell(35, 23).text = "535"
    Grid1.Range(35, 17, 35, 22).Merge
    Grid1.Cell(35, 25).text = "+"
    
    Grid1.Cell(36, 2).text = "30"
    Grid1.Range(36, 6, 36, 15).Merge
    Grid1.Cell(36, 6).text = " Declaraciones de Ingreso (DIN) importaciones activo fijo"
    Grid1.Cell(36, 6).Alignment = cellLeftGeneral
    Grid1.Cell(36, 16).text = "536"
    Grid1.Cell(36, 23).text = "553"
    Grid1.Range(36, 17, 36, 22).Merge
    Grid1.Cell(36, 25).text = "+"
    
    Grid1.Cell(37, 2).text = "31"
    Grid1.Range(37, 3, 37, 22).Merge
    Grid1.Cell(37, 3).Alignment = cellLeftGeneral
    Grid1.Cell(37, 3).text = " Remanente Crédito Fiscal mes anterior"
    Grid1.Cell(37, 23).text = "504"
    Grid1.Cell(37, 25).text = "+"
    
    Grid1.Cell(38, 2).text = "32"
    Grid1.Range(38, 3, 38, 22).Merge
    Grid1.Cell(38, 3).Alignment = cellLeftGeneral
    Grid1.Cell(38, 3).text = " Devolución Solicitud Art. 36 (Exportadores)"
    Grid1.Cell(38, 23).text = "593"
    Grid1.Cell(38, 25).text = "-"
    
    
    Grid1.Cell(39, 2).text = "33"
    Grid1.Range(39, 3, 39, 22).Merge
    Grid1.Cell(39, 3).Alignment = cellLeftGeneral
    Grid1.Cell(39, 3).text = " Devolución Solicitud Art. 27 bis (Activo fijo)"
    Grid1.Cell(39, 23).text = "594"
    Grid1.Cell(39, 25).text = "-"
    
    Grid1.Cell(40, 2).text = "34"
    Grid1.Range(40, 3, 40, 22).Merge
    Grid1.Cell(40, 3).Alignment = cellLeftGeneral
    Grid1.Cell(40, 3).text = " Certificado Imputación Art. 27 bis (Activo fijo)"
    Grid1.Cell(40, 23).text = "592"
    Grid1.Cell(40, 25).text = "-"
    
    Grid1.Cell(41, 2).text = "35"
    Grid1.Range(41, 3, 41, 22).Merge
    Grid1.Cell(41, 3).Alignment = cellLeftGeneral
    Grid1.Cell(41, 3).text = " Devolución Solicitud Art. 3° (Cambio de Sujeto)"
    Grid1.Cell(41, 23).text = "539"
    Grid1.Cell(41, 25).text = "-"
    
    Grid1.Cell(42, 2).text = "36"
    Grid1.Range(42, 3, 42, 22).Merge
    Grid1.Cell(42, 3).Alignment = cellLeftGeneral
    Grid1.Cell(42, 3).text = " Devolución Solicitud Ley N° 20.258 por remanente CF IVA originado en Impuesto específico Petróleo Diesel (Generadoras Eléctricas)"
    Grid1.Cell(42, 23).text = "718"
    Grid1.Cell(42, 25).text = "-"
    
    Grid1.Cell(43, 2).text = "37"
    Grid1.Range(43, 3, 43, 22).Merge
    Grid1.Cell(43, 3).Alignment = cellLeftGeneral
    Grid1.Cell(43, 3).text = " Monto Reintegrado por Devolución Indebida de Crédito Fiscal D.S. 348 (Exportadores)"
    Grid1.Cell(43, 23).text = "164"
    Grid1.Cell(43, 25).text = "+"
    
   
    Grid1.Range(44, 4, 44, 11).Merge
    Grid1.Range(44, 12, 44, 18).Merge
    Grid1.Cell(44, 12).Alignment = cellCenterGeneral
    Grid1.Cell(44, 12).text = " M3 Comprados con derecho a crédito"
    Grid1.Range(44, 19, 44, 22).Merge
    Grid1.Cell(44, 19).Alignment = cellCenterGeneral
    Grid1.Cell(44, 19).text = "Componentes del Impuesto"
    
    Grid1.Range(45, 2, 46, 2).Merge
    Grid1.Cell(45, 2).text = "38"
    Grid1.Range(45, 4, 46, 11).Merge
    Grid1.Range(45, 4, 46, 11).WrapText = True
    Grid1.Cell(45, 4).Alignment = cellLeftGeneral
    Grid1.Cell(45, 4).text = " Recuperación del Impuesto Específico al Petróleo Diesel                              (Art. 7º Ley 18.502, Arts. 1º y 3º D.S. Nº311/86)"
    Grid1.Range(45, 12, 46, 12).Merge
    Grid1.Cell(45, 12).text = "730"
    Grid1.Range(45, 13, 46, 16).Merge
    Grid1.Range(45, 17, 45, 18).Merge
    Grid1.Range(46, 17, 46, 18).Merge
    Grid1.Cell(45, 17).text = " Base"
    Grid1.Cell(45, 17).Alignment = cellLeftGeneral
    Grid1.Cell(46, 17).text = " Variable"
    Grid1.Cell(46, 17).Alignment = cellLeftGeneral
    Grid1.Cell(45, 19).text = " 742"
    Grid1.Cell(46, 19).text = " 743"
    Grid1.Range(45, 20, 45, 22).Merge
    Grid1.Range(46, 20, 46, 22).Merge
    Grid1.Range(45, 23, 46, 23).Merge
    Grid1.Cell(45, 23).text = "127"
    Grid1.Range(45, 24, 46, 24).Merge
    Grid1.Range(45, 25, 46, 25).Merge
    Grid1.Cell(45, 25).text = "+"
    
    Grid1.Range(47, 2, 48, 2).Merge
    Grid1.Cell(47, 2).text = "39"
    Grid1.Range(47, 4, 48, 11).Merge
    Grid1.Range(47, 4, 48, 11).WrapText = True
    Grid1.Cell(47, 4).Alignment = cellLeftGeneral
    Grid1.Cell(47, 4).text = " Recuperación del Impuesto Específico al Petróleo Diesel                              soportado por Transportistas de Carga (Art. 2º Ley 19.764)"
    Grid1.Range(47, 12, 48, 12).Merge
    Grid1.Cell(47, 12).text = "729"
    Grid1.Range(47, 13, 48, 16).Merge
    Grid1.Range(47, 17, 47, 18).Merge
    Grid1.Range(48, 17, 48, 18).Merge
    Grid1.Cell(47, 17).text = " Base"
    Grid1.Cell(47, 17).Alignment = cellLeftGeneral
    Grid1.Cell(48, 17).text = " Variable"
    Grid1.Cell(48, 17).Alignment = cellLeftGeneral
    Grid1.Cell(47, 19).text = " 744"
    Grid1.Cell(48, 19).text = " 745"
    Grid1.Range(47, 20, 47, 22).Merge
    Grid1.Range(48, 20, 48, 22).Merge
    Grid1.Range(47, 23, 48, 23).Merge
    Grid1.Cell(47, 23).text = "544"
    Grid1.Range(47, 24, 48, 24).Merge
    Grid1.Range(47, 25, 48, 25).Merge
    Grid1.Cell(47, 25).text = "+"
    
    
    Grid1.Cell(49, 2).text = "40"
    Grid1.Range(49, 3, 49, 22).Merge
    Grid1.Cell(49, 3).Alignment = cellLeftGeneral
    Grid1.Cell(49, 3).text = " Crédito del Art. 11º Ley 18.211 (correspondiente a Zona Franca de extensión)"
    Grid1.Cell(49, 23).text = "523"
    Grid1.Cell(49, 25).text = "+"
    
    Grid1.Cell(50, 2).text = "41"
    Grid1.Range(50, 3, 50, 22).Merge
    Grid1.Cell(50, 3).Alignment = cellLeftGeneral
    Grid1.Cell(50, 3).text = " Crédito por Impuesto de Timbres y Estampillas, Art. 3º Ley 20.259"
    Grid1.Cell(50, 23).text = "712"
    Grid1.Cell(50, 25).text = "+"
    
    Grid1.Cell(51, 2).text = "42"
    Grid1.Range(51, 3, 51, 22).Merge
    Grid1.Cell(51, 3).Alignment = cellLeftGeneral
    Grid1.Cell(51, 3).text = " TOTAL CREDITOS"
    Grid1.Cell(51, 3).Font.Bold = True
    Grid1.Cell(51, 23).text = "537"
    Grid1.Cell(51, 25).text = "="
    
    
    Grid1.Range(24, 1, 51, 1).Merge
    Grid1.Range(24, 1, 22, 1).Alignment = cellCenterCenter
    Grid1.Range(24, 1, 22, 1).WrapText = True
    Grid1.Cell(24, 1).text = "C  R  É  D  I  T  O  S    y    C  O  M  P  R  A  S "
    
    
    
    End Sub
    Sub tercercuadro()
    
    Grid1.Range(52, 1, 52, 25).Merge
    Grid1.RowHeight(52) = 10
    Grid1.Range(53, 1, 53, 15).Merge
    Grid1.Range(53, 1, 53, 15).Alignment = cellCenterGeneral
    Grid1.Cell(53, 1).text = "   Diferencia Total Débitos (línea 20, código 538) menos Total Créditos (línea 42, código 537); trasládelo a la línea 43.          Si el resultado es positivo al código 89, si es negativo al código 77 sin signo."
    Grid1.Cell(53, 1).WrapText = True
    Grid1.RowHeight(53) = 40
      
    Grid1.Range(54, 1, 54, 22).Merge
    Grid1.Range(54, 23, 54, 25).Merge
    Grid1.Cell(54, 23).Alignment = cellCenterGeneral
    Grid1.Cell(54, 23).text = "IMPUESTO DETERMINADO"
    Grid1.Cell(54, 23).Font.Bold = True
    
    Grid1.Cell(55, 2).text = "43"
    Grid1.Range(55, 3, 55, 11).Merge
    Grid1.Cell(55, 3).Alignment = cellLeftGeneral
    Grid1.Cell(55, 3).text = " Remanente de crédito fiscal para el período siguiente"
    Grid1.Cell(55, 12).text = "77"
    Grid1.Range(55, 13, 55, 19).Merge
    Grid1.Range(55, 20, 55, 22).Merge
    Grid1.Cell(55, 20).Alignment = cellLeftGeneral
    Grid1.Cell(55, 20).text = "IVA determinado"
    Grid1.Cell(55, 23).text = "89"
    Grid1.Cell(55, 25).text = "+"
    Grid1.Range(56, 1, 56, 25).Merge
    
    Grid1.Cell(57, 2).text = "44"
    Grid1.Range(57, 3, 57, 22).Merge
    Grid1.Cell(57, 3).Alignment = cellLeftGeneral
    Grid1.Cell(57, 3).text = " Retención Impuesto Primera Categoría por rentas de capitales mobiliarios del Art. 20 N° 2, según Art. 73 LIR"
    Grid1.Cell(57, 23).text = "50"
    Grid1.Cell(57, 25).text = "+"
    
    Grid1.Cell(58, 2).text = "45"
    Grid1.Range(58, 3, 58, 10).Merge
    Grid1.Cell(58, 3).Alignment = cellLeftGeneral
    Grid1.Cell(58, 3).text = " Retención Impuesto Único a los Trabajadores, según Art. 74 N° 1 LIR"
    Grid1.Range(58, 11, 58, 14).Merge
    Grid1.Cell(58, 11).Alignment = cellLeftGeneral
    Grid1.Cell(58, 11).text = "Crédito Donación Ley 20.444/2010"
    Grid1.Cell(58, 15).text = "735"
    Grid1.Range(58, 16, 58, 17).Merge
    Grid1.Range(58, 18, 58, 22).Merge
    Grid1.Cell(58, 18).text = "Impuesto Único 2da. Categoría a Pagar"
    Grid1.Cell(58, 18).Alignment = cellLeftGeneral
    Grid1.Cell(58, 23).text = "48"
    Grid1.Cell(58, 25).text = "+"
    
    
    Grid1.Cell(59, 2).text = "46"
    Grid1.Range(59, 3, 59, 22).Merge
    Grid1.Cell(59, 3).Alignment = cellLeftGeneral
    Grid1.Cell(59, 3).text = " Retención de Impuesto con tasa del 10% sobre las rentas del Art. 42 N°2, según Art. 74 N°2 LIR"
    Grid1.Cell(59, 23).text = "151"
    Grid1.Cell(59, 25).text = "+"
    
    Grid1.Cell(60, 2).text = "47"
    Grid1.Range(60, 3, 60, 22).Merge
    Grid1.Cell(60, 3).Alignment = cellLeftGeneral
    Grid1.Cell(60, 3).text = " Retención de Impuesto con tasa del 10% sobre las rentas del Art. 48, según Art. 74 N°3 LIR"
    Grid1.Cell(60, 23).text = "153"
    Grid1.Cell(60, 25).text = "+"
    
    Grid1.Cell(61, 2).text = "48"
    Grid1.Range(61, 3, 61, 22).Merge
    Grid1.Cell(61, 3).Alignment = cellLeftGeneral
    Grid1.Cell(61, 3).text = " Retención a Suplementeros, según Art. 74 N° 5 (tasa 0,5%) LIR"
    Grid1.Cell(61, 23).text = "54"
    Grid1.Cell(61, 25).text = "+"
    
    Grid1.Cell(62, 2).text = "49"
    Grid1.Range(62, 3, 62, 22).Merge
    Grid1.Cell(62, 3).Alignment = cellLeftGeneral
    Grid1.Cell(62, 3).text = " Retención por compra de productos mineros, según Art. 74 N° 6 LIR"
    Grid1.Cell(62, 23).text = "56"
    Grid1.Cell(62, 25).text = "+"
    
    Grid1.Cell(63, 2).text = "50"
    Grid1.Range(63, 3, 63, 22).Merge
    Grid1.Cell(63, 3).Alignment = cellLeftGeneral
    Grid1.Cell(63, 3).text = " Retención sobre cantidades pagadas en cumplimiento de Seguros Dotales del Art.17, N°3 (tasa 15%)"
    Grid1.Cell(63, 23).text = "588"
    Grid1.Cell(63, 25).text = "+"
    
    Grid1.Cell(64, 2).text = "51"
    Grid1.Range(64, 3, 64, 22).Merge
    Grid1.Cell(64, 3).Alignment = cellLeftGeneral
    Grid1.Cell(64, 3).text = " Retención sobre retiros de Ahorro Previsional Voluntario del Art. 42 bis LIR (tasa 15%)"
    Grid1.Cell(64, 23).text = "589"
    Grid1.Cell(64, 25).text = "+"
 
    Grid1.Range(65, 2, 65, 6).Merge
    Grid1.Range(65, 7, 65, 10).Merge
    Grid1.Cell(65, 7).Alignment = cellLeftGeneral
    Grid1.Cell(65, 7).text = "          Acogido a suspensión PPM         (Art 1º bis Ley 19.420 y 1º bis Ley 19.606)"
    Grid1.Cell(65, 7).Font.Size = 7
    Grid1.Cell(65, 7).WrapText = True
    Grid1.Range(65, 11, 65, 13).Merge
    Grid1.Cell(65, 11).Alignment = cellCenterGeneral
    Grid1.Cell(65, 11).text = "Monto Pérdida Art. 90"
    Grid1.Cell(65, 11).Font.Bold = True
    Grid1.Range(65, 14, 65, 17).Merge
    Grid1.Cell(65, 14).Alignment = cellCenterGeneral
    Grid1.Cell(65, 14).text = "Base Imponible"
    Grid1.Cell(65, 14).Font.Bold = True
    Grid1.Range(65, 18, 65, 20).Merge
    Grid1.Cell(65, 18).Alignment = cellCenterGeneral
    Grid1.Cell(65, 18).text = "Tasa"
    Grid1.Cell(65, 18).Font.Bold = True
    Grid1.Range(65, 21, 65, 22).Merge
    Grid1.Cell(65, 21).Alignment = cellLeftGeneral
    Grid1.Cell(65, 21).text = "Crédito/Tope Suspensión PPM (Arts. 1ºbis Leyes 19.420 y 19.606)"
    Grid1.Cell(65, 21).Font.Size = 6
    Grid1.Cell(65, 21).WrapText = True
    Grid1.Range(65, 23, 65, 25).Merge
    Grid1.Cell(65, 23).Alignment = cellCenterGeneral
    Grid1.Cell(65, 23).text = "PPM Neto Determinado"
    Grid1.Cell(65, 23).Font.Bold = True
    Grid1.RowHeight(65) = 35
    
    Grid1.Cell(66, 2).text = "52"
    Grid1.Range(66, 4, 66, 6).Merge
    Grid1.Cell(66, 4).Alignment = cellLeftGeneral
    Grid1.Cell(66, 4).text = " 1ra. Categoría Art. 84 a)"
    Grid1.Cell(66, 7).text = "750"
    Grid1.Range(66, 8, 66, 10).Merge
    Grid1.Cell(66, 8).CellType = cellCheckBox
    Grid1.Cell(66, 11).text = "30"
    Grid1.Range(66, 12, 66, 13).Merge
    Grid1.Cell(66, 14).text = "563"
    Grid1.Range(66, 15, 66, 17).Merge
    Grid1.Cell(66, 18).text = "115"
    Grid1.Range(66, 19, 66, 20).Merge
    Grid1.Cell(66, 21).text = "68"
    
    Grid1.Cell(66, 23).text = "62"
    Grid1.Cell(66, 25).text = "+"
    
    Grid1.Cell(67, 2).text = "53"
    Grid1.Range(67, 4, 67, 10).Merge
    Grid1.Cell(67, 4).Alignment = cellLeftGeneral
    Grid1.Cell(67, 4).text = " Mineros Art. 84 a)"
    Grid1.Cell(67, 11).text = "565"
    Grid1.Range(67, 12, 67, 13).Merge
    Grid1.Cell(67, 14).text = "120"
    Grid1.Range(67, 15, 67, 17).Merge
    Grid1.Cell(67, 18).text = "542"
    Grid1.Range(67, 19, 67, 20).Merge
    Grid1.Cell(67, 21).text = "122"
    Grid1.Cell(67, 23).text = "123"
    Grid1.Cell(67, 25).text = "+"
    
    Grid1.Cell(68, 2).text = "54"
    Grid1.Range(68, 4, 68, 10).Merge
    Grid1.Cell(68, 4).Alignment = cellLeftGeneral
    Grid1.Cell(68, 4).text = " Explotador Minero Art. 84 h)"
    Grid1.Cell(68, 11).text = "700"
    Grid1.Range(68, 12, 68, 13).Merge
    Grid1.Cell(68, 14).text = "701"
    Grid1.Range(68, 15, 68, 17).Merge
    Grid1.Cell(68, 18).text = "702"
    Grid1.Range(68, 19, 67, 20).Merge
    Grid1.Cell(68, 21).text = "711"
    Grid1.Cell(68, 23).text = "703"
    Grid1.Cell(68, 25).text = "+"
    
    Grid1.Cell(69, 2).text = "55"
    Grid1.Range(69, 4, 69, 22).Merge
    Grid1.Cell(69, 4).Alignment = cellLeftGeneral
    Grid1.Cell(69, 4).text = " Transportistas acogidos a Renta Presunta, Art. 84, e) y f) (tasa de 0,3%)"
    Grid1.Cell(69, 23).text = "66"
    Grid1.Cell(69, 25).text = "+"
    
    
    Grid1.Range(70, 2, 71, 2).Merge
    Grid1.Cell(70, 2).text = "56"
    Grid1.Range(70, 4, 71, 10).Merge
    Grid1.Range(70, 4, 71, 10).WrapText = True
    Grid1.Cell(70, 4).Alignment = cellLeftGeneral
    Grid1.Cell(70, 4).text = " Crédito Capacitación, Ley 19.518/97"
    Grid1.Range(70, 11, 70, 13).Merge
    Grid1.Cell(70, 11).text = "Crédito del Mes"
    Grid1.Cell(71, 11).text = "721"
    Grid1.Range(71, 12, 71, 13).Merge
    Grid1.Range(70, 14, 70, 17).Merge
    Grid1.Cell(70, 14).text = "Remanente Mes Anterior"
    Grid1.Cell(71, 14).text = "722"
    Grid1.Range(71, 15, 71, 17).Merge
    Grid1.Range(70, 18, 70, 21).Merge
    Grid1.Cell(70, 18).text = "Remanente Período Siguiente"
    Grid1.Cell(70, 18).Font.Size = 7
    Grid1.Cell(71, 18).text = "724"
    Grid1.Range(71, 19, 71, 20).Merge
    Grid1.Range(70, 22, 70, 25).Merge
    Grid1.Range(71, 21, 71, 22).Merge
    Grid1.Cell(71, 21).Alignment = cellCenterGeneral
    Grid1.Cell(71, 21).text = "Crédito a Imputar"
    Grid1.Cell(71, 23).text = "723"
    Grid1.Cell(71, 25).text = "-"
    
    Grid1.Cell(72, 2).text = "57"
    Grid1.Range(72, 4, 72, 22).Merge
    Grid1.Cell(72, 4).Alignment = cellLeftGeneral
    Grid1.Cell(72, 4).text = " 2da. Categoría Art. 84, b) (tasa 10%)"
    Grid1.Cell(72, 23).text = "152"
    Grid1.Cell(72, 25).text = "+"
    
    Grid1.Cell(73, 2).text = "58"
    Grid1.Range(73, 4, 73, 22).Merge
    Grid1.Cell(73, 4).Alignment = cellLeftGeneral
    Grid1.Cell(73, 4).text = " Taller artesanal Art. 84, c) (tasa de 1,5% o 3%)"
    Grid1.Cell(73, 23).text = "70"
    Grid1.Cell(73, 25).text = "+"
    
    Grid1.Range(74, 1, 74, 25).Merge
    
    Grid1.Range(57, 1, 73, 1).Merge
    Grid1.Range(57, 1, 73, 1).Alignment = cellCenterCenter
    Grid1.Range(57, 1, 73, 1).WrapText = True
    Grid1.Cell(57, 1).text = " I MPUESTO A LA RENTA D.L. 824/74"
    
    
    
    Grid1.Cell(75, 2).text = "59"
    Grid1.Range(75, 3, 75, 22).Merge
    Grid1.Cell(75, 3).Alignment = cellLeftGeneral
    Grid1.Cell(75, 3).text = " SUB TOTAL IMPUESTO DETERMINADO ANVERSO. (Suma de las líneas 43 a 58, columna Impuesto y/o PPM determinado)"
    Grid1.Cell(75, 3).Font.Bold = True
    Grid1.Cell(75, 23).text = "595"
    Grid1.Cell(75, 25).text = "="
    
    Grid1.Range(76, 1, 76, 25).Merge
    Grid1.RowHeight(76) = 10
    Grid1.Range(77, 1, 77, 25).Merge
    Grid1.Range(77, 1, 77, 25).Alignment = cellCenterGeneral
    Grid1.Cell(77, 1).text = "Si no declara tributación simplificada, Impuesto Adicional (Art. 37º o Art. 42º, DL Nº 825), cambio de sujeto y créditos especiales por concepto de Sistemas Solares Térmicos; Patentes por Derechos de Agua; Cotización Adicional; Empresas Constructoras y Peajes Empresas de Transporte de Pasajeros, traslade el valor de línea 59 (código 595) a línea 111 (código 91). En caso contrario, continúe al reverso."
    Grid1.Cell(77, 1).WrapText = True
    Grid1.RowHeight(77) = 40
    Grid1.Range(78, 1, 78, 25).Merge
    Grid1.RowHeight(78) = 10
    
    
    Grid1.Cell(79, 1).text = "01"
    Grid1.Range(79, 2, 79, 15).Merge
    Grid1.Cell(79, 2).Alignment = cellLeftGeneral
    Grid1.Cell(79, 2).text = " Apellido Paterno o Razón Social"
    Grid1.Cell(79, 2).Font.Bold = True
    Grid1.Cell(79, 16).text = "02"
    Grid1.Range(79, 17, 79, 22).Merge
    Grid1.Cell(79, 17).Alignment = cellLeftGeneral
    Grid1.Cell(79, 17).text = " Apellido Paterno o Razón Social"
    Grid1.Cell(79, 17).Font.Bold = True
    Grid1.Cell(79, 23).text = "05"
    Grid1.Range(79, 24, 79, 25).Merge
    Grid1.Cell(79, 24).Alignment = cellLeftGeneral
    Grid1.Cell(79, 24).text = " Nombres"
    Grid1.Cell(79, 24).Font.Bold = True
    Grid1.Range(80, 1, 80, 15).Merge
    Grid1.Cell(80, 1).Alignment = cellLeftGeneral
    Grid1.Range(80, 16, 80, 22).Merge
    Grid1.Cell(80, 16).Alignment = cellLeftGeneral
    Grid1.Range(80, 23, 80, 25).Merge
    Grid1.Cell(80, 24).Alignment = cellLeftGeneral
    
    Grid1.Range(81, 1, 81, 6).Merge
    Grid1.Cell(81, 1).Alignment = cellCenterGeneral
    Grid1.Cell(81, 1).text = "Cambia datos de Domicilio"
    Grid1.Cell(81, 1).Font.Bold = True
    Grid1.Cell(81, 7).text = "583"
    Grid1.Range(81, 9, 81, 20).Merge
    Grid1.Cell(81, 9).Alignment = cellLeftGeneral
    Grid1.Cell(81, 9).text = "(Si marca con X el casillero, registre los cambios al reverso)"
    Grid1.Range(81, 21, 81, 25).Merge
    Grid1.Cell(81, 21).Alignment = cellCenterGeneral
    Grid1.Cell(81, 21).text = "Viene de línea 59 código 595, ó línea 105 código 547"
    Grid1.Range(82, 1, 82, 25).Merge
    Grid1.RowHeight(82) = 10
    
    Grid1.Range(83, 1, 83, 13).Merge
    Grid1.Range(83, 14, 83, 25).Merge
    Grid1.Cell(83, 1).WrapText = True
    Grid1.Cell(83, 1).Alignment = cellCenterGeneral
    Grid1.Cell(83, 1).text = "Declaro bajo juramento que los datos contenidos en esta declaración son la expresión fiel de la verdad, por lo que asumo la responsabilidad correspondiente."
    Grid1.RowHeight(83) = 33
    
    Grid1.Range(84, 1, 84, 10).Merge
    Grid1.Cell(84, 11).text = "111"
    Grid1.Range(84, 1, 84, 10).Merge
    Grid1.Range(84, 12, 84, 22).Merge
    Grid1.Cell(84, 12).Font.Bold = True
    Grid1.Cell(84, 12).Alignment = cellLeftGeneral
    Grid1.Cell(84, 12).text = "TOTAL A PAGAR EN PLAZO LEGAL"
    Grid1.Cell(84, 23).text = "91"
    Grid1.Cell(84, 25).text = "="
    
    Grid1.Range(85, 1, 85, 10).Merge
    Grid1.Cell(85, 11).text = "112"
    Grid1.Range(85, 1, 85, 10).Merge
    Grid1.Range(85, 12, 85, 22).Merge
    Grid1.Cell(85, 12).Alignment = cellLeftGeneral
    Grid1.Cell(85, 12).text = "Más IPC"
    Grid1.Cell(85, 23).text = "92"
    Grid1.Cell(85, 25).text = "+"
    
    Grid1.Range(86, 1, 86, 10).Merge
    Grid1.Cell(86, 11).text = "113"
    Grid1.Range(86, 1, 86, 10).Merge
    Grid1.Range(86, 12, 86, 22).Merge
    Grid1.Cell(86, 12).Alignment = cellLeftGeneral
    Grid1.Cell(86, 12).text = "Más Intereses y multas"
    Grid1.Cell(86, 23).text = "93"
    Grid1.Cell(86, 25).text = "+"
    
    
    Grid1.Range(87, 1, 87, 10).Merge
    Grid1.Cell(87, 11).text = "114"
    Grid1.Range(87, 1, 87, 10).Merge
    Grid1.Range(87, 12, 87, 22).Merge
    Grid1.Cell(87, 12).Alignment = cellLeftGeneral
    Grid1.Cell(87, 12).text = "TOTAL A PAGAR CON RECARGO"
    Grid1.Cell(87, 23).text = "94"
    Grid1.Cell(87, 25).text = "="
    
 
    Grid1.Range(88, 12, 88, 22).Merge
    Grid1.Cell(88, 12).Alignment = cellLeftGeneral
    Grid1.Cell(88, 12).text = "FORM. N° 29 - 05/2013 - AMF - A. MOLINA FLORES S.A."
  
    
    
    
    
    
    
        
    End Sub

Private Sub monto_Click()
End Sub

Private Sub leer()
    Call leerposiciones("503")
     If hoja = "1" Then
        Grid1.Cell(posx, posy).text = Format(grillainformes.Grid1.Cell(1, 1).text, "###,###,##0")
        Grid1.Cell(posx, posy).Font.Bold = True
        Grid1.Cell(posx, posy).BackColor = vbYellow
     End If
     If hoja = "2" Then
        Grid3.Cell(posx, posy).text = Format(grillainformes.Grid1.Cell(1, 1).text, "###,###,##0")
        Grid3.Cell(posx, posy).Font.Bold = True
        Grid3.Cell(posx, posy).BackColor = vbYellow
     End If
            
    
    Call leerdatoscodigo("39", empresaactiva, MES, año)
    Call leerdatoscodigo("556", empresaactiva, MES, año)
      
End Sub
Function leer_556(empresa, MES, año) As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT ifnull(SUM(IF(fd.tipo='3' or fd.tipo='6',fd.monto*-1,monto)),0) FROM facturasdecompras_detalle AS fd INNER JOIN facturasdecompras AS fc "
    csql.sql = csql.sql + "oN fc.tipo = fd.tipo And fc.rut = fd.rut And fc.numero = fd.numero "
    csql.sql = csql.sql + "WHERE fc.mescontable ='" + MES + "' AND añocontable='" + año + "' AND (cuentadelmayor='11400005' or cuentadelmayor='11400012' or cuentadelmayor='11400009')  "

    
'    csql.sql = "SELECT SUM(IF(tipo='3' or tipo='6',monto*-1,monto)) FROM facturasdecompras_detalle WHERE fechacreacion LIKE '2014-05%'AND (cuentadelmayor='11400005' or cuentadelmayor='11400012' or cuentadelmayor='11400009')  GROUP BY cuentadelmayor "
    csql.Execute
    leer_556 = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
        leer_556 = resultados(0)
    End If
    
    csql.Close
    Set csql = Nothing
    
End Function

Sub leerdatoscodigo(codigo, empresa, MES, año)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Call leerposiciones(codigo)
    
    If codigo = "39" Then
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT ifnull(SUM(IF(fd.tipo='3'or fd.tipo='6',fd.monto*-1,fd.monto)),0) FROM facturasdeventas_detalle as fd INNER JOIN facturasdeventas AS fc "
        Rem csql.sql = "SELECT SUM(IF(fd.tipo='3',fd.monto*-1,monto)) FROM facturasdecompras_detalle AS fd INNER JOIN facturasdecompras AS fc "
        csql.sql = csql.sql + "oN fc.tipo = fd.tipo And fc.rut = fd.rut And fc.numero = fd.numero "
        
        csql.sql = csql.sql + "WHERE fc.fecha LIKE '" + año + "-" + MES + "%'AND (cuentadelmayor='23200005' or cuentadelmayor='23200009') AND (fc.tipo='1' OR fc.tipo='2' OR fc.tipo='3' ) "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            If hoja = "1" Then
             Grid1.Cell(posx, posy).text = Format(resultados(0), "###,###,##0")
              Grid1.Cell(posx, posy).Font.Bold = True
              Grid1.Cell(posx, posy).BackColor = vbYellow
            End If
            If hoja = "2" Then
              Grid3.Cell(posx, posy).text = Format(resultados(0), "###,###,##0")
              Grid3.Cell(posx, posy).Font.Bold = True
              Grid3.Cell(posx, posy).BackColor = vbYellow
            End If
            
        End If
        
        csql.Close
        Set csql = Nothing
    End If
    
    If codigo = "556" Then
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT ifnull(SUM(IF(fd.tipo='3' or fd.tipo='6',fd.monto*-1,monto)),0) FROM facturasdecompras_detalle AS fd INNER JOIN facturasdecompras AS fc "
        csql.sql = csql.sql + "oN fc.tipo = fd.tipo And fc.rut = fd.rut And fc.numero = fd.numero "
        csql.sql = csql.sql + "WHERE fc.mescontable ='" + MES + "' AND añocontable='" + año + "' AND (cuentadelmayor='11400005' or cuentadelmayor='11400012' or cuentadelmayor='11400009')  "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            If hoja = "1" Then
             Grid1.Cell(posx, posy).text = Format(resultados(0), "###,###,##0")
              Grid1.Cell(posx, posy).Font.Bold = True
              Grid1.Cell(posx, posy).BackColor = vbYellow
            End If
            If hoja = "2" Then
              Grid3.Cell(posx, posy).text = Format(resultados(0), "###,###,##0")
              Grid3.Cell(posx, posy).Font.Bold = True
              Grid3.Cell(posx, posy).BackColor = vbYellow
            End If
        End If
        
        csql.Close
        Set csql = Nothing
    End If
End Sub
Sub leerposiciones(codigo)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = conta
        csql.sql = "select posicionx,posiciony,hoja from "
        csql.sql = csql.sql & " maestro_codigof29 "
        csql.sql = csql.sql & " where codigosii='" & codigo & "' "
        csql.Execute
        posx = 0
        posy = 0
        hoja = 0
        
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        posx = resultados(0)
        posy = resultados(1)
        hoja = resultados(2)
    End If
    csql.Close
    Set csql = Nothing
    
    
End Sub
Sub limpiar()
End Sub

Sub CABEZAS2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle

    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
    End If
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub
Sub cabezas4(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle

    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
    End If
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 4
        .TopMargin = 2
        .BottomMargin = 1
        .LeftMargin = 1
        .RightMargin = 1
        .Orientation = cellLandscape
        .PrintFixedRow = True
        
        
        
        
End With

End Sub


Sub CARGAGRILLA2()
     Dim FORMATOGRILLA(10, 30)
    Grid3.DefaultFont.Size = 8
     Grid3.Rows = 79
     Grid3.Cols = 27
       
    FORMATOGRILLA(1, 1) = ""
    FORMATOGRILLA(1, 2) = ""
    FORMATOGRILLA(1, 3) = ""
    FORMATOGRILLA(1, 4) = ""
    FORMATOGRILLA(1, 5) = ""
    FORMATOGRILLA(1, 6) = ""
    FORMATOGRILLA(1, 7) = ""
    FORMATOGRILLA(1, 8) = ""
    FORMATOGRILLA(1, 9) = ""
    FORMATOGRILLA(1, 10) = ""
    FORMATOGRILLA(1, 11) = ""
    FORMATOGRILLA(1, 12) = ""
    FORMATOGRILLA(1, 13) = ""
    FORMATOGRILLA(1, 14) = ""
    FORMATOGRILLA(1, 15) = ""
    FORMATOGRILLA(1, 16) = ""
    FORMATOGRILLA(1, 17) = ""
    FORMATOGRILLA(1, 18) = ""
    FORMATOGRILLA(1, 20) = ""
    FORMATOGRILLA(1, 21) = ""
    FORMATOGRILLA(1, 22) = ""
    FORMATOGRILLA(1, 23) = ""
    FORMATOGRILLA(1, 24) = ""
    FORMATOGRILLA(1, 25) = ""
    FORMATOGRILLA(1, 26) = ""
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "2"
    FORMATOGRILLA(2, 2) = "3"
    FORMATOGRILLA(2, 3) = "4"
    FORMATOGRILLA(2, 4) = "4"
    FORMATOGRILLA(2, 5) = "18"
    FORMATOGRILLA(2, 6) = "4"
    FORMATOGRILLA(2, 7) = "4"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "3"
    FORMATOGRILLA(2, 10) = "5"
    FORMATOGRILLA(2, 11) = "6"
    FORMATOGRILLA(2, 12) = "6"
    FORMATOGRILLA(2, 13) = "4"
    FORMATOGRILLA(2, 14) = "4"
    FORMATOGRILLA(2, 15) = "4"
    FORMATOGRILLA(2, 16) = "4"
    FORMATOGRILLA(2, 17) = "4"
    FORMATOGRILLA(2, 18) = "4"
    FORMATOGRILLA(2, 19) = "4"
    FORMATOGRILLA(2, 20) = "4"
    FORMATOGRILLA(2, 21) = "4"
    FORMATOGRILLA(2, 22) = "4"
    FORMATOGRILLA(2, 23) = "3"
    FORMATOGRILLA(2, 24) = "3"
    FORMATOGRILLA(2, 25) = "10"
    FORMATOGRILLA(2, 26) = "2"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "N"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "N"
    FORMATOGRILLA(3, 15) = "N"
    FORMATOGRILLA(3, 16) = "N"
    FORMATOGRILLA(3, 17) = "N"
    FORMATOGRILLA(3, 18) = "N"
    FORMATOGRILLA(3, 19) = "N"
    FORMATOGRILLA(3, 20) = "N"
    FORMATOGRILLA(3, 21) = "N"
    FORMATOGRILLA(3, 22) = "N"
    FORMATOGRILLA(3, 23) = "N"
    FORMATOGRILLA(3, 24) = "N"
    FORMATOGRILLA(3, 25) = "C"
    FORMATOGRILLA(3, 26) = "C"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = ""
    FORMATOGRILLA(4, 17) = ""
    
    
    Rem LOCCKED
    For k = 1 To 26
        FORMATOGRILLA(5, k) = "FALSE"
    Next k
    
    
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    
'   grid3.BackColorFixed = RGB(90, 158, 214)
'   grid3.BackColorFixedSel = RGB(110, 180, 230)
'   grid3.BackColorBkg = RGB(90, 158, 214)
'   grid3.BackColorScrollBar = RGB(231, 235, 247)
'   grid3.BackColor1 = RGB(231, 235, 247)
'   grid3.BackColor2 = RGB(239, 243, 255)
'   grid3.GridColor = RGB(148, 190, 231)
    Grid3.Column(0).Width = 0
    
    For k = 1 To Grid3.Cols - 1
        
        Grid3.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid3.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid3.DefaultFont.Size
        Grid3.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid3.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid3.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid3.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "C" Then Grid3.Column(k).Alignment = cellCenterGeneral
        If FORMATOGRILLA(3, k) = "D" Then Grid3.Column(k).CellType = cellCalendar
        
    Next k
    
    Call primercuadro2
    Call segundocuadro2
    Call tercercuadro2
    Call cuartocuadro2
    Call quintocuadro2
     Call sextocuadro2
    Grid3.Refresh
     
    
End Sub

Private Sub leercertificado(rut, NOMBRE, numero)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim j As Double
    
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim total3 As Double
    Dim total4 As Double
    Dim tila3 As Double
    Dim ipc As Double
    Dim corre1 As Double
    Dim corre2 As Double
    Dim corre3 As Double
    Dim corre4 As Double
    Dim corre5 As Double
    Dim total5 As Double
    Dim total6 As Double
    Dim total7 As Double
    Dim total8 As Double
    Dim total9 As Double
    Dim total10 As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
        
        Set csql.ActiveConnection = contadb
'        csql.sql = "select codigo,sum(monto) from " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones where  año ='" + COMBOAÑO.text + "' and rut='" + rut + "' group by mes,codigo order by mes  "
'        csql.sql = "select codigo,sum(monto) from " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones where  año ='" + COMBOAÑO.text + "' and rut='0088977246' group by mes,codigo order by mes  "
        
        
        csql.sql = "select codigo,SUM(if(codigo='THI01',monto,0)) AS dos, "
        csql.sql = csql.sql & "SUM(if(codigo='AFP01',monto,0)), "
        csql.sql = csql.sql & "SUM(if(codigo='ISA03',monto,0))+ SUM(if(codigo='ISA01',monto,0)), "
        csql.sql = csql.sql & "SUM(if(codigo='IRE01',monto,0)), "
        csql.sql = csql.sql & "SUM(IF(mid(codigo,1,2)='HN',monto,0))+SUM(IF(mid(codigo,1,2)='FN',monto,0)),mes, "
        csql.sql = csql.sql & "SUM(if(mid(codigohd,1,2)='ST',monto,0))+ SUM(if(mid(codigohd,1,1)='P',monto,0)) "
        
        csql.sql = csql.sql & "from " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones where  año ='" & COMBOAÑO.text & "' "
        csql.sql = csql.sql & "and rut='" & rut & "' group by mes order by mes "
'        csql.sql = csql.sql & "and rut='0127433577' group by mes order by mes "
 
 
        csql.Execute
        Grid2.Rows = 22
        For j = 1 To 12
        Grid2.Cell(j, 1).text = MonthName(j)
        Grid2.Cell(j, 2).text = "0"
        Grid2.Cell(j, 3).text = "0"
        Grid2.Cell(j, 4).text = "0"
        Grid2.Cell(j, 5).text = "0"
        Grid2.Cell(j, 6).text = "0"
        Grid2.Cell(j, 7).text = "0"
        Grid2.Cell(j, 8).text = "0"
        
        Grid2.Cell(j, 9).text = "0"
        Grid2.Cell(j, 10).text = 1 + (leeripc(Format(j, "00"), COMBOAÑO.text) / 100)
        Grid2.Cell(j, 11).text = "0"
        Grid2.Cell(j, 12).text = "0"
        Grid2.Cell(j, 13).text = "0"
        Grid2.Cell(j, 14).text = "0"
        Grid2.Cell(j, 15).text = "0"
         Grid2.Cell(j, 16).text = "0"
        Next j
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
             LINEA = resultados("mes")
             ipc = 1 + (leeripc(resultados("mes"), COMBOAÑO.text) / 100)
             Grid2.Cell(LINEA, 1).text = MonthName(resultados("mes"))
             Grid2.Cell(LINEA, 2).text = resultados(1)
             
             salud = resultados(3)
             
             If salud > 4.2 * leerUFmes(Format(LINEA, "00"), COMBOAÑO.text) Then
             salud = Round(4.2 * leerUFmes(Format(LINEA, "00"), COMBOAÑO.text), 0)
             End If
             
             
             previ = resultados(2) + salud + resultados(7)
             
             Grid2.Cell(LINEA, 3).text = previ
             Grid2.Cell(LINEA, 4).text = resultados(1) - previ
             Grid2.Cell(LINEA, 5).text = resultados(4)
             
             corre1 = Round(Val(Grid2.Cell(LINEA, 4).text) * ipc, 0)
             corre2 = Round(Val(Grid2.Cell(LINEA, 5).text) * ipc, 0)
             
             
             Grid2.Cell(LINEA, 6).text = "0"
             Grid2.Cell(LINEA, 8).text = resultados(5)
             corre3 = Round(Val(Grid2.Cell(LINEA, 8).text) * ipc, 0)
             Grid2.Cell(LINEA, 7).text = "0"
             Grid2.Cell(LINEA, 9).text = "0"
             Grid2.Cell(LINEA, 10).text = ipc
             Grid2.Cell(LINEA, 11).text = corre1
             Grid2.Cell(LINEA, 12).text = corre2
             Grid2.Cell(LINEA, 13).text = "0"
             Grid2.Cell(LINEA, 14).text = "0"
             Grid2.Cell(LINEA, 15).text = corre3
             Grid2.Cell(LINEA, 16).text = "0"
             
             
             
             
             total = total + resultados(1)
             total2 = total2 + previ
             total3 = total3 + resultados(1) - previ
             total4 = total4 + resultados(4)
             total5 = total5 + resultados(5)
             total6 = total6 + corre1
             total7 = total7 + corre2
             total8 = 0
             total9 = total9 + corre3
             total10 = 0
             resultados.MoveNext
            Wend
             LINEA = 13
             
             Grid2.Range(LINEA, 1, LINEA, 14).FontBold = True
             Grid2.Range(LINEA, 1, LINEA, 14).Borders(cellEdgeTop) = cellThin
             Grid2.Cell(LINEA, 1).text = "TOTALES"
             Grid2.Cell(LINEA, 2).text = total
             Grid2.Cell(LINEA, 3).text = total2
             Grid2.Cell(LINEA, 4).text = total3
             Grid2.Cell(LINEA, 5).text = total4
             Grid2.Cell(LINEA, 6).text = "0"
             Grid2.Cell(LINEA, 7).text = "0"
             Grid2.Cell(LINEA, 8).text = total5
             Grid2.Cell(LINEA, 9).text = "0"
             
             Grid2.Cell(LINEA, 11).text = total6
             Grid2.Cell(LINEA, 12).text = total7
             Grid2.Cell(LINEA, 13).text = total8
             Grid2.Cell(LINEA, 14).text = "0"
             Grid2.Cell(LINEA, 15).text = total9
             
             Grid2.Cell(LINEA, 16).text = total10
         
            resultados.Close
            Set resultados = Nothing
            
            Grid2.Range(15, 1, 15, Grid2.Cols - 1).Merge
            Grid2.Range(15, 1, 15, Grid2.Cols - 1).FontSize = 8
            
            
            Grid2.Range(16, 1, 16, Grid2.Cols - 1).Merge
            Grid2.Range(16, 1, 16, Grid2.Cols - 1).FontSize = 8
            
'            Grid2.Range(17, 1, 17, Grid2.Cols - 1).Merge
'            Grid2.Range(17, 1, 17, Grid2.Cols - 1).FontSize = 11
           
            Grid2.Cell(15, 1).text = "Se extiende el presente certificado en cumplimiento de lo dispuesto en la Resolucion Ex Nro 6509 del Servicio de Impuestos Internos , publicada en el Diario Oficial de fecha 20 de Diciembre de 1993 y sus modificaciones"
            Grid2.Cell(16, 1).text = "posteriores. "
            
            Grid2.Range(20, 12, 20, Grid2.Cols - 1).Merge
            Grid2.Range(20, 12, 20, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
            Grid2.Cell(20, 12).Alignment = cellCenterCenter
            
            Grid2.Cell(20, 12).text = FIRMA.text
            


End If

      
End Sub

Private Sub calculacertificado(rut, row)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim j As Double
    
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim total3 As Double
    Dim total4 As Double
    Dim tila3 As Double
    Dim ipc As Double
    Dim corre1 As Double
    Dim corre2 As Double
    Dim corre3 As Double
    Dim corre4 As Double
    Dim corre5 As Double
    Dim total5 As Double
    Dim total6 As Double
    Dim total7 As Double
    Dim total8 As Double
    Dim total9 As Double
    Dim total10 As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
        
        Set csql.ActiveConnection = contadb
'        csql.sql = "select codigo,sum(monto) from " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones where  año ='" + COMBOAÑO.text + "' and rut='" + rut + "' group by mes,codigo order by mes  "
'        csql.sql = "select codigo,sum(monto) from " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones where  año ='" + COMBOAÑO.text + "' and rut='0088977246' group by mes,codigo order by mes  "
        
        
        csql.sql = "select codigo,SUM(if(codigo='THI01',monto,0)) AS dos, "
        csql.sql = csql.sql & "SUM(if(codigo='AFP01',monto,0)), "
        csql.sql = csql.sql & "SUM(if(codigo='ISA03',monto,0))+ SUM(if(codigo='ISA01',monto,0)), "
        csql.sql = csql.sql & "SUM(if(codigo='IRE01',monto,0)), "
        csql.sql = csql.sql & "SUM(IF(mid(codigo,1,2)='HN',monto,0))+SUM(IF(mid(codigo,1,2)='FN',monto,0)),mes, "
        csql.sql = csql.sql & "SUM(if(mid(codigohd,1,2)='ST',monto,0))+ SUM(if(mid(codigohd,1,1)='P',monto,0)) "
        csql.sql = csql.sql & "from " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones where  año ='" & COMBOAÑO.text & "' "
        csql.sql = csql.sql & "and rut='" & rut & "' group by mes order by mes "
'        csql.sql = csql.sql & "and rut='0127433577' group by mes order by mes "
 
        sincorre1 = 0
        sincorre2 = 0
        sincorre3 = 0
        csql.Execute
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
             LINEA = resultados("mes")
             ipc = 1 + (leeripc(resultados("mes"), COMBOAÑO.text) / 100)
             salud = resultados(3)
             
             If salud > 4.2 * leerUFmes(Format(LINEA, "00"), COMBOAÑO.text) Then
             salud = Round(4.2 * leerUFmes(Format(LINEA, "00"), COMBOAÑO.text), 0)
             End If
             previ = resultados(2) + salud + resultados(7)
             
             sincorre1 = sincorre1 + Round((resultados(1)) * ipc, 0)
             sincorre2 = sincorre2 + Round(resultados(4), 0)
             sincorre3 = sincorre3 + Round(resultados(5), 0)
             
             
             corre1 = Round((resultados(1) - previ) * ipc, 0)
             corre2 = Round(resultados(4) * ipc, 0)
             corre3 = Round(resultados(5) * ipc, 0)
             
             total = total + resultados(1)
             total2 = total2 + previ
             total3 = total3 + resultados(1) - previ
             total4 = total4 + resultados(4)
             total5 = total5 + resultados(5)
             total6 = total6 + corre1
             total7 = total7 + corre2
             total8 = 0
             total9 = total9 + corre3
             total10 = 0
             resultados.MoveNext
            Wend
             totalsincorre1 = totalsincorre1 + sincorre1
             totalsincorre2 = totalsincorre2 + total6
             totalsincorre3 = totalsincorre3 + sincorre3
             totalsincorre4 = totalsincorre4 + total3
             totalsincorre5 = totalsincorre5 + sincorre2
            
             
             Grid1.Cell(row, 3).text = total6
             Grid1.Cell(row, 4).text = total7
             Grid1.Cell(row, 6).text = total9
             Grid1.Cell(row, 7).text = 0
            
            resultados.Close
            Set resultados = Nothing
            
End If

      
End Sub


Sub cabezas3(rut, NOMBRE, numero)
Dim objReportTitle As FlexCell.ReportTitle
Grid2.ReportTitles.Clear



    'Report Title 1
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid2.ReportTitles.Add objReportTitle
    Next k

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CERTIFICADO N° 6 SOBRE SUELDOS Y OTRAS RENTAS SIMILARES"
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "N° " & numero
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellRight
    Grid2.ReportTitles.Add objReportTitle
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "  " & DATOSEMPRESA(3) & ",  " & Format(fechasistema, "dd") & " de " & MonthName(Format(fechasistema, "mm")) & " del " & Format(fechasistema, "yyyy")
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellRight
    Grid2.ReportTitles.Add objReportTitle
    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellRight
    Grid2.ReportTitles.Add objReportTitle



Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "El Empleador, Habilitado o Pagador," & DATOSEMPRESA(1) & ", certifica que el Sr. " & NOMBRE & " RUT N° " & Format(Mid(rut, 1, 9), "###,###,###") + "-" + Mid(rut, 10, 1) & ", "
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 8.5
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid2.ReportTitles.Add objReportTitle
    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "en su calidad de empleado dependiente, durante el año " & COMBOAÑO.text & ", se le han pagado las las rentas que se indican y sobre las cuales se le practicaron las retenciones de impuestos que se señalan:"
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 8.5
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid2.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 8
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid2.ReportTitles.Add objReportTitle


With Grid2.PageSetup
        
        Rem If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub

Private Sub Grid1_Click()
    If Grid1.ActiveCell.col = 22 Then
        If Grid1.Cell(Grid1.ActiveCell.row, 22).text = "1" Then
            Grid1.Cell(Grid1.ActiveCell.row, 22).text = "0"
        Else
            Grid1.Cell(Grid1.ActiveCell.row, 22).text = "1"
        End If
    End If
End Sub

Public Function leercargasvigentes(rut, empresa, fecha) As Double
    Dim csql As New rdoQuery
    Dim resultados  As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select count(rutcarga) from " + clientesistema + "remu" + empresa + ".re_cargafamiliares "
    csql.sql = csql.sql & "where rut='" & rut & "' and (fechavencimiento>='" & Format(fecha, "yyyy-mm") & "-01'  or fechavencimiento='1111-11-11') and fechaingreso<='" & Format(fecha, "yyyy-mm-dd") & "' "
    csql.Execute
    leercargasvigentes = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leercargasvigentes = resultados(0)
    End If
    csql.Close
    Set resultados = Nothing
    Set csql = Nothing
End Function
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Private Sub opt1_Click()
    If opt1.Value = True Then
        frameprimerahoja.Visible = True
        frmsegundahoja.Visible = False
    End If
End Sub

Private Sub opt2_Click()
    If opt2.Value = True Then
        frameprimerahoja.Visible = False
        frmsegundahoja.Visible = True
    End If
End Sub
Sub primercuadro2()

    Grid3.Range(1, 2, 1, 16).Merge
    Grid3.Range(1, 2, 1, 16).Alignment = cellCenterGeneral
    Grid3.Cell(1, 2).text = "SISTEMA DE TRIBUTACIÓN SIMPLIFICADA DEL IVA, ART. 29 D.L. 825"
    Grid3.Cell(1, 2).Font.Bold = True
    Grid3.Range(1, 17, 1, 21).Merge
    Grid3.Range(1, 22, 1, 26).Merge
    Grid3.Cell(1, 22).Alignment = cellCenterGeneral
    Grid3.Cell(1, 22).text = "IMPUESTO DETERMINADO"
    Grid3.Cell(1, 22).Font.Bold = True
     
    
    Grid3.Cell(2, 2).text = "60"
    Grid3.Range(2, 3, 2, 8).Merge
    Grid3.Cell(2, 3).text = "Ventas del período"
    Grid3.Cell(2, 3).Alignment = cellLeftGeneral
    Grid3.Cell(2, 9).text = "529"
    Grid3.Range(2, 10, 2, 16).Merge
    Grid3.Range(2, 17, 2, 26).Merge
    
    Grid3.Cell(3, 2).text = "61"
    Grid3.Range(3, 3, 3, 8).Merge
    Grid3.Cell(3, 3).text = " Crédito del período"
    Grid3.Cell(3, 3).Alignment = cellLeftGeneral
    Grid3.Cell(3, 9).text = "530"
    Grid3.Range(3, 10, 3, 16).Merge
    Grid3.Range(3, 17, 3, 26).Merge
   
    Grid3.Cell(4, 2).text = "62"
    Grid3.Range(4, 3, 4, 21).Merge
    Grid3.Cell(4, 3).text = " IVA determinado por concepto de Tributación Simplificada"
    Grid3.Cell(4, 3).Alignment = cellLeftGeneral
    Grid3.Cell(4, 22).text = "409"
    Grid3.Range(4, 23, 4, 25).Merge
    Grid3.Cell(4, 26).text = "+"
    Grid3.Range(5, 1, 5, 26).Merge
    
    Grid3.Cell(6, 2).text = "63"
    Grid3.Range(6, 3, 6, 21).Merge
    Grid3.Cell(6, 3).text = " Letras e), h), i), l) (tasa 15%)"
    Grid3.Cell(6, 3).Alignment = cellLeftGeneral
    Grid3.Cell(6, 22).text = "522"
    Grid3.Range(6, 23, 6, 25).Merge
    Grid3.Cell(6, 26).text = "+"
    
    Grid3.Cell(7, 2).text = "64"
    Grid3.Range(7, 3, 7, 21).Merge
    Grid3.Cell(7, 3).text = " Letra j) (tasa 50%)"
    Grid3.Cell(7, 3).Alignment = cellLeftGeneral
    Grid3.Cell(7, 22).text = "526"
    Grid3.Range(7, 23, 7, 25).Merge
    Grid3.Cell(7, 26).text = "+"
    
    Grid3.Range(8, 2, 8, 26).Merge
    
    Grid3.Cell(9, 2).text = "65"
    Grid3.Range(9, 3, 9, 13).Merge
    Grid3.Cell(9, 3).text = " Débito de Impuesto Adicional Ventas Art. 37 letras a), b) y c) y Art. 40 D.L. 825 (tasa 15%)"
    Grid3.Cell(9, 3).Alignment = cellLeftGeneral
    Grid3.Cell(9, 14).text = "113"
    Grid3.Range(9, 15, 9, 19).Merge
    Grid3.Cell(9, 20).text = "+"
    Grid3.Cell(9, 20).Alignment = cellCenterGeneral
    Grid3.Range(9, 21, 9, 22).Merge
    
    Grid3.Cell(10, 2).text = "66"
    Grid3.Range(10, 3, 10, 13).Merge
    Grid3.Cell(10, 3).text = " Crédito de Impuesto Adicional Art. 37 letras a), b) y c) D.L. 825"
    Grid3.Cell(10, 3).Alignment = cellLeftGeneral
    Grid3.Cell(10, 14).text = "28"
    Grid3.Range(10, 15, 10, 19).Merge
    Grid3.Cell(10, 20).text = "-"
    Grid3.Cell(10, 20).Alignment = cellCenterGeneral
    Grid3.Range(10, 21, 10, 22).Merge
    
    Grid3.Cell(11, 2).text = "67"
    Grid3.Range(11, 3, 11, 13).Merge
    Grid3.Cell(11, 3).text = " Monto reintegrado por devolución indebida de crédito por exportadores D.L. 825"
    Grid3.Cell(11, 3).Alignment = cellLeftGeneral
    Grid3.Cell(11, 14).text = "548"
    Grid3.Range(11, 15, 11, 19).Merge
    Grid3.Cell(11, 20).text = "-"
    Grid3.Cell(11, 20).Alignment = cellCenterGeneral
    Grid3.Range(11, 21, 11, 22).Merge
    
    Grid3.Cell(12, 2).text = "68"
    Grid3.Range(12, 3, 12, 13).Merge
    Grid3.Cell(12, 3).text = " Remanente crédito Art. 37 mes anterior D.L.825"
    Grid3.Cell(12, 3).Alignment = cellLeftGeneral
    Grid3.Cell(12, 14).text = "540"
    Grid3.Range(12, 15, 12, 19).Merge
    Grid3.Cell(12, 20).text = "-"
    Grid3.Cell(12, 20).Alignment = cellCenterGeneral
    Grid3.Range(12, 21, 12, 22).Merge
    
    Grid3.Cell(13, 2).text = "69"
    Grid3.Range(13, 3, 13, 13).Merge
    Grid3.Cell(13, 3).text = " Devolución Solicitud Art. 36 relativa al Impuesto Adicional Art. 37 letras a), b) y c) D.L. 825"
    Grid3.Cell(13, 3).Alignment = cellLeftGeneral
    Grid3.Cell(13, 14).text = "541"
    Grid3.Range(13, 15, 13, 19).Merge
    Grid3.Cell(13, 20).text = "+"
    Grid3.Cell(13, 20).Alignment = cellCenterGeneral
    Grid3.Range(13, 21, 13, 22).Merge
    
   
    Grid3.Range(9, 21, 13, 25).Merge
    Grid3.Range(9, 26, 13, 26).Merge
    Grid3.Cell(9, 21).text = "Diferencia Débito menos Crédito Impuesto Art. 37 D.L. 825/74 (operación aritmética de las líneas 65 a la 69), traslade el valor a la línea 70. Si el resultado es positivo al código 550, en caso contrario al código 549 sin signo.   "
    Grid3.Cell(9, 21).WrapText = True
    Grid3.Cell(9, 21).Font.Size = 7
    Grid3.Cell(9, 21).Alignment = cellLeftGeneral
    
    Grid3.Range(14, 1, 14, 26).Merge
    
    Grid3.Cell(15, 2).text = "70"
    Grid3.Range(15, 3, 15, 8).Merge
    Grid3.Cell(15, 3).text = " Remanente crédito impuesto Art. 37 para período siguiente"
    Grid3.Cell(15, 3).Alignment = cellLeftGeneral
    Grid3.Cell(15, 9).text = "549"
    Grid3.Range(15, 10, 15, 13).Merge
    Grid3.Range(15, 14, 15, 21).Merge
    Grid3.Cell(15, 22).text = "550"
    Grid3.Cell(15, 22).Alignment = cellCenterGeneral
    Grid3.Range(15, 23, 15, 25).Merge
    Grid3.Cell(15, 26).text = "+"
    
     Grid3.Range(16, 1, 16, 20).Merge
    
    
End Sub
Sub segundocuadro2()
       Grid3.Range(16, 21, 31, 26).Merge
        
        Grid3.Range(17, 2, 17, 13).Merge
        Grid3.Cell(17, 14).text = " Débitos"
        Grid3.Cell(17, 14).Alignment = cellCenterGeneral
        Grid3.Range(17, 14, 17, 20).Merge
       
 
        Grid3.Cell(18, 2).text = "71"
        Grid3.Range(18, 3, 18, 13).Merge
        Grid3.Cell(18, 3).text = " Pisco, Licores, Whisky y Aguardiente (tasa 27%)"
        Grid3.Cell(18, 3).Alignment = cellLeftGeneral
        Grid3.Cell(18, 14).text = "577"
        Grid3.Range(18, 15, 18, 19).Merge
        Grid3.Cell(18, 20).text = "+"
        Grid3.Cell(18, 20).Alignment = cellCenterGeneral

        Grid3.Cell(19, 2).text = "72"
        Grid3.Range(19, 3, 19, 13).Merge
        Grid3.Cell(19, 3).text = " Vinos, Champaña, Chichas (tasa 15%))"
        Grid3.Cell(19, 3).Alignment = cellLeftGeneral
        Grid3.Cell(19, 14).text = "32"
        Grid3.Range(19, 15, 19, 19).Merge
        Grid3.Cell(19, 20).text = "+"
        Grid3.Cell(19, 20).Alignment = cellCenterGeneral

        Grid3.Cell(20, 2).text = "73"
        Grid3.Range(20, 3, 20, 13).Merge
        Grid3.Cell(20, 3).text = " Cervezas (tasa 15%)"
        Grid3.Cell(20, 3).Alignment = cellLeftGeneral
        Grid3.Cell(20, 14).text = "150"
        Grid3.Range(20, 15, 20, 19).Merge
        Grid3.Cell(20, 20).text = "+"
        Grid3.Cell(20, 20).Alignment = cellCenterGeneral

        Grid3.Cell(21, 2).text = "74"
        Grid3.Range(21, 3, 21, 13).Merge
        Grid3.Cell(21, 3).text = " Bebidas analcohólicas (tasa 13%)"
        Grid3.Cell(21, 3).Alignment = cellLeftGeneral
        Grid3.Cell(21, 14).text = "146"
        Grid3.Range(21, 15, 21, 19).Merge
        Grid3.Cell(21, 20).text = "+"
        Grid3.Cell(21, 20).Alignment = cellCenterGeneral

        Grid3.Cell(22, 2).text = "75"
        Grid3.Range(22, 3, 22, 13).Merge
        Grid3.Cell(22, 3).text = " Notas de Débito emitidas"
        Grid3.Cell(22, 3).Alignment = cellLeftGeneral
        Grid3.Cell(22, 14).text = "545"
        Grid3.Range(22, 15, 22, 19).Merge
        Grid3.Cell(22, 20).text = "+"
        Grid3.Cell(22, 20).Alignment = cellCenterGeneral

        Grid3.Cell(23, 2).text = "76"
        Grid3.Range(23, 3, 23, 13).Merge
        Grid3.Cell(23, 3).text = " Notas de Crédito emitidas por Facturas"
        Grid3.Cell(23, 3).Alignment = cellLeftGeneral
        Grid3.Cell(23, 14).text = "546"
        Grid3.Range(23, 15, 23, 19).Merge
        Grid3.Cell(23, 20).text = "-"
        Grid3.Cell(23, 20).Alignment = cellCenterGeneral

        Grid3.Cell(24, 2).text = "77"
        Grid3.Range(24, 3, 24, 13).Merge
        Grid3.Cell(24, 3).text = " Notas de Crédito emitidas por Vales de máquinas autorizadas por el Servicio"
        Grid3.Cell(24, 3).Alignment = cellLeftGeneral
        Grid3.Cell(24, 14).text = "710"
        Grid3.Range(24, 15, 24, 19).Merge
        Grid3.Cell(24, 20).text = "-"
        Grid3.Cell(24, 20).Alignment = cellCenterGeneral

        Grid3.Cell(25, 2).text = "78"
        Grid3.Range(25, 3, 25, 13).Merge
        Grid3.Cell(25, 3).text = " Total Débitos Art. 42 D.L. 825"
        Grid3.Cell(25, 3).Alignment = cellLeftGeneral
        Grid3.Cell(25, 3).Font.Bold = True
        Grid3.Cell(25, 14).text = "602"
        Grid3.Range(25, 15, 25, 19).Merge
        Grid3.Cell(25, 20).text = "="
        Grid3.Cell(25, 20).Alignment = cellCenterGeneral

        Grid3.Range(26, 1, 26, 20).Merge
        
        
        Grid3.Range(27, 2, 27, 8).Merge
        Grid3.Range(27, 9, 27, 13).Merge
        Grid3.Cell(27, 9).text = " Total crédito recargado en facturas recibidas"
        Grid3.Cell(27, 9).Alignment = cellCenterGeneral
        Grid3.Cell(27, 9).Font.Size = 7
        Grid3.Range(27, 14, 27, 20).Merge
        Grid3.Cell(27, 14).text = "Crédito imputable del período"
        Grid3.Cell(27, 14).Alignment = cellCenterGeneral
        
        Grid3.Cell(28, 2).text = "79"
        Grid3.Range(28, 3, 28, 8).Merge
        Grid3.Cell(28, 3).text = " Pisco, Licores, Whisky y Aguardiente (tasa 27%)"
        Grid3.Cell(28, 3).Alignment = cellLeftGeneral
        Grid3.Cell(28, 9).text = "575"
        Grid3.Range(28, 10, 28, 13).Merge
        Grid3.Cell(28, 14).text = "576"
        Grid3.Range(28, 15, 28, 19).Merge
        Grid3.Cell(28, 20).text = "+"
        Grid3.Cell(28, 20).Alignment = cellCenterGeneral
        
        Grid3.Cell(29, 2).text = "80"
        Grid3.Range(29, 3, 29, 8).Merge
        Grid3.Cell(29, 3).text = " Vinos, Champaña, Chichas (tasa 15%)"
        Grid3.Cell(29, 3).Alignment = cellLeftGeneral
        Grid3.Cell(29, 9).text = "574"
        Grid3.Range(29, 10, 29, 13).Merge
        Grid3.Cell(29, 14).text = "33"
        Grid3.Range(29, 15, 28, 19).Merge
        Grid3.Cell(29, 20).text = "+"
        Grid3.Cell(29, 20).Alignment = cellCenterGeneral
        
        Grid3.Cell(30, 2).text = "81"
        Grid3.Range(30, 3, 30, 8).Merge
        Grid3.Cell(30, 3).text = " Cervezas (tasa 15%)"
        Grid3.Cell(30, 3).Alignment = cellLeftGeneral
        Grid3.Cell(30, 9).text = "580"
        Grid3.Range(30, 10, 30, 13).Merge
        Grid3.Cell(30, 14).text = "149"
        Grid3.Range(30, 15, 30, 19).Merge
        Grid3.Cell(30, 20).text = "+"
        Grid3.Cell(30, 20).Alignment = cellCenterGeneral
        
        Grid3.Cell(31, 2).text = "82"
        Grid3.Range(31, 3, 31, 8).Merge
        Grid3.Cell(31, 3).text = " Bebidas analcohólicas (tasa 13%)"
        Grid3.Cell(31, 3).Alignment = cellLeftGeneral
        Grid3.Cell(31, 9).text = "582"
        Grid3.Range(31, 10, 31, 13).Merge
        Grid3.Cell(31, 14).text = "85"
        Grid3.Range(31, 15, 31, 19).Merge
        Grid3.Cell(31, 20).text = "+"
        Grid3.Cell(31, 20).Alignment = cellCenterGeneral
        
        Grid3.Cell(32, 2).text = "83"
        Grid3.Range(32, 3, 32, 13).Merge
        Grid3.Cell(32, 3).text = " Notas de Débito recibidas"
        Grid3.Cell(32, 3).Alignment = cellLeftGeneral
        Grid3.Cell(32, 14).text = "551"
        Grid3.Range(32, 15, 32, 19).Merge
        Grid3.Cell(32, 20).text = "+"
        Grid3.Cell(32, 20).Alignment = cellCenterGeneral
        
        Grid3.Cell(33, 2).text = "84"
        Grid3.Range(33, 3, 33, 13).Merge
        Grid3.Cell(33, 3).text = " Notas de Crédito recibidas"
        Grid3.Cell(33, 3).Alignment = cellLeftGeneral
        Grid3.Cell(33, 14).text = "559"
        Grid3.Range(33, 15, 33, 19).Merge
        Grid3.Cell(33, 20).text = "-"
        Grid3.Cell(33, 20).Alignment = cellCenterGeneral
        
        Grid3.Cell(34, 2).text = "85"
        Grid3.Range(34, 3, 34, 13).Merge
        Grid3.Cell(34, 3).text = " Remanente crédito Art. 42 mes anterior"
        Grid3.Cell(34, 3).Alignment = cellLeftGeneral
        Grid3.Cell(34, 14).text = "508"
        Grid3.Range(34, 15, 34, 19).Merge
        Grid3.Cell(34, 20).text = "+"
        Grid3.Cell(34, 20).Alignment = cellCenterGeneral
        
        Grid3.Cell(35, 2).text = "86"
        Grid3.Range(35, 3, 35, 13).Merge
        Grid3.Cell(35, 3).text = " Devolución Art. 36 D.L. 825 relativas impuesto Art. 42"
        Grid3.Cell(35, 3).Alignment = cellLeftGeneral
        Grid3.Cell(35, 14).text = "533"
        Grid3.Range(35, 15, 35, 19).Merge
        Grid3.Cell(35, 20).text = "-"
        Grid3.Cell(35, 20).Alignment = cellCenterGeneral
        
        Grid3.Cell(36, 2).text = "87"
        Grid3.Range(36, 3, 36, 13).Merge
        Grid3.Cell(36, 3).text = " Monto reintegrado devoluciones indebidas de crédito por exportaciones"
        Grid3.Cell(36, 3).Alignment = cellLeftGeneral
        Grid3.Cell(36, 14).text = "552"
        Grid3.Range(36, 15, 36, 19).Merge
        Grid3.Cell(36, 20).text = "+"
        Grid3.Cell(36, 20).Alignment = cellCenterGeneral
        
        
        Grid3.Cell(37, 2).text = "88"
        Grid3.Range(37, 3, 37, 13).Merge
        Grid3.Cell(37, 3).text = " Total créditos Art. 42 D.L. 825"
        Grid3.Cell(37, 3).Font.Bold = True
        Grid3.Cell(37, 3).Alignment = cellLeftGeneral
        Grid3.Cell(37, 14).text = "603"
        Grid3.Range(37, 15, 37, 19).Merge
        Grid3.Range(37, 21, 37, 26).Merge
        Grid3.Cell(37, 20).text = "+"
        Grid3.Cell(37, 20).Alignment = cellCenterGeneral
        
        Grid3.Range(32, 21, 36, 25).Merge
        Grid3.Cell(32, 21).text = "Diferencia Débito menos Crédito Impuesto Art. 42 D.L. 825/74 (código 602 menos el código 603), traslade el valor a la línea 89. Si el resultado es positivo al código 506, en caso contrario al código 507 sin signo."
        Grid3.Cell(32, 21).WrapText = True
        Grid3.Cell(32, 21).Alignment = cellLeftCenter
        Grid3.Range(32, 26, 36, 26).Merge
        
        Grid3.Range(38, 2, 38, 26).Merge
        
        
        
        Grid3.Cell(39, 2).text = "89"
        Grid3.Range(39, 3, 39, 8).Merge
        Grid3.Cell(39, 3).text = " Remanente crédito imp. Adic. Art. 42 para período siguiente"
        Grid3.Cell(39, 3).Alignment = cellLeftGeneral
        Grid3.Cell(39, 9).text = "507"
        Grid3.Range(39, 10, 39, 13).Merge
        Grid3.Cell(39, 14).text = "Impuesto Adicional Art. 42 determinado"
        Grid3.Cell(39, 14).Alignment = cellLeftGeneral
        Grid3.Range(39, 14, 39, 21).Merge
        Grid3.Cell(39, 22).text = "506"
        Grid3.Range(39, 23, 39, 25).Merge
        Grid3.Cell(39, 26).text = "+"
        Grid3.Cell(39, 26).Alignment = cellCenterGeneral
        
        Grid3.Range(40, 1, 40, 26).Merge
    
End Sub


Sub tercercuadro2()
       
        Grid3.Range(41, 2, 41, 14).Merge
        Grid3.Cell(41, 2).text = " ANTICIPO CAMBIO DE SUJETO (CONTRIBUYENTES RETENIDOS)"
        Grid3.Cell(41, 2).Alignment = cellLeftGeneral
        Grid3.Cell(41, 2).Font.Bold = True
        Grid3.Range(41, 15, 41, 26).Merge
        
        Grid3.Cell(42, 2).text = "90"
        Grid3.Range(42, 3, 42, 8).Merge
        Grid3.Cell(42, 3).text = " IVA anticipado del período"
        Grid3.Cell(42, 3).Alignment = cellLeftGeneral
        Grid3.Cell(42, 9).text = "556"
        Grid3.Range(42, 10, 42, 13).Merge
        Grid3.Cell(42, 14).text = "+"
        Grid3.Cell(42, 14).Alignment = cellCenterGeneral
        
        Grid3.Cell(43, 2).text = "91"
        Grid3.Range(43, 3, 43, 8).Merge
        Grid3.Cell(43, 3).text = " Remanente del mes anterior"
        Grid3.Cell(43, 3).Alignment = cellLeftGeneral
        Grid3.Cell(43, 9).text = "557"
        Grid3.Range(43, 10, 43, 13).Merge
        Grid3.Cell(43, 14).text = "+"
        Grid3.Cell(43, 14).Alignment = cellCenterGeneral
        
        Grid3.Cell(44, 2).text = "92"
        Grid3.Range(44, 3, 44, 8).Merge
        Grid3.Cell(44, 3).text = " Devolución del mes anterior"
        Grid3.Cell(44, 3).Alignment = cellLeftGeneral
        Grid3.Cell(44, 9).text = "558"
        Grid3.Range(44, 10, 44, 13).Merge
        Grid3.Cell(44, 14).text = "-"
        Grid3.Cell(44, 14).Alignment = cellCenterGeneral
        
        Grid3.Cell(45, 2).text = "93"
        Grid3.Range(45, 3, 45, 8).Merge
        Grid3.Cell(45, 3).text = " Total de Anticipo"
        Grid3.Cell(45, 3).Font.Bold = True
        Grid3.Cell(45, 3).Alignment = cellLeftGeneral
        Grid3.Cell(45, 9).text = "543"
        Grid3.Range(45, 10, 45, 13).Merge
        Grid3.Cell(45, 14).text = "="
        Grid3.Cell(45, 14).Alignment = cellCenterGeneral
        
        Grid3.Range(42, 21, 45, 25).Merge
        Grid3.Cell(42, 21).text = "Registre Total de anticipo (código 543) en el código 598, con tope del valor del código 89 línea 43, el saldo restante se debe registrar en el remanente para el mes siguiente, código 573."
        Grid3.Cell(42, 21).WrapText = True
        Grid3.Cell(42, 21).Alignment = cellCenterGeneral
        
        Grid3.Range(42, 15, 45, 20).Merge
        
        Grid3.Range(46, 2, 46, 26).Merge
        
        Grid3.Cell(47, 2).text = "94"
        Grid3.Range(47, 3, 47, 8).Merge
        Grid3.Cell(47, 3).text = " Remanente Anticipos Cambio Sujeto para período siguiente."
        Grid3.Cell(47, 3).Alignment = cellLeftGeneral
        Grid3.Cell(47, 3).Font.Size = 7
        Grid3.Cell(47, 9).text = "573"
        Grid3.Range(47, 10, 47, 14).Merge
        Grid3.Range(47, 15, 47, 17).Merge
        Grid3.Range(47, 18, 47, 21).Merge
        Grid3.Cell(47, 18).text = "Anticipo a imputar"
        Grid3.Cell(47, 18).Alignment = cellLeftGeneral
        Grid3.Cell(47, 22).text = "598"
        Grid3.Range(47, 24, 47, 25).Merge
        Grid3.Cell(47, 26).text = "-"
        Grid3.Cell(47, 26).Alignment = cellCenterGeneral
        
        Grid3.Range(48, 2, 48, 26).Merge
        
        Grid3.Range(49, 2, 49, 14).Merge
        Grid3.Cell(49, 2).text = " CAMBIO DE SUJETO (AGENTE RETENEDOR)"
        Grid3.Cell(49, 2).Alignment = cellLeftGeneral
        Grid3.Cell(49, 2).Font.Bold = True
        Grid3.Range(49, 15, 49, 26).Merge
        
        Grid3.Cell(50, 2).text = "95"
        Grid3.Range(50, 3, 50, 8).Merge
        Grid3.Cell(50, 3).text = " IVA total retenido a terceros (tasa Art. 14 D.L. 825)"
        Grid3.Cell(50, 3).Alignment = cellLeftGeneral
        Grid3.Cell(50, 9).text = "39"
        Grid3.Range(50, 10, 50, 13).Merge
        Grid3.Cell(50, 14).text = "+"
        Grid3.Cell(50, 14).Alignment = cellCenterGeneral
        Grid3.Range(50, 15, 50, 26).Merge
        
        Grid3.Cell(51, 2).text = "96"
        Grid3.Range(51, 3, 51, 8).Merge
        Grid3.Cell(51, 3).text = " IVA parcial retenido a terceros (según tasa)"
        Grid3.Cell(51, 3).Alignment = cellLeftGeneral
        Grid3.Cell(51, 9).text = "554"
        Grid3.Range(51, 10, 51, 13).Merge
        Grid3.Cell(51, 14).text = "+"
        Grid3.Cell(51, 14).Alignment = cellCenterGeneral
        
        Grid3.Cell(52, 2).text = "97"
        Grid3.Range(52, 3, 52, 8).Merge
        Grid3.Cell(52, 3).text = " IVA Retenido por notas de crédito emitidas"
        Grid3.Cell(52, 3).Alignment = cellLeftGeneral
        Grid3.Cell(52, 9).text = "736"
        Grid3.Range(52, 10, 52, 13).Merge
        Grid3.Cell(52, 14).text = "-"
        Grid3.Cell(52, 14).Alignment = cellCenterGeneral
        Grid3.Range(51, 15, 52, 20).Merge
        
        Grid3.Range(51, 21, 52, 25).Merge
        Grid3.Cell(51, 21).text = "Registre en el código 596 la suma de las retenciones (código 39, 554, 736, 597 y 555)."
        Grid3.Cell(51, 21).WrapText = True
        Grid3.Cell(51, 21).Font.Size = 7
        Grid3.Cell(51, 21).Alignment = cellCenterGeneral
        
        
        Grid3.Cell(53, 2).text = "98"
        Grid3.Range(53, 3, 53, 8).Merge
        Grid3.Cell(53, 3).text = " Retención de margen de comercialización"
        Grid3.Cell(53, 3).Alignment = cellLeftGeneral
        Grid3.Cell(53, 9).text = "597"
        Grid3.Range(53, 10, 53, 13).Merge
        Grid3.Cell(53, 14).text = "+"
        Grid3.Cell(53, 14).Alignment = cellCenterGeneral
        Grid3.Range(53, 15, 53, 26).Merge
        
        Grid3.Cell(54, 2).text = "99"
        Grid3.Range(54, 3, 54, 8).Merge
        Grid3.Cell(54, 3).text = " Retención Anticipo de Cambio de Sujeto"
        Grid3.Cell(54, 3).Alignment = cellLeftGeneral
        Grid3.Cell(54, 9).text = "555"
        Grid3.Range(54, 10, 54, 13).Merge
        Grid3.Cell(54, 14).text = "+"
        Grid3.Cell(54, 14).Alignment = cellCenterGeneral
        Grid3.Range(54, 16, 54, 21).Merge
        Grid3.Cell(54, 16).text = " Retención Cambio de Sujeto"
        Grid3.Cell(54, 16).Font.Bold = True
        Grid3.Cell(54, 16).Alignment = cellLeftGeneral
        Grid3.Cell(54, 22).text = "596"
        Grid3.Range(54, 23, 54, 25).Merge
        Grid3.Cell(54, 26).text = "+"
        Grid3.Cell(54, 26).Alignment = cellCenterGeneral
        
         Grid3.Range(55, 1, 55, 26).Merge
        
        
       
End Sub


Sub cuartocuadro2()
        Grid3.Range(55, 1, 55, 26).Merge
        
        Grid3.Cell(56, 2).text = "100"
        Grid3.Range(56, 3, 56, 6).Merge
        Grid3.Cell(56, 3).text = " Crédito por Sistemas Solares Térmicos, Ley Nº 20.365"
        Grid3.Cell(56, 3).Alignment = cellLeftGeneral
        Grid3.Cell(56, 3).Font.Size = 7
        Grid3.Cell(56, 7).text = "725"
        Grid3.Range(56, 8, 56, 9).Merge
        Grid3.Cell(56, 10).text = "Remanente mes anterior"
        Grid3.Range(56, 10, 56, 12).Merge
        Grid3.Cell(56, 10).Alignment = cellLeftGeneral
        Grid3.Cell(56, 10).Font.Size = 7
        Grid3.Cell(56, 13).text = "737"
        Grid3.Range(56, 14, 56, 18).Merge
        Grid3.Range(56, 19, 56, 21).Merge
        Grid3.Cell(56, 19).text = "Total Crédito"
        Grid3.Cell(56, 19).Alignment = cellLeftGeneral
        Grid3.Cell(56, 22).text = "727"
        Grid3.Range(56, 23, 56, 25).Merge
        Grid3.Cell(56, 26).text = "-"
        Grid3.Cell(56, 26).Alignment = cellCenterGeneral
        
        Grid3.Cell(57, 2).text = "101"
        Grid3.Range(57, 3, 57, 6).Merge
        Grid3.Cell(57, 3).text = " Imputación del Pago Patentes Aguas, Ley Nº 20.017"
        Grid3.Cell(57, 3).Alignment = cellLeftGeneral
        Grid3.Cell(57, 3).Font.Size = 7
        Grid3.Cell(57, 7).text = "704"
        Grid3.Range(57, 8, 57, 9).Merge
        Grid3.Cell(57, 10).text = "Remanente mes anterior"
        Grid3.Range(57, 10, 57, 12).Merge
        Grid3.Cell(57, 10).Alignment = cellLeftGeneral
        Grid3.Cell(57, 10).Font.Size = 7
        Grid3.Cell(57, 13).text = "705"
        Grid3.Range(57, 14, 57, 18).Merge
        Grid3.Range(57, 19, 57, 21).Merge
        Grid3.Cell(57, 19).text = "Total Crédito"
        Grid3.Cell(57, 19).Alignment = cellLeftGeneral
        Grid3.Cell(57, 22).text = "706"
        Grid3.Range(57, 23, 57, 25).Merge
        Grid3.Cell(57, 26).text = "-"
        Grid3.Cell(57, 26).Alignment = cellCenterGeneral
        
        Grid3.Cell(58, 2).text = "102"
        Grid3.Range(58, 3, 58, 6).Merge
        Grid3.Cell(58, 3).text = " Cotización Adicional, Ley Nº 18.566"
        Grid3.Cell(58, 3).Alignment = cellLeftGeneral
        Grid3.Cell(58, 3).Font.Size = 7
        Grid3.Cell(58, 7).text = "160"
        Grid3.Range(58, 8, 58, 9).Merge
        Grid3.Cell(58, 10).text = "Remanente mes anterior"
        Grid3.Range(58, 10, 58, 12).Merge
        Grid3.Cell(58, 10).Alignment = cellLeftGeneral
        Grid3.Cell(58, 10).Font.Size = 7
        Grid3.Cell(58, 13).text = "161"
        Grid3.Range(58, 14, 58, 18).Merge
        Grid3.Range(58, 19, 58, 21).Merge
        Grid3.Cell(58, 19).text = "Total Crédito"
        Grid3.Cell(58, 19).Alignment = cellLeftGeneral
        Grid3.Cell(58, 22).text = "570"
        Grid3.Range(58, 23, 58, 25).Merge
        Grid3.Cell(58, 26).text = "-"
        Grid3.Cell(58, 26).Alignment = cellCenterGeneral
        
        Grid3.Cell(59, 2).text = "103"
        Grid3.Range(59, 3, 59, 6).Merge
        Grid3.Cell(59, 3).text = " Crédito Especial Empresas Constructoras"
        Grid3.Cell(59, 3).Alignment = cellLeftGeneral
        Grid3.Cell(59, 3).Font.Size = 7
        Grid3.Cell(59, 7).text = "126"
        Grid3.Range(59, 8, 59, 9).Merge
        Grid3.Cell(59, 10).text = "Remanente mes anterior"
        Grid3.Range(59, 10, 59, 12).Merge
        Grid3.Cell(59, 10).Alignment = cellLeftGeneral
        Grid3.Cell(59, 10).Font.Size = 7
        Grid3.Cell(59, 13).text = "128"
        Grid3.Range(59, 14, 59, 18).Merge
        Grid3.Range(59, 19, 59, 21).Merge
        Grid3.Cell(59, 19).text = "Total Crédito"
        Grid3.Cell(59, 19).Alignment = cellLeftGeneral
        Grid3.Cell(59, 22).text = "571"
        Grid3.Range(59, 23, 59, 25).Merge
        Grid3.Cell(59, 26).text = "-"
        Grid3.Cell(59, 26).Alignment = cellCenterGeneral
        
        Grid3.Cell(60, 2).text = "104"
        Grid3.Range(60, 3, 60, 6).Merge
        Grid3.Cell(60, 3).text = " Recup. Peajes Transportistas Pasajero, Ley Nº 19.764"
        Grid3.Cell(60, 3).Alignment = cellLeftGeneral
        Grid3.Cell(60, 3).Font.Size = 7
        Grid3.Cell(60, 7).text = "572"
        Grid3.Range(60, 8, 60, 9).Merge
        Grid3.Cell(60, 10).text = "Remanente mes anterior"
        Grid3.Range(60, 10, 60, 12).Merge
        Grid3.Cell(60, 10).Alignment = cellLeftGeneral
        Grid3.Cell(60, 10).Font.Size = 7
        Grid3.Cell(60, 13).text = "568"
        Grid3.Range(60, 14, 60, 18).Merge
        Grid3.Range(60, 19, 60, 21).Merge
        Grid3.Cell(60, 19).text = "Total Crédito"
        Grid3.Cell(60, 19).Alignment = cellLeftGeneral
        Grid3.Cell(60, 22).text = "590"
        Grid3.Range(60, 23, 60, 25).Merge
        Grid3.Cell(60, 26).text = "-"
        Grid3.Cell(60, 26).Alignment = cellCenterGeneral
        
        
         Grid3.Range(61, 1, 61, 26).Merge
         Grid3.Range(62, 1, 62, 26).Merge
         Grid3.Cell(62, 1).text = " Realice la operación aritmética de las líneas 59 a 104 columna Impuesto Determinado. Registre el valor resultante en el código 547 (línea 105), Si es negativo anótelo entre paréntesis."
         Grid3.Cell(62, 1).Alignment = cellLeftGeneral
         
         Grid3.Range(63, 1, 63, 26).Merge
         
        Grid3.Cell(64, 2).text = "105"
        Grid3.Range(64, 3, 64, 21).Merge
        Grid3.Cell(64, 3).text = " TOTAL DETERMINADO"
        Grid3.Cell(64, 3).Alignment = cellLeftGeneral
        Grid3.Cell(64, 3).Font.Bold = True
        Grid3.Cell(64, 22).text = "547"
        Grid3.Range(64, 23, 64, 25).Merge
        Grid3.Cell(64, 26).text = "="
        Grid3.Cell(64, 26).Alignment = cellCenterGeneral
        Grid3.Range(65, 1, 65, 26).Merge
        
        
End Sub


Sub quintocuadro2()
       Grid3.Range(65, 1, 65, 26).Merge
        
        Grid3.Cell(66, 2).text = "106"
        Grid3.Range(66, 3, 66, 8).Merge
        Grid3.Cell(66, 3).text = " Remanente Crédito por Sistemas Solares Térmicos, Ley Nº 20.365"
        Grid3.Cell(66, 3).Alignment = cellLeftGeneral
        Grid3.Cell(66, 3).Font.Size = 7
        Grid3.Cell(66, 9).text = "728"
        Grid3.Range(66, 10, 66, 13).Merge
        Grid3.Range(66, 14, 66, 26).Merge
        
        Grid3.Cell(67, 2).text = "107"
        Grid3.Range(67, 3, 67, 8).Merge
        Grid3.Cell(67, 3).text = " Remanente Crédito por Sistemas Solares Térmicos, Ley Nº 20.365"
        Grid3.Cell(67, 3).Alignment = cellLeftGeneral
        Grid3.Cell(67, 3).Font.Size = 7
        Grid3.Cell(67, 9).text = "707"
        Grid3.Range(67, 10, 67, 13).Merge
        Grid3.Range(67, 14, 67, 21).Merge
        
        Grid3.Range(67, 22, 70, 25).Merge
        Grid3.Cell(67, 22).text = "Si código 547 es positivo, trasládelo al anverso (código 91, línea 111), en caso contrario regístrelo en los códigos de remanente (línea 106 a 110) teniendo presente las instrucciones."
        Grid3.Cell(67, 22).WrapText = True
        Grid3.Cell(67, 22).Font.Size = 7
        Grid3.Cell(67, 22).Alignment = cellCenterGeneral
        
        Grid3.Cell(68, 2).text = "108"
        Grid3.Range(68, 3, 68, 8).Merge
        Grid3.Cell(68, 3).text = " Remanente de Cotizacion Adicional Ley Nº 18.566"
        Grid3.Cell(68, 3).Alignment = cellLeftGeneral
        Grid3.Cell(68, 3).Font.Size = 7
        Grid3.Cell(68, 9).text = "73"
        Grid3.Range(68, 10, 68, 13).Merge
        Grid3.Range(68, 14, 68, 21).Merge
        
        Grid3.Cell(69, 2).text = "109"
        Grid3.Range(69, 3, 69, 8).Merge
        Grid3.Cell(69, 3).text = " Remanente Crédito Especial Empresas Constructoras"
        Grid3.Cell(69, 3).Alignment = cellLeftGeneral
        Grid3.Cell(69, 3).Font.Size = 7
        Grid3.Cell(69, 9).text = "130"
        Grid3.Range(69, 10, 69, 13).Merge
        Grid3.Range(69, 14, 69, 21).Merge
        
        Grid3.Cell(70, 2).text = "110"
        Grid3.Range(70, 3, 70, 8).Merge
        Grid3.Cell(70, 3).text = " Remanente Recup. de Peajes Trans. Pasajeros Ley Nº 19.764"
        Grid3.Cell(70, 3).Alignment = cellLeftGeneral
        Grid3.Cell(70, 3).Font.Size = 7
        Grid3.Cell(70, 9).text = "591"
        Grid3.Range(70, 10, 70, 13).Merge
        Grid3.Range(70, 14, 70, 21).Merge
        
        
End Sub

Sub sextocuadro2()
       Grid3.Range(71, 1, 71, 26).Merge
        
        Grid3.Cell(72, 1).text = "REGISTRE SI CAMBIA ALGUNO DE LOS SIGUIENTES ANTECEDENTES"
        Grid3.Range(72, 1, 72, 26).Merge
        Grid3.Cell(72, 1).Font.Bold = True
        Grid3.Cell(72, 1).Alignment = cellCenterGeneral
        
        Grid3.Cell(73, 1).text = "06"
        Grid3.Range(73, 2, 73, 10).Merge
        Grid3.Cell(73, 2).text = "Calle"
        Grid3.Cell(73, 2).Alignment = cellCenterGeneral
        Grid3.Cell(73, 11).text = "610"
        Grid3.Cell(73, 12).text = "N°"
        Grid3.Range(73, 12, 73, 15).Merge
        Grid3.Cell(73, 12).Alignment = cellCenterGeneral
        Grid3.Cell(73, 16).text = "611"
        Grid3.Cell(73, 17).text = "Departamento"
        Grid3.Range(73, 17, 73, 23).Merge
        Grid3.Cell(73, 17).Alignment = cellCenterGeneral
        Grid3.Cell(73, 24).text = "612"
        Grid3.Cell(73, 25).text = "Villa o Población"
        Grid3.Range(73, 25, 73, 26).Merge
        Grid3.Cell(73, 25).Alignment = cellCenterGeneral
        
      
        Grid3.Range(74, 1, 74, 10).Merge
        Grid3.Range(74, 11, 74, 15).Merge
        Grid3.Range(74, 16, 74, 23).Merge
        Grid3.Range(74, 24, 74, 26).Merge
        
        
        Grid3.Cell(75, 1).text = "08"
        Grid3.Range(75, 2, 75, 3).Merge
        Grid3.Cell(75, 2).text = "Comuna"
        Grid3.Cell(75, 2).Alignment = cellCenterGeneral
        Grid3.Cell(75, 4).text = "53"
        Grid3.Cell(75, 5).text = "Región"
        Grid3.Range(75, 5, 75, 6).Merge
        Grid3.Cell(75, 5).Alignment = cellCenterGeneral
        Grid3.Cell(75, 7).text = "613"
        Grid3.Cell(75, 8).text = "Cód. área teléfono"
        Grid3.Range(75, 8, 75, 10).Merge
        Grid3.Cell(75, 8).Alignment = cellCenterGeneral
        Grid3.Cell(75, 11).text = "09"
        Grid3.Cell(75, 12).text = "Teléfono"
        Grid3.Range(75, 12, 75, 15).Merge
        Grid3.Cell(75, 12).Alignment = cellCenterGeneral
        Grid3.Cell(75, 16).text = "601"
        Grid3.Cell(75, 17).text = "Fax"
        Grid3.Range(75, 17, 75, 23).Merge
        Grid3.Cell(75, 17).Alignment = cellCenterGeneral
        Grid3.Cell(75, 24).text = "604"
        Grid3.Cell(75, 25).text = "Teléfono celular"
        Grid3.Range(75, 25, 75, 26).Merge
        Grid3.Cell(75, 25).Alignment = cellCenterGeneral
        
        Grid3.Range(76, 1, 76, 3).Merge
        Grid3.Range(76, 4, 76, 6).Merge
        Grid3.Range(76, 7, 76, 10).Merge
        Grid3.Range(76, 11, 76, 15).Merge
        Grid3.Range(76, 16, 76, 23).Merge
        Grid3.Range(76, 24, 76, 26).Merge
        
        
        
        Grid3.Cell(77, 1).text = "55"
        Grid3.Range(77, 2, 77, 5).Merge
        Grid3.Cell(77, 2).text = "Correo Electrónico"
        Grid3.Cell(77, 2).Alignment = cellCenterGeneral
        Grid3.Cell(77, 6).text = "44"
        Grid3.Cell(77, 7).text = "Domicilio Postal"
        Grid3.Range(77, 7, 77, 10).Merge
        Grid3.Cell(77, 7).Alignment = cellCenterGeneral
 
        Grid3.Cell(77, 11).text = "726"
        Grid3.Cell(77, 12).text = "Comuna Postal"
        Grid3.Range(77, 12, 77, 15).Merge
        Grid3.Cell(77, 12).Alignment = cellCenterGeneral
        Grid3.Cell(77, 16).text = "313"
        Grid3.Cell(77, 17).text = "Rut Contador"
        Grid3.Range(77, 17, 77, 23).Merge
        Grid3.Cell(77, 17).Alignment = cellCenterGeneral
        Grid3.Cell(77, 24).text = "314"
        Grid3.Cell(77, 25).text = "Rut Representante Legal"
        Grid3.Cell(77, 25).Font.Size = 7
        Grid3.Range(77, 25, 77, 26).Merge
        Grid3.Cell(77, 25).Alignment = cellCenterGeneral
        
        Grid3.Range(78, 1, 78, 5).Merge
        Grid3.Range(78, 6, 78, 10).Merge
        Grid3.Range(78, 11, 78, 15).Merge
        Grid3.Range(78, 16, 78, 23).Merge
        Grid3.Range(78, 24, 78, 26).Merge
        
        
End Sub

