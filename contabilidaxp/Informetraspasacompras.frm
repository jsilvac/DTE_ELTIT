VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0010 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Traspaso de Facturas de Compras"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   5400
      TabIndex        =   16
      Top             =   1440
      Width           =   3135
      _ExtentX        =   5530
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   6
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
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   7
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   90
      TabIndex        =   8
      Top             =   45
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   15187
      BackColor       =   16761024
      Caption         =   "INFORME TRASPASO DE FACTURAS DE COMPRAS"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
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
      Begin VB.CommandButton BUSCAR 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Busca Orden"
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
         Left            =   11340
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8190
         Width           =   1320
      End
      Begin VB.TextBox ORDEN 
         BackColor       =   &H00FFC0C0&
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
         Left            =   12825
         MaxLength       =   10
         TabIndex        =   14
         Top             =   8190
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "EXPORTAR A EXCEL"
         Height          =   330
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   8190
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8190
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   11
         Top             =   360
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   1852
         BackColor       =   16744576
         Caption         =   ""
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
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   6000
            TabIndex        =   5
            Top             =   360
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   915
            Left            =   240
            TabIndex        =   19
            Top             =   0
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   1614
            BackColor       =   16761024
            Caption         =   "Fecha Consultar"
            CaptionEstilo3D =   1
            BackColor       =   16761024
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
            Begin VB.TextBox DESDE3 
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1110
               MaxLength       =   4
               TabIndex        =   2
               Tag             =   "fecha"
               Top             =   405
               Width           =   615
            End
            Begin VB.TextBox DESDE2 
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   750
               MaxLength       =   2
               TabIndex        =   1
               Tag             =   "fecha"
               Top             =   405
               Width           =   375
            End
            Begin VB.TextBox DESDE1 
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   390
               MaxLength       =   2
               TabIndex        =   0
               Tag             =   "fecha"
               Top             =   405
               Width           =   375
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   915
            Left            =   3000
            TabIndex        =   20
            Top             =   0
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   1614
            BackColor       =   16761024
            Caption         =   "Hora Consulta"
            CaptionEstilo3D =   1
            BackColor       =   16761024
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
            Begin VB.TextBox txtHoraHasta 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1200
               MaxLength       =   5
               TabIndex        =   4
               Tag             =   "s"
               Text            =   "00:00"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtHoraDesde 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   3
               Tag             =   "s"
               Text            =   "00:00"
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "HASTA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   1200
               TabIndex        =   22
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "DESDE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   -120
         TabIndex        =   9
         Top             =   2640
         Visible         =   0   'False
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   11774
         BackColor       =   16761024
         Caption         =   "LISTADO DE FACTURAS DE FACTURAS RECIBIDAS"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Grid1 
            Height          =   6360
            Left            =   0
            TabIndex        =   10
            Top             =   270
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   11218
            BackColorFixed  =   14737632
            BackColorSel    =   12648447
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   0
            Rows            =   30
            DateFormat      =   2
         End
      End
   End
End
Attribute VB_Name = "prove0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String
Private FORMATOGRILLA(20, 30)
Private lin As Double
Private tipoprove As String
Private plan(2000, 3) As Variant
Private canplan As Double
Private total(20) As Double
Private detalle(20, 20) As Double
Private TIPOS(9) As String
Private MES As String
Private año As String
Private totaldocumentos As Double
Private refrescos As String
Private licores As String
Private vinos As String
Private cerveza As String
Private HARINA As String
Private CARNE As String


Private Sub BUSCAR_Click()
 Dim i As Integer
 
  For i = 1 To Grid1.Rows - 1
            If Mid(Grid1.Cell(i, 17).text, 1, 10) = ORDEN.text Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                Exit For
            End If
        Next i
End Sub

Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
    Dim infogrilla As grillainformes
    Set infogrilla = New grillainformes
    
    
For k = 1 To 2000
plan(k, 3) = 0
Next k
For k = 1 To 20
detalle(k, 1) = 0
detalle(k, 2) = 0
detalle(k, 3) = 0
detalle(k, 4) = 0
detalle(k, 5) = 0
detalle(k, 6) = 0
detalle(k, 7) = 0
detalle(k, 8) = 0
detalle(k, 9) = 0
detalle(k, 10) = 0
detalle(k, 11) = 0
detalle(k, 12) = 0
detalle(k, 13) = 0
detalle(k, 14) = 0
detalle(k, 15) = 0
detalle(k, 16) = 0
detalle(k, 17) = 0
detalle(k, 18) = 0
detalle(k, 19) = 0
detalle(k, 20) = 0

Next k


    Call CARGAGRILLA(infogrilla)
    Call Consulta_Informe(infogrilla)
    infogrilla.Visible = True
    infogrilla.Caption = "DOCUMENTOS INGRESADOS "
    
    
    grillainformes.Tag = "prove0010"
    
    infogrilla.Show
End Sub



Private Sub Command3_Click()
Dim k As Integer


End Sub


Private Sub Command4_Click()
    If Grid1.Rows > 1 Then
        Call Grid1.ExportToExcel("", True, True)
    End If
End Sub

Private Sub DESDE1_GotFocus()
    Call cargatexto(DESDE1)
End Sub

Private Sub DESDE1_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(DESDE1, DESDE2, KeyCode)
End Sub

Private Sub DESDE1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(DESDE1)
        If DESDE1.text = "00" Then DESDE1.text = Format(fechasistema, "dd")
        DESDE2.SetFocus
    End If
End Sub

Private Sub DESDE2_GotFocus()
    Call cargatexto(DESDE2)
End Sub

Private Sub DESDE2_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(DESDE1, DESDE3, KeyCode)
End Sub

Private Sub DESDE2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(DESDE2)
        If DESDE2.text = "00" Then DESDE2.text = Format(fechasistema, "mm")
        DESDE3.SetFocus
    End If
End Sub

Private Sub DESDE3_GotFocus()
    Call cargatexto(DESDE3)
End Sub

Private Sub DESDE3_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(DESDE2, txtHoraDesde, KeyCode)
End Sub

Private Sub DESDE3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(DESDE3)
        If DESDE3.text = "0000" Then DESDE3.text = Format(fechasistema, "yyyy")
        txtHoraDesde.SetFocus
    End If
End Sub

Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD
    sc = 0
'CARGAGRILLA
Call Conectarventas(Servidor, clientesistema + "ventas00", Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)

Call Conectargestionrubro(Servidor, clientesistema + "gestion" + rubro, Usuario, password)

 txtHoraDesde.text = Time
 txtHoraHasta.text = Time


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
 
 

Sub limpia()
    
    
End Sub

Sub imprimir()
Dim titulo As String
titulo = "LISTADO DE FACTURAS EMITIDAS  "
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
Sub grilla()
    
End Sub

Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
'Sub CARGAGRILLA()
'Rem DATOS DE LA COLUMNA
'    Dim FormatoGrilla(10, 20)
'    Grid1.DefaultFont.Size = 8
'    Grid1.DefaultFont.Bold = False
'
'
'    FormatoGrilla(1, 1) = "TP"
'    FormatoGrilla(1, 2) = "NUMERO"
'    FormatoGrilla(1, 3) = "RUT"
'    FormatoGrilla(1, 4) = "PROVEEDOR"
'    FormatoGrilla(1, 5) = "FECHA"
'    FormatoGrilla(1, 6) = "NETO"
'    FormatoGrilla(1, 7) = "IVA"
'    FormatoGrilla(1, 8) = "I.CERV"
'    FormatoGrilla(1, 9) = "I.REFRE"
'    FormatoGrilla(1, 10) = "I.VINO "
'    FormatoGrilla(1, 11) = "I.LICOR"
'    FormatoGrilla(1, 12) = "I.HARINA"
'    FormatoGrilla(1, 13) = "I.CARNE"
'    FormatoGrilla(1, 14) = "TOTAL  "
'    FormatoGrilla(1, 15) = "CO"
'    FormatoGrilla(1, 16) = "TP"
'    FormatoGrilla(1, 17) = "ORDEN"
'    FormatoGrilla(1, 18) = "MES"
'    FormatoGrilla(1, 19) = "AÑO"
'    FormatoGrilla(1, 20) = "RECEPCION"
'
'    Rem LARGO DE LOS DATOS
'    FormatoGrilla(2, 1) = "2"
'    FormatoGrilla(2, 2) = "10"
'    FormatoGrilla(2, 3) = "11"
'    FormatoGrilla(2, 4) = "20"
'    FormatoGrilla(2, 5) = "10"
'    FormatoGrilla(2, 6) = "8"
'    FormatoGrilla(2, 7) = "7"
'    FormatoGrilla(2, 8) = "7"
'    FormatoGrilla(2, 9) = "7"
'    FormatoGrilla(2, 10) = "7"
'    FormatoGrilla(2, 11) = "7"
'    FormatoGrilla(2, 12) = "7"
'    FormatoGrilla(2, 13) = "8"
'    FormatoGrilla(2, 14) = "8"
'    FormatoGrilla(2, 15) = "3"
'    FormatoGrilla(2, 16) = "2"
'    FormatoGrilla(2, 17) = "10"
'    FormatoGrilla(2, 18) = "4"
'    FormatoGrilla(2, 19) = "4"
'    FormatoGrilla(2, 20) = "10"
'
'
'    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
'    FormatoGrilla(3, 1) = "S"
'    FormatoGrilla(3, 2) = "N"
'    FormatoGrilla(3, 3) = "N"
'    FormatoGrilla(3, 4) = "S"
'    FormatoGrilla(3, 5) = "D"
'    FormatoGrilla(3, 6) = "N"
'    FormatoGrilla(3, 7) = "N"
'    FormatoGrilla(3, 8) = "N"
'    FormatoGrilla(3, 9) = "N"
'    FormatoGrilla(3, 10) = "N"
'    FormatoGrilla(3, 11) = "N"
'    FormatoGrilla(3, 12) = "N"
'    FormatoGrilla(3, 13) = "N"
'    FormatoGrilla(3, 14) = "N"
'    FormatoGrilla(3, 15) = "D"
'
'    Rem FORMATO GRILLA
'    FormatoGrilla(4, 2) = "0000000000"
'    FormatoGrilla(4, 3) = ""
'
'    FormatoGrilla(4, 4) = ""
'    FormatoGrilla(4, 5) = ""
'
'
'    FormatoGrilla(4, 6) = "##,###,##0"
'    FormatoGrilla(4, 7) = "##,###,##0"
'    FormatoGrilla(4, 8) = "##,###,##0"
'    FormatoGrilla(4, 9) = "##,###,##0"
'    FormatoGrilla(4, 10) = "##,###,##0"
'    FormatoGrilla(4, 11) = "##,###,##0"
'    FormatoGrilla(4, 12) = "##,###,##0"
'    FormatoGrilla(4, 13) = "##,###,##0"
'    FormatoGrilla(4, 14) = "##,###,##0"
'    FormatoGrilla(4, 17) = "0000000000"
'
'    Rem LOCCKED
'    For k = 1 To 20
'    FormatoGrilla(5, k) = "TRUE"
'
'    Next k
'
'    FormatoGrilla(5, 15) = "FALSE"
'
'
'    Grid1.Cols = 21
'    Grid1.Rows = 2
'
'    Grid1.AllowUserResizing = False
'    Grid1.DisplayFocusRect = False
'    Grid1.ExtendLastCol = True
'    Grid1.BoldFixedCell = False
'    Grid1.DrawMode = cellOwnerDraw
'
'    Grid1.Appearance = Flat
'    Grid1.ScrollBarStyle = Flat
'    Grid1.FixedRowColStyle = Flat
'
''   Grid1.BackColorFixed = RGB(90, 158, 214)
''   Grid1.BackColorFixedSel = RGB(110, 180, 230)
''   Grid1.BackColorBkg = RGB(90, 158, 214)
''   Grid1.BackColorScrollBar = RGB(231, 235, 247)
''   Grid1.BackColor1 = RGB(231, 235, 247)
''   Grid1.BackColor2 = RGB(239, 243, 255)
''   Grid1.GridColor = RGB(148, 190, 231)
'   Grid1.Column(0).Width = 0
'
'    For k = 1 To Grid1.Cols - 1
'
'        Grid1.Cell(0, k).text = FormatoGrilla(1, k)
'        Grid1.Column(k).Width = Val(FormatoGrilla(2, k)) * (Grid1.DefaultFont.Size - 1)
'        Grid1.Column(k).MaxLength = Val(FormatoGrilla(2, k))
'        Grid1.Column(k).FormatString = FormatoGrilla(4, k)
'        Grid1.Column(k).Locked = FormatoGrilla(5, k)
'        If FormatoGrilla(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
'        If FormatoGrilla(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
'
'    Next k
'    Grid1.Column(16).Width = 30
'    Grid1.Column(1).Width = 30
'    Grid1.Column(15).CellType = cellCheckBox
'    Grid1.Column(2).Mask = cellNumeric
'    Grid1.Column(6).Mask = cellNumeric
'    Grid1.Column(7).Mask = cellNumeric
'    Grid1.Column(8).Mask = cellNumeric
'    Grid1.Column(9).Mask = cellNumeric
'    Grid1.Column(10).Mask = cellNumeric
'    Grid1.Column(11).Mask = cellNumeric
'    Grid1.Column(12).Mask = cellNumeric
'
'    Grid1.Column(1).CellType = cellComboBox
'
'    Grid1.Column(16).CellType = cellComboBox
'
'
'
'    With Grid1.ComboBox(1)
'
'        '.Locked = False
'        .AutoComplete = True
'
'        .AddItem "FA FACTURA" '1
'        .AddItem "ND NOTA DEBITO" '2
'        .AddItem "NC NOTA CREDITO" '3
'        .AddItem "FAE FACTURA ELECTRONICA" '1
'        .AddItem "NDE NOTA DEBITO ELECTRONICA" '2
'        .AddItem "NCE NOTA CREDITO ELECTRONICA" '3
'        .AddItem "OE ORDEN DE ENLACE" '4
'        .AddItem "GD DESPACHO" '4
'
'
'    End With
'
'    With Grid1.ComboBox(16)
'        '.Locked = True
'        .AutoComplete = True
'        .AddItem "MERCADERIAS"
'        .AddItem "CIGARRILLOS"
'        .AddItem "FRUTAS Y VERDURAS"
'        .AddItem "CARNICERIA"
'        .AddItem "FIAMBRERIA"
'        .AddItem "PANADERIA"
'        .AddItem "EMPAQUE"
'        .AddItem "DIARIOS"
'
'    End With
'
'
'
'End Sub



Private Sub monto_Click()
End Sub

'Private Sub leer()
'
'Dim resultados As rdoResultset
'    Dim csql As New rdoQuery
'    Dim rut As String
'    Dim PASO As String
'    Dim fecha1 As String
'    Dim fecha2 As String
'    Dim linea As Double
'    Dim total As Double
'    Dim fec As Double
'    Dim fec1 As Double
'    Dim fechasum As String
'    Dim total2 As Double
'    Dim MESCONTABLE As Double
'
'    Dim AÑOCONTABLE As Double
'
'
'    linea = 0: fec = 0: fec1 = 0
'    fecha1 = año + "-" + mes + "-" + "01"
'    fecha2 = año + "-" + mes + "-" + "31"
'
'        Set csql.ActiveConnection = gestionrubro
'        csql.sql = "SELECT lof.tipo,lof.numero,lof.rut,lof.fecha,lof.neto,lof.iva,lof.total,lof.categoria,lof.bonificacion,loc.fecharecepcion,lof.ordendecompra "
'        csql.sql = csql.sql + "FROM l_ordendecompra_detalle_facturas_" + localfiltro + " as lof,l_ordendecompra_cabeza_" + localfiltro + " as loc "
'        csql.sql = csql.sql + "where lof.ordendecompra=loc.numero and "
''        If Option1.Value = False Then
'        csql.sql = csql.sql + "loc.fecharecepcion>='" + fecha1 + "' AND loc.fecharecepcion<='" + fecha2 + "' "
''        Else
''        csql.sql = csql.sql + "loc.fecharecepcion='" & Format(fechasistema, "yyyy-mm-dd") + "' "
''        End If
'
'        csql.sql = csql.sql + "and (lof.tipo='FA' or lof.tipo='NC' or lof.tipo='ND' or lof.tipo='FAE' or lof.tipo='NCE' or lof.tipo='NDE') order by loc.fecharecepcion,lof.ordendecompra "
'        csql.sql = csql.sql + ""
'        csql.Execute
'        total = 0
'        total2 = 0
'        Grid1.Rows = 1
'        Grid1.AutoRedraw = False
'
'
'        If csql.RowsAffected > 0 Then
'        Set resultados = csql.OpenResultset
'        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
'
'         While Not resultados.EOF
'
'             If leefactura(resultados(0), resultados(1), resultados(2)) = "0" Then
'             Grid1.Rows = Grid1.Rows + 1
'
'             linea = linea + 1
'             Grid1.Cell(linea, 1).text = resultados(0)
'             Grid1.Cell(linea, 2).text = resultados(1)
'             Grid1.Cell(linea, 3).text = Mid(resultados(2), 1, 9) + "-" + Mid(resultados(2), 10, 1)
'             Grid1.Cell(linea, 4).text = nombrectacte(resultados(2))
'             If IsNull(resultados(3)) = False Then
'             Grid1.Cell(linea, 5).text = resultados(3)
'             End If
'             Grid1.Cell(linea, 6).text = resultados(4)
'             Grid1.Cell(linea, 7).text = resultados(5)
'             Grid1.Cell(linea, 8).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400014")
'             Grid1.Cell(linea, 9).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400010")
'             Grid1.Cell(linea, 10).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400011")
'             Grid1.Cell(linea, 11).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400013") + LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(2), "11400014") + LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(2), "11400015")
'             Grid1.Cell(linea, 12).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400005")
'             Grid1.Cell(linea, 13).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400012")
'             Grid1.Cell(linea, 14).text = resultados(6)
'             Grid1.Cell(linea, 15).text = "0"
'             Grid1.Cell(linea, 16).text = resultados(7)
'             Grid1.Cell(linea, 17).text = resultados(10)
'             MESCONTABLE = CDbl(Format(fechasistema, "mm"))
'             AÑOCONTABLE = CDbl(Format(fechasistema, "yyyy"))
'             If Format(resultados(3), "yyyy-mm") < Format(fechasistema, "yyyy-mm") And Format(fechasistema, "dd") <= diacierrecompra Then
'             MESCONTABLE = MESCONTABLE - 1
'             If MESCONTABLE = 0 Then MESCONTABLE = 12: AÑOCONTABLE = AÑOCONTABLE - 1
'
'             End If
'
'             Grid1.Cell(linea, 18).text = Format(MESCONTABLE, "00")
'             Grid1.Cell(linea, 19).text = AÑOCONTABLE
'             Grid1.Cell(linea, 20).text = resultados(9)
'
'            End If
'            resultados.MoveNext
'
'            Wend
'End If
'      Grid1.AutoRedraw = True
'      Grid1.Refresh
'
'
'
'End Sub
Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "FOLIO"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "FECHA"
    FORMATOGRILLA(1, 5) = "RUT"
    FORMATOGRILLA(1, 6) = "PROVEEDOR"
    FORMATOGRILLA(1, 7) = "NETO"
    FORMATOGRILLA(1, 8) = "IVA"
    FORMATOGRILLA(1, 9) = "EXENTO"
    FORMATOGRILLA(1, 10) = "IMPTO DIESEL"
    FORMATOGRILLA(1, 11) = "RETENCION"
    
    FORMATOGRILLA(1, 12) = "TOTAL"
    FORMATOGRILLA(1, 13) = " CUENTA "
    FORMATOGRILLA(1, 14) = " MONTO "
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "4"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "30"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    FORMATOGRILLA(2, 11) = "9"
    FORMATOGRILLA(2, 12) = "9"
    FORMATOGRILLA(2, 13) = "30"
    FORMATOGRILLA(2, 14) = "9"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "S"
    FORMATOGRILLA(3, 14) = "N"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    FORMATOGRILLA(4, 10) = "###,###,###"
    FORMATOGRILLA(4, 11) = "###,###,###"
    FORMATOGRILLA(4, 12) = "###,###,###"
    FORMATOGRILLA(4, 14) = "###,###,###"
    
    Rem LOCCKED
    For k = 1 To 14
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 15
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    infogrilla.Grid1.Column(0).Width = 0
    
    For k = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        infogrilla.Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        infogrilla.Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub
Sub Consulta_Informe(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim PASO As String
        totaldocumentos = 0
        tipoprove = CUENTAPROVEEDOR
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT folio,fc.tipo,numero,fecha,fc.rut,cc.nombre,neto,iva,exento,impuestoespecifico,retencion,total,fc.electronica,fc.activo,fc.comentario "
        csql.sql = csql.sql + " FROM facturasdecompras as fc,cuentascorrientes as cc "
        csql.sql = csql.sql & " WHERE MID(horatraspaso,1,5) BETWEEN '" & txtHoraDesde.text & "' AND '" & txtHoraHasta.text & "' and usuario='" & USUARIOSISTEMA & "' and "
        csql.sql = csql.sql + "fc.rut=cc.rut and cc.año='" + DESDE3.text + "' and cc.tipo='" + tipoprove + "' and fechatraspaso='" + DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text & "'"
         
        csql.sql = csql.sql + " order by fecha "
        
        
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0
          total(7) = 0
        If csql.RowsAffected > 0 Then
'        barra.Max = csql.RowsAffected
'        barra.Value = 0
        Set resultados = csql.OpenResultset
        lin = 0
         While Not resultados.EOF

'         If RESUMEN1.Value = True Then
'             barra.Value = lin
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
             For k = 0 To 11
             infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
             
             Next k
             multi = 1
                totaldocumentos = totaldocumentos + 1
                If resultados(1) = "1" Then infogrilla.Grid1.Cell(lin, 2).text = "FA"
                If resultados(1) = "2" Then infogrilla.Grid1.Cell(lin, 2).text = "ND"
                If resultados(1) = "3" Then infogrilla.Grid1.Cell(lin, 2).text = "NC": multi = -1
                If resultados(1) = "4" Then infogrilla.Grid1.Cell(lin, 2).text = "FAE"
                If resultados(1) = "5" Then infogrilla.Grid1.Cell(lin, 2).text = "NDE"
                If resultados(1) = "6" Then infogrilla.Grid1.Cell(lin, 2).text = "NCE": multi = -1
                If resultados(1) = "7" Then infogrilla.Grid1.Cell(lin, 2).text = "FC"
                If resultados(1) = "8" Then infogrilla.Grid1.Cell(lin, 2).text = "IM"
                
                infogrilla.Grid1.Cell(lin, 7).text = resultados(6) * multi
                infogrilla.Grid1.Cell(lin, 8).text = resultados(7) * multi
                infogrilla.Grid1.Cell(lin, 9).text = resultados(8) * multi
                infogrilla.Grid1.Cell(lin, 10).text = resultados(9) * multi
                infogrilla.Grid1.Cell(lin, 11).text = resultados(10) * multi
                infogrilla.Grid1.Cell(lin, 12).text = resultados(11) * multi
                
                infogrilla.Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 9) + "-" + Mid(resultados(4), 10, 1)
                Rem If resultados(12) = "S" Then infogrilla.Grid1.Cell(lin, 2).text = infogrilla.Grid1.Cell(lin, 2).text + "E"
         
         
'         End If
             If resultados(1) = "3" Or resultados(1) = "6" Then multi = -1 Else multi = 1
             total(1) = total(1) + resultados(6) * multi
             total(2) = total(2) + resultados(7) * multi
             total(3) = total(3) + resultados(8) * multi
             total(4) = total(4) + resultados(9) * multi
             total(5) = total(5) + resultados(10) * multi
             total(6) = total(6) + resultados(11) * multi
                          
                          Rem If resultados(1) = "7" And resultados(13) <> "S" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9):: detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 7) = detalle(1, 7) + resultados(11)
                          If resultados(1) = "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9):: detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 7) = detalle(1, 7) + resultados(11)
                          If resultados(1) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + resultados(8):: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 7) = detalle(2, 7) + resultados(11)
                          If resultados(1) = "3" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + resultados(8):: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 7) = detalle(3, 7) + resultados(11)
                          If resultados(1) = "4" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(6): detalle(4, 3) = detalle(4, 3) + resultados(7):: detalle(4, 4) = detalle(4, 4) + resultados(8):: detalle(4, 5) = detalle(4, 5) + resultados(9): detalle(4, 6) = detalle(4, 6) + resultados(10): detalle(4, 7) = detalle(4, 7) + resultados(11)
                          If resultados(1) = "5" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(6): detalle(5, 3) = detalle(5, 3) + resultados(7):: detalle(5, 4) = detalle(5, 4) + resultados(8):: detalle(5, 5) = detalle(5, 5) + resultados(9): detalle(5, 6) = detalle(5, 6) + resultados(10): detalle(5, 7) = detalle(5, 7) + resultados(11)
                          If resultados(1) = "6" Then detalle(6, 1) = detalle(6, 1) + 1: detalle(6, 2) = detalle(6, 2) + resultados(6): detalle(6, 3) = detalle(6, 3) + resultados(7):: detalle(6, 4) = detalle(6, 4) + resultados(8):: detalle(6, 5) = detalle(6, 5) + resultados(9): detalle(6, 6) = detalle(6, 6) + resultados(10): detalle(6, 7) = detalle(6, 7) + resultados(11)
                          If resultados(13) = "S" And resultados(1) <> "3" And resultados(1) <> "6" Then detalle(7, 1) = detalle(7, 1) + 1: detalle(7, 2) = detalle(7, 2) + resultados(6): detalle(7, 3) = detalle(7, 3) + resultados(7):: detalle(7, 4) = detalle(7, 4) + resultados(8):: detalle(7, 5) = detalle(7, 5) + resultados(9): detalle(7, 6) = detalle(7, 6) + resultados(10): detalle(7, 7) = detalle(7, 7) + resultados(11)
                          If resultados(1) = "7" Then detalle(8, 1) = detalle(8, 1) + 1: detalle(8, 2) = detalle(8, 2) + resultados(6): detalle(8, 3) = detalle(8, 3) + resultados(7):: detalle(8, 4) = detalle(8, 4) + resultados(8):: detalle(8, 5) = detalle(8, 5) + resultados(9): detalle(8, 6) = detalle(8, 6) + resultados(10): detalle(8, 7) = detalle(8, 7) + resultados(11)
                          If resultados(1) = "8" Then detalle(9, 1) = detalle(9, 1) + 1: detalle(9, 2) = detalle(9, 2) + resultados(6): detalle(9, 3) = detalle(9, 3) + resultados(7):: detalle(9, 4) = detalle(9, 4) + resultados(8):: detalle(9, 5) = detalle(9, 5) + resultados(9): detalle(9, 6) = detalle(9, 6) + resultados(10): detalle(9, 7) = detalle(9, 7) + resultados(11)
                          
             
             
'                          If resultados(12) <> "S" And resultados(13) <> "S" And resultados(1) = "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(6): detalle(1, 3) = detalle(1, 3) + resultados(7):: detalle(1, 4) = detalle(1, 4) + resultados(8):: detalle(1, 5) = detalle(1, 5) + resultados(9):: detalle(1, 6) = detalle(1, 6) + resultados(10): detalle(1, 7) = detalle(1, 7) + resultados(11)
'                          If resultados(12) <> "S" And resultados(1) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(6): detalle(2, 3) = detalle(2, 3) + resultados(7):: detalle(2, 4) = detalle(2, 4) + resultados(8):: detalle(2, 5) = detalle(2, 5) + resultados(9): detalle(2, 6) = detalle(2, 6) + resultados(10): detalle(2, 7) = detalle(2, 7) + resultados(11)
'                          If resultados(12) <> "S" And resultados(1) = "3" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(6): detalle(3, 3) = detalle(3, 3) + resultados(7):: detalle(3, 4) = detalle(3, 4) + resultados(8):: detalle(3, 5) = detalle(3, 5) + resultados(9): detalle(3, 6) = detalle(3, 6) + resultados(10): detalle(3, 7) = detalle(3, 7) + resultados(11)
'                          If resultados(12) = "S" And resultados(1) = "1" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(6): detalle(4, 3) = detalle(4, 3) + resultados(7):: detalle(4, 4) = detalle(4, 4) + resultados(8):: detalle(4, 5) = detalle(4, 5) + resultados(9): detalle(4, 6) = detalle(4, 6) + resultados(10): detalle(4, 7) = detalle(4, 7) + resultados(11)
'                          If resultados(12) = "S" And resultados(1) = "2" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(6): detalle(5, 3) = detalle(5, 3) + resultados(7):: detalle(5, 4) = detalle(5, 4) + resultados(8):: detalle(5, 5) = detalle(5, 5) + resultados(9): detalle(5, 6) = detalle(5, 6) + resultados(10): detalle(5, 7) = detalle(5, 7) + resultados(11)
'                          If resultados(12) = "S" And resultados(1) = "3" Then detalle(6, 1) = detalle(6, 1) + 1: detalle(6, 2) = detalle(6, 2) + resultados(6): detalle(6, 3) = detalle(6, 3) + resultados(7):: detalle(6, 4) = detalle(6, 4) + resultados(8):: detalle(6, 5) = detalle(6, 5) + resultados(9): detalle(6, 6) = detalle(6, 6) + resultados(10): detalle(6, 7) = detalle(6, 7) + resultados(11)
'                          If resultados(13) = "S" And resultados(1) = "1" Then detalle(7, 1) = detalle(7, 1) + 1: detalle(7, 2) = detalle(7, 2) + resultados(6): detalle(7, 3) = detalle(7, 3) + resultados(7):: detalle(7, 4) = detalle(7, 4) + resultados(8):: detalle(7, 5) = detalle(7, 5) + resultados(9): detalle(7, 6) = detalle(7, 6) + resultados(10): detalle(7, 7) = detalle(7, 7) + resultados(11)
'
        If resultados("comentario") = "RECEPCION DTE" And (resultados(1) = "3" Or resultados(1) = "4" Or resultados(1) = "5") Then
        infogrilla.Grid1.Cell(lin, 3).BackColor = &H80FF80
        
        End If
        
        
'            If Check3.Value = 1 Then
              Call Consultadetalle(resultados(1), resultados(2), resultados(4), infogrilla)
'            End If
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
     
Call totallibro(infogrilla)
'barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub
Sub Consultadetalle(tipo, numero, rut, infogrilla As grillainformes)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT cuentadelmayor,monto "
        csql2.sql = csql2.sql + "FROM facturasdecompras_detalle "
        csql2.sql = csql2.sql + "where tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' order by linea "
        csql2.Execute
        
        If csql2.RowsAffected > 0 Then
'        barra.Max = barra.Max + csql2.RowsAffected - 1
        
        Set resultados2 = csql2.OpenResultset
        linpaso = 0
        While Not resultados2.EOF
          
          For k = 1 To canplan
          If tipo = 3 Or tipo = 6 Then multi = -1 Else multi = 1
          If resultados2(0) = plan(k, 1) Then plan(k, 3) = plan(k, 3) + (resultados2(1) * multi)
          If resultados2(0) = plan(k, 1) Then
            If linpaso = 1 And csql2.RowsAffected > 1 Then
            lin = lin + 1: infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
            End If
          
            infogrilla.Grid1.Cell(lin, 13).text = plan(k, 2): infogrilla.Grid1.Cell(lin, 14).text = resultados2(1): k = canplan + 1: linpaso = 1
          
          End If
          
            
          Next k
          resultados2.MoveNext
                

         Wend

          resultados2.Close

        End If

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
    objReportTitle.text = ""
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



Sub eliminafactura(tipo, numero)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_pagos_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM l_movimientos_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, gestionrubro, "")
        
        
End Sub


Private Sub Grid1_Click()
If Grid1.ActiveCell.col = 15 Then
If Mid(Grid1.Cell(Grid1.ActiveCell.row, 4).text, 1, 3) = "***" Then
Grid1.Cell(Grid1.ActiveCell.row, 15).text = "0"
End If
End If


End Sub

Private Sub Grid1_DblClick()
If Grid1.ActiveCell.col = 15 Then
If Mid(Grid1.Cell(Grid1.ActiveCell.row, 4).text, 1, 3) = "***" Then
Grid1.Cell(Grid1.ActiveCell.row, 15).text = "0"
End If
End If

localorden = localfiltro
Rcompra02.dato1.text = Grid1.Cell(Grid1.ActiveCell.row, 17).text

Rcompra02.Show vbModal


End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'End If
'leer
End Sub

Sub grabafactura(LINEA, tipo, ORDEN)
    Dim netos As Double
    Dim DH As String
    Dim DH2 As String
    Dim mesconta As String
    Dim añoconta As String
    Dim diaconta As String
    Dim CUENTA2 As String
    
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    Dim ELECTRONICA As String
    Dim tipodoc As String
    Dim fecha As Date
    Dim fechacom As Date
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "retencion"
    campos(9, 0) = "total"
    campos(10, 0) = "añocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = "electronica"
    campos(14, 0) = "activo"
    campos(15, 0) = "fechadigitacion"
    campos(16, 0) = "folio"
    campos(17, 0) = "impuestoespecifico"
    campos(18, 0) = "usuario"
    campos(19, 0) = "fechatraspaso"
    campos(20, 0) = "horatraspaso"
    campos(21, 0) = ""
    
    
    
 
    If Grid1.Cell(LINEA, 1).text = "FA" Then TIPOCON = "1": ELECTRONICA = "N": tipodoc = "FC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "ND" Then TIPOCON = "2": ELECTRONICA = "N": tipodoc = "DC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NC" Then TIPOCON = "3": ELECTRONICA = "N": tipodoc = "NC": DH = "D": DH2 = "H"
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "4": ELECTRONICA = "S": tipodoc = "FC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "5": ELECTRONICA = "S": tipodoc = "DC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "6": ELECTRONICA = "S": tipodoc = "NC": DH = "D": DH2 = "H"
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(3, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(4, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(5, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 7).text, ",", ".")
    exentos = CDbl(Grid1.Cell(LINEA, 8).text) + CDbl(Grid1.Cell(LINEA, 9).text) + CDbl(Grid1.Cell(LINEA, 10).text) + CDbl(Grid1.Cell(LINEA, 11).text) + CDbl(Grid1.Cell(LINEA, 12).text) + CDbl(Grid1.Cell(LINEA, 13).text)
    campos(7, 1) = Str(exentos)
    campos(8, 1) = "0"
    campos(9, 1) = Replace(Grid1.Cell(LINEA, 14).text, ",", ".")
    
    
    campos(10, 1) = Grid1.Cell(LINEA, 19).text
    campos(11, 1) = Grid1.Cell(LINEA, 18).text
    campos(12, 1) = "CENTRALIZACION AUTOMATICA"
        
    campos(13, 1) = ELECTRONICA
    campos(14, 1) = "N"
    campos(15, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(16, 1) = LEERULTIMOFOLIO(campos(11, 1), campos(10, 1))
    campos(17, 1) = "0"
    campos(18, 1) = USUARIOSISTEMA
    campos(19, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(20, 1) = Time
    
    
    condicion = ""
    campos(0, 2) = "facturasdecompras"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb

    Call sqlconta.sqlconta(op, condicion)
    k = sqlconta.status
    fecha = Format(campos(3, 1), "yyyy-mm-dd")
    
    
    fechacom = Format(fechasistema, "yyyy-mm") + "-" + "01"
    If fecha >= fechacom Then
    fechacom = fecha
    End If
    
    If TIPOCON = "3" Or TIPOCON = "6" Then
    CUENTA2 = "11200044"
    Else
    CUENTA2 = CUENTAPROVEEDOR
    
    End If
    
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "001", fechacom, CUENTA2, "", campos(4, 1), "", "CENTRALIZA DOCUMENTO DE COMPRAS " + Grid1.Cell(LINEA, 1).text, tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(9, 1), DH, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "002", fechacom, ivacredito, "", campos(4, 1), "", "CENTRALIZACION I.V.A", tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(6, 1), DH2, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    
    Call grabardetallefactura(LINEA, tipo, ORDEN, fechacom, campos(11, 1), campos(10, 1))


End Sub

Sub grabardetallefactura(LINEA, tipo, ORDEN, fecha, MES, año)
    
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim ilas As Double
    Dim CRCC As String
    Dim cuenta As String
    Dim DH As String
    Dim NOMBRE As String
    Dim tipodoc As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = "fechacreacion"
    campos(11, 0) = ""
    If localfiltro = "00" Then CRCC = "0101"
    If localfiltro = "41" Then CRCC = "0104"
    If localfiltro = "17" Then CRCC = "0101"
    If localfiltro = "42" Then CRCC = "0101"
    
    
    If Grid1.Cell(LINEA, 1).text = "FA" Then TIPOCON = "1": tipodoc = "FC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "ND" Then TIPOCON = "2": tipodoc = "DC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NC" Then TIPOCON = "3": tipodoc = "NC": DH = "H"
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "4": tipodoc = "FC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "5": tipodoc = "DC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "6": tipodoc = "NC": DH = "H"
    
    If tipo = "DI" Then cuenta = "11350008": NOMBRE = "DIARIOS"
    If tipo = "ME" Then cuenta = "11350001": NOMBRE = "MERCADERIAS"
    If tipo = "CI" Then cuenta = "11350007": NOMBRE = "CIGARRILLOS"
    If tipo = "FR" Then cuenta = "11350002": NOMBRE = "FRUTAS"
    If tipo = "CA" Then cuenta = "11350003": NOMBRE = "CARNICERIA"
    If tipo = "FI" Then cuenta = "11350004": NOMBRE = "FIAMBRERIA"
    If tipo = "PA" Then cuenta = "11350007": NOMBRE = "PANADERIA"
    If tipo = "EM" Then cuenta = "11350006": NOMBRE = "MATERIAL EMPAQUE"
    

Rem CALCULA NETOS

    lin = 3
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = cuenta
    campos(5, 1) = "O/C " + ORDEN + " " + NOMBRE
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(7, 1) = DH
    campos(8, 1) = leerdatoslocal(localfiltro, "codigocrcc")
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
  
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", campos(3, 1), "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    
    
Rem CALCULA ILAS CERVEZAS

    ilas = CDbl(Grid1.Cell(LINEA, 8).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailacervezas")
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA CERVEZAS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
Rem CALCULA ILAS refrescos

    ilas = CDbl(Grid1.Cell(LINEA, 9).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailarefrescos")
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA REFRESCOS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA ILAS vinos

    ilas = CDbl(Grid1.Cell(LINEA, 10).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailavinos")
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA VINOS "
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA ILAS licores

    ilas = CDbl(Grid1.Cell(LINEA, 11).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailalicores")
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA LICORES "
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA HARINA
    ilas = CDbl(Grid1.Cell(LINEA, 12).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentaharina")
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO HARINAS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA carne
    ilas = CDbl(Grid1.Cell(LINEA, 13).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentacarne")
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO CARNE"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
    
   
    
    
End Sub

Public Function leefactura(tipo, numero, rut) As String

    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    If tipo = "FA" Then tipo = "1"
    If tipo = "ND" Then tipo = "2"
    If tipo = "NC" Then tipo = "3"
    If tipo = "FAE" Then tipo = "4"
    If tipo = "NDE" Then tipo = "5"
    If tipo = "NCE" Then tipo = "6"
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
    campos(0, 2) = "facturasdecompras"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leefactura = "1"
    
    Else
    leefactura = "0"
    
    End If
    
    

End Function

Sub crearcuentacorriente(rut)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion

            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,mc.giro,mc.fono1 "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            


End Sub
'cSql.SQL = "INSERT INTO l_movimientos_detalle_" & empresaactiva & " "
'            cSql.SQL = cSql.SQL & "(tipo, numero, linea, fecha, rut, codigo, descripcion, cantidad, unidades, precio, total, costoventa, bodega, bodegatraspaso, uxc) "
'            cSql.SQL = cSql.SQL & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.unidades, dd.precio, dd.total, dd.pcosto, dd.bodega, dd.bodega, ROUND(dd.unidades / dd.cantidad, 0) "
'            cSql.SQL = cSql.SQL & "FROM " & baseVentas & rubro & ".sv_documento_detalle_" + empresaactiva + " as dd "
'            cSql.SQL = cSql.SQL & "WHERE dd.local = '" & empresaactiva & "' AND dd.tipo = '" & v.detalle.tipo & "' AND dd.numero = '" & v.detalle.numero & "'"
'            cSql.Execute

Public Function LEERULTIMOFOLIO(mesconta, añoconta) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select max(folio) from facturasdecompras where mescontable = '" & Format(mesconta, "00") & "' AND añocontable = '" & añoconta & "' "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
        If resultados(0) <> "NULO" Then
        LEERULTIMOFOLIO = resultados(0) + 1
        Else
        LEERULTIMOFOLIO = "0000000001"
        End If
        
    End If
    
End Function
Public Function LEERMONTOIMPUESTO(tipo, numero, ORDEN, cuenta) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
 
        Set csql.ActiveConnection = gestionrubro

            csql.sql = "select monto from l_ordendecompra_impuestos_" + localfiltro + " where cuenta = '" & cuenta & "' and tipo='" + tipo + "' and numero='" + numero + "' and numeroorden='" + ORDEN + "' "
            
            csql.Execute
    LEERMONTOIMPUESTO = 0
    If csql.RowsAffected > 0 Then
    
    Set resultados = csql.OpenResultset
    LEERMONTOIMPUESTO = resultados(0)
    
    End If
    
End Function
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = MES
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub



Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Static palabra As String
    Dim i As Integer
    Dim largo As Integer
    If KeyAscii = 13 Then
        palabra = ""
    Else
        palabra = palabra + UCase(Chr(KeyAscii))
        largo = Len(palabra)
        For i = 1 To Grid1.Rows - 1
            If Mid(Grid1.Cell(i, 16).text, 1, largo) = palabra Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                
                
                Exit For
            End If
        Next i
    End If
    
End Sub

Private Sub ORDEN_GotFocus()
Call cargatexto(ORDEN)
End Sub

Private Sub ORDEN_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(ORDEN)
Call BUSCAR_Click


End If

End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

 
Sub ayudahora(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("horatraspaso", "fechatraspaso")
    cabezas = Array("HORA", "FECHA")
    largo = Array("8s", "10s")
    mensajeAyuda = "Ayuda Horas"
    cfijo = "fechatraspaso='" & DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text & "' and usuario='" & USUARIOSISTEMA & "' group by fechatraspaso,mid(horatraspaso,1,5)"
    Call cargaAyudaT(Servidor, clientesistema + "conta" & empresaactiva, Usuario, password, "facturasdecompras ", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus
no:

End Sub

 

Private Sub txtHoraDesde_GotFocus()
    Call cargatexto(txtHoraDesde)
End Sub

Private Sub txtHoraDesde_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(DESDE3, txtHoraHasta, KeyCode)
    If KeyCode = vbKeyF2 Then
        Call ayudahora(txtHoraDesde)
    End If
End Sub

Private Sub txtHoraDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtHoraHasta.SetFocus
    End If
End Sub

Private Sub txtHoraHasta_GotFocus()
    Call cargatexto(txtHoraHasta)
End Sub

Private Sub txtHoraHasta_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(txtHoraDesde, txtHoraHasta, KeyCode)
     If KeyCode = vbKeyF2 Then
        Call ayudahora(txtHoraHasta)
    End If
End Sub

Private Sub txtHoraHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2.SetFocus
    End If
End Sub
Sub totallibro(infogrilla As grillainformes)
    
    Dim TOTALge As Double
      lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 7, lin, 12).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 6).text = "TOTAL DOCUMENTOS  " & Format(totaldocumentos, "###,###,###")
        infogrilla.Grid1.Cell(lin, 7).text = total(1)
        infogrilla.Grid1.Cell(lin, 8).text = total(2)
        infogrilla.Grid1.Cell(lin, 9).text = total(3)
        infogrilla.Grid1.Cell(lin, 10).text = total(4)
        infogrilla.Grid1.Cell(lin, 11).text = total(5)
        infogrilla.Grid1.Cell(lin, 12).text = total(6)
    
    TOTALge = 0
    lin = lin + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 11
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 5, lin + 9, 12).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 5).text = "Cant."
    infogrilla.Grid1.Cell(lin, 6).text = "Documentos"
    infogrilla.Grid1.Cell(lin, 7).text = "Neto"
    infogrilla.Grid1.Cell(lin, 8).text = "i.v.a"
    infogrilla.Grid1.Cell(lin, 9).text = "exento"
    infogrilla.Grid1.Cell(lin, 10).text = "diesel"
    infogrilla.Grid1.Cell(lin, 11).text = "retencion"
    infogrilla.Grid1.Cell(lin, 12).text = "total"
    
    
    
    For k = 1 To 9
    lin = lin + 1
    
    infogrilla.Grid1.Cell(lin, 6).text = TIPOS(k)
    infogrilla.Grid1.Cell(lin, 5).text = Format(detalle(k, 1), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 2), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 3), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 4), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 10).text = Format(detalle(k, 5), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 11).text = Format(detalle(k, 6), "###,###,##0")
    infogrilla.Grid1.Cell(lin, 12).text = Format(detalle(k, 7), "###,###,##0")
    
    Next k
    
    
    
    
    
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
    lin = lin + 2
    For k = 1 To canplan
    If plan(k, 3) <> 0 Then
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 5).text = plan(k, 1)
        infogrilla.Grid1.Cell(lin, 6).text = plan(k, 2)
        infogrilla.Grid1.Cell(lin, 7).text = plan(k, 3)
        TOTALge = TOTALge + plan(k, 3)
        End If
    Next k
        lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Range(lin, 6, lin, 7).Borders(cellEdgeTop) = cellThin
        
        
        
        
        
        infogrilla.Grid1.Cell(lin, 6).text = "TOTAL DETALLE"
         infogrilla.Grid1.Cell(lin, 7).text = TOTALge
               
    End Sub
