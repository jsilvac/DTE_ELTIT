VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form control02 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IMPRIME HOJAS PARA TIMBRAR"
   ClientHeight    =   6975
   ClientLeft      =   240
   ClientTop       =   1290
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   4320
      TabIndex        =   17
      Top             =   6240
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
   Begin FlexCell.Grid Grid1 
      Height          =   420
      Left            =   1260
      TabIndex        =   7
      Top             =   7830
      Visible         =   0   'False
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   741
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   6855
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   12091
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
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "Grabar Datos Timbraje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   240
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6120
         Width           =   1410
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF8080&
         Caption         =   "Datos Propuestos"
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   810
         Width           =   1860
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Datos Grabados"
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   315
         Width           =   1860
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Vista previa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   12
         Top             =   5760
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6120
         Width           =   1995
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   5760
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   4230
         Left            =   45
         TabIndex        =   1
         Top             =   1395
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   7461
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
         Begin FlexCell.Grid Grid2 
            Height          =   3075
            Left            =   45
            TabIndex        =   13
            Top             =   855
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   5424
            BackColorBkg    =   16777215
            BackColorFixed  =   16761024
            BackColorFixedSel=   16777215
            BackColorScrollBar=   16777215
            BackColorSel    =   16777215
            Cols            =   2
            DefaultFontSize =   8.25
            GridColor       =   16777215
            Rows            =   30
         End
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
            Locked          =   -1  'True
            TabIndex        =   2
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
            Left            =   855
            TabIndex        =   3
            Top             =   315
            Width           =   3255
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1035
         Left            =   810
         TabIndex        =   5
         Top             =   270
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1826
         BackColor       =   16744576
         Caption         =   "cantidad de hojas"
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
         Begin VB.TextBox termino 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   9
            Top             =   585
            Width           =   1575
         End
         Begin VB.TextBox inicio 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   405
            MaxLength       =   10
            TabIndex        =   6
            Top             =   585
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Folio Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2430
            TabIndex        =   11
            Top             =   315
            Width           =   1590
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Folio Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   405
            TabIndex        =   10
            Top             =   315
            Width           =   1590
         End
      End
   End
End
Attribute VB_Name = "control02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(10, 20) As String
Private sumas(10) As Double
Private suma(10) As Double
Private sumas2(10) As Double
Private sumas3(10) As Double
Private montos(5) As Double
Private lin As Double


 

Private Sub Command1_Click()
Dim o As Double
If inicio.text <> "" And termino.text <> "" Then
For o = CDbl(inicio.text) To CDbl(termino.text)
Call cabezas(o)
If Check1.Value <> "1" Then
Grid1.DirectPrint
Else
Grid1.PrintPreview
End If
Next o
End If
End Sub

Private Sub COMMAND2_Click()
original


End Sub

Private Sub Command3_Click()
PROPUESTA

End Sub

Private Sub Command4_Click()
For k = 1 To 10
Call modificatimbraje("timbraje" & k, Grid2.Cell(k, 1).text)
Next k
End Sub
Sub modificatimbraje(campo, dato)
    Dim campos(10, 10) As String
    Dim condicion As String
    
    Dim netos As Double
    Dim DH As String
    campos(0, 0) = campo
    campos(1, 0) = ""
    campos(0, 1) = dato
    
    
    
    condicion = "codigoempresa='" + empresaactiva + "' "
    campos(0, 2) = "maestroempresas"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato1)
    
End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub


Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    
    empresanombre.Caption = sqlconta.response(1, 3)

no:
End Sub

Private Sub Form_Load()
CENTRAR Me

 Call Conectar_BD
 Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
empresanombre.Caption = nombreempresa

Grid2.Column(0).Width = 0
Grid2.Column(1).Width = 70 * 7
Grid2.Column(1).MaxLength = 70
Grid2.Rows = 11
Grid2.Cell(0, 1).text = "DETALLE TIMBRAJE"
  
           
 original
 

End Sub
Sub PROPUESTA()
Grid2.Cell(1, 1).text = leerdatos(conta, "maestroempresas", "nombre", "codigoempresa='" + empresaactiva + "' ")
Grid2.Cell(2, 1).text = leerdatos(conta, "maestroempresas", "rut", "codigoempresa='" + empresaactiva + "' ")
Grid2.Cell(3, 1).text = leerdatos(conta, "maestroempresas", "direccion", "codigoempresa='" + empresaactiva + "' ")
Grid2.Cell(4, 1).text = leerdatos(conta, "maestroempresas", "comuna", "codigoempresa='" + empresaactiva + "' ")
Grid2.Cell(5, 1).text = leerdatos(conta, "maestroempresas", "ciudad", "codigoempresa='" + empresaactiva + "' ")
Grid2.Cell(6, 1).text = leerdatos(conta, "maestroempresas", "representantelegal", "codigoempresa='" + empresaactiva + "' ")
For k = 7 To 10
Grid2.Cell(k, 1).text = ""
Next k

End Sub
Sub original()
For k = 1 To 10
Grid2.Cell(k, 1).text = leerdatos(conta, "maestroempresas", "timbraje" & k, "codigoempresa='" + empresaactiva + "' ")
Next k
End Sub

Sub cabezas(FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


    'Report Title 1
        For k = 1 To 10
        If Grid2.Cell(k, 1).text <> "" Then
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = Grid2.Cell(k, 1).text
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 7
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        End If
    Next k
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "FOLIO :" & FOLIO
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = cellRight
        Grid1.ReportTitles.Add objReportTitle
    
With Grid1.PageSetup
        
        Rem If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .LeftMargin = 1
        .RightMargin = 1
        
        .TopMargin = 0.5
        .BottomMargin = 0.5
        
        
        
End With
For k = 1 To Grid1.PageSetup.PaperSizes.Count
        If UCase(Grid1.PageSetup.PaperSizes.item(k).PaperName) = "CARTA" Then
            Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.item(k).Kind
            Exit For
        End If
    Next k
    
End Sub

Private Sub inicio_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
termino.SetFocus

End If

End Sub
Private Sub TERMINO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Command1.SetFocus

End If

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
