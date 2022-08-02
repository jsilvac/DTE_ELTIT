VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0006 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DEVOLUCIONES"
   ClientHeight    =   9135
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   13410
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   609
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   894
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   10320
      TabIndex        =   32
      Top             =   0
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   33
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   4605
      Left            =   180
      TabIndex        =   19
      Top             =   3000
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   8123
      BackColor       =   16761024
      Caption         =   " Detalle Comprobante"
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
      Alignment       =   1
      Begin FlexCell.Grid Grid1 
         Height          =   4200
         Left            =   90
         TabIndex        =   20
         Top             =   315
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   7408
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2715
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   4789
      BackColor       =   16744576
      Caption         =   "DATOS DEVOLUCION"
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
      Alignment       =   1
      Begin VB.OptionButton TIPO 
         BackColor       =   &H00FF8080&
         Caption         =   "GUIA ELECTRONICA"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   36
         Top             =   720
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton TIPO 
         BackColor       =   &H00FF8080&
         Caption         =   "GUIA NORMAL"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   35
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmd1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ver Detalle"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2280
         Width           =   1320
      End
      Begin VB.TextBox HASTA3 
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
         Left            =   10680
         MaxLength       =   4
         TabIndex        =   30
         Tag             =   "fecha"
         Top             =   1395
         Width           =   615
      End
      Begin VB.TextBox HASTA2 
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
         Left            =   10300
         MaxLength       =   2
         TabIndex        =   29
         Tag             =   "fecha"
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox HASTA1 
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
         Left            =   9960
         MaxLength       =   2
         TabIndex        =   28
         Tag             =   "fecha"
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox DESDE3 
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
         Left            =   2480
         MaxLength       =   4
         TabIndex        =   25
         Tag             =   "fecha"
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox DESDE2 
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
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   24
         Tag             =   "fecha"
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox DESDE1 
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
         Left            =   1725
         MaxLength       =   2
         TabIndex        =   23
         Tag             =   "fecha"
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox dato2 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "tipo"
         Top             =   690
         Width           =   1215
      End
      Begin XPFrame.FrameXp GLOSACTACTE 
         Height          =   255
         Left            =   2205
         TabIndex        =   18
         Top             =   315
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   450
         BackColor       =   49344
         Caption         =   ""
         CaptionEstilo3D =   1
         BackColor       =   49344
         ForeColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox dato1 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1725
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "tipo"
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "direccion"
         Top             =   1755
         Width           =   1575
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1725
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "rut"
         Top             =   1035
         Width           =   1095
      End
      Begin VB.TextBox dato5 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9960
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Comuna"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9960
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "ciudad"
         Top             =   1035
         Width           =   1575
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "fono"
         Top             =   1755
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   90
         TabIndex        =   27
         Top             =   1380
         Width           =   1530
      End
      Begin VB.Label lblnombreproveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   26
         Top             =   1035
         Width           =   4455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " numero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   90
         TabIndex        =   22
         Top             =   690
         Width           =   1530
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   90
         TabIndex        =   17
         Top             =   1035
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   315
         Width           =   1530
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   1755
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   7890
         TabIndex        =   14
         Top             =   675
         Width           =   1890
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nº Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   7890
         TabIndex        =   13
         Top             =   1035
         Width           =   1890
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   7890
         TabIndex        =   12
         Top             =   1395
         Width           =   1890
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   7890
         TabIndex        =   11
         Top             =   1755
         Width           =   1890
      End
      Begin VB.Label dv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2925
         TabIndex        =   10
         Top             =   1035
         Width           =   255
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   13380
      TabIndex        =   8
      Top             =   9135
      Width           =   13410
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "prove0006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private MODIFI As Integer

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub
Sub todaslasrelaciones()
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigoempresa "
        csql.sql = csql.sql + "FROM maestroempresas "
        csql.sql = csql.sql + "order by codigoempresa "
        csql.Execute
        Grid1.Rows = 1
    If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
        While resultados.EOF = False
        Call leerelaciones(Me, Grid1, dato2.text + DV.Caption, resultados(0))
        resultados.MoveNext
        Wend
            resultados.Close
        Set resultados = Nothing
    End If
    csql.Close
    Set csql = Nothing
End Sub
 Private Sub imprimir()
If Grid1.Rows > 1 Then
Call Titulos("LISTADO DE SALDOS RELACIONADOS ")
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.HeaderMargin = 0.5
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.TopMargin = 3
Grid1.PageSetup.LeftMargin = 0.1
Grid1.PageSetup.RightMargin = 0.1
Grid1.PageSetup.BottomMargin = 3
Grid1.PageSetup.FooterMargin = 2
Grid1.PageSetup.BlackAndWhite = True

Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid1.PrintPreview
End If
End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.Orientation = cellLandscape
    Grid1.PageSetup.PrintTitleRows = 0
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    Grid1.PageSetup.HeaderAlignment = CellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "EMITIDO  :  " & Format(fechasistema, "dd-MM-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = dato2.text + "-" + DV.Caption & "  " & dato4.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub

Private Sub dp_Click()
If dato4.text <> "" Then
DATOSPAGO.Show vbModal
End If
End Sub

Private Sub cmd1_Click()
If dato2.text <> "" Then
If tipo(0).Value = True Then Call verdetalle(dato1.text, "DM", dato2.text)
If tipo(1).Value = True Then Call verdetalle(dato1.text, "D1", dato2.text)
'Call verdetalle(dato1.text, dato2.text)
End If
End Sub

Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub

Private Sub dato5_GotFocus()
Call cargatexto(DATO5)
End Sub
Private Sub dato6_GotFocus()
Call cargatexto(dato6)
End Sub
 
Private Sub dato8_GotFocus()
Call cargatexto(dato8)
End Sub
Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudalocales(dato1)
   
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudactacte(dato4)
    Call flechas(dato1, dato4, KeyCode)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, DATO5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(DATO5, dato8, KeyCode)
End Sub
 
 
 
 
Private Sub MANUAL_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
DOCU(1) = "ACTIVO"
DOCU(2) = "PASIVO"
DOCU(3) = "RESULTADO"
CANDO = 3

Rem Call RECUPERAFECHA
Call CARGAPERMISO(Me.Name)
 
 CARGAGRILLADETALLE

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato1.text <> "" Then Call ceros(dato1): GLOSACTACTE.Caption = leernombrelocal(dato1.text): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato2.text) <> 0 Then Call ceros(dato2): leer
 
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato4.text) <> "" Then Call Pregunta(dato4, DATO5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(DATO5.text) <> "" Then sc = 1: Call Pregunta(DATO5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato6.text) <> "" Then sc = 1: Call Pregunta(dato6, dato8)
End Sub
  Sub leer()
    campos(0, 0) = "rut"
    campos(1, 0) = "fecha"
    campos(2, 0) = "monto"
    campos(3, 0) = "tipoco"
    campos(4, 0) = "numeroco"
    campos(5, 0) = "fechaco"
    campos(6, 0) = "montoco"
    campos(7, 0) = ""
    campos(0, 2) = "devoluciones_proveedores"
    condicion = "local=" + "'" + dato1.text + "' and numero='" & dato2.text & "' "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato2.SetFocus: GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
     
        
no:
End Sub
Sub leersiguiente()
    campos(0, 0) = "rut"
    campos(1, 0) = "fecha"
    campos(2, 0) = "monto"
    campos(3, 0) = "tipoco"
    campos(4, 0) = "numeroco"
    campos(5, 0) = "fechaco"
    campos(6, 0) = "montoco"
    campos(7, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut>'" + dato2.text + DV.Caption + "' and año='" + Format(fechasistema, "yyyy") + "' order by tipo,rut asc "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
 
no:
   
    
End Sub
Sub leeranterior()
    campos(0, 0) = "rut"
    campos(1, 0) = "fecha"
    campos(2, 0) = "monto"
    campos(3, 0) = "tipoco"
    campos(4, 0) = "numeroco"
    campos(5, 0) = "fechaco"
    campos(6, 0) = "montoco"
    campos(7, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut<'" + dato2.text + DV.Caption + "' and año='" + Format(fechasistema, "yyyy") + "' order by tipo,rut desc "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
  
no:
       
End Sub

Sub carga()
    habilita (True)
    dato3.text = Mid(sqlconta.response(0, 3), 1, 9)
    DV.Caption = Mid(sqlconta.response(0, 3), 10, 1)
    DESDE1.text = Mid(sqlconta.response(1, 3), 1, 2)
    DESDE2.text = Mid(sqlconta.response(1, 3), 4, 2)
    DESDE3.text = Mid(sqlconta.response(1, 3), 7, 4)
    dato4.text = Format(sqlconta.response(2, 3), " $ ###,###,##0")
    DATO5.text = sqlconta.response(3, 3)
    dato6.text = sqlconta.response(4, 3)
    If IsNull(sqlconta.response(5, 3)) = True Then
        HASTA1.text = "00"
        HASTA2.text = "00"
        HASTA3.text = "0000"
    Else
        HASTA1.text = Mid(sqlconta.response(5, 3), 1, 2)
        HASTA2.text = Mid(sqlconta.response(5, 3), 4, 2)
        HASTA3.text = Mid(sqlconta.response(5, 3), 7, 4)
    End If
    dato8.text = Format(sqlconta.response(6, 3), " $ ###,###,##0")
     lblnombreproveedor.Caption = LEERNOMBREPROVEEDOR(sqlconta.response(0, 3))
 
End Sub

Sub habilita(ByVal condicion As Boolean)
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato4.Locked = condicion
    DATO5.Locked = condicion
    dato6.Locked = condicion
    dato8.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    dato6.Enabled = condicion
    dato8.Enabled = condicion
End Sub
Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub
Sub ayudalocales(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("10s", "40s")
    cfijo = "no"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda Locales"
        
    Call cargaAyudaT(Servidor, basebus, Usuario, password, clientesistema & "gestion" & ".g_maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
'    CAMPOS(0, 0) = dato1.Tag
'    CAMPOS(1, 0) = dato2.Tag
'    CAMPOS(2, 0) = dato4.Tag
'    CAMPOS(3, 0) = dato5.Tag
'    CAMPOS(4, 0) = dato6.Tag
'    CAMPOS(5, 0) = dato7.Tag
'    CAMPOS(6, 0) = ""
'    CAMPOS(0, 1) = dato1.text
'    CAMPOS(1, 1) = dato2.text + dv.Caption
'    CAMPOS(2, 1) = dato4.text
'    CAMPOS(3, 1) = dato5.text
'    CAMPOS(4, 1) = dato6.text
'    CAMPOS(5, 1) = dato7.text
'    CAMPOS(6, 1) = dato8.text
'
'
'    CAMPOS(0, 2) = "cuentascorrientes"
'    If MODIFI = 1 Then condicion = "tipo=" + "'" + dato1.text + "' and rut ='" + dato2.text + dv.Caption + "' and año='" + Format(fechasistema, "yyyy") + "'"
'    If MODIFI = 1 Then op = 3 Else op = 2
'    sqlconta.response = CAMPOS
'    Set sqlconta.conexion = contadb
'    Call sqlconta.sqlconta(op, condicion)
'
'    If MODIFI = 0 Then grabar2
'
    End Sub
Sub grabar2()
      
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = "debe01"
    campos(6, 0) = "debe02"
    campos(7, 0) = "debe03"
    campos(8, 0) = "debe04"
    campos(9, 0) = "debe05"
    campos(10, 0) = "debe06"
    campos(11, 0) = "debe07"
    campos(12, 0) = "debe08"
    campos(13, 0) = "debe09"
    campos(14, 0) = "debe10"
    campos(15, 0) = "debe11"
    campos(16, 0) = "debe12"
    campos(17, 0) = "haber01"
    campos(18, 0) = "haber02"
    campos(19, 0) = "haber03"
    campos(20, 0) = "haber04"
    campos(21, 0) = "haber05"
    campos(22, 0) = "haber06"
    campos(23, 0) = "haber07"
    campos(24, 0) = "haber08"
    campos(25, 0) = "haber09"
    campos(26, 0) = "haber10"
    campos(27, 0) = "haber11"
    campos(28, 0) = "haber12"
    
    campos(29, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text + DV.Caption
    campos(2, 1) = año

    For k = 3 To 28
    campos(k, 1) = "0"
    Next k
    campos(0, 2) = "saldosctacte"
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub

Sub ELIMINAR()
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut=" + "'" + dato2.text + DV.Caption + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub



Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
'If command = "modifica" Then modifica
'If command = "elimina" Then elimina

''If command = "siguiente" Then leersiguiente
'If command = "anterior" Then leeranterior
'If command = "movimientos" Then movimientos



End Sub
Sub ELIMINA()
If saldoglobal = 0 Then
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus
Else
MsgBox ("IMPOSIBLE ELIMINAR RUT CON DATOS")
End If
End Sub

Sub modifica()
disponible (True)
habilita (False)
dato1.Enabled = False
dato2.Enabled = False
dato4.SetFocus
MODIFI = 1

End Sub
Sub retorno()

disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
MODIFI = 0
no:
Grid1.Rows = 1
    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    DV.Caption = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato8.text = ""
    
    DESDE1.text = ""
    DESDE2.text = ""
    DESDE3.text = ""
   
    HASTA1.text = ""
    HASTA2.text = ""
    HASTA3.text = ""
    lblnombreproveedor.Caption = ""
    GLOSACTACTE.Caption = ""
End Sub



  

Sub LEERSALDOS()
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = "debe01"
    campos(6, 0) = "debe02"
    campos(7, 0) = "debe03"
    campos(8, 0) = "debe04"
    campos(9, 0) = "debe05"
    campos(10, 0) = "debe06"
    campos(11, 0) = "debe07"
    campos(12, 0) = "debe08"
    campos(13, 0) = "debe09"
    campos(14, 0) = "debe10"
    campos(15, 0) = "debe11"
    campos(16, 0) = "debe12"
    campos(17, 0) = "haber01"
    campos(18, 0) = "haber02"
    campos(19, 0) = "haber03"
    campos(20, 0) = "haber04"
    campos(21, 0) = "haber05"
    campos(22, 0) = "haber06"
    campos(23, 0) = "haber07"
    campos(24, 0) = "haber08"
    campos(25, 0) = "haber09"
    campos(26, 0) = "HABER10"
    campos(27, 0) = "HABER11"
    campos(28, 0) = "HABER12"
    campos(29, 0) = ""
    condicion = "tipo=" + "'" + dato1.text + "' and rut='" + dato2.text + DV.Caption + "' and año='" + Mid(fechasistema, 7, 4) + "'"
    campos(0, 2) = "saldosctacte"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
    
   Rem  If sqlconta.status = 4 Then Stop
 
End Sub
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub

Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & dato1.text & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then dato2.SetFocus: GoTo no
    dato4.Enabled = True
    dato2.text = Mid(pivote.text, 1, 9)
    DV.Caption = Mid(pivote.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus
no:

End Sub

 
Sub CARGAGRILLADETALLE()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "DESCRIPCION"
    formatogrilla2(1, 3) = "CANTIDAD"
    formatogrilla2(1, 4) = "UXC"
    formatogrilla2(1, 5) = "UNIDADES"
    formatogrilla2(1, 6) = "PRECIO"
    formatogrilla2(1, 7) = "DESC"
    formatogrilla2(1, 8) = "TOTAL"
'     select linea,codigo,descripcion,cantidad,uxc,unidades,precio,descuento,total from l_movimientos_detalle_41 where tipo='DM' and numero='0000005022'
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "13"
    formatogrilla2(2, 2) = "40"
    formatogrilla2(2, 3) = "7"
    formatogrilla2(2, 4) = "7"
    formatogrilla2(2, 5) = "7"
    formatogrilla2(2, 6) = "7"
    formatogrilla2(2, 7) = "7"
    formatogrilla2(2, 8) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    formatogrilla2(3, 8) = "N"
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 6) = " ###,###,##0"
    formatogrilla2(4, 7) = " ###,###,##0"
    formatogrilla2(4, 8) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 9
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.BackColorFixed = RGB(90, 158, 214)
    Grid1.BackColorFixedSel = RGB(110, 180, 230)
    Grid1.BackColorBkg = RGB(90, 158, 214)
    Grid1.BackColorScrollBar = RGB(231, 235, 247)
    Grid1.BackColor1 = RGB(231, 235, 247)
    Grid1.BackColor2 = RGB(239, 243, 255)
    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 8
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
    Next k
 
    End Sub
 
Sub verdetalle(loc, tipoguia, numero)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo, tipo2 As String
tipo = tipoguia


Set csql.ActiveConnection = contadb

csql.sql = "select linea,codigo,descripcion,cantidad,uxc,unidades,precio,descuento,total "
csql.sql = csql.sql & "from " & clientesistema & "gestion" & leerrubro(dato1.text) & ".l_movimientos_detalle_" & loc & " where tipo='" & tipo & "' and numero='" & numero & "' order by linea"
csql.Execute

If csql.RowsAffected > 0 Then
    Grid1.Rows = csql.RowsAffected + 1
    Set resultados = csql.OpenResultset
    
    While Not resultados.EOF
        Grid1.Cell(resultados(0), 1).text = resultados(1)
        Grid1.Cell(resultados(0), 2).text = resultados(2)
        Grid1.Cell(resultados(0), 3).text = resultados(3)
        Grid1.Cell(resultados(0), 4).text = resultados(4)
        Grid1.Cell(resultados(0), 5).text = resultados(5)
        Grid1.Cell(resultados(0), 6).text = resultados(6)
        Grid1.Cell(resultados(0), 7).text = resultados(7)
        Grid1.Cell(resultados(0), 8).text = resultados(8)
        resultados.MoveNext
    Wend
End If

csql.Close
Set csql = Nothing
Set resultados = Nothing
 
End Sub
Function leerrubro(loc) As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select rubro from " & clientesistema & "gestion.g_maestroempresas where "
    csql.sql = csql.sql & "codigo='" & loc & "' "
    csql.Execute
    
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    leerrubro = resultado(0)
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
