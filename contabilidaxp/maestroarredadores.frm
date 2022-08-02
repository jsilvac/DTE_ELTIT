VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form maestro05 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Arredadores"
   ClientHeight    =   10140
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   8370
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   676
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   3525
      Left            =   180
      TabIndex        =   29
      Top             =   5265
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   6218
      BackColor       =   16761024
      Caption         =   " Arrendatarios"
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
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   285
         Left            =   4050
         TabIndex        =   32
         Top             =   3195
         Width           =   1320
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Todas Las Empresas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   10350
         TabIndex        =   31
         Top             =   3105
         Width           =   2580
      End
      Begin FlexCell.Grid Grid1 
         Height          =   2760
         Left            =   90
         TabIndex        =   30
         Top             =   315
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   4868
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5010
      Left            =   180
      TabIndex        =   16
      Top             =   180
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8837
      BackColor       =   16744576
      Caption         =   "DATOS  "
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
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "tipo"
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox dato3 
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
         Top             =   1035
         Width           =   6015
      End
      Begin VB.TextBox dato2 
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
         TabIndex        =   1
         Tag             =   "nombre"
         Top             =   675
         Width           =   6015
      End
      Begin VB.TextBox dato4 
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
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Comuna"
         Top             =   1395
         Width           =   3255
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
         Left            =   1725
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "ciudad"
         Top             =   1755
         Width           =   3255
      End
      Begin VB.TextBox dato11 
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
         TabIndex        =   5
         Tag             =   "giro"
         Top             =   3920
         Width           =   495
      End
      Begin VB.TextBox dato7 
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
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "celular"
         Top             =   2475
         Width           =   1815
      End
      Begin VB.TextBox dato8 
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
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "fax"
         Top             =   2835
         Width           =   1815
      End
      Begin VB.TextBox dato6 
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
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "fono"
         Top             =   2115
         Width           =   1815
      End
      Begin VB.TextBox dato10 
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
         TabIndex        =   10
         Tag             =   "contacto"
         Top             =   3555
         Width           =   5895
      End
      Begin VB.TextBox dato9 
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
         TabIndex        =   9
         Tag             =   "email"
         Top             =   3195
         Width           =   5895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut"
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
         Left            =   100
         TabIndex        =   28
         Top             =   315
         Width           =   1530
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   100
         TabIndex        =   27
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         Left            =   100
         TabIndex        =   26
         Top             =   1035
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comuna"
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
         Left            =   100
         TabIndex        =   25
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ciudad"
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
         Left            =   100
         TabIndex        =   24
         Top             =   1755
         Width           =   1530
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Contable"
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
         Left            =   100
         TabIndex        =   23
         Top             =   3920
         Width           =   1530
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fono"
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
         Left            =   100
         TabIndex        =   22
         Top             =   2115
         Width           =   1530
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fax"
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
         Index           =   0
         Left            =   100
         TabIndex        =   21
         Top             =   2835
         Width           =   1530
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Email"
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
         Left            =   100
         TabIndex        =   20
         Top             =   3195
         Width           =   1530
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Celular"
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
         Left            =   100
         TabIndex        =   19
         Top             =   2475
         Width           =   1530
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contacto"
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
         Left            =   100
         TabIndex        =   18
         Top             =   3555
         Width           =   1530
      End
      Begin VB.Label dv 
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
         Left            =   2880
         TabIndex        =   17
         Top             =   315
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
      ScaleWidth      =   8340
      TabIndex        =   15
      Top             =   10140
      Width           =   8370
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3735
      Left            =   8400
      TabIndex        =   12
      Top             =   240
      Width           =   4695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SALDOS 
         Height          =   3495
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6165
         _Version        =   393216
         BackColor       =   16776436
         ForeColor       =   12582912
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   16107953
         BackColorSel    =   16777215
         ForeColorSel    =   16744576
         BackColorBkg    =   16776436
         GridColor       =   -2147483635
         GridColorFixed  =   12582912
         GridLinesFixed  =   1
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   3735
         Left            =   0
         Top             =   0
         Width           =   4695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   180
      TabIndex        =   11
      Top             =   8820
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   3735
      Left            =   8520
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "maestro05"
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

Private Sub Check1_Click()
If Check1.Value = "1" Then

todaslasrelaciones
Else
Grid1.Rows = 1
Call leerelaciones(Me, Grid1, dato2.text + dv.Caption, empresaactiva)
End If

End Sub
Sub todaslasrelaciones()
Dim cSql As New rdoQuery
Dim resultados As rdoResultset

Set cSql.ActiveConnection = conta
        cSql.sql = "SELECT codigoempresa "
        cSql.sql = cSql.sql + "FROM maestroempresas "
        cSql.sql = cSql.sql + "order by codigoempresa "
        cSql.Execute
        Grid1.Rows = 1
    If cSql.RowsAffected > 0 Then
        
        Set resultados = cSql.OpenResultset
        While resultados.EOF = False
        Call leerelaciones(Me, Grid1, dato2.text + dv.Caption, resultados(0))
        
        resultados.MoveNext
        
        Wend
        
            resultados.Close
        Set resultados = Nothing
            
    End If
    
    cSql.Close
    Set cSql = Nothing

End Sub

Private Sub Command1_Click()
impRimir
End Sub
Private Sub impRimir()
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
    
    'Logo
'    grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    grid1.ReportTitles.Add objReportTitle
    
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
    objReportTitle.text = dato2.text + "-" + dv.Caption & "  " & dato4.text
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

 

Private Sub dato1_GotFocus()
Call cargatexto(dato1)
End Sub

Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub
Private Sub dato4_GotFocus()
dv.Caption = rut(dato2.text)


If MODIFI = 0 And scrut <> "S" Then leer

Call cargatexto(dato4)
End Sub
Private Sub dato5_GotFocus()
Call cargatexto(dato5)
End Sub
Private Sub dato6_GotFocus()

Call cargatexto(dato6)
End Sub
Private Sub dato7_GotFocus()
Call cargatexto(dato7)
End Sub
Private Sub dato8_GotFocus()
Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()
Call cargatexto(dato9)
End Sub
Private Sub dato10_GotFocus()
Call cargatexto(dato10)
End Sub
Private Sub dato11_GotFocus()
Call cargatexto(dato11)
End Sub
Private Sub dato12_GotFocus()
Call cargatexto(dato12)
End Sub
Private Sub dato13_GotFocus()
Call cargatexto(dato13)
End Sub
Private Sub dato14_GotFocus()
Call cargatexto(dato14)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then Call ayudaarrendatarios(dato1)
    Call flechas(dato1, dato2, KeyCode)

End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
       Call flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato7, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato8, dato10, KeyCode)
End Sub
Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato9, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato10, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato11, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato13, dato14, KeyCode)
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
Call CARGAGRILLArelacion
dp.Visible = False

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato1.text) <> 0 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato2.text) <> 0 Then Call ceros(dato2): Call Pregunta(dato2, dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato4.text) <> "" Then Call Pregunta(dato4, dato5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato5.text) <> "" Then Call Pregunta(dato5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato6.text) <> "" Then Call Pregunta(dato6, dato7)
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato7.text) <> "" Then Call Pregunta(dato7, dato8)
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato8.text) <> "" Then Call Pregunta(dato8, dato9)
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato9, dato10)
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato10, dato11)
End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        grabar
        retorno
    End If
End Sub

Sub leer()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = dato4.Tag
    CAMPOS(3, 0) = dato5.Tag
    CAMPOS(4, 0) = dato6.Tag
    CAMPOS(5, 0) = dato7.Tag
    CAMPOS(6, 0) = dato8.Tag
    CAMPOS(7, 0) = dato9.Tag
    CAMPOS(8, 0) = dato10.Tag
    CAMPOS(9, 0) = dato11.Tag
    CAMPOS(10, 0) = ""
    CAMPOS(0, 2) = "maestro_arrendadores"
    condicion = "rut= '" & dato1.text & dv.Caption & "' "

    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
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
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = dato4.Tag
    CAMPOS(3, 0) = dato5.Tag
    CAMPOS(4, 0) = dato6.Tag
    CAMPOS(5, 0) = dato7.Tag
    CAMPOS(6, 0) = dato8.Tag
    CAMPOS(7, 0) = dato9.Tag
    CAMPOS(8, 0) = dato10.Tag
    CAMPOS(9, 0) = dato11.Tag
    CAMPOS(10, 0) = ""
    CAMPOS(0, 2) = "maestro_arrendadores"
    condicion = " rut > '" & dato1.text & dv.Caption & "' order by rut asc "

    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
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
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = dato4.Tag
    CAMPOS(3, 0) = dato5.Tag
    CAMPOS(4, 0) = dato6.Tag
    CAMPOS(5, 0) = dato7.Tag
    CAMPOS(6, 0) = dato8.Tag
    CAMPOS(7, 0) = dato9.Tag
    CAMPOS(8, 0) = dato10.Tag
    CAMPOS(9, 0) = dato11.Tag
    CAMPOS(10, 0) = ""
    CAMPOS(0, 2) = "maestro_arrendadores"
    condicion = " rut < '" & dato1.text & dv.Caption & "' order by rut desc "
    
    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
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
    dato1.text = sqlconta.response(0, 3)
    dato2.text = Mid(sqlconta.response(1, 3), 1, 9)
    dv.Caption = Mid(sqlconta.response(1, 3), 10, 1)
    dato4.text = sqlconta.response(2, 3)
    dato5.text = sqlconta.response(3, 3)
    dato6.text = sqlconta.response(4, 3)
    dato7.text = sqlconta.response(5, 3)
    dato8.text = sqlconta.response(6, 3)
    dato9.text = sqlconta.response(7, 3)
    dato10.text = sqlconta.response(8, 3)
    dato11.text = sqlconta.response(9, 3)
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato4.Locked = condicion
    dato5.Locked = condicion
    dato6.Locked = condicion
    dato7.Locked = condicion
    dato8.Locked = condicion
    dato9.Locked = condicion
    dato10.Locked = condicion
    dato11.Locked = condicion
 
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion

    dato4.Enabled = condicion
    dato5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    dato9.Enabled = condicion
    dato10.Enabled = condicion
    dato11.Enabled = condicion
 
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudaarrendatarios(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("rut", "nombre")
    largo = Array("11s", "40s")
    cfijo = "NO"
    cabezas = Array("Rut", "Nombre")
    mensajeAyuda = "Ayuda de Arrendatarios"
        
    Call cargaAyudaT(servidor, basebus, usuario, password, "maestro_arrendadores", dato1, CAMPOS, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = dato4.Tag
    CAMPOS(3, 0) = dato5.Tag
    CAMPOS(4, 0) = dato6.Tag
    CAMPOS(5, 0) = dato7.Tag
    CAMPOS(6, 0) = dato8.Tag
    CAMPOS(7, 0) = dato9.Tag
    CAMPOS(8, 0) = dato10.Tag
    CAMPOS(9, 0) = dato11.Tag
    CAMPOS(10, 0) = dato12.Tag
    CAMPOS(11, 0) = dato13.Tag
    CAMPOS(12, 0) = dato14.Tag
    CAMPOS(13, 0) = "año"
    CAMPOS(14, 0) = ""
    CAMPOS(0, 1) = dato1.text
    CAMPOS(1, 1) = dato2.text + dv.Caption
    CAMPOS(2, 1) = dato4.text
    CAMPOS(3, 1) = dato5.text
    CAMPOS(4, 1) = dato6.text
    CAMPOS(5, 1) = dato7.text
    CAMPOS(6, 1) = dato8.text
    CAMPOS(7, 1) = dato9.text
    CAMPOS(8, 1) = dato10.text
    CAMPOS(9, 1) = dato11.text
    CAMPOS(10, 1) = dato12.text
    CAMPOS(11, 1) = dato13.text
    CAMPOS(12, 1) = dato14.text
    CAMPOS(13, 1) = Format(fechasistema, "yyyy")
    
    CAMPOS(0, 2) = "cuentascorrientes"
    If MODIFI = 1 Then condicion = "tipo=" + "'" + dato1.text + "' and rut ='" + dato2.text + dv.Caption + "' and año='" + Format(fechasistema, "yyyy") + "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    
    If MODIFI = 0 Then grabar2
    
    End Sub
Sub grabar2()
      
    CAMPOS(0, 0) = "tipo"
    CAMPOS(1, 0) = "rut"
    CAMPOS(2, 0) = "año"
    CAMPOS(3, 0) = "debeanterior"
    CAMPOS(4, 0) = "haberanterior"
    CAMPOS(5, 0) = "debe01"
    CAMPOS(6, 0) = "debe02"
    CAMPOS(7, 0) = "debe03"
    CAMPOS(8, 0) = "debe04"
    CAMPOS(9, 0) = "debe05"
    CAMPOS(10, 0) = "debe06"
    CAMPOS(11, 0) = "debe07"
    CAMPOS(12, 0) = "debe08"
    CAMPOS(13, 0) = "debe09"
    CAMPOS(14, 0) = "debe10"
    CAMPOS(15, 0) = "debe11"
    CAMPOS(16, 0) = "debe12"
    CAMPOS(17, 0) = "haber01"
    CAMPOS(18, 0) = "haber02"
    CAMPOS(19, 0) = "haber03"
    CAMPOS(20, 0) = "haber04"
    CAMPOS(21, 0) = "haber05"
    CAMPOS(22, 0) = "haber06"
    CAMPOS(23, 0) = "haber07"
    CAMPOS(24, 0) = "haber08"
    CAMPOS(25, 0) = "haber09"
    CAMPOS(26, 0) = "haber10"
    CAMPOS(27, 0) = "haber11"
    CAMPOS(28, 0) = "haber12"
    
    CAMPOS(29, 0) = ""
    CAMPOS(0, 1) = dato1.text
    CAMPOS(1, 1) = dato2.text + dv.Caption
    CAMPOS(2, 1) = año

    For k = 3 To 28
    CAMPOS(k, 1) = "0"
    Next k
    CAMPOS(0, 2) = "saldosctacte"
    op = 2
    
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    

End Sub

Sub ELIMINAR()
    CAMPOS(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut=" + "'" + dato2.text + dv.Caption + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 4
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)

    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub



Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then modifica
If command = "elimina" Then elimina

If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "movimientos" Then movimientos



End Sub
Sub elimina()
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
If cierrect = "S" Then cierrect = "": Unload Me
    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dv.Caption = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
End Sub



Sub DATOSSALDOS()
Dim debe As Double
Dim haber As Double

LEERSALDOS
SUMADOR = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
SALDOS.TextMatrix(1, 1) = Format(sqlconta.response(3, 3), "###,###,##0")
SALDOS.TextMatrix(1, 2) = Format(sqlconta.response(4, 3), "###,###,##0")
SALDOS.TextMatrix(1, 3) = Format(SUMADOR, "###,###,##0")
debe = 0
haber = 0

For k = 5 To 16

SALDOS.TextMatrix(k - 3, 1) = Format(sqlconta.response(k, 3), "###,###,##0")
SALDOS.TextMatrix(k - 3, 2) = Format(sqlconta.response(k + 12, 3), "###,###,##0")
SUMADOR = SUMADOR + Val(sqlconta.response(k, 3)) - Val(sqlconta.response(k + 12, 3))
SALDOS.TextMatrix(k - 3, 3) = Format(SUMADOR, "###,###,##0")
debe = debe + Val(sqlconta.response(k, 3))
haber = haber + Val(sqlconta.response(k + 12, 3))

Next k
saldoglobal = debe + haber

End Sub
Sub grillasaldos()
SALDOS.Cols = 4
SALDOS.Rows = 14
SALDOS.ColWidth(0) = 120 * 12
SALDOS.ColWidth(1) = 120 * 8
SALDOS.ColWidth(2) = 120 * 8
SALDOS.ColWidth(3) = 120 * 8
SALDOS.TextMatrix(0, 0) = "MESES   "
SALDOS.TextMatrix(0, 1) = "DEBE    "
SALDOS.TextMatrix(0, 2) = "HABER   "
SALDOS.TextMatrix(0, 3) = "SALDO   "
SALDOS.TextMatrix(1, 0) = "AÑO ANTERIOR"
SALDOS.TextMatrix(2, 0) = "ENERO"
SALDOS.TextMatrix(3, 0) = "FEBRERO"
SALDOS.TextMatrix(4, 0) = "MARZO"
SALDOS.TextMatrix(5, 0) = "ABRIL"
SALDOS.TextMatrix(6, 0) = "MAYO"
SALDOS.TextMatrix(7, 0) = "JUNIO"
SALDOS.TextMatrix(8, 0) = "JULIO"
SALDOS.TextMatrix(9, 0) = "AGOSTO"
SALDOS.TextMatrix(10, 0) = "SEPTIEMBRE"
SALDOS.TextMatrix(11, 0) = "OCTUBRE"
SALDOS.TextMatrix(12, 0) = "NOVIEMBRE "
SALDOS.TextMatrix(13, 0) = "DICIEMBRE "
For k = 1 To 13
SALDOS.TextMatrix(k, 1) = "0"
SALDOS.TextMatrix(k, 2) = "0"
SALDOS.TextMatrix(k, 3) = "0"
Next k
End Sub

Sub LEERSALDOS()
    CAMPOS(0, 0) = "tipo"
    CAMPOS(1, 0) = "rut"
    CAMPOS(2, 0) = "año"
    CAMPOS(3, 0) = "debeanterior"
    CAMPOS(4, 0) = "haberanterior"
    CAMPOS(5, 0) = "debe01"
    CAMPOS(6, 0) = "debe02"
    CAMPOS(7, 0) = "debe03"
    CAMPOS(8, 0) = "debe04"
    CAMPOS(9, 0) = "debe05"
    CAMPOS(10, 0) = "debe06"
    CAMPOS(11, 0) = "debe07"
    CAMPOS(12, 0) = "debe08"
    CAMPOS(13, 0) = "debe09"
    CAMPOS(14, 0) = "debe10"
    CAMPOS(15, 0) = "debe11"
    CAMPOS(16, 0) = "debe12"
    CAMPOS(17, 0) = "haber01"
    CAMPOS(18, 0) = "haber02"
    CAMPOS(19, 0) = "haber03"
    CAMPOS(20, 0) = "haber04"
    CAMPOS(21, 0) = "haber05"
    CAMPOS(22, 0) = "haber06"
    CAMPOS(23, 0) = "haber07"
    CAMPOS(24, 0) = "haber08"
    CAMPOS(25, 0) = "haber09"
    CAMPOS(26, 0) = "HABER10"
    CAMPOS(27, 0) = "HABER11"
    CAMPOS(28, 0) = "HABER12"
    CAMPOS(29, 0) = ""
    condicion = "tipo=" + "'" + dato1.text + "' and rut='" + dato2.text + dv.Caption + "' and año='" + Mid(fechasistema, 7, 4) + "'"
    CAMPOS(0, 2) = "saldosctacte"
    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    
    
    
   Rem  If sqlconta.status = 4 Then Stop
grillasaldos
End Sub




Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus

End Sub

Sub ayudactacte(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & dato1.text & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    
    Call cargaAyudaT(servidor, basebus, usuario, password, "cuentascorrientes", pivote, CAMPOS, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then dato2.SetFocus: GoTo no
    dato4.Enabled = True
    dato2.text = Mid(pivote.text, 1, 9)
    dv.Caption = Mid(pivote.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub

Sub LEETIPOCTACTE()
    CAMPOS(0, 0) = "codigo"
    CAMPOS(1, 0) = "nombre"
    CAMPOS(2, 0) = ""
    CAMPOS(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    
    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    
    Call sqlconta.sqlconta(op, condicion)

   If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
   GLOSACTACTE.Caption = sqlconta.response(1, 3)

no:
End Sub

Sub movimientos()
Rem cartola = "mayor:" + dato1.text + dato2.text + dato3.text
informa04.ctdato1.text = dato1.text
informa04.ctdato2.text = dato2.text
informa04.dv.text = dv.Caption
informa04.nombrectacte = GLOSACTACTE.Caption



informa04.ctnombre = dato4.text
informa04.sbtab1.Tab = 1
informa04.ctindi = True


informa04.Show

End Sub

Sub CARGAGRILLArelacion()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CUENTA"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "SALDO ANTE."
    formatogrilla2(1, 4) = "DEBE"
    formatogrilla2(1, 5) = "HABER"
    formatogrilla2(1, 6) = "SALDO ACTUAL"
    formatogrilla2(1, 7) = "EMPRESA"
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "6"
    formatogrilla2(2, 2) = "20"
    formatogrilla2(2, 3) = "10"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "17"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "C"
    formatogrilla2(3, 2) = "C"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "S"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 3) = " ###,###,##0"
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " ###,###,##0"
    formatogrilla2(4, 6) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 8
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
        
        
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
 
    End Sub

