VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form activo03 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LISTADO ACTIVO FIJO"
   ClientHeight    =   10095
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   17310
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   673
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1154
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "ASIENTO"
      Height          =   255
      Left            =   15480
      TabIndex        =   18
      Top             =   9840
      Width           =   1815
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1380
      Left            =   120
      TabIndex        =   2
      Top             =   45
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   2434
      BackColor       =   16744576
      Caption         =   "DATOS "
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
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
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EXCEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   15000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1000
         Width           =   1965
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   15000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1965
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PROCESAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   15000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1980
      End
      Begin XPFrame.FrameXp frmrut 
         Height          =   675
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   1191
         BackColor       =   16761024
         Caption         =   "FILTRAR POR FAMILIA"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         ColorBarraArriba=   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox dato3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   7
            Tag             =   "rut"
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FAMILIA"
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
            Height          =   255
            Left            =   225
            TabIndex        =   9
            Top             =   270
            Width           =   1530
         End
         Begin VB.Label lblnombreFAMILIA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Top             =   270
            Width           =   4455
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   675
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   "LOCAL"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BordeColor      =   -2147483635
         ColorBarraArriba=   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox ComboLOCAL 
            Height          =   315
            Left            =   90
            TabIndex        =   5
            Top             =   270
            Width           =   4395
         End
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   675
         Left            =   8280
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   "AÑO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BordeColor      =   -2147483635
         ColorBarraArriba=   4194304
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
         Begin VB.ComboBox comboaño 
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   270
            Width           =   1635
         End
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VENDIDO"
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
         Height          =   255
         Left            =   11400
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "DOBLE CLICK SOBRE GRILLA PARA VER DETALLE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3975
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   8400
      Left            =   90
      TabIndex        =   3
      Top             =   1440
      Width           =   17145
      _ExtentX        =   30242
      _ExtentY        =   14817
      BackColor       =   16761024
      Caption         =   "Listado de Acivos Fijos "
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin FlexCell.Grid Grid1 
         Height          =   8025
         Left            =   10
         TabIndex        =   19
         Top             =   360
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   14155
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.PictureBox CmdFavoritos 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   15240
         ScaleHeight     =   195
         ScaleWidth      =   2715
         TabIndex        =   12
         Top             =   10920
         Width           =   2775
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
      ScaleWidth      =   17280
      TabIndex        =   1
      Top             =   10095
      Width           =   17310
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DOBLE CLIC EN LA GRILLA PARA VISUALIZAR ACTIVO Y FACTURA"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   9840
      Width           =   15255
   End
End
Attribute VB_Name = "activo03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public saldoglobal As Double
Private localfiltro As String
Private GeneraAsiento As Boolean
Private MODIFI As Integer
Dim lin As Double
'Private Sub codigo_Click()
'    Call dato1_KeyDown(vbKeyF2, 0)
'End Sub
 Private Sub imprimir()
If Grid1.Rows > 1 Then
Call Titulos("LISTADO DE ACTIVOS FIJOS")
Grid1.PageSetup.Orientation = cellLandscape
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


 For k = 1 To Grid1.PageSetup.PaperSizes.Count
            If UCase(Grid1.PageSetup.PaperSizes.item(k).PaperName) = "LEGAL" Then 'Or UCase(Grid1.PageSetup.PaperSizes.Item(k).PaperName) = "LETTER" Then
            
                Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.item(k).Kind
                Exit For
            End If
        Next k


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
    

    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub


Private Sub CmdFavoritos_Click()
    Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

'Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then Call ayudactacte(dato2)
'    Call flechas(dato1, dato4, KeyCode)
'End Sub
 

Private Sub Command1_Click()
imprimir

End Sub

Private Sub COMMAND2_Click()
Call LEERactivofijos
End Sub

Private Sub Command3_Click()
Call Grid1.ExportToExcel("", True, True)
End Sub

Private Sub Command4_Click()
GeneraAsiento = True
Call LEERactivofijos
GeneraAsiento = False
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudafamilia(dato3)

End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(dato3)
lblnombreFAMILIA.Caption = LEERNOMBREFAMILIA(dato3.text)
If lblnombreFAMILIA.Caption = "" Then
    dato3.SetFocus
Else
    LEERactivofijos
End If

End If

End Sub

Private Sub Form_Load()
'Me.Width = Screen.Width - 1000
'Me.Refresh
'FrameXp1.Width = Me.Width - 1000
'Me.Refresh
'FrameXp2.Width = FrameXp1.Width
'Me.Refresh
'Grid1.Width = FrameXp2.Width
'Call CENTRAR(Me)

    Call Conectar_BD
    sc = 0
  
Call CARGAPERMISO(Me.Name)
 
 Call CargarFormatoGrilla
LEErlocales
LeerAños
End Sub
Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion.g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub

Sub LeerAños()
'    Dim resultados As rdoResultset
'    Dim csql As New rdoQuery
'
'        Set csql.ActiveConnection = contadb
'        csql.sql = "SELECT año FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo"
'        csql.sql = csql.sql & " GROUP BY año"
'        csql.Execute
'        comboaño.Clear
'        If csql.RowsAffected > 0 Then
'            Set resultados = csql.OpenResultset
'            While Not resultados.EOF
'                comboaño.AddItem (resultados(0))
'                resultados.MoveNext
'            Wend
'            resultados.Close
'            Set resultados = Nothing
'        comboaño.text = comboaño.List(0)
'        End If
          COMBOAÑO.AddItem Format(fechasistema, "yyyy")
        COMBOAÑO.text = COMBOAÑO.List(0)
        FrameXp2.Caption = " Listado de Acivos Fijos " & COMBOAÑO
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub


Sub ayudafamilia(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12n", "40s")
    cfijo = "no"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda Familias"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestro_familias_nuevo", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
     
    caja.text = pivote.text
    caja.Enabled = True
    caja.SetFocus

no:

End Sub

Sub CARGAGRILLADETALLE()
    Dim formatogrilla2(10, 40)
 
    
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "CRCC"
    formatogrilla2(1, 4) = "FECHA P." & vbCrLf & "SERVICIO"
    formatogrilla2(1, 5) = "VALOR" & vbCrLf & "LIBRO"
    formatogrilla2(1, 6) = "DEPRECIACION" & vbCrLf & "ACUMULADA"
    formatogrilla2(1, 7) = "VALOR" & vbCrLf & "NETO"
    formatogrilla2(1, 8) = "FACTOR" & vbCrLf & "CORREC."
    formatogrilla2(1, 9) = "CORRECION" & vbCrLf & "MONETARIA"
    formatogrilla2(1, 10) = "VALOR" & vbCrLf & "ACTUAL."
    formatogrilla2(1, 11) = "VIDA" & vbCrLf & "UTIL"
    formatogrilla2(1, 12) = "VIDA" & vbCrLf & "USADA"
    formatogrilla2(1, 13) = "DEPRECIACION" & vbCrLf & "DEL EJERCICIO"
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "6"
    formatogrilla2(2, 2) = "15"
    formatogrilla2(2, 3) = "17"
    formatogrilla2(2, 4) = "7"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "8"
    formatogrilla2(2, 8) = "8"
    formatogrilla2(2, 9) = "8"
    formatogrilla2(2, 10) = "8"
    formatogrilla2(2, 11) = "8"
    formatogrilla2(2, 12) = "8"
    formatogrilla2(2, 13) = "15"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
 
    
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "D"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    formatogrilla2(3, 8) = "N"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    formatogrilla2(3, 12) = "N"
    formatogrilla2(3, 13) = "N"
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 5) = "$ ###,###,###,##0"
    formatogrilla2(4, 6) = "$ ###,###,##0"
    formatogrilla2(4, 7) = "$ ###,###,##0"
    formatogrilla2(4, 8) = "% ###,###,##0.00"
    formatogrilla2(4, 9) = "$ ###,###,##0.00"
    formatogrilla2(4, 10) = "$ ###,###,##0"
    formatogrilla2(4, 13) = "$ ###,###,##0.00"
      
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    formatogrilla2(5, 10) = "TRUE"
    formatogrilla2(5, 11) = "TRUE"
    formatogrilla2(5, 12) = "TRUE"
    formatogrilla2(5, 13) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 14
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
'    Grid1.BackColorFixed = RGB(90, 158, 214)
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
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
  Grid1.DefaultFont.Size = 6.5
   Grid1.RowHeight(0) = 30
  Grid1.Column(0).Width = 0
  Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
  Grid1.Range(0, 1, 0, Grid1.Cols - 1).WrapText = True
  
 
    End Sub

 
Sub CargarFormatoGrilla()
    Dim formatogrilla2(10, 40)
 
    
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "CRCC"
    formatogrilla2(1, 4) = "FECHA P." & vbCrLf & "SERVICIO"
    formatogrilla2(1, 5) = "VALOR" & vbCrLf & "LIBRO"
    formatogrilla2(1, 6) = "DEPRECIACION" & vbCrLf & "ACUMULADA"
    formatogrilla2(1, 7) = "CORRECCION " & vbCrLf & "ACUMULADA"
    formatogrilla2(1, 8) = "CREDITO" & vbCrLf & "4 %"
    formatogrilla2(1, 9) = "VALOR" & vbCrLf & "INICIAL"
    formatogrilla2(1, 10) = "DEPRECIACION" & vbCrLf & " DEL EJERCICIO"
    formatogrilla2(1, 11) = "CORRECCION " & vbCrLf & "DEL EJERCICIO"
    formatogrilla2(1, 12) = "VALOR" & vbCrLf & "FINAL"
    formatogrilla2(1, 13) = "FACTOR" & vbCrLf & "CORREC."
    formatogrilla2(1, 14) = "VIDA" & vbCrLf & "UTIL"
    formatogrilla2(1, 15) = "VIDA" & vbCrLf & "USADA"
    formatogrilla2(1, 16) = "VIDA" & vbCrLf & "DEL EJERCICIO"
    formatogrilla2(1, 17) = "NUEVA VIDA" & vbCrLf & "UTIL"
    
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "6"
    formatogrilla2(2, 2) = "20"
    formatogrilla2(2, 3) = "0"
    formatogrilla2(2, 4) = "7"
    formatogrilla2(2, 5) = "15"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "0"
    formatogrilla2(2, 8) = "7"
    formatogrilla2(2, 9) = "10"
    formatogrilla2(2, 10) = "10"
    formatogrilla2(2, 11) = "10"
    formatogrilla2(2, 12) = "10"
    formatogrilla2(2, 13) = "7"
    formatogrilla2(2, 14) = "5"
    formatogrilla2(2, 15) = "5"
    formatogrilla2(2, 16) = "8"
    formatogrilla2(2, 17) = "8"
    formatogrilla2(2, 18) = "5"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
 
    
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "D"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    formatogrilla2(3, 8) = "N"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    formatogrilla2(3, 12) = "N"
    formatogrilla2(3, 13) = "N"
    formatogrilla2(3, 14) = "N"
    formatogrilla2(3, 15) = "N"
    formatogrilla2(3, 16) = "N"
    formatogrilla2(3, 17) = "N"
    formatogrilla2(3, 18) = "N"
    formatogrilla2(3, 19) = "N"
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 5) = "###,###,###,##0"
    formatogrilla2(4, 6) = "###,###,##0"
    formatogrilla2(4, 7) = "###,###,##0"
    formatogrilla2(4, 8) = "###,###,##0"
    formatogrilla2(4, 9) = "###,###,##0"
    formatogrilla2(4, 10) = "###,###,##0"
    formatogrilla2(4, 11) = "###,###,##0"
    formatogrilla2(4, 12) = "###,###,##0"
    formatogrilla2(4, 13) = "###,###,##0.00"
    formatogrilla2(4, 14) = "###,###,##0"
    formatogrilla2(4, 15) = "###,###,##0"
    formatogrilla2(4, 16) = "###,###,##0"
    formatogrilla2(4, 17) = "###,###,##0"
    formatogrilla2(4, 18) = "###,###,##0"
    formatogrilla2(4, 19) = "###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    formatogrilla2(5, 10) = "TRUE"
    formatogrilla2(5, 11) = "TRUE"
    formatogrilla2(5, 12) = "TRUE"
    formatogrilla2(5, 13) = "TRUE"
    formatogrilla2(5, 14) = "TRUE"
    formatogrilla2(5, 15) = "TRUE"
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 18
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
'    Grid1.BackColorFixed = RGB(90, 158, 214)
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 8
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then
            Grid1.Column(k).Alignment = cellRightCenter
            Grid1.Column(k).Mask = cellNumeric
        End If
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
    Next k
  Grid1.DefaultFont.Size = 6.5
   Grid1.RowHeight(0) = 40
  Grid1.Column(0).Width = 0
  Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
  Grid1.Range(0, 1, 0, Grid1.Cols - 1).WrapText = True
  
 
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


Sub LEERactivofijos()
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo As String
Dim filtro As String
Dim FILTRO2 As String
Dim totales(10) As Double
Dim totales2(10) As Double
Dim cuentapublicidad As String
Dim factor As Double
Dim CORRECCION As Double
Dim depreciacion As Double
Dim vidaanterior As Double
Dim valorlibro As Double
Dim depreciacionmes As Double
Dim vidaejercicio As Double
Dim libro As Double
Dim glosa As String
Dim numero As String
Dim td As String
Dim cuenta As String
Dim monto As Double
Dim DH As String
If GeneraAsiento = True Then
    td = "AF"
    numero = LEERULTIMOFOLIO(td)
    LINEA = 0
End If


Call CargarFormatoGrilla
 For k = 1 To Grid1.Cols - 1
    Grid1.Column(k).Locked = False
 Next k
    Set csql.ActiveConnection = contadb
    csql.sql = "select codigo,nombre,crcc,fechapuestaenmarcha,valorcompra+correcionmonetaria+IFNULL((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0)"
    csql.sql = csql.sql & " ,depreciacion+ifnull((SELECT SUM(depreciacion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
    csql.sql = csql.sql & " WHERE año <c.año AND codigo=c.codigo),0) AS depreciacion "
    csql.sql = csql.sql & " ,correcionmonetaria+ifnull((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0) AS correccionmonetaria"
    csql.sql = csql.sql & ",valorcompra+correcionmonetaria+IFNULL((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0)-(depreciacion+IFNULL((SELECT SUM(depreciacion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0)) -valorcredito,vidautil"
    csql.sql = csql.sql & ",ifnull((SELECT SUM(vida_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0) AS  vidausada,'0',familia"
    
    csql.sql = csql.sql & ", valorcredito,fechaventa "
    csql.sql = csql.sql & " from activo_fijo_nuevo  as c where año='" & COMBOAÑO & "' " 'AND (codigo='00001016' OR codigo='00000749' OR codigo='00000723') " 'and codigo='00000767' "
    If dato3.text <> "" Then
        csql.sql = csql.sql & " and familia='" & dato3.text & "'    "
    End If
    
    
    csql.sql = csql.sql & "order by familia,fechapuestaenmarcha,codigo  "
    
    csql.Execute
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        filtro = resultados(11)
        FILTRO2 = filtro
        Grid1.AutoRedraw = False
        
        While Not resultados.EOF
        libro = 0
   '     If Val(resultados("codigo")) = 227 Then Stop
            If filtro <> FILTRO2 Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                Grid1.Column(3).Locked = False
                Grid1.Column(4).Locked = False
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 4).FontBold = True
                
                Grid1.Cell(Grid1.Rows - 1, 4).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 5).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 8).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 10).Font.Bold = True
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTAL FAMILIA " & filtro & "-" & LEERNOMBREFAMILIA(filtro)
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Alignment = cellLeftGeneral
                Grid1.Cell(Grid1.Rows - 1, 5).text = totales(1)
                Grid1.Cell(Grid1.Rows - 1, 6).text = totales(2)
                Grid1.Cell(Grid1.Rows - 1, 7).text = totales(3)
                Grid1.Cell(Grid1.Rows - 1, 8).text = totales(4)
                Grid1.Cell(Grid1.Rows - 1, 9).text = totales(5)
                Grid1.Cell(Grid1.Rows - 1, 10).text = totales(6)
                Grid1.Cell(Grid1.Rows - 1, 11).text = totales(7)
                Grid1.Cell(Grid1.Rows - 1, 17).text = totales(8)
                
                If GeneraAsiento = True Then
                    Rem DEPRECIACION
                    LINEA = LINEA + 1
                    glosa = LEERNOMBREFAMILIA(filtro)
                    cuenta = leerdatos(conta, clientesistema & "conta" & ".maestro_familias_nuevo", "depreciacionacumulada", "codigo='" & filtro & "'")
                    monto = Grid1.Cell(Grid1.Rows - 1, 10).text
                    DH = "H"
                    Call grabarcomprobante_lineas(td, numero, LINEA, Format(fechasistema, "YYYY-MM-DD"), cuenta, "", "", "", "DEPRECIACION " & glosa, td, numero, Format(fechasistema, "YYYY-MM-DD"), Format(fechasistema, "YYYY-MM-DD"), monto, DH, USUARIOSISTEMA, Format(fechasistema, "MM"), Format(fechasistema, "YYYY"), Format(fechasistema, "YYYY-MM-DD"), Time, "")
                    
                    LINEA = LINEA + 1
                    cuenta = leerdatos(conta, clientesistema & "conta" & ".maestro_familias_nuevo", "cuentadepreciacion", "codigo='" & filtro & "'")
                    monto = Grid1.Cell(Grid1.Rows - 1, 10).text
                    DH = "D"
                    Call grabarcomprobante_lineas(td, numero, LINEA, Format(fechasistema, "YYYY-MM-DD"), cuenta, "", "", "000054", "DEPRECIACION " & glosa, "", "", "", "", monto, DH, USUARIOSISTEMA, Format(fechasistema, "MM"), Format(fechasistema, "YYYY"), Format(fechasistema, "YYYY-MM-DD"), Time, "")
                    
                    Rem CORRECION
                    LINEA = LINEA + 1
                    glosa = LEERNOMBREFAMILIA(filtro)
                    cuenta = leerdatos(conta, clientesistema & "conta" & ".maestro_familias_nuevo", "cuentamayor", "codigo='" & filtro & "'")
                    monto = Grid1.Cell(Grid1.Rows - 1, 11).text
                    DH = "D"
                    Call grabarcomprobante_lineas(td, numero, LINEA, Format(fechasistema, "YYYY-MM-DD"), cuenta, "", "", "", "CORRECION " & glosa, td, numero, Format(fechasistema, "YYYY-MM-DD"), Format(fechasistema, "YYYY-MM-DD"), monto, DH, USUARIOSISTEMA, Format(fechasistema, "MM"), Format(fechasistema, "YYYY"), Format(fechasistema, "YYYY-MM-DD"), Time, "")
                    
                    LINEA = LINEA + 1
                    cuenta = leerdatos(conta, clientesistema & "conta" & ".maestro_familias_nuevo", "cuentacorreccion", "codigo='" & filtro & "'")
                    monto = Grid1.Cell(Grid1.Rows - 1, 11).text
                    DH = "H"
                    Call grabarcomprobante_lineas(td, numero, LINEA, Format(fechasistema, "YYYY-MM-DD"), cuenta, "", "", "000054", "CORRECION " & glosa, "", "", "", "", monto, DH, USUARIOSISTEMA, Format(fechasistema, "MM"), Format(fechasistema, "YYYY"), Format(fechasistema, "YYYY-MM-DD"), Time, "")
                    
                    
                    
                End If
                
                For k = 0 To 10
                    totales(k) = 0
                Next k
                filtro = FILTRO2
                libro = 0
            End If
            
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 0).text = "1"
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
       '     Grid1.Cell(Grid1.Rows - 1, 3).text = leerNOMBREcrcc(resultados(2))
            Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultados(3), "dd-mm-yyyy")
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
             Grid1.Cell(Grid1.Rows - 1, 7).text = 0 'resultados(6)
            
            libro = resultados(7)
            CORRECCION = 0
            depreciacion = 0
            vidaejercicio = 0
            If libro < 0 Then libro = 1
            Grid1.Cell(Grid1.Rows - 1, 9).text = libro
            
If libro > 1 Then
          If Format(resultados(3), "yyyy") < COMBOAÑO Then
                Grid1.Cell(Grid1.Rows - 1, 13).text = leeripc("00", COMBOAÑO)
             Else
                Grid1.Cell(Grid1.Rows - 1, 13).text = leeripc(Format(resultados(3), "mm"), Format(resultados(3), "yyyy"))
            End If
            
            factor = Grid1.Cell(Grid1.Rows - 1, 13).text / 100
            CORRECCION = Round(libro * factor)
             
             If Format(resultados(3), "yyyy") < COMBOAÑO Then
                vidaejercicio = 12
             Else
                If IsDate(resultados("fechaventa")) = False Then
                    vidaejercicio = Round(DateDiff("m", resultados(3), fechasistema)) + 1
                Else
                    vidaejercicio = Round(DateDiff("m", resultados(3), resultados("fechaventa"))) + 1
                End If
             End If
              If vidaejercicio > 12 Then
                vidaejercicio = 12
              End If
              
              
             If resultados(8) - resultados(9) < vidaejercicio Then
                vidaejercicio = resultados(8) - resultados(9)
             End If
            Grid1.Cell(Grid1.Rows - 1, 16).text = vidaejercicio
                If IsDate(resultados("fechaventa")) = False Then
                    Grid1.Cell(Grid1.Rows - 1, 17).text = resultados(8) - resultados(9) - vidaejercicio
                Else
                    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = vbGreen
                    Grid1.Cell(Grid1.Rows - 1, 17).text = 0
                End If
            vidaanterior = resultados(8) - resultados(9)
            valorlibro = (libro)
            If vidaanterior = 0 Then vidaanterior = 1
            depreciacionmes = (valorlibro + CORRECCION) / vidaanterior
             
            depreciacion = Round(depreciacionmes * vidaejercicio)
            
            If TieneDepreciacion(resultados("familia")) = False Then
                depreciacion = 0
            End If
            
            If TieneCorreccion(resultados("familia")) = False Then
                CORRECCION = 0
            End If
            
End If
            
            If libro = 0 Then libro = 1
            Grid1.Cell(Grid1.Rows - 1, 11).text = CORRECCION
            Grid1.Cell(Grid1.Rows - 1, 10).text = depreciacion
            Grid1.Cell(Grid1.Rows - 1, 12).text = libro - depreciacion + CORRECCION
            If Grid1.Cell(Grid1.Rows - 1, 12).text = 0 Then Grid1.Cell(Grid1.Rows - 1, 12).text = 1
            Grid1.Cell(Grid1.Rows - 1, 14).text = resultados(8)
            Grid1.Cell(Grid1.Rows - 1, 15).text = resultados(9)
            
            If Grid1.Cell(Grid1.Rows - 1, 17).text = "" Then Grid1.Cell(Grid1.Rows - 1, 17).text = 0
            
    
 
             
 '           Grid1.Cell(Grid1.Rows - 1, 14).text = 0
'            If Format(resultados(3), "yyyy") = Format(fechasistema, "yyyy") Then
                 Grid1.Cell(Grid1.Rows - 1, 8).text = Round(resultados("valorcredito"))
'            End If
            
'            If ExistemMovimientos(resultados(0), comboaño) = False Then
                Call ActualizaValores(resultados(0), COMBOAÑO, depreciacion, CORRECCION, vidaejercicio)
'            End If
        
            totales(1) = totales(1) + Grid1.Cell(Grid1.Rows - 1, 5).text
            totales(2) = totales(2) + Grid1.Cell(Grid1.Rows - 1, 6).text
            totales(3) = totales(3) + Grid1.Cell(Grid1.Rows - 1, 7).text
            totales(4) = totales(4) + Grid1.Cell(Grid1.Rows - 1, 8).text
            totales(5) = totales(5) + resultados(7)
            totales(6) = totales(6) + Grid1.Cell(Grid1.Rows - 1, 10).text
            totales(7) = totales(7) + Grid1.Cell(Grid1.Rows - 1, 11).text
            totales(8) = totales(8) + Grid1.Cell(Grid1.Rows - 1, 17).text
            
            
            totales2(1) = totales2(1) + Grid1.Cell(Grid1.Rows - 1, 5).text
            totales2(2) = totales2(2) + Grid1.Cell(Grid1.Rows - 1, 6).text
            totales2(3) = totales2(3) + Grid1.Cell(Grid1.Rows - 1, 7).text
            totales2(4) = totales2(4) + Grid1.Cell(Grid1.Rows - 1, 8).text
            totales2(5) = totales2(5) + resultados(7)
            totales2(6) = totales2(6) + Grid1.Cell(Grid1.Rows - 1, 10).text
            totales2(7) = totales2(7) + Grid1.Cell(Grid1.Rows - 1, 11).text
            totales2(8) = totales2(8) + Grid1.Cell(Grid1.Rows - 1, 17).text
            
        
            resultados.MoveNext
            If Not resultados.EOF Then
                FILTRO2 = resultados("familia")
            End If
        Wend
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                Grid1.Column(3).Locked = False
                Grid1.Column(4).Locked = False
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 4).FontBold = True
                
                Grid1.Cell(Grid1.Rows - 1, 4).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 5).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 8).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 10).Font.Bold = True
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTAL FAMILIA " & filtro & "-" & LEERNOMBREFAMILIA(filtro)
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Alignment = cellLeftGeneral
                Grid1.Cell(Grid1.Rows - 1, 5).text = totales(1)
                Grid1.Cell(Grid1.Rows - 1, 6).text = totales(2)
                Grid1.Cell(Grid1.Rows - 1, 7).text = totales(3)
                Grid1.Cell(Grid1.Rows - 1, 8).text = totales(4)
                Grid1.Cell(Grid1.Rows - 1, 9).text = totales(5)
                Grid1.Cell(Grid1.Rows - 1, 10).text = totales(6)
                Grid1.Cell(Grid1.Rows - 1, 11).text = totales(7)
                Grid1.Cell(Grid1.Rows - 1, 17).text = totales(8)
                
                If GeneraAsiento = True Then
                    Rem DEPRECIACION
                    LINEA = LINEA + 1
                    glosa = LEERNOMBREFAMILIA(filtro)
                    cuenta = leerdatos(conta, clientesistema & "conta" & ".maestro_familias_nuevo", "cuentamayor", "codigo='" & filtro & "'")
                    monto = Grid1.Cell(Grid1.Rows - 1, 10).text
                    DH = "H"
                    Call grabarcomprobante_lineas(td, numero, LINEA, Format(fechasistema, "YYYY-MM-DD"), cuenta, "", "", "", "DEPRECIACION " & glosa, td, numero, Format(fechasistema, "YYYY-MM-DD"), Format(fechasistema, "YYYY-MM-DD"), monto, DH, USUARIOSISTEMA, Format(fechasistema, "MM"), Format(fechasistema, "YYYY"), Format(fechasistema, "YYYY-MM-DD"), Time, "")
                    
                    LINEA = LINEA + 1
                    cuenta = leerdatos(conta, clientesistema & "conta" & ".maestro_familias_nuevo", "cuentadepreciacion", "codigo='" & filtro & "'")
                    monto = Grid1.Cell(Grid1.Rows - 1, 10).text
                    DH = "D"
                    Call grabarcomprobante_lineas(td, numero, LINEA, Format(fechasistema, "YYYY-MM-DD"), cuenta, "", "", "000054", "DEPRECIACION " & glosa, "", "", "", "", monto, DH, USUARIOSISTEMA, Format(fechasistema, "MM"), Format(fechasistema, "YYYY"), Format(fechasistema, "YYYY-MM-DD"), Time, "")
                    
                    
                    Rem CORRECION
                    LINEA = LINEA + 1
                    glosa = LEERNOMBREFAMILIA(filtro)
                    cuenta = leerdatos(conta, clientesistema & "conta" & ".maestro_familias_nuevo", "cuentamayor", "codigo='" & filtro & "'")
                    monto = Grid1.Cell(Grid1.Rows - 1, 11).text
                    DH = "H"
                    Call grabarcomprobante_lineas(td, numero, LINEA, Format(fechasistema, "YYYY-MM-DD"), cuenta, "", "", "", "CORRECION " & glosa, td, numero, Format(fechasistema, "YYYY-MM-DD"), Format(fechasistema, "YYYY-MM-DD"), monto, DH, USUARIOSISTEMA, Format(fechasistema, "MM"), Format(fechasistema, "YYYY"), Format(fechasistema, "YYYY-MM-DD"), Time, "")
                    
                    LINEA = LINEA + 1
                    cuenta = leerdatos(conta, clientesistema & "conta" & ".maestro_familias_nuevo", "cuentacorreccion", "codigo='" & filtro & "'")
                    monto = Grid1.Cell(Grid1.Rows - 1, 11).text
                    DH = "D"
                    Call grabarcomprobante_lineas(td, numero, LINEA, Format(fechasistema, "YYYY-MM-DD"), cuenta, "", "", "000054", "CORRECION " & glosa, "", "", "", "", monto, DH, USUARIOSISTEMA, Format(fechasistema, "MM"), Format(fechasistema, "YYYY"), Format(fechasistema, "YYYY-MM-DD"), Time, "")
                    
                    
                End If
                
                
                 Grid1.Rows = Grid1.Rows + 2
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                Grid1.Column(3).Locked = False
                Grid1.Column(4).Locked = False
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 4).FontBold = True
                
                Grid1.Cell(Grid1.Rows - 1, 4).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 5).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 8).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 10).Font.Bold = True
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTAL GENERAL"
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Alignment = cellLeftGeneral
                Grid1.Cell(Grid1.Rows - 1, 5).text = totales2(1)
                Grid1.Cell(Grid1.Rows - 1, 6).text = totales2(2)
                Grid1.Cell(Grid1.Rows - 1, 7).text = totales2(3)
                Grid1.Cell(Grid1.Rows - 1, 8).text = totales2(4)
                Grid1.Cell(Grid1.Rows - 1, 9).text = totales2(5)
                Grid1.Cell(Grid1.Rows - 1, 10).text = totales2(6)
                Grid1.Cell(Grid1.Rows - 1, 11).text = totales2(7)
                Grid1.Cell(Grid1.Rows - 1, 17).text = totales2(8)
                
    End If
    Grid1.AutoRedraw = True
    Grid1.Refresh
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
 For k = 1 To Grid1.Cols - 1
    Grid1.Column(k).Locked = True
 Next k
End Sub
Function LEERNOMBREFAMILIA(codigo) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "maestro_familias_nuevo"
    condicion = "codigo= '" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LEERNOMBREFAMILIA = sqlconta.response(0, 3)
    Else
        LEERNOMBREFAMILIA = ""
    End If
    

End Function
  Function leeripc(MES, año) As Double
Dim csql As New rdoQuery
Dim resultados  As rdoResultset

Set csql.ActiveConnection = conta

csql.sql = "select porcentaje from ipc_nuevo where mes='" & MES & "' and año='" & año & "' "

csql.Execute

 leeripc = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leeripc = resultados(0)
End If

End Function

Private Sub Grid1_DblClick()
    If Grid1.Rows > 1 Then
        If Grid1.Cell(Grid1.ActiveCell.row, 0).text = "1" Then
            Unload activo02
            Load activo02
            activo02.Tag = Grid1.Cell(Grid1.ActiveCell.row, 1).text
            activo02.Show
          
        End If
    End If
    
End Sub

Private Sub Option1_Click()
    dato3.text = ""
    lblnombreFAMILIA.Caption = ""
    frmrut.Enabled = False
    LEERactivofijos
End Sub

Private Sub Option2_Click()
frmrut.Enabled = True
dato3.SetFocus

End Sub




Sub ActualizaValores(codigo, AÑOACTUAL, depreciacion_ejercicio, correccion_ejercicio, vida_ejercicio)
    Dim campos(10, 10) As String
    Dim op As Double
    Dim condicion As String
    campos(0, 0) = "depreciacion_ejercicio"
    campos(1, 0) = "correccion_ejercicio"
    campos(2, 0) = "vida_ejercicio"
    campos(3, 0) = ""
    
    campos(0, 1) = depreciacion_ejercicio
    campos(1, 1) = correccion_ejercicio
    campos(2, 1) = vida_ejercicio
     
    campos(0, 2) = clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    condicion = "codigo='" & codigo & "' and año='" & AÑOACTUAL & "' "
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    End Sub




Function ExistemMovimientos(codigo, año) As Boolean
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select count(codigo) from " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
    csql.sql = csql.sql & " where codigo='" & codigo & "' "
    csql.sql = csql.sql & " and año >'" & año & "' "
    csql.Execute
    ExistemMovimientos = False
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    If resultado(0) > 0 Then
        ExistemMovimientos = True
    Else
        ExistemMovimientos = False
    End If
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function





Function TieneCorreccion(familia) As Boolean
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = conta
    csql.sql = "select correccion_monetaria from " & clientesistema & "conta.maestro_familias_nuevo"
    csql.sql = csql.sql & " where codigo='" & familia & "' "
    
    csql.Execute
    TieneCorreccion = False
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    If resultado(0) > 0 Then
        TieneCorreccion = True
    Else
        TieneCorreccion = False
    End If
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function




Function TieneDepreciacion(familia) As Boolean
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = conta
    csql.sql = "select depreciacion from " & clientesistema & "conta.maestro_familias_nuevo"
    csql.sql = csql.sql & " where codigo='" & familia & "' "
    
    csql.Execute
    TieneDepreciacion = False
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    If resultado(0) > 0 Then
        TieneDepreciacion = True
    Else
        TieneDepreciacion = False
    End If
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function





Function AsientoTemporal(cuenta, glosa, monto, DH) As Boolean
'    Dim csql As New rdoQuery
'    Dim resultado As rdoResultset
'
'    Set csql.ActiveConnection = contadb
Dim numero As String

    
    ' call grabarcomprobante_lineas("AF",
'    If glosa = "elimina" Then
'         csql.sql = "delete from " & clientesistema & "conta" & empresaactiva & ".temporal_asiento_activofijo"
'         csql.sql = csql.sql & " where año='" & Format(fechasistema, "yyyy") & "'"
'         csql.sql = csql.sql & " and ip='" & GetWanIP & "' "
'    Else
'
'    csql.sql = "replace into " & clientesistema & "conta" & empresaactiva & ".temporal_asiento_activofijo"
'    csql.sql = csql.sql & " values ('" & GetWanIP & "','"
'    csql.sql = csql.sql & Format(fechasistema, "yyyy") & "','"
'    csql.sql = csql.sql & cuenta & "','"
'    csql.sql = csql.sql & glosa & "','"
'    csql.sql = csql.sql & monto & "','"
'    csql.sql = csql.sql & dh & "','"
'    csql.sql = csql.sql & LINEA & "') "
'    End If
'    csql.Execute
''    TieneDepreciacion = False
'' If csql.RowsAffected > 0 Then
''    Set resultado = csql.OpenResultset
''    If resultado(0) > 0 Then
''        TieneDepreciacion = True
''    Else
''        TieneDepreciacion = False
''    End If
'' End If
' csql.Close
' Set csql = Nothing
' Set resultado = Nothing
 
End Function


Public Function LEERULTIMOFOLIO(tipo) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from movimientoscontables where mes = '" & Format(MES, "00") & "' AND año = '" & año & "' and tipo='" + tipo + "' "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIO = Format(resultados(0) + 1, "0000000000")
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
    If monto = 0 Then Exit Sub
    
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
    
    If tipodocumento = "" Then tipodocumento = tipo
    If numerodocumento = "" Then numerodocumento = numero
    If fechadocumento = "" Then fechadocumento = fecha
    If fechavencimiento = "" Then fechavencimiento = fecha
    
    
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

