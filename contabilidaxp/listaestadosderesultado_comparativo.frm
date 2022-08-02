VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form infoge05 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista Estado de Resultados Comparativos"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   636
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FRMBALA 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16960
      BackColor       =   16744576
      Caption         =   "Estados de Resultados"
      BackColor       =   16744576
      ForeColor       =   65535
      BordeColor      =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   16744576
      Begin VB.CheckBox chksimula 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Incluye Simulacion Ventas"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   8520
         Width           =   2295
      End
      Begin VB.CheckBox chkrenta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Incluye Impto Proyectado"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   8520
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   8760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         BackColor       =   16761024
         Caption         =   "Taza de Proyeccion Renta"
         BackColor       =   16761024
         ColorBarraArriba=   4194304
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
         HabilitarArrastre=   -1  'True
         Begin VB.TextBox TXTRENTA 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Text            =   "20"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EXPORTAR EXCEL"
         Height          =   375
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8280
         Width           =   2055
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   7800
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "PROCESAR INFORME"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   8280
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIME"
         Height          =   375
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8280
         Width           =   2055
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7335
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   12938
         Appearance      =   0
         BackColor1      =   14737632
         BackColorActiveCellSel=   8454016
         BackColorBkg    =   -2147483643
         BackColorFixed  =   -2147483647
         BackColorFixedSel=   -2147483639
         BackColorScrollBar=   -2147483639
         BackColorSel    =   16777215
         Cols            =   5
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         FixedRowColStyle=   0
         ForeColorFixed  =   16777215
         GridColor       =   -2147483641
         Rows            =   30
         SelectionMode   =   1
      End
      Begin XPFrame.FrameXp FrameXp9 
         Height          =   735
         Left            =   11160
         TabIndex        =   5
         Top             =   8760
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1296
         BackColor       =   16761024
         Caption         =   "Centros de Costo"
         BackColor       =   16761024
         ColorBarraArriba=   4194304
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
         Begin VB.ComboBox Combocrcc 
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
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   3735
         End
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   735
         Left            =   4920
         TabIndex        =   7
         Top             =   8760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1296
         BackColor       =   16761024
         Caption         =   "Meses a Comparar"
         BackColor       =   16761024
         ColorBarraArriba=   4194304
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
         Begin VB.TextBox txtmes 
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
            Height          =   405
            Left            =   120
            TabIndex        =   8
            Text            =   "3"
            Top             =   240
            Width           =   1455
         End
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   735
         Left            =   2520
         TabIndex        =   13
         Top             =   8760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         BackColor       =   16761024
         Caption         =   "% Simulacion Ventas (+) o (-)"
         BackColor       =   16761024
         ColorBarraArriba=   4194304
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
         HabilitarArrastre=   -1  'True
         Begin VB.TextBox TXTVENTA 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   720
            MaxLength       =   2
            TabIndex        =   14
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   735
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   615
         Left            =   12000
         TabIndex        =   18
         Top             =   8160
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
            TabIndex        =   20
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton botonmisfavoritos 
            Caption         =   "Mis Favoritos"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   735
         Left            =   6720
         TabIndex        =   21
         Top             =   8760
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1296
         BackColor       =   16761024
         Caption         =   "Maestro de Empresas"
         BackColor       =   16761024
         ColorBarraArriba=   4194304
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
         Begin VB.ComboBox comboempresas 
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
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "PRESIONE O PARA AGREGAR OBSERVACIONES"
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
         Left            =   120
         TabIndex        =   23
         Top             =   8160
         Width           =   4575
      End
   End
End
Attribute VB_Name = "infoge05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LINEAREMU As Double
Public linea_RESULTADO As Double

Public saldoglobal As Double
Public ROW1 As Double
Dim totales(40) As Double
Dim totales2(40, 40) As Double
Private LINEA As Double
Private linea2 As Double
Private PORCE2 As Double
Private empresa As String
Private PRIMERAVEZ As Double










Private Sub Combocrcc_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub comboempresas_Click()

Call CARGAcrcc(Mid(comboempresas.text, 1, 2))
End Sub

Private Sub comboempresas_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
imprimir

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
        For k = 1 To 5
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
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub



Private Sub COMMAND2_Click()
If chksimula.Value = 1 Then

End If
fechadata = leerdatos(conta, "maestroempresas", "flujoactualizado", "codigoempresa='" + Mid(comboempresas.text, 1, 2) + "'")
If fechadata <> 0 And fechadata <= fechasistema Then
If Verifica_Permiso(Me.Caption, "autoriza") = False Then
    fechasistema = Format(fechadata, "yyyy-mm-dd")
End If
End If
txtmes.text = CDbl(Format(fechasistema, "mm"))
frmbala.Caption = "ESTADO RESULTADOS ULTIMOS " & txtmes.text & " de " & Format(fechasistema, "mm-yyyy")

If Val(txtmes.text) < 1 Or Val(txtmes.text) > 12 Then
MsgBox "PERIODO NO PUEDE SUPERAR LOS 12 MESES"
txtmes.text = "12"
txtmes.SetFocus

Exit Sub

End If

frmbala.Caption = "ESTADO RESULTADOS ULTIMOS " & txtmes.text & " de " & Format(fechasistema, "mm-yyyy")

CARGAGRILLA
Call cargadatos(Mid(Combocrcc.text, 1, 2) + Mid(Combocrcc.text, 4, 2), Mid(comboempresas.text, 1, 2))
Call leecapital(Mid(Combocrcc.text, 1, 2) + Mid(Combocrcc.text, 4, 2), Mid(comboempresas.text, 1, 2))

PRIMERAVEZ = 0

    Grid1.Column(2).Locked = False
    Grid1.Column(3).Locked = False
    Grid1.Column(4).Locked = False
    Grid1.Column(5).Locked = False
    Grid1.Column(6).Locked = False
    Grid1.Column(7).Locked = False
'    Grid1.Column(8).Locked = False
'    Grid1.Column(9).Locked = False
'    Grid1.Column(10).Locked = False
     
        Call cargaObservaciones(Mid(Combocrcc.text, 1, 5), Mid(comboempresas.text, 1, 2), Grid1.Cell(0, Grid1.Cols - 5).text)
    
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
'    Grid1.Column(8).Locked = True
'    Grid1.Column(9).Locked = True
'    Grid1.Column(10).Locked = True

    
End Sub
Sub cargaObservaciones(CRCC, emp, fecha)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim k As Double
    Dim ANOANTE As String
    
    Set csql.ActiveConnection = conta
    ANOANTE = "01-" & Format(DateAdd("d", -365, Mid(fecha, 4, 4) & "-" & Mid(fecha, 1, 2) & "-01"), "yyyy")
    csql.sql = "select crcc,nombre,glosa,fecha "
    csql.sql = csql.sql & "from analisis_eerr "
    csql.sql = csql.sql & "where (fecha like '%" & Format(fecha, "yyyy") & "' or fecha like '%" & Format(ANOANTE, "yyyy") & "%') and crcc='" & CRCC & "' and empresa='" & emp & "' "
    csql.sql = csql.sql & " order by fecha,nombre "
    csql.Execute
    If csql.RowsAffected > 0 Then
        If PRIMERAVEZ = 0 Then
            PRIMERAVEZ = 1
            Grid1.Rows = Grid1.Rows + 2
            Grid1.Cell(Grid1.Rows - 1, 1).text = "OBSERVACIONES"
            Grid1.Cell(Grid1.Rows - 1, 1).Font.Bold = True
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 1).Alignment = cellCenterGeneral
            Grid1.Range(Grid1.Rows - 1, 2, Grid1.Rows - 1, 5).Merge
            Grid1.Range(Grid1.Rows - 1, 2, Grid1.Rows - 1, 5).Alignment = cellCenterGeneral
            
'            Grid1.Range(Grid1.Rows - 1, 6, Grid1.Rows - 1, 7).Merge
'            Grid1.Range(Grid1.Rows - 1, 6, Grid1.Rows - 1, 7).Alignment = cellCenterGeneral
            
            Grid1.Cell(Grid1.Rows - 1, 1).text = "OBSERVACION"
            Grid1.Cell(Grid1.Rows - 1, 1).Font.Bold = True
            Grid1.Cell(Grid1.Rows - 1, 2).text = "NOMBRE"
            Grid1.Cell(Grid1.Rows - 1, 2).Font.Bold = True
            Grid1.Cell(Grid1.Rows - 1, 6).text = "PERIODO"
            Grid1.Cell(Grid1.Rows - 1, 6).Font.Bold = True
        End If
        
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Grid1.Rows = Grid1.Rows + 1
            
            
            
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 1).Alignment = cellLeftCenter
            Grid1.Range(Grid1.Rows - 1, 2, Grid1.Rows - 1, 5).Merge
            Grid1.Range(Grid1.Rows - 1, 2, Grid1.Rows - 1, 5).Alignment = cellLeftCenter
            Grid1.Range(Grid1.Rows - 1, 6, Grid1.Rows - 1, 7).Merge
            Grid1.Range(Grid1.Rows - 1, 6, Grid1.Rows - 1, 7).Alignment = cellLeftCenter
            
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(2)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
'
           'Grid1.Cell(Grid1.Rows - 1, 6).CellType = cellCalendar
            Grid1.Cell(Grid1.Rows - 1, 6).text = "PERIODO " & resultados(3)
             
             For k = 1 To Grid1.Rows - 1
                If resultados(1) = Grid1.Cell(k, 1).text Then
                    Dim E As Double
                    For E = 1 To Grid1.Cols - 1
                    If resultados(3) = Grid1.Cell(0, E).text Then
                    Grid1.Cell(k, E).BackColor = &HC0FFC0
                    End If
                    Next E
                
                
                
                End If
             Next k
           
             
            
            resultados.MoveNext
        Wend
        
    End If
End Sub
Private Sub Command3_Click()
Call Grid1.ExportToExcel("estado_resultado_" & empresaactiva & ".xls", True)

End Sub

Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

End Sub

Private Sub Form_Load()


'dibu1.FileName = App.path & "\archivo.gif"
'dibu2.FileName = App.path & "\archivo.gif"
Call Conectar_BD
Call CARGAPERMISO(Me.Name)
CARGAGRILLA
barra.Value = 0
barra.Refresh
fechadata = leerdatos(conta, "maestroempresas", "flujoactualizado", "codigoempresa='" + empresaactiva + "'")
If fechadata <> 0 And fechadata <= fechasistema Then
If Verifica_Permiso(Me.Caption, "autoriza") = False Then
    fechasistema = Format(fechadata, "yyyy-mm-dd")
End If
End If

txtmes.text = CDbl(Format(fechasistema, "mm"))

frmbala.Caption = "ESTADO RESULTADOS ULTIMOS " & txtmes.text & " de " & Format(fechasistema, "mm-yyyy")

CARGAempresas
If Format(fechasistema, "yyyy") < "2014" Then TXTRENTA.text = "20"
If Format(fechasistema, "yyyy") = "2014" Then TXTRENTA.text = "21"
If Format(fechasistema, "yyyy") = "2015" Then TXTRENTA.text = "22,5"
If Format(fechasistema, "yyyy") = "2016" Then TXTRENTA.text = "24"
If Format(fechasistema, "yyyy") = "2017" Then TXTRENTA.text = "25,5"
If Format(fechasistema, "yyyy") = "2018" Then TXTRENTA.text = "27"
If Format(fechasistema, "yyyy") = "2019" Then TXTRENTA.text = "27"
If Format(fechasistema, "yyyy") = "2020" Then TXTRENTA.text = "27"

TXTRENTA.text = leerdatos(conta, "maestroempresas", "flujoactualizado", "codigoempresa='" + empresaactiva + "'")


 Call CENTRAR(Me)
End Sub



Sub CARGAGRILLA()
    Dim formatogrilla2(10, 50)
    Dim fec As String
    Dim fec2 As String
    Dim contador As Double
    Dim PORCE10 As Double
    
    formatogrilla2(1, 1) = "NOMBRE"
    fec = fechasistema
    contador = 1
    MESES = Val(txtmes.text)
    Grid1.Cols = 2 + (MESES * 3)
    Grid1.DefaultFont.Size = 7
    
    For k = 1 To MESES
    fec = DateAdd("m", (MESES - k) * -1, fechasistema)
    fec2 = DateAdd("yyyy", -1, fec)
    contador = contador + 1
    formatogrilla2(1, contador) = Format(fec2, "mm-yyyy")
    formatogrilla2(2, contador) = "10"
    formatogrilla2(3, contador) = "N"
    formatogrilla2(4, contador) = " ###,###,###,##0"
    formatogrilla2(5, contador) = "TRUE"
    
    contador = contador + 1
    formatogrilla2(1, contador) = Format(fec, "mm-yyyy")
    formatogrilla2(2, contador) = "10"
    formatogrilla2(3, contador) = "N"
    formatogrilla2(4, contador) = " ###,###,###,##0"
    formatogrilla2(5, contador) = "TRUE"
    
    
    contador = contador + 1
    formatogrilla2(1, contador) = "%"
    formatogrilla2(2, contador) = "4"
    formatogrilla2(3, contador) = "N"
    formatogrilla2(4, contador) = "##0.0"
    formatogrilla2(5, contador) = "TRUE"
    
    
    Next k
    Grid1.Cols = Grid1.Cols + 3
    contador = contador + 1
    formatogrilla2(1, contador) = Format(fechasistema, "YYYY") - 1
    formatogrilla2(2, contador) = "10"
    formatogrilla2(3, contador) = "N"
    formatogrilla2(4, contador) = " ###,###,###,##0"
    formatogrilla2(5, contador) = "TRUE"
    
    contador = contador + 1
    formatogrilla2(1, contador) = Format(fechasistema, "YYYY")
    formatogrilla2(2, contador) = "10"
    formatogrilla2(3, contador) = "N"
    formatogrilla2(4, contador) = " ###,###,###,##0"
    formatogrilla2(5, contador) = "TRUE"
    
    
    contador = contador + 1
    formatogrilla2(1, contador) = "%"
    formatogrilla2(2, contador) = "4"
    formatogrilla2(3, contador) = "N"
    formatogrilla2(4, contador) = "##0.0"
    formatogrilla2(5, contador) = "TRUE"
    
    
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "30"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    
    Rem FORMATO GRILLA
    
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    
    Grid1.Rows = 1
'    Grid1.AllowUserResizing = False
'    Grid1.DisplayFocusRect = False
'    Grid1.ExtendLastCol = True
'    Grid1.BoldFixedCell = False
'    Grid1.DrawMode = cellOwnerDraw
'    Grid1.Appearance = Flat
'    Grid1.ScrollBarStyle = Flat
'    Grid1.FixedRowColStyle = Flat
Rem     Grid1.BackColorFixed = RGB(90, 158, 214)
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter: Grid1.Column(k).Mask = cellNumeric
        
        
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
 
 Grid1.SelectionMode = cellSelectionFree
 
 
    End Sub


Sub leecapital(CRCC, empresa2)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim SALTO1 As Double
    Dim SALTO2 As Double
    Dim SALTO3 As Double
    empresa = empresa2
    
    Dim PASO As String
    Dim LINEAS2 As Double
    For k = 1 To 40
    totales(k) = 0
    For j = 1 To 40
    totales2(k, j) = 0
    
    Next j
    
    Next k
    
           
        Set csql2.ActiveConnection = conta
        csql2.sql = "SELECT codigo,glosa,signo "
        csql2.sql = csql2.sql + "FROM balanceclasificado_titulos where codigo>'10'  "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        Grid1.AutoRedraw = False
        Grid1.Rows = 1
        barra.Value = 0
        If csql2.RowsAffected > 0 Then
        barra.Max = csql2.RowsAffected
        
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &HFFC0C0
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).ForeColor = 0
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontBold = True
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(1)
      
        Call leeCAPITALDETALLE(resultados2(0), resultados2(2), CRCC)
        
        LINEAS2 = 2
        MESES = txtmes.text
        
        
        For k = 1 To MESES
        
        If resultados2(2) = "T" Then
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
        
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales(LINEAS2)
        totales2(CDbl(resultados2(0)), LINEAS2) = totales(LINEAS2)
        totales(LINEAS2) = 0
        End If
        If resultados2(0) = "05" Then
        totales2(5, LINEAS2) = totales2(2, LINEAS2) + totales2(4, LINEAS2)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(5, LINEAS2)
        End If
        
        
        If resultados2(0) = "10" Then
        totales2(10, LINEAS2) = totales2(7, LINEAS2) + totales2(9, LINEAS2)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(10, LINEAS2)
        End If

        If resultados2(0) = "15" Then
        totales2(15, LINEAS2) = totales2(12, LINEAS2) - totales2(14, LINEAS2)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(15, LINEAS2)
        End If
        If resultados2(0) = "18" Then
        totales2(18, LINEAS2) = totales2(15, LINEAS2) - totales2(17, LINEAS2)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(18, LINEAS2)
        End If
        If resultados2(0) = "27" Then
        totales2(27, LINEAS2) = (totales2(19, LINEAS2) + totales2(20, LINEAS2) + totales2(21, LINEAS2)) - (totales2(22, LINEAS2) + totales2(23, LINEAS2) + totales2(24, LINEAS2) + totales2(25, LINEAS2) + totales2(26, LINEAS2))
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(27, LINEAS2)
        End If

        If resultados2(0) = "28" Then
        totales2(28, LINEAS2) = totales2(18, LINEAS2) + totales2(27, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(28, LINEAS2)
        End If
        If resultados2(0) = "29" Then
        totales2(29, LINEAS2) = totales2(29, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(29, LINEAS2)
        End If
        
        If resultados2(0) = "30" Then
        totales2(30, LINEAS2) = totales2(28, LINEAS2) - totales2(29, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(30, LINEAS2)
        End If
        If resultados2(0) = "32" Then
        totales2(32, LINEAS2) = totales2(30, LINEAS2) - totales2(31, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(32, LINEAS2)
        End If
        If resultados2(0) = "34" Then
        totales2(34, LINEAS2) = totales2(32, LINEAS2) - totales2(33, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(34, LINEAS2)
        End If
        LINEAS2 = LINEAS2 + 3
        Next k
        LINEAS2 = 3
        MESES = txtmes.text
        
        
        For k = 1 To MESES
        
        If resultados2(2) = "T" Then
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
        
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales(LINEAS2)
        totales2(CDbl(resultados2(0)), LINEAS2) = totales(LINEAS2)
        totales(LINEAS2) = 0
        End If
        If resultados2(0) = "05" Then
        totales2(5, LINEAS2) = totales2(2, LINEAS2) + totales2(4, LINEAS2)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(5, LINEAS2)
        End If
        
        
        If resultados2(0) = "10" Then
        totales2(10, LINEAS2) = totales2(7, LINEAS2) + totales2(9, LINEAS2)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(10, LINEAS2)
        End If

        If resultados2(0) = "15" Then
        totales2(15, LINEAS2) = totales2(12, LINEAS2) - totales2(14, LINEAS2)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(15, LINEAS2)
        
        End If
        If resultados2(0) = "18" Then
        totales2(18, LINEAS2) = totales2(15, LINEAS2) - totales2(17, LINEAS2)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(18, LINEAS2)
        End If
        If resultados2(0) = "27" Then
        totales2(27, LINEAS2) = (totales2(19, LINEAS2) + totales2(20, LINEAS2) + totales2(21, LINEAS2)) - (totales2(22, LINEAS2) + totales2(23, LINEAS2) + totales2(24, LINEAS2) + totales2(25, LINEAS2) + totales2(26, LINEAS2))
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(27, LINEAS2)
        End If

        If resultados2(0) = "28" Then
        totales2(28, LINEAS2) = totales2(18, LINEAS2) + totales2(27, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(28, LINEAS2)
        End If
        If resultados2(0) = "29" Then
        totales2(29, LINEAS2) = totales2(29, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(29, LINEAS2)
        End If
        
        If resultados2(0) = "30" Then
        totales2(30, LINEAS2) = totales2(28, LINEAS2) - totales2(29, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(30, LINEAS2)
        End If
        If resultados2(0) = "32" Then
        totales2(32, LINEAS2) = totales2(30, LINEAS2) - totales2(31, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(32, LINEAS2)
        End If
        If resultados2(0) = "34" Then
        totales2(34, LINEAS2) = totales2(32, LINEAS2) - totales2(33, LINEAS2)
       Rem  totales2(27) = totales2(18) - totales2(27)
        Grid1.Cell(Grid1.Rows - 1, LINEAS2).text = totales2(34, LINEAS2)
        linea_RESULTADO = Grid1.Rows - 1
        
        End If
        LINEAS2 = LINEAS2 + 3
        Next k
        
        barra.Value = barra.Value + 1
        barra.Refresh
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    MESES = txtmes.text
    For k = 1 To Grid1.Rows - 1
    suma1 = 0
    suma2 = 0
    
    largo = 0
    For o = 1 To MESES
    largo = largo + 3
    porce = 0
    If o = 1 Then
            
                Grid1.Cell(k, 4).BackColor = &HC0FFC0
                If Val(Grid1.Cell(k, 2).text) <> "0" Or Val(Grid1.Cell(k, 3).text) <> "0" Then
                 
                If Val(Grid1.Cell(k, 2).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, 3).text) / CDbl(Grid1.Cell(k, 2).text)) - 1) * 100
                End If
                
                Grid1.Cell(k, 4).text = Round(porce, 1)
                
                
                suma1 = suma1 + CDbl(Grid1.Cell(k, 2).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 3).text)
                End If
    End If
    If o = 2 Then
                Grid1.Cell(k, 7).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 5).text) <> "0" Or Val(Grid1.Cell(k, 6).text) <> "0" Then
                If Val(Grid1.Cell(k, 5).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, 6).text) / CDbl(Grid1.Cell(k, 5).text)) - 1) * 100
                End If
                Grid1.Cell(k, 7).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 5).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 6).text)
                End If
    End If
    If o = 3 Then
                Grid1.Cell(k, 10).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 8).text) <> "0" Or Val(Grid1.Cell(k, 9).text) <> "0" Then
                If Val(Grid1.Cell(k, 8).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, 9).text) / CDbl(Grid1.Cell(k, 8).text)) - 1) * 100
                End If
                Grid1.Cell(k, 10).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 8).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 9).text)
                
                End If
    End If
    If o = 4 Then
                Grid1.Cell(k, 13).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 11).text) <> "0" Or Val(Grid1.Cell(k, 12).text) <> "0" Then
                If Val(Grid1.Cell(k, 11).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, 12).text) / CDbl(Grid1.Cell(k, 11).text)) - 1) * 100
                End If
                Grid1.Cell(k, 13).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 11).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 12).text)
                
                End If
    End If
    If o = 5 Then
                Grid1.Cell(k, 16).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 14).text) <> "0" Or Val(Grid1.Cell(k, 15).text) <> "0" Then
                If Val(Grid1.Cell(k, 14).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, 15).text) / CDbl(Grid1.Cell(k, 14).text)) - 1) * 100
                End If
                Grid1.Cell(k, 16).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 14).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 15).text)
                
                End If
    End If
    If o = 6 Then
                Grid1.Cell(k, 19).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 17).text) <> "0" Or Val(Grid1.Cell(k, 18).text) <> "0" Then
                If Val(Grid1.Cell(k, 17).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, 18).text) / CDbl(Grid1.Cell(k, 17).text)) - 1) * 100
                End If
                Grid1.Cell(k, 19).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 17).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 18).text)
                
                End If
    End If
    If o = 7 Then
                Grid1.Cell(k, 22).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 20).text) <> "0" Or Val(Grid1.Cell(k, 21).text) <> "0" Then
                If Val(Grid1.Cell(k, 20).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, 21).text) / CDbl(Grid1.Cell(k, 20).text)) - 1) * 100
                End If
                Grid1.Cell(k, 22).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 20).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 21).text)
                
                End If
    End If
    If o = 8 Then
                Grid1.Cell(k, 25).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 23).text) <> "0" Or Val(Grid1.Cell(k, 24).text) <> "0" Then
                If Val(Grid1.Cell(k, 23).text) <> 0 Then
                
                porce = ((CDbl(Grid1.Cell(k, 24).text) / CDbl(Grid1.Cell(k, 23).text)) - 1) * 100
                End If
                Grid1.Cell(k, 25).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 23).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 24).text)
                
                End If
    End If
    If o = 9 Then
                Grid1.Cell(k, 28).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 26).text) <> "0" Or Val(Grid1.Cell(k, 27).text) <> "0" Then
                If Val(Grid1.Cell(k, 26).text) <> 0 Then
                
                porce = ((CDbl(Grid1.Cell(k, 27).text) / CDbl(Grid1.Cell(k, 26).text)) - 1) * 100
                End If
                Grid1.Cell(k, 28).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 26).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 27).text)
                
                End If
    End If
    If o = 10 Then
                Grid1.Cell(k, 31).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 29).text) <> "0" Or Val(Grid1.Cell(k, 30).text) <> "0" Then
                If Val(Grid1.Cell(k, 29).text) <> 0 Then
                
                porce = ((CDbl(Grid1.Cell(k, 30).text) / CDbl(Grid1.Cell(k, 29).text)) - 1) * 100
                End If
                Grid1.Cell(k, 31).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 29).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 30).text)
                
                End If
    End If
    If o = 11 Then
                Grid1.Cell(k, 34).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 32).text) <> "0" Or Val(Grid1.Cell(k, 33).text) <> "0" Then
                If Val(Grid1.Cell(k, 32).text) <> 0 Then
                
                porce = ((CDbl(Grid1.Cell(k, 33).text) / CDbl(Grid1.Cell(k, 32).text)) - 1) * 100
                End If
                Grid1.Cell(k, 34).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 32).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 33).text)
                
                End If
    End If
    If o = 12 Then
                Grid1.Cell(k, 37).BackColor = &HC0FFC0
                
                If Val(Grid1.Cell(k, 35).text) <> "0" Or Val(Grid1.Cell(k, 36).text) <> "0" Then
                If Val(Grid1.Cell(k, 35).text) <> 0 Then
                
                porce = ((CDbl(Grid1.Cell(k, 36).text) / CDbl(Grid1.Cell(k, 35).text)) - 1) * 100
                End If
                Grid1.Cell(k, 37).text = Round(porce, 1)
                suma1 = suma1 + CDbl(Grid1.Cell(k, 35).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, 36).text)
                
                End If
    End If
    
    Next o
    Grid1.Cell(k, Grid1.Cols - 1).BackColor = &HC0FFC0
                
    If suma1 <> 0 Or suma2 <> 0 Then
    
                
                If suma1 <> 0 Then
                porce = ((suma2 / suma1) - 1) * 100
                End If
                Grid1.Cell(k, Grid1.Cols - 3).text = suma1
                Grid1.Cell(k, Grid1.Cols - 2).text = suma2
                
                Grid1.Cell(k, Grid1.Cols - 1).text = Round(porce, 1)
                
                End If
    
    Next k
    
    
    Grid1.FrozenCols = 1
    
    Rem TOTALES PARA RESUMEN
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "CUADRO RESUMEN "
    Grid1.Cell(Grid1.Rows - 1, 1).Font.Bold = True
    
    
    For k = 1 To Grid1.Rows - 1
    If Mid(Grid1.Cell(k, 1).text, 1, 28) = "INGRESOS POR VENTA COMERCIAL" Then
    Grid1.Rows = Grid1.Rows + 1
    For s = 1 To Grid1.Cols - 1
    
    Grid1.Cell(Grid1.Rows - 1, s).text = Grid1.Cell(k, s).text
    
    If s = 1 Then
        Grid1.Cell(Grid1.Rows - 1, s).text = Grid1.Cell(Grid1.Rows - 1, s).text & " R"
    End If
    Next s
    Rem TOTALES
    suma1 = 0
    suma2 = 0
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES
    porce = 0
    
            
                Grid1.Cell(k, largo + 2).BackColor = &HC0FFC0
                If Val(Grid1.Cell(k, largo).text) <> "0" Or Val(Grid1.Cell(k, largo + 1).text) <> "0" Then
                If Val(Grid1.Cell(k, largo).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, largo + 1).text) / CDbl(Grid1.Cell(k, largo).text)) - 1) * 100
                End If
                Grid1.Cell(k, largo + 2).text = Round(porce, 1)
                
                suma1 = suma1 + CDbl(Grid1.Cell(k, largo).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, largo + 1).text)
                largo = largo + 3
                
    
    End If
    Next o
    Grid1.Cell(k, Grid1.Cols - 1).BackColor = &HC0FFC0
                
    If suma1 <> 0 Or suma2 <> 0 Then
                If suma1 <> 0 Then
                porce = ((suma2 / suma1) - 1) * 100
                End If
                Grid1.Cell(k, Grid1.Cols - 3).text = suma1
                Grid1.Cell(k, Grid1.Cols - 2).text = suma2
                
                Grid1.Cell(k, Grid1.Cols - 1).text = Round(porce, 1)
                
                End If
    
    
    End If
    
    
    Next k
    
    Rem INGRESOS OPERACIONALES
    For k = 1 To Grid1.Rows - 1
    If Mid(Grid1.Cell(k, 1).text, 1, 22) = "INGRESOS OPERACIONALES" Then
    Grid1.Rows = Grid1.Rows + 1
    For s = 1 To Grid1.Cols - 1
    Grid1.Cell(Grid1.Rows - 1, s).text = Grid1.Cell(k, s).text
     If s = 1 Then
        Grid1.Cell(Grid1.Rows - 1, s).text = Grid1.Cell(Grid1.Rows - 1, s).text & " R"
    End If
    Next s
    Rem TOTALES
    suma1 = 0
    suma2 = 0
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES
    
    porce = 0
            
                Grid1.Cell(k, largo + 2).BackColor = &HC0FFC0
                If Val(Grid1.Cell(k, largo).text) <> "0" Or Val(Grid1.Cell(k, largo + 1).text) <> "0" Then
                If Val(Grid1.Cell(k, largo).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, largo + 1).text) / CDbl(Grid1.Cell(k, largo).text)) - 1) * 100
                End If
                Grid1.Cell(k, largo + 2).text = Round(porce, 1)
                
                suma1 = suma1 + CDbl(Grid1.Cell(k, largo).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, largo + 1).text)
                largo = largo + 3
                
    
    End If
    Next o
    Grid1.Cell(k, Grid1.Cols - 1).BackColor = &HC0FFC0
                
    If suma1 <> 0 Or suma2 <> 0 Then
                
                If suma1 <> 0 Then
                porce = ((suma2 / suma1) - 1) * 100
                End If
                Grid1.Cell(k, Grid1.Cols - 3).text = suma1
                Grid1.Cell(k, Grid1.Cols - 2).text = suma2
                
                Grid1.Cell(k, Grid1.Cols - 1).text = Round(porce, 1)
                
                End If
    
    
    End If
    
    
    Next k
    
    
    Rem TOTAL COSTO DE VENTA
    For k = 1 To Grid1.Rows - 1
    If Mid(Grid1.Cell(k, 1).text, 1, 24) = "COSTO DE VENTA COMERCIAL" Then
    Grid1.Rows = Grid1.Rows + 1
    For s = 1 To Grid1.Cols - 1
    Grid1.Cell(Grid1.Rows - 1, s).text = Grid1.Cell(k, s).text
     If s = 1 Then
        Grid1.Cell(Grid1.Rows - 1, s).text = Grid1.Cell(Grid1.Rows - 1, s).text & " R"
    End If
    Next s
    Rem TOTALES
    suma1 = 0
    suma2 = 0
    k = Grid1.Rows - 1
    largo = 0
    largo = 2
    For o = 1 To MESES
    porce = 0
    
            
                Grid1.Cell(k, largo + 2).BackColor = &HC0FFC0
                If Val(Grid1.Cell(k, largo).text) <> "0" Or Val(Grid1.Cell(k, largo + 1).text) <> "0" Then
                If Val(Grid1.Cell(k, largo).text) <> 0 Then
                porce = ((CDbl(Grid1.Cell(k, largo + 1).text) / CDbl(Grid1.Cell(k, largo).text)) - 1) * 100
                End If
                Grid1.Cell(k, largo + 2).text = Round(porce, 1)
                
                suma1 = suma1 + CDbl(Grid1.Cell(k, largo).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, largo + 1).text)
                largo = largo + 3
                
    
    End If
    Next o
    
    Grid1.Cell(k, Grid1.Cols - 1).BackColor = &HC0FFC0
                
    If suma1 <> 0 Or suma2 <> 0 Then
                If suma1 <> 0 Then
                porce = ((suma2 / suma1) - 1) * 100
                
                End If
                Grid1.Cell(k, Grid1.Cols - 3).text = suma1
                Grid1.Cell(k, Grid1.Cols - 2).text = suma2
                
                Grid1.Cell(k, Grid1.Cols - 1).text = Round(porce, 1)
                
                End If
    

    End If
    
    
    Next k
    
    Rem MARGEN DE COMERCIALIZACION
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "MARGEN DE EXPLOTACION R"
    
    
    suma1 = 0
    suma2 = 0
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES
                
                
                Grid1.Cell(k, largo).text = Val(Grid1.Cell(k - 3, largo).text) - Val(Grid1.Cell(k - 1, largo).text)
                Grid1.Cell(k, largo + 1).text = Val(Grid1.Cell(k - 3, largo + 1).text) - Val(Grid1.Cell(k - 1, largo + 1).text)
                suma1 = suma1 + CDbl(Grid1.Cell(k, largo).text)
                suma2 = suma2 + CDbl(Grid1.Cell(k, largo + 1).text)
                
                
                largo = largo + 3
    Next o
    
    If suma1 <> 0 Or suma2 <> 0 Then
                If suma1 <> 0 Then
                porce = ((suma2 / suma1) - 1) * 100
                End If
                Grid1.Cell(k, Grid1.Cols - 3).text = suma1
                Grid1.Cell(k, Grid1.Cols - 2).text = suma2
                
                Grid1.Cell(k, Grid1.Cols - 1).text = Round(porce, 1)
                
                End If
    

    
    
    
    Rem MARGEN DE COMERCIALIZACION
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "MARGEN DE COMERCIALIZACION R"
    
    
    suma1 = 0
    suma2 = 0
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES + 1
           porce = 0
                
                If Val(Grid1.Cell(k - 1, largo).text) <> "0" Or Val(Grid1.Cell(k - 2, largo).text) <> "0" Then
                
                If Grid1.Cell(k - 2, largo).text <> "" Then
                 If CDbl(Grid1.Cell(k - 2, largo).text) <> 0 Then
                 porce = ((CDbl(Grid1.Cell(k - 1, largo).text) / CDbl(Grid1.Cell(k - 2, largo).text))) * 100
                 End If
                End If
                
                Grid1.Cell(k, largo).text = Format(Round(porce, 1), "% ##,##0.00")
                End If
                If Val(Grid1.Cell(k - 1, largo + 1).text) <> "0" Or Val(Grid1.Cell(k - 2, largo + 1).text) <> "0" Then
                If Val(Grid1.Cell(k - 2, largo + 1).text) <> 0 Then
                PORCE1 = ((CDbl(Grid1.Cell(k - 1, largo + 1).text) / CDbl(Grid1.Cell(k - 2, largo + 1).text))) * 100
                End If
                
                Grid1.Cell(k, largo + 1).text = Format(Round(PORCE1, 1), "% ##,##0.00")
                End If
                Grid1.Cell(k, largo + 2).text = Format(PORCE1 - porce, "##,##0.00")
    
                largo = largo + 3
    Next o
    
    
    
    
    Rem MARGEN DE CONTRIBUCION
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "MARGEN DE CONTRIBUCION R"
    suma1 = 0
    suma2 = 0
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES + 1
             porce = 0
             PORCE1 = 0
             
                
                If Val(Grid1.Cell(k - 2, largo).text) <> "0" Or Val(Grid1.Cell(k - 5, largo).text) <> "0" Then
                If Val(Grid1.Cell(k - 5, largo).text) <> 0 Then
                
                porce = ((CDbl(Grid1.Cell(k - 2, largo).text) / Val(Grid1.Cell(k - 5, largo).text))) * 100
                End If
                Grid1.Cell(k, largo).text = Format(Round(porce, 1), "% ##,##0.00")
                End If
                If Val(Grid1.Cell(k - 2, largo + 1).text) <> "0" Or Val(Grid1.Cell(k - 5, largo + 1).text) <> "0" Then
                If Val(Grid1.Cell(k - 5, largo + 1).text) <> 0 Then
                
                PORCE1 = ((CDbl(Grid1.Cell(k - 2, largo + 1).text) / CDbl(Grid1.Cell(k - 5, largo + 1).text))) * 100
                End If
                Grid1.Cell(k, largo + 1).text = Format(Round(PORCE1, 1), "% ##,##0.00")
                End If
                Grid1.Cell(k, largo + 2).text = Format(PORCE1 - porce, "##,##0.00")
    
                largo = largo + 3
    Next o
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "MARGEN ULTIMA LINEA DE LA VENTA R"
    suma1 = 0
    suma2 = 0
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES + 1

porce = 0
PORCE1 = 0

                If Val(Grid1.Cell(k - 8, largo).text) <> "0" Or Val(Grid1.Cell(k - 6, largo).text) <> "0" Then
                'ariel cambia cdbl por VAL
                If Val(Grid1.Cell(k - 6, largo).text) <> 0 Then
                
                porce = ((CDbl(Grid1.Cell(k - 8, largo).text) / CDbl(Grid1.Cell(k - 6, largo).text))) * 100
                End If
                Grid1.Cell(k, largo).text = Format(Round(porce, 1), "% ##,##0.00")
                End If
                If Val(Grid1.Cell(k - 8, largo + 1).text) <> "0" Or Val(Grid1.Cell(k - 6, largo + 1).text) <> "0" Then
                If Val(Grid1.Cell(k - 6, largo + 1).text) <> 0 Then
                
                PORCE1 = ((CDbl(Grid1.Cell(k - 8, largo + 1).text) / CDbl(Grid1.Cell(k - 6, largo + 1).text))) * 100
                End If
                Grid1.Cell(k, largo + 1).text = Format(Round(PORCE1, 1), "% ##,##0.00")
                End If
                Grid1.Cell(k, largo + 2).text = Format(PORCE1 - porce, "##,##0.00")

                largo = largo + 3
    Next o
    
    Rem
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "METROS CUADRADOS DE SALA R"
    suma1 = 0
    suma2 = 0
    MTS = SUMAMTS(Mid(Combocrcc.text, 1, 5))
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES + 1


                Grid1.Cell(k, largo).text = Format(MTS, " #,###,##0")
                Grid1.Cell(k, largo + 1).text = Format(MTS, "#,###,##0")
                
                largo = largo + 3
    Next o
    
    
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "VENTA POR METRO CUADRADO R"
    suma1 = 0
    suma2 = 0
    MTS = SUMAMTS(Mid(Combocrcc.text, 1, 5))
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES + 1
porce = 0
PORCET = 0


                
                If MTS <> 0 Then
                porce = (Round(Val(Grid1.Cell(k - 8, largo).text) / MTS))
                Grid1.Cell(k, largo).text = Format(Round(porce, 1), "$ #,###,##0")
                PORCE1 = (Round(Val(Grid1.Cell(k - 8, largo + 1).text) / MTS))
                Grid1.Cell(k, largo + 1).text = Format(Round(PORCE1, 1), "$ ###,###,##0")
                End If
                If porce <> 0 Then
                PORCET = ((PORCE1 - porce) / porce) * 100
                End If
                Grid1.Cell(k, largo + 2).text = Format(PORCET, " ###.00")

                largo = largo + 3
                
    Next o
    
        
Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "COSTO REMUNERACIONES POR METRO CUADRADO R"
    suma1 = 0
    suma2 = 0
    MTS = SUMAMTS(Mid(Combocrcc.text, 1, 5))
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES + 1
PORCE1 = 0
porce = 0
PORCET = 0

                
                If MTS <> 0 Then
                porce = (Round(Val(Grid1.Cell(LINEAREMU, largo).text) / MTS))
                Grid1.Cell(k, largo).text = Format(Round(porce, 1), "$ #,###,##0")
                PORCE1 = (Round(Val(Grid1.Cell(LINEAREMU, largo + 1).text) / MTS))
                Grid1.Cell(k, largo + 1).text = Format(Round(PORCE1, 1), "$ ###,###,##0")
                End If
                If porce <> 0 Then
                PORCET = ((PORCE1 - porce) / porce) * 100
                End If
                Grid1.Cell(k, largo + 2).text = Format(PORCET, " ###.00")

                largo = largo + 3
                
    Next o
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "UTILIDAD POR METRO CUADRADO R"
    suma1 = 0
    suma2 = 0
    MTS = SUMAMTS(Mid(Combocrcc.text, 1, 5))
    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES + 1
            
                If MTS <> 0 Then
                porce = (Round(Val(Grid1.Cell(k - 12, largo).text) / MTS))
                Grid1.Cell(k, largo).text = Format(Round(porce, 1), "$ #,###,##0")
                PORCE1 = (Round(Val(Grid1.Cell(k - 12, largo + 1).text) / MTS))
                Grid1.Cell(k, largo + 1).text = Format(Round(PORCE1, 1), "$ ###,###,##0")
                End If
                If porce <> 0 Then
                PORCET = ((PORCE1 - porce) / porce) * 100
                End If
                Grid1.Cell(k, largo + 2).text = Format(PORCET, " ###.00")

                largo = largo + 3
                
    Next o
    Dim EBITDA(2, 12) As String
    
    For k = 1 To Grid1.Rows - 1
    If Mid(Grid1.Cell(k, 1).text, 1, 26) = "TOTAL INGRESOS EXPLOTACION" Then
    LINEA1 = k
    End If
    If Mid(Grid1.Cell(k, 1).text, 1, 21) = "RESULTADO OPERACIONAL" Then
    linea2 = k
    End If
    
    Next k
    
    
    
    
    
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = "EBITDA R"
    suma1 = 0
    suma2 = 0
    Dim INGRESOEXPLOTACION As Double
    Dim EGRESOEXPLOTACION As Double
    
    

    k = Grid1.Rows - 1
    largo = 2
    For o = 1 To MESES + 1
                INGRESOEXPLOTACION = Val(Grid1.Cell(LINEA1, largo).text)
                EGRESOEXPLOTACION = Val(Grid1.Cell(linea2, largo).text)
                If INGRESOEXPLOTACION = 0 Then INGRESOEXPLOTACION = 1
                EBITA = EGRESOEXPLOTACION / INGRESOEXPLOTACION * 100
                
                Grid1.Cell(k, largo).text = Format(EBITA, "% ###.#0")
                
                INGRESOEXPLOTACION = Val(Grid1.Cell(LINEA1, largo + 1).text)
                EGRESOEXPLOTACION = Val(Grid1.Cell(linea2, largo + 1).text)
                If INGRESOEXPLOTACION = 0 Then INGRESOEXPLOTACION = 1
                EBITA = EGRESOEXPLOTACION / INGRESOEXPLOTACION * 100
                
                Grid1.Cell(k, largo + 1).text = Format(EBITA, "% ###.#0")
'
'                porce = (Round(Val(Grid1.Cell(k - LINEA1, largo).text) / MTS))
'                Grid1.Cell(k, largo).text = Format(Round(porce, 1), "$ #,###,##0")
'                PORCE1 = (Round(Val(Grid1.Cell(k - 12, largo + 1).text) / MTS))
'                Grid1.Cell(k, largo + 1).text = Format(Round(PORCE1, 1), "$ ###,###,##0")
'
                largo = largo + 3
    Next o
    

    Call leeCAPITALDETALLE("35", "-", CRCC)
    
    
    
    Grid1.AutoRedraw = True
        Grid1.Refresh
    
    

End Sub

Sub leeCAPITALDETALLE(codigo, signo, CRCC)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim monto As Double
    empresa = Mid(comboempresas.text, 1, 2)
        
        If codigo <> "35" Then
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT cpd.codigo,cm.nombre "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta" + empresa + ".balanceclasificado_detalle as cpd left join  " + clientesistema + "conta" + empresa + ".cuentasdelmayor as cm on cpd.codigo=cm.codigo and cm.año='" + Format(fechasistema, "yyyy") + "' "
        csql2.sql = csql2.sql + " where cpd.codigotitulo='" + codigo + "' "
        csql2.sql = csql2.sql + "order by cpd.codigo"
        csql2.Execute
        End If
        
        If codigo = "35" Then
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT cpd.codigo,cm.nombre "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta" + empresa + ".balanceclasificado_detalle as cpd left join  " + clientesistema + "conta" + empresa + ".cuentasdelmayor as cm on cpd.codigo=cm.codigo and cm.año='" + Format(fechasistema, "yyyy") + "' "
        csql2.sql = csql2.sql + " where cpd.codigotitulo='35' "
        csql2.sql = csql2.sql + "order by cpd.codigo"
        csql2.Execute
        End If
        
        
        LINEAS = 0
        MESES = Val(txtmes.text)
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        
        While Not resultados2.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 0).text = codigo
        If codigo <> 35 Then
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(1)
        End If
        If codigo = 35 Then
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(1)
        End If
        
        
        Call leersaldomayor2(resultados2(0), MESES, codigo, CRCC)
        Call leersaldomayor3(resultados2(0), MESES, codigo, CRCC)

        Call leeCAPITALDETALLE_detalle(resultados2(0), signo, CRCC, codigo)

                resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    
    
    

End Sub
Sub leeCAPITALDETALLE_detalle(codigo, signo, CRCC, codigo2)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim monto As Double
    empresa = Mid(comboempresas.text, 1, 2)
        Set csql2.ActiveConnection = contadb
        If codigo2 <> "35" Then
        csql2.sql = "SELECT cm.codigo,cm.nombre "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta" + empresa + ".cuentasdelmayor as cm where  cm.año='" + Format(fechasistema, "yyyy") + "' "
        csql2.sql = csql2.sql + " and mid(codigo,1,4)='" + Mid(codigo, 1, 4) + "' and mid(codigo,5,6)<>'0000' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        End If
        
        If codigo2 = "35" Then
        csql2.sql = "SELECT codigo,glosa "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta" + empresa + ".estado_resultado where  codigo='47100001' and glosa<>'' and año='3' "
        csql2.sql = csql2.sql + "  "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        End If
        
        
        
        
        LINEAS = 0
        MESES = Val(txtmes.text)
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        
        While Not resultados2.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 0).text = resultados2(0)
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(1)
        Call leersaldomayor20(resultados2(0), MESES, codigo, CRCC, codigo2, resultados2(1))
        Call leersaldomayor30(resultados2(0), MESES, codigo, CRCC, codigo2, resultados2(1))
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    
    
    

End Sub




Private Sub Grid2_DblClick()

End Sub
Sub grabar(codigo, codigotitulo)
    campos(0, 0) = "codigo"
    campos(1, 0) = "codigotitulo"
    campos(2, 0) = ""
   
    campos(0, 1) = codigo
    campos(1, 1) = codigotitulo
  
    campos(0, 2) = "balanceclasificado_detalle"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If Mid(codigo, 5, 4) = "0000" Then
    Call eliminasubCAPITALDETALLE(codigo)
    
    End If
    
End Sub

Sub eliminaCAPITALDETALLE(codigo)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "delete FROM balanceclasificado_detalle "
        csql2.sql = csql2.sql + " where codigo='" + codigo + "' "
        csql2.Execute
        Call sincronizadatos(csql2.sql, contadb, "")
        
        
End Sub

Sub eliminasubCAPITALDETALLE(codigo)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "delete FROM balanceclasificado_detalle "
        csql2.sql = csql2.sql + " where mid(codigo,1,4)='" + Mid(codigo, 1, 4) + "' and mid(codigo,5,4)<>'0000'  "
        csql2.Execute
        Call sincronizadatos(csql2.sql, contadb, "")
        
        
End Sub

Public Function existeCAPITALDETALLE(codigo) As Boolean


Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "select * FROM balanceclasificado_detalle "
        csql2.sql = csql2.sql + " where codigo='" + codigo + "' "
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        existeCAPITALDETALLE = True
        Else
        existeCAPITALDETALLE = False
        End If
        
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "select * FROM balanceclasificado_detalle "
        csql2.sql = csql2.sql + " where codigo='" + Mid(codigo, 1, 4) + "0000" + "' "
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        existeCAPITALDETALLE = True
        End If
        
        
        
End Function


Private Sub Grid2_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
ROW1 = NewRow
End Sub
Sub imprimir()
Dim titulo As String


titulo = "BALANCE CLASIFICADO AL " + Format(fechasistema, "dd-mm-yyyy")
Call CABEZAS2(titulo, "N", 1)
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellLandscape

Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick


Grid1.PageSetup.CenterHorizontally = True


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub

Public Function leersaldomayor2(codigo, MESES, codigo2, CRCC) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim fecha1 As String
Dim fecha2 As String
Dim NIVEL As String
Dim suma2 As Double
Dim LINEAS As Double

empresa = Mid(comboempresas.text, 1, 2)
Set csql.ActiveConnection = contadb
        
        NIVEL = "3"
        If Mid(codigo, 5, 5) = "0000" Then NIVEL = "2"
        If Mid(codigo, 3, 6) = "000000" Then NIVEL = "1"
        
        
        csql.sql = "select  "
        For k = 1 To MESES
        csql.sql = csql.sql + "datoa_" & k & ","
        
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
       If codigo2 = 35 Then
        csql.sql = csql.sql + " from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" & ipusada & "' and codigo='47100000' AND glosa<>'' and año='3' "
       End If
       If codigo2 <> 35 Then
        csql.sql = csql.sql + " from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" & ipusada & "' and codigo='" + codigo + "' and año='1' "
       End If
       
        csql.Execute
        leersaldomayor2 = 0
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LINEAS = 2
        For k = 1 To MESES
        If k = 1 Then monto = resultados(0)
        If k = 2 Then monto = resultados(1)
        If k = 3 Then monto = resultados(2)
        If k = 4 Then monto = resultados(3)
        If k = 5 Then monto = resultados(4)
        If k = 6 Then monto = resultados(5)
        If k = 7 Then monto = resultados(6)
        If k = 8 Then monto = resultados(7)
        If k = 9 Then monto = resultados(8)
        If k = 10 Then monto = resultados(9)
        If k = 11 Then monto = resultados(10)
        If k = 12 Then monto = resultados(11)
        
         If chksimula = 1 Then
                    If Option2.Value = True Then
                    PORCE10 = 1 + (CDbl(TXTVENTA.text) / 100)
                    Else
                    PORCE10 = 1 - (CDbl(TXTVENTA.text) / 100)
                    End If
                    If Mid(codigo, 1, 2) = "35" Then
                    monto = monto * PORCE10
                    End If
                    If Mid(codigo, 1, 4) = "4710" Then
                    monto = monto * PORCE10
                    End If
                
                
         End If
                
          
        If Mid(codigo, 1, 2) < "40" Then monto = monto * -1
        If IsNull(monto) = True Then monto = 0
        If codigo = "47150000" Then LINEAREMU = Grid1.Rows - 1
        Grid1.Cell(Grid1.Rows - 1, LINEAS).text = monto
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = "0"
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &HC0C0C0
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
                
        totales2(codigo2, LINEAS) = totales2(codigo2, LINEAS) + monto
        totales(LINEAS) = totales(LINEAS) + monto
        LINEAS = LINEAS + 3
        Next k
        
            resultados.Close
        Set resultados = Nothing
            
    End If
    If Mid(codigo, 1, 4) = "4780" Then
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = "0"
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &HC0C0C0
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
        
    End If
        
        If chkrenta = 1 And Mid(codigo, 1, 4) = "4780" Then
        LINEAS = 2
        For k = 1 To MESES
                    PORCE10 = (CDbl(TXTRENTA.text) / 100)
                    monto = Round(totales2(28, LINEAS) * PORCE10)
                    totales2(codigo2, LINEAS) = totales2(codigo2, LINEAS) + monto
                    totales(LINEAS) = totales(LINEAS) + monto
                    LINEAS = LINEAS + 3
                    
        Next k
                
         End If
     
    csql.Close
    Set csql = Nothing

End Function
Public Function leersaldomayor3(codigo, MESES, codigo2, CRCC) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim fecha1 As String
Dim fecha2 As String
Dim NIVEL As String
Dim suma2 As Double
Dim LINEAS As Double

empresa = Mid(comboempresas.text, 1, 2)
Set csql.ActiveConnection = contadb
        
        NIVEL = "3"
        If Mid(codigo, 5, 5) = "0000" Then NIVEL = "2"
        If Mid(codigo, 3, 6) = "000000" Then NIVEL = "1"
        
        
        csql.sql = "select  "
        For k = 1 To MESES
        csql.sql = csql.sql + "datoa_" & k & ","
        
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        If codigo2 = "35" Then
        csql.sql = csql.sql + " from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" & ipusada & "' and codigo='" + codigo + "' and glosa<>'' and año='4' "
        End If
        If codigo2 <> "35" Then
        csql.sql = csql.sql + " from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" & ipusada & "' and codigo='" + codigo + "'  and año='2' "
        End If
        
       
        csql.Execute
        leersaldomayor3 = 0
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LINEAS = 3
        For k = 1 To MESES
        If k = 1 Then monto = resultados(0)
        If k = 2 Then monto = resultados(1)
        If k = 3 Then monto = resultados(2)
        If k = 4 Then monto = resultados(3)
        If k = 5 Then monto = resultados(4)
        If k = 6 Then monto = resultados(5)
        If k = 7 Then monto = resultados(6)
        If k = 8 Then monto = resultados(7)
        If k = 9 Then monto = resultados(8)
        If k = 10 Then monto = resultados(9)
        If k = 11 Then monto = resultados(10)
        If k = 12 Then monto = resultados(11)
         
            
                
                If chksimula = 1 Then
                    If Option2.Value = True Then
                    PORCE10 = 1 + (CDbl(TXTVENTA.text) / 100)
                    Else
                    PORCE10 = 1 - (CDbl(TXTVENTA.text) / 100)
                    End If
                    If Mid(codigo, 1, 2) = "35" Then
                    monto = monto * PORCE10
                    End If
                    If Mid(codigo, 1, 4) = "4710" Then
                    monto = monto * PORCE10
                    End If
                
                
         End If
         
                
        
        If Mid(codigo, 1, 2) < "40" Then monto = monto * -1
        If IsNull(monto) = True Then monto = 0
        Grid1.Cell(Grid1.Rows - 1, LINEAS).text = monto
        totales2(codigo2, LINEAS) = totales2(codigo2, LINEAS) + monto
        totales(LINEAS) = totales(LINEAS) + monto
        LINEAS = LINEAS + 3
        Next k
        
            resultados.Close
        Set resultados = Nothing
            
    End If
    If Mid(codigo, 1, 4) = "4780" Then
    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = "0"
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &HC0C0C0
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
        
    End If
    
    If chkrenta = 1 And Mid(codigo, 1, 4) = "4780" Then
        LINEAS = 3
        For k = 1 To MESES
                       PORCE10 = (CDbl(TXTRENTA.text) / 100)
                    monto = Round(totales2(28, LINEAS) * PORCE10)
                     
                    totales2(codigo2, LINEAS) = totales2(codigo2, LINEAS) + monto
                    totales(LINEAS) = totales(LINEAS) + monto
                    LINEAS = LINEAS + 3
                    
        Next k
            
         End If
     
    csql.Close
    Set csql = Nothing

End Function
Public Function leersaldomayor20(codigo, MESES, codigo2, CRCC, codigo3, glosa) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim fecha1 As String
Dim fecha2 As String
Dim NIVEL As String
Dim suma2 As Double
Dim LINEAS As Double

empresa = Mid(comboempresas.text, 1, 2)
Set csql.ActiveConnection = contadb
        
        NIVEL = "3"
        If Mid(codigo, 5, 5) = "0000" Then NIVEL = "2"
        If Mid(codigo, 3, 6) = "000000" Then NIVEL = "1"
        
        
        csql.sql = "select  "
        For k = 1 To MESES
        csql.sql = csql.sql + "datoa_" & k & ","
        
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        If codigo3 <> "35" Then
        csql.sql = csql.sql + " from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" & ipusada & "' and codigo='" + codigo + "' and año='1' "
        End If
        If codigo3 = "35" Then
        csql.sql = csql.sql + " from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" & ipusada & "' and codigo='" + codigo + "' and año='3' AND glosa='" + glosa + "' "
        End If
        
       
        csql.Execute
        
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LINEAS = 2
        For k = 1 To MESES
        If k = 1 Then monto = resultados(0)
        If k = 2 Then monto = resultados(1)
        If k = 3 Then monto = resultados(2)
        If k = 4 Then monto = resultados(3)
        If k = 5 Then monto = resultados(4)
        If k = 6 Then monto = resultados(5)
        If k = 7 Then monto = resultados(6)
        If k = 8 Then monto = resultados(7)
        If k = 9 Then monto = resultados(8)
        If k = 10 Then monto = resultados(9)
        If k = 11 Then monto = resultados(10)
        If k = 12 Then monto = resultados(11)
            
        If chksimula = 1 Then
                    If Option2.Value = True Then
                    PORCE10 = 1 + (CDbl(TXTVENTA.text) / 100)
                    Else
                    PORCE10 = 1 - (CDbl(TXTVENTA.text) / 100)
                    End If
                    If Mid(codigo, 1, 2) = "35" Then
                    monto = monto * PORCE10
                    End If
                    If Mid(codigo, 1, 4) = "4710" Then
                    monto = monto * PORCE10
                    End If
                
                
         End If
         
        
        If Mid(codigo, 1, 2) < "40" Then monto = monto * -1
        If IsNull(monto) = True Then monto = 0
        Grid1.Cell(Grid1.Rows - 1, LINEAS).text = monto
        LINEAS = LINEAS + 3
        Next k
        
            resultados.Close
        Set resultados = Nothing
            
    End If
      
        If chkrenta = 1 And Mid(codigo, 1, 4) = "4780" Then
        LINEAS = 2
        For k = 1 To MESES
                    PORCE10 = (CDbl(TXTRENTA.text) / 100)
                    monto = Round(totales2(28, LINEAS) * PORCE10)
                    Grid1.Cell(Grid1.Rows - 1, LINEAS).text = monto
                    LINEAS = LINEAS + 3
                    
        Next k
                
         End If
    csql.Close
    Set csql = Nothing

End Function
Public Function leersaldomayor30(codigo, MESES, codigo2, CRCC, codigo3, glosa) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim fecha1 As String
Dim fecha2 As String
Dim NIVEL As String
Dim suma2 As Double
Dim LINEAS As Double

empresa = Mid(comboempresas.text, 1, 2)
Set csql.ActiveConnection = contadb
        
        NIVEL = "3"
        If Mid(codigo, 5, 5) = "0000" Then NIVEL = "2"
        If Mid(codigo, 3, 6) = "000000" Then NIVEL = "1"
        
        
        csql.sql = "select  "
        For k = 1 To MESES
        csql.sql = csql.sql + "datoa_" & k & ","
        
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        If codigo3 <> "35" Then
        csql.sql = csql.sql + " from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" & ipusada & "' and codigo='" + codigo + "' and año='2' "
        End If
        If codigo3 = "35" Then
        csql.sql = csql.sql + " from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" & ipusada & "' and codigo='" + codigo + "' and año='4' AND glosa='" + glosa + "' "
        End If
        
        
       
        csql.Execute
        
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LINEAS = 3
        For k = 1 To MESES
        If k = 1 Then monto = resultados(0)
        If k = 2 Then monto = resultados(1)
        If k = 3 Then monto = resultados(2)
        If k = 4 Then monto = resultados(3)
        If k = 5 Then monto = resultados(4)
        If k = 6 Then monto = resultados(5)
        If k = 7 Then monto = resultados(6)
        If k = 8 Then monto = resultados(7)
        If k = 9 Then monto = resultados(8)
        If k = 10 Then monto = resultados(9)
        If k = 11 Then monto = resultados(10)
        If k = 12 Then monto = resultados(11)
            
        If chksimula = 1 Then
                    If Option2.Value = True Then
                    PORCE10 = 1 + (CDbl(TXTVENTA.text) / 100)
                    Else
                    PORCE10 = 1 - (CDbl(TXTVENTA.text) / 100)
                    End If
                    If Mid(codigo, 1, 2) = "35" Then
                    monto = monto * PORCE10
                    End If
                    If Mid(codigo, 1, 4) = "4710" Then
                    monto = monto * PORCE10
                    End If
                
                
         End If
         
        
        If Mid(codigo, 1, 2) < "40" Then monto = monto * -1
        If IsNull(monto) = True Then monto = 0
        Grid1.Cell(Grid1.Rows - 1, LINEAS).text = monto
        LINEAS = LINEAS + 3
        Next k
        
            resultados.Close
        Set resultados = Nothing
            
    End If
        If chkrenta = 1 And Mid(codigo, 1, 4) = "4780" Then
        LINEAS = 3
        For k = 1 To MESES
                    PORCE10 = (CDbl(TXTRENTA.text) / 100)
                    monto = Round(totales2(28, LINEAS) * PORCE10)
                    Grid1.Cell(Grid1.Rows - 1, LINEAS).text = monto
                    LINEAS = LINEAS + 3
                    
        Next k
                
         End If
    csql.Close
    Set csql = Nothing

End Function

Sub CARGAcrcc(empresa2)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    empresa = empresa2
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM " + clientesistema + "conta" + empresa + ".centrosdecosto where año='" + Format(fechasistema, "YYYY") + "' "
        csql.sql = csql.sql + "order by codigo"
        csql.Execute
        LINEA = 0
Combocrcc.Clear

        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             Combocrcc.AddItem (Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + " " + resultados(1))
             
            resultados.MoveNext
            Wend
        End If
        Combocrcc.AddItem ("99.99" + " " + "TODOS")
            
        Combocrcc.text = Combocrcc.List(LINEA)
        
   

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
        End If
        ComboLOCAL.AddItem ("99 todos")
        ComboLOCAL.text = ComboLOCAL.List(ComboLOCAL.ListIndex - 1)
        
        
End Sub

Public Function cargadatos(CRCC, empresa2) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim fecha1 As String
Dim fecha2 As String
Dim NIVEL As String
Dim suma2 As Double
Dim LINEAS As Double
empresa = empresa2

Set csql.ActiveConnection = contadb
        
Rem         insert into estado_resultado
Rem select '" & ipusada & "',codigocuenta,'2',sum(if(fecha like '%2012-03%',if(mo.dh='D',monto,monto*-1),0)),sum(if(fecha like '%2012-04%',if(mo.dh='D',monto,monto*-1),0)),sum(if(fecha like '%2012-05%',if(mo.dh='D',monto,monto*-1),0)),0,0,0,0,0,0,0,0,0 from movimientoscontables as mo where fecha >= '%2011%' and codigocuenta>'30000000' group by codigocuenta;

Rem insert into estado_resultado
Rem select '" & ipusada & "',codigocuenta,'1',sum(if(fecha like '%2011-03%',if(mo.dh='D',monto,monto*-1),0)),sum(if(fecha like '%2011-04%',if(mo.dh='D',monto,monto*-1),0)),sum(if(fecha like '%2011-05%',if(mo.dh='D',monto,monto*-1),0)),0,0,0,0,0,0,0,0,0 from movimientoscontables as mo where fecha >= '%2011%' and codigocuenta>'30000000' group by codigocuenta;

Rem insert into estado_resultado
Rem select '" & ipusada & "',concat(mid(codigocuenta,1,4),'0000'),'2',sum(if(fecha like '%2012-03%',if(mo.dh='D',monto,monto*-1),0)),sum(if(fecha like '%2012-04%',if(mo.dh='D',monto,monto*-1),0)),sum(if(fecha like '%2012-05%',if(mo.dh='D',monto,monto*-1),0)),0,0,0,0,0,0,0,0,0 from movimientoscontables as mo where fecha >= '%2011%' and codigocuenta>'30000000' group by mid(codigocuenta,1,4);

Rem insert into estado_resultado
Rem select '" & ipusada & "',concat(mid(codigocuenta,1,4),'0000'),'1',sum(if(fecha like '%2011-03%',if(mo.dh='D',monto,monto*-1),0)),sum(if(fecha like '%2011-04%',if(mo.dh='D',monto,monto*-1),0)),sum(if(fecha like '%2011-05%',if(mo.dh='D',monto,monto*-1),0)),0,0,0,0,0,0,0,0,0 from movimientoscontables as mo where fecha >= '%2011%' and codigocuenta>'30000000' group by mid(codigocuenta,1,4);
MESES = txtmes.text
    Rem elimina
        csql.sql = "delete from " + clientesistema + "conta" + empresa + ".estado_resultado where ip='" + ipusada + "'; "
        csql.Execute
    Rem carga venta 1 año detalle
        csql.sql = "insert into " + clientesistema + "conta" + empresa + ".estado_resultado  "
        csql.sql = csql.sql + "select  '" + ipusada + "',codigocuenta,'1', "
        LINEAS = 2
        For k = 1 To MESES
        csql.sql = csql.sql + "sum(if(fecha like '%" & Format(Grid1.Cell(0, LINEAS).text, "yyyy-mm") & "%',if(mo.dh='D',monto,monto*-1),0)),"
        LINEAS = LINEAS + 3
        Next k
        For k = MESES + 1 To 12
        csql.sql = csql.sql + "0,"
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        csql.sql = csql.sql + ",'' from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo where fecha >= '%" & Grid1.Cell(0, 2).text & "%' and codigocuenta>'30000000'  "
        If Mid(CRCC, 1, 2) <> "99" Then
        csql.sql = csql.sql + "and centrocosto='" + CRCC + "' "
        End If
        csql.sql = csql.sql + "group by codigocuenta "
        
       
        csql.Execute
        Rem solo  1 año detalle  costo venta
        
       
        csql.sql = "insert into " + clientesistema + "conta" + empresa + ".estado_resultado  "
        csql.sql = csql.sql + "select  '" + ipusada + "',codigocuenta,'3', "
        LINEAS = 2
        For k = 1 To MESES
        csql.sql = csql.sql + "sum(if(fecha like '%" & Format(Grid1.Cell(0, LINEAS).text, "yyyy-mm") & "%',if(mo.dh='D',monto,monto*-1),0)),"
        LINEAS = LINEAS + 3
        Next k
        For k = MESES + 1 To 12
        csql.sql = csql.sql + "0,"
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        csql.sql = csql.sql + ",replace(glosacontable,'CENTRALIZACION','') from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo where fecha >= '%" & Grid1.Cell(0, 2).text & "%' and codigocuenta='47100001' and tipo='MI' "
        If Mid(CRCC, 1, 2) <> "99" Then
        csql.sql = csql.sql + "and centrocosto='" + CRCC + "' "
        End If
        csql.sql = csql.sql + "group by codigocuenta,glosacontable "
        
       
        csql.Execute
        
        Rem carga vta 1 año sumado
        csql.sql = "insert into " + clientesistema + "conta" + empresa + ".estado_resultado  "
        csql.sql = csql.sql + "select  '" + ipusada + "',concat(mid(codigocuenta,1,4),'0000'),'1', "
        LINEAS = 2
        For k = 1 To MESES
        csql.sql = csql.sql + "sum(if(fecha like '%" & Format(Grid1.Cell(0, LINEAS).text, "yyyy-mm") & "%',if(mo.dh='D',monto,monto*-1),0)),"
        LINEAS = LINEAS + 3
        Next k
        For k = MESES + 1 To 12
        csql.sql = csql.sql + "0,"
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        csql.sql = csql.sql + ",'' from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo where fecha >= '%" & Grid1.Cell(0, 2).text & "%' and codigocuenta>'30000000' "
        If Mid(CRCC, 1, 2) <> "99" Then
        csql.sql = csql.sql + "and centrocosto='" + CRCC + "' "
        End If
        csql.sql = csql.sql + "group by mid(codigocuenta,1,4) "
        
        
        csql.Execute
        Rem carga vta 1 año sumado solo costo
        csql.sql = "insert into " + clientesistema + "conta" + empresa + ".estado_resultado  "
        csql.sql = csql.sql + "select  '" + ipusada + "',concat(mid(codigocuenta,1,4),'0000'),'3', "
        LINEAS = 2
        For k = 1 To MESES
        csql.sql = csql.sql + "sum(if(fecha like '%" & Format(Grid1.Cell(0, LINEAS).text, "yyyy-mm") & "%',if(mo.dh='D',monto,monto*-1),0)),"
        LINEAS = LINEAS + 3
        Next k
        For k = MESES + 1 To 12
        csql.sql = csql.sql + "0,"
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        csql.sql = csql.sql + ",replace(glosacontable,'CENTRALIZACION','') from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo where fecha >= '%" & Grid1.Cell(0, 2).text & "%' and codigocuenta='47100001' and tipo='MI' "
        If Mid(CRCC, 1, 2) <> "99" Then
        csql.sql = csql.sql + "and centrocosto='" + CRCC + "' "
        End If
        csql.sql = csql.sql + "group by mid(codigocuenta,1,4) "
        
        
        csql.Execute
        
        
        
    Rem carga venta 2 año detalle
        
        
        csql.sql = "insert into " + clientesistema + "conta" + empresa + ".estado_resultado  "
        csql.sql = csql.sql + "select  '" + ipusada + "',codigocuenta,'2', "
        LINEAS = 3
        For k = 1 To MESES
        csql.sql = csql.sql + "sum(if(fecha like '%" & Format(Grid1.Cell(0, LINEAS).text, "yyyy-mm") & "%',if(mo.dh='D',monto,monto*-1),0)),"
        LINEAS = LINEAS + 3
        Next k
        For k = MESES + 1 To 12
        csql.sql = csql.sql + "0,"
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        csql.sql = csql.sql + ",'' from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo where fecha >= '%" & Grid1.Cell(0, 2).text & "%' and codigocuenta>'30000000' "
        If Mid(CRCC, 1, 2) <> "99" Then
        csql.sql = csql.sql + "and centrocosto='" + CRCC + "' "
        End If
        csql.sql = csql.sql + "group by codigocuenta "
        csql.Execute
        
        
        
        
        Rem carga solo gastos
        
        csql.sql = "insert into " + clientesistema + "conta" + empresa + ".estado_resultado  "
        csql.sql = csql.sql + "select  '" + ipusada + "',codigocuenta,'4', "
        LINEAS = 3
        For k = 1 To MESES
        csql.sql = csql.sql + "sum(if(fecha like '%" & Format(Grid1.Cell(0, LINEAS).text, "yyyy-mm") & "%',if(mo.dh='D',monto,monto*-1),0)),"
        LINEAS = LINEAS + 3
        Next k
        For k = MESES + 1 To 12
        csql.sql = csql.sql + "0,"
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        csql.sql = csql.sql + ",replace(glosacontable,'CENTRALIZACION','') from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo where fecha >= '%" & Grid1.Cell(0, 2).text & "%' and codigocuenta='47100001' and tipo='MI' "
        If Mid(CRCC, 1, 2) <> "99" Then
        csql.sql = csql.sql + "and centrocosto='" + CRCC + "' "
        End If
        csql.sql = csql.sql + "group by codigocuenta,glosacontable "
        csql.Execute
        
        
        
        Rem carga vta 2 año sumado
        csql.sql = "insert into " + clientesistema + "conta" + empresa + ".estado_resultado  "
        csql.sql = csql.sql + "select  '" + ipusada + "',concat(mid(codigocuenta,1,4),'0000'),'2', "
        LINEAS = 3
        For k = 1 To MESES
        csql.sql = csql.sql + "sum(if(fecha like '%" & Format(Grid1.Cell(0, LINEAS).text, "yyyy-mm") & "%',if(mo.dh='D',monto,monto*-1),0)),"
        LINEAS = LINEAS + 3
        Next k
        For k = MESES + 1 To 12
        csql.sql = csql.sql + "0,"
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        csql.sql = csql.sql + ",'' from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo where fecha >= '%" & Grid1.Cell(0, 2).text & "%' and codigocuenta>'30000000' "
        If Mid(CRCC, 1, 2) <> "99" Then
        csql.sql = csql.sql + "and centrocosto='" + CRCC + "' "
        End If
        csql.sql = csql.sql + "group by mid(codigocuenta,1,4) "
        
        csql.Execute
    
    Rem carga vta 2 año sumado solo costo
        csql.sql = "insert into " + clientesistema + "conta" + empresa + ".estado_resultado  "
        csql.sql = csql.sql + "select  '" + ipusada + "',concat(mid(codigocuenta,1,4),'0000'),'4', "
        LINEAS = 3
        For k = 1 To MESES
        csql.sql = csql.sql + "sum(if(fecha like '%" & Format(Grid1.Cell(0, LINEAS).text, "yyyy-mm") & "%',if(mo.dh='D',monto,monto*-1),0)),"
        LINEAS = LINEAS + 3
        Next k
        For k = MESES + 1 To 12
        csql.sql = csql.sql + "0,"
        Next k
        csql.sql = Mid(csql.sql, 1, Len(csql.sql) - 1)
        csql.sql = csql.sql + ",replace(glosacontable,'CENTRALIZACION','') from " + clientesistema + "conta" + empresa + ".movimientoscontables as mo where fecha >= '%" & Grid1.Cell(0, 2).text & "%' and codigocuenta='47100001' and tipo='MI' "
        If Mid(CRCC, 1, 2) <> "99" Then
        csql.sql = csql.sql + "and centrocosto='" + CRCC + "' "
        End If
        csql.sql = csql.sql + "group by mid(codigocuenta,1,4) "
        
        csql.Execute
    
    
    
    csql.Close
    Set csql = Nothing

End Function

Private Sub Grid1_DblClick()

If Val(Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text) <> 0 Then
    If Grid1.Cell(0, Grid1.ActiveCell.col).text <> "%" Then
        If Grid1.ActiveCell.row > 2 Then
            If Val(Mid(Grid1.Cell(Grid1.ActiveCell.row, 0).text, 5, 4)) <> 0 Then
            informa04.cmdato1.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 0).text, 1, 2)
            informa04.cmdato2.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 0).text, 3, 2)
            informa04.cmdato3.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 0).text, 5, 4)
            informa04.desdefecha.Caption = "01" + "-" + Format(Grid1.Cell(0, Grid1.ActiveCell.col).text, "mm-yyyy")
            informa04.cmnombre.Caption = Grid1.Cell(Grid1.ActiveCell.row, 1).text
            informa04.hastafecha.Caption = "31" + "-" + Format(Grid1.Cell(0, Grid1.ActiveCell.col).text, "mm-yyyy")
            informa04.frm_crcc.Visible = False
                If Mid(Combocrcc.text, 1, 5) <> "99.99" Then
                    informa04.txt_crcc.text = Mid(Combocrcc.text, 1, 2) + Mid(Combocrcc.text, 4, 2)
                    informa04.lbl_crcc.Caption = Mid(Combocrcc.text, 7, 10)
                    informa04.frm_crcc.Visible = True
                End If
            informa04.Show
            End If
        End If
    End If
End If

End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyO Then
        If Val(Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text) <> 0 Then
            If Grid1.Cell(0, Grid1.ActiveCell.col).text <> "%" Then
                If IsDate(Grid1.Cell(0, Grid1.ActiveCell.col).text) = True Then
                    Load GLOSAEERR
                    GLOSAEERR.lblcentrocosto.Caption = Mid(Combocrcc.text, 1, 5)
                    GLOSAEERR.LBLEMPRESA.Caption = Mid(comboempresas.text, 1, 2)
                    GLOSAEERR.lblNOMBRE.Caption = Grid1.Cell(Grid1.ActiveCell.row, 1).text
                    GLOSAEERR.lblfecha.Caption = Grid1.Cell(0, Grid1.ActiveCell.col).text
                    GLOSAEERR.leer
                    GLOSAEERR.Show vbModal
                    
                End If
            End If
        End If
    End If

End Sub

Private Sub txtmes_KeyPress(KeyAscii As Integer)
Call esNumero(KeyAscii)

End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Sub CARGAempresas()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigoempresa,nombre "
        csql.sql = csql.sql + "FROM maestroempresas "
        csql.sql = csql.sql + "order by codigoempresa"
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             comboempresas.AddItem (Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + " " + resultados(1))
             
            resultados.MoveNext
            Wend
        End If
        comboempresas.AddItem ("99.99" + " " + "TODOS")
            
        comboempresas.text = empresaactiva + " " + leerempresa(empresaactiva)
        
        
Call CARGAcrcc(Mid(comboempresas.text, 1, 2))

   

End Sub

Private Sub TXTVENTA_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

End Sub
Public Function SUMAMTS(CRCC)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT sum(mts) "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion.g_maestroempresas where codigocontable='" + Mid(comboempresas.text, 1, 2) + "' "
        If CRCC <> "99.99" Then
        csql.sql = csql.sql + "and codigocrcc='" + Mid(CRCC, 1, 2) + Mid(CRCC, 4, 2) + "' "
        End If
        
        csql.sql = csql.sql + "group by codigocontable "
        csql.Execute
        LINEA = 0
        SUMAMTS = 0
        If csql.RowsAffected > 0 Then
            
            Set resultados = csql.OpenResultset
        SUMAMTS = resultados(0)
        End If
        
   

End Function

