VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "clbutn.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form grilla1847 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9735
   ClientLeft      =   645
   ClientTop       =   1110
   ClientWidth     =   14925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp CABEZA 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   18653
      BackColor       =   16761024
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
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
      ColorTextShadow =   16711680
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   855
         Left            =   90
         TabIndex        =   2
         Top             =   8775
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   1508
         BackColor       =   16777152
         Caption         =   "OPCIONES"
         CaptionEstilo3D =   1
         BackColor       =   16777152
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
         Begin CoolButtons.cool_Button command1 
            Height          =   495
            Left            =   3015
            TabIndex        =   3
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "Imprimir"
         End
         Begin CoolButtons.cool_Button COMMAND2 
            Height          =   495
            Left            =   5805
            TabIndex        =   4
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "Exportar Excel"
         End
         Begin CoolButtons.cool_Button command4 
            Height          =   495
            Left            =   8400
            TabIndex        =   5
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "Salir"
         End
         Begin MSComctlLib.Slider TAMAÑOS 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
         End
         Begin CoolButtons.cool_Button cmd_xml 
            Height          =   495
            Left            =   11280
            TabIndex        =   9
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "Genera UPLOAD SII"
         End
         Begin VB.Label registros 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   10800
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label LETRA 
            Height          =   255
            Left            =   720
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   8415
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14843
         Cols            =   5
         DefaultFontSize =   8.25
         DefaultRowHeight=   15
         Rows            =   30
      End
      Begin VB.Label titulofinal 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "grilla1847"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 

Private Sub cmd_xml_Click()
Dim LINEA As String
Dim s As Integer

Close 20
Open "C:\SII\f1847" + empresaactiva + ".txt" For Output As #20

LINEA = "1;" + codigoae + ";NO APLICA;;" & FOLIOini2 & ";" & foliofin2 & ";;;;;;;;;;;"
Print #20, LINEA
For k = 1 To Grid1.Rows - 9
LINEA = "2;;;;;"
For s = 1 To 11
LINEA = LINEA + ";" + Grid1.Cell(k, s).text
Next s
LINEA = Replace(LINEA, "¢", "O")

Print #20, LINEA
Next k
Close 20

Shell "notepad C:\SII\f1847" + empresaactiva + ".txt"


End Sub

Private Sub Command1_Click()
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
    Grid1.PageSetup.PaperWidth = 21.59
    Grid1.PageSetup.PaperHeight = 27.94
    
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar01" Then Call imprime_balancetributario(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar02" Then Call imprime_balanceanalitico(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar03" Then Call imprime_mayoranalitico(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10), titulofinal.Caption)
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar04" Then Call imprime_librodiario(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10), titulofinal.Caption)
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar05" Then Call imprime_librocompras(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar06" Then Call imprime_librohonorarios(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "publi0004" Then Call imprime_publicidad(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar44" Then Call imprime_libroventas(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "auxiliar07" Then Call imprime_libroboletas(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 7) = "banco02" Then Call imprime_cartolamayor("N", Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 12) = "CARTOLAMAYOR" Then Call imprime_cartolamayor("N", Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 13) = "CARTOLACTACTE" Then Call imprime_cartolaCTACTE("N", Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 11) = "informa01_1" Then Call imprime_NORMALES("PLAN DE CUENTAS")
    If Mid(grillainformes.Tag, 1, 11) = "informa01_2" Then Call imprime_NORMALES("CUENTAS CORRIENTES")
    If Mid(grillainformes.Tag, 1, 11) = "informa01_3" Then Call imprime_NORMALES("CENTROS DE COSTO")
    If Mid(grillainformes.Tag, 1, 8) = "infoge01" Then Call imprime_estadoresultado(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 8) = "infoge02" Then Call imprime_facturasporpagar(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 8) = "infoge03" Then Call imprime_honorariosporpagar(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 8) = "infoge04" Then Call imprime_ventasporpagar(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "control03" Then Call imprime_buscapormonto(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "control01" Then Call imprime_descuadrados(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "control04" Then Call imprime_buscacuentaseliminadas(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 10) = "INFOHARINA" Then Call imprime_harina(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 9) = "INFOCARNE" Then Call imprime_CARNE(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    If Mid(grillainformes.Tag, 1, 7) = "PRESU04" Then Call imprime_estadoresultado(Mid(grillainformes.Tag, 11, 1), Mid(grillainformes.Tag, 12, 10))
    
    Unload Me
End Sub

Private Sub Command2_Click()
Grid1.ExportToExcel (""), True



End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_Load()


Call CENTRAR(Me)
LETRA.Caption = Grid1.DefaultFont.Name
TAMAÑOS.Min = Grid1.DefaultFont.Size - 5
TAMAÑOS.Max = Grid1.DefaultFont.Size + 10
TAMAÑOS.Value = Grid1.DefaultFont.Size
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Locked = True


Next k
If xmlcompra = True Then
cmd_xml.Visible = True

End If
If xmlventa = True Then
cmd_xml.Visible = True

End If



End Sub

Sub imprime_balancetributario(Tipo, FOLIO)
Dim titulo As String
Dim subtitulo As String

subtitulo = Mid(CABEZA.Caption, 20, ((InStr(CABEZA.Caption, "empresa") - 6) - 20))
titulo = "BALANCE TRIBUTARIO"
Call cabezas(titulo, Tipo, FOLIO, subtitulo)
Grid1.DefaultFont.Size = 7

If Grid1.Column(1).Width <> "0" Then
    Grid1.PageSetup.Orientation = cellLandscape
'    For k = 1 To 11 - 1
'        Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
'    Next k
Else
    Grid1.PageSetup.Orientation = cellPortrait
    Grid1.Column(1).Width = 0
    For k = 2 To 11 - 1
        Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
    Next k
End If

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PrintPreview 120

End Sub
Sub imprime_balanceanalitico(Tipo, FOLIO)
Dim titulo As String
titulo = "BALANCE ANALITICO"
Call cabezas(titulo, Tipo, FOLIO, "")
'grid1.DefaultFont.Size = 6
Grid1.Column(2).Width = 20 * 8
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh

Grid1.PrintPreview 120

End Sub
Sub imprime_mayoranalitico(Tipo, FOLIO, TITULOalfinal)
Dim titulo As String
titulo = "MAYOR ANALITICO"
titulo = TITULOalfinal
Grid1.DefaultFont.Size = 8
For k = 1 To 10 - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 8
Next k
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh

Grid1.PrintPreview 120

End Sub
Sub imprime_cartolamayor(Tipo, FOLIO)

Dim titulo As String
titulo = "CARTOLA DEL MAYOR"
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
For k = 1 To 15 - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
Next k
Grid1.Column(6).Width = 200
Grid1.Column(5).Width = 0

Grid1.Column(14).Width = 0
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_cartolaCTACTE(Tipo, FOLIO)
Dim titulo As String
titulo = "CARTOLA DEL CUENTAS CORRIENTES"
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
For k = 1 To 10 - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
Next k

Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh

Grid1.PrintPreview 120

End Sub



Sub imprime_librodiario(Tipo, FOLIO, titulocabeza)
Dim titulo As String
titulo = "LIBRO DIARIO"
titulo = titulocabeza
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0

For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 8
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub

Sub imprime_librocompras(Tipo, FOLIO)
Dim titulo As String
titulo = "LIBRO DE COMPRAS " + auxiliar05.COMBOMES.text + " de " + auxiliar05.COMBOAÑO.text


Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape


If Tipo <> "N" Then
Grid1.PageSetup.Orientation = cellPortrait

Grid1.Cols = 13
End If
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_publicidad(Tipo, FOLIO)
Dim titulo As String
titulo = "LIBRO DE COMPRAS a PROVEEDORES PERIODO " + publi0004.desdefecha.Caption + " HASTA " + publi0004.hastafecha.Caption

Tipo = "N"
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub

Sub imprime_libroventas(Tipo, FOLIO)
Dim titulo As String
titulo = "LIBRO DE VENTAS " + auxiliar44.Combocrcc.text + " " + auxiliar44.COMBOMES.text + " de " + auxiliar44.COMBOAÑO.text

Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 7
Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_libroboletas(Tipo, FOLIO)
Dim titulo As String
titulo = "LIBRO DE BOLETAS " + auxiliar07.Combocrcc.text + " " + auxiliar07.COMBOMES.text + " de " + auxiliar07.COMBOAÑO.text

Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 7
Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_facturasporpagar(Tipo, FOLIO)
Dim titulo As String
titulo = "FACTURAS POR PAGAR"
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_estadoresultado(Tipo, FOLIO)
Dim titulo As String
titulo = "ESTADO DE RESULTADO "
Call cabezas(titulo, "N", FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

 For i = 1 To Grid1.PageSetup.PaperSizes.Count
            If UCase(Grid1.PageSetup.PaperSizes.Item(i).PaperName) = "OFICIO" Then
                Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 0
Grid1.PageSetup.TopMargin = 0.5
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k
Grid1.SelectionMode = cellSelectionFree


Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_honorariosporpagar(Tipo, FOLIO)
Dim titulo As String
titulo = "HONORARIOS POR PAGAR"
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_ventasporpagar(Tipo, FOLIO)
Dim titulo As String
titulo = "VENTAS POR PAGAR"
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_buscapormonto(Tipo, FOLIO)
Dim titulo As String
titulo = "BUSCA POR MONTO "
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_buscacuentaseliminadas(Tipo, FOLIO)
Dim titulo As String
titulo = "LISTADO DE CUENTAS ELIMINADAS "
Call cabezas(titulo, Tipo, FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellPortrait

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub imprime_librohonorarios(Tipo, FOLIO)
    Dim titulo As String
    titulo = "LIBRO DE HONORARIOS " + auxiliar06.COMBOMES.text + " de " + auxiliar06.COMBOAÑO.text
    Call cabezas(titulo, Tipo, FOLIO, "")
    Grid1.DefaultFont.Size = 6
    Grid1.PageSetup.Orientation = cellPortrait
    Grid1.PageSetup.PrintFixedRow = True
    Grid1.PageSetup.BottomMargin = 2
    Grid1.PageSetup.TopMargin = 1
    Grid1.PageSetup.LeftMargin = 0.5
    Grid1.PageSetup.RightMargin = 0
    
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
    Next k
    Grid1.DisablePrintButton = False
    Grid1.Refresh
    Grid1.PrintPreview 120
   
End Sub


Sub cabezas(titulo, Tipo, FOLIO, subtitulo)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    If subtitulo <> "" Then
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = UCase(subtitulo)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 9
        objReportTitle.Font.Bold = True
        objReportTitle.PrintOnAllPages = True
        Grid1.ReportTitles.Add objReportTitle
    End If
    
    'Report Title 1
    
    If Tipo = "N" Then
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
        
        If Tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload informa04

End Sub

Private Sub Grid1_DblClick()
Dim dia As String

If Mid(CABEZA.Caption, 1, 7) = "CARTOLA" Then
da0 = Grid1.Cell(Grid1.ActiveCell.row, 2).text
da1 = Grid1.Cell(Grid1.ActiveCell.row, 3).text
da2 = Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "dd")
da3 = Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "mm")
da4 = Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "yyyy")
muestracomprobantes.Show vbModal
End If
If Mid(grillainformes.Tag, 1, 7) = "PRESU04" Then

End If
If Mid(CABEZA.Caption, 1, 6) = "ESTADO" Then
Load informa04

informa04.cmdato1.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 2)
informa04.cmdato2.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 3, 2)
informa04.cmdato3.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 5, 4)
informa04.desdefecha.Caption = "01" + "-" + Format(Grid1.ActiveCell.col - 2, "00") + "-" + Format(fechasistema, "yyyy")
informa04.cmnombre.Caption = Grid1.Cell(Grid1.ActiveCell.row, 2).text
dia = Format(DateSerial(Format(fechasistema, "yyyy"), (Grid1.ActiveCell.col - 2) + 1, 0), "dd")
informa04.hastafecha.Caption = dia + "-" + Format(Grid1.ActiveCell.col - 2, "00") + "-" + Format(fechasistema, "yyyy")


informa04.Show



End If

End Sub

Private Sub TAMAÑOS_Click()

Grid1.DefaultFont.Size = TAMAÑOS.Value
'For K = 1 To Grid1.Cols - 1
'Grid1.Column(K).Width = Len(Grid1.Cell(2, K).text) * TAMAÑOS.Value


'Next K



Grid1.Refresh

End Sub
Sub imprime_NORMALES(Titulos)
Dim titulo As String
titulo = Titulos
Call cabezas(titulo, "N", "000000000", "")
'grid1.DefaultFont.Size = 6

Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.Refresh

Grid1.PrintPreview 120

End Sub

Sub imprime_harina(Tipo, FOLIO)
Dim titulo As String
Dim titulo2 As String

titulo = "ANEXO INFORME MENSUAL VENDEDORES DE HARINA "
titulo2 = "INFORMACION DEL MES DE " + infoharina.COMBOMES.text + " AÑO " + infoharina.COMBOAÑO.text



Call CABEZAS2(titulo, titulo2)
Grid1.DefaultFont.Size = 7
Grid1.PageSetup.Orientation = cellPortrait



Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub
Sub imprime_CARNE(Tipo, FOLIO)
Dim titulo As String
Dim titulo2 As String

titulo = "ANEXO INFORME MENSUAL RETENCION CARNE "
titulo2 = "INFORMACION DEL MES DE " + infocarne.COMBOMES.text + " AÑO " + infocarne.COMBOAÑO.text



Call CABEZAS2(titulo, titulo2)
Grid1.DefaultFont.Size = 7
Grid1.PageSetup.Orientation = cellPortrait



Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


Sub CABEZAS2(titulo, titulo2)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo2
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
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
        Grid1.ReportTitles.Add objReportTitle
    Next k
    
With Grid1.PageSetup
        
        If Tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub
Sub imprime_descuadrados(Tipo, FOLIO)
Dim titulo As String
titulo = "LISTADOS DESCUADRADOS "
Call cabezas(titulo, "N", FOLIO, "")
Grid1.DefaultFont.Size = 6
Grid1.PageSetup.Orientation = cellPortrait


 For i = 1 To Grid1.PageSetup.PaperSizes.Count
            If UCase(Grid1.PageSetup.PaperSizes.Item(i).PaperName) = "OFICIO" Then
                Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 0
Grid1.PageSetup.TopMargin = 0.5
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0
For k = 1 To Grid1.Cols - 1
Grid1.Column(k).Width = Grid1.Column(k).Width / 7 * 6
Next k

Grid1.Refresh
Grid1.PrintPreview 120
End Sub


