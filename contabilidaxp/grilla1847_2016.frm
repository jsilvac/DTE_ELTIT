VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "clbutn.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form grilla1847_2016 
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
      BackColor       =   16384
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16384
      ForeColor       =   65535
      ColorBarraArriba=   33023
      ColorBarraAbajo =   8438015
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
         Begin FlexCell.Grid Grid3 
            Height          =   255
            Left            =   1800
            TabIndex        =   11
            Top             =   360
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
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
Attribute VB_Name = "grilla1847_2016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 

Private Sub cmd_xml_Click()
Dim LINEA As String
Dim s As Integer
On Error GoTo err:

If Format(fechasistema, "yyyy") < "2019" Then


        Close 20
        Open "C:\SII\f1847_2017_" + empresaactiva + ".txt" For Output As #20
        
        'LINEA = "1;" + codigoae + ";NO APLICA;;" & FOLIOini2 & ";" & foliofin2 & ";2;;;;;;;;;;;;;;"
        LINEA = "1;" + codigoae + ";NO APLICA;;" & FOLIOini2 & ";" & foliofin2 & ";2;;;;;;;;;;;;;"
        Print #20, LINEA
        For k = 1 To GRID1.rows - 9
        If GRID1.Cell(k, 2).text = "TOTALES" Then GoTo PASO:
        LINEA = "2;;;;;;"
        For s = 1 To 13
        LINEA = LINEA + ";" + GRID1.Cell(k, s).text
        Next s
        LINEA = Replace(LINEA, "¢", "O")
        
        Print #20, LINEA
        Next k
PASO:
        Close 20
        
        Shell "notepad C:\SII\f1847_2017_" + empresaactiva + ".txt"
        Exit Sub
err:
        MsgBox "debe crear la carpeta c:\sii"

Else

' para exportar a csv

        Grid3.cols = 22
        Grid3.rows = 1

        Grid3.rows = Grid3.rows + 1
        Grid3.Cell(Grid3.rows - 1, 1).text = "1"
        Grid3.Cell(Grid3.rows - 1, 2).text = "1"
        Grid3.Cell(Grid3.rows - 1, 3).text = codigoae
        Grid3.Cell(Grid3.rows - 1, 4).text = "NO APLICA"
        Grid3.Cell(Grid3.rows - 1, 5).text = ""
        Grid3.Cell(Grid3.rows - 1, 6).text = FOLIOini2
        Grid3.Cell(Grid3.rows - 1, 7).text = foliofin2
        Grid3.Cell(Grid3.rows - 1, 8).text = "2"
        Grid3.Cell(Grid3.rows - 1, 9).text = ""
        Grid3.Cell(Grid3.rows - 1, 10).text = ""
        Grid3.Cell(Grid3.rows - 1, 11).text = ""
        Grid3.Cell(Grid3.rows - 1, 12).text = ""
        Grid3.Cell(Grid3.rows - 1, 13).text = ""
        Grid3.Cell(Grid3.rows - 1, 14).text = ""
        Grid3.Cell(Grid3.rows - 1, 15).text = ""
        Grid3.Cell(Grid3.rows - 1, 16).text = ""
        Grid3.Cell(Grid3.rows - 1, 17).text = ""
        Grid3.Cell(Grid3.rows - 1, 18).text = ""
        Grid3.Cell(Grid3.rows - 1, 19).text = ""
        Grid3.Cell(Grid3.rows - 1, 20).text = ""
        Grid3.Cell(Grid3.rows - 1, 21).text = ""
        
        
        
        
For k = 1 To GRID1.rows - 9
    If GRID1.Cell(k, 2).text = "TOTALES" Then GoTo paso1:
    D1 = k
    D2 = "2"
    D3 = ""
    D4 = ""
    D5 = ""
    D6 = ""
    D7 = ""
    D8 = ""
 

    D9 = GRID1.Cell(k, 1).text
    D10 = GRID1.Cell(k, 2).text
    D11 = Format(GRID1.Cell(k, 3).text, "")
    D12 = Format(GRID1.Cell(k, 4).text, "")
    D13 = Format(GRID1.Cell(k, 5).text, "")
    D14 = Format(GRID1.Cell(k, 6).text, "")
    D15 = Format(GRID1.Cell(k, 7).text, "")
    D16 = Format(GRID1.Cell(k, 8).text, "")
    D17 = Format(GRID1.Cell(k, 9).text, "")
    D18 = Format(GRID1.Cell(k, 10).text, "")
    D19 = Format(GRID1.Cell(k, 11).text, "")
    D20 = Format(GRID1.Cell(k, 12).text, "")
    D21 = Format(GRID1.Cell(k, 13).text, "")
 
    If D10 = "1.03.05.00" Then D21 = "0"
'    If Mid(D10, 1, 1) = "2" Then Stop
    If Mid(D10, 1, 4) = "2.03" Then D21 = "0"
    If Mid(D10, 1, 1) = "3" Then D21 = "0"
    
    
        
       Grid3.rows = Grid3.rows + 1
        Grid3.Cell(Grid3.rows - 1, 1).text = D1
        Grid3.Cell(Grid3.rows - 1, 2).text = D2
        Grid3.Cell(Grid3.rows - 1, 3).text = D3
        Grid3.Cell(Grid3.rows - 1, 4).text = D4
        Grid3.Cell(Grid3.rows - 1, 5).text = D5
        Grid3.Cell(Grid3.rows - 1, 6).text = D6
        Grid3.Cell(Grid3.rows - 1, 7).text = D7
        Grid3.Cell(Grid3.rows - 1, 8).text = D8
        Grid3.Cell(Grid3.rows - 1, 9).text = D9
        Grid3.Cell(Grid3.rows - 1, 10).text = D10
        Grid3.Cell(Grid3.rows - 1, 11).text = D11
        Grid3.Cell(Grid3.rows - 1, 12).text = D12
        Grid3.Cell(Grid3.rows - 1, 13).text = D13
        Grid3.Cell(Grid3.rows - 1, 14).text = D14
        Grid3.Cell(Grid3.rows - 1, 15).text = D15
        Grid3.Cell(Grid3.rows - 1, 16).text = D16
        Grid3.Cell(Grid3.rows - 1, 17).text = D17
        Grid3.Cell(Grid3.rows - 1, 18).text = D18
        Grid3.Cell(Grid3.rows - 1, 19).text = D19
        Grid3.Cell(Grid3.rows - 1, 20).text = D20
        
        Grid3.Cell(Grid3.rows - 1, 21).text = D21
        
        
     
    
    
paso1:
'    Print #10, D1 + ";" + D21 + ";" + D2 + ";" + D3 + ";" + D4 + ";" + D5 + ";" + D6 + ";" + D7 + ";" + D8 + ";" + D9 + ";" + D10 + ";" + D11 + ";" + D12 + ";" + D13 + ";" + D14 + ";" + D15 + ";" + D16 + ";" + D17 + ";" + D18 + ";" + D19 + ";" + D20
Next k
 


     If ExportarCSV("c:\" & "f1847_" & Format(fechasistema, "yyyy") & "_" + empresaactiva & ".csv", Grid3, ";") = True Then
        Shell "NOTEPAD " + "c:\" & "f1847_" & Format(fechasistema, "yyyy") & "_" + empresaactiva & ".csv"
     End If
'    Call Grid3.ExportToCSV("", False, False)
End If





End Sub

Public Function ExportarCSV(ByVal rutadestino As String, ByVal grilla As grid, ByVal SEPARADOR As String) As Boolean
Dim N As Double
Dim c As Double
Dim columnas As Double
Dim trama As String
Dim MENSAJE As String
Dim campos(999) As Double
MENSAJE = "ARCHIVO YA EXISTE Y NO SE PUDO REEMPLAZAR" & vbNewLine & " NO SE PUDO CONTINUAR"
If ExisteArchivo(rutadestino) = True Then
    On Local Error Resume Next
    Call Kill(rutadestino)
End If
If ExisteArchivo(rutadestino) = True Then
    MsgBox MENSAJE
    Exit Function
End If

columnas = grilla.cols - 1
For N = 1 To grilla.rows - 1
    For c = 1 To columnas
        trama = trama & grilla.Cell(N, c).text
        If c < columnas Then
                trama = trama & SEPARADOR
        End If
        
    Next c
    Call GrabarLineaArchivo(rutadestino, trama)
        trama = ""
Next N
If ExisteArchivo(rutadestino) = True Then
    ExportarCSV = True
End If
End Function


Public Sub GrabarLineaArchivo(ARCHIVO, lineanueva)


Close 20
Open ARCHIVO For Append As #20
Print #20, lineanueva
Close 20
End Sub


Private Sub Command1_Click()
    GRID1.Range(0, 1, 0, GRID1.cols - 1).Borders(cellEdgeBottom) = cellThick
    GRID1.Range(0, 1, 0, GRID1.cols - 1).Borders(cellEdgeLeft) = cellThick
    GRID1.Range(0, 1, 0, GRID1.cols - 1).Borders(cellEdgeTop) = cellThick
    GRID1.Range(0, 1, 0, GRID1.cols - 1).Borders(cellEdgeRight) = cellThick
    GRID1.Range(0, 1, 0, GRID1.cols - 1).Borders(cellInsideHorizontal) = cellThick
    GRID1.Range(0, 1, 0, GRID1.cols - 1).Borders(cellInsideVertical) = cellThick
    GRID1.PageSetup.PaperWidth = 21.59
    GRID1.PageSetup.PaperHeight = 27.94
    
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

Private Sub COMMAND2_Click()
GRID1.ExportToExcel (""), True



End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_Load()


Call CENTRAR(Me)
LETRA.Caption = GRID1.DefaultFont.Name
TAMAÑOS.Min = GRID1.DefaultFont.Size - 5
TAMAÑOS.Max = GRID1.DefaultFont.Size + 10
TAMAÑOS.Value = GRID1.DefaultFont.Size
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Locked = True


Next k
If xmlcompra = True Then
cmd_xml.Visible = True

End If
If xmlventa = True Then
cmd_xml.Visible = True

End If



End Sub

Sub imprime_balancetributario(tipo, folio)
Dim titulo As String
Dim subtitulo As String

subtitulo = Mid(cabeza.Caption, 20, ((InStr(cabeza.Caption, "empresa") - 6) - 20))
titulo = "BALANCE TRIBUTARIO"
Call cabezas(titulo, tipo, folio, subtitulo)
GRID1.DefaultFont.Size = 7

If GRID1.Column(1).Width <> "0" Then
    GRID1.PageSetup.Orientation = cellLandscape
'    For k = 1 To 11 - 1
'        Grid1.Column(k).Width = Grid1.Column(k).Width / 8 * 6
'    Next k
Else
    GRID1.PageSetup.Orientation = cellPortrait
    GRID1.Column(1).Width = 0
    For k = 2 To 11 - 1
        GRID1.Column(k).Width = GRID1.Column(k).Width / 8 * 6
    Next k
End If

GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 1
GRID1.PageSetup.RightMargin = 0
GRID1.PrintPreview 120

End Sub
Sub imprime_balanceanalitico(tipo, folio)
Dim titulo As String
titulo = "BALANCE ANALITICO"
Call cabezas(titulo, tipo, folio, "")
'grid1.DefaultFont.Size = 6
GRID1.Column(2).Width = 20 * 8
GRID1.PageSetup.Orientation = cellPortrait
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 1
GRID1.PageSetup.RightMargin = 0
GRID1.Refresh

GRID1.PrintPreview 120

End Sub
Sub imprime_mayoranalitico(tipo, folio, TITULOalfinal)
Dim titulo As String
titulo = "MAYOR ANALITICO"
titulo = TITULOalfinal
GRID1.DefaultFont.Size = 8
For k = 1 To 10 - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 8 * 8
Next k
Call cabezas(titulo, tipo, folio, "")
GRID1.PageSetup.Orientation = cellPortrait
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 1
GRID1.PageSetup.RightMargin = 0
GRID1.Refresh

GRID1.PrintPreview 120

End Sub
Sub imprime_cartolamayor(tipo, folio)

Dim titulo As String
titulo = "CARTOLA DEL MAYOR"
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 6
For k = 1 To 15 - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 8 * 6
Next k
GRID1.Column(6).Width = 200
GRID1.Column(5).Width = 0

GRID1.Column(14).Width = 0
GRID1.PageSetup.Orientation = cellPortrait
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 1
GRID1.PageSetup.RightMargin = 0
GRID1.Refresh
GRID1.PrintPreview 120
End Sub
Sub imprime_cartolaCTACTE(tipo, folio)
Dim titulo As String
titulo = "CARTOLA DEL CUENTAS CORRIENTES"
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 6
For k = 1 To 10 - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 8 * 6
Next k

GRID1.PageSetup.Orientation = cellLandscape

GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 1
GRID1.PageSetup.RightMargin = 0
GRID1.Refresh

GRID1.PrintPreview 120

End Sub



Sub imprime_librodiario(tipo, folio, titulocabeza)
Dim titulo As String
titulo = "LIBRO DIARIO"
titulo = titulocabeza
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 8
GRID1.PageSetup.Orientation = cellPortrait
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 1
GRID1.PageSetup.RightMargin = 0

For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 8
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub

Sub imprime_librocompras(tipo, folio)
Dim titulo As String
titulo = "LIBRO DE COMPRAS " + auxiliar05.COMBOMES.text + " de " + auxiliar05.COMBOAÑO.text


Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 6
GRID1.PageSetup.Orientation = cellLandscape


If tipo <> "N" Then
GRID1.PageSetup.Orientation = cellPortrait

GRID1.cols = 13
End If
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub
Sub imprime_publicidad(tipo, folio)
Dim titulo As String
titulo = "LIBRO DE COMPRAS a PROVEEDORES PERIODO " + publi0004.desdefecha.Caption + " HASTA " + publi0004.hastafecha.Caption

tipo = "N"
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 8
GRID1.PageSetup.Orientation = cellPortrait


GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub

Sub imprime_libroventas(tipo, folio)
Dim titulo As String
titulo = "LIBRO DE VENTAS " + auxiliar44.Combocrcc.text + " " + auxiliar44.COMBOMES.text + " de " + auxiliar44.COMBOAÑO.text

Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 7
GRID1.PageSetup.Orientation = cellPortrait


GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub
Sub imprime_libroboletas(tipo, folio)
Dim titulo As String
titulo = "LIBRO DE BOLETAS " + auxiliar07.Combocrcc.text + " " + auxiliar07.COMBOMES.text + " de " + auxiliar07.COMBOAÑO.text

Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 7
GRID1.PageSetup.Orientation = cellPortrait


GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub


Sub imprime_facturasporpagar(tipo, folio)
Dim titulo As String
titulo = "FACTURAS POR PAGAR"
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 6
GRID1.PageSetup.Orientation = cellLandscape

GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub
Sub imprime_estadoresultado(tipo, folio)
Dim titulo As String
titulo = "ESTADO DE RESULTADO "
Call cabezas(titulo, "N", folio, "")
GRID1.DefaultFont.Size = 6
GRID1.PageSetup.Orientation = cellLandscape

 For i = 1 To GRID1.PageSetup.PaperSizes.Count
            If UCase(GRID1.PageSetup.PaperSizes.Item(i).PaperName) = "OFICIO" Then
                GRID1.PageSetup.PaperSize = GRID1.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        


GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 0
GRID1.PageSetup.TopMargin = 0.5
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k
GRID1.SelectionMode = cellSelectionFree


GRID1.Refresh
GRID1.PrintPreview 120
End Sub


Sub imprime_honorariosporpagar(tipo, folio)
Dim titulo As String
titulo = "HONORARIOS POR PAGAR"
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 6
GRID1.PageSetup.Orientation = cellLandscape

GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub
Sub imprime_ventasporpagar(tipo, folio)
Dim titulo As String
titulo = "VENTAS POR PAGAR"
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 6
GRID1.PageSetup.Orientation = cellLandscape

GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub
Sub imprime_buscapormonto(tipo, folio)
Dim titulo As String
titulo = "BUSCA POR MONTO "
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 6
GRID1.PageSetup.Orientation = cellLandscape

GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub


Sub imprime_buscacuentaseliminadas(tipo, folio)
Dim titulo As String
titulo = "LISTADO DE CUENTAS ELIMINADAS "
Call cabezas(titulo, tipo, folio, "")
GRID1.DefaultFont.Size = 6
GRID1.PageSetup.Orientation = cellPortrait

GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub


Sub imprime_librohonorarios(tipo, folio)
    Dim titulo As String
    titulo = "LIBRO DE HONORARIOS " + auxiliar06.COMBOMES.text + " de " + auxiliar06.COMBOAÑO.text
    Call cabezas(titulo, tipo, folio, "")
    GRID1.DefaultFont.Size = 6
    GRID1.PageSetup.Orientation = cellPortrait
    GRID1.PageSetup.PrintFixedRow = True
    GRID1.PageSetup.BottomMargin = 2
    GRID1.PageSetup.TopMargin = 1
    GRID1.PageSetup.LeftMargin = 0.5
    GRID1.PageSetup.RightMargin = 0
    
    
    For k = 1 To GRID1.cols - 1
        GRID1.Column(k).Width = GRID1.Column(k).Width / 8 * 6
    Next k
    GRID1.DisablePrintButton = False
    GRID1.Refresh
    GRID1.PrintPreview 120
   
End Sub


Sub cabezas(titulo, tipo, folio, subtitulo)
Dim objReportTitle As FlexCell.ReportTitle
GRID1.ReportTitles.Clear
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    If subtitulo <> "" Then
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = UCase(subtitulo)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 9
        objReportTitle.Font.Bold = True
        objReportTitle.PrintOnAllPages = True
        GRID1.ReportTitles.Add objReportTitle
    End If
    
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
        GRID1.ReportTitles.Add objReportTitle
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
        GRID1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        GRID1.ReportTitles.Add objReportTitle
        
    End If
    
With GRID1.PageSetup
        
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

Private Sub Form_Unload(Cancel As Integer)
Unload informa04

End Sub

Private Sub Grid1_DblClick()
Dim dia As String

If Mid(cabeza.Caption, 1, 7) = "CARTOLA" Then
da0 = GRID1.Cell(GRID1.ActiveCell.row, 2).text
da1 = GRID1.Cell(GRID1.ActiveCell.row, 3).text
da2 = Format(GRID1.Cell(GRID1.ActiveCell.row, 1).text, "dd")
da3 = Format(GRID1.Cell(GRID1.ActiveCell.row, 1).text, "mm")
da4 = Format(GRID1.Cell(GRID1.ActiveCell.row, 1).text, "yyyy")
muestracomprobantes.Show vbModal
End If
If Mid(grillainformes.Tag, 1, 7) = "PRESU04" Then

End If
If Mid(cabeza.Caption, 1, 6) = "ESTADO" Then
Load informa04

informa04.cmdato1.text = Mid(GRID1.Cell(GRID1.ActiveCell.row, 1).text, 1, 2)
informa04.cmdato2.text = Mid(GRID1.Cell(GRID1.ActiveCell.row, 1).text, 3, 2)
informa04.cmdato3.text = Mid(GRID1.Cell(GRID1.ActiveCell.row, 1).text, 5, 4)
informa04.desdefecha.Caption = "01" + "-" + Format(GRID1.ActiveCell.col - 2, "00") + "-" + Format(fechasistema, "yyyy")
informa04.cmnombre.Caption = GRID1.Cell(GRID1.ActiveCell.row, 2).text
dia = Format(DateSerial(Format(fechasistema, "yyyy"), (GRID1.ActiveCell.col - 2) + 1, 0), "dd")
informa04.hastafecha.Caption = dia + "-" + Format(GRID1.ActiveCell.col - 2, "00") + "-" + Format(fechasistema, "yyyy")


informa04.Show



End If

End Sub

Private Sub TAMAÑOS_Click()

GRID1.DefaultFont.Size = TAMAÑOS.Value
'For K = 1 To Grid1.Cols - 1
'Grid1.Column(K).Width = Len(Grid1.Cell(2, K).text) * TAMAÑOS.Value


'Next K



GRID1.Refresh

End Sub
Sub imprime_NORMALES(Titulos)
Dim titulo As String
titulo = Titulos
Call cabezas(titulo, "N", "000000000", "")
'grid1.DefaultFont.Size = 6

GRID1.PageSetup.Orientation = cellPortrait
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 1
GRID1.PageSetup.RightMargin = 0
GRID1.Refresh

GRID1.PrintPreview 120

End Sub

Sub imprime_harina(tipo, folio)
Dim titulo As String
Dim titulo2 As String

titulo = "ANEXO INFORME MENSUAL VENDEDORES DE HARINA "
titulo2 = "INFORMACION DEL MES DE " + infoharina.COMBOMES.text + " AÑO " + infoharina.COMBOAÑO.text



Call CABEZAS2(titulo, titulo2)
GRID1.DefaultFont.Size = 7
GRID1.PageSetup.Orientation = cellPortrait



GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub
Sub imprime_CARNE(tipo, folio)
Dim titulo As String
Dim titulo2 As String

titulo = "ANEXO INFORME MENSUAL RETENCION CARNE "
titulo2 = "INFORMACION DEL MES DE " + infocarne.COMBOMES.text + " AÑO " + infocarne.COMBOAÑO.text



Call CABEZAS2(titulo, titulo2)
GRID1.DefaultFont.Size = 7
GRID1.PageSetup.Orientation = cellPortrait



GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub


Sub CABEZAS2(titulo, titulo2)
Dim objReportTitle As FlexCell.ReportTitle
GRID1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo2
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
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
        GRID1.ReportTitles.Add objReportTitle
    Next k
    
With GRID1.PageSetup
        
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
Sub imprime_descuadrados(tipo, folio)
Dim titulo As String
titulo = "LISTADOS DESCUADRADOS "
Call cabezas(titulo, "N", folio, "")
GRID1.DefaultFont.Size = 6
GRID1.PageSetup.Orientation = cellPortrait


 For i = 1 To GRID1.PageSetup.PaperSizes.Count
            If UCase(GRID1.PageSetup.PaperSizes.Item(i).PaperName) = "OFICIO" Then
                GRID1.PageSetup.PaperSize = GRID1.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        


GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.BottomMargin = 0
GRID1.PageSetup.TopMargin = 0.5
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0
For k = 1 To GRID1.cols - 1
GRID1.Column(k).Width = GRID1.Column(k).Width / 7 * 6
Next k

GRID1.Refresh
GRID1.PrintPreview 120
End Sub


