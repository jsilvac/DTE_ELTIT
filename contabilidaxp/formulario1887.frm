VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form form1887 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTA CERTIFICADOS DE RENTA"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15120
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1008
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
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
      Width           =   15105
      _ExtentX        =   26644
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
      Begin FlexCell.Grid Grid3 
         Height          =   255
         Left            =   14160
         TabIndex        =   20
         Top             =   8280
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
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
         Width           =   14880
         _ExtentX        =   26247
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
         Begin VB.CommandButton Command4 
            Caption         =   "IMPRIME PLANILLA"
            Height          =   495
            Left            =   4440
            TabIndex        =   16
            Top             =   480
            Width           =   3975
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "LISTAR"
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
            Left            =   13080
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   1455
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
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   0
         TabIndex        =   3
         Top             =   1485
         Width           =   15045
         _ExtentX        =   26538
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
            Top             =   225
            Width           =   15015
            _ExtentX        =   26485
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
Attribute VB_Name = "form1887"
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
    
    If LeerAgricola(Grid1.Cell(s, 1).text, "", Format(fechasistema, "YYYY"), empresaactiva) = False Then
    
        Call IMPRIMIR2(Grid1.Cell(s, 1).text, Grid1.Cell(s, 2).text, Grid1.Cell(s, 21).text)
    Else
        
        Call IMPRIMIR3(Grid1.Cell(s, 1).text, Grid1.Cell(s, 2).text, Grid1.Cell(s, 21).text)
    End If
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
Dim D21 As String


Close 10
Grid3.Cols = 22
Grid3.Rows = 1

Open "F1887_" + empresaactiva + ".TXT" For Output As #10
For k = 2 To Grid1.Rows - 8
'    D1 = CDbl(Mid(Grid1.Cell(k, 1).text, 2, 8))
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
'    D21 = Mid(GRID1.Cell(k, 1).text, 10, 1)
    
   If Format(fechasistema, "yyyy") < "2019" Then
        D1 = CDbl(Mid(Grid1.Cell(k, 1).text, 2, 8)) & Mid(Grid1.Cell(k, 1).text, 10, 1)
        Print #10, D1 + ";" + D2 + ";" + D3 + ";" + D4 + ";" + D5 + ";" + D6 + ";" + D7 + ";" + D8 + ";" + D9 + ";" + D10 + ";" + D11 + ";" + D12 + ";" + D13 + ";" + D14 + ";" + D15 + ";" + D16 + ";" + D17 + ";" + D18 + ";" + D19 + ";" + D20

    Else
        D1 = CDbl(Mid(Grid1.Cell(k, 1).text, 2, 8))
        D21 = Mid(Grid1.Cell(k, 1).text, 10, 1)
        
        Print #10, D1 + ";" + D21 + ";" + D2 + ";" + D3 + ";" + D4 + ";" + D5 + ";" + D6 + ";" + D7 + ";" + D8 + ";" + D9 + ";" + D10 + ";" + D11 + ";" + D12 + ";" + D13 + ";" + D14 + ";" + D15 + ";" + D16 + ";" + D17 + ";" + D18 + ";" + D19 + ";" + D20
        Grid3.Rows = Grid3.Rows + 1
        Grid3.Cell(Grid3.Rows - 1, 1).text = D1
        Grid3.Cell(Grid3.Rows - 1, 2).text = D21
        Grid3.Cell(Grid3.Rows - 1, 3).text = D2
        Grid3.Cell(Grid3.Rows - 1, 4).text = D3
        Grid3.Cell(Grid3.Rows - 1, 5).text = D4
        Grid3.Cell(Grid3.Rows - 1, 6).text = D5
        Grid3.Cell(Grid3.Rows - 1, 7).text = D6
        Grid3.Cell(Grid3.Rows - 1, 8).text = D7
        Grid3.Cell(Grid3.Rows - 1, 9).text = D8
        Grid3.Cell(Grid3.Rows - 1, 10).text = D9
        Grid3.Cell(Grid3.Rows - 1, 11).text = D10
        Grid3.Cell(Grid3.Rows - 1, 12).text = D11
        Grid3.Cell(Grid3.Rows - 1, 13).text = D12
        Grid3.Cell(Grid3.Rows - 1, 14).text = D13
        Grid3.Cell(Grid3.Rows - 1, 15).text = D14
        Grid3.Cell(Grid3.Rows - 1, 16).text = D15
        Grid3.Cell(Grid3.Rows - 1, 17).text = D16
        Grid3.Cell(Grid3.Rows - 1, 18).text = D17
        Grid3.Cell(Grid3.Rows - 1, 19).text = D18
        Grid3.Cell(Grid3.Rows - 1, 20).text = D19
        Grid3.Cell(Grid3.Rows - 1, 21).text = D20
        
        
    End If
    
    
    
'    Print #10, D1 + ";" + D21 + ";" + D2 + ";" + D3 + ";" + D4 + ";" + D5 + ";" + D6 + ";" + D7 + ";" + D8 + ";" + D9 + ";" + D10 + ";" + D11 + ";" + D12 + ";" + D13 + ";" + D14 + ";" + D15 + ";" + D16 + ";" + D17 + ";" + D18 + ";" + D19 + ";" + D20
Next k
Close #10

If Format(fechasistema, "yyyy") < "2019" Then
    Shell "NOTEPAD " + "F1887_" + empresaactiva + ".txt"
Else
     If ExportarCSV("c:\" & "F1887_" + empresaactiva & ".csv", Grid3, ";") = True Then
        Shell "NOTEPAD " + "c:\" & "F1887_" + empresaactiva & ".csv"
     End If
'    Call Grid3.ExportToCSV("", False, False)
End If




End Sub

Public Function ExportarCSV(ByVal rutadestino As String, ByVal grilla As Grid, ByVal separador As String) As Boolean
Dim n As Double
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

columnas = grilla.Cols - 1
For n = 1 To grilla.Rows - 1
    For c = 1 To columnas
        trama = trama & grilla.Cell(n, c).text
        If c < columnas Then
                trama = trama & separador
        End If
        
    Next c
    Call GrabarLineaArchivo(rutadestino, trama)
        trama = ""
Next n
If ExisteArchivo(rutadestino) = True Then
    ExportarCSV = True
End If
End Function


Public Sub GrabarLineaArchivo(archivo, lineanueva)


Close 20
Open archivo For Append As #20
Print #20, lineanueva
Close 20
End Sub

Private Sub Command4_Click()
Call cabezas4("INFORME CERTIFICADO 1887", "N", 0)

Grid1.PrintPreview

End Sub

Private Sub COMMAND2_Click()
    leer
End Sub
Private Sub FIRMA_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD

    sc = 0
CARGAGRILLA
Call Conectarventas(Servidor, clientesistema + "ventas00", Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion00", Usuario, password)

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

Sub IMPRIMIR3(rut, NOMBRE, numero)
Dim titulo As String
Call cabezas5(rut, NOMBRE, numero)

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
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 30)
    Grid1.DefaultFont.Size = 8
    Grid1.FixedRows = 2
       
    FORMATOGRILLA(1, 1) = "RUT"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = "RENTA TOTAL" + vbCrLf + "NETA PAGADA"
    FORMATOGRILLA(1, 4) = "IMPUESTO UNICO" & vbCrLf & "RETENIDO"
    FORMATOGRILLA(1, 5) = "MAYOR RETENCION" & vbCrLf & "SOLICITADA"
    FORMATOGRILLA(1, 6) = "RENTA TOTAL" & vbCrLf & "NO" & vbCrLf & "GRAVADA"
    FORMATOGRILLA(1, 7) = "RENTA TOTAL" & vbCrLf & "EXENTA "
    FORMATOGRILLA(1, 8) = "REBAJA POR ZONAS" & vbCrLf & "EXTREMAS" & vbCrLf & "(FRANQUICIA" & vbCrLf & "D.L.889)"
    FORMATOGRILLA(1, 9) = "ENE"
    FORMATOGRILLA(1, 10) = "FEB"
    FORMATOGRILLA(1, 11) = "MAR"
    FORMATOGRILLA(1, 12) = "ABR"
    FORMATOGRILLA(1, 13) = "MAY"
    FORMATOGRILLA(1, 14) = "JUN"
    FORMATOGRILLA(1, 15) = "JUL"
    FORMATOGRILLA(1, 16) = "AGO"
    FORMATOGRILLA(1, 17) = "SEP"
    FORMATOGRILLA(1, 18) = "OCT"
    FORMATOGRILLA(1, 19) = "NOV"
    FORMATOGRILLA(1, 20) = "DIC"
    FORMATOGRILLA(1, 21) = "NUMERO" & vbCrLf & "CERTIFICADO"
    FORMATOGRILLA(1, 22) = "IMPRIMIR"
    
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "25"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "8"
    FORMATOGRILLA(2, 8) = "8"
    
    'MESES
    FORMATOGRILLA(2, 9) = "3"
    FORMATOGRILLA(2, 10) = "3"
    FORMATOGRILLA(2, 11) = "3"
    FORMATOGRILLA(2, 12) = "3"
    FORMATOGRILLA(2, 13) = "3"
    FORMATOGRILLA(2, 14) = "3"
    FORMATOGRILLA(2, 15) = "3"
    FORMATOGRILLA(2, 16) = "3"
    FORMATOGRILLA(2, 17) = "3"
    FORMATOGRILLA(2, 18) = "3"
    FORMATOGRILLA(2, 19) = "3"
    FORMATOGRILLA(2, 20) = "3"
    'MESES
    FORMATOGRILLA(2, 21) = "8"
    FORMATOGRILLA(2, 22) = "1"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
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
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = "##,###,##0"
    FORMATOGRILLA(4, 4) = "##,###,##0"
    FORMATOGRILLA(4, 5) = "##,###,##0"
    FORMATOGRILLA(4, 6) = "##,###,##0"
    FORMATOGRILLA(4, 7) = "##,###,##0"
    
    
    Rem LOCCKED
    For k = 1 To 22
    FORMATOGRILLA(5, k) = "FALSE"
    Next k
    FORMATOGRILLA(5, k) = "FALSE"
    
    Grid1.Cols = 23
    Grid1.Rows = 2
    
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
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
      Grid1.RowHeight(1) = 80
      Grid1.Range(0, 1, 1, Grid1.Cols - 1).WrapText = True
      Grid1.Range(0, 1, 1, Grid1.Cols - 1).FontSize = 6.5
      Grid1.Cell(0, 1).Alignment = cellCenterCenter
      Grid1.Cell(0, 2).Alignment = cellCenterCenter
      Grid1.Cell(0, 3).Alignment = cellCenterCenter
      Grid1.Cell(0, 4).Alignment = cellCenterCenter
      Grid1.Cell(0, 5).Alignment = cellCenterCenter
      Grid1.Cell(0, 6).Alignment = cellCenterCenter
      Grid1.Cell(0, 7).Alignment = cellCenterCenter
      Grid1.Cell(0, 8).Alignment = cellCenterCenter
      Grid1.Cell(0, 9).Alignment = cellCenterCenter
      Grid1.Cell(0, 10).Alignment = cellCenterCenter
      Grid1.Cell(0, 11).Alignment = cellCenterCenter
      Grid1.Cell(0, 12).Alignment = cellCenterCenter
      Grid1.Cell(0, 13).Alignment = cellCenterCenter
      Grid1.Cell(0, 14).Alignment = cellCenterCenter
      Grid1.Cell(0, 15).Alignment = cellCenterCenter
      Grid1.Cell(0, 16).Alignment = cellCenterCenter
      Grid1.Cell(0, 17).Alignment = cellCenterCenter
      Grid1.Cell(0, 18).Alignment = cellCenterCenter
      Grid1.Cell(0, 19).Alignment = cellCenterCenter
      Grid1.Cell(0, 20).Alignment = cellCenterCenter
      Grid1.Cell(0, 21).Alignment = cellCenterCenter
      Grid1.Column(22).CellType = cellCheckBox
      
 
    
    Grid1.Range(0, 1, 1, 1).Merge
    Grid1.Range(0, 2, 1, 2).Merge
    Grid1.Range(0, 3, 0, 8).Merge
    Grid1.Range(0, 9, 0, 20).Merge
    Grid1.Range(0, 21, 1, 21).Merge
    Grid1.Range(0, 22, 1, 22).Merge
    
   Grid1.Cell(0, 3).text = "MONTOS ANUALES ACTUALIZADOS"
   Grid1.Cell(0, 8).text = "PERIODO AL CUAL CORRESPONDEN LAS RENTAS"
   Grid1.Cell(1, 3).text = "RENTA TOTAL" + vbCrLf + "NETA PAGADA"
   Grid1.Cell(1, 4).text = "IMPUESTO UNICO" & vbCrLf & "RETENIDO"
   Grid1.Cell(1, 5).text = "MAYOR RETENCION" & vbCrLf & "SOLICITADA"
   Grid1.Cell(1, 6).text = "RENTA TOTAL" & vbCrLf & "NO" & vbCrLf & "GRAVADA"
   Grid1.Cell(1, 7).text = "RENTA TOTAL" & vbCrLf & "EXENTA "
   Grid1.Cell(1, 8).text = "REBAJA POR ZONAS" & vbCrLf & "EXTREMAS" & vbCrLf & "(FRANQUICIA" & vbCrLf & "D.L.889)"
   Grid1.Cell(1, 9).text = "ENE"
   Grid1.Cell(1, 10).text = "FEB"
   Grid1.Cell(1, 11).text = "MAR"
   Grid1.Cell(1, 12).text = "ABR"
   Grid1.Cell(1, 13).text = "MAY"
   Grid1.Cell(1, 14).text = "JUN"
   Grid1.Cell(1, 15).text = "JUL"
   Grid1.Cell(1, 16).text = "AGO"
   Grid1.Cell(1, 17).text = "SEP"
   Grid1.Cell(1, 18).text = "OCT"
   Grid1.Cell(1, 19).text = "NOV"
   Grid1.Cell(1, 20).text = "DIC"
      
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
    Grid1.Column(8).Locked = True
    Grid1.Column(9).Locked = True
    Grid1.Column(10).Locked = True
    Grid1.Column(11).Locked = True
    Grid1.Column(12).Locked = True
    Grid1.Column(13).Locked = True
    Grid1.Column(14).Locked = True
    Grid1.Column(15).Locked = True
    Grid1.Column(16).Locked = True
    Grid1.Column(17).Locked = True
    Grid1.Column(18).Locked = True
    Grid1.Column(19).Locked = True
    Grid1.Column(20).Locked = True
    Grid1.Column(21).Locked = True
    Grid1.Column(22).Locked = False
   
   Grid1.Refresh
    
    
End Sub



Private Sub monto_Click()
End Sub

Private Sub leer()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim TOTAL As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim tila1 As Double
    Dim tila2 As Double
    Dim tila3 As Double
    Dim total3 As Double
    Dim total4 As Double
    Dim total6 As Double
    Dim total7 As Double
    Dim totalcorregido As Double
    Dim totalsincorregir As Double
    Dim EsAgricola As Boolean
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    
        Set csql.ActiveConnection = contadb
'        csql.sql = "select rut,sum(monto),sum(retencion) from boletasdehonorarios where retencion<>'0' and añocontable='" + COMBOAÑO.text + "' group by rut order by rut "
        
        csql.sql = "select rut,SUM(if(codigo='THI01',monto,0)) AS dos, "
        csql.sql = csql.sql & "SUM(if(codigo='AFP01',monto,0)) + SUM(if(codigo='ISA01',monto,0)) + "
        csql.sql = csql.sql & "SUM(if(codigo='ISA03',monto,0)) AS tres, "
        csql.sql = csql.sql & "SUM(if(codigo='IRE02',monto,0))- "
        csql.sql = csql.sql & "SUM(if(codigo='AFP01',monto,0))- SUM(if(codigo='ISA01',monto,0))- "
        csql.sql = csql.sql & "SUM(if(codigo='ISA03',monto,0)) AS cuatro, "
        csql.sql = csql.sql & "SUM(if(codigo='IRE01',monto,0)) AS cinco, "
        csql.sql = csql.sql & "SUM(IF(mid(codigo,1,2)='HN',monto,0))+SUM(IF(mid(codigo,1,2)='FN',monto,0)) as seis, "
        csql.sql = csql.sql & "mes, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='01',monto,0)) AS mes01, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='02',monto,0)) AS mes02, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='03',monto,0)) AS mes03, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='04',monto,0)) AS mes04, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='05',monto,0)) AS mes05, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='06',monto,0)) AS mes06, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='07',monto,0)) AS mes07, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='08',monto,0)) AS mes08, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='09',monto,0)) AS mes09, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='10',monto,0)) AS mes10, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='11',monto,0)) AS mes11, "
        csql.sql = csql.sql & "SUM(if(codigo='THI01' and mes='12',monto,0)) AS mes12  "
        csql.sql = csql.sql & "from " & clientesistema & "remu" & empresaactiva & ".calculoliquidaciones where  año ='" & COMBOAÑO.text & "' "
'        csql.sql = csql.sql & "and rut='" & rut & "' group by mes order by mes "
        csql.sql = csql.sql & "group by rut order by rut"
 
        totalsincorre1 = 0
        totalsincorre2 = 0
        totalsincorre3 = 0
        totalsincorre4 = 0
        totalsincorre5 = 0
        Grid1.AutoRedraw = False
        
        
        csql.Execute
        Grid1.Rows = 2
        LINEA = Grid1.Rows - 1
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
             LINEA = LINEA + 1
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = leerdatos(contadb, clientesistema & "remu" & empresaactiva & ".mt_fijo", "nombre", "rut='" + resultados(0) + "' ")
             Call calculacertificado(resultados(0), LINEA)
              
             If resultados(0) = "0173200056" Then
             Print "hola"
             End If
             
             'B : TRABAJADOR AGRICOLA CON JORNADA PARCIAL       ** NUEVO 2015
             'A : TRABAJADOR AGRICOLA JORNADA COMPLETA          ** NUEVO 2015
             
             'P : TRABAJADOR NO AGRICOLA CON JORNADA PARCIAL    ** ORIGINAL
             'C : TRABAJADOR NO AGRICOLA CON JORNADA COMPLETA   ** ORIGINA
             
             
             'If resultados(7) > 0 Then
             If (resultados(7) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-01-31") > 0) Then
                EsAgricola = LeerAgricola(resultados(0), "01", Format(fechasistema, "yyyy"), empresaactiva)
                
                If leerpartime(resultados(0), "01", Format(fechasistema, "yyyy"), empresaactiva) = True Then
                    If EsAgricola = True Then
                        Grid1.Cell(LINEA, 9).text = "B"
                    Else
                        Grid1.Cell(LINEA, 9).text = "P"
                    End If
                Else
                    If EsAgricola = True Then
                        Grid1.Cell(LINEA, 9).text = "A"
                    Else
                        Grid1.Cell(LINEA, 9).text = "C"
                    End If
                End If
             End If
             
             If (resultados(8) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-02-28") > 0) Then
             
             EsAgricola = LeerAgricola(resultados(0), "02", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "02", Format(fechasistema, "yyyy"), empresaactiva) = True Then
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 10).text = "B"
                    Else
                        Grid1.Cell(LINEA, 10).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 10).text = "A"
                    Else
                        Grid1.Cell(LINEA, 10).text = "C"
                    End If
             End If
             End If
             'If resultados(9) > 0 Then
             If (resultados(9) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-03-31") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "03", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "03", Format(fechasistema, "yyyy"), empresaactiva) = True Then
             If EsAgricola = True Then
                        Grid1.Cell(LINEA, 11).text = "B"
                    Else
                        Grid1.Cell(LINEA, 11).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 11).text = "A"
                    Else
                        Grid1.Cell(LINEA, 11).text = "C"
                    End If
             End If
             
             End If
             'If resultados(10) > 0 Then
             If (resultados(10) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-04-30") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "04", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "04", Format(fechasistema, "yyyy"), empresaactiva) = True Then
              If EsAgricola = True Then
                        Grid1.Cell(LINEA, 12).text = "B"
                    Else
                        Grid1.Cell(LINEA, 12).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 12).text = "A"
                    Else
                        Grid1.Cell(LINEA, 12).text = "C"
                    End If
             End If
             
             
             End If
             'If resultados(11) > 0 Then
             If (resultados(11) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-05-31") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "05", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "05", Format(fechasistema, "yyyy"), empresaactiva) = True Then
              If EsAgricola = True Then
                        Grid1.Cell(LINEA, 13).text = "B"
                    Else
                        Grid1.Cell(LINEA, 13).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 13).text = "A"
                    Else
                        Grid1.Cell(LINEA, 13).text = "C"
                    End If
             End If
             
             
             End If
             'If resultados(12) > 0 Then
             If (resultados(12) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-06-30") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "06", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "06", Format(fechasistema, "yyyy"), empresaactiva) = True Then
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 14).text = "B"
                    Else
                        Grid1.Cell(LINEA, 14).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 14).text = "A"
                    Else
                        Grid1.Cell(LINEA, 14).text = "C"
                    End If
             End If
             
             
             End If
             
             'original If resultados(13) > 0 Then
             If (resultados(13) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-07-31") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "07", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "07", Format(fechasistema, "yyyy"), empresaactiva) = True Then
            If EsAgricola = True Then
                        Grid1.Cell(LINEA, 15).text = "B"
                    Else
                        Grid1.Cell(LINEA, 15).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 15).text = "A"
                    Else
                        Grid1.Cell(LINEA, 15).text = "C"
                    End If
             End If
             
             
             End If
             If (resultados(14) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-08-31") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "08", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "08", Format(fechasistema, "yyyy"), empresaactiva) = True Then
 If EsAgricola = True Then
                        Grid1.Cell(LINEA, 16).text = "B"
                    Else
                        Grid1.Cell(LINEA, 16).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 16).text = "A"
                    Else
                        Grid1.Cell(LINEA, 16).text = "C"
                    End If
             End If
             
             
             End If
             'If resultados(15) > 0 Then
             If (resultados(15) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-09-30") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "09", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "09", Format(fechasistema, "yyyy"), empresaactiva) = True Then
                    If EsAgricola = True Then
                        Grid1.Cell(LINEA, 17).text = "B"
                    Else
                        Grid1.Cell(LINEA, 17).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 17).text = "A"
                    Else
                        Grid1.Cell(LINEA, 17).text = "C"
                    End If
             End If
                          
             End If
             'If resultados(16) > 0 Then
             If (resultados(16) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-10-31") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "10", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "10", Format(fechasistema, "yyyy"), empresaactiva) = True Then
                    If EsAgricola = True Then
                        Grid1.Cell(LINEA, 18).text = "B"
                    Else
                        Grid1.Cell(LINEA, 18).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 18).text = "A"
                    Else
                        Grid1.Cell(LINEA, 18).text = "C"
                    End If
             End If
             
             
             End If
             'If resultados(17) > 0 Then
             If (resultados(17) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-11-30") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "11", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "11", Format(fechasistema, "yyyy"), empresaactiva) = True Then
                    If EsAgricola = True Then
                        Grid1.Cell(LINEA, 19).text = "B"
                    Else
                        Grid1.Cell(LINEA, 19).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 19).text = "A"
                    Else
                        Grid1.Cell(LINEA, 19).text = "C"
                    End If
             End If
             
             End If
             If (resultados(18) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-12-31") > 0) Then
             EsAgricola = LeerAgricola(resultados(0), "12", Format(fechasistema, "yyyy"), empresaactiva)
             If leerpartime(resultados(0), "12", Format(fechasistema, "yyyy"), empresaactiva) = True Then
                     If EsAgricola = True Then
                        Grid1.Cell(LINEA, 20).text = "B"
                    Else
                        Grid1.Cell(LINEA, 20).text = "P"
                    End If
             Else
                If EsAgricola = True Then
                        Grid1.Cell(LINEA, 20).text = "A"
                    Else
                        Grid1.Cell(LINEA, 20).text = "C"
                    End If
             End If
             End If
             
             
             
             
             'ORIGINAL
'             If (resultados(18) > 0) Or (leercargasvigentes(resultados(0), empresaactiva, Format(fechasistema, "yyyy") & "-12-31") > 0) Then
'             EsAgricola = LeerAgricola(resultados(0), "12", Format(fechasistema, "yyyy"), empresaactiva)
'             If leerpartime(resultados(0), "12", Format(fechasistema, "yyyy"), empresaactiva) = True Then
'             GRID1.Cell(LINEA, 20).text = "P"
'             Else
'             GRID1.Cell(LINEA, 20).text = "C"
'             End If
'             End If
             
             
             Grid1.Cell(LINEA, 21).text = Format(LINEA - 1, "0000000000")
             Grid1.Cell(LINEA, 22).text = "1"
             TOTAL = TOTAL + Val(Grid1.Cell(LINEA, 3).text)
             
             total2 = total2 + Val(Grid1.Cell(LINEA, 4).text)
             total3 = total3 + 0
             total4 = total4 + Val(Grid1.Cell(LINEA, 6).text)
           
             
             resultados.MoveNext
            
            Wend
             LINEA = LINEA + 1
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Range(LINEA, 1, LINEA, 6).FontBold = True
             Grid1.Range(LINEA, 1, LINEA, 6).Borders(cellEdgeTop) = cellThin
             
             'GRID1.Cell(linea, 3).text = total
             'GRID1.Cell(linea, 4).text = total2
             'GRID1.Cell(linea, 5).text = total3
             'GRID1.Cell(linea, 6).text = total4
             'GRID1.Cell(linea, 7).text = "0"
             
             
             
             Grid1.Cell(LINEA, 3).text = TOTAL
             Grid1.Cell(LINEA, 4).text = total2
             Grid1.Cell(LINEA, 5).text = total3
             Grid1.Cell(LINEA, 6).text = total4
             Grid1.Cell(LINEA, 7).text = "0"
             Grid1.Cell(LINEA, 8).text = "0"
             Grid1.Rows = Grid1.Rows + 6
             
'             Grid1.Cell(LINEA + 2, 2).text = "Total Imponible Act"
'             Grid1.Cell(LINEA + 2, 3).text = totalsincorre1
'
'             Grid1.Cell(LINEA + 3, 2).text = "Total Rent.Afec.Act"
'             Grid1.Cell(LINEA + 3, 3).text = totalsincorre2
'
'             Grid1.Cell(LINEA + 4, 2).text = "Total exenta no grab"
'             Grid1.Cell(LINEA + 4, 3).text = totalsincorre3
'
'             Grid1.Cell(LINEA + 5, 2).text = "Total Afecta No Act"
'             Grid1.Cell(LINEA + 5, 3).text = totalsincorre4


            Grid1.Cell(LINEA + 3, 2).text = "Total Remun.Imponible Actualizada"
             Grid1.Cell(LINEA + 3, 3).text = totalsincorre1
             
'             Grid1.Cell(LINEA + 3, 2).text = "Total Renta Neta Afec.Act"
'             Grid1.Cell(LINEA + 3, 3).text = totalsincorre2
             
             Grid1.Cell(LINEA + 4, 2).text = "Total exenta No Actualizada"
             Grid1.Cell(LINEA + 4, 3).text = totalsincorre3
             
             Grid1.Cell(LINEA + 5, 2).text = "Total Renta Neta No Actualizada"
             Grid1.Cell(LINEA + 5, 3).text = totalsincorre4
             
             Grid1.Cell(LINEA + 6, 2).text = "Total Impuesto Unico No Actualizado"
             Grid1.Cell(LINEA + 6, 3).text = totalsincorre5
         
         resultados.Close
            Set resultados = Nothing

End If
Grid1.AutoRedraw = True
Grid1.Refresh
        
      
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
Grid2.Column(2).Locked = False
Grid2.Column(3).Locked = False
Grid2.Column(4).Locked = False

Grid2.Column(2).Width = 100
Grid2.Column(3).Width = 100
Grid2.Column(4).Width = 100
'Grid2.Column(9).Width = 100
End Sub


Sub CARGAGRILLA2()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(20, 20)
    Grid2.DefaultFont.Size = 7
       
    FORMATOGRILLA(1, 1) = "PERIODO"
    FORMATOGRILLA(1, 2) = "SUELDO BRUTO "
    FORMATOGRILLA(1, 3) = "COTIZACIÓN " + vbCrLf + "PREVISIONAL  O DE" + vbCrLf + "SALUD DE CARGO " + vbCrLf + "DEL TRABAJADOR"
    FORMATOGRILLA(1, 4) = "RENTA IMPONIBLE  " + vbCrLf + "AFECTA AL IMPTO." + vbCrLf + " ÚNICO DE 2° CAT."
    FORMATOGRILLA(1, 5) = "IMPTO. ÚNICO" + vbCrLf + "RETENIDO"
    FORMATOGRILLA(1, 6) = "MAYOR " + vbCrLf + "RETENCIÓN DE" + vbCrLf + "IMPTO." + vbCrLf + "SOLICITADA" + vbCrLf + "ART. 88 LIR"
    FORMATOGRILLA(1, 7) = "RENTA TOTAL " + vbCrLf + "EXENTA "
    FORMATOGRILLA(1, 8) = "RENTA TOTAL " + vbCrLf + " NO " + vbCrLf + "GRAVADA"
    FORMATOGRILLA(1, 9) = "REBAJA POR" + vbCrLf + "ZONAS EXTREMAS" + vbCrLf + "(FRANQUICIA D.L." + vbCrLf + "889)"
    FORMATOGRILLA(1, 10) = "FACTOR" + vbCrLf + "ACTUALIZACIÓN"
    FORMATOGRILLA(1, 11) = "RENTA AFECTA A" + vbCrLf + "IMPTO. ÚNICO" + vbCrLf + "DE 2° CAT."
    FORMATOGRILLA(1, 12) = "IMPTO. ÚNICO " + vbCrLf + "RETENIDO"
    FORMATOGRILLA(1, 13) = "MAYOR RETENCIÓN" + vbCrLf + "DE IMPTO." + vbCrLf + "SOLICITADA ART.88" & vbCrLf & "LIR."
    FORMATOGRILLA(1, 14) = "RENTA TOTAL " + vbCrLf + "EXENTA "
    FORMATOGRILLA(1, 15) = "RENTA TOTAL " + vbCrLf + " NO " + vbCrLf + "GRAVADA"
    
    FORMATOGRILLA(1, 16) = "REBAJA POR" + vbCrLf + "ZONAS EXTREMAS" + vbCrLf + "(FRANQUICIA " + vbCrLf + "D.L. 889)"

    
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "9"
    FORMATOGRILLA(2, 2) = "9"
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "9"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "9"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    FORMATOGRILLA(2, 11) = "9"
    FORMATOGRILLA(2, 12) = "9"
    FORMATOGRILLA(2, 13) = "9"
    FORMATOGRILLA(2, 14) = "9"
    FORMATOGRILLA(2, 15) = "9"
    FORMATOGRILLA(2, 16) = "9"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
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
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 2) = "##,###,##0"
    FORMATOGRILLA(4, 3) = "##,###,##0"
    FORMATOGRILLA(4, 4) = "##,###,##0"
    FORMATOGRILLA(4, 5) = "##,###,##0"
    FORMATOGRILLA(4, 6) = "##,###,##0"
    FORMATOGRILLA(4, 7) = "##,###,##0"
    FORMATOGRILLA(4, 8) = "##,###,##0"
    
    FORMATOGRILLA(4, 9) = "##,###,##0"
    FORMATOGRILLA(4, 10) = "##,###,##0.000"
    FORMATOGRILLA(4, 11) = "##,###,##0"
    FORMATOGRILLA(4, 12) = "##,###,##0"
    FORMATOGRILLA(4, 13) = "##,###,##0"
    FORMATOGRILLA(4, 14) = "##,###,##0"
    FORMATOGRILLA(4, 15) = "##,###,##0"
    FORMATOGRILLA(4, 16) = "##,###,##0"
    
    
    Rem LOCCKED
    For k = 1 To 16
    FORMATOGRILLA(5, k) = "false"
    
    Next k
        
    
    Grid2.Cols = 17
    Grid2.Rows = 2
    
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    
'   grid2.BackColorFixed = RGB(90, 158, 214)
'   grid2.BackColorFixedSel = RGB(110, 180, 230)
'   grid2.BackColorBkg = RGB(90, 158, 214)
'   grid2.BackColorScrollBar = RGB(231, 235, 247)
'   grid2.BackColor1 = RGB(231, 235, 247)
'   grid2.BackColor2 = RGB(239, 243, 255)
'   grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        
        Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid2.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid2.DefaultFont.Size
        Grid2.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid2.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid2.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
        Grid2.RowHeight(0) = 80
   
      Grid2.Range(0, 1, 0, Grid2.Cols - 1).WrapText = True
      Grid2.Range(0, 1, 0, Grid2.Cols - 1).FontSize = 5
      Grid2.Cell(0, 1).Alignment = cellCenterCenter
      Grid2.Cell(0, 2).Alignment = cellCenterCenter
      Grid2.Cell(0, 3).Alignment = cellCenterCenter
      Grid2.Cell(0, 4).Alignment = cellCenterCenter
      Grid2.Cell(0, 5).Alignment = cellCenterCenter
      Grid2.Cell(0, 6).Alignment = cellCenterCenter
      Grid2.Cell(0, 7).Alignment = cellCenterCenter
      Grid2.Cell(0, 8).Alignment = cellCenterCenter
      Grid2.Cell(0, 9).Alignment = cellCenterCenter
      Grid2.Cell(0, 10).Alignment = cellCenterCenter
      Grid2.Cell(0, 11).Alignment = cellCenterCenter
      Grid2.Cell(0, 12).Alignment = cellCenterCenter
      Grid2.Cell(0, 13).Alignment = cellCenterCenter
      Grid2.Cell(0, 14).Alignment = cellCenterCenter
      Grid2.Cell(0, 15).Alignment = cellCenterCenter
      Grid2.Cell(0, 16).Alignment = cellCenterCenter
    
End Sub

Private Sub leercertificado(rut, NOMBRE, numero)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim j As Double
    
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim TOTAL As Double
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
             
             
             
             
             TOTAL = TOTAL + resultados(1)
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
             Grid2.Cell(LINEA, 2).text = TOTAL
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
    Dim TOTAL As Double
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
             
             TOTAL = TOTAL + resultados(1)
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

Grid2.Column(2).Width = 80
Grid2.Column(3).Width = 80
Grid2.Column(16).Width = 80
Grid2.Column(9).Width = 100


End Sub

Sub cabezas5(rut, NOMBRE, numero)
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
    objReportTitle.text = "CERTIFICADO N° 41 SOBRE SUELDOS Y OTRAS RENTAS SIMILARES DE LOS TRABAJADORES AGRICOLA"
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
Grid2.Column(2).Width = 0
Grid2.Column(3).Width = 0
Grid2.Column(16).Width = 0
Grid2.Column(9).Width = 0
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

Public Function LeerAgricola(rut, MES, año, empresa) As Boolean

    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "agricola"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema + "remu" + empresa + ".mt_fijo"
    
    condicion = "rut='" + rut + "' and año='" + año + "'  "
    If MES <> "" Then condicion = condicion & " and  mes='" + MES + "' "
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    LeerAgricola = False
    
    If sqlconta.status = 0 Then
        LeerAgricola = sqlconta.response(0, 3)
    End If
End Function
