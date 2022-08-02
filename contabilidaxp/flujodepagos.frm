VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0008 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTA FLUJOS DE PAGO"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14880
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   992
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   19
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   280
         Width           =   1455
      End
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
      TabIndex        =   5
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8925
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15743
      BackColor       =   16744576
      Caption         =   "INFORME CERTIFICADOS DE HONORARIOS"
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
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8280
         Width           =   1365
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   8
         Top             =   360
         Width           =   14640
         _ExtentX        =   25823
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
         Begin VB.CommandButton cmdguardar 
            Caption         =   "Guardar"
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
            Left            =   9360
            TabIndex        =   4
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox DATO5 
            BackColor       =   &H00E1FFFD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7560
            TabIndex        =   3
            Tag             =   "fecha"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox DATO2 
            BackColor       =   &H00E1FFFD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "fecha"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox DATO3 
            BackColor       =   &H00E1FFFD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5280
            MaxLength       =   2
            TabIndex        =   1
            Tag             =   "fecha"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox DATO4 
            BackColor       =   &H00E1FFFD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5640
            MaxLength       =   4
            TabIndex        =   2
            Tag             =   "fecha"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
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
            Left            =   13080
            TabIndex        =   10
            Top             =   720
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   135
            TabIndex        =   11
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
               TabIndex        =   12
               Top             =   270
               Width           =   2865
            End
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "La fecha ingresada debe ser Lunes"
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
            Left            =   3840
            TabIndex        =   17
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MONTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6600
            TabIndex        =   15
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FECHA :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3840
            TabIndex        =   14
            Top             =   600
            Width           =   975
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   135
         TabIndex        =   7
         Top             =   1485
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   11774
         BackColor       =   16744576
         Caption         =   "LISTADO DE PRESUPUESTOS INGRESADOS"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid GRID1 
            Height          =   6330
            Left            =   90
            TabIndex        =   13
            Top             =   225
            Width           =   14550
            _ExtentX        =   25665
            _ExtentY        =   11165
            Cols            =   5
            DefaultFontName =   "Arial"
            DefaultFontSize =   8.25
            FixedRowColStyle=   0
            Rows            =   30
            SelectionMode   =   1
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Doble clic sobre la grilla para modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   8280
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPR sobre la grilla para eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   16
         Top             =   8280
         Width           =   2895
      End
   End
End
Attribute VB_Name = "prove0008"
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
Private esta As String
Private Sub Command3_Click()
Dim D1 As String
Dim D2 As String
Dim D3 As String
Dim D4 As String
Dim D5 As String
Close 10

Open "F1879_" + empresaactiva + ".TXT" For Output As #10
For k = 1 To Grid1.Rows - 2
D1 = CDbl(Mid(Grid1.Cell(k, 1).text, 2, 8)) & Mid(Grid1.Cell(k, 1).text, 10, 1)

D2 = Format(Grid1.Cell(k, 6).text, "000000000000")
D3 = "000000000000"
D4 = "000000000000"
D5 = Format(CDbl(Grid1.Cell(k, 7).text), "0000000")
Print #10, D1 + ";" + D2 + ";" + D3 + ";" + D4 + ";" + D5
Next k
Close #10
Shell "NOTEPAD " + "F1879_" + empresaactiva + ".TXT"




End Sub


Private Sub Command4_Click()
End Sub

Private Sub cmdguardar_Click()
If IsDate(dato2.text & "-" & dato3.text & "-" & dato4.text) = True And DATO5.text <> "" Then
    Call guardarmonto(dato4.text & "-" & dato3.text & "-" & dato2.text, DATO5.text, empresaactiva, esta)
    Call limpia
    Call leer
End If

End Sub

Private Sub Command1_Click()
If Grid1.Rows > 1 Then
    imprimir
End If

End Sub

Private Sub COMMAND2_Click()
leer
End Sub

Private Sub command8_Click()

End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato3, KeyCode)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
snum = 0: KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    Call ceros(dato2)
    If dato2.text = "00" Then dato2.text = Format(fechasistema, "dd")
    dato3.SetFocus
End If
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
snum = 0: KeyAscii = esNumero(KeyAscii)
 If KeyAscii = 13 Then
    Call ceros(dato3)
    If dato3.text = "00" Then dato3.text = Format(fechasistema, "mm")
    dato4.SetFocus
End If
End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, dato4, KeyCode)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
Dim dia As String
snum = 0: KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
    Call ceros(dato4)
    If dato4.text = "0000" Then
        dato4.text = Format(fechasistema, "yyyy")
    End If
    If IsDate(dato2.text & "-" & dato3.text & "-" & dato4.text) = True Then
        dia = Weekday(dato4.text & "/" & dato3.text & "/" & dato2.text, vbMonday)
        If dia <> "1" Then
            MsgBox "Fecha ingresada No corresponde a un Dia Lunes", vbCritical, "Atencion"
            dato2.text = ""
            dato3.text = ""
            dato4.text = ""
            dato2.SetFocus
        Else
            Call leerfecha(dato4.text & "-" & dato3.text & "-" & dato2.text, empresaactiva)
            DATO5.SetFocus
        End If
    Else
        MsgBox "FECHA NO ES VALIDA", vbCritical, "ATENCION"
    End If
  End If

End Sub
Sub leerfecha(fecha, localconsulta)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = conta
    csql.sql = "select montoplazo from maximopagoproveedores where fecha='" & fecha & "' and empresa='" & localconsulta & "' "
    csql.Execute
    esta = "0"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        DATO5.text = resultados(0)
        esta = "1"
    End If
End Sub


Private Sub dato5_KeyPress(KeyAscii As Integer)
 snum = 0: KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And DATO5.text <> "" Then
    If CDbl(DATO5.text) > 0 Then
        cmdguardar.SetFocus
    End If
End If
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
  
Sub limpia()
dato2.text = ""
dato3.text = ""
dato4.text = ""
DATO5.text = ""
dato2.SetFocus
    
End Sub

Sub imprimir()
Dim titulo As String
titulo = "PRESUPUESTOS DE PAGOS"
Call CABEZAS2(titulo, "N", "000000000")
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait

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
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "PRESUPUESTO"
    FORMATOGRILLA(1, 3) = "USADO"
    FORMATOGRILLA(1, 4) = "DISPONIBLE"
    FORMATOGRILLA(1, 5) = "HONORARIOS"
    FORMATOGRILLA(1, 6) = "RETENCION"
    FORMATOGRILLA(1, 7) = "NUMERO"
    FORMATOGRILLA(1, 8) = "IMPRIMIR"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "30"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 2) = "##,###,##0"
    FORMATOGRILLA(4, 4) = "##,###,##0"
    FORMATOGRILLA(4, 5) = "##,###,##0"
    FORMATOGRILLA(4, 6) = "##,###,##0"
    
    
    Rem LOCCKED
    For k = 1 To 8
    FORMATOGRILLA(5, k) = "TRUE"
    
    Next k
        
    
    Grid1.Cols = 5
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
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim USADO As Double
    
    Dim fechasum As String
    Dim total2 As Double
    Dim tila1 As Double
    Dim tila2 As Double
    Dim tila3 As Double
    Dim total3 As Double
    Dim total4 As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    
        Set csql.ActiveConnection = conta
        csql.sql = "select fecha,montoplazo from maximopagoproveedores where empresa='" + empresaactiva + "' and fecha>='" + dato4.text + "-" + dato3.text + "-" + dato2.text + "' "
        csql.Execute
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
             LINEA = LINEA + 1
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
                USADO = 0
             Grid1.Cell(LINEA, 3).text = USADO
             Grid1.Cell(LINEA, 4).text = resultados(1) - USADO
                 
             total = total + resultados(1)
             
             resultados.MoveNext
            
            Wend
             LINEA = LINEA + 1
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Range(LINEA, 1, LINEA, Grid1.Cols - 1).FontBold = True
             Grid1.Range(LINEA, 1, LINEA, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
             Grid1.Cell(LINEA, 3).text = total
             
         
         resultados.Close
            Set resultados = Nothing

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


Private Sub Grid1_DblClick()
If Grid1.Rows > 1 Then
     If Grid1.ActiveCell.row > 0 Then
         
         dato2.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 2)
         dato3.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 4, 2)
         dato4.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 7, 4)
         DATO5.text = Grid1.Cell(Grid1.ActiveCell.row, 2).text
         esta = "1"
         DATO5.SetFocus
         
    End If
End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
 Select Case KeyCode
            Case 46
            If Verifica_Permiso(Me.Caption, "elimina") = True Then
                If Grid1.ActiveCell.row > 0 Then
                    Call ELIMINAR(Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "yyyy-mm-dd"), empresaactiva)
                    Grid1.RemoveItem (Grid1.ActiveCell.row)
                    Call leer
                    
                End If
            End If
            
        End Select
End Sub
Sub ELIMINAR(fecha, empresa)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = conta
    csql.sql = "delete from maximopagoproveedores where fecha='" & fecha & "' and empresa='" & empresa & "'"
    csql.Execute
    
    csql.Close
    Call sincronizadatos(csql.sql, conta, "")
    
    
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
