VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form confi08 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fecha de Cierre"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   525
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   926
      BackColor       =   49344
      Caption         =   "Año"
      CaptionEstilo3D =   1
      BackColor       =   49344
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.ComboBox comboaño 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   200
         Width           =   1815
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   3570
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6297
      Cols            =   3
      DefaultFontSize =   8.25
      Rows            =   13
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1005
      BackColor       =   49344
      Caption         =   "Empresa"
      CaptionEstilo3D =   1
      BackColor       =   49344
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.ComboBox comboempresa 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Doble Click Sobre la Grilla Para Modificar"
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
      Left            =   600
      TabIndex        =   5
      Top             =   5160
      Width           =   3495
   End
End
Attribute VB_Name = "confi08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Sub CARGAGRILLA()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "MES"
    formatogrilla2(1, 2) = "ESTADO"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    formatogrilla2(2, 2) = "5"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    Rem FORMATO GRILLA
    formatogrilla2(4, 1) = ""
    formatogrilla2(4, 2) = ""
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "FALSE"
    
    Grid1.Cols = 3
    Grid1.Rows = 13
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
    Grid1.SelectionMode = cellSelectionFree
    
     For k = 1 To 12
     Grid1.Cell(k, 1).text = UCase(MonthName(k))
     Next k
     Grid1.Column(2).CellType = cellCheckBox
     Grid1.Refresh
    End Sub

Private Sub comboaño_Click()
    Call leer(Mid(comboempresa.text, 1, 2), COMBOAÑO.text)
End Sub

Private Sub comboempresa_Click()
    Call leer(Mid(comboempresa.text, 1, 2), COMBOAÑO.text)
End Sub

Private Sub Form_Load()
Dim palabras As String
CARGAGRILLA
LEErlocales
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
For k = 1 To comboempresa.ListCount
 palabras = comboempresa.List(k)
If empresaactiva = Mid(palabras, 1, 2) Then
    comboempresa.text = comboempresa.List(k)
    Exit For
End If
Next k

Call leer(Mid(comboempresa.text, 1, 2), COMBOAÑO.text)

 
End Sub
Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigoempresa,nombre "
        csql.sql = csql.sql + "FROM maestroempresas "
        csql.sql = csql.sql + "ORDER BY codigoempresa "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                comboempresa.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        comboempresa.text = comboempresa.List(0)
        End If
End Sub

Sub grabar(MES, estado, año, empresa)
    campos(0, 0) = "empresa"
    campos(1, 0) = "mes"
    campos(2, 0) = "año"
    campos(3, 0) = "estado"
    campos(4, 0) = ""
    campos(0, 1) = empresa
    campos(1, 1) = MES
    campos(2, 1) = año
    campos(3, 1) = estado
    campos(0, 2) = "fechacierre"
    condicion = "empresa='" & empresa & "' and año='" & año & "' and mes='" & MES & "'"
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    
        op = 5
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 4 Then
            op = 2
            condicion = ""
        Else
            op = 3
        End If
        Call sqlconta.sqlconta(op, condicion)
End Sub
 

Private Sub Grid1_Click()
Call grabar(Grid1.ActiveCell.row, Grid1.Cell(Grid1.ActiveCell.row, 2).text, COMBOAÑO.text, Mid(comboempresa.text, 1, 2))
End Sub

Sub leer(empresa, año)
Dim k As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = conta
csql.sql = "select mes,estado from " + clientesistema + "conta.fechacierre "
csql.sql = csql.sql & "where empresa='" & empresa & "' and año='" & año & "' "
csql.sql = csql.sql & "order by mes "
csql.Execute


If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    For k = 1 To 12
        Grid1.Cell(k, 2).text = "0"
    Next k
    While Not resultados.EOF
        Grid1.Cell(resultados(0), 2).text = resultados(1)
        resultados.MoveNext
    Wend
Else
    For k = 1 To 12
        Grid1.Cell(k, 2).text = "0"
    Next k
    
End If
Set resultados = Nothing
csql.Close
Set csql = Nothing
Grid1.Refresh
End Sub

