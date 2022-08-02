VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form balance01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Stock de Productos"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14760
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   583
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   984
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Imprime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exportar a Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Width           =   1725
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   6180
      Left            =   -45
      TabIndex        =   2
      Top             =   2250
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   10901
      BackColor       =   16761024
      Caption         =   "Stock de Productos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
         Height          =   5895
         Left            =   45
         TabIndex        =   0
         Top             =   240
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   10398
         BackColorSel    =   16777215
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   3836
      BackColor       =   16744576
      Caption         =   "Informe"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "GENERA PRECIO PROMEDIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   3885
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "GRABAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1680
         Width           =   1725
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Salir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1440
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Genera Informe"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1725
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1095
         Left            =   9960
         TabIndex        =   4
         Top             =   600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1931
         BackColor       =   16761024
         Caption         =   "Fecha Stock"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.TextBox DESDE3 
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
            IMEMode         =   3  'DISABLE
            Left            =   840
            MaxLength       =   4
            TabIndex        =   7
            Tag             =   "fecha"
            Top             =   705
            Width           =   615
         End
         Begin VB.TextBox DESDE2 
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
            Left            =   480
            MaxLength       =   2
            TabIndex        =   6
            Tag             =   "fecha"
            Top             =   705
            Width           =   375
         End
         Begin VB.TextBox DESDE1 
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
            Left            =   120
            MaxLength       =   2
            TabIndex        =   5
            Tag             =   "fecha"
            Top             =   705
            Width           =   375
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AL DIA"
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
            Height          =   285
            Left            =   135
            TabIndex        =   8
            Top             =   405
            Width           =   1335
         End
      End
      Begin MSAdodcLib.Adodc movimientos 
         Height          =   375
         Left            =   0
         Top             =   1920
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   855
         Left            =   0
         TabIndex        =   9
         Top             =   1125
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "BODEGA"
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
         Begin VB.ComboBox combobodega 
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   3855
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   855
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "LOCAL"
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
         Begin VB.ComboBox combolocal 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   3855
         End
      End
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   11640
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
   End
   Begin MSComctlLib.ProgressBar progreso 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8460
      Visible         =   0   'False
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "balance01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private stockinicial As Double
Private BODEGAFILTRO As String
Private localfiltro As String
Private FORMATOGRILLA(20, 20)
Private stockfinal As Double
Private fechacompra As String
Private factor As Double
Private mesipc As String





Private Sub Command1_Click()
rubrolocal = leerrubro(Mid(ComboLOCAL.text, 1, 2))

If leergrabados(Mid(ComboLOCAL.text, 1, 2), DESDE3.text) = True Then
If MsgBox("ya tiene un informe generado desea usarlo ", vbYesNo) = vbYes Then
Call leestockleido(Mid(ComboLOCAL.text, 1, 2), DESDE3.text)
Exit Sub
End If

End If


    Call leestock
    
    Command2.Visible = True
End Sub

Private Sub COMMAND2_Click()
    Titulos

    Grid1.PrintPreview
End Sub

Private Sub Command3_Click()
If Grid1.Rows > 1 Then
Grid1.ExportToExcel ("")
End If


End Sub

Private Sub Command4_Click()
Unload Me
End Sub

 




 

Private Sub Command5_Click()
Dim k, LINEA As Double
If leergrabados(Mid(ComboLOCAL.text, 1, 2), DESDE3.text) = True Then
    If MsgBox("YA TIENE UNA VERSION GUARDADA DESEA ELIMINAR ", vbYesNo) = vbYes Then
        Call borrargrabados(Mid(ComboLOCAL.text, 1, 2), Format(fechasistema, "yyyy"))
        For k = 1 To Grid1.Rows - 2
        LINEA = k
        Call grabarinventarios(Mid(ComboLOCAL.text, 1, 2), DESDE3.text, Grid1.Cell(LINEA, 1).text, Grid1.Cell(LINEA, 2).text, Grid1.Cell(LINEA, 3).text, Grid1.Cell(LINEA, 4).text, Grid1.Cell(LINEA, 5).text, Grid1.Cell(LINEA, 6).text, Grid1.Cell(LINEA, 7).text, Grid1.Cell(LINEA, 9).text, Grid1.Cell(LINEA, 10).text, Grid1.Cell(LINEA, 11).text)
        Next k
    
    End If
Else
For k = 1 To Grid1.Rows - 3
LINEA = k
Call grabarinventarios(Mid(ComboLOCAL.text, 1, 2), DESDE3.text, Grid1.Cell(LINEA, 1).text, Grid1.Cell(LINEA, 2).text, Grid1.Cell(LINEA, 3).text, Grid1.Cell(LINEA, 4).text, Grid1.Cell(LINEA, 5).text, Grid1.Cell(LINEA, 6).text, Grid1.Cell(LINEA, 7).text, Grid1.Cell(LINEA, 8).text, Grid1.Cell(LINEA, 10).text, Grid1.Cell(LINEA, 11).text)
Next k

End If
End Sub

Private Sub Command6_Click()
Dim Precio As Double

rubrolocal = leerrubro(Mid(ComboLOCAL.text, 1, 2))
Call borrarpreciocargado(fechasistema, Mid(ComboLOCAL.text, 1, 2))
For k = 1 To Grid1.Rows - 2
LINEA = k
Precio = CDbl(Grid1.Cell(LINEA, 9).text)
Call grabarPrecios(Grid1.Cell(LINEA, 1).text, Mid(ComboLOCAL.text, 1, 2), DESDE3.text + "-" + DESDE2.text + "-" + DESDE1.text, Grid1.Cell(LINEA, 4).text, Precio, Grid1.Cell(LINEA, 4).text, Precio, Date, USUARIOSISTEMA, "CM")
Next k
Call Command1_Click
End Sub

Private Sub DESDE1_GotFocus()
    Call cargatexto(DESDE1)
End Sub
'
Private Sub DESDE1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE1, DESDE2, KeyCode)
End Sub

Private Sub DESDE1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call ceros(DESDE1): DESDE2.SetFocus
End Sub

Private Sub DESDE2_GotFocus()
    Call cargatexto(DESDE2)
End Sub

Private Sub DESDE2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE1, DESDE3, KeyCode)
End Sub

Private Sub DESDE2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And DESDE1.text <> "" Then Call ceros(DESDE2): DESDE3.SetFocus
End Sub

Private Sub DESDE3_GotFocus()
    Call cargatexto(DESDE3)
End Sub

Private Sub DESDE3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE2, DESDE3, KeyCode)
End Sub

Private Sub DESDE3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And DESDE2.text Then Call ceros(DESDE3)
End Sub


Private Sub Form_Activate()
      
   ' dato1.SetFocus
End Sub

Private Sub Form_Load()
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 800

    DESDE1.text = "31"
    DESDE2.text = "12"
    
    DESDE3.text = Format(fechasistema, "yyyy")

    Command1.Enabled = True
    Command4.Enabled = True
      Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
     Call Conectargestionrubro(Servidor, clientesistema + "gestion" + rubro, Usuario, password)
   
    
    Call LEErlocales(Me)
    Call LEErbodegas
    
    

    Call CARGAGRILLA


End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub CARGAGRILLA()
    Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7
    Grid1.DefaultFont.Bold = False

    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "DESCRIPCION"
    FORMATOGRILLA(1, 3) = "ULTIMA" & vbCrLf & "COMPRA "
    FORMATOGRILLA(1, 4) = "STOCK"
    FORMATOGRILLA(1, 5) = " PRECIO "
    FORMATOGRILLA(1, 6) = " TOTAL  "
    FORMATOGRILLA(1, 7) = "FACTOR "
    FORMATOGRILLA(1, 8) = "REAJUSTE"
    FORMATOGRILLA(1, 9) = "PRECIO " & vbCrLf & "REAJ."
    FORMATOGRILLA(1, 10) = "CORREC." & vbCrLf & "MONETARIA"
    FORMATOGRILLA(1, 11) = "TOTAL " & vbCrLf & "REAJUSTADO"
    FORMATOGRILLA(1, 12) = "COSTO " & vbCrLf & "GENERADO"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "11"
    FORMATOGRILLA(2, 2) = "25"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "12"
    FORMATOGRILLA(2, 7) = "7"
    FORMATOGRILLA(2, 8) = "0"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "12"
    FORMATOGRILLA(2, 11) = "12"
    FORMATOGRILLA(2, 12) = "12"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "D"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    
    Rem FORMATO GRILLA
  
    FORMATOGRILLA(4, 4) = "#,###,##0.00"
    FORMATOGRILLA(4, 5) = "#,###,###,##0.00"
    FORMATOGRILLA(4, 6) = "#,###,###,##0"
    
    FORMATOGRILLA(4, 7) = "#,###,##0.000"
    FORMATOGRILLA(4, 8) = "#,###,##0.00"
    FORMATOGRILLA(4, 9) = "#,###,##0.00"
    FORMATOGRILLA(4, 10) = "#,###,###,##0"
    FORMATOGRILLA(4, 11) = "#,###,###,##0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    
    Grid1.Cols = 13
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
    Grid1.BackColorFixedSel = RGB(110, 180, 214)
    Grid1.BackColorBkg = RGB(90, 158, 214)
    Grid1.BackColorScrollBar = RGB(231, 235, 247)
    Grid1.BackColor1 = RGB(231, 235, 247)
    Grid1.BackColor2 = RGB(239, 243, 255)
    Grid1.GridColor = RGB(148, 190, 231)
   
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
    Next k
    Grid1.Column(0).Width = 0
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterGeneral
    Grid1.RowHeight(0) = 30
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).WrapText = True
    Grid1.Column(12).CellType = cellCheckBox
    
End Sub

Sub leestock()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim LINEA As Double
    Dim saldo As Double
    Dim canti As Double
    Dim fechade As String
    Dim fechaha As String
    Dim palabra As String
    Dim filtro As String
    Dim TOTAL1 As Double
    Dim total2 As Double
    Dim FILTRO2 As String
    Dim stock As Double
    Dim Precio As Double
    Dim TOTAL As Double
    Dim iva As Double
    Dim CAJAS As Double
    Dim uniemba As Double
    Dim bases As String
    Dim prome As Double
   
  
    Dim NUEVO As Double
    Dim CORRECCION As Double
    
    CARGAGRILLA
    iva = 1.19
    
        Set csql.ActiveConnection = gestion
         bases = clientesistema + "gestion" + rubrolocal + "."
        csql.sql = "SELECT mpf.codigobarra,mpf.descripcion,mpf.importado "
        csql.sql = csql.sql + "FROM " & bases & "r_maestroproductos_fijo_" + rubrolocal + " as mpf order by mpf.descripcion "
        csql.Execute
        
        progreso.Visible = True
        progreso.Min = 0
        progreso.Value = 0
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            Grid1.AutoRedraw = False
            progreso.Max = csql.RowsAffected
            While Not resultados.EOF
            stock = leersaldomo(resultados(0), DESDE3.text + "-" + DESDE2.text + "-" + DESDE1.text, Mid(ComboLOCAL.text, 1, 2), Mid(combobodega.text, 1, 2))
            If stock <> 0 Then
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
         
            fechacompra = leerultimacompra(resultados(0), DESDE3.text + "-" + DESDE2.text + "-" + DESDE1.text)
            
            prome = leerpreciomayor(resultados(0), fechacompra) / iva
            Grid1.Cell(Grid1.Rows - 1, 3).text = fechacompra
            Grid1.Cell(Grid1.Rows - 1, 4).text = stock
            Grid1.Cell(Grid1.Rows - 1, 5).text = prome
            Grid1.Cell(Grid1.Rows - 1, 6).text = Round(prome * stock, 0)
            
            Rem If fechacompra < DESDE3.text + "-01-01" Then factor = 1.08
            
            factor = leeripc(mesipc, Format(fechasistema, "yyyy"), resultados(2))
            Grid1.Cell(Grid1.Rows - 1, 7).text = factor
            CORRECCION = 0
            If factor > 0 Then
            CORRECCION = prome * (factor - 1)
            End If
            NUEVO = CORRECCION + prome
            Grid1.Cell(Grid1.Rows - 1, 8).text = CORRECCION
            Grid1.Cell(Grid1.Rows - 1, 9).text = NUEVO
            Grid1.Cell(Grid1.Rows - 1, 10).text = Round(CORRECCION * stock, 0)
            
            Grid1.Cell(Grid1.Rows - 1, 11).text = Round(NUEVO * stock, 0)
            TOTAL1 = TOTAL1 + Round((prome * stock), 0)
            total2 = total2 + Round(NUEVO * stock, 0)
            
Rem                  stock = leersaldomo(resultados(0), DESDE3.text + "-" + DESDE2.text + "-" + DESDE1.text, Mid(combolocal.text, 1, 2), Mid(combobodega.text, 1, 2))
             Rem End If
             
    Rem                    Precio = leerpromedio(resultados(0), DESDE3.text + "-" + DESDE2.text + "-" + DESDE1.text) / iva
      
            End If
            progreso.Value = progreso.Value + 1
            progreso.Refresh
            
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            Command2.Visible = True
           
        Else
            MsgBox "No se encontraron resultados para la búsqueda, elija otro criterio e intente nuevamente.", vbInformation + vbOKOnly, "Consulta sin Resultados"
        End If

' SUMA FINAL

If Grid1.Rows > 1 Then
'

 Grid1.Rows = Grid1.Rows + 1
                progreso.Max = progreso.Max + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 10).FontBold = True
        
        
        Grid1.Cell(Grid1.Rows - 1, 2).text = "TOTAL GENERAL VALORIZADO "

        Grid1.Cell(Grid1.Rows - 1, 6).Border(cellEdgeTop) = cellThin
        
        Grid1.Cell(Grid1.Rows - 1, 10).Border(cellEdgeTop) = cellThin
        Grid1.Cell(Grid1.Rows - 1, 11).Border(cellEdgeTop) = cellThin
        
        
    
        Grid1.Cell(Grid1.Rows - 1, 6).text = TOTAL1
        Grid1.Cell(Grid1.Rows - 1, 10).text = total2 - TOTAL1
        
        Grid1.Cell(Grid1.Rows - 1, 11).text = total2
        
        TOTAL1 = 0
        total2 = 0
'        Grid1.Rows = Grid1.Rows + 1
'        progreso.Max = progreso.Max + 1
'
       End If
        
        Command1.Enabled = True
        progreso.Visible = False
        Grid1.AutoRedraw = True
        Grid1.Refresh


End Sub

Sub leestockleido(loc, año)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim LINEA As Double
    Dim saldo As Double
    Dim canti As Double
    Dim fechade As String
    Dim fechaha As String
    Dim palabra As String
    Dim filtro As String
    Dim TOTAL1 As Double
    Dim total2 As Double
    Dim FILTRO2 As String
    Dim stock As Double
    Dim Precio As Double
    Dim TOTAL As Double
    Dim iva As Double
    Dim CAJAS As Double
    Dim uniemba As Double
    Dim bases As String
    Dim prome As Double
   
  
    Dim NUEVO As Double
    Dim CORRECCION As Double
    
    CARGAGRILLA
    iva = 1.19
    
        Set csql.ActiveConnection = contadb
         
        csql.sql = "SELECT * "
        csql.sql = csql.sql + "FROM  inventario where local='" + loc + "' and año='" + año + "' "
        csql.Execute
        
        progreso.Visible = True
        progreso.Min = 0
        progreso.Value = 0
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            Grid1.AutoRedraw = False
            progreso.Max = csql.RowsAffected
            While Not resultados.EOF
            If resultados(2) <> "" Then
            Grid1.Rows = Grid1.Rows + 1
            
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(2)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(3)
            Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(4)
            Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(5)
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(6)
            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(7)
            Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(8)
            Grid1.Cell(Grid1.Rows - 1, 9).text = resultados(9)
            Grid1.Cell(Grid1.Rows - 1, 10).text = resultados(10)
            Grid1.Cell(Grid1.Rows - 1, 11).text = resultados(11)
            If preciocargado(resultados(2), DESDE3.text + "-" + DESDE2.text + "-" + DESDE1.text, loc) = True Then
            Grid1.Cell(Grid1.Rows - 1, 12).text = "1"
            Else
            Grid1.Cell(Grid1.Rows - 1, 12).text = "0"
            
            End If
            
            
            TOTAL1 = TOTAL1 + resultados(7)
            total2 = total2 + resultados(11)
            
            End If
            
            
            
            progreso.Value = progreso.Value + 1
            progreso.Refresh
            
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            Command2.Visible = True
           
        Else
            MsgBox "No se encontraron resultados para la búsqueda, elija otro criterio e intente nuevamente.", vbInformation + vbOKOnly, "Consulta sin Resultados"
        End If

' SUMA FINAL

If Grid1.Rows > 1 Then
'

 Grid1.Rows = Grid1.Rows + 1
                progreso.Max = progreso.Max + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 10).FontBold = True
        
        
        Grid1.Cell(Grid1.Rows - 1, 2).text = "TOTAL GENERAL VALORIZADO "

        Grid1.Cell(Grid1.Rows - 1, 6).Border(cellEdgeTop) = cellThin
        
        Grid1.Cell(Grid1.Rows - 1, 10).Border(cellEdgeTop) = cellThin
        Grid1.Cell(Grid1.Rows - 1, 11).Border(cellEdgeTop) = cellThin
        
        
    
        Grid1.Cell(Grid1.Rows - 1, 6).text = TOTAL1
        Grid1.Cell(Grid1.Rows - 1, 10).text = total2 - TOTAL1
        
        Grid1.Cell(Grid1.Rows - 1, 11).text = total2
        
        TOTAL1 = 0
        total2 = 0
'        Grid1.Rows = Grid1.Rows + 1
'        progreso.Max = progreso.Max + 1
'
       End If
        
        Command1.Enabled = True
        progreso.Visible = False
        Grid1.AutoRedraw = True
        Grid1.Refresh


End Sub

 
 

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub Titulos()
    

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.PageSetup.Orientation = cellPortrait
    
    Grid1.DefaultFont.Size = 7.5
   
    Grid1.PageSetup.PrintFixedRow = True
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.PrintTitleRows = 0
    Grid1.PageSetup.BlackAndWhite = True
    
    
    
    
    
    
    'Logo
'    GRID1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    GRID1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid1.PageSetup.HeaderAlignment = CellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE STOCK VALORIZADO " + combobodega.text
    
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ComboLOCAL.text + "  PRECIO COSTO NETO"
    
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "AL DIA :  " & Format(DESDE1.text & "-" & DESDE2.text & "-" & DESDE3.text, "dd-mm-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
        
        
    Set objReportTitle = New FlexCell.ReportTitle
  
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Underline = True
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D " + "usuario:" + USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    Grid1.PageSetup.LeftMargin = 0.2
    Grid1.PageSetup.RightMargin = 0.2
    
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
    
    
    
    
    
End Sub




Sub LEErbodegas()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim locales As String
    Dim bases As String
    
    locales = Mid(ComboLOCAL.text, 1, 2)
    
            Set csql.ActiveConnection = gestion
           If rubro = "" Then
           Exit Sub
           End If
           
            bases = clientesistema + "gestion" + rubro + "."
        
        csql.sql = "SELECT codigobodega,nombre "
        csql.sql = csql.sql + "FROM " & bases & "r_maestrobodegas_" + rubro + " "
        csql.sql = csql.sql + "where local='" + locales + "' "
        csql.sql = csql.sql + "ORDER BY codigobodega "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                combobodega.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            combobodega.AddItem ("99 TODAS")
            Set resultados = Nothing
            combobodega.text = combobodega.List(0)
            
        End If
    
        
End Sub
 
Sub LEErlocales(ByRef frm As Form)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim PASO As Integer
    Dim i As Integer
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas "
        csql.sql = csql.sql + "WHERE codigocontable='" + empresaactiva + "' ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            i = 0
            While Not resultados.EOF
                i = i + 1
                frm.ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                If resultados(0) = empresaactiva Then
                PASO = i - 1
                End If
                
                resultados.MoveNext
                
            Wend
            resultados.Close
            Set resultados = Nothing
        frm.ComboLOCAL.text = frm.ComboLOCAL.List(PASO)
        End If
        localfiltro = Mid(frm.ComboLOCAL.text, 1, 2)
        
End Sub

Public Function leersaldomo(codigo, hasta, loc, bodega) As Double
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim añode As String
 

añode = "2007"
  
Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT saldoanterior + ifnull((SELECT sum(if (mt.operacion ='+',lmd.unidades,lmd.unidades*-1)) FROM " & clientesistema & "gestion" & rubrolocal & ".l_movimientos_detalle_" + loc + " as lmd ," + clientesistema
        csql.sql = csql.sql + "gestion.g_maestrotipodedocumentos as mt WHERE mt.tipos=lmd.tipo and lmd.codigo=mps.codigo and lmd.bodega = mps.bodega AND lmd.fecha >='" & añode & "-01-01" + "' AND lmd.fecha <='" + hasta + "' group BY lmd.codigo ORDER BY lmd.codigo),'0') as saldo "
        csql.sql = csql.sql + "FROM " & clientesistema & "gestion" & rubrolocal & ".r_maestroproductos_stock_" + rubrolocal + " as mps "
        csql.sql = csql.sql + "WHERE local = '" & loc & "' "
'        If bodega <> "99" Then
'        csql.sql = csql.sql + "AND bodega = '" & bodega & "' "
'        End If
        csql.sql = csql.sql + "AND año = '" & añode & "' "
        csql.sql = csql.sql + "and codigo='" + codigo + "' order by  año asc  limit 0,1"
        csql.Execute
        leersaldomo = 0
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
            leersaldomo = resultados(0)
            resultados.Close
        Set resultados = Nothing
            
    End If
    
    csql.Close
    Set csql = Nothing

End Function
Public Function leerrubro(empresa) As String
Dim csql As New rdoQuery
Dim resultados  As rdoResultset
leerrubro = ""
Set csql.ActiveConnection = gestion
csql.sql = "select rubro from g_maestroempresas where codigo='" & empresa & "' "
csql.Execute
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leerrubro = resultados(0)
End If
csql.Close
Set csql = Nothing


End Function

Function leerpromedio(codigo, fecha) As Double

Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String

    Set csql.ActiveConnection = gestion
    bases = clientesistema + "gestion" + rubrolocal + "."
csql.sql = "select preciopromedio,fecha from "
csql.sql = csql.sql & bases & "l_maestroproductos_costos_local_" & Mid(ComboLOCAL.text, 1, 2) + " "
csql.sql = csql.sql & " where codigo='" & codigo & "' and fecha <='" & Format(fecha, "yyyy-mm-dd") & "' and local='" & Mid(ComboLOCAL.text, 1, 2) & "' order by fecha desc  limit 0,1 "
csql.Execute
If csql.RowsAffected > 0 Then
 Set resultados = csql.OpenResultset
 
 While Not resultados.EOF
 leerpromedio = resultados(0)
 
 resultados.MoveNext
 Wend
 
 End If
End Function
Function leerultimacompra(codigo, fecha) As String

Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String
 
    Set csql.ActiveConnection = gestion
    bases = clientesistema + "gestion" + rubrolocal + "."
csql.sql = "select fecha from "
csql.sql = csql.sql & bases & "l_movimientos_detalle_" + Mid(ComboLOCAL.text, 1, 2) + " "
csql.sql = csql.sql & " where codigo='" & codigo & "' and fecha between '" + Format(fechasistema, "yyyy") + "-01-01' and '" & Format(fecha, "yyyy-mm-dd") & "' and (tipo='OC' OR TIPO='RL') order by fecha desc limit 0,1 "
csql.Execute
    
    leerultimacompra = Format(fechasistema, "yyyy") - 1 & "-12-31"

If csql.RowsAffected > 0 Then
 Set resultados = csql.OpenResultset
    leerultimacompra = Format(resultados(0), "yyyy-mm-dd")
End If

End Function

Function leerpreciomayor(codigo, fecha) As String

Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String
Dim fecha1 As String
Dim fecha2 As String
Dim FECHA3 As String
Dim fecha4 As String
Dim fecha5 As String
Dim fecha6 As String
 
FECHA3 = Format(fechasistema, "yyyy") & "-01-01"
fecha4 = Format(fechasistema, "yyyy") & "-06-30"
fecha5 = Format(fechasistema, "yyyy") & "-07-01"
fecha6 = Format(fechasistema, "yyyy") & "-12-31"

If fecha < FECHA3 Then
fecha1 = fecha
fecha2 = fecha
leerpreciomayor = leerpromedio(codigo, fecha)
mesipc = "00"
If leerpreciomayor = 0 Then
leerpreciomayor = leerpreciocosto(codigo)

End If

Exit Function
End If
If fecha >= FECHA3 And fecha <= fecha4 Then
fecha1 = Format(fechasistema, "yyyy") + "-01-01"
fecha2 = Format(fechasistema, "yyyy") + "-06-30"
mesipc = "06"
End If
If fecha >= fecha5 And fecha <= fecha6 Then
fecha1 = Format(fechasistema, "yyyy") + "-07-01"
fecha2 = Format(fechasistema, "yyyy") + "-12-31"
mesipc = "13"
End If



    Set csql.ActiveConnection = gestion
bases = clientesistema + "gestion" + rubrolocal + "."
csql.sql = "select precio from "
csql.sql = csql.sql & bases & "l_movimientos_detalle_" & Mid(ComboLOCAL.text, 1, 2) & " "
csql.sql = csql.sql & " where codigo='" & codigo & "' and "
csql.sql = csql.sql & "fecha between '" & Format(fecha1, "yyyy-mm-dd") & "' and '" & Format(fecha2, "yyyy-mm-dd") & "' and (tipo='OC' OR TIPO='RL') order by precio desc limit 0,1 "
csql.Execute
    leerpreciomayor = 0

If csql.RowsAffected > 0 Then
 Set resultados = csql.OpenResultset
    leerpreciomayor = resultados(0)
End If

End Function

Function leerpreciocosto(codigo) As Double

Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String

    Set csql.ActiveConnection = gestion
    bases = clientesistema + "gestion" + rubrolocal + "."
csql.sql = "select pcosto from "
csql.sql = csql.sql & bases & "r_maestroproductos_fijo_" & rubrolocal
csql.sql = csql.sql & " where codigobarra='" & codigo & "' "
csql.Execute
If csql.RowsAffected > 0 Then
 Set resultados = csql.OpenResultset
 
 While Not resultados.EOF
 leerpreciocosto = resultados(0)
 
 resultados.MoveNext
 Wend
 
 End If
End Function

Function leeripc(MES, año, importado) As Double

Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String

    Set csql.ActiveConnection = gestion
    bases = clientesistema + "conta."
csql.sql = "select porcentaje,porcentaje_dolar from "
csql.sql = csql.sql & bases & "ipc"
csql.sql = csql.sql & " where mes='" & MES & "' and año='" + año + "' "
csql.Execute
If csql.RowsAffected > 0 Then
 Set resultados = csql.OpenResultset
 
If importado = 0 Then
 leeripc = 1 + (resultados(0) / 100)
 Else
 leeripc = 1 + (resultados(1) / 100)
 
 End If
 
 
 End If
End Function

Sub grabarinventarios(loc, año, codigo, descripcion, ultimacompra, stock, Precio, TOTAL, factor, precioreajustado, correcion, totalreajustado)
    Dim campos(20, 10) As String
    
    campos(0, 0) = "local"
    campos(1, 0) = "año"
    campos(2, 0) = "codigo"
    campos(3, 0) = "descripcion"
    campos(4, 0) = "ultimacompra"
    campos(5, 0) = "stock"
    campos(6, 0) = "precio"
    campos(7, 0) = "total"
    campos(8, 0) = "factor"
    campos(9, 0) = "precioreajustado"
    campos(10, 0) = "correcion"
    campos(11, 0) = "totalreajustado"
    campos(12, 0) = ""
    
    campos(0, 1) = loc
    campos(1, 1) = año
    campos(2, 1) = codigo
    campos(3, 1) = descripcion
    campos(4, 1) = Format(ultimacompra, "yyyy-mm-dd")
    campos(5, 1) = Replace(stock, ",", ".")
    campos(6, 1) = Replace(Precio, ",", ".")
    campos(7, 1) = Replace(TOTAL, ",", ".")
    campos(8, 1) = Replace(factor, ",", ".")
    campos(9, 1) = Replace(precioreajustado, ",", ".")
    campos(10, 1) = Replace(correcion, ",", ".")
    campos(11, 1) = Replace(totalreajustado, ",", ".")
    
    campos(0, 2) = "inventario"
    condicion = ""
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    sqlconta.audit = False
    
    Call sqlconta.sqlconta(op, condicion)
    status = sqlconta.status
    sqlconta.audit = True





End Sub

Function leergrabados(loc, año) As Boolean


Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String

Set csql.ActiveConnection = contadb
    
csql.sql = "select * from "
csql.sql = csql.sql & clientesistema + "conta" + empresaactiva + ".inventario"
csql.sql = csql.sql & " where local='" & loc & "' and año='" + año + "' "
csql.Execute

If csql.RowsAffected > 0 Then
leergrabados = True
Else
leergrabados = False
End If

End Function

Sub grabarPrecios(codigo, loc, fecha, stock, preciocompra, cantidadcompra, preciopromedio, fechacreacion, usuariocreacion, tipo)
    Dim campos(20, 10) As String
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "local"
    campos(2, 0) = "fecha"
    campos(3, 0) = "stock"
    campos(4, 0) = "preciocompra"
    campos(5, 0) = "cantidadcompra"
    campos(6, 0) = "preciopromedio"
    campos(7, 0) = "fechacreacion"
    campos(8, 0) = "usuariocreacion"
    campos(9, 0) = "tipo"
    campos(10, 0) = ""
    
    campos(0, 1) = codigo
    campos(1, 1) = loc
    campos(2, 1) = Format(fecha, "yyyy-mm-dd")
    campos(3, 1) = Replace(stock, ",", ".")
    campos(4, 1) = Replace(preciocompra, ",", ".")
    campos(5, 1) = Replace(cantidadcompra, ",", ".")
    campos(6, 1) = Replace(preciopromedio, ",", ".")
    campos(7, 1) = Format(Date, "yyyy-mm-dd")
    campos(8, 1) = USUARIOSISTEMA
    campos(9, 1) = "CM"
    
    campos(0, 2) = clientesistema + "gestion" + rubrolocal + ".l_maestroproductos_costos_local_" + loc
    condicion = ""
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    sqlconta.audit = False
    
    Call sqlconta.sqlconta(op, condicion)
    status = sqlconta.status
    sqlconta.audit = True





End Sub

Function preciocargado(codigo, fecha, loc) As Boolean


Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String

Set csql.ActiveConnection = contadb
    
csql.sql = "select * from "
csql.sql = csql.sql & clientesistema + "gestion" + rubrolocal + ".l_maestroproductos_costos_local_" + loc + " "
csql.sql = csql.sql & " where codigo='" & codigo & "' and fecha='" + fecha + "' and tipo='CM' and local='" + loc + "' "
csql.Execute

If csql.RowsAffected > 0 Then
preciocargado = True
Else
preciocargado = False
End If

End Function

Sub borrargrabados(loc, año)


Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String

Set csql.ActiveConnection = contadb
    
csql.sql = "delete from "
csql.sql = csql.sql & "inventario"
csql.sql = csql.sql & " where local='" & loc & "' and año='" + año + "' "
csql.Execute
Call sincronizadatos(csql.sql, contadb, "")


End Sub


Sub borrarpreciocargado(fecha, loc)


Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim bases As String

Set csql.ActiveConnection = contadb
    
csql.sql = "delete from "
csql.sql = csql.sql & clientesistema + "gestion" + rubrolocal + ".l_maestroproductos_costos_local_" + loc + " "
csql.sql = csql.sql & " where fecha='" & Format(fecha, "yyyy-mm-dd") & "' and tipo='CM' and local='" + loc + "' "
csql.Execute
Call sincronizadatos(csql.sql, contadb, "")

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

