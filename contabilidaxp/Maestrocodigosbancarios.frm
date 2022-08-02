VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Maestrocodigosbancarios 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6930
   ClientLeft      =   2565
   ClientTop       =   2595
   ClientWidth     =   14100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   14100
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp frmDatos 
      Height          =   6915
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12197
      BackColor       =   16744576
      Caption         =   "Maestro Codigos Bancarios"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ColorBarraAbajo =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   10080
         TabIndex        =   22
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
         Begin VB.CommandButton botonmisfavoritos 
            Caption         =   "Mis Favoritos"
            Height          =   255
            Left            =   1800
            TabIndex        =   24
            Top             =   280
            Width           =   1335
         End
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   280
            Width           =   1455
         End
      End
      Begin VB.CommandButton impimir 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6120
         Width           =   1935
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13140
         MaxLength       =   1
         TabIndex        =   17
         Top             =   5760
         Width           =   795
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         MaxLength       =   9
         TabIndex        =   16
         Top             =   6480
         Width           =   1875
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   15
         Top             =   5760
         Width           =   4995
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   12
         Top             =   5760
         Width           =   1155
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         MaxLength       =   8
         TabIndex        =   11
         Top             =   5760
         Width           =   2130
      End
      Begin VB.TextBox dato7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         MaxLength       =   100
         TabIndex        =   2
         Top             =   6480
         Width           =   9165
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         MaxLength       =   3
         TabIndex        =   0
         Top             =   5760
         Width           =   1155
      End
      Begin XPFrame.FrameXp frmLista 
         Height          =   4875
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   8599
         BackColor       =   16761024
         Caption         =   "Lista de Codigos"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ColorBarraArriba=   12648447
         ColorBarraAbajo =   16711680
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
         Begin FlexCell.Grid lista 
            Height          =   4395
            Left            =   90
            TabIndex        =   3
            Top             =   420
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   7752
            BackColorFixed  =   16761024
            Cols            =   5
            DefaultFontSize =   9.75
            Rows            =   1
            SelectionMode   =   1
         End
         Begin MSAdodcLib.Adodc data 
            Height          =   330
            Left            =   60
            Top             =   4560
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   13680
         TabIndex        =   5
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         BackColor       =   16711680
         Caption         =   "X"
         BackColor       =   16711680
         ForeColor       =   0
         ColorBarraAbajo =   16761024
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
      End
      Begin VB.Label lbldv 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   360
         Left            =   2160
         TabIndex        =   20
         Top             =   6480
         Width           =   465
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   6120
         Width           =   2475
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   13140
         TabIndex        =   18
         Top             =   5400
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   4680
         TabIndex        =   14
         Top             =   5400
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   5880
         TabIndex        =   13
         Top             =   5400
         Width           =   4995
      End
      Begin VB.Label lblbanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1380
         TabIndex        =   10
         Top             =   5760
         Width           =   3225
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   6120
         Width           =   9165
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   10920
         TabIndex        =   8
         Top             =   5400
         Width           =   2130
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   5400
         Width           =   3225
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   5400
         Width           =   1155
      End
   End
End
Attribute VB_Name = "Maestrocodigosbancarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private FORMATOGRILLA(10, 10) As String
    Private modifica As Boolean

 

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
 
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
        End If
    End Sub
    
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato1.text <> "" Then
            Call ceros(dato1)
            lblBanco.Caption = leebanco(dato1.text)
            dato2.SetFocus
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        Dim Precio As String
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato2.text <> "" Then
            Call ceros(dato2)
             If lblBanco.Caption <> "" Then
                If existecodigobancario(dato1.text, dato2.text) = False Then
                dato3.SetFocus
                modifica = False
                Else
                dato3.SetFocus
                modifica = True
                End If
             End If
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato3.text <> "" Then
            dato4.SetFocus
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
   
'MANEJO DE LOS CONTOLES
'============================================================

'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaLista(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        FORMATOGRILLA(1, 1) = "BANCO"
        FORMATOGRILLA(1, 2) = "CODIGO"
        FORMATOGRILLA(1, 3) = "NOMBRE"
        FORMATOGRILLA(1, 4) = "C. CONTABLE"
        FORMATOGRILLA(1, 5) = "TIPO"
        FORMATOGRILLA(1, 6) = "GLOSA"
        FORMATOGRILLA(1, 7) = "RUT"
        
        
        Rem LARGO DE LOS DATOS
        FORMATOGRILLA(2, 1) = "4"
        FORMATOGRILLA(2, 2) = "5"
        FORMATOGRILLA(2, 3) = "30"
        FORMATOGRILLA(2, 4) = "8"
        FORMATOGRILLA(2, 5) = "4"
        FORMATOGRILLA(2, 6) = "50"
        FORMATOGRILLA(2, 7) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        FORMATOGRILLA(3, 1) = "N"
        FORMATOGRILLA(3, 2) = "N"
        FORMATOGRILLA(3, 3) = "S"
        FORMATOGRILLA(3, 4) = "N"
        FORMATOGRILLA(3, 5) = "S"
        FORMATOGRILLA(3, 6) = "S"
        FORMATOGRILLA(3, 7) = "N"
        
        Rem FORMATO GRILLA
        FORMATOGRILLA(4, 1) = "000"
        FORMATOGRILLA(4, 2) = "00000"
        FORMATOGRILLA(4, 3) = ""
        FORMATOGRILLA(4, 4) = ""
        FORMATOGRILLA(4, 5) = ""
        FORMATOGRILLA(4, 6) = ""
        FORMATOGRILLA(4, 7) = "0000000000"
        
        Rem LOCCKED
        FORMATOGRILLA(5, 1) = "TRUE"
        FORMATOGRILLA(5, 2) = "TRUE"
        FORMATOGRILLA(5, 3) = "TRUE"
        FORMATOGRILLA(5, 4) = "TRUE"
        FORMATOGRILLA(5, 5) = "TRUE"
        FORMATOGRILLA(5, 6) = "TRUE"
        FORMATOGRILLA(5, 7) = "TRUE"
        
        Rem VALOR MINIMO
        FORMATOGRILLA(6, 1) = ""
        FORMATOGRILLA(6, 2) = ""
        FORMATOGRILLA(6, 3) = ""
        FORMATOGRILLA(6, 4) = ""
        FORMATOGRILLA(6, 5) = ""
        FORMATOGRILLA(6, 6) = ""
        FORMATOGRILLA(6, 7) = ""
        
        Rem VALOR MAXIMO
        FORMATOGRILLA(7, 1) = ""
        FORMATOGRILLA(7, 2) = ""
        FORMATOGRILLA(7, 3) = ""
        FORMATOGRILLA(7, 4) = ""
        FORMATOGRILLA(7, 5) = ""
        FORMATOGRILLA(7, 6) = ""
        FORMATOGRILLA(7, 7) = ""
        
        Rem ANCHO
        FORMATOGRILLA(8, 1) = "4"
        FORMATOGRILLA(8, 2) = "5"
        FORMATOGRILLA(8, 3) = "25"
        FORMATOGRILLA(8, 4) = "8"
        FORMATOGRILLA(8, 5) = "4"
        FORMATOGRILLA(8, 6) = "25"
        FORMATOGRILLA(8, 7) = "10"
    
            
        lista.Cols = col
        lista.Rows = row
        lista.AllowUserResizing = False
        lista.DisplayFocusRect = False
        lista.ExtendLastCol = True
        lista.BoldFixedCell = False
        lista.DrawMode = cellOwnerDraw
        lista.Appearance = Flat
        lista.ScrollBarStyle = Flat
        lista.FixedRowColStyle = Flat
'        lista.BackColorFixed = RGB(90, 158, 214)
'        lista.BackColorFixedSel = RGB(110, 180, 230)
'        lista.BackColorBkg = RGB(90, 158, 214)
'        lista.BackColorScrollBar = RGB(231, 235, 247)
'        lista.BackColor1 = RGB(231, 235, 247)
'        lista.BackColor2 = RGB(239, 243, 255)
'        lista.GridColor = RGB(148, 190, 231)
        
        lista.Column(0).Width = 0
        For i = 1 To col - 1
            lista.Cell(0, i).text = FORMATOGRILLA(1, i)
            lista.Column(i).Width = Val(FORMATOGRILLA(8, i)) * (lista.Cell(0, i).Font.Size + 1.25)
            lista.Column(i).MaxLength = Val(FORMATOGRILLA(2, i))
            lista.Column(i).FormatString = FORMATOGRILLA(4, i)
            lista.Column(i).Locked = FORMATOGRILLA(5, i)
            If FORMATOGRILLA(3, i) = "N" Then
                lista.Column(i).Alignment = cellRightCenter
            Else
                lista.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        lista.Range(0, 1, 0, lista.Cols - 1).Alignment = cellCenterCenter
        lista.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================
    Private Sub leerEspeciales()
        Dim tabla As String
        tabla = "SELECT CONCAT(banco, '" & vbTab & "', codigo, '" & vbTab & "',nombre, '" & vbTab & "', codigocontable, '" & vbTab & "', tipo, '" & vbTab & "', glosa, '" & vbTab & "', rut) AS item "
        tabla = tabla & "FROM cartolasbancarias_codigoscontables  "
        tabla = tabla & "ORDER BY banco,codigo ASC"
        Call ConectarControlData(data, Servidor, clientesistema & "conta", Usuario, password, tabla)
        lista.Rows = 1
        lista.AutoRedraw = False
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                lista.AddItem data.Recordset.Fields("item"), True
                data.Recordset.MoveNext
            Wend
        lista.AutoRedraw = True
        lista.Refresh
        End If
    End Sub
'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub grabarEspeciales()
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        campos(0, 0) = "banco"
        campos(1, 0) = "codigo"
        campos(2, 0) = "nombre"
        campos(3, 0) = "codigocontable"
        campos(4, 0) = "tipo"
        campos(5, 0) = "glosa"
        campos(6, 0) = "rut"
        campos(7, 0) = ""
        
        campos(0, 1) = dato1.text
        campos(1, 1) = dato2.text
        campos(2, 1) = dato3.text
        campos(3, 1) = dato4.text
        campos(4, 1) = DATO5.text
        campos(5, 1) = dato7.text
        campos(6, 1) = dato6.text & LBLDV.Caption
        campos(7, 1) = ""
        
        campos(0, 2) = "cartolasbancarias_codigoscontables"
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
    End Sub
'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub modificaEspeciales()
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        campos(0, 0) = "nombre"
        campos(1, 0) = "codigocontable"
        campos(2, 0) = "tipo"
        campos(3, 0) = "glosa"
        campos(4, 0) = "rut"
        campos(5, 0) = ""
        
        campos(0, 1) = dato3.text
        campos(1, 1) = dato4.text
        campos(2, 1) = DATO5.text
        campos(3, 1) = dato7.text
        campos(4, 1) = dato6.text & LBLDV.Caption
        campos(5, 1) = ""
        
        campos(0, 2) = "cartolasbancarias_codigoscontables"
        
        condicion = "banco = '" & dato1.text & "' AND codigo= '" & dato2.text & "'"
        op = 3
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
        modifica = False
    End Sub
'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================
    Private Sub eliminarEspeciales(ByVal banco As String, ByVal codigo As String)
        Dim condicion As String
        Dim campos(1, 3) As String
        Dim op As Integer
        condicion = "banco = '" & banco & "' AND codigo = '" & codigo & "' "
        op = 4
        campos(0, 2) = "cartolasbancarias_codigoscontables "
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================
 
Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato4.text <> "" Then
    Call ceros(dato4)
    DATO5.SetFocus
End If
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 13 And DATO5.text <> "" And (DATO5.text = "A" Or DATO5.text = "C") Then
    dato6.SetFocus
End If
End Sub


Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato6.text <> "" Then
    Call ceros(dato6)
    LBLDV.Caption = rut(dato6.text)
    dato7.SetFocus
End If
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii = 13 And dato7.text <> "" Then
             If modifica = False Then
                Call grabarEspeciales
            Else
                Call modificaEspeciales
            End If
            Call limpia
           Call leerEspeciales
            
        
 End If
End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub
    
    Private Sub Form_Load()
        Call CENTRAR(Me)
        Call CargaGrillaLista(1, 8)
        modifica = False
        cargaLista
    End Sub
    
    Private Sub frmCerrar_BarClick()
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = RAISED
    End Sub
 
Private Sub impimir_Click()
If lista.Rows > 1 Then
Call Titulos("LISTADO DE CODIGOS BANCARIOS ")
lista.PageSetup.Orientation = cellLandscape
lista.PageSetup.HeaderMargin = 0.5
lista.PageSetup.PrintFixedRow = True
lista.PageSetup.TopMargin = 1
lista.PageSetup.LeftMargin = 0.5
lista.PageSetup.RightMargin = 0.5
lista.PageSetup.BottomMargin = 3
lista.PageSetup.FooterMargin = 2
lista.PageSetup.BlackAndWhite = True

lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeTop) = cellThin
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeBottom) = cellThin
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeLeft) = cellThin
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeRight) = cellThin
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellInsideVertical) = cellThin
lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellInsideHorizontal) = cellThin
lista.PrintPreview
End If
End Sub

    Private Sub lista_DblClick()
 
        
        dato1.text = lista.Cell(lista.ActiveCell.row, 1).text
        dato2.text = lista.Cell(lista.ActiveCell.row, 2).text
        dato3.text = lista.Cell(lista.ActiveCell.row, 3).text
        dato4.text = lista.Cell(lista.ActiveCell.row, 4).text
        DATO5.text = lista.Cell(lista.ActiveCell.row, 5).text
        dato6.text = Mid(lista.Cell(lista.ActiveCell.row, 7).text, 1, 9)
        LBLDV.Caption = Mid(lista.Cell(lista.ActiveCell.row, 7).text, 10, 1)
        dato7.text = lista.Cell(lista.ActiveCell.row, 6).text
        lblBanco.Caption = leebanco(dato1.text)
        
        
        modifica = True
        dato3.SetFocus
    End Sub

    Private Sub lista_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
'        Select Case KeyCode
'            Case 46
'            If MsgBox("ESTA SEGURO QUE DESEA ELIMINAR CODIGO", vbYesNo, "ATENCION") = vbYes Then
'                If lista.ActiveCell.Row > 0 Then
'                    Call eliminarEspeciales(lista.Cell(lista.ActiveCell.Row, 1).text, lista.Cell(lista.ActiveCell.Row, 2).text)
'                    lista.RemoveItem (lista.ActiveCell.Row)
'                End If
'            End If
'        End Select
        
    End Sub
      Public Sub cargaLista()
        Call leerEspeciales
    End Sub

Private Function existecodigobancario(banco, codigo) As Boolean
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = conta

csql.sql = "select nombre,codigocontable,tipo,glosa,rut from cartolasbancarias_codigoscontables "
csql.sql = csql.sql & "where banco='" & banco & "' and  codigo='" & codigo & "' "
csql.Execute
existecodigobancario = False
If csql.RowsAffected > 0 Then
    existecodigobancario = True
    Set resultados = csql.OpenResultset
    dato3.text = resultados(0)
    dato4.text = resultados(1)
    DATO5.text = resultados(2)
    dato6.text = Mid(resultados(4), 1, 9)
    LBLDV.Caption = Mid(resultados(4), 10, 1)
    dato7.text = resultados(3)
    
End If
End Function
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = ""
    lblBanco.Caption = ""
    LBLDV.Caption = ""
    dato1.SetFocus
    
End Sub


Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    lista.FixedRowColStyle = Fixed3D
    lista.CellBorderColorFixed = vbButtonShadow
    lista.ShowResizeTips = False
    lista.ReportTitles.Clear
    lista.PageSetup.CenterHorizontally = True
    lista.PageSetup.Orientation = cellLandscape
    
      
    lista.PageSetup.PrintTitleRows = 0
    
    'Logo
'    lista.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    lista.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    lista.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    lista.PageSetup.HeaderAlignment = CellLeft
    lista.PageSetup.HeaderFont.Name = "Verdana"
    lista.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "EMITIDO  :  " & Format(fechasistema, "dd-MM-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle
    
      
    
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = tipoListado
'    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
'    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold = True
'    objReportTitle.Align = cellCenter
'    objReportTitle.PrintOnAllPages = True
'    lista.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    lista.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    lista.PageSetup.FooterAlignment = cellRight
    lista.PageSetup.FooterFont.Name = "Verdana"
    lista.PageSetup.FooterFont.Size = 7
    
End Sub






Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub
