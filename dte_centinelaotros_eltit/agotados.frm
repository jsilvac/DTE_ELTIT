VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Agotados 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7335
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmDatos 
      Height          =   7230
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   12753
      BackColor       =   12648384
      Caption         =   "Productos Agotados"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "RETORNO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   855
         Width           =   1815
      End
      Begin VB.CommandButton Imprime 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   450
         Width           =   1815
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
         MaxLength       =   13
         TabIndex        =   0
         Top             =   6840
         Width           =   1635
      End
      Begin XPFrame.FrameXp frmLista 
         Height          =   5010
         Left            =   120
         TabIndex        =   3
         Top             =   1395
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   8837
         BackColor       =   12648447
         Caption         =   "Lista de Productos"
         CaptionEstilo3D =   1
         BackColor       =   12648447
         ColorBarraArriba=   12648447
         ColorBarraAbajo =   32896
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
            Height          =   4620
            Left            =   45
            TabIndex        =   2
            Top             =   315
            Width           =   11715
            _ExtentX        =   20664
            _ExtentY        =   8149
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
         Left            =   11700
         TabIndex        =   4
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   32896
         ColorBarraAbajo =   12648447
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
      Begin XPFrame.FrameXp frmimprime 
         Height          =   885
         Left            =   3465
         TabIndex        =   16
         Top             =   450
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1561
         BackColor       =   12648384
         Caption         =   "Fecha Consultar"
         CaptionEstilo3D =   1
         BackColor       =   12648384
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BordeEstilo     =   5
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
            Left            =   750
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "fecha"
            Top             =   525
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
            Left            =   390
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "fecha"
            Top             =   525
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
            Left            =   30
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "fecha"
            Top             =   525
            Width           =   375
         End
         Begin VB.TextBox HASTA3 
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
            Left            =   2235
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "fecha"
            Top             =   525
            Width           =   615
         End
         Begin VB.TextBox HASTA2 
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
            Left            =   1875
            MaxLength       =   2
            TabIndex        =   18
            Tag             =   "fecha"
            Top             =   525
            Width           =   375
         End
         Begin VB.TextBox HASTA1 
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
            Left            =   1515
            MaxLength       =   2
            TabIndex        =   17
            Tag             =   "fecha"
            Top             =   525
            Width           =   375
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HASTA"
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
            Left            =   1515
            TabIndex        =   24
            Top             =   285
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DESDE"
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
            Left            =   30
            TabIndex        =   23
            Top             =   285
            Width           =   1335
         End
      End
      Begin VB.TextBox dato3 
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
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   9
         Top             =   5760
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPR Elimina Individual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7965
         TabIndex        =   25
         Top             =   900
         Width           =   2985
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ESC Para Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9675
         TabIndex        =   14
         Top             =   45
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   90
         TabIndex        =   13
         Top             =   0
         Width           =   4605
      End
      Begin VB.Label fechaactual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9765
         TabIndex        =   11
         Top             =   495
         Width           =   2130
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7740
         TabIndex        =   10
         Top             =   495
         Width           =   1965
      End
      Begin VB.Label lblProducto 
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
         Left            =   1860
         TabIndex        =   8
         Top             =   6840
         Width           =   7920
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9840
         TabIndex        =   7
         Top             =   6000
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Top             =   6480
         Width           =   7920
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   6480
         Width           =   1650
      End
   End
End
Attribute VB_Name = "Agotados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(10, 10) As String
    Private modifica As Boolean
    Private fecha1 As String
    Private fecha2 As String
    
    

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
'        Call selecciona(dato1)
        Label1.Caption = "F2 Ayuda Producto"
    End Sub
 
    Private Sub dato3_GotFocus()
'        Call selecciona(dato3)
        Label1.Caption = "F2 Ayuda Vendedores"
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaProducto(dato1)
        Else
'            Call flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudavendedor(dato3)
        Else
'           Call flechas(KeyCode, dato1)
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
                 dato1.text = ceros(dato1)
                  lblProducto.Caption = leerNombreProducto(dato1.text)
             If lblProducto.Caption <> "" Then
                Call leerEspeciales
                     If modifica = False Then
                        Call grabarEspeciales
                        dato1.text = ""
                        lblProducto.Caption = ""
'                        dato3.text = ""
                     Else
                        Call modificaEspeciales
                        dato1.text = ""
                        lblProducto.Caption = ""
'                        dato3.text = ""
                     End If
                 dato1.SetFocus
                 Call leerEspeciales
              Else
                 dato1.text = ""
                 dato1.SetFocus
              End If
       
      

            Else
'                Call selecciona(dato1)
            End If
        
    End Sub
    
'    Private Sub dato2_KeyPress(KeyAscii As Integer)
'        Dim Precio As String
'        KeyAscii = esNumero(KeyAscii, "N")
'        If KeyAscii = 13 And dato2.text <> "" Then
'            Precio = revisaCodigo
'            If Precio <> "" Then
'                dato3.text = Precio
'                modifica = True
'            Else
'                modifica = False
'            End If
'            dato3.SetFocus
'        End If
'    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato3.text <> "" Then
        Call ceros(dato3)
          If existevendedor(dato3.text) = True Then
               If modifica = False Then
                Call grabarEspeciales
                dato1.text = ""
                lblProducto.Caption = ""
'                dato3.text = ""
                
                
               Else
                Call modificaEspeciales
                dato1.text = ""
                lblProducto.Caption = ""
'                dato3.text = ""
              
               End If
              dato1.SetFocus
              Call leerEspeciales
              Else
'              dato3.text = ""
'              dato3.SetFocus
          End If
       
        
        
        End If
        
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
'        Call limpiaBarra(2)
        Label1.Caption = ""
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaLista(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "REF "
        formatogrilla(1, 3) = "DESCRIPCION"
        formatogrilla(1, 4) = "B00"
        formatogrilla(1, 5) = "B01"
        formatogrilla(1, 6) = "BOTRAS"
        formatogrilla(1, 7) = "CODIGO"
        formatogrilla(1, 8) = "VENDEDOR"
        formatogrilla(1, 9) = "FECHA"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "13"
        formatogrilla(2, 2) = "20"
        formatogrilla(2, 3) = "30"
        formatogrilla(2, 4) = "4"
        formatogrilla(2, 5) = "4"
        formatogrilla(2, 6) = "4"
        formatogrilla(2, 7) = "0"
        formatogrilla(2, 8) = "9"
        formatogrilla(2, 9) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "S"
        formatogrilla(3, 9) = "D"
        
        

        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = "###0"
        formatogrilla(4, 5) = "###0"
        formatogrilla(4, 6) = "###0"
        formatogrilla(4, 7) = "00"
        formatogrilla(4, 8) = ""
        formatogrilla(4, 9) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        formatogrilla(7, 9) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "15"
        formatogrilla(8, 3) = "30"
        formatogrilla(8, 4) = "4"
        formatogrilla(8, 5) = "4"
        formatogrilla(8, 6) = "5"
        formatogrilla(8, 7) = "0"
        formatogrilla(8, 8) = "20"
        formatogrilla(8, 9) = "10"
            
'        FORMATOGRILLA(1, 1) = "CODIGO"
'        FORMATOGRILLA(1, 2) = "REF "
'        FORMATOGRILLA(1, 3) = "DESCRIPCION"
'        FORMATOGRILLA(1, 4) = "B00"
'        FORMATOGRILLA(1, 5) = "B01"
'        FORMATOGRILLA(1, 6) = "BOTRAS"
'        FORMATOGRILLA(1, 7) = "CODIGO"
'        FORMATOGRILLA(1, 8) = "VENDEDOR"
'        FORMATOGRILLA(1, 9) = "FECHA"
           
            
            
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
        lista.BackColorFixed = RGB(90, 214, 158)
        lista.BackColorFixedSel = RGB(110, 230, 180)
        lista.BackColorBkg = RGB(90, 214, 158)
        lista.BackColorScrollBar = RGB(231, 247, 235)
        lista.BackColor1 = RGB(231, 247, 235)
        lista.BackColor2 = RGB(239, 255, 243)
        lista.GridColor = RGB(148, 231, 190)
        
        lista.Column(0).Width = 0
        For i = 1 To col - 1
            lista.Cell(0, i).text = formatogrilla(1, i)
            lista.Column(i).Width = Val(formatogrilla(8, i)) * (lista.Cell(0, i).Font.Size + 1.25)
            lista.Column(i).MaxLength = Val(formatogrilla(2, i))
            lista.Column(i).FormatString = formatogrilla(4, i)
            lista.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
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
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Double
    
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "SELECT codigo,descripcion,vendedor,fecha  "
    csql.sql = csql.sql + "FROM l_agotados_" & rubro & "  where impreso='0' "
    csql.sql = csql.sql + "ORDER BY fecha asc "
    csql.Execute
    lista.Rows = csql.RowsAffected + 1
    
    linea = 0
    If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
        lista.AutoRedraw = False
        While Not resultados.EOF
           linea = linea + 1
           lista.Cell(linea, 1).text = resultados(0)
           lista.Cell(linea, 2).text = leerref(resultados(0))
           lista.Cell(linea, 3).text = resultados(1)
           lista.Cell(linea, 4).text = leerstock(resultados(0), "00")
           lista.Cell(linea, 5).text = leerstock(resultados(0), "01")
           lista.Cell(linea, 6).text = leerstock(resultados(0), "02") + leerstock(resultados(0), "03") + leerstock(resultados(0), "04") + leerstock(resultados(0), "05") + leerstock(resultados(0), "06") + leerstock(resultados(0), "07")
           lista.Cell(linea, 7).text = resultados(2)
           lista.Cell(linea, 8).text = vendedor(resultados(2) & rut(resultados(2)))
           lista.Cell(linea, 9).text = Format(resultados(3), "dd-mm-yyyy")
            
            resultados.MoveNext
        Wend
        resultados.Close
        Set resultados = Nothing
        lista.AutoRedraw = True
        lista.Refresh
    Else
'        FORMATOGRILLA(1, 1) = "CODIGO"
'        FORMATOGRILLA(1, 2) = "REF "
'        FORMATOGRILLA(1, 3) = "DESCRIPCION"
'        FORMATOGRILLA(1, 4) = "B00"
'        FORMATOGRILLA(1, 5) = "B01"
'        FORMATOGRILLA(1, 6) = "BOTRAS"
'        FORMATOGRILLA(1, 7) = "CODIGO"
'        FORMATOGRILLA(1, 8) = "VENDEDOR"
'        FORMATOGRILLA(1, 9) = "FECHA"
    End If
   End Sub
   
   Private Sub leerinformeporfecha(desde, hasta)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Double
    
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "SELECT codigo,descripcion,vendedor,fecha  "
    csql.sql = csql.sql + "FROM l_agotados_" & rubro & " "
    csql.sql = csql.sql + "WHERE fecha between '" & desde & "' and '" & hasta & "' and impreso='0' "
    csql.sql = csql.sql + "ORDER BY fecha asc "
    csql.Execute
    lista.Rows = csql.RowsAffected + 1
    
    linea = 0
    If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
        lista.AutoRedraw = False
        While Not resultados.EOF
           linea = linea + 1
           lista.Cell(linea, 1).text = resultados(0)
           lista.Cell(linea, 2).text = leerref(resultados(0))
           lista.Cell(linea, 3).text = resultados(1)
           lista.Cell(linea, 4).text = leerstock(resultados(0), "00")
           lista.Cell(linea, 5).text = leerstock(resultados(0), "01")
           lista.Cell(linea, 6).text = leerstock(resultados(0), "02") + leerstock(resultados(0), "03") + leerstock(resultados(0), "04") + leerstock(resultados(0), "05") + leerstock(resultados(0), "06") + leerstock(resultados(0), "07")
           lista.Cell(linea, 8).text = vendedor(resultados(2) & rut(resultados(2)))
           lista.Cell(linea, 9).text = Format(resultados(3), "dd-mm-yyyy")
            
            resultados.MoveNext
        Wend
        resultados.Close
        Set resultados = Nothing
        lista.AutoRedraw = True
        lista.Refresh
    Else
    
    End If
   
    End Sub
'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub grabarEspeciales()
        
        Dim CAMPOS(10, 3) As String
        Dim op As Integer
        CAMPOS(0, 0) = "codigo"
        CAMPOS(1, 0) = "descripcion"
        CAMPOS(2, 0) = "vendedor"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = ""
        
        CAMPOS(0, 1) = dato1.text
        CAMPOS(1, 1) = lblProducto.Caption
        CAMPOS(2, 1) = dato3.text
        CAMPOS(3, 1) = Format(fechasistema, "yyyy-mm-dd")
                
        CAMPOS(0, 2) = "l_agotados_" & rubro
        
        condicion = ""
        op = 2
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub modificaEspeciales()
        
        Dim CAMPOS(10, 3) As String
        Dim op As Integer
        CAMPOS(0, 0) = "codigo"
        CAMPOS(1, 0) = "descripcion"
        CAMPOS(2, 0) = "vendedor"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = ""
        
        CAMPOS(0, 1) = dato1.text
        CAMPOS(1, 1) = lblProducto.Caption
        CAMPOS(2, 1) = dato3.text
        CAMPOS(3, 1) = Format(fechasistema, "yyyy-mm-dd")
        CAMPOS(4, 1) = ""
        
        CAMPOS(0, 2) = "l_agotados_" & rubro
        
        condicion = "codigo = '" & dato1.text & "' AND fecha= '" + Format(fechaactual.Caption, "yyyy-mm-dd") + "' "
        op = 3
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
        modifica = False
    End Sub
'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================
    Private Sub eliminarEspeciales(ByVal CODIGO As String, ByVal fecha1 As String)
        
        Dim CAMPOS(1, 3) As String
        Dim op As Integer
        condicion = "codigo = '" & CODIGO & "' AND fecha = '" & Format(fecha1, "yyyy-mm-dd") & "'"
        op = 4
        CAMPOS(0, 2) = "l_agotados_" & rubro
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================

Private Sub dato3_LostFocus()
Label1.Caption = ""
End Sub

   '========================================================
    Private Sub DESDE1_GotFocus()
    Call cargatexto(DESDE1)
End Sub

Private Sub DESDE1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, DESDE1)
End Sub

Private Sub DESDE1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    DESDE1.text = ceros(DESDE1)
    If DESDE1.text = "00" Then
                DESDE1.text = Format(Now, "dd")
                DESDE2.SetFocus
                Else
                DESDE2.SetFocus
            End If
    Call esfecha(DESDE1, DESDE2, DESDE3, "dd")
    End If
    
End Sub

Private Sub DESDE2_GotFocus()
    Call cargatexto(DESDE2)
End Sub

Private Sub DESDE2_KeyDown(KeyCode As Integer, Shift As Integer)
   Call Flechas(KeyCode, DESDE1)
End Sub

Private Sub DESDE2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    DESDE2.text = ceros(DESDE2)
    If DESDE2.text = "00" Then
                DESDE2.text = Format(Now, "mm")
                DESDE3.SetFocus
                Else
                DESDE3.SetFocus
            End If
    
    Call esfecha(DESDE1, DESDE2, DESDE3, "mm")
End If
End Sub

Private Sub DESDE3_GotFocus()
    Call cargatexto(DESDE3)
End Sub

Private Sub DESDE3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, DESDE2)
End Sub

Private Sub DESDE3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    DESDE3.text = ceros(DESDE3)
    If DESDE3.text = "0000" Then
                DESDE3.text = Format(Now, "yyyy")
                HASTA1.SetFocus
                Else
                HASTA1.SetFocus
            End If
   
    Call esfecha(DESDE1, DESDE2, DESDE3, "yyyy")
End If
End Sub

Private Sub HASTA1_GotFocus()
    Call cargatexto(HASTA1)
End Sub

Private Sub HASTA1_KeyDown(KeyCode As Integer, Shift As Integer)
   Call Flechas(KeyCode, DESDE3)
End Sub

Private Sub HASTA1_KeyPress(KeyAscii As Integer)
     KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    HASTA1.text = ceros(HASTA1)
    If HASTA1.text = "00" Then
                HASTA1.text = Format(Now, "dd")
                HASTA2.SetFocus
                Else
                
                HASTA2.SetFocus
            End If
   
    Call esfecha(HASTA1, HASTA2, HASTA3, "dd")
End If

End Sub

Private Sub HASTA2_GotFocus()
    Call cargatexto(HASTA2)
End Sub

Private Sub HASTA2_KeyDown(KeyCode As Integer, Shift As Integer)
   Call Flechas(KeyCode, HASTA1)
End Sub

Private Sub hasta2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    HASTA2.text = ceros(HASTA2)
    If HASTA2.text = "00" Then
                HASTA2.text = Format(Now, "mm")
                HASTA3.SetFocus
                Else
                HASTA3.SetFocus
            End If
    
Call esfecha(HASTA1, HASTA2, HASTA3, "mm")
End If
End Sub

Private Sub HASTA3_GotFocus()
    Call cargatexto(HASTA3)
End Sub

Private Sub HASTA3_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Flechas(KeyCode, HASTA2)
End Sub

Private Sub hasta3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    HASTA3.text = ceros(HASTA3)
     If HASTA3.text = "0000" Then
                HASTA3.text = Format(Now, "yyyy")
                Else
              Call esfecha(HASTA1, HASTA2, HASTA3, "yyyy")
              fecha1 = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
              fecha2 = HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text
                Call leerinformeporfecha(fecha1, fecha2)
            
            End If
   
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
        Call CargaGrillaLista(1, 10)
        cargaLista
        modifica = False
        fechaactual.Caption = Format(fechasistema, "dd-mm-yyyy")
        dato3.text = usuarioSistema
    End Sub
    
    Private Sub frmCerrar_BarClick()
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

Private Sub Command1_Click()
DESDE1.text = ""
DESDE2.text = ""
DESDE3.text = ""
HASTA1.text = ""
HASTA2.text = ""
HASTA3.text = ""
dato1.text = ""
lblProducto.Caption = ""
'dato3.text = ""

lista.Rows = 1
dato1.SetFocus
fechaactual.Caption = Format(fechasistema, "dd-mm-yyyy")
modifica = False
End Sub

Private Sub Imprime_Click()
 fecha1 = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
 fecha2 = HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text
 Call leerinformeporfecha(fecha1, fecha2)
If lista.Rows > 1 Then
      MsgBox "Despues de Imprimir Se Limpiara La Pantalla", vbOKOnly, "ATENCION"
       Titulos
       lista.PrintPreview
       Call eliminarlistado(fecha1, fecha2)
       Call Command1_Click
     
    Else
     DESDE1.SetFocus
End If
End Sub


Sub Titulos()

  
    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
 
    
    lista.FixedRowColStyle = Fixed3D
    lista.CellBorderColorFixed = vbButtonShadow
    lista.ShowResizeTips = False
    lista.PageSetup.Orientation = cellLandscape
    
    
    
    lista.PageSetup.PrintFixedRow = True
    lista.ReportTitles.Clear
    lista.PageSetup.CenterHorizontally = True
    lista.PageSetup.PrintTitleRows = 0
    lista.PageSetup.BlackAndWhite = True
    
    
    
    
    'ENCABEZADO DE PAGINA
    lista.PageSetup.Header = nombreempresa & vbCrLf & leerDireccionEmpresa(empresaActiva)
    lista.PageSetup.HeaderAlignment = cellLeft
    lista.PageSetup.HeaderFont.Name = "Verdana"
    lista.PageSetup.HeaderFont.Size = 8
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO PRODUCTOS AGOTADOS "
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = "PERIODO  :  DEL " & fechasistema
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle
        
        
    Set objReportTitle = New FlexCell.ReportTitle
  
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Underline = True
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    lista.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D " & vbCrLf & "Usuario:" + usuarioSistema
    lista.PageSetup.FooterAlignment = cellRight
    lista.PageSetup.FooterFont.Name = "Verdana"
    lista.PageSetup.FooterFont.Size = 7
    lista.PageSetup.LeftMargin = 0.5
    lista.PageSetup.RightMargin = 0.5
    
    lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeLeft) = cellThick
    lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeTop) = cellThick
    lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeBottom) = cellThick
    lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellEdgeRight) = cellThick
    lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    lista.Range(0, 1, 0, lista.Cols - 1).Borders(cellInsideVertical) = cellThick
    
    
    
    
End Sub


    Private Sub Lista_DblClick()
        dato1.text = lista.Cell(lista.ActiveCell.row, 1).text
        lblProducto.Caption = lista.Cell(lista.ActiveCell.row, 3).text
        dato3.text = lista.Cell(lista.ActiveCell.row, 7).text
        fechaactual.Caption = lista.Cell(lista.ActiveCell.row, 9).text
        modifica = True
        dato3.SetFocus
    End Sub



    Private Sub lista_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        Select Case KeyCode
            Case 46
                If lista.ActiveCell.row > 0 Then
                    Call eliminarEspeciales(lista.Cell(lista.ActiveCell.row, 1).text, lista.Cell(lista.ActiveCell.row, 9).text)
                    lista.RemoveItem (lista.ActiveCell.row)
                End If
        End Select
    End Sub

'    Private Function revisaCodigo() As String
'        Dim i As Long
'        revisaCodigo = ""
'        For i = 1 To lista.Rows - 1
'            If lista.Cell(i, 1).text = dato1.text And lista.Cell(i, 3).text = dato2.text Then
'                revisaCodigo = lista.Cell(i, 4).text
'                Exit For
'            End If
'        Next i
'    End Function

    Private Sub ayudaProducto(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = basedatos & rubro
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "r_maestroproductos_fijo_" & rubro & " AS mpf"
        mensajeAyuda = "Ayuda de Productos"
        camposAyuda = Array("mpf.codigobarra", "mpf.descripcion")
        cabezasAyuda = Array("codigo", "descripcion")
        largoAyuda = Array("13n", "50s")
        condicionAyuda = "mpf.descontinuado='0'"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub
     Private Sub ayudavendedor(ByRef txt As TextBox)
        servidorAyuda = servidor
        basedatosAyuda = baseVentas
        usuarioAyuda = usuario
        passAyuda = password
        tablaAyuda = "sv_maestrovendedores"
        mensajeAyuda = "Ayuda de Vendedores"
        camposAyuda = Array("codigo", "nombre")
        cabezasAyuda = Array("Codigo", "Nombre")
        largoAyuda = Array("13n", "50s")
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
        
    End Sub


    Public Sub cargaLista()
        Call leerEspeciales
    End Sub

Function existevendedor(CODIGO) As Boolean
   
    Dim op As Integer
    Dim cad As String
    Dim p As String
    Dim CAMPOS(2, 2) As String
    
    CAMPOS(0, 0) = "nombre"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_maestrocajeras"
    condicion = "rut= '" & CODIGO & "'"
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
        existevendedor = True
       
       
        
    Else
        existevendedor = False
    End If
End Function
Function vendedor(CODIGO) As String
   
    Dim op As Integer
    Dim cad As String
    Dim p As String
    Dim CAMPOS(3, 3) As String
    
    CAMPOS(0, 0) = "nombre"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_maestrocajeras"
    
    condicion = "rut= '" & CODIGO & "'"
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    
    If sqlventas.Status = 0 Then
    
        vendedor = sqlventas.response(0, 3)
        
    Else
    
        vendedor = ""
        
    End If
End Function

 Private Function leerref(ByVal CODIGO As String) As String
        
        Dim op As Integer
        Dim CAMPOS(3, 3) As String
        
        'Set sql = New CSQLUtil
        CAMPOS(0, 0) = "proveedor"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "r_maestroproductos_fijo_" & rubro
        
        condicion = "codigobarra = '" & CODIGO & "'"
        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
            leerref = sqlventas.response(0, 3)
        Else
            leerref = ""
        End If
    End Function

Private Function leerstock(CODIGO, bodega)
    Dim a As Integer
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim saldo As Double

        Set csql.ActiveConnection = gestionRubro
        csql.sql = "SELECT stockactual "
        csql.sql = csql.sql + "FROM r_maestroproductos_stock_" & rubro & " "
        csql.sql = csql.sql + "WHERE año='" + Format(fechasistema, "yyyy") + "' AND codigo='" + CODIGO + "' AND bodega='" + bodega + "' "
        csql.Execute
       
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
           leerstock = resultados(0)
            resultados.Close
            Set resultados = Nothing
            Else
            leerstock = 0
        End If
       
End Function
'
'Private Sub eliminarlistado(ByVal fecha As String, ByVal fecha_2 As String)
'
'        Dim campos(1, 3) As String
'        Dim op As Integer
'        condicion = "fecha between '" & Format(fecha, "yyyy-mm-dd") & "' and '" & Format(fecha_2, "yyyy-mm-dd") & "'"
'        op = 4
'        campos(0, 2) = "l_agotados_" & rubro
'        SQLUTIL.datos = campos
'        Set SQLUTIL.conexion = GESTIONrubro
'        Call SQLUTIL.SQLUTIL(op, condicion)
'    End Sub
    
     Private Sub eliminarlistado(ByVal fecha As String, ByVal fecha_2 As String)
        
        Dim CAMPOS(3, 3) As String
        Dim op As Integer
        
        CAMPOS(0, 0) = "impreso"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 1) = "1"
        
        
        CAMPOS(0, 2) = "l_agotados_" & rubro
        
        
        
        condicion = "fecha between '" & Format(fecha, "yyyy-mm-dd") & "' and '" & Format(fecha_2, "yyyy-mm-dd") & "'"
        op = 3
        CAMPOS(0, 2) = "l_agotados_" & rubro
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = gestionRubro
        Call sqlventas.sqlventas(op, condicion)
    End Sub






