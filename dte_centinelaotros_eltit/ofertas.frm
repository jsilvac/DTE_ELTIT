VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ofertas 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7575
   ClientLeft      =   645
   ClientTop       =   2775
   ClientWidth     =   14550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7575
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp frmDatos 
      Height          =   7500
      Left            =   90
      TabIndex        =   1
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13229
      BackColor       =   12648384
      Caption         =   "INGRESO DE OFERTAS"
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
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   1335
         Left            =   240
         TabIndex        =   26
         Top             =   6120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2355
         BackColor       =   8454016
         Caption         =   "OPCIONES GRILLA"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         ForeColor       =   255
         BordeColor      =   16777215
         ColorBarraArriba=   8454016
         ColorBarraAbajo =   49152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorTextShadow =   16777215
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DOBLE CLICK MODIFICA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ESC SALIR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "(SUPR) ELIMINA OFERTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6300
         Width           =   5865
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   9630
         MaxLength       =   4
         TabIndex        =   24
         Top             =   5760
         Width           =   525
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   9315
         MaxLength       =   2
         TabIndex        =   23
         Top             =   5760
         Width           =   255
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   9000
         MaxLength       =   2
         TabIndex        =   22
         Top             =   5760
         Width           =   255
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   8370
         MaxLength       =   4
         TabIndex        =   21
         Top             =   5760
         Width           =   525
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   8055
         MaxLength       =   2
         TabIndex        =   20
         Top             =   5760
         Width           =   255
      End
      Begin VB.TextBox dato11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   12960
         MaxLength       =   9
         TabIndex        =   18
         Top             =   5760
         Width           =   1290
      End
      Begin VB.TextBox dato10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   11610
         MaxLength       =   9
         TabIndex        =   17
         Top             =   5760
         Width           =   1290
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10215
         MaxLength       =   9
         TabIndex        =   16
         Top             =   5760
         Width           =   1290
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6390
         MaxLength       =   9
         TabIndex        =   10
         Top             =   5760
         Width           =   1290
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   7740
         MaxLength       =   2
         TabIndex        =   2
         Top             =   5760
         Width           =   255
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         MaxLength       =   13
         TabIndex        =   0
         Top             =   5760
         Width           =   1545
      End
      Begin XPFrame.FrameXp frmLista 
         Height          =   4875
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   14220
         _ExtentX        =   25083
         _ExtentY        =   8599
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
            Height          =   4395
            Left            =   60
            TabIndex        =   3
            Top             =   360
            Width           =   14100
            _ExtentX        =   24871
            _ExtentY        =   7752
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
         Left            =   14040
         TabIndex        =   5
         Top             =   45
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
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   13005
         TabIndex        =   19
         Top             =   5400
         Width           =   1290
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max.vta"
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
         Left            =   8190
         TabIndex        =   15
         Top             =   5040
         Width           =   1470
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max.vta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   10215
         TabIndex        =   14
         Top             =   5400
         Width           =   1290
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7740
         TabIndex        =   13
         Top             =   5400
         Width           =   1155
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.termino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9000
         TabIndex        =   12
         Top             =   5400
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max.stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11610
         TabIndex        =   11
         Top             =   5400
         Width           =   1290
      End
      Begin VB.Label lblProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1860
         TabIndex        =   9
         Top             =   5760
         Width           =   4500
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6390
         TabIndex        =   8
         Top             =   5400
         Width           =   1290
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Top             =   5400
         Width           =   4500
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   5400
         Width           =   1515
      End
   End
End
Attribute VB_Name = "ofertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatoGrilla(10, 10) As String
    Private modifica As Boolean

Private Sub Command1_Click()
Titulos



lista.PrintPreview

End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call selecciona(dato1)
    End Sub
    
Private Sub dato10_GotFocus()

            Call selecciona(dato10)
        
End Sub

Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
 Call flechas(dato9, dato10, KeyCode)
End Sub

Private Sub dato10_KeyPress(KeyAscii As Integer)
        Dim valida As Double
        If Len(dato10.text) > 1 Then
        valida = Mid(dato10.text, 1, 2)
        Else
        valida = 1
        End If
        KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 And dato10.text <> "" And dato10.text <> "0" And valida <> "0" Then
            dato11.SetFocus
        End If
               
End Sub

Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
         Call flechas(dato10, dato11, KeyCode)
End Sub

Private Sub dato11_gotFocus()

            Call selecciona(dato11)
    
End Sub

    Private Sub dato2_GotFocus()
        Call selecciona(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
  
            Call selecciona(dato3)
    
       
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
            Call flechas(dato1, dato2, KeyCode)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato1, dato3, KeyCode)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato2, dato4, KeyCode)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 And dato1.text <> "" Then
            Call ceros(dato1)
            lblProducto.caption = leerNombreProducto(dato1.text)
            dato11.text = leerPrecioProducto(dato1.text)
            
            If lblProducto.caption <> "" Then
                Call leerEspeciales
                dato2.text = "2"
                dato2.SetFocus
            Else
                Call selecciona(dato1)
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        Dim Precio As String
        Dim valida As Double
        If Len(dato2.text) > 1 Then
        valida = Mid(dato2.text, 1, 2)
        Else
        valida = 1
        End If
        
        KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 And dato2.text <> "" And dato2.text <> "0" And valida <> "0" Then
            Precio = revisaCodigo
            If Precio <> "" Then
                dato11.text = Precio
                modifica = True
            Else
                modifica = False
            End If
            dato3.SetFocus
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
    
    
        KeyAscii = esNumero(KeyAscii, "N")
        
        If KeyAscii = 13 And dato11.text <> "" And dato11.text <> "0" Then
            If modifica = False Then
                Call grabarEspeciales
            Else
                Call modificaEspeciales
            End If
            
            Call leerEspeciales
            limpiar
            
            dato1.Enabled = True
            dato2.Enabled = True
            dato1.SetFocus
                        
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        Call limpiaBarra(2)
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
    Private Sub CargaGrillaLista(ByVal Row As Integer, ByVal Col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "DESCRIPCION"
        formatoGrilla(1, 3) = "CANTIDAD"
        formatoGrilla(1, 4) = "F.INI."
        formatoGrilla(1, 5) = "F.TER."
        formatoGrilla(1, 6) = "MAX.VTA"
        formatoGrilla(1, 7) = "MAX.STOCK"
        formatoGrilla(1, 8) = "$ PRECIO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "10"
        formatoGrilla(2, 2) = "30"
        formatoGrilla(2, 3) = "10"
        formatoGrilla(2, 4) = "10"
        formatoGrilla(2, 5) = "10"
        formatoGrilla(2, 6) = "10"
        formatoGrilla(2, 7) = "10"
        formatoGrilla(2, 8) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "D"
        formatoGrilla(3, 5) = "D"
        formatoGrilla(3, 6) = "N"
        formatoGrilla(3, 7) = "N"
        formatoGrilla(3, 8) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 3) = "###,###"
        formatoGrilla(4, 4) = "dd/mm/yyyy"
        formatoGrilla(4, 5) = "dd/mm/yyyy"
        formatoGrilla(4, 6) = "##,###,##0"
        formatoGrilla(4, 7) = "##,###,##0"
        formatoGrilla(4, 8) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "TRUE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "TRUE"
        formatoGrilla(5, 4) = "TRUE"
        formatoGrilla(5, 5) = "TRUE"
        formatoGrilla(5, 6) = "TRUE"
        formatoGrilla(5, 7) = "TRUE"
        formatoGrilla(5, 8) = "TRUE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        formatoGrilla(6, 8) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "24"
        formatoGrilla(8, 3) = "7"
        formatoGrilla(8, 4) = "10"
        formatoGrilla(8, 5) = "10"
        formatoGrilla(8, 6) = "8"
        formatoGrilla(8, 7) = "8"
        formatoGrilla(8, 8) = "8"
            
        lista.Cols = Col
        lista.Rows = Row
        lista.AllowUserResizing = False
        lista.DisplayFocusRect = False
        lista.ExtendLastCol = False
        
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
        For i = 1 To Col - 1
            lista.Cell(0, i).text = formatoGrilla(1, i)
            lista.Column(i).Width = Val(formatoGrilla(8, i)) * (lista.Cell(0, i).Font.Size + 1.25)
            lista.Column(i).MaxLength = Val(formatoGrilla(2, i))
            lista.Column(i).FormatString = formatoGrilla(4, i)
            lista.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                lista.Column(i).Alignment = cellRightCenter
            Else
                lista.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        lista.Range(0, 1, 0, lista.Cols - 1).Alignment = cellCenterCenter
        lista.Enabled = True
    lista.Column(4).Alignment = cellCenterCenter
    lista.Column(5).Alignment = cellCenterCenter
    
    
    
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================
    Private Sub leerEspeciales()
        Dim tabla As String
        Dim linea As Double
        
        tabla = "SELECT mpo.codigo,mpf.descripcion,mpo.cantidad,mpo.fechainicio,mpo.fechatermino,mpo.maximoxcliente,mpo.maximostockalaventa,mpo.preciooferta "
        tabla = tabla & "FROM r_maestroproductos_ofertas_" & rubro & " as mpo, r_maestroproductos_fijo_" & rubro & " as mpf "
        tabla = tabla & "where mpo.local='" + empresaactiva + "' and mpf.codigobarra=mpo.codigo "
        
        Call ConectarControlData(data, servidor, basedatos & rubro, usuario, password, tabla)
        lista.Rows = data.Recordset.RecordCount + 1
        
        lista.AutoRedraw = False
        linea = 0
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
               linea = linea + 1
                
                lista.Cell(linea, 1).text = data.Recordset.Fields(0)
                lista.Cell(linea, 2).text = data.Recordset.Fields(1)
                lista.Cell(linea, 3).text = data.Recordset.Fields(2)
                lista.Cell(linea, 4).text = data.Recordset.Fields(3)
                lista.Cell(linea, 5).text = data.Recordset.Fields(4)
                lista.Cell(linea, 6).text = data.Recordset.Fields(5)
                lista.Cell(linea, 7).text = data.Recordset.Fields(6)
                lista.Cell(linea, 8).text = data.Recordset.Fields(7)
                
                
                
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
        Dim preciocosto As Double
        preciocosto = leerCostoProducto(dato1)
        campos(0, 0) = "local"
        campos(1, 0) = "codigo"
        campos(2, 0) = "codigoprecio"
        campos(3, 0) = "cantidad"
        campos(4, 0) = "fechainicio"
        campos(5, 0) = "fechatermino"
        campos(6, 0) = "maximoxcliente"
        campos(7, 0) = "maximostockalaventa"
        campos(8, 0) = "preciooferta"
        campos(9, 0) = "costooferta"
        campos(10, 0) = ""
        
        campos(0, 1) = empresaactiva
        campos(1, 1) = dato1.text
        campos(2, 1) = "01"
        campos(3, 1) = dato2.text
        campos(4, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
        campos(5, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
        campos(6, 1) = dato9.text
        campos(7, 1) = dato10.text
        campos(8, 1) = dato11.text
        campos(9, 1) = preciocosto
        
        
        campos(0, 2) = "r_maestroproductos_ofertas_" & rubro
        
        condicion = ""
        op = 2
        SQLUTIL.datos = campos
        Set SQLUTIL.conexion = GESTIONrubro
        Call SQLUTIL.SQLUTIL(op, condicion)
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
        Dim preciocosto As Double
        
        preciocosto = leerCostoProducto(dato1)
        campos(0, 0) = "local"
        campos(1, 0) = "codigo"
        campos(2, 0) = "codigoprecio"
        campos(3, 0) = "cantidad"
        campos(4, 0) = "fechainicio"
        campos(5, 0) = "fechatermino"
        campos(6, 0) = "maximoxcliente"
        campos(7, 0) = "maximostockalaventa"
        campos(8, 0) = "preciooferta"
        campos(9, 0) = "costooferta"
        campos(10, 0) = ""
        
        campos(0, 1) = empresaactiva
        campos(1, 1) = dato1.text
        campos(2, 1) = "01"
        campos(3, 1) = dato2.text
        campos(4, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
        campos(5, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
        campos(6, 1) = dato9.text
        campos(7, 1) = dato10.text
        campos(8, 1) = dato11.text
        campos(9, 1) = preciocosto
        
        
        campos(0, 2) = "r_maestroproductos_ofertas_" & rubro
        
        condicion = "codigo = '" & dato1.text & "' AND codigoprecio = '01' AND cantidad = '" & dato2.text & "'"
        op = 3
        SQLUTIL.datos = campos
        Set SQLUTIL.conexion = GESTIONrubro
        Call SQLUTIL.SQLUTIL(op, condicion)
        modifica = False
    End Sub
'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================
    Private Sub eliminarEspeciales(ByVal codigo As String, ByVal Cantidad As String)
        Dim condicion As String
        Dim campos(1, 3) As String
        Dim op As Integer
        condicion = "codigo = '" & codigo & "' AND tipoprecio = '01' AND cantidad = '" & Cantidad & "'"
        op = 4
        campos(0, 2) = "r_maestroproductos_ofertas_" & rubro
        SQLUTIL.datos = campos
        Set SQLUTIL.conexion = GESTIONrubro
        Call SQLUTIL.SQLUTIL(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================

Private Sub dato3_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 Then
            If dato3.text = "" Then
                dato3.text = Format(Now, "dd")
                dato4.text = Format(Now, "mm")
                dato5.text = Format(Now, "yyyy")
                dato6.SetFocus
            Else
                dato4.SetFocus
            End If
        End If
End Sub

Private Sub dato3_LostFocus()
            Call ceros(dato3)
            Call esfecha(dato3, dato4, dato5, "dd")
        
End Sub

Private Sub dato4_GotFocus()

            Call selecciona(dato4)
 
 
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
 Call flechas(dato3, dato5, KeyCode)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
  KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 Then
            If dato4.text = "" Then
                dato4.text = Format(Now, "mm")
                dato5.text = Format(Now, "yyyy")
                dato6.SetFocus
            Else
                dato5.SetFocus
            End If
        End If
End Sub

Private Sub dato4_LostFocus()
 
            Call ceros(dato4)
            Call esfecha(dato3, dato4, dato5, "mm")
 
End Sub

Private Sub dato5_GotFocus()
 
            Call selecciona(dato5)
 
End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
 Call flechas(dato4, dato6, KeyCode)
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 Then
            If dato5.text = "" Then
                dato5.text = Format(Now, "yyyy")
                dato6.SetFocus
                
            End If
            dato6.SetFocus
        End If
End Sub

Private Sub dato5_LostFocus()
            Call ceros(dato5)
            Call esfecha(dato3, dato4, dato5, "yyyy")
        
End Sub

Private Sub dato6_GotFocus()
            Call selecciona(dato6)
 End Sub

Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
 Call flechas(dato5, dato7, KeyCode)
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 Then
            If dato6.text = "" Then
                dato6.text = Format(Now, "dd")
                dato7.text = Format(Now, "mm")
                dato8.text = Format(Now, "yyyy")
                dato9.SetFocus
            Else
                dato7.SetFocus
            End If
        End If
End Sub

Private Sub dato6_LostFocus()
 If dato6.text <> "" Then
            Call ceros(dato6)
            Call esfecha(dato6, dato7, dato8, "dd")
        End If
End Sub

Private Sub dato7_GotFocus()
            Call selecciona(dato7)
 
End Sub

Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
 Call flechas(dato6, dato8, KeyCode)
End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 Then
            If dato7.text = "" Then
                dato7.text = Format(Now, "mm")
                dato8.text = Format(Now, "yyyy")
                dato9.SetFocus
            Else
                dato8.SetFocus
            End If
        End If
End Sub

Private Sub dato7_LostFocus()
 If dato7.text <> "" Then
            Call ceros(dato7)
            Call esfecha(dato6, dato7, dato8, "mm")
        End If
End Sub

Private Sub dato8_GotFocus()
            Call selecciona(dato8)
 
End Sub

Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
 Call flechas(dato7, dato9, KeyCode)
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii, "N")
        If KeyAscii = 13 Then
            If dato8.text = "" Then
                dato8.text = Format(Now, "yyyy")
                dato9.SetFocus
                
            End If
            dato9.SetFocus
        End If
End Sub

Private Sub dato8_LostFocus()
            Call ceros(dato8)
            Call esfecha(dato6, dato7, dato8, "yyyy")
 
End Sub

Private Sub dato9_GotFocus()
            Call selecciona(dato9)
 
End Sub

Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)

 Call flechas(dato8, dato10, KeyCode)

 
 
End Sub

Private Sub dato9_KeyPress(KeyAscii As Integer)
 Dim valida As Double
        If Len(dato9.text) > 1 Then
        valida = Mid(dato9.text, 1, 2)
        Else
        valida = 1
        End If
        If KeyAscii = 13 And dato1.text <> "" And dato9.text <> "0" And valida <> "0" Then
            dato10.SetFocus
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
        Call CargaGrillaLista(1, 9)
        modifica = False
    leerEspeciales
    
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

    Private Sub lista_DblClick()
        If lista.ActiveCell.Row >= 1 Then
        
        dato1.text = lista.Cell(lista.ActiveCell.Row, 1).text
        lblProducto.caption = lista.Cell(lista.ActiveCell.Row, 2).text
        dato2.text = lista.Cell(lista.ActiveCell.Row, 3).text
        dato3.text = Mid(lista.Cell(lista.ActiveCell.Row, 4).text, 1, 2)
        dato4.text = Mid(lista.Cell(lista.ActiveCell.Row, 4).text, 4, 2)
        dato5.text = Mid(lista.Cell(lista.ActiveCell.Row, 4).text, 7, 4)
        dato6.text = Mid(lista.Cell(lista.ActiveCell.Row, 5).text, 1, 2)
        dato7.text = Mid(lista.Cell(lista.ActiveCell.Row, 5).text, 4, 2)
        dato8.text = Mid(lista.Cell(lista.ActiveCell.Row, 5).text, 7, 4)
        
        dato9.text = lista.Cell(lista.ActiveCell.Row, 6).text
        dato10.text = lista.Cell(lista.ActiveCell.Row, 7).text
        dato11.text = lista.Cell(lista.ActiveCell.Row, 8).text
        
        modifica = True
        dato1.Enabled = False
        dato2.Enabled = False
        
        dato3.SetFocus
    End If
    
    End Sub

    Private Sub lista_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        Select Case KeyCode
            Case 46
                If lista.ActiveCell.Row > 0 Then
                    Call eliminarEspeciales(lista.Cell(lista.ActiveCell.Row, 1).text, lista.Cell(lista.ActiveCell.Row, 3).text)
                    lista.RemoveItem (lista.ActiveCell.Row)
                End If
        End Select
    End Sub

    Private Function revisaCodigo() As String
        Dim i As Long
        revisaCodigo = ""
        For i = 1 To lista.Rows - 1
            If lista.Cell(i, 1).text = dato1.text And lista.Cell(i, 3).text = dato2.text Then
                revisaCodigo = lista.Cell(i, 4).text
                Exit For
            End If
        Next i
    End Function

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
        condicionAyuda = "no"
        cantidadAyuda = 2
        Call Mayuda.cargaAyuda(txt)
    End Sub

    Public Sub cargaLista()
        Call leerEspeciales
    End Sub













Private Sub Text4_Change()

End Sub

Private Sub Text9_Change()

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
    lista.PageSetup.PrintTitleRows = 1
    lista.PageSetup.BlackAndWhite = True
    
    
    
    
    
    
    'Logo
'    LISTA.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    LISTA.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    lista.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    lista.PageSetup.HeaderAlignment = CellLeft
    lista.PageSetup.HeaderFont.Name = "Verdana"
    lista.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE OFERTAS VIGENTES"
    objReportTitle.Font.Name = "Verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "VIGENTES AL : " + fechasistema
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    lista.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    lista.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D " + "usuario:" + USUARIOSISTEMA
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


Sub limpiar()
dato1.text = ""
dato2.text = ""
dato3.text = ""
dato4.text = ""
dato5.text = ""
dato6.text = ""
dato7.text = ""
dato8.text = ""
dato9.text = ""
dato10.text = ""
dato11.text = ""

End Sub
