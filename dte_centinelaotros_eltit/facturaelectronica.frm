VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form electro04 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERA OTROS  ELECTRONICOS"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   14565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   480
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   120
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   -1
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
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   7260
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   12806
      BackColor       =   8421631
      Caption         =   "Informe"
      CaptionEstilo3D =   1
      BackColor       =   8421631
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
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
      Begin FlexCell.Grid impresion 
         Height          =   6615
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   11668
         AllowUserResizing=   0   'False
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2175
      Left            =   75
      TabIndex        =   6
      Top             =   90
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   3836
      BackColor       =   8421631
      Caption         =   "OTROS ELECTRINICOS"
      CaptionEstilo3D =   1
      BackColor       =   8421631
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
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
      Begin VB.CheckBox sihayerror 
         BackColor       =   &H008080FF&
         Caption         =   "Detener en error"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   14160
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H008080FF&
         Caption         =   "Generacion Automatica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.OptionButton opt5 
         BackColor       =   &H00FF8080&
         Caption         =   "NC Factura"
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
         Height          =   300
         Left            =   16320
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.OptionButton opt4 
         BackColor       =   &H00FF8080&
         Caption         =   "NC Boleta"
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
         Height          =   300
         Left            =   14160
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Informe"
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
         Left            =   12120
         TabIndex        =   15
         Top             =   480
         Width           =   1635
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Facturas"
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
         Height          =   300
         Left            =   13920
         TabIndex        =   14
         Top             =   2100
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Boletas"
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
         Height          =   300
         Left            =   15930
         TabIndex        =   13
         Top             =   2100
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton opt3 
         BackColor       =   &H00FF8080&
         Caption         =   "Zetas"
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
         Height          =   300
         Left            =   20280
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   12420
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   1020
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   11940
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   1020
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   12900
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   1020
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   420
         Width           =   795
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1740
         Left            =   3960
         TabIndex        =   22
         Top             =   360
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   3069
         BackColor       =   16744576
         Caption         =   "Locales"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ColorBarraArriba=   4194304
         ColorBarraAbajo =   4194304
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
         Begin FlexCell.Grid Grid1 
            Height          =   1335
            Left            =   0
            TabIndex        =   23
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   2355
            Cols            =   3
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.Label lbllocal 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   8400
         TabIndex        =   21
         Top             =   1680
         Width           =   5895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Para Modificar folio:  Presione enter para grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   15480
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   660
         TabIndex        =   8
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   10500
         TabIndex        =   7
         Top             =   1020
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   10680
      TabIndex        =   11
      Top             =   9720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "I   M   P   R   I   M   I   R"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "electro04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private TIPO As String
    Private detalle As Boolean
    Private fecha1 As String
    Private fecha2 As String

Private Sub Command1_Click()
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            Call CargaGrillaInforme(1, 16)
            If LOCAL_PROCESO = "" Then LOCAL_PROCESO = empresaActiva
            
            Call generaInformeLV(data, impresion, TIPO, detalle, LOCAL_PROCESO, fecha1, fecha1)
            
End Sub

Private Sub Command2_Click()
'Call modificafacturadepublicidad("FV", "2011-01-17", "0000000100", "98", "08", "0000001156")


End Sub

Private Sub Command3_Click()

End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
    End Sub

    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub

    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call VerificarCajas(Me, dato4)
        Call selecciona(dato4)
    End Sub

    Private Sub dato5_GotFocus()
        Call VerificarCajas(Me, dato5)
        Call selecciona(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call VerificarCajas(Me, dato6)
        Call selecciona(dato6)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub

    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato5)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            SendKeys "{Tab}"
            Call generaInformeLV(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
'    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato1.text) = dato1.MaxLength Then
'            Call dato1_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato2.text) = dato2.MaxLength Then
'            Call dato2_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato3.text) = dato3.MaxLength Then
'            Call dato3_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato4_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato4.text) = dato4.MaxLength Then
'            Call dato4_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato5_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato5.text) = dato5.MaxLength Then
'            Call dato5_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato6_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato6.text) = dato6.MaxLength Then
'            Call dato6_KeyPress(13)
'        End If
'    End Sub
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    
    Private Sub dato1_LostFocus()
    Call limpiaBarra(2)
    Call esfecha(dato1, dato2, dato3, "dd")
    End Sub
    Private Sub dato2_LostFocus()
    Call esfecha(dato1, dato2, dato3, "mm")
    End Sub
    Private Sub dato3_LostFocus()
    Call esfecha(dato1, dato2, dato3, "yyyy")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato4, dato5, dato6, "dd")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato4, dato5, dato6, "mm")
    End Sub
    Private Sub dato6_LostFocus()
    Call esfecha(dato4, dato5, dato6, "yyyy")
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 27
                Unload Me
            Case 38
                If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                    Unload Me
                End If
        End Select
    End Sub
    
    Private Sub Form_Load()
    Dim ss As String
    Dim K As Integer
    Dim pat1 As String
    Dim pat2 As String
    Dim pater(10) As String
    ProcesaNC = False
    
    Close 20
        Open App.Path + "\confiotros.txt" For Input As #20
    While EOF(20) = False
    Input #20, ss
    
    If Mid(ss, 1, 8) = "SERVIDOR" Then
        Servidor = Mid(ss, 10, Len(ss) - 9)
    End If
    If Mid(ss, 1, 8) = "SERVIDO2" Then
        Servidor2 = Mid(ss, 10, Len(ss) - 9)
    End If
    If Mid(ss, 1, 9) = "BASEDATOS" Then
        basedatos = Mid(ss, 11, Len(ss) - 10)
    End If
    If Mid(ss, 1, 10) = "BASEVENTAS" Then
        baseVentas = Mid(ss, 12, Len(ss) - 11)
    End If
    If Mid(ss, 1, 7) = "EMPRESA" Then
        empresaActiva = Mid(ss, 9, Len(ss) - 8)
    End If
    If Mid(ss, 1, 6) = "BODEGA" Then
        bodega = Mid(ss, 8, Len(ss) - 7)
    End If
    If Mid(ss, 1, 4) = "CAJA" Then
        idCaja = Mid(ss, 6, Len(ss) - 5)
    End If
    If Mid(ss, 1, 4) = "RUTA" Then
        rutaUpdate = Mid(ss, 6, Len(ss) - 5)
    End If
    If Mid(ss, 1, 13) = "IMPRESORAPAGO" Then
        IMPRESORAPAGO = Mid(ss, 15, Len(ss) - 5)
    End If
    If Mid(ss, 1, 12) = "BODEGARETIRO" Then
        BODEGARETIRO = Mid(ss, 14, Len(ss) - 5)
    End If
    If Mid(ss, 1, 16) = "IMPRESORACREDITO" Then
        impresoracredito = Mid(ss, 18, Len(ss) - 5)
    End If
    If Mid(ss, 1, 11) = "ASEGURADORA" Then
        ASEGURADORA = Mid(ss, 13, Len(ss) - 12)
    End If
    If Mid(ss, 1, 14) = "IMPRIMEDIRECTO" Then
    If Mid(ss, 16, Len(ss) - 14) = "S" Then
        imprimeDirecto = True
        Else
        imprimeDirecto = False
    End If
    End If
    If Mid(ss, 1, 11) = "IMPRIMETIPO" Then
        imprIMETIPO = Mid(ss, 13, Len(ss) - 11)
    End If
    
    If Mid(ss, 1, 9) = "PROCESANC" Then
        ProcesaNC = Mid(ss, 11, Len(ss) - 10)
    End If
    
    
    Wend
        Close 20
 
 
        usuario = "admixp"
        password = "adminplus_76465111"
 
 
 
        pat1 = "_licencia"
        pat2 = "mifranchitaflan"
        pater(1) = "erp_"
        pater(2) = "licencia_"
        pater(3) = "775753404"
        
'     Usuario = "erp_licencia"
'        password = "erp_licencia_775753404"
'
        

Call Conectartemporal(Servidor, "adminerp_inicio", "erp" + pat1, pater(1) + pater(2) + pater(3))



    
'    Call Conectartemporal(Servidor, "mysql", usuario, password)
'
    Call leerdatosconeccion("facturaelectronica.exe")
    basedatos = clientesistema + "gestion"
            
        baseteso = clientesistema & "teso"
        baseauditoria = clientesistema
        segundosespera = "60"


        Call Conectar(Servidor, basedatos, usuario, password)
        Call Conectar3(Servidor2, basedatos, usuario, password)
        
        
        rubro = leerRubro(empresaActiva)
        Call ConectarRubro(Servidor, basedatos, usuario, password)
        Call Conectarventas(Servidor, baseVentas & empresaActiva, usuario, password)
        iva = leerImpuesto("IVA")
        iha = leerImpuesto("IHA")
        fechasistema = Format(Now, "yyyy-mm-dd")
        empresa
        
    
    envia = False
    mensaje_nopermiso = "Usted no tiene privilegios suficientes para realizar esta operación."
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda General"
                Call Conectar_Auditoria
            Set sqlventas.conauditoria = conexionauditoria
'    Call COMPARAPROGRAMAS

    
        
        Call Centrar(Me)
        Call CargaGrillaInforme(1, 16)
        Call CargaGrillaLocales(1, 3)
        TIPO = "(dc.tipo = 'FV')"
        detalle = False
       
        dato1.text = Format(DateAdd("d", -10, fechasistema), "dd")
        dato2.text = Format(DateAdd("d", -10, fechasistema), "mm")
        dato3.text = Format(DateAdd("d", -10, fechasistema), "yyyy")
        dato4.text = dato1.text
        dato5.text = dato2.text
        dato6.text = dato3.text
         LEErlocales
     
    End Sub
Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventas
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion.g_maestroempresas WHERE facturadorelectronico='1' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        
        End If
        
End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatogrilla(10, 20) As String
        Dim I As Integer
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "TD"
        formatogrilla(1, 2) = "NUMERO"
        formatogrilla(1, 3) = "FECHA"
        formatogrilla(1, 4) = "RUT"
        formatogrilla(1, 5) = "CLIENTE"
        formatogrilla(1, 6) = "NETO"
        formatogrilla(1, 7) = "I.V.A"
        formatogrilla(1, 8) = "I.REF"
        formatogrilla(1, 9) = "I.VINOS"
        formatogrilla(1, 10) = "I.LIC"
        formatogrilla(1, 11) = "IHA "
        formatogrilla(1, 12) = "ICA "
        formatogrilla(1, 13) = "EXENTO"
        formatogrilla(1, 14) = "TOTAL"
        formatogrilla(1, 15) = "Nº SISTEMA"
        Rem LARGO DE LOS DATOS
        
        formatogrilla(2, 1) = "4"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "9"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "30"
        formatogrilla(2, 6) = "9"
        formatogrilla(2, 7) = "9"
        formatogrilla(2, 8) = "9"
        formatogrilla(2, 9) = "9"
        formatogrilla(2, 10) = "9"
        formatogrilla(2, 11) = "9"
        formatogrilla(2, 12) = "9"
        formatogrilla(2, 13) = "0"
        formatogrilla(2, 14) = "9"
        formatogrilla(2, 15) = "9"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "D"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        formatogrilla(3, 11) = "N"
        formatogrilla(3, 12) = "N"
        formatogrilla(3, 13) = "N"
        formatogrilla(3, 14) = "N"
        formatogrilla(3, 15) = "N"
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = "0000000000"
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = "###,###,##0"
        formatogrilla(4, 7) = "##,###,##0"
        formatogrilla(4, 8) = "##,###,##0"
        formatogrilla(4, 9) = "##,###,##0"
        formatogrilla(4, 10) = "##,###,##0"
        formatogrilla(4, 11) = "##,###,##0"
        formatogrilla(4, 12) = "##,###,##0"
        formatogrilla(4, 13) = "##,###,##0"
        formatogrilla(4, 14) = "###,###,##0"
        formatogrilla(4, 15) = "0000000000"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        If Verifica_Permiso(Me.Caption, "modifica") = True Then
        formatogrilla(5, 2) = "TRUE"
        Else
        formatogrilla(5, 2) = "TRUE"
        End If
        
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "TRUE"
        formatogrilla(5, 13) = "TRUE"
        formatogrilla(5, 14) = "TRUE"
        formatogrilla(5, 15) = "TRUE"
        
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
        formatogrilla(6, 10) = ""
        formatogrilla(6, 11) = ""
        formatogrilla(6, 12) = ""
        formatogrilla(6, 13) = ""
        formatogrilla(6, 14) = ""
        formatogrilla(6, 15) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        Rem ANCHO
        formatogrilla(8, 1) = "2"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "7"
        formatogrilla(8, 4) = "7"
        formatogrilla(8, 5) = "24"
       If opt2.Value = False Then
        formatogrilla(8, 6) = "7"
        formatogrilla(8, 7) = "5"
        formatogrilla(8, 8) = "5"
        formatogrilla(8, 9) = "5"
        formatogrilla(8, 10) = "5"
        formatogrilla(8, 11) = "5"
        formatogrilla(8, 12) = "5"
        Else
        formatogrilla(8, 6) = "0"
        formatogrilla(8, 7) = "0"
        formatogrilla(8, 8) = "0"
        formatogrilla(8, 9) = "0"
        formatogrilla(8, 10) = "0"
        formatogrilla(8, 11) = "0"
        formatogrilla(8, 12) = "0"
       End If
        formatogrilla(8, 13) = "0"
        
        formatogrilla(8, 14) = "7"
        formatogrilla(8, 15) = "7"
        formatogrilla(8, 16) = "7"
        
'        formatoGrilla(1, 7) = "I.V.A"
'        formatoGrilla(1, 8) = "I.REF"
'        formatoGrilla(1, 9) = "I.VINOS"
'        formatoGrilla(1, 10) = "I.LICORES"
'        formatoGrilla(1, 11) = "IHA "
'        formatoGrilla(1, 12) = "ICA "
        
                
        impresion.Cols = col + 3
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        impresion.Column(0).Alignment = cellLeftGeneral
        
        
        impresion.Column(0).Width = 16
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1.75
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        
        For I = 1 To impresion.Cols - 4
            impresion.Cell(0, I).text = formatogrilla(1, I)
            impresion.Column(I).Width = Val(formatogrilla(8, I)) * (impresion.Cell(0, I).Font.Size + 1.25)
            impresion.Column(I).MaxLength = Val(formatogrilla(2, I))
            impresion.Column(I).FormatString = formatogrilla(4, I)
            impresion.Column(I).Locked = formatogrilla(5, I)
            If formatogrilla(3, I) = "N" Then
                impresion.Column(I).Alignment = cellRightCenter
            End If
            If formatogrilla(3, I) = "S" Then
                impresion.Column(I).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, I) = "C" Then
                impresion.Column(I).Alignment = cellCenterCenter
            End If
        Next I
  impresion.Column(2).Mask = cellNumeric
  
        
        
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        If impresion.Rows > 1 Then
        Call imprimir
        End If
        
    End Sub
    
    Private Sub imprimir()
        Dim I As Long
        impresion.AutoRedraw = False
        impresion.Range(1, 1, 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThick
       
        impresion.PageSetup.HeaderMargin = 2
    
        impresion.PageSetup.TopMargin = 1
        impresion.PageSetup.LeftMargin = 0.5
        impresion.PageSetup.RightMargin = 0
        impresion.PageSetup.BottomMargin = 1
        
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.Orientation = cellLandscape
        impresion.PageSetup.PrintFixedRow = True
        
        
        Call verificaImpresora(5, impresion)
        
        impresion.AutoRedraw = True
    End Sub


Private Sub impresion_DblClick()
Dim K As Double
K = impresion.ActiveCell.row
Call GENERADTE(LOCAL_PROCESO, Mid(impresion.Cell(K, 0).text, 1, 2), impresion.Cell(K, 1).text, impresion.Cell(K, 2).text, impresion.Cell(K, 3).text, Mid(impresion.Cell(K, 4).text, 1, 10), impresion.Cell(K, 6).text, impresion.Cell(K, 7).text, impresion.Cell(K, 14).text, impresion.Cell(K, 8).text, impresion.Cell(K, 9).text, impresion.Cell(K, 10).text, impresion.Cell(K, 11).text, impresion.Cell(K, 12).text, Mid(impresion.Cell(K, 4).text, 11, 1), impresion.Cell(K, 16).text, impresion.Cell(K, 13).text)

Command1_Click


End Sub

    Private Sub opt1_Click()
        If opt1.Value = True Then
            TIPO = "(dc.tipo = 'FV')"
           Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    
    Private Sub opt2_Click()
        If opt2.Value = True Then
            TIPO = "(dc.tipo = 'BV')"
            Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    Private Sub opt3_Click()
        If opt3.Value = True Then
            TIPO = "(dc.tipo = 'ZE')"
            Call generaInformeLV(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    
    
Private Sub opt4_Click()
      If opt4.Value = True Then
            TIPO = "(dc.tipo = 'NB')"
             Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
      End Sub

    Private Sub opt5_Click()
        If opt5.Value = True Then
            TIPO = "(dc.tipo = 'NF')"
             Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If

End Sub

Private Sub Timer1_Timer()
Dim t As Double

If VERIFICAPING(Servidor) = True Then

Rem     sincronizarFechaHora
    
    
    If Check1.Value = 1 Then
        For t = 1 To Grid1.Rows - 1
        ' For T = 1 To 1
        LOCAL_PROCESO = Grid1.Cell(t, 1).text
    
        Call BORRARFOLIOSVENCIDOS(LOCAL_PROCESO)
        
        lbllocal.Caption = Grid1.Cell(t, 2).text & " " & Time
        lbllocal.Refresh
        empresa
        Command1_Click
        
        If impresion.Rows > 1 Then
        impresion.Cell(1, 1).SetFocus
        Call impresion_DblClick
        End If
        Sleep (200)
      Next t
    End If
End If



End Sub
Private Function listadte(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim tabla As String
    Dim rubAux As String
    Dim harinas As Double
    Dim subproductos As Double
    Dim envases As Double
    Dim trigo As Double
    Dim maquila As Double
    Dim otros As Double
    Dim cadena As String
    Dim tipodoc As String
    Dim numeroDoc As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim resultados As rdoResultset
    Dim sucursal As String
    
    Dim I As Integer

    rubAux = rubro
    Call Conectarventas(Servidor, baseVentas + codLoc, usuario, password)
    
    
    Set csql.ActiveConnection = ventasRubro
    sucursal = "0"
    csql.sql = "SELECT dc.tipo, dc.numero , dc.fecha, dc.rut, "
    csql.sql = csql.sql & "IFNULL(mc.nombre,'') as nombre, dc.neto, "
    csql.sql = csql.sql & "dc.iva, dc.exento, dc.total,dc.impuestoilarefrescos, "
    csql.sql = csql.sql & "dc.impuestoilavinos,dc.impuestoilalicores, "
    csql.sql = csql.sql & "dc.impuestoharina,dc.impuestocarne,dc.foliosii, "
    csql.sql = csql.sql & "dc.caja,dc.sucursal "
    csql.sql = csql.sql & "FROM sv_otros_documento_cabeza_" + codLoc + " AS dc left JOIN " & baseVentas & ".sv_maestroclientes AS mc ON "
    csql.sql = csql.sql & "dc.rut = mc.rut AND mc.sucursal = '" + sucursal + "'"
    csql.sql = csql.sql & "WHERE "
    fecha2 = DateAdd("m", -2, fechasistema)
    csql.sql = csql.sql & "fecha >= '" + Format(fecha2, "yyyy-mm-dd") + "' and "
    ' csql.sql = csql.sql & "AND dc.tipo='G1' "
    
    csql.sql = csql.sql & " dc.contabilizado='E' ": Rem  OR dc.tipo='NF' or dc.tipo='ND')  "
    csql.sql = csql.sql & "ORDER BY dc.tipo,dc.foliosii "
    csql.Execute
  
    linea = 0
    If csql.RowsAffected > 0 Then
       impresion.Rows = 1
       Set resultados = csql.OpenResultset
        While Not resultados.EOF
        
              If existedte(codLoc, resultados("tipo"), resultados("numero"), resultados("fecha"), resultados("caja"), "0") = True Then
            If ExistePDF(codLoc, resultados("tipo"), resultados("numero"), resultados("caja")) = False Then
            '   Stop
            If resultados("tipo") = "FV" Then dte_tipodte = "33"
             If resultados("tipo") = "NF" Or resultados("tipo") = "NB" Then dte_tipodte = "61"
            dte_folio = resultados("foliosii")
            Call imprimelectronica(dte_tipodte, CDbl(dte_folio), Format(resultados("fecha"), "yyyy-mm-dd"), Format(Mid(resultados("rut"), 1, 9), "#########") & "-" & Mid(resultados("rut"), 10, 1), resultados("numero"), resultados("caja"))
            End If
          End If
          
           If existedte(codLoc, resultados("tipo"), resultados("numero"), resultados("fecha"), resultados("caja"), "0") = False Then
           impresion.Rows = impresion.Rows + 1
           linea = linea + 1
            impresion.Cell(linea, 0).text = resultados("caja") + resultados("numero")
            impresion.Cell(linea, 1).text = resultados("tipo")
            impresion.Cell(linea, 2).text = resultados("numero")
            impresion.Cell(linea, 3).text = resultados("fecha")
            impresion.Cell(linea, 4).text = resultados("rut") & resultados("sucursal")
            impresion.Cell(linea, 5).text = resultados("nombre")
            impresion.Cell(linea, 6).text = resultados("neto")
            impresion.Cell(linea, 7).text = resultados("iva")
            impresion.Cell(linea, 8).text = resultados("impuestoilarefrescos")
            impresion.Cell(linea, 9).text = resultados("impuestoilavinos")
            impresion.Cell(linea, 10).text = resultados("impuestoilalicores")
            impresion.Cell(linea, 11).text = resultados("impuestoharina")
            impresion.Cell(linea, 12).text = resultados("impuestocarne")
            impresion.Cell(linea, 13).text = resultados("exento")
            impresion.Cell(linea, 14).text = resultados("total"): Rem resultados("descuentoporce")
            impresion.Cell(linea, 15).text = resultados("foliosii")
            
            End If
            resultados.MoveNext
        Wend
    
    End If
Set csql = Nothing
csql.Close
Set resultados = Nothing

    'Call sumaGrilla(impresion)
End Function
Public Function ExistePDF(empresa, TIPO, numero, caja) As Boolean
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = ventasRubro
csql.sql = "SELECT dte.tipo ,dte.numero ,COUNT(pdf.tipo)"
csql.sql = csql.sql & "  FROM " & clientesistema & "fae" & empresa & ".sv_dte" & empresa & " AS dte"
csql.sql = csql.sql & " LEFT JOIN " & clientesistema & "fae" & empresa & ".sv_dtepdf_" & empresa & " AS pdf "
csql.sql = csql.sql & " ON pdf.tipo=dte.tipo AND pdf.numero=dte.numero"
csql.sql = csql.sql & " WHERE dte.tipodocumento='" & TIPO & "'"
csql.sql = csql.sql & " AND dte.numerodocumento='" & numero & "'"
Rem - csql.sql = csql.sql & " AND dte.fechadocumento='" & Format(fecha, "yyyy-mm-dd") & "' "
csql.sql = csql.sql & " AND dte.cajadocumento='" & caja & "' "

csql.Execute
ExistePDF = False
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
        While Not resultados.EOF
 
            If resultados(2) >= 2 Then
                ExistePDF = True
            End If
            If resultados(2) < 2 Then
                ExistePDF = False
                dte_tipodte = resultados(0)
                dte_folio = resultados(1)
                Call EliminaPDF(empresa, resultados(0), resultados(1))
            End If
          
            resultados.MoveNext
        Wend
End If
csql.Close
 Set csql = Nothing
    
End Function
   Public Sub EliminaPDF(empresa, TIPO, numero)
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "delete FROM " & clientesistema & "fae" & empresa & ".sv_dtepdf_" & empresa
        csql.sql = csql.sql & "  WHERE  tipo = '" & TIPO & "' AND numero = '" & numero & "' "
        csql.Execute
       
'        If csql.RowsAffected > 0 Then
'            Set resultados = csql.OpenResultset
'            leerfoliodte = resultados(0) + 1
'        End If
        csql.Close
        Set csql = Nothing
    End Sub

Public Sub generaInformeLV(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim I As Long
    Dim documento As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    If TIPO = "FV" Then documento = "FACTURAS"
    If TIPO = "BV" Then documento = "BOLETAS "
    If TIPO = "ZE" Then documento = "ZETAS   "
    
    Call cargaCabeza("LISTADO DOCUMENTOS EMITIDOS  " + documento + " DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), codLoc, impresion)
    Call listadte(data, impresion, TIPO, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Sub leerdatosconeccion(nombre)
Dim CAMPOS(18, 3) As String
Dim op As Integer
Set sql = New sqlventas.sqlventa
    Call leerdatos_Certificado(usuario, password)
    CAMPOS(0, 0) = "usuariomysql"
    CAMPOS(1, 0) = "passwordmysql"
    CAMPOS(2, 0) = "cliente"
    CAMPOS(3, 0) = "rutaactualizaciones"
    CAMPOS(0, 2) = "admin_confi.clientes_admin "
    condicion = "sistema=" + "'" + nombre + "'"
    op = 5
    sql.response = CAMPOS
    Set sql.conexion = gestion
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
'    usuario = sql.response(0, 3)
'    password = sql.response(1, 3)
    clientesistema = sql.response(2, 3)
    rutaUpdate = sql.response(3, 3)
    sql.cliente_sql = clientesistema
    Else
    MsgBox ("NO EXISTE CONFIGURACION NI LICENCIA PARA ESTE SOFTWARE")
    Unload Me
    End If
    
no:

    
End Sub
 Private Sub CargaGrillaLocales(ByVal row As Integer, ByVal col As Integer)
        Dim formatogrilla(10, 20) As String
        Dim I As Integer
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "COD"
        formatogrilla(1, 2) = "NOMBRE"
        
        Rem LARGO DE LOS DATOS
        
        formatogrilla(2, 1) = "3"
        formatogrilla(2, 2) = "7"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "00"
        formatogrilla(4, 2) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
      
     
        
        col = 3
        Grid1.Cols = col
        Grid1.Rows = row
        Grid1.AllowUserResizing = True
        Grid1.AllowUserSort = True
        Grid1.DisplayFocusRect = False
        Grid1.ExtendLastCol = True
        Grid1.BoldFixedCell = False
        Grid1.Column(0).Alignment = cellLeftGeneral
        
        For I = 1 To Grid1.Cols - 1
            Grid1.Cell(0, I).text = formatogrilla(1, I)
            Grid1.Column(I).Width = Val(formatogrilla(2, I)) * (Grid1.Cell(0, I).Font.Size + 1.25)
            Grid1.Column(I).MaxLength = Val(formatogrilla(2, I))
            Grid1.Column(I).FormatString = formatogrilla(4, I)
            Grid1.Column(I).Locked = formatogrilla(5, I)
            If formatogrilla(3, I) = "N" Then
                Grid1.Column(I).Alignment = cellRightCenter
            End If
            If formatogrilla(3, I) = "S" Then
                Grid1.Column(I).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, I) = "C" Then
                Grid1.Column(I).Alignment = cellCenterCenter
            End If
        Next I
        Grid1.Column(0).Width = 0
        Grid1.Range(0, 1, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
        Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
   
       


    End Sub


