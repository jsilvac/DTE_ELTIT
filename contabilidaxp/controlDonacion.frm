VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ControlDonacion01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Personal"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   15525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "EXPORTAR A EXCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9360
      Width           =   3735
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   12120
      TabIndex        =   8
      Top             =   9360
      Width           =   3375
      _ExtentX        =   5953
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   280
         Width           =   1455
      End
   End
   Begin MSComctlLib.ProgressBar progreso 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   9000
      Visible         =   0   'False
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdimprimir 
      BackColor       =   &H00FF8080&
      Caption         =   "I M P R I M I R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9360
      Width           =   3735
   End
   Begin XPFrame.FrameXp FrameXp5 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   13573
      BackColor       =   16744576
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
      Begin XPFrame.FrameXp frmcheque 
         Height          =   2535
         Left            =   10080
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   4471
         BackColor       =   16761024
         Caption         =   "PANTALLA DATOS DEL CHEQUE"
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
         Begin VB.TextBox dato4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   30
            Top             =   1530
            Width           =   1815
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H0080FFFF&
            Caption         =   "INICIAR PROCESO"
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
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2040
            Width           =   2535
         End
         Begin VB.TextBox dato3 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   3420
            MaxLength       =   4
            TabIndex        =   28
            Top             =   405
            Width           =   735
         End
         Begin VB.TextBox dato2 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   2940
            MaxLength       =   2
            TabIndex        =   27
            Top             =   405
            Width           =   375
         End
         Begin VB.TextBox dato1 
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   2520
            MaxLength       =   2
            TabIndex        =   26
            Tag             =   "codigo"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox pivote 
            Height          =   285
            Left            =   4680
            TabIndex        =   25
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NUMERO INICIAL"
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
            Height          =   240
            Left            =   1800
            TabIndex        =   33
            Top             =   1260
            Width           =   1815
         End
         Begin VB.Label lblBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   90
            TabIndex        =   32
            Top             =   810
            Width           =   5145
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " BANCO"
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
            Height          =   285
            Left            =   990
            TabIndex        =   31
            Top             =   390
            Width           =   1455
         End
      End
      Begin FlexCell.Grid impresion 
         Height          =   7335
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   12938
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         DateFormat      =   2
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   0
         Top             =   0
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
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   3413
      BackColor       =   16744576
      Caption         =   ""
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
      Begin VB.TextBox codigotxt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8040
         MaxLength       =   5
         TabIndex        =   21
         Top             =   510
         Width           =   975
      End
      Begin VB.CheckBox reimprime 
         BackColor       =   &H00FF8080&
         Caption         =   "RE Impresion Comprobantes"
         Height          =   255
         Left            =   8880
         TabIndex        =   20
         Top             =   960
         Width           =   3015
      End
      Begin VB.CheckBox chk2 
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
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "MARCAR/DESMARCAR TODOS"
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
         Left            =   5400
         TabIndex        =   17
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Generar Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   12720
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   800
         Width           =   2625
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   615
         Left            =   7920
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "Seleccion del Local De Trabajo"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   4485
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Listar Personal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   12720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2625
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   615
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         BackColor       =   16744576
         Caption         =   "MES"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox COMBOMES 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   3615
         End
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         BackColor       =   16744576
         Caption         =   "AÑO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   240
            Width           =   3615
         End
      End
      Begin FlexCell.Grid Grid4 
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label lblnombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   9120
         TabIndex        =   23
         Top             =   510
         Width           =   3330
      End
      Begin VB.Label Lbl14 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Donacion"
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
         Left            =   8040
         TabIndex        =   22
         Top             =   240
         Width           =   4440
      End
   End
End
Attribute VB_Name = "ControlDonacion01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private tipo As String
    Private detalle As Boolean
    Private fecha1 As String
    Private fecha2 As String
    Private codigoempresa As String
    Private codigovendedor As String
    Private bonos(4, 4) As String
     Dim NCHEQUE As Double
    Private TOTALCheque As Double
    Private fechacheque As String
    Private NOMBREGIRADO As String
    Private FECHACONTABLE As String
    Private numerocontable As String
    Private tipocontable As String
    Private lineacontable As Double
    Private rutcontable As String
    Private cliqueado As Double
    Private empresaconsulta As String
    

Private Sub cajas_Click()
'If combolocal.text <> "" Then Call Command1_Click
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
Private Sub cajeras_Click()
'If combolocal.text <> "" Then Call Command1_Click
End Sub

Private Sub Check1_Click()
    Dim k As Double
    For k = 1 To impresion.Rows - 1
        If Val(impresion.Cell(k, 4).text) > 0 Then
            impresion.Cell(k, 5).text = Check1.Value
        End If
    Next k
End Sub

Private Sub cmdimprimir_Click()
If impresion.Rows > 1 Then Call imprimir
End Sub

Private Sub Command1_Click()
    Screen.MousePointer = vbHourglass
    progreso.Visible = True
        If lblnombre.Caption = "" Then
           codigotxt.text = ""
        End If
        
    codigoempresa = Mid(ComboLOCAL.text, 1, 2)
    Call CargaGrillaInforme(1, 6)
        Call generaInformeLV(data, impresion, tipo, detalle, codigoempresa, fecha1, fecha2)
        Call LEERTRABAJADORES(codigoempresa, Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, codigotxt.text)
         empresaconsulta = empresaactiva
        Screen.MousePointer = vbNormal
        progreso.Visible = False
End Sub
'Public Sub ayudaDonacion(ByRef txt As TextBox)
'        servidorAyuda = servidor
'        basedatosAyuda = clientesistema & "remu"
'        usuarioAyuda = Usuario
'        passAyuda = password
'        tablaAyuda = "mt_maestrodonaciones"
'        mensajeAyuda = "Ayuda Donaciones"
'        camposAyuda = Array("codigo", "glosa")
'        cabezasAyuda = Array("Codigo", "Nombre")
'        largoAyuda = Array("7n", "50s")
'        condicionAyuda = "no"
'        cantidadAyuda = 2
'        Call Mayuda.cargaAyuda(txt)
' End Sub
 Sub ayudaDonacion(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    Dim basedatosAyuda As String
     basedatosAyuda = clientesistema & "remu"
     
    campos = Array("codigo", "glosa")
    largo = Array("7n", "50s")
    cfijo = "autorizado='1'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda Donaciones"
        
    Call cargaAyudaT(Servidor, basedatosAyuda, Usuario, password, "mt_maestrodonaciones", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Private Sub codigotxt_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
        Call ayudaDonacion(codigotxt)
    End If
End Sub

Private Sub codigotxt_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(codigotxt)
        lblnombre.Caption = leernombredonacion(codigotxt.text)
        If lblnombre.Caption <> "" Then
            Command1.SetFocus
        Else
            codigotxt.text = ""
            
        End If
    End If
End Sub
 Function leernombredonacion(codigo) As String
    Dim op As Integer
    Dim campos(3, 3) As String
    Dim condicion As String
    
        campos(0, 0) = "glosa"
        campos(1, 0) = ""
        campos(0, 2) = clientesistema + "remu.mt_maestrodonaciones"
        condicion = "codigo  = '" & codigo & "' and autorizado='1' "

        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
            leernombredonacion = sqlconta.response(0, 3)
        Else
            leernombredonacion = ""
        End If
 End Function
Private Sub COMMAND2_Click()
    If impresion.Rows > 1 Then
        Call impresion.ExportToExcel("", True, True)
    End If
End Sub

Private Sub Command3_Click()
'    Dim k As Double
'    Dim i As Double
'    Dim THE As Double
'    Dim FHE As Double
'    Dim SB As Double
'    Dim total As Double
''    LeerParametrosdecalculo
''    FHE = param_he50
'
'
''                    FormatoGrilla(1, 4) = "REPOSO O LICENCIA"
''                    FormatoGrilla(1, 5) = "FALTANTES"
''                    FormatoGrilla(1, 6) = "HRS. EXTRAS"
''                    FormatoGrilla(1, 7) = "INCENTIVO 1"
''                    FormatoGrilla(1, 8) = "INCENTIVO 2 " & Format(bonos(1, 3), "$ ###,###,##0")
''                    FormatoGrilla(1, 9) = "INCENTIVO 3 " & Format(bonos(2, 3), "$ ###,###,##0")
''                    FormatoGrilla(1, 10) = "INCENTIVO 4 " & Format(bonos(3, 3), "$ ###,###,##0")
''                    FormatoGrilla(1, 11) = "INCENTIVO 5 " & Format(bonos(4, 3), "$ ###,###,##0")
'
'    If MsgBox("¿DESEA ENVIAR A LIQUIDACION LOS SELECCIONADOS?", vbYesNo, "ATENCION") = vbYes Then
'        SB = 0
'        total = 0
'        For k = 1 To impresion.Rows - 1
'            If impresion.Cell(k, 5).text = "1" Then
'                If impresion.Cell(k, 0).text <> "1" Then
'                    SB = leerdatostrabajador("monto", clientesistema & "remu" & empresaactiva & ".liquidacionhd", "rut='" & impresion.Cell(k, 1).text & "' and codtablacalculo='00001' and mes='" & Format(COMBOMES.ListIndex + 1, "00") & "' and año='" & COMBOAÑO & "'", db)
'                    Call grabarcontrolpersonal(Mid(combolocal.text, 1, 2), Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), impresion.Cell(k, 1).text, impresion.Cell(k, 2).text, impresion.Cell(k, 3).text, impresion.Cell(k, 4).text, impresion.Cell(k, 5).text)
'                    For i = 4 To 4
'                        If i = 4 And Val(impresion.Cell(k, i).text) > 0 And impresion.Cell(k, i).BackColor <> vbRed Then
'                            Call cargaraliquidacion(impresion.Cell(k, 1).text, Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, "00097", "DONACION A " & lblnombre.Caption, "D$", impresion.Cell(k, i).text, "00" & Mid(combolocal.text, 1, 2), "1", "1")
'                            Call imprimiranticipo(impresion.Cell(k, 1).text, impresion.Cell(k, 2).text, impresion.Cell(k, 4).text)
'                        End If
'                    Next i
'                 End If
'
'               If reimprime.Value = 1 Then
'                  SB = leerdatostrabajador("monto", clientesistema & "remu" & empresaactiva & ".liquidacionhd", "rut='" & impresion.Cell(k, 1).text & "' and codtablacalculo='00001' and mes='" & Format(COMBOMES.ListIndex + 1, "00") & "' and año='" & COMBOAÑO & "'", db)
'                    For i = 4 To 4
'                    Call cargaraliquidacion(impresion.Cell(k, 1).text, Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, "00097", "DONACION A " & lblnombre.Caption, "D$", impresion.Cell(k, i).text, "00" & Mid(combolocal.text, 1, 2), "1", "1")
'                    Call imprimiranticipo(impresion.Cell(k, 1).text, impresion.Cell(k, 2).text, impresion.Cell(k, 4).text)
'                    Next i
'               End If
'
'
'             End If
'
'
'        Next k
'        MsgBox "TODOS LOS SELECCIONADOS HAN SIDO PASADOS", vbInformation, "ATENCION "
'        Command1_Click
'    End If

    año = Format(fechasistema, "yyyy")
    MES = Format(fechasistema, "mm")
    
   If estacerrado(Format(fechasistema, "yyyy-mm-dd")) <> True Then
    If MsgBox("DESDEA PAGO MANUAL O ELECTRONICO,SI PARA MANUAL , NO PARA ELECTRONICO", vbYesNo, "ATENCION") = vbYes Then
        frmcheque.Visible = True
        dato1.SetFocus
    Else
        Call pagoelectronico
    End If
Else
    MsgBox "MES YA CERRADO"
End If

End Sub
Sub imprimiranticipo(ruttrabajador, nombretrabajador, monto)
    Dim row As Integer
    Dim FINROW As Integer
    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    CARGAGRILLA4
    'Logo
    'lista.Images.Add App.Path & "\Logo.gif", "Logo"
    'Set objReportTitle = New FlexCell.ReportTitle
    'objReportTitle.ImageKey = "Logo"
    'Grid3.ReportTitles.Add objReportTitle
    If monto = "0" Then Exit Sub
    Grid4.PageSetup.BlackAndWhite = False
    Grid4.PageSetup.BottomMargin = 1
    Grid4.PageSetup.LeftMargin = 1
    Grid4.PageSetup.RightMargin = 1
    Grid4.PageSetup.TopMargin = 1
    Grid4.PageSetup.PrintFixedRow = True
    Grid4.Column(1).Width = 13 * 8
    
    Call cabeza(ruttrabajador, nombretrabajador, monto)
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeTop) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeBottom) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeLeft) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeRight) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellInsideVertical) = cellThin
    FINROW = Grid4.Rows
'    Grid4.Cell(FINROW + 1, 0).text = ""
    Grid4.Rows = Grid4.Rows + 3
'    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).Borders(cellEdgeTop) = cellThin
    Grid4.Column(2).Locked = False
    Grid4.Column(3).Locked = False
    Grid4.Column(4).Locked = False
    Grid4.Column(5).Locked = False
    Grid4.Column(6).Locked = False
    Grid4.Column(7).Locked = False
    Grid4.Range(Grid4.Rows - 1, 2, Grid4.Rows - 1, 2).Merge
    Grid4.Range(Grid4.Rows - 1, 2, Grid4.Rows - 1, 2).Borders(cellEdgeTop) = cellThin
    Grid4.Cell(Grid4.Rows - 1, 2).text = "                   FIRMA EMPLEADOR"
    
    Grid4.Range(Grid4.Rows - 1, 4, Grid4.Rows - 1, 6).Merge
    Grid4.Range(Grid4.Rows - 1, 4, Grid4.Rows - 1, 6).Alignment = cellLeftCenter
    Grid4.Range(Grid4.Rows - 1, 4, Grid4.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
    Grid4.Cell(Grid4.Rows - 1, 4).text = "                   FIRMA TRABAJADOR"
    Grid4.Rows = Grid4.Rows + 1
    
    Grid4.Range(Grid4.Rows - 1, 4, Grid4.Rows - 1, 6).Merge
    Grid4.Range(Grid4.Rows - 1, 4, Grid4.Rows - 1, 6).Alignment = cellLeftCenter
'    Grid4.Range(Grid4.Rows - 1, 4, Grid4.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
    Grid4.Cell(Grid4.Rows - 1, 4).text = "                   " & Format(Mid(ruttrabajador, 1, 9), "###,###,###") & "-" & Mid(ruttrabajador, 10, 1)
    
    
    Grid4.Column(2).Locked = True
    Grid4.Column(3).Locked = True
    Grid4.Column(4).Locked = True
    Grid4.Column(5).Locked = True
    Grid4.Column(6).Locked = True
    Grid4.Column(7).Locked = True
    Grid4.PageSetup.BlackAndWhite = True
    
    
    
     'Logo
    'lista.Images.Add App.Path & "\Logo.gif", "Logo"
    'Set objReportTitle = New FlexCell.ReportTitle
    'objReportTitle.ImageKey = "Logo"
    'Grid3.ReportTitles.Add objReportTitle
    Grid4.PageSetup.BlackAndWhite = False
    Grid4.PageSetup.BottomMargin = 1
    Grid4.PageSetup.LeftMargin = 1
    Grid4.PageSetup.RightMargin = 1
    Grid4.PageSetup.TopMargin = 1
    Grid4.PageSetup.PrintFixedRow = True
    Grid4.Column(1).Width = 13 * 8
'    Call cabeza
 
    Grid4.PageSetup.BlackAndWhite = True
    If chk2.Value = 1 Then
        Grid4.PrintPreview
    Else
        Grid4.DirectPrint
    End If
    
    Grid4.Rows = FINROW
End Sub
Sub cabeza(ruttrabajador, nombretrabajador, monto)
Dim k As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid4.ReportTitles.Clear
    'Report Title 1
   
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = nombreempresa
        objReportTitle.Font.Name = "verdana"
        objReportTitle.Font.Size = 7
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        objReportTitle.color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid4.ReportTitles.Add objReportTitle
 
     
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = "COMPROBANTE DE ANTICIPO Nº " & numerotxt.text
'    objReportTitle.Font.Name = "verdana"
'    objReportTitle.Font.Size = 11
'    objReportTitle.Font.Bold = True
'    objReportTitle.PrintOnAllPages = True
'    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "En " & DATOSEMPRESA(4) & "                 Con Fecha : " & Format(fechasistema, "dd") & " de " & MonthName(Format(fechasistema, "mm")) & " " & Format(fechasistema, "yyyy")
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle

    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "En       cumplimiento     a       las     diposiciones    legales     vigentes       se     deja    costancia "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "que Don (a) : " & nombretrabajador & "                      RUT : " & Format(Mid(ruttrabajador, 1, 9), "###,###,###") & "-" & Mid(ruttrabajador, 10, 1)
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Recibe      conforme    y autoriza descontar    de sus     remuneraciones el    siguiente anticipo "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "MONTO ANTICIPO :$ " & Format(monto, "###,###,###")
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " SON " & WORDNUM(Format(monto, "########0"), "PESO", "PESOS", 0)
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle
    Grid4.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    Grid4.PageSetup.FooterAlignment = cellRight
    Grid4.PageSetup.FooterFont.Name = "Verdana"
    Grid4.PageSetup.FooterFont.Size = 7
    
    With Grid4.PageSetup
        .HeaderFont.Size = 6
        '.Header = "                                                                                                                   PAGINAS &P/&N EMITIDO:&D USUARIO " + USUARIOSISTEMA
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Verdana"
        .HeaderMargin = 4
    End With
End Sub
Sub CARGAGRILLA4()
Dim k As Integer
Dim FORMATOGRILLA(40, 40) As String

    Rem DATOS DE LA COLUMNA
    Grid4.DefaultFont.Size = 8
    Grid4.DefaultFont.Bold = False
    
    
    FORMATOGRILLA(1, 1) = ""
    FORMATOGRILLA(1, 2) = ""
    FORMATOGRILLA(1, 3) = ""
    FORMATOGRILLA(1, 4) = ""
    FORMATOGRILLA(1, 5) = ""
    FORMATOGRILLA(1, 6) = ""
    FORMATOGRILLA(1, 7) = ""
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "13"
    FORMATOGRILLA(2, 2) = "30"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"

    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = "###,##0.0"
    FORMATOGRILLA(4, 4) = "###,##0.0"
    FORMATOGRILLA(4, 5) = "###,##0.0"
    FORMATOGRILLA(4, 6) = "###,##0.0"
    FORMATOGRILLA(4, 7) = "###,##0.0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "FALSE"
    FORMATOGRILLA(5, 7) = "TRUE"
    
    Grid4.FixedRows = 1
    Grid4.Cols = 8
    Grid4.Rows = 1
    
    Grid4.AllowUserResizing = False
    Grid4.DisplayFocusRect = False
    Grid4.ExtendLastCol = True
    Grid4.BoldFixedCell = False
    Grid4.DrawMode = cellOwnerDraw
    Grid4.Appearance = Flat
    Grid4.ScrollBarStyle = Flat
    Grid4.FixedRowColStyle = Flat
    Grid4.BackColorFixed = RGB(90, 158, 214)
    Grid4.BackColorFixedSel = RGB(110, 190, 230)
    Grid4.BackColorBkg = RGB(90, 158, 214)
    Grid4.BackColorScrollBar = RGB(231, 235, 247)
    Grid4.BackColor1 = RGB(231, 235, 247)
    Grid4.BackColor2 = RGB(239, 243, 255)
    Grid4.GridColor = RGB(148, 190, 231)
    For k = 1 To Grid4.Cols - 1
        Grid4.Cell(0, k).text = FORMATOGRILLA(1, k)
        'Grid4.Cell(1, k).text = FORMATOGRILLA(8, k)
        Grid4.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid4.Cell(0, k).Font.Size
        Grid4.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid4.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid4.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid4.Column(k).Alignment = cellRightCenter
    Next k
    Grid4.Column(0).Width = 0
    Grid4.Range(0, 0, 0, Grid4.Cols - 1).Alignment = cellCenterCenter
    Grid4.Column(3).UserSortIndicator = cellSortIndicatorDescending
    Rem Grid4.Enabled = False
End Sub
    Sub cargaraliquidacion(ruttrabajador, MES, año, codtablacalculo, glosa, um, monto, codcentrocosto, duracion, duracionoriginal)
    
        campos(0, 0) = "rut"
        campos(1, 0) = "mes"
        campos(2, 0) = "año"
        campos(3, 0) = "codtablacalculo"
        campos(4, 0) = "glosa"
        campos(5, 0) = "um"
        campos(6, 0) = "monto"
        campos(7, 0) = "codcentrocosto"
        campos(8, 0) = "duracion"
        campos(9, 0) = "duracionoriginal"
        campos(10, 0) = "" '--
        
        campos(0, 1) = ruttrabajador
        campos(1, 1) = MES
        campos(2, 1) = año
        campos(3, 1) = codtablacalculo 'cod. tabla de calculo
        campos(4, 1) = glosa 'glosa
        campos(5, 1) = um 'um
        campos(6, 1) = Replace(monto, ",", ".") 'monto
        campos(7, 1) = codcentrocosto 'cod.centro costo
        campos(8, 1) = duracion 'duracion
        campos(9, 1) = duracionoriginal 'duracion
        campos(0, 2) = clientesistema & "remu" & empresaactiva & ".liquidacionhd" '--
        
        
        
'        If MODIFI = 1 Then 'si modifica
            condicion = "rut = '" & ruttrabajador & "'"
            condicion = condicion & " and mes = '" & MES & "'"
            condicion = condicion & " and año = '" & año & "'"
            condicion = condicion & " and codtablacalculo = '" & codtablacalculo & "'"
            op = 5 'modificar
'        Else

'        End If
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
        
        If sqlconta.status = 4 Then
            op = 2 'insertar
            condicion = ""
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
        Else
             campos(6, 1) = sqlconta.response(6, 3) + Replace(monto, ",", ".")
            op = 3 'insertar
            condicion = "rut = '" & ruttrabajador & "'"
            condicion = condicion & " and mes = '" & MES & "'"
            condicion = condicion & " and año = '" & año & "'"
            condicion = condicion & " and codtablacalculo = '" & codtablacalculo & "'"
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
            
        End If
    End Sub

Private Sub Command5_Click()
        
Dim k As Double
Dim rutprove As String
Dim tipo As String
Dim CUENTABANCO As String
Dim fechavencimiento  As String
Dim monto As Double
Dim DH As String
Dim MES As String
Dim año As String
Dim mesremu As String
Dim añoremu As String


If Format(fechasistema, "dd") <= "18" Then
    mesremu = Format(fechasistema, "mm")
    añoremu = Format(fechasistema, "yyyy")
Else
    mesremu = Format(DateAdd("m", 1, fechasistema), "mm")
    añoremu = Format(DateAdd("m", 1, fechasistema), "yyyy")
End If

    MES = Format(fechasistema, "mm")
    año = Format(fechasistema, "yyyy")
 
 If Verifica_FORM29(Format(fechasistema, "yyyy-mm-dd"), empresaactiva) = False Then
        If lblBanco.Caption <> "" Then
            If empresaconsulta <> empresaactiva Then MsgBox "EL LISTADO ES DE OTRA EMPRESA " & empresaconsulta, vbCritical, "ATENCION": GoTo no:
            NCHEQUE = CDbl(dato4.text) - 1
            tipocontable = "PA"
            numerocontable = LEERFOLIOCE("PA")
            lineacontable = 0
            TOTALCheque = 0
            For k = 1 To impresion.Rows - 1
                If impresion.Cell(k, 5).text = "1" And Val(impresion.Cell(k, 4).text) > 0 Then
                    fechacheque = Format(fechasistema, "yyyy-mm-dd")
                    NOMBREGIRADO = impresion.Cell(k, 2).text
                    FECHACONTABLE = Format(fechasistema, "yyyy-mm-dd")
                    rutprove = Mid(impresion.Cell(k, 1).text, 1, 9) + Mid(impresion.Cell(k, 1).text, 10, 1)
                    rutcontable = rutprove
                    CUENTABANCO = "11250007"
                    fechavencimiento = Format(fechasistema, "yyyy-mm-dd")
                    monto = CDbl(impresion.Cell(k, 4).text)
                    DH = "D"
        '  If verificasiexiste2(rutcontable, CUENTABANCO, FECHACONTABLE, tipocontable, "CANC. DONACION ") = False Then
                    lineacontable = lineacontable + 1
                    Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, _
                    FECHACONTABLE, CUENTABANCO, " ", rutcontable, " ", "CANC. DONACION " & NOMBREGIRADO, _
                    tipocontable, numerocontable, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, _
                    MES, año, _
                    Format(Date, "yyyy-mm-dd"), Time, rutcontable)
                    TOTALCheque = TOTALCheque + monto
        '        Else
        '            MsgBox "YA EXISTE COMPROBANTE  ", vbCritical, "ATENCION"
        '            frmcheque.Visible = False
        '            dato4.text = ""
        '            Exit Sub
        '       End If
                
            End If
        Next k
            
                If TOTALCheque <> 0 Then
                Call grabarcheque(TOTALCheque, MES, año)
                TOTALCheque = 0
                End If
                Call cargaraliquidaciones(mesremu, añoremu)
        'leer
            imprimir
            
        End If
             MsgBox " COMPROBANTE PA Nº" & numerocontable & " FUE GENERADO CON EXITO", vbInformation, "ATENCION"
no:
        frmcheque.Visible = False
        dato4.text = ""
Else
    MsgBox mensaje_nopermiso, vbCritical, "ATENCION"
End If
End Sub
Sub cargaraliquidaciones(MES, año)
    Dim SB As Double
    Dim TOTAL As Double
    Dim i As Double
    
        SB = 0
        TOTAL = 0
        For k = 1 To impresion.Rows - 1
            If impresion.Cell(k, 5).text = "1" Then
'                If impresion.Cell(k, 0).text <> "1" Then ' para no validar que ya esta
'                    SB = leerdatostrabajador("monto", clientesistema & "remu" & empresaactiva & ".liquidacionhd", "rut='" & impresion.Cell(k, 1).text & "' and codtablacalculo='00001' and mes='" & Format(fechasistema, "mm") & "' and año='" & Format(fechasistema, "yyyy") & "'", conta)
'                    Call grabarcontrolpersonal(Mid(ComboLOCAL.text, 1, 2), Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), impresion.Cell(k, 1).text, impresion.Cell(k, 2).text, impresion.Cell(k, 3).text, impresion.Cell(k, 4).text, codigotxt.text, empresaactiva)
                    For i = 4 To 4
                        If i = 4 And Val(impresion.Cell(k, i).text) > 0 Then 'And impresion.Cell(k, i).BackColor <> vbRed
                            Call cargaraliquidacion(impresion.Cell(k, 1).text, MES, año, "00097", "DONACION ", "D$", impresion.Cell(k, i).text, "00" & Mid(ComboLOCAL.text, 1, 2), "1", "1")
'                            Call imprimiranticipo(impresion.Cell(k, 1).text, impresion.Cell(k, 2).text, impresion.Cell(k, 4).text)
                        End If
                    Next i
'                 End If
                 
               If reimprime.Value = 1 Then
'                  SB = leerdatostrabajador("monto", clientesistema & "remu" & empresaactiva & ".liquidacionhd", "rut='" & impresion.Cell(k, 1).text & "' and codtablacalculo='00001' and mes='" & Format(fechasistema, "mm") & "' and año='" & Format(fechasistema, "yyyy") & "'", conta)
                    For i = 4 To 4
                    Call cargaraliquidacion(impresion.Cell(k, 1).text, MES, año, "00097", "DONACION", "D$", impresion.Cell(k, i).text, "00" & Mid(ComboLOCAL.text, 1, 2), "1", "1")
'                    Call imprimiranticipo(impresion.Cell(k, 1).text, impresion.Cell(k, 2).text, impresion.Cell(k, 4).text)
                    Next i
               End If
             End If
        
        Next k
      
End Sub
Private Sub dato1_GotFocus()
        Call cargatexto(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub
'****************************************************************************
'GOTFOCUS
'****************************************************************************

'****************************************************************************
'KEYDOWN
'****************************************************************************
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 38 Then Unload Me: GoTo no:
        If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)
        Call flechas(dato1, dato2, KeyCode)
no:
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato1, dato3, KeyCode)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato2, dato4, KeyCode)
    End Sub
'*********************************************
'KEYDOWN
'****************************************************************************

'****************************************************************************
'KEYPRESS
'****************************************************************************
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
           
          dato2.SetFocus
          
           
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
           dato3.SetFocus
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            lblBanco.Caption = leerNombreCuentaMayor(dato1.text & dato2.text & dato3.text, 3)
            If lblBanco.Caption <> "" Then
                
            dato4.SetFocus
            End If
        
        End If
    End Sub
    Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(dato4)
If leercheque(dato1.text + dato2.text + dato3.text, dato4.text) = True Then
MsgBox ("EL NUMERO DE CHEQUE YA ESTA EMITIDO")
dato4.text = ""
dato4.SetFocus
Else

Command5.SetFocus
End If

End If

End Sub
    Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "' AND banco='1'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    basebus = clientesistema + "conta" + empresaactiva
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato3.Enabled = True
    dato1.text = Mid(pivote.text, 1, 2)
    dato2.text = Mid(pivote.text, 3, 2)
    dato3.text = Mid(pivote.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
no:
End Sub
Sub grabarcheque(montocheque As Double, MES, año)
Dim tipodocumento As String
Dim numerodocumento As String
Dim CUENTABANCO As String
Dim fechavencimiento As String
Dim monto As Double
Dim DH As String



    Rem graba cheque
        
        NCHEQUE = NCHEQUE + 1
        lineacontable = lineacontable + 1
        
        If tipocontable = "CE" Then
            tipodocumento = "CH"
            numerodocumento = Format(NCHEQUE, "0000000000")
            CUENTABANCO = dato1.text + dato2.text + dato3.text
            fechavencimiento = fechacheque
            monto = montocheque
            
            
            Else
            
            tipodocumento = "CH"
            numerodocumento = Format(NCHEQUE, "0000000000")
            CUENTABANCO = "11120001"
            fechavencimiento = fechacheque
            monto = montocheque
            
        End If
        
        DH = "H"
        NOMBREGIRADO = lblnombre.Caption
        Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, FECHACONTABLE, CUENTABANCO, " ", "", " ", NOMBREGIRADO, tipodocumento, numerodocumento, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rutcontable)
        If tipocontable = "PA" Then
        fecha = Format(fechasistema, "yyyy-mm-dd")
        Call grabacheque(CUENTABANCO, numerodocumento, fecha, monto, fechavencimiento, tipocontable, numerocontable, NOMBREGIRADO, "0")
        End If
End Sub
Public Function LEERFOLIOCE(tipo) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = contadb
            csql.sql = "select max(numero) from movimientoscontables where mes = '" & Format(Format(fechasistema, "mm"), "00") & "' AND año = '" & Format(fechasistema, "yyyy") & "' and tipo='" + tipo + "' "
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
        If IsNull(resultados(0)) = False Then
        LEERFOLIOCE = Format(resultados(0) + 1, "0000000000")
        Else
        LEERFOLIOCE = Format(1, "0000000000")
        End If
        
    End If
    
End Function


Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = MES
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If rutctacte <> "" Then
        Call existerut(año, codigocuenta, rutctacte, empresaactiva)
    End If
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub
Sub grabacheque(cuenta, numero, emision, monto, vencimiento, tipocomprobante, numerocomprobante, giradoa, ubicacion)
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "tipocomprobante"
    campos(6, 0) = "numerocomprobante"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "ubicacion"
    campos(9, 0) = ""
    
    campos(0, 1) = cuenta
    campos(1, 1) = numero
    campos(2, 1) = emision
    campos(3, 1) = monto
    campos(4, 1) = vencimiento
    campos(5, 1) = tipocomprobante
    campos(6, 1) = numerocomprobante
    campos(7, 1) = giradoa
    campos(8, 1) = "0"
    campos(0, 2) = "chequesdocumento"
       
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
End Sub

Sub grabarrevisado(MES, año, ruttrabajador, empresa, estado, fecha)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "delete from " & clientesistema & "remu.mt_fijo_revisado "
    csql.sql = csql.sql & "where empresa='" & empresa & "' and año='" & año & "' "
    csql.sql = csql.sql & "and mes='" & MES & "' and ruttrabajador='" & ruttrabajador & "' "
    csql.Execute
'    Call sincronizadatos(csql.sql, contadb, "")
    
    
    csql.sql = "insert into " & clientesistema & "remu.mt_fijo_revisado (empresa,año,mes,ruttrabajador,estado,fecha)"
    csql.sql = csql.sql & "values('" & empresa & "','" & año & "','" & MES & "','" & ruttrabajador & "','" & estado & "','" & fecha & "') "
    csql.Execute
    
'    Call sincronizadatos(csql.sql, contadb, "")
    csql.Close
    Set csql = Nothing
    
    
End Sub
Function leerestado(empresaconsulta, MES, año, ruttrabajador) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select estado "
    csql.sql = csql.sql & "from " & clientesistema & "remu.mt_fijo_revisado "
    csql.sql = csql.sql & "where empresa='" & empresaconsulta & "' and año='" & año & "' "
    csql.sql = csql.sql & "and mes='" & MES & "' and ruttrabajador='" & ruttrabajador & "' "
    csql.Execute
        leerestado = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerestado = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
    
End Function


Public Sub existerut(año, tipo, rut, empresa)
 
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = nombretraba(rut, empresa)
    condicion = "tipo='" + tipo + "' and rut='" + rut + "' and año='" + año + "'  "
    campos(0, 2) = clientesistema + "conta" + empresa + ".cuentascorrientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    Else
    
    Call grabar(año, tipo, rut, nombretraba(rut, empresa), empresa)
    
    End If

    
    End Sub
    Sub grabar(año, tipo, rut, NOMBRE, empresa)
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = NOMBRE
    
    campos(0, 2) = clientesistema + "conta" + empresa + ".cuentascorrientes"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
     Call grabar2(año, tipo, rut, empresa)
    
    End Sub
Sub grabar2(año, tipo, rut, empresa)
      
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = ""
    
    campos(0, 1) = año
    campos(1, 1) = tipo
    campos(2, 1) = rut
    
    campos(0, 2) = clientesistema + "conta" + empresa + ".saldosctacte"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub
Public Function nombretraba(rut, empresa) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    condicion = "rut='" + rut + "' "
    campos(0, 2) = clientesistema + "remu" + empresa + ".mt_fijo "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    nombretraba = sqlconta.response(0, 3)
    Else
    nombretraba = ""
    End If
    End Function
    

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
        Call CENTRAR(Me)
        tipo = "(dc.tipo = 'FV')"
        detalle = False
        For k = 1 To 12
            COMBOMES.AddItem UCase(MonthName(k))
        Next k
'        If Format(fechasistema, "dd") <= 18 Then
            COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
'        Else
'            COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm"))
'        End If
        
        For k = 2000 To Val(Format(fechasistema, "yyyy")) + 1
            COMBOAÑO.AddItem k
        Next k
        COMBOAÑO.ListIndex = k - 2002
        
        LEErlocales
        Call CargaGrillaInforme(1, 6)
End Sub
 
    Private Sub imprimir()
        Dim i As Long
        Dim k As Double
        
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
      
        impresion.Cols = 7
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeTop) = cellThick
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThick
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThick
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeRight) = cellThick
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellInsideHorizontal) = cellThick
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellInsideVertical) = cellThick

        impresion.PageSetup.PrintFixedRow = True
        impresion.Column(5).Width = 0
        impresion.Cols = 7
        impresion.Cell(0, 6).text = "FIRMA"
        impresion.Column(6).Width = 250
         
        For k = 1 To impresion.Rows - 1
            If impresion.Cell(k, 4).text = 0 Then
                impresion.RowHeight(k) = 0
            End If
            impresion.Cell(k, 6).Border(cellEdgeTop) = cellThin
        Next k
        impresion.PrintPreview
        impresion.Column(5).Width = 66
        impresion.Cols = 6
         For k = 1 To impresion.Rows - 1
                impresion.RowHeight(k) = 18
             
        Next k
        
        
        impresion.AutoRedraw = True
    End Sub

Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM " & clientesistema & "gestion.g_maestroempresas "
        csql.sql = csql.sql & "where codigocontable='" & empresaactiva & "' "
        ' original cSql.sql = cSql.sql + "ORDER BY codigo "
        ' ariel agrega condicion local < 50 para que no liste locales 50 y 51
        csql.sql = csql.sql + "  and CODIGO < '53' ORDER BY codigo "
        
        If empresaactiva = "25" Or empresaactiva = "24" Or empresaactiva = "05" Or empresaactiva = "06" Then
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM " & clientesistema & "gestion.g_maestroempresas "
        'csql.sql = csql.sql & "where codigocontable='" & EMPRESAACTIVA & "' "
        ' original cSql.sql = cSql.sql + "ORDER BY codigo "
        ' ariel agrega condicion local < 50 para que no liste locales 50 y 51
        csql.sql = csql.sql + "  where CODIGO < '53' ORDER BY codigo "
        
        End If
        
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        
                
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        
End Sub
Public Sub generaInformeLV(ByRef data As Adodc, ByRef impresion As Grid, ByVal tipo As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Dim documento As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    Call cargaCabeza("HORAS DEL PERSONAL MES DE " & MonthName(Format(fechasistema, "mm")) & " DE " & Format(fechasistema, "yyyy"), Mid(ComboLOCAL, 1, 2), impresion)
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub
  Public Sub cargaCabeza(ByVal titulo As String, ByVal codLoc As String, ByRef impresion As Grid)
        Dim cabeza As FlexCell.ReportTitle
        Dim i As Integer
        
        impresion.ReportTitles.Clear
        'For i = impresion.ReportTitles.Count To 0 Step -1
        '    impresion.ReportTitles.Remove (i)
        'Next i
        
        Set cabeza = New FlexCell.ReportTitle
                
        cabeza.text = titulo
        cabeza.Align = cellCenter
        cabeza.Font.Bold = True
        cabeza.Font.Underline = True
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
        
'        Set CABEZA = New FlexCell.ReportTitle
'        CABEZA.text = leernombreempresa(codLoc)
'        CABEZA.Align = CellLeft
'        impresion.ReportTitles.Add CABEZA
'        CABEZA.PrintOnAllPages = True
'
'        Set CABEZA = New FlexCell.ReportTitle
'        CABEZA.text = leerDireccionEmpresa(codLoc)
'        CABEZA.Align = CellLeft
'        impresion.ReportTitles.Add CABEZA
'        CABEZA.PrintOnAllPages = True
'
'        Set CABEZA = New FlexCell.ReportTitle
'        CABEZA.text = "RUT: " & leerRutEmpresa(codLoc)
'        CABEZA.Align = CellLeft
'        impresion.ReportTitles.Add CABEZA
'        CABEZA.PrintOnAllPages = True
        
        Set cabeza = New FlexCell.ReportTitle
        cabeza.text = " "
        cabeza.Align = CellLeft
        impresion.ReportTitles.Add cabeza
        cabeza.PrintOnAllPages = True
    End Sub

Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim FORMATOGRILLA(10, 20) As String
        Dim i As Integer
 
        FORMATOGRILLA(1, 1) = "RUT"
        FORMATOGRILLA(1, 2) = "NOMBRE"
        FORMATOGRILLA(1, 3) = "INGRESO"
        FORMATOGRILLA(1, 4) = "MONTO"
        FORMATOGRILLA(1, 5) = "OK"
    
    
    Rem ANCHO DE LAS CELDAS
        FORMATOGRILLA(8, 1) = "10"
        FORMATOGRILLA(8, 2) = "24"
        FORMATOGRILLA(8, 3) = "24"
        FORMATOGRILLA(8, 4) = "24"
        FORMATOGRILLA(8, 5) = "7"
        
    
    Rem LARGO DE LOS DATOS
        FORMATOGRILLA(2, 1) = "10"
        FORMATOGRILLA(2, 2) = "30"
        FORMATOGRILLA(2, 3) = "30"
        FORMATOGRILLA(2, 4) = "50"
        FORMATOGRILLA(2, 5) = "3"
    
        Rem TIPO DE DATOS
        FORMATOGRILLA(3, 1) = "C"
        FORMATOGRILLA(3, 2) = "S"
        FORMATOGRILLA(3, 3) = "S"
        FORMATOGRILLA(3, 4) = "N"
        FORMATOGRILLA(3, 5) = "C"
        
        Rem FORMATO GRILLA
        FORMATOGRILLA(4, 1) = ""
        FORMATOGRILLA(4, 2) = ""
        FORMATOGRILLA(4, 3) = ""
        FORMATOGRILLA(4, 4) = ""
        FORMATOGRILLA(4, 5) = ""
        
        Rem LOCCKED

        FORMATOGRILLA(5, 1) = "TRUE"
        FORMATOGRILLA(5, 2) = "TRUE"
        FORMATOGRILLA(5, 3) = "TRUE"
        FORMATOGRILLA(5, 4) = "TRUE"
        FORMATOGRILLA(5, 5) = "FALSE"
     
        
        
        
        Rem VALOR MINIMO
        FORMATOGRILLA(6, 1) = ""
        FORMATOGRILLA(6, 2) = ""
        FORMATOGRILLA(6, 3) = ""
        FORMATOGRILLA(6, 4) = ""
        FORMATOGRILLA(6, 5) = ""
        
        
        Rem VALOR MAXIMO
        FORMATOGRILLA(7, 1) = ""
        FORMATOGRILLA(7, 2) = ""
        FORMATOGRILLA(7, 3) = ""
        FORMATOGRILLA(7, 4) = ""
        FORMATOGRILLA(7, 5) = ""
  
        Rem ANCHO

        
                
        impresion.Cols = col
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
        
        impresion.Column(0).Width = 0
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1.75
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = FORMATOGRILLA(1, i)
            impresion.Column(i).Width = Val(FORMATOGRILLA(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(FORMATOGRILLA(2, i))
            impresion.Column(i).FormatString = FORMATOGRILLA(4, i)
            impresion.Column(i).Locked = FORMATOGRILLA(5, i)
            If FORMATOGRILLA(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If FORMATOGRILLA(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If FORMATOGRILLA(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.Column(impresion.Cols - 1).CellType = cellCheckBox
    End Sub

Sub LEERTRABAJADORES(loc_consulta, mesconsulta, AÑOCONSULTA, codigodonacion)
   Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim cod_contable As String
    Dim TOTAL As Double
    
    cod_contable = empresaactiva
    Set csql.ActiveConnection = contadb
    
    csql.sql = "  SELECT mt.rut,mt.nombre,IFNULL(usuariocreacion,'') AS ingreso,IFNULL(monto,'0') AS monto,'" & cod_contable & "' "
    csql.sql = csql.sql & "FROM " & clientesistema & "remu" & cod_contable & ".mt_fijo AS mt INNER JOIN "
    csql.sql = csql.sql & "" & clientesistema & "remu" & cod_contable & ".mt_semipermanente AS sp ON mt.rut=sp.rut AND mt.mes=sp.mes AND mt.año=sp.año "
    csql.sql = csql.sql & "LEFT JOIN " & clientesistema & "remu.mt_donaciones_trabajadores AS cp ON cp.rut=mt.rut AND cp.año=mt.año AND "
    csql.sql = csql.sql & "cp.mes=mt.mes and cp.codigo='" & codigodonacion & "' "
    csql.sql = csql.sql & "WHERE mt.mes='" & mesconsulta & "' AND "
    csql.sql = csql.sql & "mt.año='" & AÑOCONSULTA & "' AND codigotg='0011' " 'AND codigog='" & Format(loc_consulta, "0000") & "' "
    
    
    csql.sql = csql.sql & "GROUP BY rut ORDER BY nombre "
    csql.Execute
    If csql.RowsAffected > 0 Then
        impresion.Rows = 1
        impresion.AutoRedraw = False
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            impresion.Rows = impresion.Rows + 1
           
            impresion.Cell(impresion.Rows - 1, 1).text = resultados(0)
            impresion.Cell(impresion.Rows - 1, 2).text = resultados(1)
            impresion.Cell(impresion.Rows - 1, 3).text = resultados(2)
            impresion.Cell(impresion.Rows - 1, 4).text = resultados(3)
'            impresion.Cell(impresion.Rows - 1, 5).text = resultados(4)
            TOTAL = TOTAL + resultados(3)
            Call verifica(impresion.Rows - 1)
            resultados.MoveNext
        Wend
        impresion.Rows = impresion.Rows + 1
        impresion.Cell(impresion.Rows - 1, 2).text = "TOTAL"
        impresion.Cell(impresion.Rows - 1, 4).text = Format(TOTAL, "$ ###,###,##0")
        impresion.AutoRedraw = True
        impresion.Refresh
        
    End If
    
 
End Sub
Sub verifica(LINEA)
    If leehdtrabajador(impresion.Cell(LINEA, 1).text, Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, "00097", "", empresaactiva) <> 0 Then
'        impresion.Cell(LINEA, 4).BackColor = vbRed
'        impresion.Cell(LINEA, 4).Locked = True
'        impresion.Cell(LINEA, 0).text = "1"
'        impresion.Cell(LINEA, 4).text = leehdtrabajador(impresion.Cell(LINEA, 1).text, Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, "00097", "", empresaactiva)
    End If
    
End Sub
Private Sub impresion_Click()
    If impresion.ActiveCell.col = 5 Then
'        If impresion.Cell(impresion.ActiveCell.row, 0).text <> "1" Then
'        Call grabarcontrolpersonal(Mid(combolocal.text, 1, 2), Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), impresion.Cell(impresion.ActiveCell.row, 1).text, impresion.Cell(impresion.ActiveCell.row, 2).text, USUARIOSISTEMA, impresion.Cell(impresion.ActiveCell.row, 4).text, codigotxt.text, empresaactiva)
'        Else
'            MsgBox "DONACION YA TRASPASADA A LIQUIDACION, NO SE PUEDE MODIFICAR", vbCritical, "ATENCION"
'            impresion.Cell(impresion.ActiveCell.row, 5).text = 0
'        End If
    End If
End Sub

Private Sub impresion_KeyPress(KeyAscii As Integer)
   If impresion.ActiveCell.col = 3 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = esNumero(KeyAscii)
    End If
     
End Sub

Private Sub impresion_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
 
'    If row <> NewRow Or col <> NewCol Then
'        Call grabarcontrolpersonal(Mid(combolocal.text, 1, 2), Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), impresion.Cell(row, 1).text, impresion.Cell(row, 2).text, USUARIOSISTEMA, impresion.Cell(row, 4).text, codigotxt.text, empresaactiva)
'
'    End If
End Sub
Sub grabarcontrolpersonal(loc, MES, año, ruttrabajador, nombretrabajador, usuariocreacion, monto, codigodonacion, empr)
     Dim op As Integer
     Dim campos(40, 3) As String
     
     
        
        campos(0, 0) = "codigo"
        campos(1, 0) = "rut"
        campos(2, 0) = "nombre"
        campos(3, 0) = "monto"
        campos(4, 0) = "mes"
        campos(5, 0) = "año"
        campos(6, 0) = "fechacreacion"
        campos(7, 0) = "usuariocreacion"
        campos(8, 0) = "localtrabajo"
        campos(9, 0) = "empresa"
        campos(10, 0) = ""
        
        
      
        campos(0, 1) = codigodonacion
        campos(1, 1) = ruttrabajador
        campos(2, 1) = nombretrabajador
        campos(3, 1) = Replace(monto, ",", ".")
        campos(4, 1) = MES
        campos(5, 1) = año
        campos(6, 1) = Format(fechasistema, "yyyy-mm-dd")
        campos(7, 1) = usuariocreacion
        campos(8, 1) = loc
        campos(9, 1) = empr
      
        
      
        campos(0, 2) = clientesistema & "remu.mt_donaciones_trabajadores"
        
        condicion = "localtrabajo='" & loc & "' and mes='" & MES & "' and año='" & año & "' and rut='" & ruttrabajador & "' "
        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 4 Then
            condicion = ""
            op = 2
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
        Else
            condicion = "localtrabajo='" & loc & "' and mes='" & MES & "' and año='" & año & "' and rut='" & ruttrabajador & "' "
            op = 3
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
        End If
End Sub
Public Function leehdtrabajador(rut, MES, año, codigo, donde, empr) As Double
    campos(0, 0) = "monto"
    campos(1, 0) = ""
    campos(0, 2) = clientesistema & "remu" & empr & ".liquidacionhd" 'tabla
    condicion = "rut='" + rut + "' and año='" + año + "' and mes='" + MES + "' and codtablacalculo='" + codigo + "' "
    op = 5
    leehdtrabajador = 0
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leehdtrabajador = sqlconta.response(0, 3)
    End If
    If donde = "PRE" Then
        campos(0, 0) = "monto"
        campos(1, 0) = ""
        campos(0, 2) = clientesistema & "remu" & empr & ".finiquitohd" 'tabla
        condicion = "rut='" + rut + "' and año='" + año + "' and mes='" + MES + "' and codtablacalculo='" + codigo + "' "
        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
        leehdtrabajador = leehdtrabajador + sqlconta.response(0, 3)
        End If
   End If
End Function
Sub pagoelectronico()
        
Dim k As Double
Dim rutprove As String
Dim tipo As String
Dim CUENTABANCO As String
Dim fechavencimiento  As String
Dim monto As Double
Dim DH As String
Dim MES As String
Dim año As String
Dim mesremu As String
Dim añoremu As String


If Format(fechasistema, "dd") <= "18" Then
    mesremu = Format(fechasistema, "mm")
    añoremu = Format(fechasistema, "yyyy")
Else
    mesremu = Format(DateAdd("m", 1, fechasistema), "mm")
    añoremu = Format(DateAdd("m", 1, fechasistema), "yyyy")
End If

    MES = Format(fechasistema, "mm")
    año = Format(fechasistema, "yyyy")
 
 If Verifica_FORM29(Format(fechasistema, "yyyy-mm-dd"), empresaactiva) = False Then
         
            If empresaconsulta <> empresaactiva Then MsgBox "EL LISTADO ES DE OTRA EMPRESA " & empresaconsulta, vbCritical, "ATENCION": GoTo no:
            NCHEQUE = CDbl(dato4.text) - 1
            tipocontable = "PA"
            numerocontable = LEERFOLIOCE("PA")
            lineacontable = 0
            TOTALCheque = 0
            For k = 1 To impresion.Rows - 1
                If impresion.Cell(k, 5).text = "1" And Val(impresion.Cell(k, 4).text) > 0 Then
                    fechacheque = Format(fechasistema, "yyyy-mm-dd")
                    NOMBREGIRADO = impresion.Cell(k, 2).text
                    FECHACONTABLE = Format(fechasistema, "yyyy-mm-dd")
                    rutprove = Mid(impresion.Cell(k, 1).text, 1, 9) + Mid(impresion.Cell(k, 1).text, 10, 1)
                    rutcontable = rutprove
                    CUENTABANCO = "11250007"
                    fechavencimiento = Format(fechasistema, "yyyy-mm-dd")
                    monto = CDbl(impresion.Cell(k, 4).text)
                    DH = "D"
                           lineacontable = lineacontable + 1
                    Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, _
                    FECHACONTABLE, CUENTABANCO, " ", rutcontable, " ", "CANC. DONACION " & NOMBREGIRADO, _
                    tipocontable, numerocontable, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, _
                    MES, año, _
                    Format(Date, "yyyy-mm-dd"), Time, rutcontable)
                    TOTALCheque = TOTALCheque + monto
        
                
            End If
        Next k
            
                If TOTALCheque <> 0 Then
                Call grabarcheque2(TOTALCheque, MES, año)
                TOTALCheque = 0
                End If
                Call cargaraliquidaciones(mesremu, añoremu)
        'leer
            imprimir
            
         
             MsgBox " COMPROBANTE PA Nº" & numerocontable & " FUE GENERADO CON EXITO", vbInformation, "ATENCION"
no:
        frmcheque.Visible = False
        dato4.text = ""
Else
    MsgBox mensaje_nopermiso, vbCritical, "ATENCION"
End If

End Sub
Sub grabarcheque2(montocheque As Double, MES, año)
Dim tipodocumento As String
Dim numerodocumento As String
Dim CUENTABANCO As String
Dim fechavencimiento As String
Dim monto As Double
Dim DH As String



    Rem graba cheque
        
     
        lineacontable = lineacontable + 1
        
      
            tipodocumento = "DB"
            numerodocumento = Format(NCHEQUE, "0000000000")
            CUENTABANCO = "11500160"
            fechavencimiento = fechacheque
            monto = montocheque
      
        
        DH = "H"
        NOMBREGIRADO = lblnombre.Caption
        Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, FECHACONTABLE, CUENTABANCO, " ", "", " ", NOMBREGIRADO, tipodocumento, numerodocumento, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rutcontable)
        
End Sub

