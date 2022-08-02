VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form activos_maestro 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "GRABAR"
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   5520
      Width           =   3015
   End
   Begin XPFrame.FrameXp frmBaja 
      Height          =   1095
      Left            =   120
      TabIndex        =   35
      Top             =   4080
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1931
      BackColor       =   16761024
      Caption         =   "ESTADO DEL ACTIVO"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin MSComCtl2.DTPicker fechabaja 
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   161218561
         CurrentDate     =   42039
      End
      Begin VB.Label lblusuariobaja 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " USUARIO AUTORIZACION  BAJA DE ACTIVO"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   2040
         TabIndex        =   39
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA BAJA"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog explorafoto 
      Left            =   11880
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   4095
      Left            =   8040
      TabIndex        =   8
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7223
      BackColor       =   16761024
      Caption         =   "FOTO DEL ACTIVO"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
         Height          =   375
         Left            =   3840
         Picture         =   "activos_maestro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label kb 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3840
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   3495
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4095
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2175
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3836
      BackColor       =   16761024
      Caption         =   "Datos de la adquisicion"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComCtl2.DTPicker fechacompra 
         Height          =   285
         Left            =   6000
         TabIndex        =   34
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   161218561
         CurrentDate     =   42039
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   7
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox dato 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   6
         Left            =   2040
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   5
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   4
         Left            =   2040
         MaxLength       =   9
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA COMPRA"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   4080
         TabIndex        =   33
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label dv 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3600
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblNombreEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2640
         TabIndex        =   21
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label lblNombreProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMPRESA CONTABLE"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " COSTO CON IVA"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FACTURA"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROVEEDOR"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3201
      BackColor       =   16761024
      Caption         =   "Maestro de Activos"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   3
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   13
         TabIndex        =   1
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   13
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblnombretipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO ACTIVO"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DESCRIPCION"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO SERIE"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CODIGO BARRA"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin XPFrame.FrameXp FrameXp6 
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      Top             =   7800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2143
      BackColor       =   16761024
      Caption         =   "OPCIONES"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   192
      ColorBarraArriba=   255
      ColorBarraAbajo =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   192
      Begin Contabilidadxp.BotonMyERP opcion 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Modificar"
         PicturePosition =   0
         Picture         =   "activos_maestro.frx":043E
         PictureHover    =   "activos_maestro.frx":10F4
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XPColor_Pressed =   16761024
         XPColor_Hover   =   16761024
         BackColor       =   16761024
      End
      Begin Contabilidadxp.BotonMyERP opcion 
         Height          =   855
         Index           =   1
         Left            =   960
         TabIndex        =   26
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Eliminar"
         PicturePosition =   0
         Picture         =   "activos_maestro.frx":1E55
         PictureHover    =   "activos_maestro.frx":2B72
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XPColor_Pressed =   16761024
         XPColor_Hover   =   16761024
         BackColor       =   16761024
      End
      Begin Contabilidadxp.BotonMyERP opcion 
         Height          =   855
         Index           =   2
         Left            =   2760
         TabIndex        =   27
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Eventos"
         PicturePosition =   0
         Picture         =   "activos_maestro.frx":390B
         PictureHover    =   "activos_maestro.frx":4592
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XPColor_Pressed =   16761024
         XPColor_Hover   =   16761024
         BackColor       =   16761024
      End
      Begin Contabilidadxp.BotonMyERP opcion 
         Height          =   855
         Index           =   3
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Historia"
         PicturePosition =   0
         Picture         =   "activos_maestro.frx":52D5
         PictureHover    =   "activos_maestro.frx":6080
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XPColor_Pressed =   16761024
         XPColor_Hover   =   16761024
         BackColor       =   16761024
      End
      Begin Contabilidadxp.BotonMyERP opcion 
         Height          =   855
         Index           =   4
         Left            =   3720
         TabIndex        =   29
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         Caption         =   "Retorno"
         PicturePosition =   0
         Picture         =   "activos_maestro.frx":6E89
         PictureHover    =   "activos_maestro.frx":7BB2
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XPColor_Pressed =   16761024
         XPColor_Hover   =   16761024
         BackColor       =   16761024
      End
   End
   Begin XPFrame.FrameXp FrmMovimientos 
      Height          =   3495
      Left            =   120
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6165
      BackColor       =   16761024
      Caption         =   "ULTIMOS MOVIMIENTOS"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command1 
         Caption         =   "RETORNO"
         Height          =   375
         Left            =   2160
         TabIndex        =   41
         Top             =   3000
         Width           =   3015
      End
      Begin FlexCell.Grid Grid1 
         Height          =   2655
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4683
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DAR DE BAJA"
      Height          =   615
      Left            =   5520
      TabIndex        =   42
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin XPFrame.FrameXp FrmEventos 
      Height          =   5775
      Left            =   120
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10186
      BackColor       =   16761024
      Caption         =   "EVENTOS"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox txtGlosa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   47
         Top             =   4920
         Width           =   6615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "RETORNO"
         Height          =   375
         Left            =   4680
         TabIndex        =   44
         Top             =   5280
         Width           =   3015
      End
      Begin FlexCell.Grid Grid2 
         Height          =   4575
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   8070
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " GLOSA"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   46
         Top             =   4920
         Width           =   855
      End
   End
End
Attribute VB_Name = "activos_maestro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ****************datos tamaño
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long
Option Explicit
Public cnn As ADODB.Connection
Public rst As ADODB.Recordset
Dim FOLIO, rut, Conex, CONSULTA, csql As String
Dim UserSQL, PassSQL, ServerSQL, BdSQL, TablaSQL, ImgNombre As String

' ××××××××××××××××DATOS FOTOS
Dim conn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim mystream As New ADODB.Stream
Dim Img As String
Dim KbImagen As Long
Dim loadStr, ImgTemporal, Imgtemporal2 As String
Dim n As Double
' ****************datos scan
Dim iX As Integer
Dim iY As Integer
Dim clrHashForeColor
Dim clrHashBackColor
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
      lpPictDesc As PictDesc, _
      riid As Guid, _
      ByVal fPictureOwnsHandle As Long, _
      ipic As IPicture _
    ) As Long


Private Sub Command1_Click()
FrmMovimientos.Visible = False
End Sub

Private Sub cmdgrabar_Click()
Call grabar
End Sub

Private Sub COMMAND2_Click()
If ExisteActivo(dato(0)) = False Then
    MsgBox "DEBE GRABAR PRIMERO EL ACTIVO PARA SUBIR IMAGEN", vbExclamation, "ATENCION"
    Exit Sub
End If

If TieneFoto(dato(0).text) = True Then
    If MsgBox("ACTIVO CON FOTO ASIGNADA" & vbNewLine & "DESEA CONTINUAR?" & vbNewLine & "LA IMAGEN ACTUAL SE ELIMINARA", vbYesNo) = vbNo Then
    Exit Sub
    End If
End If

explorafoto.DialogTitle = "BUSCA IMAGEN DEL DISPOSITIVO"
explorafoto.DefaultExt = "*.JPG"
explorafoto.ShowOpen
If explorafoto.FileName <> "" Then
     Call ConexionImg(1)
     Call ConexionImg(2)
End If
End Sub

Private Sub Command3_Click()
If MsgBox("SE VA A DAR DE BAJA ESTE ACTIVO" & vbNewLine & " NO PODRA ASIGNARLO, NI MODIFICARLO, NI ELIMINARLO" & vbNewLine & " DESEA CONTINUAR?", vbYesNo, "ATENCION") = vbYes Then
    Call DardeBaja

End If
End Sub

Private Sub Command4_Click()
FrmEventos.Visible = False
End Sub

Private Sub dato_Change(Index As Integer)
Select Case Index
Case 3
    lblnombretipo.Caption = Empty
Case 4
    DV.Caption = Empty
    lblnombreproveedor.Caption = Empty
Case 7
    lblNombreEmpresa.Caption = Empty
End Select
End Sub

Private Sub dato_GotFocus(Index As Integer)
Call cargatexto(dato(Index))
End Sub

Private Sub dato_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 On Error GoTo error
If KeyCode = vbKeyF2 Then
    Select Case Index
        Case 3
            Call AyudaActivos_Tipos(dato(Index))
        Case 4
            Call ayudactacte(dato(Index))
            Call dato_KeyPress(Index, 13)
        Case 7
            Call AyudaActivos_EmpresaContable(dato(Index))
        Case 8
            Call AyudaActivos_Empresa(dato(Index))
        Case 9
            Call AyudaActivos_Ubicaciones(dato(Index))
        Case 10
            Call AyudaActivos_Usuarios(dato(Index))
    End Select
End If
Call flechas(dato(Index - 1), dato(Index + 1), KeyCode)
error:
End Sub

Private Sub dato_KeyPress(Index As Integer, KeyAscii As Integer)
   On Error GoTo error
If KeyAscii = 13 Then
    Select Case Index
    Case 1
        Case 0
           If Index <> 0 Then Call ceros(dato(Index))
        If Index = 0 Then
            If LeerActivo(dato(0)) = True Then
            Command2.Visible = True
            CmdGrabar.Visible = False
            Call CargarHistorialRadios(dato(0))
            If TieneFoto(dato(0).text) = True Then
                Call ConexionImg(2)
            End If
            Else
                dato(1).SetFocus
            End If
        End If
        Case 4
            Call ceros(dato(Index))
            DV.Caption = Validaciones.rut(dato(4).text)
            lblnombreproveedor.Caption = LEERNOMBREPROVEEDOR(dato(4) & DV.Caption)
            If lblnombreproveedor.Caption = "" Then
                dato(4).text = Empty
                dato(4).SetFocus
                Exit Sub
            End If
        Case 3
            Call ceros(dato(Index))
            lblnombretipo.Caption = LeerNombreActivos_tipo(dato(Index))
        Case 2, 6, 10
        Case 7
            Call ceros(dato(Index))
            lblNombreEmpresa.Caption = leerNombreEmpresa(dato(7).text)
        Case Else
            Call ceros(dato(Index))
    End Select
If dato(Index).text <> "" Then dato(Index + 1).SetFocus
    
End If




Select Case Index
    Case 1, 2, 10
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case Else
        snum = 0: KeyAscii = esNumero(KeyAscii)
End Select
error:
'CUANDO LLEGA AL FINAL DE LA MATRIZ
End Sub



Public Function LeerActivo(codigo) As Boolean
Dim campos(20, 20) As String
Dim op As Double
Dim condicion As String
    campos(0, 0) = "numeroserie"
    campos(1, 0) = "descripcion"
    campos(2, 0) = "tipo"
    campos(3, 0) = "proveedor"
    campos(4, 0) = "factura"
    campos(5, 0) = "pcosto"
    campos(6, 0) = "empresacontable"
    campos(7, 0) = "ifnull(fecha_compra,0)"
    campos(8, 0) = "ifnull(fechabaja,0)"
    campos(9, 0) = "usuariobaja"
    campos(10, 0) = ""
    campos(11, 0) = ""
    campos(12, 0) = ""
    
    LeerActivo = False
    
    campos(0, 2) = clientesistema & "conta.af_maestro_activos"
    condicion = "codigobarras='" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LeerActivo = True
        dato(1).text = sqlconta.response(0, 3)
        dato(2).text = sqlconta.response(1, 3)
        dato(3).text = sqlconta.response(2, 3)
        dato(4).text = sqlconta.response(3, 3)
        dato(5).text = sqlconta.response(4, 3)
        dato(6).text = sqlconta.response(5, 3)
        dato(7).text = sqlconta.response(6, 3)
        fechacompra = sqlconta.response(7, 3)
        DV.Caption = Validaciones.rut(dato(4).text)
        If sqlconta.response(8, 3) <> "0000-00-00" Then
            frmBaja.Visible = True
            fechabaja = sqlconta.response(8, 3)
            lblusuariobaja = sqlconta.response(9, 3)
            Command3.Visible = False
        Else
            frmBaja.Visible = False
            Command3.Visible = True
        End If
        
        lblnombreproveedor = LEERNOMBREPROVEEDOR(sqlconta.response(3, 3))
        lblnombretipo = LeerNombreActivos_tipo(dato(3).text)
        lblNombreEmpresa.Caption = leerNombreEmpresa(dato(7).text)
        Call CargarEventos(codigo)
    For n = 0 To dato.Count - 1
    dato(n).Locked = True
    Next n
    Else
    For n = 0 To dato.Count - 1
    dato(n).Locked = False
    Next n

    End If
    
End Function


Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & CUENTAPROVEEDOR & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Proveedores "
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes", caja, campos, cfijo, largo, 2)

    If Val(caja.text) = 0 Then caja.SetFocus: GoTo no
 
    caja.text = Mid(caja.text, 1, 9)
    DV.Caption = Mid(caja.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub


Public Function leerNombreEmpresa(ByVal codigo As String) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa = '" & codigo & "'  "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        leerNombreEmpresa = sqlconta.response(0, 3)
    Else
        leerNombreEmpresa = ""
    End If
End Function



Sub CARGAGRILLA()
Grid1.Rows = 1
Grid1.Cols = 4
Grid1.Cell(0, 1).text = "USUARIO A CARGO"
Grid1.Cell(0, 2).text = "FECHA ENTREGA"
Grid1.Cell(0, 3).text = "FECHA DEVOLUCION"
Grid1.RowHeight(0) = 40
Grid1.Column(0).Width = 0
Grid1.Column(1).Width = 300
Grid1.Column(1).Locked = True
Grid1.Column(2).Locked = True
Grid1.Column(3).Locked = True
Grid1.Range(0, 1, 0, Grid1.Cols - 1).WrapText = True
Grid1.ExtendLastCol = True

Grid2.Rows = 1
Grid2.Cols = 4
Grid2.RowHeight(0) = 40
Grid2.Cell(0, 1).text = "FECHA"
Grid2.Cell(0, 2).text = "USUARIO"
Grid2.Cell(0, 3).text = "GLOSA"
Grid2.Column(0).Width = 0
Grid2.Column(1).Width = 100
Grid2.Column(2).Width = 100
Grid2.Column(3).Width = 300
Grid2.Column(1).Locked = True
Grid2.Column(2).Locked = True
Grid2.Column(3).Locked = True


End Sub


Private Sub Form_Load()
Call CARGAGRILLA
Imgtemporal2 = App.path & "\tmp.jpg"
fechacompra = fechasistema
End Sub



Sub CargarHistorialRadios(codigo)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    LINEA = 1
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT nombre,fechaentrega,IFNULL(fechadevolucion,0) AS devolucion  "

    csql.sql = csql.sql & " FROM " & clientesistema & "conta.af_movimiento_activos "
    csql.sql = csql.sql & " where CODIGO = '" & codigo & "' ORDER BY fechaentrega"
    csql.Execute
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    While resultados.EOF = False
        Grid1.AddItem "", True
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
        Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
        resultados.MoveNext
    Wend
    
    End If
   Grid1.AutoRedraw = True
   Grid1.Refresh
   

End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 46 Then
    If Verifica_Permiso(Me.Caption, "autoriza") = True Then
        If Grid2.Cell(Grid2.ActiveCell.row, 1).Tag <> "" Then
            If MsgBox("SE ELIMINARA EL EVENTO SELECCIONADO" & vbNewLine & "DESEA CONTINUAR?", vbYesNo, "ATENCION") = vbYes Then
            
                Call EliminarEvento(dato(0).text, Grid2.Cell(Grid2.ActiveCell.row, 1).Tag)
                Call CargarEventos(dato(0).text)
            End If
        End If
    Else
        MsgBox "USUARIO SIN PRIVILEGIOS PARA ELIMINAR EVENTOS", vbExclamation, "ATENCION"
    End If
End If
End Sub

Private Sub opcion_Click(Index As Integer)
Dim n As Double
Select Case Index
Case 0
    If Verifica_Permiso(Me.Caption, "modifica") = True Then
        If frmBaja.Visible = False Then
            CmdGrabar.Visible = True
            For n = 1 To dato.Count - 1
                dato(n).Locked = False
            Next n
            dato(0).Locked = True
            dato(1).SetFocus
        Else
            MsgBox "ACTIVO DE BAJA" & vbNewLine & " NO PUEDE MODIFICAR", vbExclamation, "ATENCION"
        End If
    End If
Case 1
If dato(0).text <> "" Then
    If Grid1.Rows - 1 = 0 Then
        If MsgBox("DESEA ELIMINAR ESTE ACTIVO?", vbYesNo, "ATENCION") = vbYes Then
            If frmBaja.Visible = False Then
            Call ELIMINARFoto(dato(0).text)
            Call EliminarActivo(dato(0).text)
            Else
                MsgBox "ACTIVO DE BAJA" & vbNewLine & " NO PUEDE ELIMINAR", vbExclamation, "ATENCION"
            End If
            GoTo retorno
        End If
    Else
        MsgBox "ACTIVO CON MOVIMIENTOS" & vbNewLine & " NO PUEDE ELIMINAR", vbExclamation, "ATENCION"
        Exit Sub
    End If
End If
Case 2
If dato(1).text <> "" Then
Call CargarEventos(dato(0).text)
    FrmEventos.Visible = True
    txtGlosa.text = Empty
    FrmEventos.ZOrder (0)
    txtGlosa.Tag = LeerUltimoItem(dato(0).text)
    txtGlosa.SetFocus
'    FrmMovimientos.Width = Me.Width
'    Grid1.Width = Me.Width - 100
'    FrmMovimientos.Height = Me.Height
'    FrmMovimientos.Top = 0
'    FrmMovimientos.Left = 0
    
 End If

Case 3
If dato(1).text <> "" Then
    FrmMovimientos.Visible = True
'    FrmMovimientos.Width = Me.Width
'    Grid1.Width = Me.Width - 100
'    FrmMovimientos.Height = Me.Height
'    FrmMovimientos.Top = 0
'    FrmMovimientos.Left = 0
    FrmMovimientos.ZOrder (0)
 End If
Case 4
retorno:
        Image1.Picture = LoadPicture("")
        Grid1.Rows = 1
        For n = 0 To dato.Count - 1
        dato(n).Locked = False
        dato(n).text = Empty
        Next n
        CmdGrabar.Visible = True
        dato(0).SetFocus
        frmBaja.Visible = False
        fechacompra = fechasistema
        Command3.Visible = False
End Select
End Sub





Public Sub cons(CONSULTA, OPERACION)
 On Error GoTo error
Set cnn = Nothing
Set rst = Nothing
Set cnn = New ADODB.Connection
Set rst = New ADODB.Recordset
cnn.Open Conex
Set rst = Nothing
Set rst = New ADODB.Recordset
rst.Open CONSULTA, cnn, adLockOptimistic
'If OPERACION = 1 Then       'CARGA DATOS AL LISTADO
'        'Dim ITEM As String
'    While rst.EOF = False
'        With Listado
'        'If rst(0) = 1 Then ITEM = "FACTURA"
'            .AddItem rst(0) & vbTab & rst(1) & vbTab & rst(2)
'            rst.MoveNext
'        End With
'    Wend
'End If

Exit Sub
If OPERACION = 2 Then       'CARGA LA IMAGEN DESDE EL LISTADO
If rst.EOF = False Then
    Call ConexionImg(2)
End If
End If
Exit Sub
error:
    MsgBox err.Description
Exit Sub
End Sub

Private Sub RESIZE()
kb = Empty
KbImagen = Empty
Dim retval As Long
MousePointer = vbHourglass 'CAMBIO EL PUNTERO A OCUPADO
loadStr = explorafoto.FileName
SavePicture Image1.Picture, ImgTemporal
 On Error GoTo error
retval = DIWriteJpg(loadStr, 80, 1)
If retval = 1 Then  'correcto
   Image1.Picture = LoadPicture(loadStr)
   KbImagen = Mid(Str(FileLen(loadStr)), 1, Len(Str(FileLen(loadStr))) - 3)
Else                'ocurrió un error
   MsgBox "La conversión NO fue exitosa. Intentelo de nuevo."
   Exit Sub
End If
    Kill ImgTemporal
    MousePointer = vbNormal
    Image1.Picture = LoadPicture(explorafoto.FileName)
    kb = KbImagen
    
    Exit Sub
error:
    MousePointer = vbNormal
    MsgBox err.Description
End Sub



Public Sub ConexionImg(OPERACION)
'On Error GoTo error



Conex = "driver={MySQL ODBC 3.51 Driver};server=" & Servidor & ";uid=" & _
Usuario & ";pwd=" & password & ";database=" & clientesistema & "conta;connection=adUseClient"

mystream.Type = adTypeBinary
On Local Error Resume Next
cnn.Close
conn.ConnectionString = Conex
conn.CursorLocation = adUseClient
conn.Open
If OPERACION = 1 Then
    
    Call ELIMINARFoto(dato(0).text)

    Rs.Open "Select * From " & clientesistema & "conta.af_maestro_activos_foto limit 0,1", conn, adOpenStatic, adLockOptimistic
    Rs.AddNew
    mystream.Open
    mystream.LoadFromFile explorafoto.FileName
  
        Rs("codigo") = dato(0)
        Rs("linea") = 1
        Rs("foto") = mystream.Read
        Rs.Update
 
    mystream.Close
MsgBox "Se ha agregado la imagen satisfactoriamente", vbInformation, "Agregada"
'Unload Me

End If
If OPERACION = 2 Then
If ExisteArchivo(Imgtemporal2) = True Then
Call Kill(Imgtemporal2)
End If
Rs.Open "Select * From " & clientesistema & "conta.af_maestro_activos_foto WHERE codigo = '" & dato(0).text & "' limit 0,1", conn, adOpenStatic, adLockOptimistic   '
'On Local Error Resume Next
mystream.Type = adTypeBinary
mystream.Open
mystream.Write Rs.Fields("foto")
mystream.SaveToFile Imgtemporal2, adSaveCreateOverWrite
mystream.Close
Image1.Picture = LoadPicture(Imgtemporal2)
'Shell ("rundll32.exe C:\WINDOWS\System32\shimgvw.dll,ImageView_Fullscreen " & Imgtemporal2), vbMaximizedFocus
'KbImagen = ""
'KbImagen = Mid(Str(FileLen(Imgtemporal2)), 1, Len(Str(FileLen(Imgtemporal2))) - 3)
'kb = KbImagen & " KB"
Me.Image1.Picture = LoadPicture(Imgtemporal2)
 cnn.Close
End If
Rs.Close
cnn.Close
Exit Sub
error:
MsgBox err.Description
Exit Sub
End Sub
 


Public Function TieneFoto(codigo) As Boolean
Dim campos(20, 20) As String
Dim op As Double
Dim condicion As String
    campos(0, 0) = "codigo"
    campos(1, 0) = ""
    
    TieneFoto = False
    campos(0, 2) = clientesistema & "conta.af_maestro_activos_foto"
    condicion = "codigo ='" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        TieneFoto = True
    End If
    
End Function


Sub ELIMINARFoto(codigo)
    campos(0, 2) = "af_maestro_activos_foto"
    condicion = "codigo = '" & codigo & "' "

    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
End Sub



Public Function ExisteActivo(codigo) As Boolean
Dim campos(20, 20) As String
Dim op As Double
Dim condicion As String
    campos(0, 0) = "codigobarras"
    campos(1, 0) = ""
    
    ExisteActivo = False
    campos(0, 2) = clientesistema & "conta.af_maestro_activos"
    condicion = "codigobarras ='" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        ExisteActivo = True
    End If
    
End Function


Sub grabar()
Dim n As Double
Dim PASO As Boolean
'If lblNombreEmpresa.Caption = "" Then Exit Sub
'If lblnombreproveedor.Caption = "" Then Exit Sub
'If lblnombretipo.Caption = "" Then Exit Sub
PASO = True
For n = 0 To dato.Count - 1
    If dato(n).text = "" Then
    dato(n).SetFocus
    PASO = False
    Exit For
    End If
Next n
If PASO = True Then

Else
    Exit Sub
End If

Call EliminarActivo(dato(0))

    campos(0, 0) = "codigobarras"
    campos(1, 0) = "numeroserie"
    campos(2, 0) = "descripcion"
    campos(3, 0) = "tipo"
    campos(4, 0) = "empresacontable"
    campos(5, 0) = "factura"
    campos(6, 0) = "proveedor"
    campos(7, 0) = "pcosto"
    campos(8, 0) = "empresa"
    campos(9, 0) = "fechacreacion"
    campos(10, 0) = "usuariocreacion"
    campos(11, 0) = "fecha_compra"
    
 
    campos(12, 0) = ""
    
    campos(0, 1) = dato(0).text
    campos(1, 1) = dato(1).text
    campos(2, 1) = dato(2).text
    campos(3, 1) = dato(3).text
    campos(4, 1) = dato(7).text
    campos(5, 1) = dato(5).text
    campos(6, 1) = dato(4).text & DV.Caption
    campos(7, 1) = Replace(dato(6).text, ".", ",")
    campos(8, 1) = empresaactiva
    campos(9, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(10, 1) = USUARIOSISTEMA
    campos(11, 1) = Format(fechacompra, "yyyy-mm-dd")
    
  
    campos(0, 2) = "af_maestro_activos"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)

Call dato_KeyPress(0, 13)
End Sub


Sub EliminarActivo(codigo)
    campos(0, 2) = "af_maestro_activos"
    condicion = "codigobarras = '" & codigo & "' "

    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
End Sub



Sub DardeBaja()
txtGlosa.text = "DIO DE BAJA EL ACTIVO " & Now
txtGlosa.Tag = 99
Call grabarEvento
    campos(0, 0) = "fechabaja"
    campos(1, 0) = "usuariobaja"
    campos(2, 0) = ""
    campos(3, 0) = ""
   
    campos(0, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(1, 1) = USUARIOSISTEMA
  
    campos(0, 2) = "af_maestro_activos"
    condicion = "codigobarras = '" & dato(0).text & "' "

    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    Call LeerActivo(dato(0).text)
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call grabarEvento
      Call CargarEventos(dato(0).text)
      txtGlosa.SetFocus
End If
End Sub



Sub CargarEventos(codigo)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    LINEA = 1
    Set csql.ActiveConnection = conta
    csql.sql = "SELECT usuario,fecha,glosa,item  "

    csql.sql = csql.sql & " FROM " & clientesistema & "conta.af_maestro_activos_observaciones "
    csql.sql = csql.sql & " where CODIGO = '" & codigo & "' ORDER BY fecha,item "
    csql.Execute
    Grid2.Rows = 1
    Grid2.AutoRedraw = False
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    While resultados.EOF = False
        Grid2.AddItem "", True
        Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0)
        Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
        Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
        Grid2.Cell(Grid2.Rows - 1, 1).Tag = resultados(3)
        resultados.MoveNext
    Wend
    
    End If
   Grid2.AutoRedraw = True
   Grid2.Refresh
   

End Sub



Sub grabarEvento()
Call EliminarEvento(dato(0).text, txtGlosa.Tag)
    campos(0, 0) = "codigo"
    campos(1, 0) = "fecha"
    campos(2, 0) = "usuario"
    campos(3, 0) = "item"
    campos(4, 0) = "glosa"
    campos(5, 0) = ""
    
    campos(0, 1) = dato(0).text
    campos(1, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(2, 1) = USUARIOSISTEMA
    campos(3, 1) = txtGlosa.Tag
    campos(4, 1) = Replace(txtGlosa, "'", "")
 
  
    campos(0, 2) = "af_maestro_activos_observaciones"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
  
    txtGlosa.Tag = LeerUltimoItem(dato(0).text)
    txtGlosa.text = Empty
    
End Sub


Public Function LeerUltimoItem(codigo) As String
    campos(0, 0) = "ifnull(max(item),0)+1"
    campos(1, 0) = ""
    campos(0, 2) = "af_maestro_activos_observaciones"
    condicion = "codigo = '" & codigo & "'  "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        LeerUltimoItem = sqlconta.response(0, 3)
    Else
        LeerUltimoItem = ""
    End If
End Function



Sub EliminarEvento(codigo, item)
    campos(0, 2) = "af_maestro_activos_observaciones"
    condicion = "codigo = '" & codigo & "' and item='" & item & "' "

    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
End Sub
