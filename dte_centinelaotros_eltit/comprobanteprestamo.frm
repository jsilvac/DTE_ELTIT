VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Prestamouf 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " COMPROBANTE DE PRESTAMO"
   ClientHeight    =   9210
   ClientLeft      =   300
   ClientTop       =   315
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   3255
      Left            =   120
      TabIndex        =   41
      Top             =   5760
      Width           =   13700
      _ExtentX        =   24156
      _ExtentY        =   5741
      BackColor       =   16773879
      Caption         =   "LISTADO DE DESCUENTOS"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid GridIngresoHD 
         Height          =   2850
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   5027
         BackColor1      =   15859171
         BackColor2      =   15859171
         BackColorActiveCellSel=   16777088
         BackColorBkg    =   15859171
         BackColorFixedSel=   16761024
         BackColorScrollBar=   16744576
         BackColorSel    =   16777215
         BorderColor     =   16744576
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         SelectionBorderColor=   16744576
         DefaultFontName =   "Arial"
         DefaultFontSize =   9.75
         DefaultFontBold =   -1  'True
         DisplayRowIndex =   -1  'True
         ForeColorFixed  =   8388608
         GridColor       =   16744576
         Rows            =   4
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp FrameXP1 
      Height          =   5535
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   9763
      BackColor       =   16773879
      Caption         =   " ANTECEDENTES"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtempresa 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "rut"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "nombres"
         Top             =   2880
         Width           =   1260
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Tag             =   "nombres"
         Top             =   2880
         Width           =   900
      End
      Begin VB.TextBox valoruf 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "rut"
         Top             =   720
         Width           =   1575
      End
      Begin FlexCell.Grid Grid4 
         Height          =   135
         Left            =   5760
         TabIndex        =   34
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   238
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.TextBox numerotxt 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "rut"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Comboaño 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3240
         Width           =   1335
      End
      Begin VB.ComboBox Combomes 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "nombres"
         Top             =   2880
         Width           =   420
      End
      Begin VB.CommandButton btnGrabar 
         Caption         =   "Grabar"
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
         Left            =   2280
         TabIndex        =   11
         Top             =   3840
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox dato1 
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
         Height          =   285
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   3
         Tag             =   "rut"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox dato4 
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
         Height          =   285
         Left            =   6180
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "rol"
         Top             =   1065
         Width           =   675
      End
      Begin VB.TextBox dato3 
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
         Height          =   285
         Left            =   5700
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "rol"
         Top             =   1065
         Width           =   375
      End
      Begin VB.TextBox dato2 
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
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "rol"
         Top             =   1065
         Width           =   375
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "nombres"
         Top             =   2520
         Width           =   1740
      End
      Begin VB.TextBox pivote1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9480
         TabIndex        =   14
         Top             =   -240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblsueldobase 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1920
         TabIndex        =   40
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblnombreempresa 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1680
         TabIndex        =   39
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMPRESA"
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
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label lblvalor 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor Cuota U.F"
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
         Left            =   3720
         TabIndex        =   37
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lbltaza 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Taza"
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
         Left            =   1920
         TabIndex        =   36
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VALOR U.F"
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
         Left            =   4080
         TabIndex        =   35
         Top             =   720
         Width           =   1170
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
         Height          =   1815
         Left            =   120
         TabIndex        =   33
         Top             =   3840
         Visible         =   0   'False
         Width           =   7095
         _cx             =   12515
         _cy             =   3201
         FlashVars       =   ""
         Movie           =   "c:\barra_opciones.swf"
         Src             =   "c:\barra_opciones.swf"
         WMode           =   "Transparent"
         Play            =   "0"
         Loop            =   "-1"
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   "-1"
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   "0"
         EmbedMovie      =   "0"
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   "1"
         Profile         =   "0"
         ProfileAddress  =   ""
         ProfilePort     =   "0"
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO"
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
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Año A Partir"
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
         Left            =   4380
         TabIndex        =   29
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cuotas"
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
         Left            =   120
         TabIndex        =   28
         Top             =   2880
         Width           =   1000
      End
      Begin VB.Label lblnombre 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1200
         TabIndex        =   26
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Mes A Partir"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFF2F7&
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   22
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " R.U.T."
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
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1000
      End
      Begin VB.Label dv 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2400
         TabIndex        =   18
         Top             =   1080
         Width           =   285
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
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
         Left            =   4080
         TabIndex        =   17
         Top             =   1065
         Width           =   1170
      End
      Begin VB.Label lblDato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1000
      End
      Begin VB.Label lblfechacontrato 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1920
         TabIndex        =   27
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha de Contrato"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sueldo Base"
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
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1695
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   0
      Width           =   495
   End
   Begin XPFrame.FrameXp FrameXP1 
      Height          =   5535
      Index           =   1
      Left            =   7320
      TabIndex        =   24
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9763
      BackColor       =   16773879
      Caption         =   " LISTA DE PRESTAMOS VIGENTES"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Gridprestamos 
         Height          =   5055
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   8916
         BackColor1      =   16761024
         BackColor2      =   16761024
         BackColorActiveCellSel=   16777088
         BackColorBkg    =   16761024
         BackColorFixedSel=   16761024
         BackColorScrollBar=   16744576
         BorderColor     =   16744576
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         SelectionBorderColor=   16744576
         Cols            =   3
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         ForeColorFixed  =   8388608
         GridColor       =   16744576
         Rows            =   4
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
End
Attribute VB_Name = "prestamouf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim codtemporal As String
Dim ruttemporal As String
Dim fechatemporal1 As String
Dim tiempo1 As Double
Dim tiempo2 As Double
Dim cabezas1 As Variant
Dim fila As Integer
Dim columna As Integer
Dim i As Integer
Dim domicilio As String
Dim ciudad As String
Dim rutempresa As String




Private Sub btnGrabar_Click()
    If MsgBox("DESEA GENERAR LAS CUOTAS", vbYesNo, "ATENCION") = vbYes Then
        Call grabarprestamo(dato1.text & dv.Caption, numerotxt.text, dato4.text & "-" & dato3.text & "-" & dato2.text, lblsueldobase.Caption, dato6.text, dato7.text, Mid(Combomes.text, 1, 2), Comboaño.text, "", "", txtempresa.text, dato9.text, valoruf.text, dato8.text, "0")
        'Call imprimir
        Call retorno
    End If
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Call AyudaTrabajador(txtempresa.text)
    End If
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato1.text <> "" Then
        Call ceros(dato1)
        dv.Caption = rut(dato1.text)
        If leertrabajador(dato1.text & dv.Caption, txtempresa.text) = True Then
            Call CargaGrillaHD(GridIngresoHD, txtempresa.text, dato3.text, dato4.text)
            dato6.Locked = False
            dato6.SetFocus
        Else
            MsgBox "RUT NO INGRESADO EN NUESTRA BASE DE DATOS ", vbExclamation, "ATENCION"
        End If
    End If
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato6.text <> "" Then
    If CDbl(dato6.text) > 0 Then
        dato6.text = Format(dato6.text, "###,###,##0")
        dato7.Locked = False
        dato7.SetFocus
    End If
End If
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato7.text <> "" Then
    If CDbl(dato7.text) < 25 Then
      dato8.SetFocus
    Else
        MsgBox "SOLO PERMITE 24 CUOTAS", vbCritical, "ATENCION"
        dato7.text = ""
        dato7.SetFocus
    End If
End If
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
KeyAscii = esNumeroDecimal(dato8, KeyAscii)
If KeyAscii = 13 And dato8.text <> "" Then
    dato9.text = CALCULACUOTA(dato6.text, dato7.text, dato8.text, valoruf.text)
    btnGrabar.Visible = True
End If
End Sub

Private Sub Form_Activate()
sqlventas.audit = True
sqlventas.programaactivo = Me.Caption
End Sub
'******************************************************************
Private Sub Form_Load()
    Dim K As Double
'    Me.width = 13275
'    Me.height = 5370
'    Call Funciones.Centrar(Me)
'    Call Configuracion.Conectar_BD
'    Call Configuracion.ConectarRemu(servidor, clientesistema + "remu", usuario, password) 'remu
    For K = 1 To 12
        Combomes.AddItem Format(K, "00") & " - " & MonthName(K)
    Next K
    Combomes.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
    For K = Val(Format(fechasistema, "yyyy")) To Val(Format(fechasistema, "yyyy")) + 3
        Comboaño.AddItem K
    Next K
    Comboaño.ListIndex = 0
    dato2.text = Format(fechasistema, "dd")
    dato3.text = Format(fechasistema, "mm")
    dato4.text = Format(fechasistema, "yyyy")
    Call CARGAGRILLAprestamo(Gridprestamos)
    Call IniciaGrid1
End Sub

Private Sub Gridprestamos_DblClick()
    If Gridprestamos.Rows > 1 Then
        Call existeprestamo(Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 1).text, txtempresa.text)
    End If
End Sub

Private Sub Label9_Click()

End Sub

'************************************************************************
'************************************************************************
Private Sub MANUAL_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 27 And dato1.text <> "") Then
        dato1.SetFocus
    ElseIf (KeyAscii = 27 And dato1.text = "") Then
        Unload Me
    End If
End Sub
  

 '********************************************************************************
'********************************************************************************
Sub AyudaTrabajador(Empresaconsulta)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("10s", "25s")
    cfijo = " rut <> '00'"
    cabezas = Array("RUT", "APELLIDOS", "NOMBRES")
    mensajeAyuda = "AYUDA TRABAJADOR"
    Call Mayuda.cargaAyudaT(servidor, clientesistema + "remu" & Empresaconsulta, usuario, password, "mt_fijo", pivote1, campos, cfijo, largo, 2)
    ruttemporal = pivote1.text
    If (ruttemporal <> "") Then
        dato1.text = Left(ruttemporal, 9) 'rut
        dv.Caption = Right(ruttemporal, 1)  'dv
    End If
    pivote1.text = ""
End Sub
 
Function leertrabajador(rut, Empresaconsulta) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventas
    
    csql.sql = "select nombre,fechaing,direccion,comuna "
    csql.sql = csql.sql & "from " & clientesistema & "remu" & Empresaconsulta & ".mt_fijo "
    csql.sql = csql.sql & "where rut='" & rut & "' "
    csql.Execute
    leertrabajador = False
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        lblnombre.Caption = resultados(0)
        lblfechacontrato.Caption = resultados(1)
        domicilio = resultados(2)
        ciudad = resultados(3)
        leertrabajador = True
        Call leerprestamos(rut, txtempresa.text)
        lblsueldobase.Caption = leersueldobase(rut, txtempresa.text)
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
End Function
 

Private Sub numerotxt_GotFocus()
Call cargatexto(numerotxt)
End Sub

Private Sub numerotxt_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And numerotxt.text <> "" Then
        If CDbl(numerotxt.text) > 0 Then
            numerotxt.text = ceros(numerotxt)
            If existeprestamo(numerotxt, txtempresa.text) = False Then
                valoruf.Locked = False
                valoruf.SetFocus
            End If
        End If
    End If

End Sub
Function leerultimofolioprestamo(CODIGO) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventas
    csql.sql = " select IFNULL(MAX(comprobante) + 1,'0000000001')  "
    csql.sql = csql.sql & "from  sv_prestamo where empresa='" & CODIGO & "' limit 0,1 "
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerultimofolioprestamo = Format(resultados(0), "0000000000")
    Else
        leerultimofolioprestamo = "0000000001"
    End If
    
End Function

Sub leerprestamos(rut, Empresaconsulta)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventas
    
    csql.sql = "select pc.empresa,pc.comprobante,pc.numerocuota,p.cantidadcuotas,pc.mes,pc.año,pc.monto "
    csql.sql = csql.sql & "from sv_prestamo_cuota as pc inner join sv_prestamo as p "
    csql.sql = csql.sql & "on pc.comprobante=p.comprobante and pc.empresa=p.empresa and pc.rut=p.rut "
    csql.sql = csql.sql & "where pc.rut='" & rut & "' and pc.empresa='" & Empresaconsulta & "' and pc.mesrebajado='' "
    csql.sql = csql.sql & "order by pc.comprobante,pc.numerocuota asc "
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        Gridprestamos.Rows = 1
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Gridprestamos.Rows = Gridprestamos.Rows + 1
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 1).text = resultados(0)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 2).text = resultados(1)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 3).text = resultados(2) & " / " & resultados(3)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 4).text = resultados(4) & "-" & resultados(5)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 5).text = resultados(6)
            resultados.MoveNext
        Wend
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
End Sub
Sub CARGAGRILLAprestamo(Grid1 As Grid)
    Dim formatogrilla(10, 10) As String
    Dim K As Double
    
    Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 8
    
    formatogrilla(1, 1) = "EMPRESA"
    formatogrilla(1, 2) = "NUMERO"
    formatogrilla(1, 3) = "Nº CUOTA"
    formatogrilla(1, 4) = "FECHA"
    formatogrilla(1, 5) = "MONTO EN U.F"
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "8"
    formatogrilla(2, 2) = "10"
    formatogrilla(2, 3) = "8"
    formatogrilla(2, 4) = "8"
    formatogrilla(2, 5) = "10"
 
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "N"
    formatogrilla(3, 2) = "N"
    formatogrilla(3, 3) = "N"
    formatogrilla(3, 4) = "D"
    formatogrilla(3, 5) = "N"
 
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = "0000000000"
    formatogrilla(4, 3) = ""
    formatogrilla(4, 4) = ""
    formatogrilla(4, 5) = ""
 
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    Grid1.Cols = 6
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
    Grid1.BackColorFixedSel = RGB(110, 180, 230)
    Grid1.BackColorBkg = RGB(90, 158, 214)
    Grid1.BackColorScrollBar = RGB(231, 235, 247)
    Grid1.BackColor1 = RGB(231, 235, 247)
    Grid1.BackColor2 = RGB(239, 243, 255)
    Grid1.GridColor = RGB(148, 190, 231)
    For K = 1 To Grid1.Cols - 1
        Grid1.Cell(0, K).text = formatogrilla(1, K)
        Grid1.Column(K).Width = CDbl(formatogrilla(2, K)) * Grid1.DefaultFont.Size
        
        Grid1.Column(K).MaxLength = CDbl(formatogrilla(2, K))
        Grid1.Column(K).FormatString = formatogrilla(4, K)
        Grid1.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then Grid1.Column(K).Alignment = cellRightCenter
       
    Next K
    Grid1.Column(0).Width = 0
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
    Rem Grid1.Enabled = False
End Sub

Private Function existeprestamo(NUMERO, Empresaconsulta) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = ventas
    csql.sql = "select rut,fecha,monto,cantidadcuotas,mesapartir,añoapartir,sueldobase,comprobante,valoruf,taza "
    csql.sql = csql.sql & "from sv_prestamo "
    csql.sql = csql.sql & "where comprobante='" & NUMERO & "' and empresa='" & Empresaconsulta & "' "
    csql.Execute
    existeprestamo = False
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        dato1.text = Mid(resultados(0), 1, 9)
        dv.Caption = Mid(resultados(0), 10, 1)
        dato2.text = Format(resultados(1), "dd")
        dato3.text = Format(resultados(1), "mm")
        dato4.text = Format(resultados(1), "yyyy")
        lblsueldobase.Caption = Format(resultados(6), "###,###,###")
        dato6.text = Format(resultados(2), "###,###,###")
        dato7.text = resultados(3)
        Combomes.ListIndex = resultados(4) - 1
        Comboaño.text = resultados(5)
        numerotxt.text = resultados(7)
        valoruf.text = Format(resultados(8), "###,###,###")
        dato8.text = resultados(9)
        dato9.text = CALCULACUOTA(dato6.text, dato7.text, dato8.text, valoruf.text)
        Call dato1_KeyPress(13)
        
        existeprestamo = True
        dato6.Locked = True
        opciones.Visible = True
        opciones.SetFocus
    End If
    
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
End Function

 Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
    If command = "retorno" Then Call retorno
    If command = "modifica" Then
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
    End If
    If command = "elimina" Then
       If MsgBox("REALMENTE DESEA ELIMINAR", vbYesNo) = vbYes Then
        If Verifica_Permiso(Me.Caption, "elimina") = True Then
            Call ELIMINAR(numerotxt.text, dato1.text & dv.Caption, dato4.text & "-" & dato3.text & "-" & dato2.text, txtempresa.text)
            Call retorno
         Else
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
         End If
        End If

    End If
 
    If command = "imprime" Then Call imprimir
End Sub

Sub ELIMINAR(NUMERO, rut, fecha, Empresaconsulta)
   Dim campos(30, 30) As String
    Dim op As Integer
    
   
    campos(0, 2) = "sv_prestamo" 'tabla
    condicion = "empresa='" & Empresaconsulta & "' and comprobante='" & NUMERO & "' and rut='" & rut & "' and fecha='" & fecha & "'"
    op = 4
    
    sqlventas.response = campos
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    
    campos(0, 2) = "sv_prestamo_cuota" 'tabla
    condicion = "empresa='" & Empresaconsulta & "' and rut='" & rut & "' and comprobante='" & NUMERO & "' "
    op = 4
    
    sqlventas.response = campos
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    
    
End Sub
Sub retorno()
    Gridprestamos.Rows = 1
    Combomes.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
    Comboaño.ListIndex = 0
    dato7.Locked = True
    dato7.text = ""
    dato6.Locked = True
    dato6.text = ""
    lblsueldobase.Caption = ""
    dato4.text = Format(fechasistema, "yyyy")
    dato3.text = Format(fechasistema, "mm")
    dato2.text = Format(fechasistema, "dd")
    dato1.Locked = True
    dato1.text = ""
    dv.Caption = ""
    lblfechacontrato.Caption = ""
    lblnombre.Caption = ""
    numerotxt.text = ""
    valoruf.text = ""
    dato8.text = ""
    dato9.text = ""
    txtempresa.text = ""
    lblnombreempresa.Caption = ""
    opciones.Visible = False
    btnGrabar.Visible = False
    txtempresa.SetFocus
    
End Sub
Sub imprimir()
' contratoprestamo

Dim Word As Word.Application
Dim documento As Word.Documents

Dim MiDoc As String

If dato6.text <> "" Then
        
        Set Word = CreateObject("Word.Application")
        Word.Documents.Open "Z:\RESPALDO\promotora\contratoprestamo.doc"
        Word.Visible = False
        Word.Application.WindowState = wdWindowStateMinimize
        Word.Application.Documents.Open "Z:\RESPALDO\promotora\contratoprestamo.doc"
        Word.Documents(1).Bookmarks("folio").Range = numerotxt.text
        Word.Documents(1).Bookmarks("nombretra").Range = lblnombre.Caption
        Word.Documents(1).Bookmarks("ruttra").Range = Format(dato1.text, "###,###,###") & "-" & dv.Caption
        Word.Documents(1).Bookmarks("domiciliotra").Range = domicilio
        Word.Documents(1).Bookmarks("ciudadtra").Range = ciudad
        Word.Documents(1).Bookmarks("monto").Range = Format(CDbl(dato6.text), "$ ###,###,###")
        Word.Documents(1).Bookmarks("montopalabras").Range = WORDNUM(Replace(dato6.text, ".", ""), "PESO", "PESOS", "", "")
        Word.Documents(1).Bookmarks("nombreempresa").Range = lblnombreempresa.Caption
        Word.Documents(1).Bookmarks("rutempresa").Range = rutempresa
        Word.Documents(1).Bookmarks("cuotas").Range = dato7.text
        Word.Documents(1).Bookmarks("ufs").Range = dato9.text
        Word.Documents(1).Bookmarks("mesinicio").Range = MonthName(Combomes.ListIndex + 1)
        Word.Documents(1).Bookmarks("anoinicio").Range = Comboaño.text
        Word.Documents(1).Bookmarks("nombretra2").Range = lblnombre.Caption
        Word.Documents(1).Bookmarks("ruttra2").Range = Format(dato1.text, "###,###,###") & "-" & dv.Caption
        'ariel
        Word.Documents(1).Bookmarks("fechafirma").Range = Format(fechasistema, "dd-mm-YYYY")
        Word.ActiveDocument.SaveAs "Z:\RESPALDO\promotora\" & dato1.text & "-Prestamo.doc"
        Word.ActiveDocument.Close savechanges:=wdDoNotSaveChanges
        Word.Application.Documents.Open "Z:\RESPALDO\promotora\" & dato1.text & "-Prestamo.doc"
        
        Exit Sub
controlerror:
  MsgBox "DEBE TENER INSTALADO MICROSOFT OFFICE EN SU PC O " & vbCrLf & "DOCUMENTOS NO SE ENCUENTRAN EN " & rutaUpdate, vbCritical, "ATENCION"
Else
    MsgBox "DEBE INGRESAR UN CLIENTE ", vbCritical, "ATENCION"
End If
     
End Sub
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub
Sub cargagrilla4()
Dim formatogrilla(8, 8) As String
Dim K As Integer
    Rem DATOS DE LA COLUMNA
    Grid4.DefaultFont.Size = 8
    Grid4.DefaultFont.Bold = False
    
    formatogrilla(1, 1) = ""
    formatogrilla(1, 2) = ""
    formatogrilla(1, 3) = ""
    formatogrilla(1, 4) = ""
    formatogrilla(1, 5) = ""
    formatogrilla(1, 6) = ""
    formatogrilla(1, 7) = ""
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "13"
    formatogrilla(2, 2) = "30"
    formatogrilla(2, 3) = "8"
    formatogrilla(2, 4) = "8"
    formatogrilla(2, 5) = "8"
    formatogrilla(2, 6) = "8"
    formatogrilla(2, 7) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "N"
    formatogrilla(3, 4) = "N"
    formatogrilla(3, 5) = "N"
    formatogrilla(3, 6) = "N"
    formatogrilla(3, 7) = "N"

    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = "###,##0.0"
    formatogrilla(4, 4) = "###,##0.0"
    formatogrilla(4, 5) = "###,##0.0"
    formatogrilla(4, 6) = "###,##0.0"
    formatogrilla(4, 7) = "###,##0.0"
    
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "FALSE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    formatogrilla(5, 6) = "FALSE"
    formatogrilla(5, 7) = "TRUE"
    
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
    For K = 1 To Grid4.Cols - 1
        Grid4.Cell(0, K).text = formatogrilla(1, K)
        'Grid4.Cell(1, k).text = FORMATOGRILLA(8, k)
        Grid4.Column(K).Width = Val(formatogrilla(2, K)) * Grid4.Cell(0, K).Font.Size
        Grid4.Column(K).MaxLength = Val(formatogrilla(2, K))
        Grid4.Column(K).FormatString = formatogrilla(4, K)
        Grid4.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then Grid4.Column(K).Alignment = cellRightCenter
    Next K
    Grid4.Column(0).Width = 0
    Grid4.Range(0, 0, 0, Grid4.Cols - 1).Alignment = cellCenterCenter
    Grid4.Column(3).UserSortIndicator = cellSortIndicatorDescending
    Rem Grid4.Enabled = False
End Sub

Sub cabeza()
Dim K As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    Dim año As String
    Dim mes As Double
    Grid4.ReportTitles.Clear
    'Report Title 1
   
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = nombreempresa
        objReportTitle.Font.Name = "verdana"
        objReportTitle.Font.Size = 7
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        objReportTitle.color = RGB(128, 0, 0)
        objReportTitle.Align = cellLeft
        Grid4.ReportTitles.Add objReportTitle
 
     
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "COMPROBANTE DE PRESTAMO "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 11
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " FECHA : " & dato2.text & "-" & dato3.text & "-" & dato4.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellRight
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle

    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Nombre del Trabajador :" & lblnombre.Caption
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "R.u.t del Trabajador  :" & Format(dato1.text, "###,###,###") & "-" & dv.Caption
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Fecha de Contrato  " & lblfechacontrato.Caption
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Sueldo Base :" & Format(lblsueldobase.Caption, "$ ###,###,###")
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " DATOS DEL PRESTAMO "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Monto Prestado    :" & Format(dato6.text, "$ ###,###,###")
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "     " & WORDNUM(Format(dato6.text, "########0"), "PESO", "PESOS", 0)
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Cantidad de Cuotas  : " & dato7.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Descontados de su liquidación o finiquito según corresponda de la siguiente forma : "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "REMUNERACION                               CUOTA "
    objReportTitle.Font.Name = "Courier"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Font.Underline = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    
 
    
        mes = Combomes.ListIndex + 1
        año = Comboaño.text
    For K = 0 To dato7.text - 1
        Set objReportTitle = New FlexCell.ReportTitle
        
        If mes > 12 Then
        año = CDbl(Comboaño.text) + 1
        mes = 1
        End If
        objReportTitle.text = MonthName(mes) & String(13 - Len(MonthName(mes)), " ") & " de " & año & "                      " & Format(dato9.text, "###,###,###0.000")
        objReportTitle.Font.Name = "Courier"
        objReportTitle.Font.Size = 10
        objReportTitle.Font.Bold = True
        objReportTitle.PrintOnAllPages = True
        objReportTitle.Align = cellCenter
        Grid4.ReportTitles.Add objReportTitle
        mes = mes + 1
    Next K
    
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
 
 Sub grabarprestamo(rut, comprobante, fecha, sueldobase, MONTO, cantidadcuotas, mesapartir, añoapartir, TIPO, GLOSA, Empresaconsulta, montocuota, valorufcuota, tazacuota, numeroprestamo)
    Dim campos(30, 30) As String
    Dim op As Integer
    
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim fechacuota As Date
    Dim codigotablacalculo As String
    Dim codcalculo As String
    Dim K As Double
    campos(0, 0) = "rut"
    campos(1, 0) = "fecha"
    campos(2, 0) = "sueldobase"
    campos(3, 0) = "monto"
    campos(4, 0) = "cantidadcuotas"
    campos(5, 0) = "mesapartir"
    campos(6, 0) = "añoapartir"
    campos(7, 0) = "comprobante"
    campos(8, 0) = "empresa"
    campos(9, 0) = "valoruf"
    campos(10, 0) = "taza"
    campos(11, 0) = "numeroprestamo"
    campos(12, 0) = ""
    campos(0, 1) = rut
    campos(1, 1) = fecha
    campos(2, 1) = CDbl(sueldobase)
    campos(3, 1) = CDbl(MONTO)
    campos(4, 1) = cantidadcuotas
    campos(5, 1) = mesapartir
    campos(6, 1) = añoapartir
    campos(7, 1) = comprobante
    campos(8, 1) = Empresaconsulta
    campos(9, 1) = CDbl(valorufcuota)
    campos(10, 1) = Replace(tazacuota, ",", ".")
    campos(11, 1) = numerodeprestamo(rut)
    
    
    campos(0, 2) = "sv_prestamo"  'tabla
    condicion = ""
    op = 2
    
    sqlventas.response = campos
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
 
    fechacuota = "01-" & mesapartir & "-" & añoapartir
    fechacuota = DateAdd("m", -1, fechacuota)
 
    campos(0, 0) = "rut"
    campos(1, 0) = "comprobante"
    campos(2, 0) = "numerocuota"
    campos(3, 0) = "mes"
    campos(4, 0) = "año"
    campos(5, 0) = "monto"
    campos(6, 0) = "empresa"
    campos(7, 0) = ""
    
    For K = 1 To CDbl(cantidadcuotas)
        campos(0, 1) = rut
        campos(1, 1) = comprobante
        campos(2, 1) = K
        campos(3, 1) = Format(DateAdd("m", K, fechacuota), "mm")
        campos(4, 1) = Format(DateAdd("m", K, fechacuota), "yyyy")
        campos(5, 1) = Replace(montocuota, ",", ".")
        campos(6, 1) = Empresaconsulta
    
        campos(0, 2) = "sv_prestamo_cuota"  'tabla
        condicion = ""
        op = 2
     
        sqlventas.response = campos
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
     Next K
     
    
End Sub
Private Sub txtempresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Call ayudaEmpresaremu(txtempresa)
    End If
End Sub
Private Sub txtempresa_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And txtempresa.text <> "" Then
        txtempresa.text = ceros(txtempresa)
        lblnombreempresa.Caption = leernombreempresaremu(txtempresa.text)
        
        rutempresa = leerrutempresaremu(txtempresa.text)
        If lblnombreempresa.Caption <> "" Then
            numerotxt.text = leerultimofolioprestamo(txtempresa.text)
            numerotxt.SetFocus
        Else
            MsgBox "EMPRESA NO EXISTE", vbCritical, "ATENCION"
            txtempresa.text = ""
            txtempresa.SetFocus
        End If
    End If
End Sub
Function CALCULACUOTA(MONTO, CUOTAS, taza, valoruf) As String
Dim montofinal As Double
Dim cuotasfinal As Double
Dim tazafinal As Double

tazafinal = (CUOTAS * taza) / 100
montofinal = (MONTO * (1 + tazafinal))
montofinal = montofinal / valoruf
cuotasfinal = montofinal / CUOTAS
CALCULACUOTA = Round(cuotasfinal, 3)

End Function

Private Sub valoruf_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If valoruf.text <> "" And KeyAscii = 13 Then
        If CDbl(valoruf.text) > 0 Then
            valoruf.text = Format(valoruf.text, "###,###,##0")
            dato1.SetFocus
        Else
            MsgBox "DEBE INGRESAR UN VALOR > 0 ", vbCritical, "ATENCION"
            valoruf.text = ""
            valoruf.SetFocus
        End If
    End If
End Sub
Function leersueldobase(rut, Empresaconsulta) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = ventas
    csql.sql = "select monto "
    csql.sql = csql.sql & "from " & clientesistema & "remu" & Empresaconsulta & ".liquidacionhd "
    csql.sql = csql.sql & "where rut='" & rut & "' and año='" & Format(fechasistema, "yyyy") & "' and mes='" & Format(Format(fechasistema, "mm") - 1, "00") & "' and um='SB' "
    csql.Execute
    leersueldobase = "0"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leersueldobase = Format(resultados(0), "###,###,##0")
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
End Function

Sub IniciaGrid1()
    cabezas1 = Array("CODIGO", "DESCRIPCION", "U/M", "MONTO", "CR.CC", "DURAC.", "")
    Call CargaGrilla1(1, 7, GridIngresoHD, cabezas1)
End Sub
Sub CargaGrilla1(numRow, numCol, grilla As Grid, camposgrid As Variant)
    Dim i As Integer
    Dim formatogrilla(30, 30) As String
    Dim K As Double
    
    i = 0
    While (camposgrid(i) <> "")
        formatogrilla(1, i + 1) = camposgrid(i) 'encabezados
        i = i + 1
    Wend
    
    Rem LARGO DE LOS DATOS
    For i = 1 To 6
        formatogrilla(2, i) = "10"
    Next i
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE
    formatogrilla(3, 1) = "C"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "C"
    formatogrilla(3, 4) = "N"
    formatogrilla(3, 5) = "C"
    formatogrilla(3, 6) = "N"
    Rem FORMATO GRILLA
    For i = 1 To 6
        formatogrilla(4, i) = ""
    Next i
'   formatogrilla(4, 4) = "###,###,##0.0"
    Rem LOCCKED
    For i = 1 To 6
        formatogrilla(5, i) = "FALSE"
    Next i
    Rem ancho
    formatogrilla(6, 1) = "7"
    formatogrilla(6, 2) = "30"
    formatogrilla(6, 3) = "4"
    formatogrilla(6, 4) = "8"
    formatogrilla(6, 5) = "7"
    formatogrilla(6, 6) = "6"
    With grilla
        .Cols = numCol
        .Rows = numRow
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .ExtendLastCol = False
        .BoldFixedCell = False
        .DrawMode = cellOwnerDraw
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        .Column(0).Width = 30
        For K = 1 To numCol - 1
            .Cell(0, K).text = formatogrilla(1, K)
            .Column(K).Width = Val(formatogrilla(6, K)) * .Cell(0, K).Font.Size + 1.25
            .Column(0).Width = 15
            .Column(K).MaxLength = Val(formatogrilla(2, K))
            .Column(K).FormatString = formatogrilla(4, K)
            .Column(K).Locked = True 'formatogrilla(5, K)
            If formatogrilla(3, K) = "N" Then
                .Column(K).Alignment = cellRightCenter
                .Column(K).Mask = cellNumeric
            End If
            If formatogrilla(3, K) = "S" Then
                .Column(K).Alignment = cellLeftCenter
                .Column(K).Mask = cellUpper
            End If
            If formatogrilla(3, K) = "C" Then
                .Column(K).Alignment = cellCenterCenter
                .Column(K).Mask = cellUpper
            End If
            If formatogrilla(3, K) = "D" Then
                .Column(K).CellType = cellCalendar
                .Column(K).Mask = cellNumeric
            End If
            '.Column(7).CellType = cellComboBox
        Next K
        '.Range(0, 1, 0, 3).Merge
        '.Cell(0, 1).text = "CUENTA"
        .Range(0, 0, 0, .Cols - 1).Alignment = cellCenterCenter
    End With '//grilla
End Sub
Sub CargaGrillaHD(ByVal grilla As Grid, Empresaconsulta As String, mestemporal, año)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Integer
    
    If Len(mestemporal) < 2 Then
        mestemporal = "0" & mestemporal
    End If
    Set csql.ActiveConnection = ventas
    csql.sql = "SELECT codtablacalculo, glosa, um, monto, codcentrocosto, duracion"
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & Empresaconsulta & ".liquidacionhd "
    csql.sql = csql.sql + " WHERE rut= '" & ruttemporal & "' "
    csql.sql = csql.sql + " AND mes= '" & mestemporal & "' "
    csql.sql = csql.sql + " AND año= '" & año & "' and (um='D$' or um='DU') "
    csql.sql = csql.sql + " ORDER BY codtablacalculo"
    csql.Execute
    grilla.Rows = csql.RowsAffected + 1
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        linea = 1
        While Not resultados.EOF
            grilla.Cell(linea, 1).text = resultados(0) 'cod. tabla calculo
            'grilla.Cell(linea, 2).text = " " & Validaciones.RetornaValor("glosa", "tabladecalculo", "codigo='" & resultados(0) & "'", remu)
            grilla.Cell(linea, 2).text = resultados(1) 'glosa
            grilla.Cell(linea, 3).text = resultados(2) 'um
            If resultados(3) - Int(resultados(3)) = 0 Then
            grilla.Cell(linea, 4).text = Format(resultados(3), "###,###,###") 'monto
            Else
            grilla.Cell(linea, 4).text = Format(resultados(3), "###,###,##0.000") 'monto
            
            End If
            
            grilla.Cell(linea, 5).text = resultados(4) 'crcc
            grilla.Cell(linea, 6).text = resultados(5) 'duracion
            resultados.MoveNext
            linea = linea + 1
        Wend
        resultados.Close
        Set resultados = Nothing
    End If
    grilla.AutoRedraw = True
    grilla.Refresh
    '--------------------------------
   
End Sub


