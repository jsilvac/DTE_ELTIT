VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9e.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prestamosacargar 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " LISTADO PRESTAMOS A CARGAR"
   ClientHeight    =   8355
   ClientLeft      =   180
   ClientTop       =   315
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXP1 
      Height          =   8295
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   14631
      BackColor       =   16744576
      Caption         =   " LISTA DE PRESTAMOS"
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "Acumulado"
         Height          =   255
         Left            =   10200
         TabIndex        =   40
         Top             =   7680
         Width           =   1695
      End
      Begin VB.CommandButton cmdimprimir 
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
         Height          =   375
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   7560
         Width           =   1575
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos"
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
         Left            =   1800
         TabIndex        =   38
         Top             =   7680
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Mes Actual"
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
         Left            =   120
         TabIndex        =   37
         Top             =   7680
         Width           =   1695
      End
      Begin VB.CommandButton cmdgenerar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERAR"
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   7560
         Width           =   1575
      End
      Begin VB.CommandButton cmdenviar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ENVIAR "
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin FlexCell.Grid Gridprestamos 
         Height          =   7095
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   13275
         _ExtentX        =   23416
         _ExtentY        =   12515
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
         DateFormat      =   2
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
      TabIndex        =   17
      Top             =   0
      Width           =   495
   End
   Begin XPFrame.FrameXp FrameXP1 
      Height          =   5535
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   1
         Tag             =   "rut"
         Top             =   720
         Width           =   1575
      End
      Begin FlexCell.Grid Grid4 
         Height          =   135
         Left            =   5760
         TabIndex        =   30
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
         TabIndex        =   0
         Tag             =   "rut"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Comboaño 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3240
         Width           =   1335
      End
      Begin VB.ComboBox Combomes 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   26
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
         TabIndex        =   6
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
         Left            =   5640
         TabIndex        =   9
         Top             =   4920
         Visible         =   0   'False
         Width           =   1155
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   5
         Tag             =   "nombres"
         Top             =   2520
         Width           =   1740
      End
      Begin VB.TextBox pivote1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9480
         TabIndex        =   12
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
         TabIndex        =   34
         Top             =   2160
         Width           =   1935
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   720
         Width           =   1170
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
         Height          =   1215
         Left            =   240
         TabIndex        =   29
         Top             =   4200
         Visible         =   0   'False
         Width           =   5655
         _cx             =   9975
         _cy             =   2143
         FlashVars       =   ""
         Movie           =   "c:\barra_opciones.swf"
         Src             =   "c:\barra_opciones.swf"
         WMode           =   "Transparent"
         Play            =   0   'False
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
         Profile         =   0   'False
         ProfileAddress  =   ""
         ProfilePort     =   0
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
         TabIndex        =   28
         Top             =   720
         Width           =   1050
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
         TabIndex        =   25
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
         Left            =   180
         TabIndex        =   24
         Top             =   2880
         Width           =   975
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
         TabIndex        =   22
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
         Left            =   180
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   2520
         Width           =   135
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
         Left            =   180
         TabIndex        =   16
         Top             =   1440
         Width           =   975
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
         TabIndex        =   15
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
         Left            =   180
         TabIndex        =   13
         Top             =   2520
         Width           =   975
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
         TabIndex        =   23
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
         Left            =   180
         TabIndex        =   21
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
         Left            =   180
         TabIndex        =   14
         Top             =   2160
         Width           =   1695
      End
   End
End
Attribute VB_Name = "prestamosacargar"
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
Dim primera As String

 
Private Sub cmdenviar_Click()
Dim K As Double
Dim cargar As Double

cargar = 0
    If Gridprestamos.Rows > 1 Then
        For K = 1 To Gridprestamos.Rows - 1
            If Gridprestamos.Cell(K, 8).text = "1" Then
                cargar = 1
                Exit For
            End If
        Next K
        If cargar = 1 Then
            Call enviararemuneraciones(Gridprestamos)
            MsgBox "SE HAN ENVIADO A REMUNERACIONES EXITOSAMENTE", vbInformation, "ATENCION"
            Call leerprestamosvigentes(Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), opt1.Value)
        Else
            MsgBox "NO HAY CUOTAS MARCADAS PARA ENVIAR ", vbCritical, "ATENCION"
        End If
      Else
        MsgBox "NO HAY DATOS PARA ENVIAR", vbCritical, "ATENCION"
    End If
End Sub

Private Sub cmdgenerar_Click()
   Call leerprestamosvigentes(Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), opt1.Value)
End Sub

Private Sub cmdimprimir_Click()
   If Gridprestamos.Rows > 1 Then
    Call imprimir
   End If
End Sub

Private Sub Form_Activate()
sqlventas.audit = True
sqlventas.programaactivo = Me.Caption
End Sub
'******************************************************************
Private Sub Form_Load()
    Call CARGAGRILLAprestamo(Gridprestamos)
    Call leerprestamosvigentes(Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), opt1.Value)
End Sub
  

Sub leerprestamosvigentes(mes, año, FILTRO)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventas
    If Check1.Value = 1 Then
    csql.sql = "select pc.empresa,pc.comprobante,pc.numerocuota,p.cantidadcuotas,pc.mes,pc.año,sum(pc.monto),pc.rut,pc.mesrebajado,pc.añorebajado,p.numeroprestamo "
    Else
    csql.sql = "select pc.empresa,pc.comprobante,pc.numerocuota,p.cantidadcuotas,pc.mes,pc.año,pc.monto,pc.rut,pc.mesrebajado,pc.añorebajado,p.numeroprestamo "
    
    End If
    
    csql.sql = csql.sql & "from sv_prestamo_cuota as pc inner join sv_prestamo as p "
    csql.sql = csql.sql & "on pc.comprobante=p.comprobante and pc.rut=p.rut "
    If FILTRO = "Verdadero" Then
    csql.sql = csql.sql & " where pc.mes='" & mes & "' and  pc.año='" & año & "' "
    End If
    If Check1.Value = 1 Then
    csql.sql = csql.sql & "group by pc.rut,pc.mes,pc.año "
    End If
    
    csql.sql = csql.sql & "order by pc.rut,pc.numerocuota asc "
    csql.Execute
     Gridprestamos.Rows = 1
    If csql.RowsAffected > 0 Then
        Gridprestamos.AutoRedraw = False
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Gridprestamos.Rows = Gridprestamos.Rows + 1
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 1).text = resultados(0)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 2).text = resultados(7)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 3).text = leernombretrabajador(resultados(7), resultados(0))
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 4).text = resultados(1)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 5).text = resultados(2) & " / " & resultados(3)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 6).text = resultados(4) & "-" & resultados(5)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 7).text = resultados(6)
            Gridprestamos.Cell(Gridprestamos.Rows - 1, 10).text = Format(200 + resultados(10), "00000")
            
            If FILTRO = "Verdadero" Then
                Gridprestamos.Cell(Gridprestamos.Rows - 1, 8).text = 1
            End If
                Gridprestamos.Cell(Gridprestamos.Rows - 1, 9).text = leercargado(resultados(0), resultados(7), resultados(4), resultados(5), Format(resultados(10) + 200, "00000"))
            If Gridprestamos.Cell(Gridprestamos.Rows - 1, 9).text <> "0" Then
                Gridprestamos.Cell(Gridprestamos.Rows - 1, 8).text = 0
            End If
            resultados.MoveNext
        Wend
    Else
        MsgBox "NO SE HAN ENCONTRADOS DATOS ", vbInformation, "ATENCION"
    End If
    csql.Close
    Gridprestamos.AutoRedraw = True
    Gridprestamos.Refresh
    Set csql = Nothing
    Set resultados = Nothing
End Sub
Sub CARGAGRILLAprestamo(Grid1 As Grid)
    Dim formatogrilla(10, 10) As String
    Dim K As Double
    
    Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 8
    
    formatogrilla(1, 1) = "EMPRESA"
    formatogrilla(1, 2) = "RUT"
    formatogrilla(1, 3) = "NOMBRE"
    formatogrilla(1, 4) = "NUMERO"
    formatogrilla(1, 5) = "Nº CUOTA"
    formatogrilla(1, 6) = "FECHA"
    formatogrilla(1, 7) = "MONTO EN U.F"
    formatogrilla(1, 8) = "CARGAR"
    formatogrilla(1, 9) = "MONTO CARGADO"
    formatogrilla(1, 10) = "PRESTAMO"
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "7"
    formatogrilla(2, 2) = "10"
    formatogrilla(2, 3) = "25"
    formatogrilla(2, 4) = "8"
    formatogrilla(2, 5) = "8"
    formatogrilla(2, 6) = "10"
    formatogrilla(2, 7) = "10"
    formatogrilla(2, 8) = "7"
    formatogrilla(2, 9) = "10"
    formatogrilla(2, 10) = "7"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "N"
    formatogrilla(3, 2) = "N"
    formatogrilla(3, 3) = "S"
    formatogrilla(3, 4) = "S"
    formatogrilla(3, 5) = "N"
    formatogrilla(3, 6) = "D"
    formatogrilla(3, 7) = "N"
    formatogrilla(3, 8) = ""
    formatogrilla(3, 9) = ""
    formatogrilla(3, 10) = ""
 
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = ""
    formatogrilla(4, 4) = ""
    formatogrilla(4, 5) = ""
    formatogrilla(4, 6) = "mm-yyyy"
    formatogrilla(4, 7) = ""
    formatogrilla(4, 8) = ""
    formatogrilla(4, 9) = ""
 
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    formatogrilla(5, 6) = "FALSE"
    formatogrilla(5, 7) = "TRUE"
    formatogrilla(5, 8) = "FALSE"
    formatogrilla(5, 9) = "TRUE"
    formatogrilla(5, 10) = "TRUE"
    Grid1.Cols = 11
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
    Grid1.Column(6).CellType = cellCalendar
    Grid1.Column(8).CellType = cellCheckBox
   
    Grid1.Column(0).Width = 0
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
    Rem Grid1.Enabled = False
End Sub

 


 

Private Sub Gridprestamos_Click()
    If Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 9).text <> "0" Then
        Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 8).text = "0"
    End If
End Sub

Private Sub Gridprestamos_DblClick()
 If Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 9).text <> "0" Then
        Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 8).text = "0"
    End If
End Sub

Private Sub Gridprestamos_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    If col = "6" Then
        If col <> NewCol Or row <> NewRow Then
            If Gridprestamos.Rows > 1 Then
                Call modificafecha(Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 2).text, Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 4).text, Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 5).text, Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 6).text, Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 1).text)
            End If
        End If
    End If
  If col = "1" Then
        If col <> NewCol Or row <> NewRow Then
            If Gridprestamos.Rows > 1 Then
                Call modificaempresa(Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 2).text, Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 4).text, Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 5).text, Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 1).text, Gridprestamos.Cell(Gridprestamos.ActiveCell.row, 1).text)
            End If
        End If
    End If
End Sub

 Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
    If command = "retorno" Then Call retorno
    If command = "modifica" Then
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
    End If
    If command = "elimina" Then
'       If MsgBox("REALMENTE DESEA ELIMINAR", vbYesNo) = vbYes Then
'        If Verifica_Permiso(Me.Caption, "elimina") = True Then
'            Call eliminar(numerotxt.text, dato1.text & dv.Caption, dato4.text & "-" & dato3.text & "-" & dato2.text)
'            Call Retorno
'         Else
'            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
'         End If
'        End If
        MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
    End If
 
'    If command = "imprime" Then
End Sub

Sub retorno()
    Gridprestamos.Rows = 1
   
    
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
    Grid4.AllowUserSort = True
    
    
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

 
Function leernombretrabajador(rut, Empresaconsulta) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventas
    
    csql.sql = "select nombre "
    csql.sql = csql.sql & "from " & clientesistema & "remu" & Empresaconsulta & ".mt_fijo "
    csql.sql = csql.sql & "where rut='" & rut & "' "
    csql.Execute
   
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leernombretrabajador = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
End Function
 

Sub modificafecha(rut, NUMERO, cuota, fecha, Empresaconsulta)
    Dim CAMPOS(6, 3) As String
    Dim op As Integer
    
    Dim CUOTAS As Variant
    
    CUOTAS = Split(cuota, "/")
    
    CAMPOS(0, 0) = "mes"
    CAMPOS(1, 0) = "año"
    CAMPOS(2, 0) = ""
    CAMPOS(0, 1) = Format(fecha, "mm")
    CAMPOS(1, 1) = Format(fecha, "yyyy")
    CAMPOS(0, 2) = "sv_prestamo_cuota"
    condicion = "empresa='" & Empresaconsulta & "' and rut='" & rut & "' and comprobante='" & NUMERO & "' and numerocuota='" & CDbl(CUOTAS(0)) & "'"
    op = 3
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    
    
End Sub
Sub modificaempresa(rut, NUMERO, cuota, empresa, Empresaconsulta)
    Dim CAMPOS(6, 3) As String
    Dim op As Integer
    
    Dim CUOTAS As Variant
    
    CUOTAS = Split(cuota, "/")
    
    CAMPOS(0, 0) = "empresa"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 1) = empresa
    CAMPOS(0, 2) = "sv_prestamo_cuota"
    condicion = "rut='" & rut & "' and comprobante='" & NUMERO & "' and numerocuota='" & CDbl(CUOTAS(0)) & "'"
    op = 3
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    
    
End Sub

Private Sub opt1_Click()
    cmdenviar.Visible = True
     Call leerprestamosvigentes(Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), opt1.Value)
End Sub

Private Sub opt2_Click()
    cmdenviar.Visible = False
    Call leerprestamosvigentes(Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), opt1.Value)

End Sub
Sub enviararemuneraciones(Grid1 As Grid)
    Dim K As Double
    For K = 1 To Grid1.Rows - 1
        If Grid1.Cell(K, 8).text = "1" Then
            Call grabaprestamo(Grid1.Cell(K, 1).text, Grid1.Cell(K, 2).text, Grid1.Cell(K, 7).text, Grid1.Cell(K, 6).text, Grid1.Cell(K, 4).text, Grid1.Cell(K, 5).text, Grid1.Cell(K, 10).text)
        End If
    Next K
End Sub
Sub grabaprestamo(Empresaconsulta, rut, MONTO, fecha, numerocomprobante, numerocuota, CODIGO)
    Dim CAMPOS(12, 3) As String
    Dim op As Integer
    
    Dim CUOTAS As Variant
    
    CUOTAS = Split(numerocuota, "/")
    CAMPOS(0, 0) = "rut"
    CAMPOS(1, 0) = "mes"
    CAMPOS(2, 0) = "año"
    CAMPOS(3, 0) = "codtablacalculo"
    CAMPOS(4, 0) = "glosa"
    CAMPOS(5, 0) = "um"
    CAMPOS(6, 0) = "monto"
    CAMPOS(7, 0) = "codcentrocosto"
    CAMPOS(8, 0) = "duracion"
    CAMPOS(9, 0) = "duracionoriginal"
    CAMPOS(10, 0) = ""
    
    CAMPOS(0, 1) = rut
    CAMPOS(1, 1) = Format(fecha, "mm")
    CAMPOS(2, 1) = Format(fecha, "yyyy")
    CAMPOS(3, 1) = CODIGO
    CAMPOS(4, 1) = "PRESTAMO EN UF PROMOTORA " & (Val(CODIGO) - 200)
    CAMPOS(5, 1) = "DU"
    CAMPOS(6, 1) = Replace(MONTO, ",", ".")
    CAMPOS(7, 1) = "0300"
    CAMPOS(8, 1) = CUOTAS(0)
    CAMPOS(9, 1) = CUOTAS(1)
    
    CAMPOS(0, 2) = clientesistema & "remu" & Empresaconsulta & ".liquidacionhd"
    condicion = ""
    op = 2
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    
    
    '--------------------------------------------------------
    
    CAMPOS(0, 0) = "mesrebajado"
    CAMPOS(1, 0) = "añorebajado"
    CAMPOS(2, 0) = ""
 
    CAMPOS(0, 1) = Format(fecha, "mm")
    CAMPOS(1, 1) = Format(fecha, "yyyy")
 
    
    CAMPOS(0, 2) = "sv_prestamo_cuota"
    condicion = "empresa='" & Empresaconsulta & "' and rut='" & rut & "'and comprobante='" & numerocomprobante & "' and numerocuota='" & CDbl(CUOTAS(0)) & "' "
    op = 3
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    
End Sub
Function leercargado(Empresaconsulta, rut, mes, año, codcalculo) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventas
    
    csql.sql = "select monto "
    csql.sql = csql.sql & "from " & clientesistema & "remu" & Empresaconsulta & ".liquidacionhd "
    csql.sql = csql.sql & "where rut='" & rut & "' and mes='" & mes & "' and año='" & año & "' and codtablacalculo='" & codcalculo & "' "
    csql.Execute
    leercargado = "0"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leercargado = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
End Function

 Private Sub imprimir()
        Dim i As Long
        Call Titulos("LISTADO PARA CARGAR PRESTAMOS ")
        Gridprestamos.AutoRedraw = False
        Gridprestamos.PageSetup.HeaderMargin = 2
    
        Gridprestamos.PageSetup.TopMargin = 2
        Gridprestamos.PageSetup.LeftMargin = 1
        Gridprestamos.PageSetup.RightMargin = 1
        Gridprestamos.PageSetup.BottomMargin = 2
        
        Gridprestamos.PageSetup.FooterMargin = 2
        Gridprestamos.PageSetup.BlackAndWhite = True
        Gridprestamos.PageSetup.PrintFixedRow = True
        'Gridprestamos.PageSetup.Orientation = cellLandscape
        Gridprestamos.PageSetup.PrintFixedRow = True
        
        Gridprestamos.PrintPreview
        
        Gridprestamos.AutoRedraw = True
    End Sub

Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
   Gridprestamos.FixedRowColStyle = Fixed3D
    Gridprestamos.CellBorderColorFixed = vbButtonShadow
    Gridprestamos.ShowResizeTips = False
    Gridprestamos.ReportTitles.Clear
    Gridprestamos.PageSetup.CenterHorizontally = True
    Gridprestamos.PageSetup.Orientation = cellLandscape
    
      
    Gridprestamos.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    Gridprestamos.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Gridprestamos.PageSetup.HeaderAlignment = cellLeft
    Gridprestamos.PageSetup.HeaderFont.Name = "Verdana"
    Gridprestamos.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
   
  
   
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Gridprestamos.ReportTitles.Add objReportTitle
   
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "  |  " & "PERIODO  :  " & Format(fechasistema, "dd-mm-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Gridprestamos.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    Gridprestamos.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & usuarioSistema
    Gridprestamos.PageSetup.FooterAlignment = cellRight
    Gridprestamos.PageSetup.FooterFont.Name = "Verdana"
    Gridprestamos.PageSetup.FooterFont.Size = 7
    
    Gridprestamos.Range(0, 1, 0, Gridprestamos.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Gridprestamos.Range(0, 1, 0, Gridprestamos.Cols - 1).Borders(cellEdgeTop) = cellThick
    Gridprestamos.Range(0, 1, 0, Gridprestamos.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Gridprestamos.Range(0, 1, 0, Gridprestamos.Cols - 1).Borders(cellEdgeRight) = cellThick
    Gridprestamos.Range(0, 1, 0, Gridprestamos.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Gridprestamos.Range(0, 1, 0, Gridprestamos.Cols - 1).Borders(cellInsideVertical) = cellThick

End Sub

