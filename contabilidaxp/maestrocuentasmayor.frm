VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form maestro01 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Cuentas del Mayor"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14535
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   642
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   969
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11160
      TabIndex        =   18
      Top             =   8880
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
         TabIndex        =   20
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   8175
      Left            =   7875
      TabIndex        =   5
      Top             =   45
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   14420
      BackColor       =   16744576
      Caption         =   "Plan de Cuentas"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraAbajo =   4194304
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
      ColorTextShadow =   16744576
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Diseña Formulario 1846"
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   7200
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Diseña Formulario 1847"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7200
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Diseña Capital Propio"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6720
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Diseña Balance Consolidado"
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6720
         Width           =   2415
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         Height          =   6105
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6255
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   7223
      BackColor       =   16744576
      Caption         =   "Historico"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraAbajo =   4194304
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SALDOS 
         Height          =   3615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   6376
         _Version        =   393216
         BackColor       =   16776436
         ForeColor       =   12582912
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   16761024
         BackColorSel    =   16777215
         ForeColorSel    =   16744576
         BackColorBkg    =   16776436
         GridColor       =   -2147483635
         GridColorFixed  =   12582912
         GridLinesFixed  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   45
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   7011
      BackColor       =   16744576
      Caption         =   "Datos de la Cuenta"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   16744576
      ColorBarraAbajo =   4194304
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
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   375
         Left            =   6120
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin XPFrame.FrameXp tipo 
         Height          =   735
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   1296
         BackColor       =   16761024
         Caption         =   "Tipo de cuenta"
         BackColor       =   16761024
         ColorBarraArriba=   16744576
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
         Begin VB.OptionButton tipocuenta3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Resultado"
            Height          =   255
            Left            =   3600
            TabIndex        =   41
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton tipocuenta2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Pasivo"
            Height          =   255
            Left            =   2160
            TabIndex        =   40
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton tipocuenta1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Activo"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   1455
         End
      End
      Begin XPFrame.FrameXp auxiliares 
         Height          =   1575
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2778
         BackColor       =   16761024
         Caption         =   "Tipos de Analisis"
         BackColor       =   16761024
         ColorBarraArriba=   16744576
         ColorBarraAbajo =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   43
            Top             =   1320
            Width           =   255
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   29
            Top             =   1080
            Width           =   255
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   28
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   27
            Top             =   840
            Width           =   255
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   26
            Top             =   600
            Width           =   255
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   25
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   600
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   23
            Top             =   360
            Width           =   255
         End
         Begin VB.CheckBox check7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   22
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Patrimonio"
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
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Capital Propio"
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
            Height          =   255
            Left            =   3120
            TabIndex        =   37
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Activo FIjo"
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
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto Harina"
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
            Height          =   255
            Left            =   3120
            TabIndex        =   35
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto Carne"
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
            Height          =   255
            Left            =   3120
            TabIndex        =   34
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto Licor"
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
            Height          =   255
            Left            =   3120
            TabIndex        =   33
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
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
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Centro de Costo"
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
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Correntista"
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
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "nombre"
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         MaxLength       =   4
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   10
         Top             =   360
         Width           =   375
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
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "codigo"
         Top             =   360
         Width           =   375
      End
      Begin CoolButtons.cool_Button graba 
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Grabar"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   45
      Top             =   8400
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
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
End
Attribute VB_Name = "maestro01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub





Private Sub Command1_Click()
maestro21.Show

End Sub


Private Sub COMMAND2_Click()
MAESTRO20.Show

End Sub

Private Sub Command3_Click()
ma1847.Show

End Sub

Private Sub Command4_Click()
ma1846.Show

End Sub



Private Sub Command5_Click()
    form29.Show
End Sub

Private Sub dato1_GotFocus()
leecrcc
grillasaldos
Call cargatexto(dato1)
End Sub
Private Sub dato2_GotFocus()
tipocuenta1.Value = False
tipocuenta2.Value = False
tipocuenta3.Value = False

If Mid(dato1.text, 1, 1) = "1" Then tipocuenta1.Value = True
If Mid(dato1.text, 1, 1) = "2" Then tipocuenta2.Value = True
If Mid(dato1.text, 1, 1) <> "1" And Mid(dato1.text, 1, 1) <> "2" Then tipocuenta3.Value = True

Call cargatexto(dato2)
End Sub
Private Sub dato3_GotFocus()
Call cargatexto(dato3)
End Sub

Private Sub dato4_GotFocus()
Call cargatexto(dato4)
If MODIFI = 0 Then leer
End Sub


Private Sub dato4_LostFocus()
Rem If modifi = 0 Then graba.Visible = True

End Sub




Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudamayor(dato4)
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, dato4, KeyCode)
End Sub







Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption
If dato1.text <> "" And dato2.text <> "" And dato3.text <> "" Then
Call dato3_KeyPress(13)
End If

End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

'dibu1.FileName = App.path & "\archivo.gif"
'dibu2.FileName = App.path & "\archivo.gif"


    
    Call Conectar_BD
 
    sc = 0
    opciones.Visible = False
DOCU(1) = "ACTIVO"
DOCU(2) = "PASIVO"
DOCU(3) = "RESULTADO"
CANDO = 3


Rem Call RECUPERAFECHA
Call CARGAPERMISO(Me.Name)
    

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)

    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then graba.Visible = True
    
End Sub


Sub leer()
    Rem lee cuenta madre
    
  If PermisosCuentasDelMayor(USUARIOSISTEMA, Format(dato1.text & dato2.text & dato3.text, "00000000")) = False Then
    MsgBox "USTED NO TIENE PRIVILEGIOS PARA ACCEDER A ESTA CUENTA", vbCritical, "ATENCION"
  dato1.SetFocus
  Exit Sub
  End If
    
    
    If dato2.text = "00" And dato3.text = "0000" Then GoTo lee3:
    If dato2.text <> "00" And dato3.text <> "0000" Then GoTo lee2:
    If dato2.text <> "00" And dato3.text = "0000" Then GoTo lee1

lee1:
         
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "tipo"
    campos(3, 0) = "ctacte"
    campos(4, 0) = "crcc"
    campos(5, 0) = "banco"
    campos(6, 0) = "ila"
    campos(7, 0) = "ica"
    campos(8, 0) = "iha"
    campos(9, 0) = "activo"
    campos(10, 0) = ""
      campos(0, 2) = "cuentasdelmayor"
      condicion = "codigo=" + "'" + dato1.text + "00" + "0000" + "' and año='" + Format(fechasistema, "yyyy") + "'"
      op = 5
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
      If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
      GoTo lee3:
lee2:    Rem lee cuenta madre
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato4.Tag 'DESCRIPCION
    campos(2, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + "0000" + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    
lee3:    Rem lee cuenta madre 2
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "tipo"
    campos(3, 0) = "ctacte"
    campos(4, 0) = "crcc"
    campos(5, 0) = "banco"
    campos(6, 0) = "ila"
    campos(7, 0) = "ica"
    campos(8, 0) = "iha"
    campos(9, 0) = "activo"
    campos(10, 0) = "patrimonio"
    campos(11, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then graba.Visible = False: dato4.SetFocus: GoTo no:
    
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    
    DATOSSALDOS
   
    
    opciones.SetFocus
    graba.Visible = False
    
no:
   
End Sub
Sub leersiguiente()

    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "tipo"
    campos(3, 0) = "ctacte"
    campos(4, 0) = "crcc"
    campos(5, 0) = "banco"
    campos(6, 0) = "ila"
    campos(7, 0) = "ica"
    campos(8, 0) = "iha"
    campos(9, 0) = "activo"
    campos(10, 0) = "patrimonio"
    campos(11, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo>" + "'" + dato1.text + dato2.text + dato3.text + "' and año='" + Format(fechasistema, "yyyy") + "' order by codigo"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
    
    
no:
   
    
End Sub
Sub leeranterior()

    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "tipo"
    campos(3, 0) = "ctacte"
    campos(4, 0) = "crcc"
    campos(5, 0) = "banco"
    campos(6, 0) = "ila"
    campos(7, 0) = "ica"
    campos(8, 0) = "iha"
    campos(9, 0) = "activo"
    campos(10, 0) = "patrimonio"
    campos(11, 0) = ""
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo<" + "'" + dato1.text + dato2.text + dato3.text + "' and año='" + Format(fechasistema, "yyyy") + "' order by codigo desc"
    

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
  
no:
   
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
    dato4.text = sqlconta.response(1, 3)
    If sqlconta.response(2, 3) = "1" Then tipocuenta1.Value = True
    If sqlconta.response(2, 3) = "2" Then tipocuenta2.Value = True
    If sqlconta.response(2, 3) = "3" Then tipocuenta3.Value = True
    check7.Value = Val(sqlconta.response(3, 3))
    Check1.Value = Val(sqlconta.response(4, 3))
    Check2.Value = Val(sqlconta.response(5, 3))
    Check3.Value = Val(sqlconta.response(6, 3))
    Check4.Value = Val(sqlconta.response(7, 3))
    Check5.Value = Val(sqlconta.response(8, 3))
    Check6.Value = Val(sqlconta.response(9, 3))
    Check8.Value = Val(sqlconta.response(10, 3))
    
    graba.Visible = False
    auxiliares.Enabled = False
    
    
    

fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
   
  
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
 

End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    
    
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
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


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub grabar()
       
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "tipo"
    campos(3, 0) = "ctacte"
    campos(4, 0) = "crcc"
    campos(5, 0) = "banco"
    campos(6, 0) = "ila"
    campos(7, 0) = "ica"
    campos(8, 0) = "iha"
    campos(9, 0) = "activo"
    campos(10, 0) = "año"
    campos(11, 0) = "patrimonio"
    campos(12, 0) = ""
    campos(0, 1) = dato1.text + dato2.text + dato3.text
    campos(1, 1) = dato4.text
    If tipocuenta1.Value = True Then campos(2, 1) = "1"
    If tipocuenta2.Value = True Then campos(2, 1) = "2"
    If tipocuenta3.Value = True Then campos(2, 1) = "3"
    campos(3, 1) = Str(check7.Value)
    campos(4, 1) = Str(Check1.Value)
    campos(5, 1) = Str(Check2.Value)
    campos(6, 1) = Str(Check3.Value)
    campos(7, 1) = Str(Check4.Value)
    campos(8, 1) = Str(Check5.Value)
    campos(9, 1) = Str(Check6.Value)
    campos(10, 1) = Format(fechasistema, "yyyy")
    campos(11, 1) = Str(Check8.Value)
    
    campos(0, 2) = "cuentasdelmayor"
    If MODIFI = 1 Then condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If MODIFI = 0 Then grabar2
    MODIFI = 0
no:
retorno
End Sub
Sub grabar2()
     sqlconta.audit = False
     
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = ""
    campos(0, 1) = dato1.text + dato2.text + dato3.text
    campos(1, 1) = Format(fechasistema, "yyyy")
    campos(0, 2) = "saldosdelmayor"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call GRABACRCC
    

End Sub
Sub GRABACRCC()
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM centrosdecosto "
       
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
       
        If csql2.RowsAffected > 0 Then
     
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
         Call GRABAR3(resultados2(0), dato1.text + dato2.text + dato3.text)
         
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    

End Sub
Sub GRABAR3(CRCC, cuenta)
     sqlconta.audit = False
     
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "cuenta"
    campos(3, 0) = ""
    campos(0, 1) = CRCC
    campos(1, 1) = Mid(fechasistema, 7, 4)
    campos(2, 1) = cuenta

    campos(0, 2) = "saldoscentrosdecosto"
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    sqlconta.audit = True
    

End Sub


Sub ELIMINAR()
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    sqlconta.audit = False
    
    campos(0, 2) = "saldoscentrosdecosto"
    condicion = "cuenta=" + "'" + dato1.text + dato2.text + dato3.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    campos(0, 2) = "saldosdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    sqlconta.audit = True
    
    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub






Private Sub graba_Click()
grabar
End Sub

Private Sub Image1_Click()
Unload Me



End Sub

Private Sub List1_Click()
If Val(Mid(List1.text, 17, 2)) <> 0 Then

dato1.Enabled = True
dato2.Enabled = True
dato3.Enabled = True
dato4.Enabled = True

dato1.text = Mid(List1.text, 17, 2)
dato2.text = Mid(List1.text, 20, 2)
dato3.text = Mid(List1.text, 23, 4)
dato4.SetFocus
End If



End Sub

Private Sub List1_DblClick()
If Mid(List1.text, 17, 2) <> "  " Then
dato1.Enabled = True
dato2.Enabled = True
dato3.Enabled = True
dato4.Enabled = True

dato1.text = Mid(List1.text, 17, 2)
dato2.text = Mid(List1.text, 20, 2)
dato3.text = Mid(List1.text, 23, 4)
dato4.SetFocus
End If


End Sub

Private Sub MANUAL_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub


Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then retorno

If command = "modifica" Then

 If Verifica_Permiso(Me.Caption, "modifica") = True Then
    modifica
 Else
     MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
 End If
    
End If

If command = "elimina" Then
 If Verifica_Permiso(Me.Caption, "elimina") = True Then
    ELIMINA
 Else
     MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
 End If
End If
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "imprime" Then imprimir
If command = "movimientos" Then movimientos

End Sub
Sub ELIMINA()

If saldoglobal = 0 Then
If Verifica_Permiso(Me.Caption, "elimina") = True Then
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus

Else
 MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
End If
Else
MsgBox ("CUENTA CON MOVIMIENTOS IMPOSIBLE ELIMINAR ")
End If
End Sub
Sub modifica()
If saldoglobal = 0 Or Verifica_Permiso(Me.Caption, "autoriza") = True Then
auxiliares.Enabled = True

disponible (True)
habilita (False)
dato1.Enabled = False
dato2.Enabled = False
dato3.Enabled = False
dato4.SetFocus
MODIFI = 1
graba.Visible = True
Else
MsgBox ("CUENTA CON MOVIMIENTOS IMPOSIBLE MODIFICAR ")
End If

End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
opciones.Visible = False
graba.Visible = False

dato1.Enabled = True
dato1.SetFocus
MODIFI = 0
auxiliares.Enabled = True


End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
Check1.Value = False
Check1.Value = False
Check2.Value = False
Check3.Value = False
Check4.Value = False
Check5.Value = False
Check6.Value = False
check7.Value = False
Check8.Value = False




graba.Visible = False

End Sub

Sub imprimir()
    informa01.Show
    
End Sub
Sub grilla()

End Sub
Sub cabeza()
    
End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo,ctacte,crcc "
        csql.sql = csql.sql + "FROM cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' "
        csql.sql = csql.sql + " order by codigo"
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                
                dato(1) = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4): colu(1) = 15: tipodato(1) = "s"
                dato(2) = resultados(1): colu(2) = 52: tipodato(2) = "s"
                dato(3) = resultados(2) + " " + DOCU$(Val(resultados(2)))
                dato(4) = resultados(3)
                dato(5) = resultados(4)
                dato(6) = resultados(5) + " " + DOCU2$(Val(resultados(5)))
                colu(3) = 15: tipodato(3) = "s"
                colu(4) = 3: tipodato(4) = "s"
                colu(5) = 20: tipodato(5) = "s"
                colu(6) = 20: tipodato(6) = "s"
                 cancolu = 6
                grilla
                resultados.MoveNext
            Wend
            resultados.Close
            
            Set resultados = Nothing

        End If
    

End Sub

Sub DATOSSALDOS()
Dim debe As Double
Dim haber As Double
Call LEERSALDOS(dato1.text + dato2.text + dato3.text, fechasistema)
saldoglobal = leersaldomayor(dato1.text + dato2.text + dato3.text, fechasistema)


End Sub
Sub grillasaldos()
SALDOS.Cols = 4
SALDOS.Rows = 14
SALDOS.ColWidth(0) = 120 * 14
SALDOS.ColWidth(1) = 100 * 15
SALDOS.ColWidth(2) = 100 * 15
SALDOS.ColWidth(3) = 100 * 15
SALDOS.TextMatrix(0, 0) = "MESES   "
SALDOS.TextMatrix(0, 1) = "DEBE    "
SALDOS.TextMatrix(0, 2) = "HABER   "
SALDOS.TextMatrix(0, 3) = "SALDO   "
SALDOS.TextMatrix(1, 0) = "AÑO ANTERIOR"
SALDOS.TextMatrix(2, 0) = "ENERO"
SALDOS.TextMatrix(3, 0) = "FEBRERO"
SALDOS.TextMatrix(4, 0) = "MARZO"
SALDOS.TextMatrix(5, 0) = "ABRIL"
SALDOS.TextMatrix(6, 0) = "MAYO"
SALDOS.TextMatrix(7, 0) = "JUNIO"
SALDOS.TextMatrix(8, 0) = "JULIO"
SALDOS.TextMatrix(9, 0) = "AGOSTO"
SALDOS.TextMatrix(10, 0) = "SEPTIEMBRE"
SALDOS.TextMatrix(11, 0) = "OCTUBRE"
SALDOS.TextMatrix(12, 0) = "NOVIEMBRE "
SALDOS.TextMatrix(13, 0) = "DICIEMBRE "
For k = 1 To 13
SALDOS.TextMatrix(k, 1) = "0"
SALDOS.TextMatrix(k, 2) = "0"
SALDOS.TextMatrix(k, 3) = "0"
Next k
End Sub

Sub LEERSALDOS(codigo, fecha)
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim fecha1 As String
    Dim fecha2 As String
    Dim resultados As rdoResultset
  Dim NIVEL As String
  
  grillasaldos
    fecha1 = Format(fecha, "YYYY") + "-01-01"
    fecha2 = Format(fecha, "YYYY-MM-DD")
        Set csql2.ActiveConnection = contadb
       NIVEL = "3"
        If Mid(codigo, 5, 5) = "0000" Then NIVEL = "2"
        If Mid(codigo, 3, 6) = "000000" Then NIVEL = "1"
        csql2.sql = "SELECT fecha,sum(monto),dh "
        csql2.sql = csql2.sql + "FROM movimientoscontables WHERE fecha between '" + fecha1 + "' and '" + fecha2 + "' "
        If NIVEL = "1" Then
        csql2.sql = csql2.sql + "and mid(codigocuenta,1,2)='" + Mid(codigo, 1, 2) + "' "
        End If
        If NIVEL = "2" Then
        csql2.sql = csql2.sql + "and mid(codigocuenta,1,4)='" + Mid(codigo, 1, 4) + "' "
        End If
        If NIVEL = "3" Then
        csql2.sql = csql2.sql + "and codigocuenta='" + codigo + "' "
        End If
        
        
        csql2.sql = csql2.sql + " group by mid(fecha,1,7),dh "
        csql2.Execute
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
         
        Set resultados = csql2.OpenResultset
        While Not resultados.EOF
        If resultados(2) = "D" Then
        SALDOS.TextMatrix(Format(resultados(0), "mm") + 1, 1) = Format(resultados(1), "###,###,###,##0")
        Else
        SALDOS.TextMatrix(Format(resultados(0), "mm") + 1, 2) = Format(resultados(1), "###,###,###,##0")
        End If
        
        
        
        resultados.MoveNext
        Wend
          
          resultados.Close
            Set resultados = Nothing
        End If
  

sumador = leersaldomayoranterior(dato1.text + dato2.text + dato3.text)
If sumador > 0 Then
SALDOS.TextMatrix(1, 1) = Format(sumador, "###,###,###,##0")
Else
SALDOS.TextMatrix(1, 2) = Format(sumador, "###,###,###,##0")
End If
SALDOS.TextMatrix(1, 3) = Format(sumador, "###,###,###,##0")
''debe = 0
''haber = 0
For k = 2 To 13
sumador = sumador + CDbl(SALDOS.TextMatrix(k, 1)) - CDbl(SALDOS.TextMatrix(k, 2))
SALDOS.TextMatrix(k, 3) = Format(sumador, "###,###,##0")
Next k

'
'
    

  
End Sub



Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub



Private Sub opciones_GotFocus()
graba.Visible = False

MANUAL.SetFocus

End Sub


Sub leeplandecuenta()
leecrcc
End Sub
Sub leecrcc()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta" & empresaactiva & ".cuentasdelmayor WHERE año='" + Format(fechasistema, "yyyy") + "'"
       
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
         List1.Clear
         
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        
        If Mid(resultados2(0), 3, 6) = "000000" Then List1.AddItem "": List1.AddItem Mid(resultados2(0), 1, 2) + "." + Mid(resultados2(0), 3, 2) + "." + Mid(resultados2(0), 5, 4) + "  " + resultados2(1): GoTo no:
        If Mid(resultados2(0), 5, 4) = "0000" Then List1.AddItem "": List1.AddItem "      " + Mid(resultados2(0), 1, 2) + "." + Mid(resultados2(0), 3, 2) + "." + Mid(resultados2(0), 5, 4) + "  " + resultados2(1): List1.AddItem "": GoTo no:
        If Mid(resultados2(0), 5, 4) <> "0000" Then List1.AddItem "                " + Mid(resultados2(0), 1, 2) + "." + Mid(resultados2(0), 3, 2) + "." + Mid(resultados2(0), 5, 4) + "  " + resultados2(1): GoTo no:
no:
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    
    
    

End Sub

Sub movimientos()
Rem cartola = "mayor:" + dato1.text + dato2.text + dato3.text
informa04.cmdato1.text = dato1.text
informa04.cmdato2.text = dato2.text
informa04.cmdato3.text = dato3.text
informa04.cmnombre = dato4.text
informa04.sbtab1.Tab = 0
informa04.cmindi = True
informa04.Show
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
