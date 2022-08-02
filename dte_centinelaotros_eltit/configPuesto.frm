VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form configPuesto 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Configurar Puesto de Trabajo"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4935
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   8705
      BackColor       =   12648384
      Caption         =   " Configurar Puesto de Trabajo"
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
      Begin XPFrame.FrameXp frmConectar 
         Height          =   375
         Left            =   2820
         TabIndex        =   14
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Conectar"
         CaptionEstilo3D =   1
         BackColor       =   49344
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
      End
      Begin VB.Frame frmAdicionales 
         BackColor       =   &H00C0FFC0&
         Caption         =   " Informacion Adicional "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   3975
         Left            =   4560
         TabIndex        =   31
         Top             =   540
         Width           =   4515
         Begin VB.CommandButton cmdRuta 
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
            Left            =   3960
            Picture         =   "configPuesto.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   3240
            Width           =   315
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0FFC0&
            Caption         =   " Impresión Directa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   300
            TabIndex        =   13
            Top             =   2880
            Width           =   3075
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFC0&
            Caption         =   " Activar Seguridad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   300
            TabIndex        =   12
            Top             =   2520
            Width           =   3075
         End
         Begin VB.TextBox dato11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   10
            Top             =   1080
            Width           =   435
         End
         Begin VB.TextBox dato12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   11
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox dato10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   9
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lblRuta 
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
            Height          =   315
            Left            =   180
            TabIndex        =   39
            Top             =   3600
            Width           =   4125
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Ruta de Actualizacion del Sistema"
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
            TabIndex        =   38
            Top             =   3240
            Width           =   3735
         End
         Begin VB.Label lblCaja 
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
            Height          =   315
            Left            =   180
            TabIndex        =   37
            Top             =   2160
            Width           =   4125
         End
         Begin VB.Label lblBodega 
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
            Height          =   315
            Left            =   180
            TabIndex        =   36
            Top             =   1440
            Width           =   4125
         End
         Begin VB.Label lblEmpresa 
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
            Height          =   315
            Left            =   180
            TabIndex        =   35
            Top             =   720
            Width           =   4125
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Empresa"
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
            TabIndex        =   34
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Bodega"
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
            TabIndex        =   33
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Caja"
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
            TabIndex        =   32
            Top             =   1800
            Width           =   1455
         End
      End
      Begin VB.Frame frmServidor 
         BackColor       =   &H00C0FFC0&
         Caption         =   " Informacion del Servidor "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   3375
         Left            =   240
         TabIndex        =   20
         Top             =   540
         Width           =   4215
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0FFC0&
            Caption         =   " Ver Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   1920
            TabIndex        =   17
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox dato4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3540
            MaxLength       =   3
            TabIndex        =   3
            Top             =   360
            Width           =   435
         End
         Begin VB.TextBox dato3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   2
            Top             =   360
            Width           =   435
         End
         Begin VB.TextBox dato2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2460
            MaxLength       =   3
            TabIndex        =   1
            Top             =   360
            Width           =   435
         End
         Begin VB.TextBox dato1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   0
            Top             =   360
            Width           =   435
         End
         Begin VB.TextBox dato9 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2460
            TabIndex        =   8
            Top             =   2880
            Width           =   1515
         End
         Begin VB.TextBox dato8 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2460
            TabIndex        =   7
            Top             =   2520
            Width           =   1515
         End
         Begin VB.TextBox dato7 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2460
            TabIndex        =   6
            Top             =   2160
            Width           =   1515
         End
         Begin VB.TextBox dato6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1920
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox dato5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            TabIndex        =   4
            Top             =   720
            Width           =   2055
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   3900
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   " Bases de Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1740
            Width           =   3675
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3420
            TabIndex        =   29
            Top             =   420
            Width           =   105
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2880
            TabIndex        =   28
            Top             =   420
            Width           =   105
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2340
            TabIndex        =   27
            Top             =   420
            Width           =   105
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Sistema de Trigo"
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
            TabIndex        =   26
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Sistema de Ventas"
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
            TabIndex        =   25
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Sistema de Gestión"
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
            TabIndex        =   24
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Password"
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
            TabIndex        =   23
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Usuario"
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
            TabIndex        =   22
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Dirección IP"
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
            TabIndex        =   21
            Top             =   360
            Width           =   1455
         End
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   8760
         TabIndex        =   19
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
      Begin XPFrame.FrameXp frmGuardar 
         Height          =   375
         Left            =   2820
         TabIndex        =   15
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Guardar"
         CaptionEstilo3D =   1
         BackColor       =   49344
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
         Enabled         =   0   'False
      End
      Begin XPFrame.FrameXp frmModificar 
         Height          =   375
         Left            =   300
         TabIndex        =   16
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Modificar"
         CaptionEstilo3D =   1
         BackColor       =   49344
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
      End
   End
End
Attribute VB_Name = "configPuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private campos(5, 3) As String
'Private Sub Command1_Click()
'    Text3.text = EncryptStr(Text1.text, Text2.text)
'End Sub
'Private Sub Command2_Click()
'    Text4.text = UnEncryptStr(Text3.text, Text2.text)
'End Sub

Private Sub cmdRuta_Click()
    Ruta.Show vbModal
    lblRuta.Caption = rutaUpdate
End Sub

Private Sub Check3_Click()
    If Check3.Value = 0 Then
        dato6.PasswordChar = "*"
    Else
        dato6.PasswordChar = ""
    End If
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
    
    Private Sub dato7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
    End Sub
    
    Private Sub dato8_GotFocus()
        Call VerificarCajas(Me, dato8)
        Call selecciona(dato8)
    End Sub
    
    Private Sub dato9_GotFocus()
        Call VerificarCajas(Me, dato9)
        Call selecciona(dato9)
    End Sub
    
    Private Sub dato10_GotFocus()
        Call VerificarCajas(Me, dato10)
        Call selecciona(dato10)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Empresa"
    End Sub
    
    Private Sub dato11_GotFocus()
        Call VerificarCajas(Me, dato11)
        Call selecciona(dato11)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Bodega"
    End Sub
    
    Private Sub dato12_GotFocus()
        Call VerificarCajas(Me, dato12)
        Call selecciona(dato12)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Caja"
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
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato7)
    End Sub
    
    Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato8)
    End Sub
    
    Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaEmpresa(dato10)
        Else
            Call Flechas(KeyCode, dato9)
        End If
    End Sub
    
    Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaBodega(dato11)
        Else
            Call Flechas(KeyCode, dato10)
        End If
    End Sub
    
    Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaCaja(dato12, dato11.text)
        Else
            Call Flechas(KeyCode, dato11)
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
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato2.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato3.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato4.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato5.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato6.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato7.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato8.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato9.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato10_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato10.text = ceros(dato10)
            lblEmpresa.Caption = leerEmpresa(dato10.text)
            rubro = leerRubro(dato10.text)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            dato11.text = ceros(dato11)
            lblBodega.Caption = leerBodega(dato11.text)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato12_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            dato12.text = ceros(dato12)
            lblCaja.Caption = leerCaja(dato12.text)
            SendKeys "{Tab}"
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
        Call Centrar(Me)
        If leer = True Then
            frmServidor.Enabled = False
            frmConectar.Enabled = False
            frmModificar.Visible = True
        End If
    End Sub

    Private Sub frmCerrar_BarClick()
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmConectar_BarClick()
        Dim cadena As String
        cadena = dato1.text & "." & dato2.text & "." & dato3.text & "." & dato4.text
        Call cambiaColor(frmConectar)
        frmConectar.CaptionEstilo3D = Inserted
        Call ConectarConfiguracion(cadena, LCase(dato5.text), LCase(dato6.text))
        servidor = cadena
        usuario = LCase(dato5.text)
        password = LCase(dato6.text)
        baseDatos = LCase(dato7.text)
        baseVentas = LCase(dato8.text)
        baseTrigo = LCase(dato9.text)
        frmAdicionales.Enabled = True
        frmGuardar.Enabled = True
        dato10.SetFocus
    End Sub
    
    Private Sub frmConectar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmConectar)
        frmConectar.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmGuardar_BarClick()
        Call cambiaColor(frmGuardar)
        frmGuardar.CaptionEstilo3D = Inserted
        Call guardarArchivo
        Call leerDatosConectar
        Unload Me
    End Sub
    
    Private Sub frmGuardar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmGuardar)
        frmGuardar.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmModificar_BarClick()
        Call cambiaColor(frmModificar)
        frmModificar.CaptionEstilo3D = Inserted
        frmAdicionales.Enabled = True
        frmGuardar.Enabled = True
        frmServidor.Enabled = True
        frmConectar.Enabled = True
        frmModificar.Visible = False
        dato1.SetFocus
    End Sub
    
    Private Sub frmModificar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmModificar)
        frmModificar.CaptionEstilo3D = Raised
    End Sub

    Private Sub guardarArchivo()
        Dim cadena As String
        'SERVIDOR
        cadena = dato1.text & "." & dato2.text & "." & dato3.text & "." & dato4.text
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("SERVIDOR", cadena, "confiVentas.txt")
        'USUARIO
        cadena = LCase(dato5.text)
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("USUARIO", cadena, "confiVentas.txt")
        'PASSWORD
        cadena = LCase(dato6.text)
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("PASSWORD", cadena, "confiVentas.txt")
        'BASEDATOS
        cadena = LCase(dato7.text)
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("BASEDATOS", cadena, "confiVentas.txt")
        'BASEVENTAS
        cadena = LCase(dato8.text)
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("BASEVENTAS", cadena, "confiVentas.txt")
        'BASETRIGO
        cadena = LCase(dato9.text)
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("BASETRIGO", cadena, "confiVentas.txt")
        'EMPRESA
        cadena = dato10.text
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("EMPRESA", cadena, "confiVentas.txt")
        'BODEGA
        cadena = dato11.text
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("BODEGA", cadena, "confiVentas.txt")
        'CAJA
        cadena = dato12.text
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("CAJA", cadena, "confiVentas.txt")
        'SEGURIDAD
        If Check1.Value = 0 Then
            cadena = "FALSE"
        Else
            cadena = "TRUE"
        End If
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("SEGURIDAD", cadena, "confiVentas.txt")
        'IMPRIMEDIRECTO
        If Check2.Value = 0 Then
            cadena = "FALSE"
        Else
            cadena = "TRUE"
        End If
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("IMPRIMEDIRECTO", cadena, "confiVentas.txt")
        'RUTA
        cadena = lblRuta.Caption
        cadena = EncryptStr(cadena, pass)
        Call escribeArchivo("RUTA", cadena, "confiVentas.txt")
    End Sub

    Private Function leerEmpresa(ByVal codigo As String) As String
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        campos(0, 0) = "nombre"
        campos(1, 0) = ""
        
        campos(0, 2) = baseDatos & ".g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.datos = campos
        Set sql.conexion = config
        Call sql.SQLUTIL(op, condicion)
        If sql.estado = 0 Then
            leerEmpresa = sql.datos(0, 3)
        Else
            leerEmpresa = ""
        End If
    End Function

    Private Function leerBodega(ByVal codigo As String) As String
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        campos(0, 0) = "nombre"
        campos(1, 0) = ""
        
        campos(0, 2) = baseDatos & rubro & ".r_maestrobodegas_" & rubro
        
        condicion = "local = '" & empresaActiva & "' AND codigobodega = '" & codigo & "'"
        op = 5
        sql.datos = campos
        Set sql.conexion = config
        Call sql.SQLUTIL(op, condicion)
        If sql.estado = 0 Then
            leerBodega = sql.datos(0, 3)
        Else
            leerBodega = ""
        End If
    End Function
    
    Private Function leerCaja(ByVal codigo As String) As String
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        campos(0, 0) = "descripcion"
        campos(1, 0) = ""
        
        campos(0, 2) = baseVentas & ".sv_maestrodecajas"
        
        condicion = "local = '" & empresaActiva & "' AND codigo = '" & codigo & "'"
        op = 5
        sql.datos = campos
        Set sql.conexion = config
        Call sql.SQLUTIL(op, condicion)
        If sql.estado = 0 Then
            leerCaja = sql.datos(0, 3)
        Else
            leerCaja = ""
        End If
    End Function

    Private Function leerRubro(ByVal codigo As String) As String
        Dim condicion As String
        Dim op As Integer
        Set sql = New CSQLUtil
        campos(0, 0) = "rubro"
        campos(1, 0) = ""
        
        campos(0, 2) = baseDatos & ".g_maestroempresas"
        
        condicion = "codigo = '" & codigo & "'"
        op = 5
        sql.datos = campos
        Set sql.conexion = config
        Call sql.SQLUTIL(op, condicion)
        If sql.estado = 0 Then
            leerRubro = sql.datos(0, 3)
        Else
            leerRubro = ""
        End If
    End Function

    Private Function leer() As Boolean
        Dim cadena As String
        Dim cad As Variant
        If ExisteArchivo(App.Path & "\confiVentas.txt") = True Then
            cadena = UnEncryptStr(leeArchivo("SERVIDOR", App.Path & "\confiVentas.txt"), pass)
            cad = Split(cadena, ".")
            dato1.text = cad(0)
            dato2.text = cad(1)
            dato3.text = cad(2)
            dato4.text = cad(3)
            dato5.text = UCase(UnEncryptStr(leeArchivo("USUARIO", App.Path & "\confiVentas.txt"), pass))
            dato6.text = UCase(UnEncryptStr(leeArchivo("PASSWORD", App.Path & "\confiVentas.txt"), pass))
            dato7.text = UCase(UnEncryptStr(leeArchivo("BASEDATOS", App.Path & "\confiVentas.txt"), pass))
            dato8.text = UCase(UnEncryptStr(leeArchivo("BASEVENTAS", App.Path & "\confiVentas.txt"), pass))
            dato9.text = UCase(UnEncryptStr(leeArchivo("BASETRIGO", App.Path & "\confiVentas.txt"), pass))
            Call ConectarConfiguracion(cadena, LCase(dato5.text), LCase(dato6.text))
            dato10.text = UnEncryptStr(leeArchivo("EMPRESA", App.Path & "\confiVentas.txt"), pass)
            lblEmpresa.Caption = leerEmpresa(dato10.text)
            dato11.text = UnEncryptStr(leeArchivo("BODEGA", App.Path & "\confiVentas.txt"), pass)
            lblBodega.Caption = leerBodega(dato11.text)
            dato12.text = UnEncryptStr(leeArchivo("CAJA", App.Path & "\confiVentas.txt"), pass)
            lblCaja.Caption = leerCaja(dato12.text)
            cadena = UnEncryptStr(leeArchivo("SEGURIDAD", App.Path & "\confiVentas.txt"), pass)
            If cadena = "TRUE" Then
                Check1.Value = 1
            Else
                Check1.Value = 0
            End If
            cadena = UnEncryptStr(leeArchivo("IMPRIMEDIRECTO", App.Path & "\confiVentas.txt"), pass)
            If cadena = "TRUE" Then
                Check2.Value = 1
            Else
                Check2.Value = 0
            End If
            lblRuta.Caption = UCase(UnEncryptStr(leeArchivo("RUTA", App.Path & "\confiVentas.txt"), pass))
            leer = True
        Else
            leer = False
        End If
    End Function

