VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ingreso01 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   10275
   ClientLeft      =   1050
   ClientTop       =   1245
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   14235
   WindowState     =   2  'Maximized
   Begin VB.Frame CONTROL 
      BackColor       =   &H000040C0&
      Caption         =   "DATOS PARA GRILLA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   4680
      TabIndex        =   81
      Top             =   8400
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command2 
         BackColor       =   &H00D9EFFF&
         Caption         =   "Retorno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame eliminatodo 
      BackColor       =   &H000040C0&
      Caption         =   "Elimina Todo El Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   4560
      TabIndex        =   75
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton sieliminatodo 
         BackColor       =   &H00D9EFFF&
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton noeliminatodo 
         BackColor       =   &H00D9EFFF&
         Caption         =   "No eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame opcioneliminalinea 
      BackColor       =   &H000040C0&
      Caption         =   "Elimina Linea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   4455
      Left            =   4320
      TabIndex        =   72
      Top             =   2880
      Visible         =   0   'False
      Width           =   4335
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid ELI 
         Height          =   3135
         Left            =   120
         TabIndex        =   78
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5530
         _Version        =   393216
         BackColor       =   16576
         ForeColor       =   8454143
         Rows            =   10
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   0
         BackColorBkg    =   16576
         GridColor       =   4210816
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton noeliminalinea 
         BackColor       =   &H00D9EFFF&
         Caption         =   "No eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton sieliminalinea 
         BackColor       =   &H00D9EFFF&
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   3960
         Width           =   1935
      End
   End
   Begin VB.Frame Evento 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   70
      Top             =   2400
      Visible         =   0   'False
      Width           =   10575
      Begin VB.Label textoevento 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   120
         Width           =   10095
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   3
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   10575
      End
   End
   Begin VB.Frame opcionelimina 
      BackColor       =   &H000040C0&
      Caption         =   "Elimina Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2055
      Left            =   2280
      TabIndex        =   65
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton noelimina 
         BackColor       =   &H00D9EFFF&
         Caption         =   "no elimina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton eliminalinea 
         BackColor       =   &H00D9EFFF&
         Caption         =   "Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton eliminacomprobante 
         BackColor       =   &H00D9EFFF&
         Caption         =   "comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame opcionmodifica 
      BackColor       =   &H000040C0&
      Caption         =   "Modifica Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   62
      Top             =   6720
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton nomodifica 
         BackColor       =   &H00D9EFFF&
         Caption         =   "no modifica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton agregalinea 
         BackColor       =   &H00D9EFFF&
         Caption         =   "agrega Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton modificacabeza 
         BackColor       =   &H00D9EFFF&
         Caption         =   "Encabezado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton modificalinea 
         BackColor       =   &H00D9EFFF&
         Caption         =   "Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame TIPOS 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tipo de Documentos"
      Height          =   1575
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
      Begin VB.ListBox LISTATIPOS 
         BackColor       =   &H00FDFBE3&
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
         Height          =   1230
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame comprodatos 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   1455
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   13815
      Begin VB.TextBox dato6 
         BackColor       =   &H00E1FFFD&
         DataField       =   "DSADSADSA"
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
         Left            =   720
         MaxLength       =   2
         TabIndex        =   40
         Tag             =   "codigocuenta"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   8640
         MaxLength       =   2
         TabIndex        =   39
         Tag             =   "tipodocumento"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato7 
         BackColor       =   &H00E1FFFD&
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
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   38
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato8 
         BackColor       =   &H00E1FFFD&
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
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   37
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox dato12 
         BackColor       =   &H00E1FFFD&
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
         Left            =   4080
         MaxLength       =   50
         TabIndex        =   36
         Tag             =   "glosacontable"
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox dato17 
         BackColor       =   &H00E1FFFD&
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
         Left            =   11040
         MaxLength       =   4
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox dato16 
         BackColor       =   &H00E1FFFD&
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
         Left            =   10680
         MaxLength       =   2
         TabIndex        =   34
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato15 
         BackColor       =   &H00E1FFFD&
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
         Left            =   10320
         MaxLength       =   2
         TabIndex        =   33
         Tag             =   "fechavencimiento"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   11760
         MaxLength       =   15
         TabIndex        =   32
         Tag             =   "monto"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox dato19 
         BackColor       =   &H00E1FFFD&
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
         Left            =   13320
         MaxLength       =   1
         TabIndex        =   31
         Tag             =   "dh"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   2040
         MaxLength       =   9
         TabIndex        =   30
         Tag             =   "rutctacte"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox dato10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   3120
         MaxLength       =   1
         TabIndex        =   29
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox dato11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   28
         Tag             =   "centrocosto"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox dato14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "numerodocumento"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox NUMEROLINEA 
         BackColor       =   &H00E1FFFD&
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
         Left            =   120
         MaxLength       =   3
         TabIndex        =   26
         Tag             =   "linea"
         Text            =   "001"
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   1455
         Left            =   0
         Top             =   0
         Width           =   13815
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   58
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   57
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CUENTA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   56
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GLOSA CONTABLE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   55
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TD"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8640
         TabIndex        =   54
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9000
         TabIndex        =   53
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MONTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11760
         TabIndex        =   52
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D/H"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   13320
         TabIndex        =   51
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VENCIMIENTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10320
         TabIndex        =   50
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUT"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   49
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CRCC"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   48
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DAF9FE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE CUENTA CONTABLE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DAF9FE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE RUT CUENTA CORRIENTE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   46
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label nombremayor 
         BackColor       =   &H00FFFDF2&
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
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DAF9FE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE CENTRO DE COSTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9600
         TabIndex        =   44
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label nombrecuentacorriente 
         BackColor       =   &H00FFFDF2&
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
         Height          =   255
         Left            =   4920
         TabIndex        =   43
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label nombrecentrocosto 
         BackColor       =   &H00FFFDF2&
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
         Height          =   255
         Left            =   9600
         TabIndex        =   42
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NL"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame comprocabeza 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   13815
      Begin VB.TextBox dato2 
         BackColor       =   &H00E1FFFD&
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "numero"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   600
         MaxLength       =   2
         TabIndex        =   19
         Tag             =   "tipo"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00E1FFFD&
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   18
         Tag             =   "fecha"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00E1FFFD&
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   7080
         MaxLength       =   2
         TabIndex        =   17
         Tag             =   "fecha"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox dato5 
         BackColor       =   &H00E1FFFD&
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "fecha"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(F2) Busqueda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12240
         TabIndex        =   79
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label nombrecomprobante 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   960
         TabIndex        =   61
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO   :"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
         Height          =   255
         Left            =   9960
         TabIndex        =   22
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   13815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA EMISION"
         Height          =   255
         Left            =   5400
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   735
      Left            =   7560
      TabIndex        =   5
      Top             =   8760
      Width           =   6015
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEBE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   13
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   11
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HABER"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALDO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label debe 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label saldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label haber 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame comprocuerpo 
      BackColor       =   &H00FFF2F7&
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   13815
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilladocumentoS 
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   9340
         _Version        =   393216
         BackColor       =   16776436
         ForeColor       =   4194304
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   12640511
         BackColorSel    =   16777215
         ForeColorSel    =   16744576
         BackColorBkg    =   16776690
         GridColor       =   -2147483635
         GridColorFixed  =   12582912
         GridLinesFixed  =   1
         Appearance      =   0
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
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   4080
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   615
      Left            =   240
      Top             =   360
      Width           =   13815
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   0
      TabIndex        =   59
      Top             =   8400
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "\\eltitxp\contabilidad 2005\barra_opciones.swf"
      Src             =   "\\eltitxp\contabilidad 2005\barra_opciones.swf"
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
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   5535
      Left            =   240
      Top             =   2880
      Width           =   13815
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   1455
      Left            =   240
      Top             =   1200
      Width           =   13815
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   735
      Left            =   7680
      Top             =   8880
      Width           =   6015
   End
End
Attribute VB_Name = "ingreso01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub agregalinea_Click()
comprodatos.Enabled = True
opcionmodifica.Visible = False

dato6.Enabled = True
dato6.SetFocus
End Sub

Private Sub dato1_Change()
If Val(dato1.text) < 1 Or Val(dato1.text) > CANDO Then dato1.Enabled = True: dato1.text = "": dato1.SetFocus
End Sub

Private Sub dato1_LostFocus()
TIPOS.Visible = False

End Sub
Private Sub DATO1_GotFocus()

limpia
limpia2
Evento.Visible = False

Call cargatexto(dato1)
TIPOS.Visible = True
End Sub

Private Sub dato19_Change()
If dato19.text <> "D" And dato19.text <> "H" Then dato19.text = ""
End Sub

Private Sub dato19_GotFocus()
If dato19.text = "H" And modifi = 0 Then dato19.text = "D": GoTo PASO:
If dato19.text = "D" And modifi = 0 Then dato19.text = "H": GoTo PASO:
PASO:
Call cargatexto(dato19)
End Sub

Private Sub dato2_GotFocus()

Call cargatexto(dato2)
If Val(dato1.text) < 1 Or Val(dato1.text) > CANDO Then dato1.text = "": dato1.SetFocus: GoTo NO:
nombrecomprobante.Caption = DOCU(Val(dato1.text))
NO:
End Sub

Private Sub dato3_GotFocus()

If modifi = 0 Then LEERMOVIMIENTOS

If VARIPASO <> "0" Then opciones.Visible = True: comprodatos.Enabled = False: comprocabeza.Enabled = False: comprocuerpo.Enabled = False: opciones.SetFocus: limpia2: CREANDO = "": GoTo NO:
CREANDO = "S": grilladocumentoS.Enabled = True

If Val(dato2.text) = 0 Then dato2.text = "": dato2.Enabled = True: dato2.SetFocus
Call cargatexto(dato3)
NO:
End Sub
Private Sub dato4_GotFocus()
If dato3.text = "00" Then dato4.Enabled = True: dato5.Enabled = True: dato6.Enabled = True: dato3.text = Mid(fechasistema, 1, 2): dato4.text = Mid(fechasistema, 4, 2): dato5.text = Mid(fechasistema, 7, 4): dato6.SetFocus
Call cargatexto(dato4)
End Sub

Private Sub dato5_GotFocus()
Call cargatexto(dato5)
End Sub




Private Sub dato6_GotFocus()

Call cargatexto(dato6)
Call ESFECHA(Val(dato3.text), Val(dato4.text), Val(dato5.text))

If VARIPASO = "N" Then dato3.text = "": dato4.text = "": dato5.text = "": dato3.SetFocus

End Sub


Private Sub dato7_GotFocus()
Call cargatexto(dato7)
End Sub
Private Sub dato8_GotFocus()
Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()
Call cargatexto(dato9)
PIVOTE.text = dato6.text + dato7.text + dato8.text
leemayor
If tipocuenta = "00" Then dato9.text = "": dato10.Enabled = True: dato10.text = "": dato11.Enabled = True: dato11.SetFocus
End Sub

Private Sub dato10_GotFocus()
Call cargatexto(dato10)
   If tipocuenta = "00" Then dato10.Enabled = True: dato11.Enabled = True: dato11.SetFocus

End Sub

Private Sub dato11_GotFocus()
Call cargatexto(dato11)
If tipocuenta <> "00" Then pivote2.text = dato9.text + dato10.text: leectacte: GoTo NO:
If tipocentro <> "2" Then dato11.text = "": dato12.Enabled = True: dato12.SetFocus
NO:
If tipocuenta = "00" Then leemayor
End Sub

Private Sub dato12_GotFocus()
Call cargatexto(dato12)
If tipocentro = "2" Then leecrcc
If tipocentro <> "2" Then dato11.text = ""
Rem If tipocuenta <> "00" Then leectacte
If tipocentro <> "2" And tipocuenta = "00" Then leemayor
End Sub
Private Sub dato13_GotFocus()
Call cargatexto(dato13)
End Sub

Private Sub dato14_GotFocus()
Call cargatexto(dato14)
leetipos
End Sub
Private Sub dato15_GotFocus()
Call cargatexto(dato15)
End Sub
Private Sub dato16_GotFocus()
If dato15.text = "00" Then dato18.Enabled = True: dato16.Enabled = True: dato17.Enabled = True: dato15.text = Mid(fechasistema, 1, 2): dato16.text = Mid(fechasistema, 4, 2): dato17.text = Mid(fechasistema, 7, 4): dato18.SetFocus
Call cargatexto(dato16)
End Sub
Private Sub dato17_GotFocus()
Call cargatexto(dato17)
End Sub

Private Sub dato18_GotFocus()
Call cargatexto(dato18)
Call ESFECHA(Val(dato15.text), Val(dato16.text), Val(dato17.text))
If VARIPASO = "N" Then dato15.text = "": dato16.text = "": dato17.text = "": dato15.SetFocus

End Sub


Sub ESFECHA(ByRef DIAS As Integer, ByRef mes As Integer, ByRef ANO As Integer)
If DIAS < 1 Or DIAS > 31 Then VARIPASO = "N": GoTo NO
If mes < 1 Or mes > 12 Then VARIPASO = "N": GoTo NO
If ANO < 2005 Then VARIPASO = "N": GoTo NO:
VARIPASO = "S"
NO:

End Sub


Private Sub eliminacomprobante_Click()

opcionelimina.Visible = False
eliminatodo.Visible = True
Evento.Visible = True
textoevento.Caption = "PROCESO DE ELIMINACION DEL COMPROBANTE COMPLETO"


End Sub

Private Sub eliminalinea_Click()
comprocuerpo.Enabled = True
grilladocumentoS.Enabled = True
Evento.Visible = True
textoevento.Caption = "PROCESO DE ELIMINACION SELECCIONE UNA LINEA Y DOBLE CLICK"
opcionelimina.Visible = False
End Sub

Private Sub Form_Load()
    
    
    sc = 0
    opciones.Visible = False
DOCU(1) = "INGRESOS"
DOCU(2) = "EGRESOS "
DOCU(3) = "TRASPASOS"
DOCU(4) = "TIPO4 "
DOCU(5) = "TIPO5 "



CANDO = 3
For K = 1 To CANDO
LISTATIPOS.AddItem (Str$(K) + "=" + DOCU(K))
Next K
PLANTILLA
grilladocumentoS.Enabled = False
fechasistema = Date
dia = Mid(Date, 1, 2)
mes = Mid(Date, 4, 2)
año = Mid(Date, 7, 4)
CREANDO = ""
limpia2
End Sub



Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo NO:
    Call flechas(dato1, dato2, KeyCode)
NO:
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub
 Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, dato5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudamayor(dato9)
    Call flechas(dato5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato7, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudactacte(dato11)
    Call flechas(dato8, dato10, KeyCode)
End Sub
Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato9, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then Call ayudacrcc(dato12)
    Call flechas(dato10, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)

    Call flechas(dato6, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudatipos(dato14)
    Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato13, dato15, KeyCode)
End Sub
Private Sub dato15_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato14, dato16, KeyCode)
End Sub
Private Sub dato16_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato15, dato17, KeyCode)
End Sub
Private Sub dato17_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato16, dato18, KeyCode)
End Sub
Private Sub dato18_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato17, dato19, KeyCode)
End Sub
Private Sub dato19_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(dato18, dato19, KeyCode)
End Sub


Private Sub DATO1_KeyPress(KeyAscii As Integer)
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
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato4): Call Pregunta(dato4, dato5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato5): Call Pregunta(dato5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 42 And SUMADEBE = SUMAHABER Then retorno: dato3.Enabled = True: dato3.SetFocus
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato6): Call Pregunta(dato6, dato7)
NO:
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato7): Call Pregunta(dato7, dato8)
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato8): Call Pregunta(dato8, dato9)
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato9): Call Pregunta(dato9, dato10)
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato10): Call Pregunta(dato10, dato11)
End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato11): Call Pregunta(dato11, dato12)
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato12, dato13)
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato13, dato14)
End Sub
Private Sub dato14_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato14): Call Pregunta(dato14, dato15)
End Sub
Private Sub dato15_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato15): Call Pregunta(dato15, dato16)
End Sub
Private Sub dato16_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato16): Call Pregunta(dato16, dato17)
End Sub
Private Sub dato17_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato17): Call Pregunta(dato17, dato18)
End Sub
Private Sub dato18_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato18.text) <> 0 Then Call formato(dato18, 0): Call Pregunta(dato18, dato19)
End Sub

Private Sub dato19_KeyPress(KeyAscii As Integer)
      
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      
      If KeyAscii = 13 And dato19.text = "D" Then GRABAR2
      If KeyAscii = 13 And dato19.text = "H" Then GRABAR2
End Sub

Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub

Sub leersiguiente()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'UNIDAD MEDIDA
    campos(3, 0) = dato4.Tag 'SECCION
    campos(4, 0) = dato5.Tag 'DEPARTAMENTO
    campos(5, 0) = dato6.Tag 'LINEA
    campos(6, 0) = dato7.Tag 'IMPUESTO
    campos(7, 0) = dato8.Tag 'PRECIO COSTO IVA
    campos(8, 0) = dato9.Tag 'PRECIO VENTA MAYOR
    campos(9, 0) = dato10.Tag 'PRECIO VENTA DETALLE
    campos(10, 0) = dato11.Tag 'STOCK CRITICO
    campos(11, 0) = dato12.Tag 'DESCUENTO
    campos(12, 0) = dato13.Tag 'DATO EXTRA
    campos(13, 0) = dato14.Tag 'UBICACION
    campos(0, 2) = "maestroproductos"
    condicion = "codigoproducto>" + "'" + dato1.text + "' order by codigoproducto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'UNIDAD MEDIDA
    campos(3, 0) = dato4.Tag 'SECCION
    campos(4, 0) = dato5.Tag 'DEPARTAMENTO
    campos(5, 0) = dato6.Tag 'LINEA
    campos(6, 0) = dato7.Tag 'IMPUESTO
    campos(7, 0) = dato8.Tag 'PRECIO COSTO IVA
    campos(8, 0) = dato9.Tag 'PRECIO VENTA MAYOR
    campos(9, 0) = dato10.Tag 'PRECIO VENTA DETALLE
    campos(10, 0) = dato11.Tag 'STOCK CRITICO
    campos(11, 0) = dato12.Tag 'DESCUENTO
    campos(12, 0) = dato13.Tag 'DATO EXTRA
    campos(13, 0) = dato14.Tag 'UBICACION
    campos(0, 2) = "maestroproductos"
    condicion = "codigoproducto<" + "'" + dato1.text + "' order by codigoproducto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = SQLUTIL.datos(0, 3)
    dato2.text = SQLUTIL.datos(1, 3)
    dato3.text = SQLUTIL.datos(2, 3)
    dato4.text = SQLUTIL.datos(3, 3)
    dato5.text = SQLUTIL.datos(4, 3)
    dato6.text = SQLUTIL.datos(5, 3)
    dato7.text = SQLUTIL.datos(6, 3)
    dato8.text = SQLUTIL.datos(7, 3)
    dato9.text = SQLUTIL.datos(8, 3)
    dato10.text = SQLUTIL.datos(9, 3)
    dato11.text = SQLUTIL.datos(10, 3)
    dato12.text = SQLUTIL.datos(11, 3)
    dato13.text = SQLUTIL.datos(12, 3)
    dato14.text = SQLUTIL.datos(13, 3)

fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    dato5.Locked = condicion
    dato6.Locked = condicion
    dato7.Locked = condicion
    dato8.Locked = condicion
    dato9.Locked = condicion
    dato10.Locked = condicion
    dato11.Locked = condicion
    dato12.Locked = condicion
    dato13.Locked = condicion
    dato14.Locked = condicion
    dato15.Locked = condicion
    dato16.Locked = condicion
    dato17.Locked = condicion
    dato18.Locked = condicion
    dato19.Locked = condicion
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    dato5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    dato9.Enabled = condicion
    dato10.Enabled = condicion
    dato11.Enabled = condicion
    dato12.Enabled = condicion
    dato13.Enabled = condicion
    dato14.Enabled = condicion
    dato15.Enabled = condicion
    dato16.Enabled = condicion
    dato17.Enabled = condicion
    dato18.Enabled = condicion
    dato19.Enabled = condicion
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus: caja.SelStart = 0
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus: caja.SelStart = 0
End Sub
Sub GRABAR()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato6.Tag
    campos(4, 0) = dato9.Tag
    campos(5, 0) = dato11.Tag
    campos(6, 0) = dato12.Tag
    campos(7, 0) = dato13.Tag
    campos(8, 0) = dato14.Tag
    campos(9, 0) = dato15.Tag
    campos(10, 0) = dato18.Tag
    campos(11, 0) = dato19.Tag
    campos(12, 0) = NUMEROLINEA.Tag
    campos(13, 0) = "tipoctacte"
    campos(14, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = dato5.text + dato4.text + dato3.text
    campos(3, 1) = dato6.text + dato7.text + dato8.text
    campos(4, 1) = dato9.text + dato10.text
    campos(5, 1) = dato11.text
    campos(6, 1) = dato12.text
    campos(7, 1) = dato13.text
    campos(8, 1) = dato14.text
    campos(9, 1) = dato17.text + dato16.text + dato15.text
    campos(10, 1) = Replace(dato18.text, ".", "")
    campos(11, 1) = dato19.text
    campos(12, 1) = NUMEROLINEA.text
    campos(13, 1) = "00"

    campos(0, 2) = "movimientoscontables"
    If modifi = 1 Then condicion = "TIPO=" + "'" + dato1.text + "' AND NUMERO=" + "'" + dato2.text + "' AND LINEA=" + "'" + NUMEROLINEA.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.ESTADO

    actualizamayor

End Sub
Sub GRABAR2()
GRABAR
LEERMOVIMIENTOS
If modifi = 1 And VARIPASO <> "0" And CREANDO <> "S" Then opciones.Visible = True: comprodatos.Enabled = False: comprocabeza.Enabled = False: comprocuerpo.Enabled = False: opciones.SetFocus: Evento.Visible = False: GoTo NO:
disponible (False)
dato6.Enabled = True
dato6.SetFocus
modifi = 0
NO:
End Sub
Sub ELIMINAR()

    campos(0, 2) = "movimientoscontables"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "' order by linea"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    eliminagrilla
    

End Sub


Private Sub grilladocumentoS_DblClick()
If CONTROLGRILLA = "S" Then GoTo NO:
If CREANDO = "S" And modifi = 0 And Val(NUMEROLINEA.text) > 1 Then modifi = 1: CARGADATOSAMODIFICAR: dato6.SetFocus: GoTo NO:
If CREANDO = "S" And modifi = 0 And Val(NUMEROLINEA.text) = 1 Then GoTo NO:
If modifi = 1 Then GoTo MODIFICADOR
opcioneliminalinea.Visible = True
NUMEROLINEA.text = grilladocumentoS.TextMatrix(grilladocumentoS.Row, 0)
Call ceros(NUMEROLINEA)
opcioneliminalinea.Caption = "ELIMINARA LA LINEA " + NUMEROLINEA.text
CARGADATOSAELIMINAR
GoTo NO:
MODIFICADOR:
comprodatos.Enabled = True
CARGADATOSAMODIFICAR
textoevento.Caption = "EL SISTEMA ESTA MODIFICANDO LA LINEA " + NUMEROLINEA.text
NO:

End Sub

Private Sub grilladocumentoS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Evento.Visible = False: opciones.SetFocus: grilladocumentoS.Enabled = False




End Sub

Private Sub LISTATIPOS_DblClick()
dato1.text = Mid(LISTATIPOS.text, 2, 1)
End Sub

Private Sub modificacabeza_Click()
opcionmodifica.Visible = False
modifi = 1
comprocabeza.Enabled = True
dato1.SetFocus

End Sub

Private Sub modificalinea_Click()
comprocuerpo.Enabled = True
grilladocumentoS.Enabled = True
Evento.Visible = True
textoevento.Caption = "PROCESO DE MODIFICACION SELECCIONE UNA LINEA Y DOBLE CLICK"
modifi = 1
opcionmodifica.Visible = False

End Sub

Private Sub noelimina_Click()
opcionelimina.Visible = False

End Sub

Private Sub noeliminalinea_Click()
opcioneliminalinea.Visible = False
Evento.Visible = False
grilladocumentoS.Enabled = False: opciones.SetFocus

End Sub

Private Sub noeliminatodo_Click()
eliminatodo.Visible = False
Evento.Visible = False
grilladocumentoS.Enabled = False: opciones.SetFocus

End Sub

Private Sub nomodifica_Click()
opcionmodifica.Visible = False

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
modifi = 0
If command = "retorno" Then retorno
If command = "modifica" Then opcionmodifica.Visible = True
If command = "elimina" Then opcionelimina.Visible = True
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
End Sub
Sub retorno()
If SUMADEBE <> SUMAHABER Then MENSAJE ("SUMA DEL DEBE CON EL HABER DEBEN SER IGUALES"): GoTo NO:
comprocabeza.Enabled = True
comprodatos.Enabled = True
comprocuerpo.Enabled = True
opcionelimina.Visible = False
opcionmodifica.Visible = False

disponible (False): habilita (False): opciones.Visible = False: PLANTILLA: dato1.Locked = False: dato1.Enabled = True: dato1.SetFocus
NO:
End Sub
Sub limpia()


    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
GoTo NO:
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
    dato15.text = ""
    dato16.text = ""
    dato17.text = ""
    dato18.text = ""
    dato19.text = ""
NO:
End Sub
Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("10n", "40s")
    cfijo = "no"
    Call cargaAyudaT("eltitxp", "conta01", "root", "123", "cuentasdelmayor", PIVOTE, campos, cfijo, largo, 2)
      
    If Val(PIVOTE.text) = 0 Then dato15.SetFocus: GoTo NO
    dato16.Enabled = True
    dato17.Enabled = True
    dato15.text = Mid(PIVOTE.text, 1, 2)
    dato16.text = Mid(PIVOTE.text, 3, 2)
    dato17.text = Mid(PIVOTE.text, 5, 4)
 
    caja.Enabled = True
    caja.SetFocus

NO:
End Sub
Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    
    largo = Array("12n", "40s")
    cfijo = "tipo='" & tipocuenta & "'"
    Call cargaAyudaT("eltitxp", "conta01", "root", "123", "cuentascorrientes", pivote2, campos, cfijo, largo, 2)
     
    If Val(pivote2.text) = 0 Then dato9.SetFocus: GoTo NO
    dato10.Enabled = True
    dato9.text = pivote2.text
    dato10.text = Mid(pivote2.text, 10, 1)
    leectacte
    caja.Enabled = True
    caja.SetFocus

NO:

End Sub
Sub ayudatipos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("tipos", "nombredocumento")
    largo = Array("4s", "40s")
    cfijo = "no"
    Call cargaAyudaT("eltitxp", "conta00", "root", "123", "maestrotipodedocumentos", dato13, campos, cfijo, largo, 2)
    If dato13.text = "" Then dato13.SetFocus: GoTo NO
    leetipos
    caja.Enabled = True
    caja.SetFocus


NO:

End Sub

Sub ayudacrcc(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("5s", "50s")
    cfijo = "no"
    Call cargaAyudaT("eltitxp", "conta01", "root", "123", "centrosdecosto", dato11, campos, cfijo, largo, 2)
      
    If dato11.text = "" Then dato11.SetFocus: GoTo NO
    leecrcc
    caja.Enabled = True
    caja.SetFocus
NO:
End Sub

Sub grilla(pasada As ListBox)
palabra = ""


For K = 1 To cancolu
If tipodato(K) = "s" Or tipodato(K) = "S" Then dato(K) = dato(K) & String(colu(K) - Len(dato(K)), 32)
If tipodato(K) = "n" Or tipodato(K) = "N" Then dato(K) = String(colu(K) - Len(dato(K)), 32) & dato(K)

palabra = palabra & dato(K)
Next K
pasada.AddItem (palabra)
End Sub
Sub calcula()





End Sub
Sub suma()

End Sub
Sub PLANTILLA()
grilladocumentoS.Clear
grilladocumentoS.Cols = 10
grilladocumentoS.Rows = 2
grilladocumentoS.ColWidth(0) = 120 * 6
grilladocumentoS.ColWidth(1) = 120 * 11
grilladocumentoS.ColWidth(2) = 120 * 11
grilladocumentoS.ColWidth(3) = 120 * 5
grilladocumentoS.ColWidth(4) = 120 * 39
grilladocumentoS.ColWidth(5) = 120 * 4
grilladocumentoS.ColWidth(6) = 120 * 10
grilladocumentoS.ColWidth(7) = 120 * 12
grilladocumentoS.ColWidth(8) = 120 * 10
grilladocumentoS.ColWidth(9) = 120 * 2

Rem TITULOS
grilladocumentoS.TextMatrix(0, 0) = "LINEA"
grilladocumentoS.TextMatrix(0, 1) = "CUENTA"
grilladocumentoS.TextMatrix(0, 2) = "RUT"
grilladocumentoS.TextMatrix(0, 3) = "CRCC"
grilladocumentoS.TextMatrix(0, 4) = "GLOSA"
grilladocumentoS.TextMatrix(0, 5) = "TD"
grilladocumentoS.TextMatrix(0, 6) = "NUMERO"
grilladocumentoS.TextMatrix(0, 7) = "VENCIMIENTO"
grilladocumentoS.TextMatrix(0, 8) = "MONTO"
grilladocumentoS.TextMatrix(0, 9) = "D/H"


End Sub

Sub LEERMOVIMIENTOS()

PLANTILLA

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT linea,codigocuenta,rutctacte,centrocosto,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh,fecha "
        cSql.SQL = cSql.SQL + "FROM movimientoscontables"
        cSql.SQL = cSql.SQL + " where tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "' order by linea"
                ' cSql.SQL = cSql.SQL + " where tipo=1 and numero=0000000005 order by linea"
        cSql.Execute
        
        linea = 0: SUMADOR = 0
        If cSql.RowsAffected > 0 Then
            
            
            Set resultados = cSql.OpenResultset
            SUMADEBE = 0
            SALDOS = 0
            SUMAHABER = 0
            While Not resultados.EOF
                linea = linea + 1
                grilladocumentoS.Rows = linea + 1
                For K = 0 To 9
                grilladocumentoS.TextMatrix(linea, K) = resultados(K)
                Next K
                grilladocumentoS.TextMatrix(linea, 1) = Mid(resultados(1), 1, 2) + "." + Mid(resultados(1), 3, 2) + "." + Mid(resultados(1), 5, 4)
                grilladocumentoS.TextMatrix(linea, 8) = Format(resultados(8), "#,###,###,##0")
                NUMEROLINEA.text = resultados(0)
                dato3.Enabled = True
                dato4.Enabled = True
                dato5.Enabled = True
                dato3.text = Mid(resultados(10), 1, 2)
                dato4.text = Mid(resultados(10), 4, 2)
                dato5.text = Mid(resultados(10), 7, 4)
                If resultados(9) = "D" Then SUMADEBE = SUMADEBE + CDbl(resultados(8))
                If resultados(9) = "H" Then SUMAHABER = SUMAHABER + CDbl(resultados(8))
                reordenalineas
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing

        End If
    VARIPASO = cSql.RowsAffected
    NUMEROLINEA.text = CDbl(linea) + 1
    Call ceros(NUMEROLINEA)
    If VARIPASO <> "0" Then CARGADATAFIELD
    saldo = SUMADEBE - SUMAHABER
    debe.Caption = Format(SUMADEBE, "##,###,###,##0")
    haber.Caption = Format(SUMAHABER, "##,###,###,##0")
    saldo.Caption = Format(saldo, "##,###,###,##0")
    End With

End Sub

Sub reordenalineas()
If Val(NUMEROLINEA.text) = linea Then GoTo NO:

    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = NUMEROLINEA.Tag
    campos(3, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = NUMEROLINEA.text
    campos(13, 1) = "00"

    campos(0, 2) = "movimientoscontables"
    If modifi = 1 Then condicion = "TIPO=" + "'" + dato1.text + "' AND NUMERO=" + "'" + dato2.text + "' AND LINEA=" + "'" + NUMEROLINEA.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.ESTADO

    actualizamayor


NO:
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus: caja.SelStart = 0
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub leemayor()
    PIVOTE.text = dato6.text + dato7.text + dato8.text
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "ctacte"
    campos(3, 0) = "centrocosto"
    campos(4, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + PIVOTE.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then dato6.text = "": dato7.text = "": dato8.text = "": dato6.SetFocus: GoTo NO:
    nombremayor.Caption = SQLUTIL.datos(1, 3)
    tipocuenta = SQLUTIL.datos(2, 3)
    tipocentro = SQLUTIL.datos(3, 3)

NO:

End Sub
Sub leectacte()
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + tipocuenta + "' and rut=" + "'" + pivote2.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then dato9.text = "": dato10.text = "": dato9.SetFocus: GoTo NO:
    nombrecuentacorriente.Caption = SQLUTIL.datos(1, 3)
    dato12.Enabled = True
    dato12.SetFocus
NO:

End Sub


Sub leecrcc()
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + dato11.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    If SQLUTIL.ESTADO = 4 Then dato11.text = "": dato11.SetFocus:  GoTo NO:
    VARIPASO = "S"
    nombrecentrocosto.Caption = SQLUTIL.datos(1, 3)

NO:

End Sub
Sub leetipos()
    
    campos(0, 0) = "tipos"
    campos(1, 0) = "nombredocumento"
    campos(2, 0) = ""
    
    campos(0, 2) = "maestrotipodedocumentos"
    condicion = "tipos=" + "'" + dato13.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta00
    Call SQLUTIL.SQLUTIL(op, condicion)

    If SQLUTIL.ESTADO = 4 Then dato13.text = "": dato13.SetFocus:  GoTo NO:
    VARIPASO = "S"
    

NO:

End Sub
Sub cargatexto(ByRef caja As TextBox)
Rem If caja.text = "" Then caja.text = String(caja.MaxLength, "")

caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub


Private Sub sieliminalinea_Click()
eliminandolinea
opcioneliminalinea.Visible = False



Evento.Visible = False

End Sub
Sub eliminandolinea()
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and linea=" + "'" + NUMEROLINEA.text + "' order by linea"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    cargadato
    desactualizamayor
    LEERMOVIMIENTOS
    
End Sub

Private Sub sieliminatodo_Click()
eliminatodo.Visible = False
Evento.Visible = False

ELIMINAR

retorno
End Sub
Sub CARGADATOSAELIMINAR()
ELI.Clear
ELI.Cols = 2
ELI.Rows = 10
ELI.ColWidth(0) = 120 * 12
ELI.ColWidth(1) = 120 * 20

Rem TITULOS
For VARINUM = 0 To 9
K = grilladocumentoS.Row
ELI.TextMatrix(VARINUM, 0) = grilladocumentoS.TextMatrix(0, VARINUM)
ELI.TextMatrix(VARINUM, 1) = grilladocumentoS.TextMatrix(K, VARINUM)
Next VARINUM

End Sub
Sub CARGADATOSAMODIFICAR()
Rem TITULOS

K = grilladocumentoS.Row
habilita (False)
dato6.Enabled = True
dato7.Enabled = True
dato8.Enabled = True
dato9.Enabled = True
dato10.Enabled = True
dato11.Enabled = True
dato12.Enabled = True
dato13.Enabled = True
dato14.Enabled = True
dato15.Enabled = True
dato16.Enabled = True
dato17.Enabled = True
dato18.Enabled = True
dato19.Enabled = True
NUMEROLINEA.text = grilladocumentoS.TextMatrix(K, 0)
cargadato
desactualizamayor
modifi = 1

dato6.SetFocus

End Sub
Sub CARGADATAFIELD()
K = Val(NUMEROLINEA.text) - 1
cargadato

End Sub
Sub limpia2()



    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
    dato15.text = ""
    dato16.text = ""
    dato17.text = ""
    dato18.text = ""
    dato19.text = ""
End Sub

Sub actualizamayor()


    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If dato19.text = "D" Then campos(2, 0) = "debe" + mes
    If dato19.text = "H" Then campos(2, 0) = "haber" + mes
    campos(3, 0) = ""
    
    condicion = "codigo=" + "'" + dato6.text + dato7.text + dato8.text + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    
    If SQLUTIL.ESTADO = 4 Then Stop
    
    VARIPASO = Replace(dato18.text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto + Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
    

    
End Sub
Sub desactualizamayor()


    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If dato19.text = "D" Then campos(2, 0) = "debe" + mes
    If dato19.text = "H" Then campos(2, 0) = "haber" + mes
    campos(3, 0) = ""
    
    condicion = "codigo=" + "'" + dato6.text + dato7.text + dato8.text + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop:
    
    VARIPASO = Replace(dato18.text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto - Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
NO:


    
End Sub


Sub eliminagrilla()
For K = 1 To Val(NUMEROLINEA.text) - 1
cargadato
desactualizamayor

Next K
End Sub
Sub cargadato()


dato6.text = Mid(grilladocumentoS.TextMatrix(K, 1), 1, 2)
dato7.text = Mid(grilladocumentoS.TextMatrix(K, 1), 4, 2)
dato8.text = Mid(grilladocumentoS.TextMatrix(K, 1), 7, 4)
dato9.text = Mid(grilladocumentoS.TextMatrix(K, 2), 1, 9)
dato10.text = Mid(grilladocumentoS.TextMatrix(K, 2), 10, 1)
dato11.text = grilladocumentoS.TextMatrix(K, 3)
dato12.text = grilladocumentoS.TextMatrix(K, 4)
dato13.text = grilladocumentoS.TextMatrix(K, 5)
dato14.text = grilladocumentoS.TextMatrix(K, 6)
dato15.text = Mid(grilladocumentoS.TextMatrix(K, 7), 1, 2)
dato16.text = Mid(grilladocumentoS.TextMatrix(K, 7), 4, 2)
dato17.text = Mid(grilladocumentoS.TextMatrix(K, 7), 7, 4)
dato18.text = grilladocumentoS.TextMatrix(K, 8)
dato19.text = grilladocumentoS.TextMatrix(K, 9)
End Sub

