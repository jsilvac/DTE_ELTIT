VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8b.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form infoge03 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso Libro de Compras"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   1230
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   15045
   Begin VB.TextBox LINEAS 
      Height          =   285
      Left            =   5760
      MaxLength       =   3
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame TIPOS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2040
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRILLATIPO 
         Height          =   975
         Left            =   120
         TabIndex        =   65
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16107953
         ForeColor       =   16711680
         Rows            =   3
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   16777152
         BackColorBkg    =   16761024
         GridColor       =   16744576
         GridColorFixed  =   14282751
         GridColorUnpopulated=   14282751
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00808080&
         Height          =   1215
         Left            =   0
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Frame comprodatos 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   1335
      Left            =   5760
      TabIndex        =   37
      Top             =   240
      Width           =   9135
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
         Left            =   720
         MaxLength       =   2
         TabIndex        =   15
         Tag             =   "codigocuenta"
         Top             =   360
         Width           =   375
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
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   16
         Top             =   360
         Width           =   375
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
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   17
         Top             =   360
         Width           =   615
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
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "glosacontable"
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox dato20 
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
         Left            =   6960
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "monto"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox dato21 
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
         Left            =   8520
         MaxLength       =   1
         TabIndex        =   21
         Tag             =   "dh"
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
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   18
         Tag             =   "centrocosto"
         Top             =   360
         Width           =   615
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
         TabIndex        =   38
         Tag             =   "linea"
         Text            =   "001"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DAF9FE&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE CUENTA CONTABLE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label nombremayor 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DAF9FE&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE CENTRO DE COSTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         TabIndex        =   62
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label nombrecentrocosto 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         TabIndex        =   61
         Top             =   960
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   1335
         Left            =   0
         Top             =   0
         Width           =   9135
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
         TabIndex        =   45
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CUENTA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   44
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GLOSA CONTABLE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   43
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MONTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6960
         TabIndex        =   42
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D/H"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8520
         TabIndex        =   41
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CRCC"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   40
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NL"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame comprocabeza 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   6615
      Left            =   120
      TabIndex        =   36
      Top             =   240
      Width           =   5415
      Begin VB.TextBox folios 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Left            =   3720
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   72
         Tag             =   "tipo"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox añoconta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   70
         Tag             =   "tipo"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox mesconta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   68
         Tag             =   "tipo"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
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
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   0
         Tag             =   "tipo"
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox dato2 
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "numero"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox comentariofactura 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   5160
         Width           =   5175
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "fecha"
         Top             =   2640
         Width           =   615
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
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "fecha"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox dato6 
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "fechavencimiento"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox total 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   13
         Tag             =   "monto"
         Text            =   "0"
         Top             =   4440
         Width           =   1455
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   12
         Tag             =   "retencion"
         Text            =   "0"
         Top             =   4080
         Width           =   1455
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "exento"
         Text            =   "0"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox dato12 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   10
         Tag             =   "iva"
         Text            =   "0"
         Top             =   3360
         Width           =   1455
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "neto"
         Text            =   "0"
         Top             =   3000
         Width           =   1455
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
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "rut"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox dato3 
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox dato4 
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
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox dato5 
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label DV 
         BackColor       =   &H00DAF9FE&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         TabIndex        =   73
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AÑO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   71
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MES"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   69
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   67
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA EMISION"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VENCIMIENTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIPO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COMENTARIO DE FACTURA DE COMPRA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   4920
         Width           =   5175
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RETENCION"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXENTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I.V.A"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NETO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PROVEEDOR"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label nombreproveedor 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1920
         Width           =   5175
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
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   2040
         TabIndex        =   48
         Top             =   840
         Width           =   1935
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   6615
         Left            =   0
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   735
      Left            =   8880
      TabIndex        =   26
      Top             =   4320
      Width           =   6015
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEBE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FF8080&
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
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HABER"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALDO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame comprocuerpo 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5760
      TabIndex        =   25
      Top             =   1800
      Width           =   9135
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilladocumentoS 
         Height          =   2055
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   16776436
         ForeColor       =   4194304
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   16107953
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
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   9135
      End
   End
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      Height          =   735
      Left            =   9000
      Top             =   4440
      Width           =   5970
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      Height          =   2295
      Left            =   5880
      Top             =   1920
      Width           =   9100
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      Height          =   6615
      Left            =   240
      Top             =   360
      Width           =   5415
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   6000
      TabIndex        =   46
      Top             =   5280
      Width           =   7695
      _cx             =   13573
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
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      Height          =   1335
      Left            =   5880
      Top             =   360
      Width           =   9100
   End
End
Attribute VB_Name = "infoge03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private tipocuenta As String
    Private cc As Integer


Private Sub Command2_Click()

End Sub

Private Sub agregalinea_Click()
comprodatos.Enabled = True
opcionmodifica.Visible = False

dato6.Enabled = True
dato6.SetFocus
End Sub

Private Sub comentariofactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 42 Then dato15.Enabled = True: dato15.SetFocus
End Sub

Private Sub dato1_Change()
If Val(dato1.text) < 1 Or Val(dato1.text) > CANDO Then dato1.Enabled = True: dato1.text = "": dato1.SetFocus
End Sub

Private Sub dato1_LostFocus()
TIPOS.Visible = False
leeFOLIO
End Sub
Private Sub DATO1_GotFocus()

Call cargatexto(dato1)
TIPOS.Visible = True
End Sub

Private Sub dato19_GotFocus()
If tipocentro = 2 Then leecrcc

Call cargatexto(dato19)
End Sub

Private Sub dato2_GotFocus()

Call cargatexto(dato2)
If Val(dato1.text) < 1 Or Val(dato1.text) > CANDO Then dato1.text = "": dato1.SetFocus: GoTo no:
nombrecomprobante.Caption = GRILLATIPO.TextMatrix(Val(dato1.text) - 1, 1)
no:

End Sub

Private Sub dato21_Change()
If dato21.text <> "D" And dato21.text <> "H" Then dato21.text = ""
End Sub

Private Sub dato21_KeyPress(KeyAscii As Integer)
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii = 13 And dato21.text = "D" Then GRABAR2
      If KeyAscii = 13 And dato21.text = "H" Then GRABAR2

End Sub

Private Sub dato3_Change()
If Val(dato3.text) > 31 Then dato3.text = ""
End Sub
Private Sub dato4_Change()
If Val(dato4.text) > 12 Then dato4.text = ""
End Sub
Private Sub dato6_Change()
If Val(dato6.text) > 31 Then dato6.text = ""
End Sub
Private Sub dato7_Change()
If Val(dato7.text) > 12 Then dato7.text = ""
End Sub

Private Sub dato3_GotFocus()
If tipocuenta <> "00" Then dv.Caption = rut(dato9.text): pivote2.text = dato9.text + dv.Caption: leectacte
leefactura
If SQLUTIL.estado = 0 Then carga: LEERMOVIMIENTOS: If VARIPASO <> "0" Then opciones.Visible = True: comprodatos.Enabled = False: comprocabeza.Enabled = False: comprocuerpo.Enabled = False: opciones.SetFocus: GoTo no:





If Val(dato2.text) = 0 Then dato2.text = "": dato2.Enabled = True: dato2.SetFocus
Call cargatexto(dato3)
no:
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
If IsDate(dato3.text + "-" + dato4.text + "-" + dato5.text) = False Then dato3.text = "": dato4.text = "": dato5.text = "": dato3.SetFocus

End Sub


Private Sub dato7_GotFocus()
If dato6.text = "00" Then dato6.Enabled = True: dato7.Enabled = True: dato8.Enabled = True: dato6.text = Mid(fechasistema, 1, 2): dato7.text = Mid(fechasistema, 4, 2): dato8.text = Mid(fechasistema, 7, 4): dato11.Enabled = True: dato11.SetFocus
Call cargatexto(dato7)
End Sub
Private Sub dato8_GotFocus()

Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()

Call cargatexto(dato9)



Call leemayor(cuentaproveedor)

End Sub

Private Sub dato10_GotFocus()
Call cargatexto(dato10)
   If tipocuenta = "00" Then dato10.Enabled = True: dato11.Enabled = True: dato11.SetFocus

End Sub

Private Sub dato11_GotFocus()
If IsDate(dato6.text + "-" + dato7.text + "-" + dato8.text) = False Then dato6.text = "": dato7.text = "": dato8.text = "": dato6.SetFocus



Call cargatexto(dato11)

no:
End Sub

Private Sub dato12_GotFocus()
SUMADOR = Int((CDbl(Replace(dato11.text, ",", "")) * iva / 100) + 0.5)
dato12.text = Format(SUMADOR, "#,###,###,##0")
totalfactura
Call cargatexto(dato12)
End Sub
Private Sub dato13_GotFocus()
totalfactura
Call cargatexto(dato13)
End Sub

Private Sub dato14_GotFocus()
totalfactura
Call cargatexto(dato14)

End Sub
Private Sub dato15_GotFocus()
If Val(NUMEROLINEA.text) = 1 Then grabafactura: LEERMOVIMIENTOS
totalfactura
Call cargatexto(dato15)
End Sub
Private Sub dato16_GotFocus()
Call cargatexto(dato16)
End Sub
Private Sub dato17_GotFocus()
Call cargatexto(dato17)
End Sub

Private Sub dato18_GotFocus()

Call leemayor(dato15.text + dato16.text + dato17.text)
If tipocentro <> 2 Then dato19.Enabled = True: dato19.SetFocus
Call cargatexto(dato18)

End Sub

Private Sub dato20_GotFocus()
Call cargatexto(dato20)

End Sub
Private Sub dato21_GotFocus()
Call cargatexto(dato21)

End Sub




Private Sub Form_Activate()
ingreso02.Left = 0
ingreso02.Top = 0

End Sub

Private Sub Form_Load()



iva = 19
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)
    
   
    sc = 0
    opciones.Visible = False
GRILLATIPOS
PLANTILLA
mesconta.text = Mid(fechasistema, 4, 2)
añoconta.text = Mid(fechasistema, 7, 4)
dia = Mid(fechasistema, 1, 2)
mes = Mid(fechasistema, 4, 2)
año = Mid(fechasistema, 7, 4)


End Sub
Sub GRILLATIPOS()
GRILLATIPO.Cols = 2
GRILLATIPO.Rows = 3
GRILLATIPO.ColWidth(0) = 200 * 2
GRILLATIPO.ColWidth(1) = 200 * 10

GRILLATIPO.TextMatrix(0, 0) = "1"
GRILLATIPO.TextMatrix(1, 0) = "2"
GRILLATIPO.TextMatrix(2, 0) = "3"
GRILLATIPO.TextMatrix(0, 1) = "FACTURA"
GRILLATIPO.TextMatrix(1, 1) = "NOTA DE DEBITO"
GRILLATIPO.TextMatrix(2, 1) = "NOTA DE CREDITO"
CANDO = 3
End Sub


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato9, KeyCode)
End Sub
 Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudactacte(dato3)
    Call flechas(dato2, dato11, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato9, dato4, KeyCode)
End Sub
 
 Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, dato5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato7, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        'If KeyCode = vbKeyF2 Then Call ayudacrcc(dato12)
    Call flechas(dato9, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)

    Call flechas(dato11, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
    ' If KeyCode = vbKeyF2 Then Call ayudatipos(dato14)
    Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato13, dato15, KeyCode)
End Sub
Private Sub dato15_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudamayor(dato19)
    Call flechas(dato14, dato16, KeyCode)
End Sub
Private Sub dato16_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato15, dato17, KeyCode)
End Sub
Private Sub dato17_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato16, dato18, KeyCode)
End Sub
Private Sub dato18_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudacrcc(dato19)
    Call flechas(dato17, dato19, KeyCode)
End Sub
Private Sub dato19_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(dato18, dato19, KeyCode)
End Sub
Private Sub dato20_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(dato18, dato19, KeyCode)
End Sub
Private Sub dato21_KeyDown(KeyCode As Integer, Shift As Integer)
     Call flechas(dato18, dato19, KeyCode)
End Sub


Private Sub DATO1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call Pregunta(dato1, dato2)
End Sub
Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato9)

End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato9): Call Pregunta(dato9, dato3)
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
    ' If KeyAscii = 42 And SUMADEBE = SUMAHABER Then grabarcomprobante:retorno: dato3.Enabled = True: dato3.SetFocus
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato6): Call Pregunta(dato6, dato7)
no:
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato7): Call Pregunta(dato7, dato8)
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato8): Call Pregunta(dato8, dato11)
End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato11.text) <> 0 Then Call formato(dato11, 0): Call Pregunta(dato11, dato12)
    
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato12, 0): Call Pregunta(dato12, dato13)
    
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato13, 0): Call Pregunta(dato13, dato14)

End Sub
Private Sub dato14_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato14, 0): Call Pregunta(dato14, comentariofactura)

End Sub
Private Sub dato15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 42 And sumadebe = sumahaber Then grabarcomprobante: retorno: dato1.Enabled = True: dato1.SetFocus
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
    If KeyAscii = 13 Then Call ceros(dato18):  Call Pregunta(dato18, dato19)
End Sub

Private Sub dato19_KeyPress(KeyAscii As Integer)
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If KeyAscii = 13 Then Call Pregunta(dato19, dato20)
           
End Sub
Private Sub dato20_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato20, 0): Call Pregunta(dato20, dato21)
           
End Sub

Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub


Sub carga()
    habilita (True)
    dato1.text = SQLUTIL.datos(0, 3)
    dato2.text = SQLUTIL.datos(1, 3)
    dato3.text = Mid(SQLUTIL.datos(2, 3), 1, 2)
    dato4.text = Mid(SQLUTIL.datos(2, 3), 4, 2)
    dato5.text = Mid(SQLUTIL.datos(2, 3), 7, 4)
    dato6.text = Mid(SQLUTIL.datos(3, 3), 1, 2)
    dato7.text = Mid(SQLUTIL.datos(3, 3), 4, 2)
    dato8.text = Mid(SQLUTIL.datos(3, 3), 7, 4)
    dato9.text = Mid(SQLUTIL.datos(4, 3), 1, 9)
    dv.Caption = Mid(SQLUTIL.datos(4, 3), 10, 1)
    dato11.text = Format(SQLUTIL.datos(5, 3), "##,###,###,##0")
    dato12.text = Format(SQLUTIL.datos(6, 3), "##,###,###,##0")
    dato13.text = Format(SQLUTIL.datos(7, 3), "##,###,###,##0")
    dato14.text = Format(SQLUTIL.datos(8, 3), "##,###,###,##0")
    total.text = Format(SQLUTIL.datos(9, 3), "##,###,###,##0")
    mescontabilizado = SQLUTIL.datos(10, 3)
    añocontabilizado = SQLUTIL.datos(11, 3)
    comentariofactura.text = SQLUTIL.datos(12, 3)
    
    
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
Sub GRABADETALLEFACTURA()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "cuentadelmayor"
    campos(4, 0) = "glosa"
    campos(5, 0) = "monto"
    campos(6, 0) = "dh"
    campos(7, 0) = "centrodecosto"
    campos(8, 0) = "tipoctacte"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = NUMEROLINEA.text
    campos(3, 1) = dato15.text + dato16.text + dato17.text
    campos(4, 1) = dato19.text
    campos(5, 1) = Replace(dato20.text, ".", "")
    campos(6, 1) = dato21.text
    campos(7, 1) = dato18.text
    campos(8, 1) = ""
    campos(9, 1) = ""
    
    campos(0, 2) = "detallefacturasdecompra"
    If modifi = 1 Then condicion = "TIPO=" + "'" + dato1.text + "' AND NUMERO=" + "'" + dato2.text + "' AND LINEA=" + "'" + NUMEROLINEA.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado


End Sub
Sub grabafactura()
    Dim NETOS As Double
        Dim DH As String
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato6.Tag
    campos(4, 0) = dato9.Tag
    campos(5, 0) = dato11.Tag
    campos(6, 0) = dato12.Tag
    campos(7, 0) = dato13.Tag
    campos(8, 0) = dato14.Tag
    campos(9, 0) = "total"
    campos(10, 0) = "comentario"
    campos(11, 0) = "añocontable"
    campos(12, 0) = "mescontable"
    campos(13, 0) = "folio"
    campos(14, 0) = "tipofactura"
    campos(15, 0) = "tipodocumento"
    
    campos(16, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = dato5.text + dato4.text + dato3.text
    campos(3, 1) = dato8.text + dato7.text + dato6.text
    campos(4, 1) = dato9.text + dv.Caption
    campos(5, 1) = Replace(dato11.text, ".", "")
    campos(6, 1) = Replace(dato12.text, ".", "")
    campos(7, 1) = Replace(dato13.text, ".", "")
    campos(8, 1) = Replace(dato14.text, ".", "")
    campos(9, 1) = Replace(total.text, ".", "")
    campos(10, 1) = comentariofactura.text
    campos(11, 1) = añoconta.text
    campos(12, 1) = mesconta.text
    campos(13, 1) = folios.text
    campos(14, 1) = "N"
    campos(15, 1) = "FA"

    campos(0, 2) = "facturasdecompras"
    If modifi = 1 Then condicion = "TIPO=" + "'" + dato1.text + "' AND NUMERO=" + "'" + dato2.text + "' AND LINEA=" + "'" + NUMEROLINEA.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
Rem graba linea PROVEEDOR
    
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = "linea"
    campos(3, 0) = "cuentadelmayor"
    campos(4, 0) = "glosa"
    campos(5, 0) = "monto"
    campos(6, 0) = "dh"
    campos(7, 0) = "centrodecosto"
    campos(8, 0) = "tipoctacte"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = "001"
    campos(3, 1) = cuentaproveedor
    If dato1.text = "3" Then VARIPASO$ = "ABONA " Else VARIPASO = "CARGA "
    campos(4, 1) = VARIPASO + DOCU$(Val(dato1.text)) + " " + nombreproveedor.Caption
    campos(5, 1) = Str(total.text)
    If dato1.text = "3" Then campos(6, 1) = "D" Else campos(6, 1) = "H"
    campos(7, 1) = ""
    campos(8, 1) = tipocuenta
    campos(9, 1) = dato9.text + dv.Caption
    
    campos(0, 2) = "detallefacturasdecompra"
    If modifi = 1 Then condicion = "TIPO=" + "'" + dato1.text + "' AND NUMERO=" + "'" + dato2.text + "' AND LINEA=" + "'" + NUMEROLINEA.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
    
Rem graba linea I.V.A
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = "linea"
    campos(3, 0) = "cuentadelmayor"
    campos(4, 0) = "glosa"
    campos(5, 0) = "monto"
    campos(6, 0) = "dh"
    campos(7, 0) = "centrodecosto"
    campos(8, 0) = "tipoctacte"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = "002"
    campos(3, 1) = cuentaiva
    campos(4, 1) = "I.V.A DOCUMENTO DE COMPRA"
    campos(5, 1) = Str(dato12.text)
    If dato1.text = "3" Then campos(6, 1) = "H" Else campos(6, 1) = "D"
    
    campos(7, 1) = ""
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(0, 2) = "detallefacturasdecompra"
    If modifi = 1 Then condicion = "TIPO=" + "'" + dato1.text + "' AND NUMERO=" + "'" + dato2.text + "' AND LINEA=" + "'" + NUMEROLINEA.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    K = SQLUTIL.estado
    NUMEROLINEA.text = "003"
    NETOS = CDbl(total.text) - CDbl(dato12.text)
    
    dato15.text = Mid(mercaderias, 1, 2)
    dato16.text = Mid(mercaderias, 3, 2)
    dato17.text = Mid(mercaderias, 5, 4)
    dato18.text = ""
    dato19.text = "MERCADERIAS"
    dato20.text = Str(NETOS)
    If dato1.text = "3" Then dato21.text = "H" Else dato21.text = "D"
    

End Sub


Sub GRABAR2()
GRABADETALLEFACTURA
LEERMOVIMIENTOS
If modifi = 1 And VARIPASO <> "0" Then opciones.Visible = True: comprodatos.Enabled = False: comprocabeza.Enabled = False: comprocuerpo.Enabled = False: opciones.SetFocus:  GoTo no:
disponible (False)
dato15.Enabled = True
dato15.SetFocus
no:
End Sub
Sub ELIMINAR()
    Dim tipocon As String
    Call acepta("DESEA ELIMINAR EL DOCUMENTO")
    If RESPUESTA = "N" Then GoTo no:
    campos(0, 2) = "detallefacturasdecompra"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "' order by linea"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    campos(0, 2) = "facturasdecompras"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and rut=" + "'" + dato9.text + dv.Caption + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    campos(0, 2) = "movimientoscontables"
    If dato1.text = "1" Then tipocon = "FC"
    If dato1.text = "2" Then tipocon = "DC"
    If dato1.text = "3" Then tipocon = "CC"
     
    condicion = "tipo=" + "'" + tipocon + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    eliminacomprobantecontable
    retorno
no:
End Sub


Private Sub glosafactura_Change()

End Sub

Private Sub grilladocumentoS_DblClick()
If modifi = 1 Then GoTo MODIFICADOR
opcioneliminalinea.Visible = True
NUMEROLINEA.text = grilladocumentoS.TextMatrix(grilladocumentoS.row, 0)
Call ceros(NUMEROLINEA)
opcioneliminalinea.Caption = "ELIMINARA LA LINEA " + NUMEROLINEA.text
CARGADATOSAELIMINAR
GoTo no:
MODIFICADOR:
comprodatos.Enabled = True
CARGADATOSAMODIFICAR
textoevento.Caption = "EL SISTEMA ESTA MODIFICANDO LA LINEA " + NUMEROLINEA.text
no:

End Sub

Private Sub grilladocumentoS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then opciones.SetFocus: grilladocumentoS.Enabled = False




End Sub

Private Sub LISTATIPOSCONTROL_DblClick()
dato1.text = Mid(LISTATIPOSCONTROL.text, 2, 1)
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

textoevento.Caption = "PROCESO DE MODIFICACION SELECCIONE UNA LINEA Y DOBLE CLICK"
modifi = 1
opcionmodifica.Visible = False

End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub noelimina_Click()
opcionelimina.Visible = False

End Sub

Private Sub noeliminalinea_Click()
opcioneliminalinea.Visible = False

grilladocumentoS.Enabled = False: opciones.SetFocus

End Sub

Private Sub noeliminatodo_Click()
eliminatodo.Visible = False

grilladocumentoS.Enabled = False: opciones.SetFocus

End Sub

Private Sub nomodifica_Click()
opcionmodifica.Visible = False

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
modifi = 0
If command = "retorno" Then retorno

If command = "elimina" Then ELIMINAR
'If command = "siguiente" Then leersiguiente

'If command = "anterior" Then leeranterior
End Sub
Sub retorno()

comprocabeza.Enabled = True
comprodatos.Enabled = True
comprocuerpo.Enabled = True
limpia

disponible (False): habilita (False): opciones.Visible = False: PLANTILLA: dato1.Locked = False: dato1.Enabled = True: dato1.SetFocus
End Sub
Sub limpia()


    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dv.Caption = ""
    dato11.text = "0"
    dato12.text = "0"
    dato13.text = "0"
    dato14.text = "0"
    dato15.text = ""
    dato16.text = ""
    dato17.text = ""
    dato18.text = ""
    dato19.text = ""
    dato20.text = ""
    dato21.text = ""
    total.text = "0"
    comentariofactura.text = ""
    NUMEROLINEA.text = "001"
no:
End Sub
Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    largo = Array("10n", "40s")
    cfijo = "no"
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato15.SetFocus: GoTo no
    dato16.Enabled = True
    dato17.Enabled = True
    dato15.text = Mid(pivote.text, 1, 2)
    dato16.text = Mid(pivote.text, 3, 2)
    dato17.text = Mid(pivote.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
no:
End Sub
Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    
    largo = Array("12n", "40s")
    cfijo = "tipo='" & tipocuenta & "'"
    cabezas = Array("RUT", "NOMBRE")
    mensajeAyuda = "Ayuda Cuentas Corrientes"

    
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentascorrientes", pivote2, campos, cfijo, largo, 2)
     
    If Val(pivote2.text) = 0 Then dato9.SetFocus: GoTo no
    dato9.text = Mid(pivote2.text, 1, 9)
    dv.Caption = Mid(pivote2.text, 10, 1)
    
    caja.Enabled = True
    caja.SetFocus

no:

End Sub

Sub ayudacrcc(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("5s", "50s")
    cfijo = "no"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Centros de Costo"
    
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "centrosdecosto", dato18, campos, cfijo, largo, 2)
      
    If dato18.text = "" Then dato18.SetFocus: GoTo no
    
    caja.Enabled = True
    caja.SetFocus
no:
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
grilladocumentoS.Cols = 6

grilladocumentoS.Rows = 2
grilladocumentoS.ColWidth(0) = 120 * 6
grilladocumentoS.ColWidth(1) = 120 * 11
grilladocumentoS.ColWidth(2) = 120 * 5
grilladocumentoS.ColWidth(3) = 120 * 40
grilladocumentoS.ColWidth(4) = 120 * 12
grilladocumentoS.ColWidth(5) = 120 * 2

Rem TITULOS
grilladocumentoS.TextMatrix(0, 0) = "LINEA"
grilladocumentoS.TextMatrix(0, 1) = "CUENTA"
grilladocumentoS.TextMatrix(0, 2) = "CRCC"
grilladocumentoS.TextMatrix(0, 3) = "GLOSA"
grilladocumentoS.TextMatrix(0, 4) = "MONTO"
grilladocumentoS.TextMatrix(0, 5) = "D/H"
grilladocumentoS.Enabled = False

End Sub

Sub LEERMOVIMIENTOS()

PLANTILLA

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim SALDOS As Double
    Dim sumadebe As Double
    Dim sumahaber As Double
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT linea,cuentadelmayor,centrodecosto,glosa,monto,dh "
        cSql.SQL = cSql.SQL + "FROM detallefacturasdecompra"
        cSql.SQL = cSql.SQL + " where tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "' order by linea"
        ' cSql.SQL = cSql.SQL + " where tipo=1 and numero=0000000005 order by linea"
        cSql.Execute
        
        linea = 0: SUMADOR = 0
        If cSql.RowsAffected > 0 Then
            
            
            Set resultados = cSql.OpenResultset
            sumadebe = 0
            SALDOS = 0
            sumahaber = 0
            While Not resultados.EOF
                linea = linea + 1
                grilladocumentoS.Rows = linea + 1
                
                For K = 0 To 5
                grilladocumentoS.TextMatrix(linea, K) = resultados(K)
                Next K
                grilladocumentoS.TextMatrix(linea, 1) = Mid(resultados(1), 1, 2) + "." + Mid(resultados(1), 3, 2) + "." + Mid(resultados(1), 5, 4)
                grilladocumentoS.TextMatrix(linea, 4) = Format(resultados(4), "#,###,###,##0")
                If resultados(5) = "D" Then sumadebe = sumadebe + CDbl(resultados(4))
                If resultados(5) = "H" Then sumahaber = sumahaber + CDbl(resultados(4))

                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing

        End If
    VARIPASO = cSql.RowsAffected
    NUMEROLINEA.text = CDbl(linea) + 1
    Call ceros(NUMEROLINEA)
    Rem If VARIPASO <> "0" Then CARGADATAFIELD
    saldo = sumadebe - sumahaber
    debe.Caption = Format(sumadebe, "##,###,###,##0")
    haber.Caption = Format(sumahaber, "##,###,###,##0")
    saldo.Caption = Format(saldo, "##,###,###,##0")
    End With
no:
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus: caja.SelStart = 0
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub leemayor(cuenta)
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "ctacte"
    campos(3, 0) = "centrocosto"
    campos(4, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo=" + "'" + cuenta + "'"
    
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    If SQLUTIL.estado = 4 Then dato15.text = "": dato16.text = "": dato17.text = "": dato15.SetFocus: GoTo no:
    nombremayor.Caption = SQLUTIL.datos(1, 3)
    If Val(SQLUTIL.datos(2, 3)) <> 0 Then tipocuenta = SQLUTIL.datos(2, 3)
    tipocentro = SQLUTIL.datos(3, 3)

no:

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
    If SQLUTIL.estado = 4 Then crearut: GoTo no:
    nombreproveedor.Caption = SQLUTIL.datos(1, 3)
    dato3.Enabled = True
    dato3.SetFocus
no:

End Sub

Sub crearut()
    scrut = "S"
    
    maestro02.dato1.Enabled = True
    maestro02.dato2.Enabled = True
    maestro02.dato3.Enabled = True
    maestro02.dato4.Enabled = True
    maestro02.dato1.text = tipocuenta
    maestro02.dato2.text = dato9.text
    maestro02.dato3.text = dv.Caption
    maestro02.dato1.Enabled = False
    maestro02.dato2.Enabled = False
    maestro02.dato3.Enabled = False
    maestro02.Show
    maestro02.SetFocus
    
End Sub

Sub leecrcc()
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + dato18.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    If SQLUTIL.estado = 4 Then dato18.text = "": dato18.SetFocus:  GoTo no:
  
    nombrecentrocosto.Caption = SQLUTIL.datos(1, 3)

no:

End Sub
Sub leetipos()
    
    campos(0, 0) = "tipos"
    campos(1, 0) = "nombredocumento"
    campos(2, 0) = ""
    
    campos(0, 2) = "maestrotipodedocumentos"
    condicion = "tipos=" + "'" + dato13.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)

    If SQLUTIL.estado = 4 Then dato13.text = "": dato13.SetFocus:  GoTo no:
    VARIPASO = "S"
    

no:

End Sub
Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub


Sub CARGADATOSAMODIFICAR()
Rem TITULOS

K = grilladocumentoS.row
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

Sub actualizamayor(codigo As String)


    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If dato21.text = "D" Then campos(2, 0) = "debe" + mes
    If dato21.text = "H" Then campos(2, 0) = "haber" + mes
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + codigo + "' and año ='" + añoconta.text + "' order by codigo"
    
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
    If SQLUTIL.estado = 4 Then Stop
    
    VARIPASO = Replace(dato20.text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto + Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
    

    
End Sub
Sub desactualizamayor(codigo As String)


    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If dato21.text = "D" Then campos(2, 0) = "debe" + mes
    If dato21.text = "H" Then campos(2, 0) = "haber" + mes
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + codigo + "' and año ='" + añoconta.text + "' order by codigo"
    
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop:
    
    VARIPASO = Replace(dato20.text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto - Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
no:


    
End Sub


Sub cargadato()


dato15.text = Mid(grilladocumentoS.TextMatrix(K, 1), 1, 2)
dato16.text = Mid(grilladocumentoS.TextMatrix(K, 1), 4, 2)
dato17.text = Mid(grilladocumentoS.TextMatrix(K, 1), 7, 4)
dato18.text = grilladocumentoS.TextMatrix(K, 2)
dato19.text = grilladocumentoS.TextMatrix(K, 3)
dato20.text = grilladocumentoS.TextMatrix(K, 4)
dato21.text = grilladocumentoS.TextMatrix(K, 5)
End Sub
Sub totalfactura()
SUMADOR = CDbl(Replace(dato11.text, ",", "")) + CDbl(Replace(dato12.text, ",", "")) + CDbl(Replace(dato13.text, ",", "") - CDbl(Replace(dato14.text, ",", "")))
total.text = Format(SUMADOR, "###,###,###,##0")
End Sub
Sub leefactura()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "retencion"
    campos(9, 0) = "total"
    campos(10, 0) = "añocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = ""
    campos(0, 2) = "facturasdecompras"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and rut=" + "'" + dato9.text + dv.Caption + "'"

    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    Rem If SQLUTIL.estado = 0 Then modifi = 1: carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Stop
End Sub



Sub grabarcomprobante()
    
    
    Dim tipocon As String
    Dim tipo2 As String
    For K = 1 To Val(NUMEROLINEA.text) - 1
    LINEAS.text = K
    Call ceros(LINEAS)
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
    campos(18, 0) = ""
    
    dato15.text = Mid(grilladocumentoS.TextMatrix(K, 1), 1, 2)
    dato16.text = Mid(grilladocumentoS.TextMatrix(K, 1), 4, 2)
    dato17.text = Mid(grilladocumentoS.TextMatrix(K, 1), 7, 4)
    dato18.text = grilladocumentoS.TextMatrix(K, 2)
    dato19.text = grilladocumentoS.TextMatrix(K, 3)
    dato20.text = Replace(grilladocumentoS.TextMatrix(K, 4), ".", "")
    dato21.text = grilladocumentoS.TextMatrix(K, 5)
    
    
    
    
    If dato1.text = "1" Then tipo2 = "FA": tipocon = "FC"
    If dato1.text = "2" Then tipo2 = "ND": tipocon = "DC"
    If dato1.text = "3" Then tipo2 = "NC": tipocon = "CC"
 
    campos(0, 1) = tipocon
    campos(1, 1) = dato2.text
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato5.text + dato4.text + dato3.text
    campos(4, 1) = dato15.text + dato16.text + dato17.text
 
    If K = 1 Then campos(5, 1) = tipocuenta
    If K <> 1 Then campos(5, 1) = ""
    If K = 1 Then campos(6, 1) = dato9.text + dv.Caption
    If K <> 1 Then campos(6, 1) = ""
    campos(7, 1) = dato18.text
    campos(8, 1) = dato19.text
    campos(9, 1) = "FA"
    campos(10, 1) = dato2.text
    campos(11, 1) = dato5.text + dato4.text + dato3.text
    campos(12, 1) = dato8.text + dato7.text + dato6.text
    campos(13, 1) = dato20.text
    campos(14, 1) = dato21.text
    campos(15, 1) = USUARIO
    campos(16, 1) = mesconta.text
    campos(17, 1) = añoconta.text
    
    campos(0, 2) = "movimientoscontables"
    condicion = ""

    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    Rem tipo2 = Mid(Grid1.Cell(K, 4).text, 2, 2)
    Rem If tipo2 <> "00" Then Call actualizactacte(K, tipo2, Grid1.Cell(K, 5).text)
    Rem If Mid(Grid1.Cell(K, 4).text, 1, 1) = "S" Then Call actualizacrcc(K, Grid1.Cell(K, 6).text)
    Rem If Mid(Grid1.Cell(K, 4).text, 1, 1) = "S" Then Call actualizacrcc(K, Mid(Grid1.Cell(K, 6).text, 1, 2) + "00")
    Call actualizamayor(dato15.text + dato16.text + dato17.text)
    Call actualizamayor(dato15.text + dato16.text + "0000")
    Call actualizamayor(dato15.text + "000000")
    
    Next K
End Sub



Sub leeFOLIO()
    campos(0, 0) = "folio"
    campos(1, 0) = ""
    campos(0, 2) = "facturasdecompras"
    condicion = "mescontable=" + "'" + mesconta.text + "'" + " and añocontable=" + "'" + añoconta.text + "' order by folio desc"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then K = SQLUTIL.datos(0, 3) Else K = 0
    folios.text = K + 1
    
    Call ceros(folios)



End Sub
Sub eliminacomprobantecontable()
For K = 1 To Val(NUMEROLINEA.text) - 1
dato15.text = Mid(grilladocumentoS.TextMatrix(K, 1), 1, 2)
dato16.text = Mid(grilladocumentoS.TextMatrix(K, 1), 4, 2)
dato17.text = Mid(grilladocumentoS.TextMatrix(K, 1), 7, 4)
dato18.text = grilladocumentoS.TextMatrix(K, 2)
dato19.text = grilladocumentoS.TextMatrix(K, 3)
dato20.text = grilladocumentoS.TextMatrix(K, 4)
dato21.text = grilladocumentoS.TextMatrix(K, 5)
Call desactualizamayor(dato15.text + dato16.text + dato17.text)
Call desactualizamayor(dato15.text + dato16.text + "0000")
Call desactualizamayor(dato15.text + "000000")
Next K
End Sub

Sub actualizactacte(row, tipo As String, rut As String)
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + dato3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + dato3.text
    campos(3, 0) = ""
    condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año ='" + dato4.text + "' order by tipo"
    campos(0, 2) = "saldosctacte"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
    If SQLUTIL.estado = 4 Then Stop
   
    VARIPASO = Replace(Grid1.Cell(row, 11).text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto + Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
    

    
End Sub

Sub desactualizactacte(row, tipo As String, rut As String)
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + dato3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + dato3.text
    campos(3, 0) = ""
    condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año ='" + dato4.text + "' order by tipo"
    campos(0, 2) = "saldosctacte"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
    If SQLUTIL.estado = 4 Then Stop
   
    VARIPASO = Replace(Grid1.Cell(row, 11).text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto - Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
    

    
End Sub



Sub actualizacrcc(row, crcc As String)
    

    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + dato3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + dato3.text
    campos(3, 0) = ""
    
    condicion = "codigo=" + "'" + crcc + "' and año ='" + dato4.text + "' order by codigo"
    campos(0, 2) = "saldoscentrosdecosto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
    If SQLUTIL.estado = 4 Then Stop
   
    VARIPASO = Replace(Grid1.Cell(row, 11).text, ".", "")
    
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto + Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
    
    
End Sub

Sub desactualizacrcc(row, crcc As String)
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    If Grid1.Cell(row, 12).text = "D" Then campos(2, 0) = "debe" + dato3.text
    If Grid1.Cell(row, 12).text = "H" Then campos(2, 0) = "haber" + dato3.text
    campos(3, 0) = ""
    condicion = "codigo=" + "'" + crcc + "' and año ='" + dato4.text + "' order by codigo"
    campos(0, 2) = "saldoscentrosdecosto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
    VARIPASO = Replace(grilladetalle(row, 11), ".", "")
    campos(0, 1) = SQLUTIL.datos(0, 3)
    campos(1, 1) = SQLUTIL.datos(1, 3)
    varimonto = SQLUTIL.datos(2, 3)
    campos(2, 1) = Str(varimonto - Val(VARIPASO))
    op = 3
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
    
End Sub

