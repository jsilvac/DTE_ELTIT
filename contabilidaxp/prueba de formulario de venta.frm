VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame datospersonales 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   2175
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   10455
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   10455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RUT"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   7080
      Width           =   9495
   End
   Begin VB.Frame Frame3 
      Caption         =   "factura final"
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   9495
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6840
         TabIndex        =   7
         Text            =   "Text4"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   6000
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7223
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "envcabezado deventa"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   9495
      Begin VB.Label Label1 
         Caption         =   "rut"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
