VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form informa04 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Mayor Analitico"
   ClientHeight    =   10230
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar barra 
      Height          =   495
      Left            =   120
      TabIndex        =   66
      Top             =   9000
      Visible         =   0   'False
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Min             =   1
      Max             =   5000
      Scrolling       =   1
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   3600
      MaxLength       =   8
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame botones 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   615
      Left            =   240
      TabIndex        =   49
      Top             =   9600
      Width           =   9015
      Begin VB.CommandButton Command7 
         Caption         =   "Nueva cartola"
         Height          =   375
         Left            =   7080
         TabIndex        =   65
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exportar Html"
         Height          =   375
         Left            =   4680
         TabIndex        =   52
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exportar Excel"
         Height          =   375
         Left            =   2400
         TabIndex        =   51
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprime Formato Grande"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   2055
      End
   End
   Begin TabDlg.SSTab opciones 
      Height          =   5295
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9340
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cuentas del Mayor"
      TabPicture(0)   =   "cartolasgenerales.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmha03"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmha02"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmha01"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmde01"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmde02"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmde03"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Cuentas Corrientes"
      TabPicture(1)   =   "cartolasgenerales.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "ctde03"
      Tab(1).Control(4)=   "ctde02"
      Tab(1).Control(5)=   "ctde01"
      Tab(1).Control(6)=   "ctha01"
      Tab(1).Control(7)=   "ctha02"
      Tab(1).Control(8)=   "ctha03"
      Tab(1).Control(9)=   "Shape5"
      Tab(1).Control(10)=   "Label3"
      Tab(1).Control(11)=   "Label2"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Centros de Costo"
      TabPicture(2)   =   "cartolasgenerales.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape3"
      Tab(2).Control(1)=   "Label10"
      Tab(2).Control(2)=   "Label11"
      Tab(2).Control(3)=   "Command6"
      Tab(2).Control(4)=   "ccde03"
      Tab(2).Control(5)=   "ccde02"
      Tab(2).Control(6)=   "ccde01"
      Tab(2).Control(7)=   "ccha01"
      Tab(2).Control(8)=   "ccha02"
      Tab(2).Control(9)=   "ccha03"
      Tab(2).Control(10)=   "Frame3"
      Tab(2).Control(11)=   "Frame6"
      Tab(2).ControlCount=   12
      Begin VB.Frame Frame6 
         Caption         =   "Opciones"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   59
         Top             =   600
         Width           =   3615
         Begin VB.OptionButton cctoda 
            Caption         =   "Todas las Cuentas"
            Height          =   495
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton ccindi 
            Caption         =   "Una Cuenta Individual"
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Opciones"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   56
         Top             =   600
         Width           =   3615
         Begin VB.OptionButton cttoda 
            Caption         =   "Todos los Tipos"
            Height          =   495
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton ctindi 
            Caption         =   "Una Cuenta corriente Individual"
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   3375
         End
         Begin VB.OptionButton cttodatipo 
            Caption         =   "Todas las Cuentas de Un tipo"
            Height          =   495
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Opciones"
         Height          =   1455
         Left            =   240
         TabIndex        =   53
         Top             =   720
         Width           =   3615
         Begin VB.OptionButton cmtoda 
            Caption         =   "Todas las Cuentas"
            Height          =   495
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton cmindi 
            Caption         =   "Una Cuenta Individual"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro de Costo"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   43
         Top             =   2280
         Width           =   7935
         Begin VB.TextBox ccdato1 
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
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   46
            Tag             =   "codigo"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox ccdato2 
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
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   45
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text16 
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
            MaxLength       =   50
            TabIndex        =   44
            Tag             =   "nombre"
            Top             =   720
            Width           =   6015
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Codigo Cuenta"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.TextBox ccha03 
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
         Left            =   -68160
         MaxLength       =   4
         TabIndex        =   40
         Tag             =   "fecha"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox ccha02 
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
         Left            =   -68520
         MaxLength       =   2
         TabIndex        =   39
         Tag             =   "fecha"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox ccha01 
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
         Left            =   -68880
         MaxLength       =   2
         TabIndex        =   38
         Tag             =   "fechavencimiento"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox ccde01 
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
         Left            =   -70440
         MaxLength       =   2
         TabIndex        =   37
         Tag             =   "fecha"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox ccde02 
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
         Left            =   -70080
         MaxLength       =   2
         TabIndex        =   36
         Tag             =   "fecha"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox ccde03 
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
         Left            =   -69720
         MaxLength       =   4
         TabIndex        =   35
         Tag             =   "fecha"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Genera Cartola"
         Height          =   255
         Left            =   -72600
         TabIndex        =   34
         Top             =   3960
         Width           =   3375
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuentas Corrientes"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   27
         Top             =   2400
         Width           =   7935
         Begin VB.TextBox dv 
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
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   63
            Tag             =   "tipo"
            Top             =   600
            Width           =   255
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   30
            Tag             =   "nombre"
            Top             =   960
            Width           =   6015
         End
         Begin VB.TextBox ctdato2 
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
            Left            =   1440
            MaxLength       =   9
            TabIndex        =   29
            Tag             =   "rut"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox ctdato1 
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
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   28
            Tag             =   "tipo"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo Cuenta"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rut"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Genera Cartola"
         Height          =   255
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta Individual"
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   7935
         Begin VB.TextBox cmdato1 
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
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   23
            Tag             =   "codigo"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox Text9 
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
            MaxLength       =   50
            TabIndex        =   22
            Tag             =   "nombre"
            Top             =   840
            Width           =   4575
         End
         Begin VB.TextBox cmdato2 
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
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   21
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox cmdato3 
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
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   20
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Codigo Cuenta"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Genera Cartola"
         Height          =   255
         Left            =   -72600
         TabIndex        =   18
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox ctde03 
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
         Left            =   -69720
         MaxLength       =   4
         TabIndex        =   15
         Tag             =   "fecha"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox ctde02 
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
         Left            =   -70080
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "fecha"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox ctde01 
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
         Left            =   -70440
         MaxLength       =   2
         TabIndex        =   13
         Tag             =   "fecha"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox ctha01 
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
         Left            =   -68880
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "fechavencimiento"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox ctha02 
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
         Left            =   -68520
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox ctha03 
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
         Left            =   -68160
         MaxLength       =   4
         TabIndex        =   10
         Tag             =   "fecha"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox cmde03 
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
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "fecha"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox cmde02 
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
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "fecha"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox cmde01 
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox cmha01 
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
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fechavencimiento"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox cmha02 
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
         Left            =   6480
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox cmha03 
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
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "fecha"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESDE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70440
         TabIndex        =   42
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HASTA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -68880
         TabIndex        =   41
         Top             =   840
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         Height          =   1215
         Left            =   -70800
         Top             =   720
         Width           =   3735
      End
      Begin VB.Shape Shape5 
         Height          =   1215
         Left            =   -70800
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HASTA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -68880
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESDE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -70440
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   4200
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HASTA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESDE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   16748
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
End
Attribute VB_Name = "informa04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private formatogrilla(20, 20)
Private lin As Double
Private saldo As Double
Private dedonde As Integer
Private tipoctacte As String





Private Sub busca_Click()

End Sub

Private Sub cmdato1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then Call ayudamayor(cmdato1)

End Sub

Private Sub Command1_Click()


 Grid1.DefaultFont.Size = 6.5
For K = 1 To 15 - 1
Grid1.Column(K).Width = Val(formatogrilla(2, K)) * Grid1.DefaultFont.Size
Next K
Grid1.PageSetup.Orientation = cellPortrait

Grid1.PageSetup.PrintFixedRow = True


'Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.BottomMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0


cabeza




Grid1.PrintPreview 75


End Sub
Sub cabeza()
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Libro Diario"
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 18
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
    For K = 1 To 5
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = DATOSEMPRESA(K)
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Italic = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Color = RGB(128, 0, 0)
    objReportTitle.Align = CellLeft
    Grid1.ReportTitles.Add objReportTitle
    Next K
With Grid1.PageSetup
        
        .Footer = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        .FooterAlignment = cellRight
        .FooterFont.Name = "Verdana"
        .FooterFont.Size = 7
        .FooterMargin = 0.1
        
End With

End Sub

Private Sub Command2_Click()
dedonde = 2

If cttoda.Value = True Then opciones.Visible = False: Call acepta(2)
If ctindi.Value = True Then opciones.Visible = False: Call acepta(2)
If cttodatipo.Value = True Then opciones.Visible = False: Call acepta(2)
End Sub

Private Sub Command5_Click()
dedonde = 1

If cmtoda.Value = True Then opciones.Visible = False: Call acepta(1)
If cmindi.Value = True Then opciones.Visible = False: Call acepta(1)

End Sub
Sub acepta(opcion)

CARGAGRILLA
barra.Visible = True

If opcion = 1 Then leecuentas
If opcion = 2 Then leercuentascorrientes
If opcion = 3 Then leecrcc
barra.Visible = False
botones.Visible = True
Grid1.Visible = True



End Sub

Private Sub Command6_Click()
dedonde = 3

If cctoda.Value = True Then opciones.Visible = False: Call acepta(3)
If ccindi.Value = True Then opciones.Visible = False: Call acepta(3)

End Sub

Private Sub Command7_Click()
lin = 0

Grid1.Visible = False
botones.Visible = False
opciones.Visible = True
barra.Visible = False

End Sub

Private Sub Command8_Click()
barra.Min = 1
barra.Max = 300


End Sub

Private Sub Form_Load()
    
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)
    'CARGAGRILLA
    'leecuentas
Grid1.Visible = False
botones.Visible = False
cmde01.text = "01"
cmde02.text = "01"
cmde03.text = "2005"
cmha01.text = Mid(fechasistema, 1, 2)
cmha02.text = Mid(fechasistema, 4, 2)
cmha03.text = Mid(fechasistema, 7, 4)
ctde01.text = "01"
ctde02.text = "01"
ctde03.text = "2005"
ctha01.text = Mid(fechasistema, 1, 2)
ctha02.text = Mid(fechasistema, 4, 2)
ctha03.text = Mid(fechasistema, 7, 4)
ccde01.text = "01"
ccde02.text = "01"
ccde03.text = "2005"
ccha01.text = Mid(fechasistema, 1, 2)
ccha02.text = Mid(fechasistema, 4, 2)
ccha03.text = Mid(fechasistema, 7, 4)
lin = 0
End Sub


    
Sub LEERMOVIMIENTOS(CUENTA, nombre)
Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    If dedonde = 1 Then fecha1 = cmde03.text + "-" + cmde02.text + "-" + cmde01.text
    If dedonde = 1 Then fecha2 = cmha03.text + "-" + cmha02.text + "-" + cmha01.text
    If dedonde = 2 Then fecha1 = ctde03.text + "-" + ctde02.text + "-" + ctde01.text
    If dedonde = 2 Then fecha2 = ctha03.text + "-" + ctha02.text + "-" + ctha01.text
    If dedonde = 3 Then fecha1 = ccde03.text + "-" + ccde02.text + "-" + ccde01.text
    If dedonde = 3 Then fecha2 = ccha03.text + "-" + ccha02.text + "-" + ccha01.text
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte "
        If dedonde = 1 Then cSql.SQL = cSql.SQL + "FROM movimientoscontables where codigocuenta='" + CUENTA + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "'"
        
        If dedonde = 2 Then cSql.SQL = cSql.SQL + "FROM movimientoscontables where tipoctacte='" + tipoctacte + "' and rutctacte='" + CUENTA + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "'"
        If dedonde = 3 Then cSql.SQL = cSql.SQL + "FROM movimientoscontables where centrocosto='" + CUENTA + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "'"
        
        cSql.SQL = cSql.SQL + "order by codigocuenta,fecha"
        cSql.Execute
        Grid1.AutoRedraw = False
        If dedonde <> 2 Then Call DATOSSALDOS(CUENTA)
        If dedonde = 2 Then Call DATOSSALDOSctacte(CUENTA)
        For K = 1 To 6
        Grid1.Column(K).Locked = False
        Next K
        If saldo <> 0 Then
        lin = lin + 1
        Grid1.Rows = Grid1.Rows + 1
                
        Grid1.Range(lin, 1, lin, 6).Merge
        
        Grid1.Cell(lin, 1).CellType = cellTextBox
        
        Grid1.Cell(lin, 10).CellType = cellTextBox
        
        Grid1.Cell(lin, 1).text = nombre
        If dedonde = 2 Then Grid1.Cell(lin, 7).text = tipoctacte
        Grid1.Cell(lin, 10).text = "SALDO-->"
        
        Grid1.Cell(lin, 13).text = saldo
        End If
        
        If cSql.RowsAffected > 0 Then
        
        
        Set resultados = cSql.OpenResultset
        
         While Not resultados.EOF
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
             For K = 0 To 9
             Grid1.Cell(lin, K + 1).text = resultados(K)
             Next K
             If resultados(11) = "D" Then Grid1.Cell(lin, 11).text = resultados(10): anted = anted + resultados(10): saldo = saldo + resultados(10)
             If resultados(11) = "H" Then Grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(10): saldo = saldo - resultados(10)
             Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
             Grid1.Cell(lin, 13).text = saldo
             resultados.MoveNext
           
         Wend
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
         
         Call totalcomprobante(lin)
          resultados.Close
            Set resultados = Nothing

        End If
 For K = 1 To 6
        Grid1.Column(K).Locked = True
        
        Next K
Grid1.AutoRedraw = True
Grid1.Refresh

End Sub

Sub totalcomprobante(ROW)
    Grid1.Range(ROW, 11, ROW, 12).Borders(cellEdgeTop) = cellThin
    Grid1.Cell(ROW, 10).CellType = cellTextBox
    Grid1.Cell(ROW, 10).text = "TOTAL "
    Grid1.Cell(ROW, 11).text = anted
    Grid1.Cell(ROW, 12).text = anteh
    lin = lin + 2
             Grid1.Rows = Grid1.Rows + 2
        
        anted = 0: anteh = 0: saldo = 0
    End Sub
    





Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7.5
    
    
    formatogrilla(1, 1) = "FECHA"
    formatogrilla(1, 2) = "TP"
    formatogrilla(1, 3) = "NUMERO"
    formatogrilla(1, 4) = "LINEA"
    formatogrilla(1, 5) = "CUENTA"
    formatogrilla(1, 6) = "GLOSA"
    formatogrilla(1, 7) = "TP"
    formatogrilla(1, 8) = "NUMERO"
    formatogrilla(1, 9) = "EMISION"
    formatogrilla(1, 10) = "VENCIMIENTO"
    formatogrilla(1, 11) = "DEBE"
    formatogrilla(1, 12) = "HABER"
    formatogrilla(1, 13) = "SALDO"
    formatogrilla(1, 14) = "NOMBRE CUENTA"
    formatogrilla(1, 15) = "CUENTA CORRIENTE"
    formatogrilla(1, 16) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 3) = "10"
    formatogrilla(2, 4) = "3"
    formatogrilla(2, 5) = "10"
    formatogrilla(2, 6) = "30"
    formatogrilla(2, 7) = "2"
    formatogrilla(2, 8) = "10"
    formatogrilla(2, 9) = "10"
    formatogrilla(2, 10) = "10"
    formatogrilla(2, 11) = "12"
    formatogrilla(2, 12) = "12"
    formatogrilla(2, 13) = "12"
    formatogrilla(2, 14) = "30"
    formatogrilla(2, 15) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "D"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "S"
    formatogrilla(3, 4) = "S"
    formatogrilla(3, 5) = "S"
    formatogrilla(3, 6) = "S"
    formatogrilla(3, 7) = "S"
    formatogrilla(3, 8) = "S"
    formatogrilla(3, 9) = "D"
    formatogrilla(3, 10) = "D"
    formatogrilla(3, 11) = "N"
    formatogrilla(3, 12) = "N"
    formatogrilla(3, 13) = "N"
    formatogrilla(3, 14) = "S"
    formatogrilla(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    formatogrilla(4, 11) = "###,###,###,###"
    formatogrilla(4, 12) = "###,###,###,###"
    formatogrilla(4, 13) = "###,###,###,###"
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
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
    
    Grid1.Cols = 15
    Grid1.Rows = 2
    
     'Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
   'Grid1.BackColorFixed = RGB(90, 158, 214)
   ' Grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' Grid1.BackColorBkg = RGB(90, 158, 214)
   ' Grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' Grid1.BackColor1 = RGB(231, 235, 247)
   ' Grid1.BackColor2 = RGB(239, 243, 255)
   ' Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For K = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, K).text = formatogrilla(1, K)
        Grid1.Column(K).Width = Val(formatogrilla(2, K)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(K).MaxLength = Val(formatogrilla(2, K))
        Grid1.Column(K).FormatString = formatogrilla(4, K)
        Grid1.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then Grid1.Column(K).Alignment = cellRightCenter
        If formatogrilla(3, K) = "D" Then Grid1.Column(K).CellType = cellCalendar
        
    Next K
End Sub


Sub leecuentas()
Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
        barra.Min = 0.1
        
        Set cSql2.ActiveConnection = db
        cSql2.SQL = "SELECT codigo,nombre "
        cSql2.SQL = cSql2.SQL + "FROM cuentasdelmayor "
        If cmindi.Value = True Then cSql2.SQL = cSql2.SQL + "where codigo='" + cmdato1.text + cmdato2.text + cmdato3.text + "' "
        
        cSql2.SQL = cSql2.SQL + "order by codigo"
        cSql2.Execute
        barra.Max = cSql2.RowsAffected + 1
        
        LINEAS = 0
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        If Mid(resultados2(0), 5, 4) <> "0000" Then Call LEERMOVIMIENTOS(resultados2(0), resultados2(1))
        barra.Value = LINEAS
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        Grid1.Column(8).Locked = True
        Grid1.Column(9).Locked = True
        Grid1.Column(10).Locked = True
        

End Sub

Sub LEERSALDOS(CUENTA)
Dim resultados3 As rdoResultset
    
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mesante As Integer
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = "debe01"
    campos(5, 0) = "debe02"
    campos(6, 0) = "debe03"
    campos(7, 0) = "debe04"
    campos(8, 0) = "debe05"
    campos(9, 0) = "debe06"
    campos(10, 0) = "debe07"
    campos(11, 0) = "debe08"
    campos(12, 0) = "debe09"
    campos(13, 0) = "debe10"
    campos(14, 0) = "debe11"
    campos(15, 0) = "debe12"
    campos(16, 0) = "haber01"
    campos(17, 0) = "haber02"
    campos(18, 0) = "haber03"
    campos(19, 0) = "haber04"
    campos(20, 0) = "haber05"
    campos(21, 0) = "haber06"
    campos(22, 0) = "haber07"
    campos(23, 0) = "haber08"
    campos(24, 0) = "haber09"
    campos(25, 0) = "HABER10"
    campos(26, 0) = "HABER11"
    campos(27, 0) = "HABER12"
    campos(28, 0) = ""
    
    If dedonde = 1 Then condicion = "codigo=" + "'" + CUENTA + "' and año='" + cmde03.text + "' order by codigo"
    If dedonde = 3 Then condicion = "codigo=" + "'" + CUENTA + "' and año='" + ccde03.text + "' order by codigo"
    
    If dedonde = 1 Then campos(0, 2) = "saldosdelmayor"
    If dedonde = 3 Then campos(0, 2) = "saldoscentrosdecosto"
    
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
   ' If SQLUTIL.ESTADO = 4 Then Stop
    SUMADOR = Val(SQLUTIL.datos(2, 3)) - Val(SQLUTIL.datos(3, 3))
    mesante = cmde02.text - 1
    
    If mesante = 0 Then mesante = 0
    
    For K = 1 To mesante
    SUMADOR = SUMADOR + Val(SQLUTIL.datos(K + 3, 3)) - Val(SQLUTIL.datos(K + 15, 3))
    Next K
    saldo = SUMADOR
Rem acumula fecha
    fecha1 = cmde03.text + "-" + cmde02.text + "-" + cmde01.text
    
        
        Set cSql3.ActiveConnection = db
        cSql3.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte "
        If dedonde = 1 Then cSql3.SQL = cSql3.SQL + "FROM movimientoscontables where codigocuenta='" + CUENTA + "' and fecha<'" + fecha1 + "'"
        If dedonde = 2 Then cSql3.SQL = cSql3.SQL + "FROM movimientoscontables where tipoctacte='" + tipoctacte + "' and rutctacte='" + CUENTA + "' and fecha<'" + fecha1 + "'"
        If dedonde = 3 Then cSql3.SQL = cSql3.SQL + "FROM movimientoscontables where centrocosto='" + CUENTA + "' and fecha<'" + fecha1 + "'"
        
        cSql3.SQL = cSql3.SQL + "order by codigocuenta,fecha"
        cSql3.Execute
        
        If cSql3.RowsAffected > 0 Then
        
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
         If resultados3(11) = "D" Then saldo = saldo + resultados3(10)
         If resultados3(11) = "H" Then saldo = saldo - resultados3(10)
         
             
             resultados3.MoveNext
           
         Wend
          resultados3.Close
            Set resultados3 = Nothing

        End If

End Sub
Sub DATOSSALDOS(CUENTA)
Call LEERSALDOS(CUENTA)






End Sub
Sub leecrcc()
Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set cSql2.ActiveConnection = db
        cSql2.SQL = "SELECT codigo,nombre "
        cSql2.SQL = cSql2.SQL + "FROM centrosdecosto "
        If ccindi.Value = True Then cSql2.SQL = cSql2.SQL + "where codigo='" + ccdato1.text + ccdato2.text + "' "
        
        cSql2.SQL = cSql2.SQL + "order by codigo"
        cSql2.Execute
        LINEAS = 0
  
        If cSql2.RowsAffected > 0 Then
     
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
    
        If Mid(resultados2(0), 3, 2) <> "00" Then Call LEERMOVIMIENTOS(resultados2(0), resultados2(1))
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        Grid1.Column(8).Locked = True
        Grid1.Column(9).Locked = True
        Grid1.Column(10).Locked = True
        

End Sub

Sub leercuentascorrientes()
Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
        Set cSql2.ActiveConnection = db
        cSql2.SQL = "SELECT tipo,rut,nombre "
        If cttoda.Value = True Then cSql2.SQL = cSql2.SQL + "FROM cuentascorrientes where tipo>'00' "
        If cttodatipo.Value = True Then cSql2.SQL = cSql2.SQL + "FROM cuentascorrientes where tipo='" + ctdato1.text + "' "
        If ctindi.Value = True Then cSql2.SQL = cSql2.SQL + "FROM cuentascorrientes where tipo='" + ctdato1.text + "' and rut='" + ctdato2.text + dv.text + "' "
        
        cSql2.SQL = cSql2.SQL + "order by tipo,nombre"
       
        cSql2.Execute
        lin = 0
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        tipoctacte = resultados2(0)
        Call LEERMOVIMIENTOS(resultados2(1), resultados2(2))
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        Grid1.Column(8).Locked = True
        Grid1.Column(9).Locked = True
        Grid1.Column(10).Locked = True
        

End Sub
Sub leerSALDOSctacte(CUENTA)
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = "debe01"
    campos(6, 0) = "debe02"
    campos(7, 0) = "debe03"
    campos(8, 0) = "debe04"
    campos(9, 0) = "debe05"
    campos(10, 0) = "debe06"
    campos(11, 0) = "debe07"
    campos(12, 0) = "debe08"
    campos(13, 0) = "debe09"
    campos(14, 0) = "debe10"
    campos(15, 0) = "debe11"
    campos(16, 0) = "debe12"
    campos(17, 0) = "haber01"
    campos(18, 0) = "haber02"
    campos(19, 0) = "haber03"
    campos(20, 0) = "haber04"
    campos(21, 0) = "haber05"
    campos(22, 0) = "haber06"
    campos(23, 0) = "haber07"
    campos(24, 0) = "haber08"
    campos(25, 0) = "haber09"
    campos(26, 0) = "HABER10"
    campos(27, 0) = "HABER11"
    campos(28, 0) = "HABER12"
    campos(29, 0) = ""
    condicion = "tipo=" + "'" + tipoctacte + "' and rut='" + CUENTA + "' and año='" + ctde03.text + "'"
    campos(0, 2) = "saldosctacte"
    
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
Stop
End Sub

Sub DATOSSALDOSctacte(CUENTA)

Call leerSALDOSctacte(CUENTA)
SUMADOR = Val(SQLUTIL.datos(3, 3)) - Val(SQLUTIL.datos(4, 3))
For K = 5 To 16
SUMADOR = SUMADOR + Val(SQLUTIL.datos(K, 3)) - Val(SQLUTIL.datos(K + 12, 3))
Next K
saldo = SUMADOR

End Sub


Sub ayudamayor(ByRef caja As TextBox)
  
   
    
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "no"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then cmdato1.SetFocus: GoTo no
    cmdato1.text = Mid(pivote.text, 1, 2)
    cmdato2.text = Mid(pivote.text, 3, 2)
    cmdato3.text = Mid(pivote.text, 5, 4)
    
    
    caja.Enabled = True
    caja.SetFocus
    caja.MaxLength = 2
no:
End Sub
   
Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Sub LEERSALDOSFECHA(CUENTA, fecha1, fecha2)
Dim resultados3 As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim PASO As String
        
        Set cSql3.ActiveConnection = db
        cSql3.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte "
        If dedonde = 1 Then cSql3.SQL = cSql3.SQL + "FROM movimientoscontables where codigocuenta='" + CUENTA + "' and fecha>'" + fecha1 + "' and fecha<'" + fecha2 + "'"
        If dedonde = 2 Then cSql3.SQL = cSql3.SQL + "FROM movimientoscontables where tipoctacte='" + tipoctacte + "' and rutctacte='" + CUENTA + "' and fecha>'" + fecha1 + "' and fecha<'" + fecha2 + "'"
        If dedonde = 3 Then cSql3.SQL = cSql3.SQL + "FROM movimientoscontables where centrocosto='" + CUENTA + "' and fecha>'" + fecha1 + "' and fecha<'" + fecha2 + "'"
        
        cSql3.SQL = cSql3.SQL + "order by codigocuenta,fecha"
        cSql3.Execute
        
        If cSql3.RowsAffected > 0 Then
        
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
             
             resultados3.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
 
End Sub

