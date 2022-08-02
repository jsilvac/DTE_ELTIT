VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form publi0001 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contrato Publicidad"
   ClientHeight    =   8535
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   15150
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1010
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   6480
      TabIndex        =   52
      Top             =   6960
      Width           =   3135
      _ExtentX        =   5530
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
         TabIndex        =   54
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   53
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   4065
      Left            =   9720
      TabIndex        =   20
      Top             =   4350
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   7170
      BackColor       =   16761024
      Caption         =   "Fechas de Pagos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Alignment       =   1
      Begin FlexCell.Grid Grid1 
         Height          =   3705
         Left            =   90
         TabIndex        =   21
         Top             =   315
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   6535
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6435
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11351
      BackColor       =   16744576
      Caption         =   "DATOS DEL CONTRATO"
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
      Alignment       =   1
      Begin XPFrame.FrameXp detalle 
         Height          =   2055
         Left            =   120
         TabIndex        =   47
         Top             =   3960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3625
         BackColor       =   16761024
         Caption         =   "Glosa"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
         ColorBarraAbajo =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin RichTextLib.RichTextBox richglosa 
            Height          =   1695
            Left            =   0
            TabIndex        =   49
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   2990
            _Version        =   393217
            TextRTF         =   $"maestropublicidad.frx":0000
         End
         Begin FlexCell.Grid Grid2 
            Height          =   1695
            Left            =   0
            TabIndex        =   48
            Top             =   1920
            Visible         =   0   'False
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   2990
            BackColor1      =   12648447
            BackColor2      =   12648447
            BackColorActiveCellSel=   16777088
            BackColorBkg    =   16761024
            BackColorFixedSel=   16761024
            BackColorScrollBar=   16744576
            BorderColor     =   16744576
            CellBorderColor =   16744576
            CellBorderColorFixed=   16744576
            SelectionBorderColor=   16744576
            Cols            =   5
            DefaultFontName =   "Arial"
            DefaultFontSize =   8.25
            DisplayRowIndex =   -1  'True
            ForeColorFixed  =   8388608
            GridColor       =   16744576
            Rows            =   30
            DateFormat      =   2
         End
      End
      Begin VB.TextBox dato8 
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
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   44
         Tag             =   "fono"
         Top             =   3600
         Width           =   6015
      End
      Begin VB.CommandButton cmdgrabar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   6090
         Width           =   1455
      End
      Begin XPFrame.FrameXp frmbase 
         Height          =   1680
         Left            =   6240
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2963
         BackColor       =   12632256
         Caption         =   "Tipo Bases"
         CaptionEstilo3D =   1
         BackColor       =   12632256
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
         Begin VB.Label lblfijopesos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 2 - FIJO EN PESOS"
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
            Height          =   285
            Left            =   135
            TabIndex        =   42
            Top             =   840
            Width           =   2475
         End
         Begin VB.Label lblfijouf 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 3 - FIJO EN U.F."
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
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Width           =   2475
         End
         Begin VB.Label lblporcentaje 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 1 - PORCENTAJE"
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
            Height          =   285
            Left            =   135
            TabIndex        =   40
            Top             =   495
            Width           =   2475
         End
      End
      Begin XPFrame.FrameXp frmTipo 
         Height          =   2280
         Left            =   6240
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   4022
         BackColor       =   12632256
         Caption         =   "Tipo Facturar"
         CaptionEstilo3D =   1
         BackColor       =   12632256
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
         Begin VB.Label lblanual 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 5 - ANUAL"
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
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Top             =   1920
            Width           =   2475
         End
         Begin VB.Label lblmensual 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 1 - MENSUAL"
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
            Height          =   285
            Left            =   135
            TabIndex        =   37
            Top             =   495
            Width           =   2475
         End
         Begin VB.Label lbltrimestral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 3 - TRIMESTRAL"
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
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   2475
         End
         Begin VB.Label lblsemestral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 4 - SEMESTRAL"
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
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Top             =   1560
            Width           =   2475
         End
         Begin VB.Label lblbimestral 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 2 - BIMESTRAL"
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
            Height          =   285
            Left            =   135
            TabIndex        =   34
            Top             =   840
            Width           =   2475
         End
      End
      Begin VB.TextBox dato7 
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
         Left            =   2880
         MaxLength       =   9
         TabIndex        =   31
         Tag             =   "rut"
         Top             =   3195
         Width           =   1140
      End
      Begin VB.TextBox HASTA1 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   28
         Tag             =   "fecha"
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox HASTA2 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   27
         Tag             =   "fecha"
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox HASTA3 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   26
         Tag             =   "fecha"
         Top             =   1395
         Width           =   615
      End
      Begin VB.TextBox DESDE1 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   25
         Tag             =   "fecha"
         Top             =   1035
         Width           =   375
      End
      Begin VB.TextBox DESDE2 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   24
         Tag             =   "fecha"
         Top             =   1035
         Width           =   375
      End
      Begin VB.TextBox DESDE3 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   23
         Tag             =   "fecha"
         Top             =   1035
         Width           =   615
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
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "tipo"
         Top             =   315
         Width           =   1140
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
         Left            =   2880
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "rut"
         Top             =   675
         Width           =   1140
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
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   2
         Tag             =   "nombre"
         Top             =   2100
         Width           =   375
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
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   3
         Tag             =   "ciudad"
         Top             =   1755
         Width           =   375
      End
      Begin VB.TextBox dato5 
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
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "giro"
         Top             =   2475
         Width           =   1815
      End
      Begin VB.TextBox dato6 
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
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   5
         Tag             =   "fono"
         Top             =   2835
         Width           =   375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre Represante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   3600
         Width           =   2610
      End
      Begin VB.Label LBLDVV 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4125
         TabIndex        =   32
         Top             =   3195
         Width           =   255
      End
      Begin VB.Label lblbase 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3360
         TabIndex        =   30
         Top             =   2100
         Width           =   2535
      End
      Begin VB.Label lblfacturar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3360
         TabIndex        =   29
         Top             =   1755
         Width           =   2535
      End
      Begin VB.Label lblnombreproveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4440
         TabIndex        =   22
         Top             =   675
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   675
         Width           =   2610
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nº Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   315
         Width           =   2610
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1035
         Width           =   2610
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1395
         Width           =   2610
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Facturar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1755
         Width           =   2610
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " En base a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   2100
         Width           =   2610
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Represante Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   3195
         Width           =   2610
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   2475
         Width           =   2610
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descontados del pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2835
         Width           =   2610
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4125
         TabIndex        =   10
         Top             =   675
         Width           =   255
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15120
      TabIndex        =   8
      Top             =   8535
      Width           =   15150
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8400
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin FlexCell.Grid impresioncontrato 
      Height          =   375
      Left            =   10320
      TabIndex        =   46
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   4065
      Left            =   9720
      TabIndex        =   50
      Top             =   180
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   7170
      BackColor       =   16761024
      Caption         =   "Contratos "
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Alignment       =   1
      Begin FlexCell.Grid Grid3 
         Height          =   3705
         Left            =   90
         TabIndex        =   51
         Top             =   315
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   6535
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
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
Attribute VB_Name = "publi0001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Public saldoglobal As Double
Private MODIFI As Integer
  
Private Sub cmdgrabar_Click()
If richglosa.text <> "" Then
    grabar
    grabar2
    retorno
End If
End Sub

Private Sub dato1_GotFocus()
dato1.text = leerultimocontrato
Call cargatexto(dato1)
CARGAGRILLAMESES
 End Sub

 

Private Sub dato2_GotFocus()
 Call cargatexto(dato2)
End Sub
  
Private Sub DATO2_LostFocus()
If lblnombreproveedor.Caption = "" Then
    dato2.SetFocus
End If
End Sub

Private Sub dato3_GotFocus()
frmTipo.Visible = True
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
Dim k As Double
Dim i As Double
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato3.text <> "" Then
     Select Case dato3.text
     
     Case 1
        lblfacturar.Caption = "MENSUAL"
     Case 2
        lblfacturar.Caption = "BIMESTRAL"
     Case 3
        lblfacturar.Caption = "TRIMESTRAL"
     Case 4
        lblfacturar.Caption = "SEMESTRAL"
     Case 5
        lblfacturar.Caption = "ANUAL"
     End Select
     
    dato4.Enabled = True
    dato4.SetFocus
    End If
End Sub

Private Sub DATO3_LostFocus()
If lblfacturar.Caption = "" Then
    dato3.SetFocus
    dato4.Enabled = False
Else
    frmTipo.Visible = False
End If
End Sub

Private Sub dato4_GotFocus()
 frmbase.Visible = True
End Sub

Private Sub dato4_LostFocus()
If lblbase.Caption = "" Then
    dato4.Enabled = True
    
    dato4.SetFocus
    DATO5.Enabled = False
Else
    frmbase.Visible = False
End If
End Sub

Private Sub dato5_GotFocus()
Call cargatexto(DATO5)
End Sub
Private Sub dato6_GotFocus()

Call cargatexto(dato6)
End Sub
Private Sub dato7_GotFocus()
Call cargatexto(dato7)
End Sub
  
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaProveedor(dato2)
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
If dato8.text <> "" And KeyAscii = 13 Then
'    Grid2.Cell(1, 1).SetFocus
     richglosa.SetFocus
End If
End Sub

 

Private Sub dato9_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Grid1_DblClick()
If Grid1.Cell(Grid1.ActiveCell.row, 1).text = "1" Then
    Grid1.Cell(Grid1.ActiveCell.row, 1).text = "0"
Else
    Grid1.Cell(Grid1.ActiveCell.row, 1).text = "1"
End If
End Sub

 

Private Sub Grid3_DblClick()
    If Grid3.Rows > 1 Then
        dato1.text = Grid3.Cell(Grid3.ActiveCell.row, 1).text
        Call dato1_KeyPress(13)
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

Private Sub Form_Load()
    Call CENTRAR(Me)
    CARGAGRILLAMESES
    Call CARGAGRILLA(20, 2)
    CARGAGRILLAVIGENTES
    Call leercontratosvigentes("")
'    cargameses
End Sub


Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato1.text <> "" Then
    Call ceros(dato1)
        If leer("=") = False Then
            dato2.Enabled = True
            dato2.SetFocus
        Else
            opciones.Visible = True
            opciones.SetFocus
        End If
    End If
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
   
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato2.text <> "" Then
    Call ceros(dato2)
    DV.Caption = rut(dato2.text)
    lblnombreproveedor.Caption = LEERNOMBREPROVEEDOR(dato2.text & DV.Caption)
    Call leercontratosvigentes(dato2.text & DV.Caption)
    DESDE1.Enabled = True
    DESDE1.SetFocus
    End If
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato4.text <> "" Then
    Select Case dato4.text
     
     Case 1
        lblbase.Caption = "PORCENTAJE"
     Case 2
        lblbase.Caption = "FIJO EN PESOS"
     Case 3
        lblbase.Caption = "FIJO EN U.F."
     
     End Select
     
    DATO5.Enabled = True
    DATO5.SetFocus
    End If
    
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And DATO5.text <> "" Then
    dato6.Enabled = True
    dato6.SetFocus
    End If
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And dato6.text <> "" And (dato6.text = "S" Or dato6.text = "N") Then
    dato7.Enabled = True
    dato7.SetFocus
    End If
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato7.text <> "" Then
        Call ceros(dato7)
        LBLDVV.Caption = rut(dato7.text)
        dato8.Enabled = True
        dato8.SetFocus
    End If
    
  End Sub
Function leer(signo As String) As Boolean
    campos(0, 0) = "numero"
    campos(1, 0) = "rut"
    campos(2, 0) = "fechainicio"
    campos(3, 0) = "fechatermino"
    campos(4, 0) = "facturar"
    campos(5, 0) = "base"
    campos(6, 0) = "monto"
    campos(7, 0) = "descontado"
    campos(8, 0) = "rutcontacto"
    campos(9, 0) = "nombrecontacto"
    campos(10, 0) = "glosa"
    campos(11, 0) = ""
    
    campos(0, 2) = "contratopublicidad"
    If signo = "=" Then
        condicion = "numero " & signo & "'" & dato1.text & "'  order by numero desc"
        
    End If
    If signo = "<" Then
        condicion = "numero " & signo & "'" & dato1.text & "'  order by numero desc"
        
    End If
    If signo = ">" Then
        condicion = "numero " & signo & "'" & dato1.text & "'  order by numero asc"
        
    End If
    
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
        leer = False
    Else
        leer = True
        carga
        opciones.Visible = True
'        disponible (True)
'        habilita (True)
        opciones.SetFocus
    End If
        
no:
End Function
  

Sub carga()

Dim k As Double
Dim conta As Double
Dim palabra As String
    habilita (True)
     
    dato1.text = sqlconta.response(0, 3)
    dato2.text = Mid(sqlconta.response(1, 3), 1, 9)
    DV.Caption = Mid(sqlconta.response(1, 3), 10, 1)
    DESDE1.text = Mid(sqlconta.response(2, 3), 1, 2)
    DESDE2.text = Mid(sqlconta.response(2, 3), 4, 2)
    DESDE3.text = Mid(sqlconta.response(2, 3), 7, 4)
    HASTA1.text = Mid(sqlconta.response(3, 3), 1, 2)
    HASTA2.text = Mid(sqlconta.response(3, 3), 4, 2)
    HASTA3.text = Mid(sqlconta.response(3, 3), 7, 4)
    dato3.text = sqlconta.response(4, 3)
    dato4.text = sqlconta.response(5, 3)
    DATO5.text = sqlconta.response(6, 3)
    
    If sqlconta.response(7, 3) = "1" Then
        dato6.text = "S"
    Else
        dato6.text = "N"
    End If
  
    dato7.text = Mid(sqlconta.response(8, 3), 1, 9)
    LBLDVV.Caption = Mid(sqlconta.response(8, 3), 10, 1)
    dato8.text = sqlconta.response(9, 3)
    If IsNull(sqlconta.response(10, 3)) = True Then
       richglosa.text = ""
    Else
'
'
'            palabra = sqlconta.response(10, 3)
'            Grid2.Cell(1, 1).text = Mid(palabra, 1, 95)
'            Grid2.Cell(2, 1).text = Mid(palabra, 96, 95)
'            Grid2.Cell(3, 1).text = Mid(palabra, 193, 95)
'            Grid2.Cell(4, 1).text = Mid(palabra, 289, 95)
'            Grid2.Cell(5, 1).text = Mid(palabra, 385, 95)
'
         richglosa.text = sqlconta.response(10, 3)
    
     End If
    
    lblnombreproveedor.Caption = LEERNOMBREPROVEEDOR(sqlconta.response(1, 3))
    Call dato3_KeyPress(13)
    Call dato4_KeyPress(13)
   
    Call cargameses(dato1.text)
    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    DATO5.Locked = condicion
    dato6.Locked = condicion
    dato7.Locked = condicion
    dato8.Locked = condicion
     
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    
 
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub
 

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
Dim glosafinal As String
Dim k As Double

    campos(0, 0) = "numero"
    campos(1, 0) = "rut"
    campos(2, 0) = "fechainicio"
    campos(3, 0) = "fechatermino"
    campos(4, 0) = "facturar"
    campos(5, 0) = "base"
    campos(6, 0) = "monto"
    campos(7, 0) = "descontado"
    campos(8, 0) = "rutcontacto"
    campos(9, 0) = "nombrecontacto"
    campos(10, 0) = "glosa"
    campos(11, 0) = ""
    
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text & DV.Caption
    campos(2, 1) = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
    campos(3, 1) = HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text
    campos(4, 1) = dato3.text
    campos(5, 1) = dato4.text
    campos(6, 1) = DATO5.text
    campos(7, 1) = dato6.text
    campos(8, 1) = dato7.text & LBLDVV.Caption
    campos(9, 1) = dato8.text
    
     glosafinal = richglosa.text
'    For k = 1 To Grid2.Rows - 1
'       If Grid2.Cell(k, 1).text <> "" Then
'       glosafinal = glosafinal & Grid2.Cell(k, 1).text
'       End If
'    Next k
     campos(10, 1) = glosafinal
     
     campos(0, 2) = "contratopublicidad"
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
     
    
    End Sub
Sub grabar2()
Dim k As Double
Dim año As Double
Dim MES As Double
Dim MESES As Double

    campos(0, 2) = "contratopublicidad_periodos"
    campos(0, 0) = "numerocontrato"
    campos(1, 0) = "mes"
    campos(2, 0) = "pagado"
    campos(3, 0) = ""
    campos(0, 1) = dato1.text
       MES = DESDE2.text
       año = DESDE3.text
       MESES = 1 + DateDiff("m", DESDE1.text & "-" & DESDE2.text & "-" & DESDE3.text, HASTA1.text & "-" & HASTA2.text & "-" & HASTA3.text)
'       meses = DatePart("m", HASTA1.text & "-" & HASTA2.text & "-" & HASTA3.text) - DatePart("m", DESDE1.text & "-" & DESDE2.text & "-" & DESDE3.text)
    For k = 1 To MESES
   Select Case dato3.text

  Case 1
            If MES > 12 Then
            año = año + 1
            MES = 1
            End If
           If (año & "-" & Format(MES, "00") & "-" & DESDE1.text <= HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text) And (año & "-" & Format(MES, "00") & "-" & DESDE1.text <> DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text) Then
                
                    If MES = 2 And CDbl(DESDE1.text) > 28 Then
                        campos(1, 1) = año & "-" & Format(MES, "00") & "-" & Format(CDbl(DESDE1.text) - 2, "00")
                    Else
                        campos(1, 1) = año & "-" & Format(MES, "00") & "-" & DESDE1.text
                    End If
                
           
                campos(2, 1) = 0
                op = 2
                sqlconta.response = campos
                Set sqlconta.conexion = contadb
                Call sqlconta.sqlconta(op, condicion)
            End If
            MES = MES + 1
  Case 2
            If MES > 12 Then
            año = año + 1
            MES = 1
            End If
          If (año & "-" & Format(MES, "00") & "-" & DESDE1.text <= HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text) And (año & "-" & Format(MES, "00") & "-" & DESDE1.text <> DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text) Then
  
            If MES = 2 And CDbl(DESDE1.text) > 28 Then
                campos(1, 1) = año & "-" & Format(MES, "00") & "-" & Format(CDbl(DESDE1.text) - 2, "00")
            Else
                campos(1, 1) = año & "-" & Format(MES, "00") & "-" & DESDE1.text
            End If
          
            campos(2, 1) = 0
            op = 2
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
          End If
            MES = MES + 2
  
  Case 3
            If MES > 12 Then
            año = año + 1
            MES = 1
            End If
            
         If (año & "-" & Format(MES, "00") & "-" & DESDE1.text <= HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text) And (año & "-" & Format(MES, "00") & "-" & DESDE1.text <> DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text) Then
                If MES = 2 And CDbl(DESDE1.text) > 28 Then
                    campos(1, 1) = año & "-" & Format(MES, "00") & "-" & Format(CDbl(DESDE1.text) - 2, "00")
                Else
                    campos(1, 1) = año & "-" & Format(MES, "00") & "-" & DESDE1.text
                End If
            
             campos(2, 1) = 0
            op = 2
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
         End If
            MES = MES + 3
  Case 4
            If MES > 12 Then
            año = año + 1
            MES = 1
            End If
        If (año & "-" & Format(MES, "00") & "-" & DESDE1.text <= HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text) And (año & "-" & Format(MES, "00") & "-" & DESDE1.text <> DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text) Then
  
            If MES = 2 And CDbl(DESDE1.text) > 28 Then
                campos(1, 1) = año & "-" & Format(MES, "00") & "-" & Format(CDbl(DESDE1.text) - 2, "00")
            Else
                campos(1, 1) = año & "-" & Format(MES, "00") & "-" & DESDE1.text
            End If
          
            campos(2, 1) = 0
            op = 2
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
        End If
            MES = MES + 6
  Case 5
            If MES > 12 Then
            año = año + 1
            MES = 1
            End If
     If (año & "-" & Format(MES, "00") & "-" & DESDE1.text <= HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text) And (año & "-" & Format(MES, "00") & "-" & DESDE1.text <> DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text) Then

            If MES = 2 And CDbl(DESDE1.text) > 28 Then
                campos(1, 1) = año & "-" & Format(MES, "00") & "-" & Format(CDbl(DESDE1.text) - 2, "00")
            Else
                campos(1, 1) = año & "-" & Format(MES, "00") & "-" & DESDE1.text
            End If
       
            campos(2, 1) = 0
            op = 2
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
     End If
            MES = MES + 12
  End Select
  Next k
    
End Sub

Sub ELIMINAR()
    campos(0, 2) = "contratopublicidad"
    condicion = "numero=" + "'" + dato1.text + "' and rut=" + "'" + dato2.text + DV.Caption + "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    campos(0, 2) = "contratopublicidad_periodos"
    condicion = "numerocontrato=" + "'" + dato1.text + "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub
 
Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then modifica
If command = "imprime" Then imprimircontrato

If command = "elimina" Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        ELIMINAR
        retorno
    Else
        MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
    End If
End If

If command = "siguiente" Then leer (">")
If command = "anterior" Then leer ("<")
 

End Sub
Sub ELIMINA()
 
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus
 
End Sub

Sub modifica()
disponible (True)
habilita (False)
dato1.Enabled = False
dato2.Enabled = False
dato4.SetFocus
MODIFI = 1

End Sub
Sub retorno()

disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
MODIFI = 0
no:
Grid1.Rows = 1
Call CARGAGRILLAVIGENTES
Call leercontratosvigentes("")
    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    DV.Caption = ""
    DESDE1.text = ""
    DESDE2.text = ""
    DESDE3.text = ""
    HASTA1.text = ""
    HASTA2.text = ""
    HASTA3.text = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = ""
    LBLDVV.Caption = ""
    dato8.text = ""
    richglosa.text = ""
    lblnombreproveedor.Caption = ""
    lblfacturar.Caption = ""
    lblbase.Caption = ""
    frmbase.Visible = False
    frmTipo.Visible = False
    Grid2.Rows = 1
    Call CARGAGRILLA(20, 2)
End Sub
  
 

Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus

End Sub

Sub ayudaProveedor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='23100026' and año='" & Format(fechasistema, "yyyy") & "' "
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Proveedores"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then dato2.SetFocus: GoTo no
    dato4.Enabled = True
    dato2.text = Mid(pivote.text, 1, 9)
    DV.Caption = Mid(pivote.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub
Sub ayudacontratos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("m.rut", "m.nombre", "p.numero", "p.fechatermino")
    largo = Array("12n", "40s", "12n", "10d")
    cfijo = "m.tipo='23100026' and m.año='" & Format(fechasistema, "yyyy") & "' "
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Proveedores"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes as m,", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then dato2.SetFocus: GoTo no
    dato4.Enabled = True
    dato2.text = Mid(pivote.text, 1, 9)
    DV.Caption = Mid(pivote.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub
 
Sub CARGAGRILLAMESES()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "PAGADOS"
    formatogrilla2(1, 2) = "FECHAS "
    Rem LARGO DE LOS DATOS
    
    formatogrilla2(2, 1) = "10"
    formatogrilla2(2, 2) = "20"
 
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    Rem VALOR MAXIMO
    
    Grid1.Cols = 3
    Grid1.Rows = 13
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
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        Grid1.Column(1).CellType = cellCheckBox
    Next k
 
    End Sub
    Sub CARGAGRILLAVIGENTES()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "NUMERO"
    formatogrilla2(1, 2) = "FECHA"
    formatogrilla2(1, 3) = "PROVEEDOR"
    Rem LARGO DE LOS DATOS
    
    formatogrilla2(2, 1) = "8"
    formatogrilla2(2, 2) = "8"
    formatogrilla2(2, 3) = "20"
 
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "D"
    formatogrilla2(3, 3) = "S"
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    Rem VALOR MAXIMO
    
    Grid3.Cols = 4
    Grid3.Rows = 13
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    Grid3.BackColorFixed = RGB(90, 158, 214)
    Grid3.BackColorFixedSel = RGB(110, 180, 230)
    Grid3.BackColorBkg = RGB(90, 158, 214)
    Grid3.BackColorScrollBar = RGB(231, 235, 247)
    Grid3.BackColor1 = RGB(231, 235, 247)
    Grid3.BackColor2 = RGB(239, 243, 255)
    Grid3.GridColor = RGB(148, 190, 231)
    Grid3.Column(0).Width = 0
    
    For k = 1 To Grid3.Cols - 1
        Grid3.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid3.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid3.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid3.Column(k).FormatString = formatogrilla2(4, k)
        Grid3.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid3.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "D" Then Grid3.Column(k).CellType = cellCalendar
  
    Next k
 Grid3.SelectionMode = cellSelectionByRow
    End Sub

Private Sub DESDE1_GotFocus()
    Call cargatexto(DESDE1)
End Sub

Private Sub DESDE1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE1, DESDE2, KeyCode)
End Sub

Private Sub DESDE1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    If DESDE1.text = "" Then DESDE1.text = Format(fechasistema, "dd")
    Call ceros(DESDE1)
    DESDE2.SetFocus
    Call esfechareal(DESDE1, DESDE2, DESDE3, "dd")
    End If
    
End Sub

Private Sub DESDE2_GotFocus()
    Call cargatexto(DESDE2)
End Sub

Private Sub DESDE2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE1, DESDE3, KeyCode)
End Sub

Private Sub DESDE2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And DESDE1.text <> "" Then
    If DESDE2.text = "" Then DESDE2.text = Format(fechasistema, "mm")
    Call ceros(DESDE2): DESDE3.SetFocus
Call esfechareal(DESDE1, DESDE2, DESDE3, "mm")
End If
End Sub

Private Sub DESDE3_GotFocus()
    Call cargatexto(DESDE3)
End Sub

Private Sub DESDE3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DESDE2, DESDE3, KeyCode)
End Sub

Private Sub DESDE3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And DESDE2.text <> "" Then
    If DESDE3.text = "" Then DESDE3.text = Format(fechasistema, "yyyy")
    Call ceros(DESDE3): HASTA1.SetFocus
    Call esfechareal(DESDE1, DESDE2, DESDE3, "yyyy")
End If
End Sub
Private Sub HASTA1_GotFocus()
    Call cargatexto(HASTA1)
End Sub

Private Sub HASTA1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(HASTA1, HASTA2, KeyCode)
End Sub

Private Sub HASTA1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    If HASTA1.text = "" Then HASTA1.text = Format(fechasistema, "dd")
    Call ceros(HASTA1)
    HASTA2.SetFocus
    Call esfechareal(HASTA1, HASTA2, HASTA3, "dd")
End If

End Sub

Private Sub HASTA2_GotFocus()
    Call cargatexto(HASTA2)
End Sub

Private Sub HASTA2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(HASTA1, HASTA3, KeyCode)
End Sub

Private Sub HASTA2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And HASTA1.text <> "" Then
    If HASTA2.text = "" Then HASTA2.text = Format(fechasistema, "mm")
    Call ceros(HASTA2): HASTA3.SetFocus
Call esfechareal(HASTA1, HASTA2, HASTA3, "mm")
End If
End Sub

Private Sub HASTA3_GotFocus()
    Call cargatexto(HASTA3)
End Sub

Private Sub HASTA3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(HASTA2, HASTA3, KeyCode)
End Sub

Private Sub HASTA3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And HASTA2.text <> "" Then
    If HASTA3.text = "" Then HASTA3.text = Format(fechasistema, "yyyy")
    Call ceros(HASTA3): dato3.Enabled = True: dato3.SetFocus
    Call esfechareal(HASTA1, HASTA2, HASTA3, "yyyy")
End If
End Sub

Private Function leerultimocontrato() As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = contadb
csql.sql = "select ifnull(max(numero + 1),1) from " + clientesistema + "conta" + empresaactiva + ".contratopublicidad "
csql.Execute
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leerultimocontrato = Format(resultados(0), "0000000000")
Else
    leerultimocontrato = "0000000001"
End If

End Function
Sub cargameses(numero)
   Dim csql As New rdoQuery
   Dim resultados As rdoResultset
   Dim i As Double
   
   
   Set csql.ActiveConnection = contadb
   csql.sql = "select mes,pagado from contratopublicidad_periodos "
   csql.sql = csql.sql & "where numerocontrato='" & numero & "'"
   csql.Execute
   
   If csql.RowsAffected > 0 Then
    Grid1.Rows = csql.RowsAffected + 1
    Set resultados = csql.OpenResultset
    i = 1
    While Not resultados.EOF
        Grid1.Cell(i, 1).text = resultados(1)
        Grid1.Cell(i, 2).text = Mid(resultados(0), 4, 10)
        i = i + 1
        resultados.MoveNext
    Wend
   
   End If
   csql.Close
   Set csql = Nothing
   Set resultados = Nothing
   
   
End Sub
 
Sub imprimircontrato()
    
    Dim row As Integer
    Dim FINROW As Integer
    Dim i As Integer
    Dim nombrerepresentante As String
    Dim profesion As String
    Dim rutrepresentante As String
    Dim nacionalidad As String
    Dim direccionproveedor As String
    Dim ciudadproveedor As String
    Dim k As Double
    Dim o As Double
    Dim rutempresa As String
    
     rutempresa = DATOSEMPRESA(5)
     rutempresa = Replace(rutempresa, "-", "")
     rutempresa = String(10 - Len(rutempresa), "0") & rutempresa
    
    nombrerepresentante = "JUAN PATRICIO ELTIT JADUE"
    profesion = "EMPRESARIO"
    rutrepresentante = "7.762.388-4"
    nacionalidad = "chilena"
    direccionproveedor = LEERdireccionproveedor(dato2.text & DV.Caption)
    ciudadproveedor = LEERciudadproveedor(dato2.text & DV.Caption)
    Dim objReportTitle As FlexCell.ReportTitle
    impresioncontrato.Rows = 1
    impresioncontrato.Cols = 10
    impresioncontrato.Column(8).Width = 130
    impresioncontrato.PageSetup.BlackAndWhite = False
    impresioncontrato.PageSetup.BottomMargin = 1
    impresioncontrato.PageSetup.LeftMargin = 1
    impresioncontrato.PageSetup.RightMargin = 1
    impresioncontrato.PageSetup.TopMargin = 1
    impresioncontrato.PageSetup.PrintFixedRow = True
    impresioncontrato.Column(1).Width = 13 * 8
    Call cabeza
 
    FINROW = impresioncontrato.Rows
 
    impresioncontrato.Rows = impresioncontrato.Rows + 50
    impresioncontrato.Range(FINROW + 1, 1, FINROW + 1, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 1, 1, FINROW + 1, 9).Merge
    impresioncontrato.Range(FINROW + 1, 1, FINROW + 1, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 1, 1).text = "      En " & DATOSEMPRESA(3) & "  a  " & Format(fechasistema, "dd") & "   de  " & MonthName(Format(fechasistema, "mm")) & "  de  " & Format(fechasistema, "yyyy") & " entre don " & nombrerepresentante & ". De profesión " & profesion & ""
 
    impresioncontrato.Range(FINROW + 2, 1, FINROW + 2, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 2, 1, FINROW + 2, 9).Merge
    impresioncontrato.Range(FINROW + 2, 1, FINROW + 2, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 2, 1).text = "C.I. " & rutrepresentante & "      " & nacionalidad & " en representación de       " & DATOSEMPRESA(1) & "  R.U.T  : " & Format(Mid(rutempresa, 1, 9), "###,###,###") & "-" & Mid(rutempresa, 10, 1) & " ambos "
        
    impresioncontrato.Range(FINROW + 3, 1, FINROW + 3, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 3, 1, FINROW + 3, 9).Merge
    impresioncontrato.Range(FINROW + 3, 1, FINROW + 3, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 3, 1).text = "domiciliados en  " & DATOSEMPRESA(2) & " de la ciudad " & DATOSEMPRESA(3) & " en adelante el (CLIENTE). "

    impresioncontrato.Range(FINROW + 6, 1, FINROW + 6, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 6, 1, FINROW + 6, 9).Merge
    impresioncontrato.Range(FINROW + 6, 1, FINROW + 6, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 6, 1).text = "      Y Don(a) " & dato8.text & " C.I. " & Format(dato7.text, "###,###,###") & "-" & LBLDVV.Caption & "  en Representación de " & lblnombreproveedor.Caption
 
    impresioncontrato.Range(FINROW + 7, 1, FINROW + 7, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 7, 1, FINROW + 7, 9).Merge
    impresioncontrato.Range(FINROW + 7, 1, FINROW + 7, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 7, 1).text = "R.U.T : " & Format(dato2.text, "###,###,###") & "-" & DV.Caption & " domiciliado en " & direccionproveedor & " Ciudad de " & ciudadproveedor
    
    impresioncontrato.Range(FINROW + 8, 1, FINROW + 8, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 8, 1, FINROW + 8, 9).Merge
    impresioncontrato.Range(FINROW + 8, 1, FINROW + 8, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 8, 1).text = "en adelante (PROVEEDOR). Acuerdan el siguiente contrato."
    
    impresioncontrato.Range(FINROW + 10, 1, FINROW + 10, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 10, 1, FINROW + 10, 9).Merge
    impresioncontrato.Range(FINROW + 10, 1, FINROW + 10, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 10, 1).text = "EL PROVEEDOR  :  Se  compromete  a  otorgar   un   aporte  publicitario  al  CLIENTE  equivalente   a  la  cantidad  "
    
    impresioncontrato.Range(FINROW + 11, 1, FINROW + 11, 3).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 11, 1, FINROW + 11, 3).Merge
    impresioncontrato.Range(FINROW + 11, 1, FINROW + 11, 3).FontSize = 10
    impresioncontrato.Cell(FINROW + 11, 1).text = lblbase.Caption & "  " & Format(DATO5.text, "###,###,###") & "   en forma "
 
    impresioncontrato.Range(FINROW + 11, 4, FINROW + 11, 4).FontBold = True
    impresioncontrato.Cell(FINROW + 11, 4).text = lblfacturar.Caption & ". "
    
    impresioncontrato.Range(FINROW + 13, 1, FINROW + 13, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 13, 1, FINROW + 13, 9).Merge
    impresioncontrato.Range(FINROW + 13, 1, FINROW + 13, 9).FontSize = 10
    impresioncontrato.Range(FINROW + 13, 1, FINROW + 13, 9).FontBold = True
    impresioncontrato.Cell(FINROW + 13, 1).text = "NOTA : ** LOS APORTES SON MAS I.V.A *** "
    
    impresioncontrato.Range(FINROW + 15, 1, FINROW + 15, 6).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 15, 1, FINROW + 15, 6).Merge
    impresioncontrato.Range(FINROW + 15, 1, FINROW + 15, 6).FontSize = 10
    impresioncontrato.Range(FINROW + 15, 1, FINROW + 15, 6).FontBold = False
    impresioncontrato.Cell(FINROW + 15, 1).text = "El  Pago de este aporte  se rebajará  del  pago de  las  facturas de  compras  en forma "
    impresioncontrato.Range(FINROW + 15, 7, FINROW + 15, 7).FontBold = True
    impresioncontrato.Cell(FINROW + 15, 7).text = lblfacturar.Caption & "."
    
    
    impresioncontrato.Range(FINROW + 18, 1, FINROW + 18, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 18, 1, FINROW + 18, 9).Merge
    impresioncontrato.Range(FINROW + 18, 1, FINROW + 18, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 18, 1).text = "El  presente  contrato  comenzará  a regir  a  partir  del  día " & DESDE1.text & "-" & DESDE2.text & "-" & DESDE3.text & " y se  le  dara  termino  el  día " & HASTA1.text & "-" & HASTA2.text & "-" & HASTA3.text

    impresioncontrato.Range(FINROW + 19, 1, FINROW + 19, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 19, 1, FINROW + 19, 9).Merge
    impresioncontrato.Range(FINROW + 19, 1, FINROW + 19, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 19, 1).text = " "

    impresioncontrato.Range(FINROW + 21, 1, FINROW + 21, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 21, 1, FINROW + 21, 9).Merge
    impresioncontrato.Range(FINROW + 21, 1, FINROW + 21, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 21, 1).text = "El   presente   contrato   sera   renovado   indefinidamente    si  ninguna    de  las  partes   manifiesta  lo contrario."
    
    impresioncontrato.Range(FINROW + 22, 1, FINROW + 22, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + 22, 1, FINROW + 22, 9).Merge
    impresioncontrato.Range(FINROW + 22, 1, FINROW + 22, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + 22, 1).text = "Con una anterioridad de 30 días antes del vencimiento de este."
     
'o = 0
'Asc(UCase(Chr
'For k = 1 To Grid2.Rows - 1
'
'    If Grid2.Cell(k, 1).text <> "" Then
'        impresioncontrato.Range(FINROW + (24 + o), 1, FINROW + (24 + o), 9).Borders(cellInsideHorizontal) = cellThin
'        impresioncontrato.Range(FINROW + (24 + o), 1, FINROW + (24 + o), 9).Merge
'        impresioncontrato.Range(FINROW + (24 + o), 1, FINROW + (24 + o), 9).FontSize = 10
'        impresioncontrato.Cell(FINROW + (24 + o), 1).text = Grid2.Cell(k, 1).text
'        o = o + 1
'    End If
'Next k
o = 24
        impresioncontrato.Range(FINROW + (24), 1, FINROW + (24), 9).Borders(cellInsideHorizontal) = cellThin
        impresioncontrato.Range(FINROW + (24), 1, FINROW + (24), 9).Merge
        impresioncontrato.Range(FINROW + (24), 1, FINROW + (24), 9).FontSize = 10
        impresioncontrato.Cell(FINROW + (24), 1).text = richglosa.text


 o = FINROW + (24 + o)
 
    impresioncontrato.Range(FINROW + o, 1, FINROW + o, 9).Borders(cellInsideHorizontal) = cellThin
    impresioncontrato.Range(FINROW + o, 1, FINROW + o, 9).Merge
    impresioncontrato.Range(FINROW + o, 1, FINROW + o, 9).FontSize = 10
    impresioncontrato.Cell(FINROW + o, 1).text = "Las partes previa lectura firman en señal de aceptación"
 
'
    o = 45 + 1
    impresioncontrato.Range(FINROW + o, 2, FINROW + o, 3).Merge
    impresioncontrato.Range(FINROW + o, 2, FINROW + o, 3).Alignment = cellCenterCenter
    impresioncontrato.Cell(FINROW + o, 2).text = nombrerepresentante
    
    impresioncontrato.Range(FINROW + o, 6, FINROW + o, 7).Merge
    impresioncontrato.Range(FINROW + o, 6, FINROW + o, 7).Alignment = cellCenterCenter
'    impresioncontrato.Range(FINROW + o, 4, FINROW + o, 6).Borders(cellEdgeTop) = cellThin
    impresioncontrato.Cell(FINROW + o, 6).text = dato8.text
    
    o = o + 1
    impresioncontrato.Range(FINROW + o, 2, FINROW + o, 3).Merge
    impresioncontrato.Range(FINROW + o, 2, FINROW + o, 3).Alignment = cellCenterCenter
    impresioncontrato.Cell(FINROW + o, 2).text = rutrepresentante
    
    impresioncontrato.Range(FINROW + o, 6, FINROW + o, 7).Merge
    impresioncontrato.Range(FINROW + o, 6, FINROW + o, 7).Alignment = cellCenterCenter
'    impresioncontrato.Range(FINROW + o, 4, FINROW + o, 6).Borders(cellEdgeTop) = cellThin
    impresioncontrato.Cell(FINROW + o, 6).text = Format(dato7.text, "###,###,###") & "-" & LBLDVV.Caption
    
    
    impresioncontrato.PageSetup.BlackAndWhite = True
'     impresioncontrato.PageSetup.PrintGridlines = True
     impresioncontrato.Column(9).Width = 0
     impresioncontrato.PrintPreview
'    impresioncontrato.Rows = FINROW
End Sub
Sub cabeza()
Dim k As Integer
Dim rutempresa As String

    Dim objReportTitle As FlexCell.ReportTitle
    
    impresioncontrato.ReportTitles.Clear
    'Report Title 1
    For k = 1 To 5
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "verdana"
        objReportTitle.Font.Size = 7
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        objReportTitle.color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
         If k = 5 Then
            rutempresa = DATOSEMPRESA(k)
            rutempresa = Replace(rutempresa, "-", "")
            rutempresa = String(10 - Len(rutempresa), "0") & rutempresa
            objReportTitle.text = Format(Mid(rutempresa, 1, 9), "###,###,###") & "-" & Mid(rutempresa, 10, 1)
        End If
        impresioncontrato.ReportTitles.Add objReportTitle
       
    Next k
 
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CONTRATO APORTE PUBLICITARIO "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 12
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    impresioncontrato.ReportTitles.Add objReportTitle
    
'    impresioncontrato.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
'    impresioncontrato.PageSetup.FooterAlignment = cellRight
'    impresioncontrato.PageSetup.FooterFont.Name = "Verdana"
'    impresioncontrato.PageSetup.FooterFont.Size = 7
    With impresioncontrato.PageSetup
        .HeaderFont.Size = 6
        '.Header = "                                                                                                                   PAGINAS &P/&N EMITIDO:&D USUARIO " + USUARIOSISTEMA
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Verdana"
        .HeaderMargin = 4
    End With
 End Sub
    
Sub CARGAGRILLA(row, col)
Dim FORMATOGRILLA(10, 10) As String
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "DETALLE DEL CONTRATO"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "150"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "false"
    
    Grid2.Cols = col
    Grid2.Rows = row
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.Column(0).Width = 0
    For k = 1 To col - 1
        Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid2.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid2.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid2.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then
            Grid2.Column(k).Alignment = cellRightCenter
            Grid2.Column(k).Mask = cellNumeric
        End If
        If FORMATOGRILLA(3, k) = "S" Then
            Grid2.Column(k).Alignment = cellLeftCenter
            Grid2.Column(k).Mask = cellUpper
        End If
        If FORMATOGRILLA(3, k) = "D" Then
            Grid2.Column(k).CellType = cellCalendar
            Grid2.Column(k).Mask = cellNumeric
        End If
        
        'Grid1.Column(7).CellType = cellComboBox
    Next k
    Grid2.Range(0, 1, 0, 1).FontSize = 12
    Grid2.Range(0, 1, 0, 1).FontBold = True
    Grid2.Range(0, 1, 0, 1).Alignment = cellCenterCenter
    Grid2.Column(1).Width = 500
    
End Sub

 

Private Sub richglosa_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub
Sub leercontratosvigentes(rutproveedor)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
        Set csql.ActiveConnection = contadb
        csql.sql = "select  numero,fechainicio,rut from " + clientesistema + "conta" + empresaactiva + ".contratopublicidad "
        csql.sql = csql.sql & "where rut like '%" & rutproveedor & "%' order by rut"
        csql.Execute
        If csql.RowsAffected > 0 Then
            Grid3.Rows = 1
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Grid3.Rows = Grid3.Rows + 1
                Grid3.Cell(Grid3.Rows - 1, 1).text = resultados(0)
                Grid3.Cell(Grid3.Rows - 1, 2).text = resultados(1)
                Grid3.Cell(Grid3.Rows - 1, 3).text = LEERNOMBREPROVEEDOR(resultados(2))
                resultados.MoveNext
            Wend
        End If
        csql.Close
        Set csql = Nothing
        Set resultados = Nothing
    
End Sub

Public Sub cargadeafueracontrato()
    Call dato1_KeyPress(13)
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
