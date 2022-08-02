VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Maestrofletes 
   Caption         =   "Maestro Fletes"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8730
      Left            =   -45
      TabIndex        =   14
      Top             =   -45
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15399
      BackColor       =   12632256
      Caption         =   "Datos Documento"
      CaptionEstilo3D =   2
      BackColor       =   12632256
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox desde3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8025
         MaxLength       =   4
         TabIndex        =   45
         Tag             =   "proveedor"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox desde2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7665
         MaxLength       =   2
         TabIndex        =   44
         Tag             =   "proveedor"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox desde1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7305
         MaxLength       =   2
         TabIndex        =   43
         Tag             =   "proveedor"
         Top             =   600
         Width           =   375
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000FF00&
         Caption         =   "FISCAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7920
         TabIndex        =   42
         Top             =   270
         Width           =   1230
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000FF00&
         Caption         =   "OFFSET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6525
         TabIndex        =   41
         Top             =   270
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.TextBox ncaja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5625
         MaxLength       =   2
         TabIndex        =   39
         Tag             =   "proveedor"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8520
         MaxLength       =   9
         TabIndex        =   35
         Tag             =   "proveedor"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox dato13 
         Appearance      =   0  'Flat
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "proveedor"
         Top             =   2040
         Width           =   4770
      End
      Begin VB.CommandButton guardar 
         BackColor       =   &H00FF8080&
         Caption         =   "Guardar"
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
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7080
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8520
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   1305
         Width           =   4815
      End
      Begin XPFrame.FrameXp frmTipo 
         Height          =   1200
         Left            =   2280
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2117
         BackColor       =   8421631
         Caption         =   "Tipo Documento"
         CaptionEstilo3D =   1
         BackColor       =   8421631
         ColorBarraArriba=   12632319
         ColorBarraAbajo =   128
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
         Begin VB.Label lbl23 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 1 - Factura"
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
            TabIndex        =   26
            Top             =   495
            Width           =   2475
         End
         Begin VB.Label lbl24 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 3 - NC Factura"
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
            TabIndex        =   25
            Top             =   1200
            Width           =   2475
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 4 - NC Boleta"
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
            TabIndex        =   24
            Top             =   1560
            Width           =   2475
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 2 - Boleta"
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
            TabIndex        =   23
            Top             =   840
            Width           =   2475
         End
      End
      Begin VB.TextBox dato11 
         Appearance      =   0  'Flat
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
         Left            =   8520
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   1680
         Width           =   4815
      End
      Begin VB.TextBox dato10 
         Appearance      =   0  'Flat
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
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   1680
         Width           =   4815
      End
      Begin VB.TextBox dato8 
         Appearance      =   0  'Flat
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
         Left            =   1755
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox dato7 
         Appearance      =   0  'Flat
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
         Left            =   5490
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   960
         Width           =   7710
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1740
         MaxLength       =   9
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox dato1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   9705
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   12105
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   12465
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   12825
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   600
         Width           =   615
      End
      Begin XPFrame.FrameXp frmDetalle 
         Height          =   4455
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   7858
         BackColor       =   16744576
         Caption         =   "Lista de Productos"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin FlexCell.Grid detalle 
            Height          =   4185
            Left            =   0
            TabIndex        =   28
            Top             =   360
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   7382
            Cols            =   6
            DefaultFontSize =   8.25
            Rows            =   15
         End
      End
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   -360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Doc."
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
         Left            =   6120
         TabIndex        =   46
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Caja"
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
         Left            =   4500
         TabIndex        =   40
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblvendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   10440
         TabIndex        =   38
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label lbldvvendedor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   9960
         TabIndex        =   37
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vendedor"
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
         Left            =   6720
         TabIndex        =   36
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Giro"
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
         Left            =   0
         TabIndex        =   34
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fono"
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
         Left            =   6720
         TabIndex        =   32
         Top             =   1320
         Width           =   1695
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
         Height          =   1755
         Left            =   270
         TabIndex        =   13
         Top             =   6885
         Visible         =   0   'False
         Width           =   6960
         _cx             =   12277
         _cy             =   3096
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
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Cliente"
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
         Left            =   0
         TabIndex        =   31
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Razón Social"
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
         Left            =   3765
         TabIndex        =   29
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo"
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
         Left            =   0
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblDocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
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
         Left            =   6720
         TabIndex        =   19
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lbl7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comuna"
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
         Left            =   0
         TabIndex        =   18
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lbl3 
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
         Left            =   11400
         TabIndex        =   17
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dirección"
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
         Left            =   0
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Numero "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8880
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Maestrofletes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Private numerocaja As String
  
  Private modifica As Boolean
  Private lectura As Boolean
  Private tipoprecio As String
  Private fecha As String
  Private formatogrilla(10, 10) As String
  Private numeroDoc As String
  Private numerofle As String
  Private horaorden As String
  Private final As Double
  Private glosafletes(20) As String
  Private fechadoc As String
  
  
Private Sub DESDE1_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            DESDE1.text = ceros(DESDE1)
            If DESDE1.text = "00" Then
                DESDE1.text = Format(fechasistema, "dd")
                DESDE2.text = Format(fechasistema, "mm")
                DESDE3.text = Format(fechasistema, "yyyy")
                fechadoc = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
             
            End If
             DESDE2.SetFocus
        End If
End Sub

 

Private Sub DESDE2_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            DESDE2.text = ceros(DESDE2)
            If DESDE2.text = "00" Then
                DESDE2.text = Format(fechasistema, "mm")
                DESDE3.text = Format(fechasistema, "yyyy")
                fechadoc = DESDE1.text & "-" & DESDE2.text & "-" & DESDE3.text
            End If
            DESDE3.SetFocus
        End If
End Sub

Private Sub DESDE3_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            DESDE3.text = ceros(DESDE3)
            If DESDE3.text = "0000" Then
                DESDE3.text = Format(fechasistema, "yyyy")
                fechadoc = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
                Else
                fechadoc = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
              End If
              If IsDate(fechadoc) = True Then
                dato2.Enabled = True
                dato2.SetFocus
              Else
                MsgBox "DEBE INGRESAR UNA FECHA VALIDA", vbCritical, "ATENCION"
                DESDE1.text = ""
                DESDE2.text = ""
                DESDE3.text = ""
                fechadoc = ""
                DESDE1.SetFocus
              End If
          
            
        End If
End Sub

 

Private Sub ncaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And ncaja.text <> "" Then
ncaja.text = Format(ncaja.text, "00")
 DESDE1.Enabled = True
 DESDE1.SetFocus
End If

End Sub

 Private Sub dato1_GotFocus()
        lectura = False
        modifica = False
        frmTipo.Visible = True
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



 Private Sub DATO7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
    End Sub
Private Sub DATO8_GotFocus()
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
    End Sub
Private Sub dato11_GotFocus()
        Call VerificarCajas(Me, dato11)
        Call selecciona(dato11)
    End Sub
    Private Sub dato13_GotFocus()
'        Call VerificarCajas(Me, dato13)
         Call selecciona(dato13)
    End Sub
  Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 97, 49
                dato1.text = "FV"
                tipoprecio = "01"
            Case 98, 50
                dato1.text = "BV"
                tipoprecio = "01"
'            Case 99, 51
'                dato1.text = "NF"
'                tipoprecio = "01"
'            Case 99, 52
'                dato1.text = "NB"
'                tipoprecio = "01"
'           Case 99, 53
'                dato1.text = "CO"
'                tipoprecio = "01"
            Case Else
                Call Flechas(KeyCode, dato1)
        End Select
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
     Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            lblDocumento.Caption = leerNombreDocumento(dato1.text)
            If lblDocumento.Caption <> "" Then
'                Select Case dato1.text
'
'                    Case Else
'
'                End Select
               ncaja.Enabled = True
               ncaja.SetFocus
'              SendKeys "{Tab}"
            End If
            tipo_doc = dato1.text
        End If
    End Sub
     Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If Val(dato2.text) = 0 Then
                Call dato1_KeyPress(13)
                Exit Sub
            End If
            If leerfletes(dato1.text, dato2.text, ncaja.text, fechadoc) = True Then
                lectura = True
                dato6_KeyPress (13)
                dato12_KeyPress (13)
                dato3.text = Format(fechasistema, "dd")
                dato4.text = Format(fechasistema, "mm")
                dato5.text = Format(fechasistema, "yyyy")
                If exiteflete(dato1.text, dato2.text) = True Then
                guardar.Visible = False
                opciones.Visible = True
                opciones.SetFocus
                Else
                If detalle.Rows > 1 Then
                guardar.Visible = True
                End If
                End If
            Call leerdetalle(dato1.text, numeroDoc, numerocaja)
            Else
            
            If Verifica_Permiso(Me.Caption, "agrega") = True Then
                lectura = False
                detalle.SelectionMode = cellSelectionFree
                'If detalle.Rows <= 1 Then
                    detalle.Rows = 1
                    detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
                'End If
                Call HabilitarCajas(Me, modifica)
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                dato2.SelStart = 0
                dato2.SelLength = Len(dato2.text)
                dato2.SetFocus
            End If
            
            
            

            End If
            numero_doc = dato2.text
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "00" Then
                dato3.text = Format(fechasistema, "dd")
                dato4.text = Format(fechasistema, "mm")
                dato5.text = Format(fechasistema, "yyyy")
                'dato4.Enabled = True
                'dato5.Enabled = True
                fecha = dato3.text & "-" & dato4.text & "-" & dato5.text
               If rut_cliente & lbldv.Caption = "0000000019" Then
               dato6.SetFocus
               Else
               
               End If
            End If
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "0000" Then
                dato5.text = Format(fechasistema, "yyyy")
                            
            End If
            fecha = dato3.text & "-" & dato4.text & "-" & dato5.text
                If rut_cliente & lbldv.Caption <> "0000000019" Then
               Call dato2_KeyPress(13)
               
               dato6.SetFocus
               
               Else
               
               End If
        End If
    
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato6.text <> "" And Val(dato6.text) <> 0 Then
             dato6.text = ceros(dato6)
             lbldv.Caption = rut(dato6.text)
             rut_cliente = dato6.text
             If rut_cliente & lbldv.Caption <> "0000000019" Then
             Call LeerClienteFlete(rut_cliente & lbldv.Caption, "0")
             guardar.Visible = True
             
             Else
             dato6.SetFocus
             End If
'            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub DATO7_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
        dato8.SetFocus
        End If
    End Sub
     Private Sub DATO8_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
        dato9.SetFocus
        End If
    End Sub
    Private Sub DATO9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
           dato10.SetFocus
        End If
    End Sub
     Private Sub DATO10_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           dato11.SetFocus
        End If
    End Sub
     Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           dato13.SetFocus
        End If
    End Sub
    Private Sub dato12_KeyPress(KeyAscii As Integer)
      
    End Sub
    Private Sub dato13_KeyPress(KeyAscii As Integer)
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           Call grabarclienteflete
           guardar.Visible = True
           guardar.SetFocus
        End If
    End Sub
    Private Sub dato1_LostFocus()
        frmTipo.Visible = False
    End Sub
    
    Private Sub dato3_LostFocus()
    Call esfecha(dato3, dato4, dato5, "dd")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato3, dato4, dato5, "mm")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato3, dato4, dato5, "yyyy")
    End Sub
    Private Sub dato7_LostFocus()
        Call limpiaBarra(2)
    End Sub
    Private Sub dato9_LostFocus()
'        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato10_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Function leerfletes(ByVal TIPO As String, ByVal NUMERO As String, ByVal caja As String, ByVal fechadoc As String) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim tabla As String
    
    Set csql.ActiveConnection = ventasRubro
    tabla = ""
    tabla = "select * from sv_documento_cabeza_" & empresaActiva & " "
    If Option1.Value = True Then tabla = tabla & "where local='" & empresaActiva & "' and tipo='" & TIPO & "' and numero='" & NUMERO & "' and caja='" + caja + "' and fecha ='" & fechadoc & "' "
    If Option2.Value = True Then tabla = tabla & "where local='" & empresaActiva & "' and tipo='" & TIPO & "' and foliosii='" & NUMERO & "' and caja='" + caja + "' and fecha ='" & fechadoc & "' "
    
    csql.sql = tabla
    csql.Execute
    leerfletes = False
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    dato6.text = resultados("rut")
    numerocaja = resultados("caja")
    numeroDoc = resultados("numero")
    leerfletes = True
    End If
    
    End Function
   Sub leerdetalle(TIPO, NUMERO, caja)
   Dim csql As New rdoQuery
   Dim resultados As rdoResultset
   Dim tabla As String
   Dim linea As Double
   
   Set csql.ActiveConnection = ventasRubro
   tabla = ""
   tabla = "select codigo,descripcion,cantidad,tipodespacho,precio,vendedor "
   tabla = tabla & "from sv_documento_detalle_" & empresaActiva
   tabla = tabla & " where caja='" + caja + "' and tipo='" & TIPO & "' and numero='" & NUMERO & "' and local='" & empresaActiva & "' and tipodespacho='05'and tipodespacho<>'' order by linea"
   csql.sql = tabla
   csql.Execute
   linea = 1
   If csql.RowsAffected > 0 Then
   Set resultados = csql.OpenResultset
   
   detalle.Rows = csql.RowsAffected + 1
   dato14.text = Mid(resultados("vendedor"), 1, 9)
   lbldvvendedor.Caption = Mid(resultados("vendedor"), 10, 1)
   lblVendedor.Caption = leerNombreVendedor(resultados("vendedor"))
   
  While Not resultados.EOF
 
    detalle.Cell(linea, 1).text = resultados("codigo")
    detalle.Cell(linea, 2).text = resultados("descripcion")
    detalle.Cell(linea, 3).text = resultados("cantidad")
    detalle.Cell(linea, 4).text = "0"
    detalle.Cell(linea, 5).text = resultados("cantidad")
    detalle.Cell(linea, 6).text = resultados("precio")
    resultados.MoveNext
    linea = linea + 1
  Wend
   End If
   csql.Close
   Set resultados = Nothing
   Set csql = Nothing
 
    
    
   End Sub
 Private Sub CARGAGRILLA(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 0) = "LN"
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "DESCRIPCION"
        formatogrilla(1, 3) = "FLETE"
        formatogrilla(1, 4) = "ENTREGADA"
        formatogrilla(1, 5) = "X DESPACHAR"
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "15"
        formatogrilla(2, 2) = "40"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
       
              
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "###,###,##0"
        formatogrilla(4, 4) = "###,###,##0"
        formatogrilla(4, 5) = "###,###,##0"
   
       
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"

      
       
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
       
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
   
       
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "46"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "10"
        formatogrilla(8, 5) = "12"
            
        detalle.Cols = 7
        detalle.Rows = row
        detalle.AllowUserResizing = False
        detalle.DisplayFocusRect = False
        detalle.ExtendLastCol = False
        detalle.BoldFixedCell = False
        detalle.DisplayRowIndex = True
        detalle.DrawMode = cellOwnerDraw
        detalle.Appearance = Flat
        detalle.ScrollBarStyle = Flat
        detalle.FixedRowColStyle = Flat
        detalle.BackColorFixed = RGB(90, 158, 214)
        detalle.BackColorFixedSel = RGB(110, 180, 230)
        detalle.BackColorBkg = RGB(90, 158, 214)
        detalle.BackColorScrollBar = RGB(231, 235, 247)
        detalle.BackColor1 = RGB(231, 235, 247)
        detalle.BackColor2 = RGB(239, 243, 255)
        detalle.GridColor = RGB(148, 190, 231)
        'detalle.DefaultFont.Size = 8
'        detalle.Column(8).Width = 0
        
        
        detalle.Cell(0, 0).text = formatogrilla(1, 0)
        For i = 1 To detalle.Cols - 2
            detalle.Cell(0, i).text = formatogrilla(1, i)
            detalle.Column(i).Width = Val(formatogrilla(8, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            detalle.Column(i).MaxLength = Val(formatogrilla(2, i))
            detalle.Column(i).FormatString = formatogrilla(4, i)
            detalle.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
                If i <> 5 And i <> 3 Then
                    detalle.Column(i).Mask = cellNumeric
                End If
            Else
                detalle.Column(i).Alignment = cellLeftCenter
                detalle.Column(i).Mask = cellUpper
            End If
        Next i
        detalle.Column(6).Width = 0
        detalle.Range(0, 0, 0, detalle.Cols - 2).Alignment = cellCenterCenter
'        detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "", True
        detalle.Enabled = True
        detalle.ExtendLastCol = True
        
    
    End Sub
Sub LeerClienteFlete(rut, sucursal)


Dim op As Integer
Dim CAMPOS(6, 6) As String

    CAMPOS(0, 0) = "nombre"
    CAMPOS(1, 0) = "direccion"
    CAMPOS(2, 0) = "comuna"
    CAMPOS(3, 0) = "ciudad"
    CAMPOS(4, 0) = "fono1"
    CAMPOS(5, 0) = "giro"
    
    CAMPOS(6, 0) = ""
    CAMPOS(0, 2) = "sv_maestroclientes"
    condicion = "rut='" & rut & "' and sucursal='" & sucursal & "'"
    op = 5
    Set sqlventas.conexion = ventas
    sqlventas.response = CAMPOS
    Call sqlventas.sqlventas(op, condicion)

    If sqlventas.Status = 0 Then
    
    dato7.text = sqlventas.response(0, 3)
    dato8.text = sqlventas.response(1, 3)
    dato9.text = sqlventas.response(4, 3)
    dato10.text = sqlventas.response(2, 3)
    dato11.text = sqlventas.response(3, 3)
    dato13.text = sqlventas.response(5, 3)
   
   
    Else
    dato7.SetFocus
    End If
    

End Sub

    
Private Sub Form_Load()
Call CARGAGRILLA(1, 6)
End Sub

Private Sub guardar_Click()
glosaflete.Show vbModal

If grabar = True Then
Call modificarut(dato1.text, dato2.text, dato6.text & lbldv.Caption, ncaja.text)
MsgBox "Guardado exitosamente", vbOKOnly, "OK"
Call exiteflete(dato1.text, dato2.text)
opciones.Visible = True
opciones.SetFocus
guardar.Visible = False
End If

End Sub

  Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
'                Call modificar
            Case "elimina"
             If Verifica_Permiso(Me.Caption, "elimina") = True Then
                    If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                    frmglosaeliminacion.Show vbModal
                    Call ELIMINAR
                    retorno
                    End If
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                End If
                 
            Case "imprime"
                Call imprimir
            Case "movimientos"
            Case "historico"
            Case "retorno"
                Call retorno
            Case "anterior"
'                Call anterior
            Case "siguiente"
'                Call siguiente
        End Select
    End Sub
    Sub imprimir()
    Dim K As Integer
    
    final = detalle.Rows
    
    detalle.Column(1).Locked = False
    detalle.Column(2).Locked = False
    detalle.Column(3).Locked = False
    detalle.Column(4).Locked = False
    detalle.Column(5).Locked = False
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 8
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellLeftGeneral
    detalle.Cell(detalle.Rows - 1, 1).text = " "
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 9
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellCenterCenter
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 2).Borders(cellEdgeTop) = cellThin
    detalle.Cell(detalle.Rows - 1, 1).text = "INFORMACION DE REPARTO  "
    
     detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 13
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellCenterCenter
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Cell(detalle.Rows - 1, 1).text = glosa1flete
    
     detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 8
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellLeftGeneral
    detalle.Cell(detalle.Rows - 1, 1).text = " "
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 8
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = False
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellLeftCenter
    detalle.Cell(detalle.Rows - 1, 1).text = "Especificaciones de acceso al lugar "
    
    For K = 1 To 10
    If glosafletes(K) <> "" Then
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 8
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellLeftGeneral
    detalle.Cell(detalle.Rows - 1, 1).text = glosafletes(K)
    End If
    Next K
    
'    detalle.Rows = detalle.Rows + 1
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeTop) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
'
'
'    detalle.Rows = detalle.Rows + 1
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
'
'    detalle.Rows = detalle.Rows + 1
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
'
'    detalle.Rows = detalle.Rows + 1
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
'
'    detalle.Rows = detalle.Rows + 1
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
'
'    detalle.Rows = detalle.Rows + 1
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
'
'    detalle.Rows = detalle.Rows + 1
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
'
'    detalle.Rows = detalle.Rows + 1
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeBottom) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
'    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    
    Call Titulos("ORDEN DE FLETE")
    detalle.PageSetup.HeaderMargin = 0.5
    detalle.PageSetup.PrintFixedRow = True
    detalle.PageSetup.TopMargin = 1
    detalle.PageSetup.LeftMargin = 1
    detalle.PageSetup.RightMargin = 1
    detalle.PageSetup.BottomMargin = 3
    detalle.PageSetup.FooterMargin = 2
    detalle.PrintPreview
    detalle.Rows = final
    
    End Sub
    Sub retorno()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato7.text = ""
    dato11.text = ""
    dato13.text = ""
    dato14.text = ""
    DESDE1.text = ""
    DESDE2.text = ""
    DESDE3.text = ""
    
    lbldvvendedor.Caption = ""
    lblVendedor.Caption = ""
    
    detalle.Rows = 1
    lblDocumento.Caption = ""
    lbldv.Caption = ""
    
   
    Call CARGAGRILLA(1, 6)
    guardar.Visible = False
    opciones.Visible = False
    
    dato1.SetFocus
    End Sub
    
    Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    detalle.FixedRowColStyle = Fixed3D
    detalle.CellBorderColorFixed = vbButtonShadow
    detalle.ShowResizeTips = False
    detalle.ReportTitles.Clear
    detalle.PageSetup.CenterHorizontally = True
    detalle.PageSetup.Orientation = cellPortrait
    detalle.PageSetup.BlackAndWhite = True
    detalle.Column(1).Width = 90
    detalle.Column(2).Width = 250
    detalle.Column(3).Width = 50
    detalle.Column(4).Width = 80
    detalle.Column(5).Width = 85
    
        
    
    
      
    detalle.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    detalle.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    detalle.PageSetup.HeaderAlignment = cellLeft
    detalle.PageSetup.HeaderFont.Name = "Verdana"
    detalle.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "   NUMERO  :  " & numerofle
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellRight
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle
    
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = "EMITIDO  :  " & Format(fechasistema, "dd-mm-yyyy") & "    HORA  :" & horaorden
'    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
'    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold = True
'    objReportTitle.Align = cellRight
'    objReportTitle.PrintOnAllPages = True
'    detalle.ReportTitles.Add objReportTitle
    
      Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CAJA : " + ncaja.text + "  DOCUMENTO :  " & dato1.text & " - " & dato2.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle

        
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "RUT  :  " & dato6.text & "-" & lbldv.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CLIENTE  :  " & dato7.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle
    
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "DIRECCION  :  " & dato8.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FONO  :  " & dato9.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CIUDAD  :  " & dato10.text & "       CONTACTO  :  " & dato7.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CONTACTO  :  " & dato7.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle
    Set objReportTitle = New FlexCell.ReportTitle
    
    objReportTitle.text = "VENDEDOR  :  " & dato14.text & lbldvvendedor.Caption & "-" & lblVendedor.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    detalle.ReportTitles.Add objReportTitle

    
    'PIE DE PAGINA
    detalle.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & " Hora: " & Time & vbCrLf & "Usuario: " & usuarioSistema
    detalle.PageSetup.FooterAlignment = cellRight
    detalle.PageSetup.FooterFont.Name = "Verdana"
    detalle.PageSetup.FooterFont.Size = 7
     detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellEdgeTop) = cellThin
     detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellEdgeBottom) = cellThin
     detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
     detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
     detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellInsideHorizontal) = cellThin
     detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellInsideVertical) = cellThin
     detalle.Range(0, 1, 0, detalle.Cols - 1).FontBold = True
    
End Sub

 Function LEERULTIMOFLETE() As String
  
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventasRubro
    
    csql.sql = " select IFNULL(MAX(numero) + 1,'0000000001')"
    csql.sql = csql.sql & " from sv_fletes_cabeza_" & empresaActiva
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
              If resultados(0) <> "" And resultados(0) <> "0" Then
                 LEERULTIMOFLETE = resultados(0)
               Else
                 LEERULTIMOFLETE = "0000000001"
              End If
    End If
    Set resultados = Nothing
    csql.Close
    
End Function
Function grabar() As Boolean
Dim csql As New rdoQuery
Dim tabla As String
Dim K As Integer
Dim CAMPOS(50, 50) As String

Dim op As Integer
Set sql = New sqlventas.sqlventa

numerofle = LEERULTIMOFLETE
pivote.MaxLength = 10
pivote.text = numerofle
numerofle = ceros(pivote)
horaorden = Time

CAMPOS(0, 0) = "numero"
CAMPOS(1, 0) = "fecha"
CAMPOS(2, 0) = "cliente"
CAMPOS(3, 0) = "direccion"
CAMPOS(4, 0) = "comuna"
CAMPOS(5, 0) = "ciudad"
CAMPOS(6, 0) = "fono"
CAMPOS(7, 0) = "tipodocumento"
CAMPOS(8, 0) = "numerodocumento"
CAMPOS(9, 0) = "glosa1"
CAMPOS(10, 0) = "glosa2"
CAMPOS(11, 0) = ""

CAMPOS(0, 1) = numerofle
CAMPOS(1, 1) = dato5.text & "-" & dato4.text & "-" & dato3.text
CAMPOS(2, 1) = dato6.text & lbldv.Caption
CAMPOS(3, 1) = dato8.text
CAMPOS(4, 1) = dato10.text
CAMPOS(5, 1) = dato11.text
CAMPOS(6, 1) = dato9.text
CAMPOS(7, 1) = dato1.text
CAMPOS(8, 1) = dato2.text
CAMPOS(9, 1) = glosa1flete
CAMPOS(10, 1) = glosa2flete

CAMPOS(0, 2) = "sv_fletes_cabeza_" & empresaActiva
condicion = ""
op = 2
    sql.response = CAMPOS
    Set sql.conexion = ventasRubro
    sql.audit = True: sql.programaactivo = Me.Caption
    Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
    Call sql.sqlventas(op, condicion)
 
For K = 1 To detalle.Rows - 1
Call grabar_detalle(numerofle, K, detalle.Cell(K, 1).text, detalle.Cell(K, 2).text, detalle.Cell(K, 3).text, detalle.Cell(K, 6).text)
Next K
grabar = True
End Function
Sub grabar_detalle(NUMERO, linea, CODIGO, descripcion, cantidad, PRECIO)
Dim csql As New rdoQuery
Dim tabla As String
Dim K As Integer
Dim CAMPOS(50, 50) As String
Set sql = New sqlventas.sqlventa

Dim op As Integer

pivote.MaxLength = 10
pivote.text = NUMERO
pivote.text = ceros(pivote)


CAMPOS(0, 0) = "numero"
CAMPOS(1, 0) = "linea"
CAMPOS(2, 0) = "codigo"
CAMPOS(3, 0) = "descripcion"
CAMPOS(4, 0) = "cantidad"
CAMPOS(5, 0) = "precio"
CAMPOS(6, 0) = ""

CAMPOS(0, 1) = pivote.text
CAMPOS(1, 1) = linea
CAMPOS(2, 1) = CODIGO
CAMPOS(3, 1) = descripcion
CAMPOS(4, 1) = cantidad
CAMPOS(5, 1) = PRECIO
CAMPOS(0, 2) = "sv_fletes_detalle_" & empresaActiva
condicion = ""
op = 2
    sql.response = CAMPOS
    Set sql.conexion = ventasRubro
    sql.audit = True: sql.programaactivo = Me.Caption
    Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
    Call sql.sqlventas(op, condicion)
 
End Sub
Private Function exiteflete(TIPO, NUMERO) As Boolean
  Dim csql As New rdoQuery
  Dim resultados As rdoResultset
  Dim sumar As Double
  Dim fina As Double
  Dim ini As Double
  Dim K As Integer
  
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select numero,glosa1,glosa2 from sv_fletes_cabeza_" & empresaActiva
    csql.sql = csql.sql & " where tipodocumento='" & TIPO & "' and numerodocumento='" & NUMERO & "'"
    csql.Execute
    
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    exiteflete = True
    numerofle = resultados("numero")
    glosa1flete = resultados("glosa1")
    sumar = 0
    ini = 1
    For K = 1 To 10
    glosafletes(K) = ""
    Next K
    For K = 1 To Len(resultados("glosa2"))
    If Mid(resultados("glosa2"), K, 1) = Chr(13) Then
    fina = K - 1
    If fina = 0 Then fina = 1
    If ini > fina Then fina = ini
    sumar = sumar + 1
    If sumar > 20 Then sumar = 20
    glosafletes(sumar) = Mid(resultados("glosa2"), ini, fina - ini)
    ini = K + 2
    End If
    Next K
    
    Else
    exiteflete = False
    
    End If
    Set resultados = Nothing
    csql.Close
    End Function
    Sub grabarclienteflete()
    Dim csql As New rdoQuery
    Dim tabla As String
    tabla = ""
    tabla = "insert into sv_maestroclientes "
    tabla = tabla & "set rut='" & dato6.text & lbldv.Caption & "', sucursal='0', nombre='" & dato7.text & "', "
    tabla = tabla & "direccion='" & dato8.text & "', fono1='" & dato9.text & "', comuna='" & dato10.text & "', "
    tabla = tabla & "ciudad='" & dato11.text & "', giro='" & dato13.text & "'"
    tabla = tabla & "on duplicate key update rut=rut "
    Set csql.ActiveConnection = ventas
    csql.sql = tabla
    csql.Execute
    csql.Close
    Set csql = Nothing
    End Sub
    Sub modificarut(TIPO, NUMERO, rut, caja)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim GLOSA As String
        Dim cSql2 As rdoQuery
        
        Set cSql2 = New rdoQuery
        Set cSql2.ActiveConnection = ventasRubro
        cSql2.sql = "update sv_documento_cabeza_" & empresaActiva & " set rut='" & rut & "' "
        If Option2.Value = True Then
            cSql2.sql = cSql2.sql & "WHERE tipo='" + TIPO + "' and foliosii='" + NUMERO + "' and local='" + empresaActiva + "' and caja='" & caja & "' "
        End If
        
        If Option1.Value = True Then
            cSql2.sql = cSql2.sql & "WHERE tipo='" + TIPO + "' and numero='" + NUMERO + "' and local='" + empresaActiva + "' and caja='" & caja & "' "
        End If
        
        cSql2.Execute
        Call sincronizadatos(cSql2.sql, ventasRubro)
        
       End Sub
Sub ELIMINAR()
        
        Dim op As Integer
        Dim CAMPOS(4, 4) As String
        Set sql = New sqlventas.sqlventa
        condicion = "numero='" & numerofle & "'"
        op = 4
        CAMPOS(0, 2) = "sv_fletes_cabeza_" & empresaActiva & ""
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        
        Set sql = New sqlventas.sqlventa
        condicion = "numero='" & numerofle & "'"
        op = 4
        CAMPOS(0, 2) = "sv_fletes_detalle_" & empresaActiva & ""
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)


 

End Sub
