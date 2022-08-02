VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form maestro02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Cuentas Corrientes"
   ClientHeight    =   10140
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   13350
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   676
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmdte 
      Height          =   1095
      Left            =   8400
      TabIndex        =   40
      Top             =   4080
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1931
      BackColor       =   16744576
      Caption         =   "Cuenta Centralizacion DTE"
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
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar"
         Height          =   255
         Left            =   3600
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox cuentadte 
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
         Left            =   120
         MaxLength       =   8
         TabIndex        =   41
         Tag             =   "contacto"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label glosadte 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   4410
      End
   End
   Begin VB.CommandButton CmdmgFactura 
      Caption         =   "VER &FACTURAS"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9360
      Width           =   1575
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   3525
      Left            =   180
      TabIndex        =   35
      Top             =   5265
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   6218
      BackColor       =   16761024
      Caption         =   "Cuentas Empresas Relacionadas"
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
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   285
         Left            =   4050
         TabIndex        =   38
         Top             =   3195
         Width           =   1320
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Todas Las Empresas"
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
         Height          =   330
         Left            =   10350
         TabIndex        =   37
         Top             =   3105
         Width           =   2580
      End
      Begin FlexCell.Grid Grid1 
         Height          =   2760
         Left            =   90
         TabIndex        =   36
         Top             =   315
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   4868
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5130
      Left            =   180
      TabIndex        =   18
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9049
      BackColor       =   16744576
      Caption         =   "DATOS CUENTAS CORRIENTES"
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
      Begin VB.CheckBox chk2 
         BackColor       =   &H00FF8080&
         Caption         =   "Aceptacion de Documentos sin Supervision"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   50
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton dp 
         BackColor       =   &H00FFFF80&
         Caption         =   "DATOS DEL PAGO"
         Height          =   330
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4635
         Width           =   2220
      End
      Begin XPFrame.FrameXp GLOSACTACTE 
         Height          =   255
         Left            =   2925
         TabIndex        =   33
         Top             =   315
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         BackColor       =   49344
         Caption         =   ""
         CaptionEstilo3D =   1
         BackColor       =   49344
         ForeColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Left            =   1725
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "tipo"
         Top             =   315
         Width           =   1095
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "direccion"
         Top             =   1395
         Width           =   6015
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
         Left            =   1725
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "rut"
         Top             =   675
         Width           =   1095
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "nombre"
         Top             =   1035
         Width           =   6015
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
         Left            =   1725
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Comuna"
         Top             =   1755
         Width           =   3255
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
         Left            =   1725
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "ciudad"
         Top             =   2115
         Width           =   3255
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "giro"
         Top             =   2475
         Width           =   6015
      End
      Begin VB.TextBox dato14 
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
         Left            =   1725
         TabIndex        =   12
         Tag             =   "dest_cheque"
         Top             =   4635
         Width           =   255
      End
      Begin VB.TextBox dato11 
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
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "celular"
         Top             =   3555
         Width           =   1815
      End
      Begin VB.TextBox dato10 
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
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "fax"
         Top             =   3195
         Width           =   1815
      End
      Begin VB.TextBox dato9 
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
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "fono"
         Top             =   2835
         Width           =   1815
      End
      Begin VB.TextBox dato13 
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "contacto"
         Top             =   4275
         Width           =   5895
      End
      Begin VB.TextBox dato12 
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "email"
         Top             =   3915
         Width           =   5895
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   972
         Left            =   3960
         TabIndex        =   47
         Top             =   2880
         Width           =   3732
         _ExtentX        =   6588
         _ExtentY        =   1720
         BackColor       =   16761024
         Caption         =   "APORTE PUBLICITARIO        %"
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "No aporta"
            Height          =   492
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1812
         End
         Begin VB.TextBox aporte 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Left            =   2040
            TabIndex        =   48
            Top             =   360
            Width           =   1572
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut"
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
         Left            =   90
         TabIndex        =   32
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Cuenta"
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
         Left            =   90
         TabIndex        =   31
         Top             =   315
         Width           =   1530
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   90
         TabIndex        =   30
         Top             =   1035
         Width           =   1530
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         Left            =   90
         TabIndex        =   29
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comuna"
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
         Left            =   90
         TabIndex        =   28
         Top             =   1755
         Width           =   1530
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ciudad"
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
         Left            =   90
         TabIndex        =   27
         Top             =   2115
         Width           =   1530
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Giro"
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
         Left            =   90
         TabIndex        =   26
         Top             =   2475
         Width           =   1530
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fono"
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
         Left            =   90
         TabIndex        =   25
         Top             =   2835
         Width           =   1530
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fax"
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
         Index           =   0
         Left            =   90
         TabIndex        =   24
         Top             =   3195
         Width           =   1530
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Email"
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
         Left            =   90
         TabIndex        =   23
         Top             =   3915
         Width           =   1530
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Celular"
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
         Left            =   90
         TabIndex        =   22
         Top             =   3555
         Width           =   1530
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contacto"
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
         Left            =   90
         TabIndex        =   21
         Top             =   4275
         Width           =   1530
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Destino Cheque"
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
         Left            =   90
         TabIndex        =   20
         Top             =   4635
         Width           =   1530
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
         Left            =   2925
         TabIndex        =   19
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
      ScaleWidth      =   13320
      TabIndex        =   17
      Top             =   10140
      Width           =   13350
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8400
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3735
      Left            =   8400
      TabIndex        =   14
      Top             =   240
      Width           =   4695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SALDOS 
         Height          =   3495
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6165
         _Version        =   393216
         BackColor       =   16776436
         ForeColor       =   12582912
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   16107953
         BackColorSel    =   16777215
         ForeColorSel    =   16744576
         BackColorBkg    =   16776436
         GridColor       =   -2147483635
         GridColorFixed  =   12582912
         GridLinesFixed  =   1
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   3735
         Left            =   0
         Top             =   0
         Width           =   4695
      End
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   615
      Left            =   9960
      TabIndex        =   44
      Top             =   9360
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   280
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   180
      TabIndex        =   13
      Top             =   8820
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
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   3735
      Left            =   8520
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "maestro02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private MODIFI As Integer

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub aporte_GotFocus()
    Call cargatexto(aporte)
End Sub

Private Sub aporte_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumeroDecimal(aporte, KeyAscii)
    If KeyAscii = 13 Then
        If aporte.text <> "" Then
            Call grabarrapel(dato2.text & DV.Caption, dato4.text, aporte.text)
        End If
    End If
End Sub
Sub grabarrapel(rutpro, nombrepro, valor)
    
    campos(0, 0) = "dato1"
    campos(1, 0) = "dato2"
    campos(2, 0) = "dato3"
    campos(3, 0) = ""
    
    campos(0, 1) = rutpro
    campos(1, 1) = nombrepro
    campos(2, 1) = Replace(valor, ".", ",")
    
    
    campos(0, 2) = "rapel"
    condicion = "dato1=" + "'" + rutpro + "' "
     op = 5
     
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
    Else
        condicion = "dato1=" + "'" + rutpro + "' "
        op = 3
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
    End If
    
    
    
    End Sub
 
Private Sub CmdmgFactura_Click()
Load IngresoImgFactura
With IngresoImgFactura
    .Label1(0) = "TIPO : FACTURAS"
    .Label1(2).Caption = "RUT PROVEEDOR :" & Me.dato2
    .Option1(1).Value = True
    .Label1(3) = "PROVEEDOR :" & dato4
    .CmdEliminaImagen.Enabled = False
    .CmdGuardaImagen.Enabled = False
    .CmdNuevaImagen.Enabled = False
End With
IngresoImgFactura.Show vbModal

End Sub

Private Sub Check1_Click()
If Check1.Value = "1" Then

todaslasrelaciones
Else
Grid1.Rows = 1
Call leerelaciones(Me, Grid1, dato2.text + DV.Caption, empresaactiva)
End If

End Sub
Sub todaslasrelaciones()
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigoempresa "
        csql.sql = csql.sql + "FROM maestroempresas "
        csql.sql = csql.sql + "order by codigoempresa "
        csql.Execute
        Grid1.Rows = 1
    If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
        While resultados.EOF = False
        Call leerelaciones(Me, Grid1, dato2.text + DV.Caption, resultados(0))
        
        resultados.MoveNext
        
        Wend
        
            resultados.Close
        Set resultados = Nothing
            
    End If
    
    csql.Close
    Set csql = Nothing

End Sub

Private Sub Command1_Click()
imprimir
End Sub
Private Sub imprimir()
If Grid1.Rows > 1 Then
Call Titulos("LISTADO DE SALDOS RELACIONADOS ")
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.HeaderMargin = 0.5
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.TopMargin = 3
Grid1.PageSetup.LeftMargin = 0.1
Grid1.PageSetup.RightMargin = 0.1
Grid1.PageSetup.BottomMargin = 3
Grid1.PageSetup.FooterMargin = 2
Grid1.PageSetup.BlackAndWhite = True

Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid1.PrintPreview
End If
End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.Orientation = cellLandscape
    
      
    Grid1.PageSetup.PrintTitleRows = 0
    
    'Logo
'    grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    Grid1.PageSetup.HeaderAlignment = CellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "EMITIDO  :  " & Format(fechasistema, "dd-MM-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = dato2.text + "-" + DV.Caption & "  " & dato4.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
      
      
      
    
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = tipoListado
'    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
'    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold = True
'    objReportTitle.Align = cellCenter
'    objReportTitle.PrintOnAllPages = True
'    grid1.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub

Private Sub COMMAND2_Click()
Call cuentadte_KeyPress(13)
If glosadte.Caption <> "" Or cuentadte.text = "" Then
Call grabar_dte(dato2.text + DV.Caption, cuentadte.text)
End If


End Sub

Private Sub cuentadte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
    Call ayudatipocuenta2(cuentadte)
    Call cuentadte_KeyPress(13)
    
    End If
    

End Sub

Private Sub cuentadte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    glosadte = leerNombreMayor(cuentadte.text)
    If glosadte = "" And cuentadte.text <> "" Then
    MsgBox ("CODIGO DE CUENTA NO EXISTE ")
    cuentadte.SetFocus
    
    Else
    Command2.SetFocus
    
    End If
    
    End If
    
    

End Sub

Private Sub dp_Click()
If dato4.text <> "" Then
DATOSPAGO.Show vbModal
End If


End Sub

Private Sub dato1_GotFocus()

grillasaldos
If scrut = "S" Then dato4.SetFocus

Call cargatexto(dato1)
End Sub

Private Sub dato2_GotFocus()
LEETIPOCTACTE
dp.Visible = False

If dato1.text = CUENTAPROVEEDOR Or dato1.text = "23100029" Then
dp.Visible = True

End If

Call cargatexto(dato2)
End Sub
Private Sub dato4_GotFocus()
DV.Caption = rut(dato2.text)


If MODIFI = 0 And scrut <> "S" Then leer

Call cargatexto(dato4)
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
Private Sub dato8_GotFocus()
Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()
Call cargatexto(dato9)
End Sub
Private Sub dato10_GotFocus()
Call cargatexto(dato10)
End Sub
Private Sub dato11_GotFocus()
Call cargatexto(dato11)
End Sub
Private Sub dato12_GotFocus()
Call cargatexto(dato12)
End Sub
Private Sub dato13_GotFocus()
Call cargatexto(dato13)
End Sub
Private Sub dato14_GotFocus()
Call cargatexto(dato14)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 And cierrect = "S" Then cierrect = "-"
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudatipocuenta(dato2)
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudactacte(dato4)
    Call flechas(dato1, dato4, KeyCode)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, DATO5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(DATO5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato7, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato8, dato10, KeyCode)
End Sub
Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato9, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato10, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato11, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato13, dato14, KeyCode)
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

    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
    Me.CmdmgFactura.Enabled = False
DOCU(1) = "ACTIVO"
DOCU(2) = "PASIVO"
DOCU(3) = "RESULTADO"
CANDO = 3

Rem Call RECUPERAFECHA
Call CARGAPERMISO(Me.Name)
Call CARGAGRILLArelacion
dp.Visible = False
frmdte.Visible = False

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato1.text) <> 0 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato2.text) <> 0 Then Call ceros(dato2): Call Pregunta(dato2, dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato4.text) <> "" Then Call Pregunta(dato4, DATO5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(DATO5.text) <> "" Then sc = 1: Call Pregunta(DATO5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato6.text) <> "" Then sc = 1: Call Pregunta(dato6, dato7)
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato7.text) <> "" Then sc = 1: Call Pregunta(dato7, dato8)
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato8.text) <> "" Then sc = 1: Call Pregunta(dato8, dato9)
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato9, dato10)
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato10, dato11)
End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato11, dato12)
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato12, dato13)
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato13, dato14)
End Sub
Private Sub dato14_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        grabar
        retorno
    End If
End Sub



Sub leer()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato4.Tag
    campos(3, 0) = DATO5.Tag
    campos(4, 0) = dato6.Tag
    campos(5, 0) = dato7.Tag
    campos(6, 0) = dato8.Tag
    campos(7, 0) = dato9.Tag
    campos(8, 0) = dato10.Tag
    campos(9, 0) = dato11.Tag
    campos(10, 0) = dato12.Tag
    campos(11, 0) = dato13.Tag
    campos(12, 0) = dato14.Tag
    campos(13, 0) = "servicio"
    campos(14, 0) = ""
    
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut=" + "'" + dato2.text + DV.Caption + "' and año='" + Format(fechasistema, "yyyy") + "'"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato4.SetFocus: GoTo no:
    carga
    opciones.Visible = True
    Me.CmdmgFactura.Enabled = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
        
no:
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato4.Tag
    campos(3, 0) = DATO5.Tag
    campos(4, 0) = dato6.Tag
    campos(5, 0) = dato7.Tag
    campos(6, 0) = dato8.Tag
    campos(7, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut>'" + dato2.text + DV.Caption + "' and año='" + Format(fechasistema, "yyyy") + "' order by tipo,rut asc "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    Me.CmdmgFactura.Enabled = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
    
    
no:
   
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato4.Tag
    campos(3, 0) = DATO5.Tag
    campos(4, 0) = dato6.Tag
    campos(5, 0) = dato7.Tag
    campos(6, 0) = dato8.Tag
    campos(7, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut<'" + dato2.text + DV.Caption + "' and año='" + Format(fechasistema, "yyyy") + "' order by tipo,rut desc "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    Me.CmdmgFactura.Enabled = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
    
    
no:
   
    
End Sub

Sub carga()
    Dim cuentadte2 As String
    
    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    dato2.text = Mid(sqlconta.response(1, 3), 1, 9)
    DV.Caption = Mid(sqlconta.response(1, 3), 10, 1)
    dato4.text = sqlconta.response(2, 3)
    DATO5.text = sqlconta.response(3, 3)
    dato6.text = sqlconta.response(4, 3)
    dato7.text = sqlconta.response(5, 3)
    dato8.text = sqlconta.response(6, 3)
    dato9.text = sqlconta.response(7, 3)
    dato10.text = sqlconta.response(8, 3)
    dato11.text = sqlconta.response(9, 3)
    dato12.text = sqlconta.response(10, 3)
    dato13.text = sqlconta.response(11, 3)
    dato14.text = sqlconta.response(12, 3)
    chk2.Value = sqlconta.response(13, 3)
    
    CARGAGRILLArelacion
    Grid1.Rows = 1
    Check1.Value = "0"
    Call Check1_Click
    If dato1.text = CUENTAPROVEEDOR Then
    frmdte.Visible = True
    cuentadte2 = electronico(dato2.text + DV.Caption)
    aporte.text = leeraporte(dato2.text & DV.Caption)
    Else
    frmdte.Visible = False
    
    End If
    
    If cuentadte2 <> "" Then
    frmdte.Visible = True
    
    cuentadte.text = cuentadte2
    glosadte = leerNombreMayor(cuentadte2)
    End If
    
fin:
End Sub
Function leeraporte(rutprove) As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = conta
    csql.sql = "select dato3 from rapel where "
    csql.sql = csql.sql & "dato1 ='" & rutprove & "' "
    csql.Execute
    leeraporte = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leeraporte = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
    
End Function
Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    
    dato4.Locked = condicion
    DATO5.Locked = condicion
    dato6.Locked = condicion
    dato7.Locked = condicion
    dato8.Locked = condicion
    dato9.Locked = condicion
    dato10.Locked = condicion
    dato11.Locked = condicion
    dato12.Locked = condicion
    dato13.Locked = condicion
    dato14.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion

    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    dato9.Enabled = condicion
    dato10.Enabled = condicion
    dato11.Enabled = condicion
    dato12.Enabled = condicion
    dato13.Enabled = condicion
    dato14.Enabled = condicion
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudatipocuenta(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "CTACTE <> '0' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("cuenta", "nombre")
    mensajeAyuda = "Ayuda tipo de Cuentas Corrientes"
        
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudatipocuenta2(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("cuenta", "nombre")
    mensajeAyuda = "Ayuda plan de cuentas"
        
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", cuentadte, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato4.Tag
    campos(3, 0) = DATO5.Tag
    campos(4, 0) = dato6.Tag
    campos(5, 0) = dato7.Tag
    campos(6, 0) = dato8.Tag
    campos(7, 0) = dato9.Tag
    campos(8, 0) = dato10.Tag
    campos(9, 0) = dato11.Tag
    campos(10, 0) = dato12.Tag
    campos(11, 0) = dato13.Tag
    campos(12, 0) = dato14.Tag
    campos(13, 0) = "año"
    campos(14, 0) = "servicio"
    campos(15, 0) = ""
    
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text + DV.Caption
    campos(2, 1) = dato4.text
    campos(3, 1) = DATO5.text
    campos(4, 1) = dato6.text
    campos(5, 1) = dato7.text
    campos(6, 1) = dato8.text
    campos(7, 1) = dato9.text
    campos(8, 1) = dato10.text
    campos(9, 1) = dato11.text
    campos(10, 1) = dato12.text
    campos(11, 1) = dato13.text
    campos(12, 1) = dato14.text
    campos(13, 1) = Format(fechasistema, "yyyy")
    campos(14, 1) = chk2.Value
    
    campos(0, 2) = "cuentascorrientes"
    If MODIFI = 1 Then condicion = "tipo=" + "'" + dato1.text + "' and rut ='" + dato2.text + DV.Caption + "' and año='" + Format(fechasistema, "yyyy") + "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If MODIFI = 0 Then grabar2
    
    End Sub
Sub grabar2()
      
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
    campos(26, 0) = "haber10"
    campos(27, 0) = "haber11"
    campos(28, 0) = "haber12"
    
    campos(29, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text + DV.Caption
    campos(2, 1) = año

    For k = 3 To 28
    campos(k, 1) = "0"
    Next k
    campos(0, 2) = "saldosctacte"
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub

Sub ELIMINAR()
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut=" + "'" + dato2.text + DV.Caption + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub



Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then modifica
If command = "elimina" Then ELIMINA

If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "movimientos" Then movimientos



End Sub
Sub ELIMINA()
If saldoglobal = 0 Then
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
Me.CmdmgFactura.Enabled = False
dato1.SetFocus
Else
MsgBox ("IMPOSIBLE ELIMINAR RUT CON DATOS")
End If
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
frmdte.Visible = False

disponible (True)
habilita (False)
limpia
opciones.Visible = False
Me.CmdmgFactura.Enabled = False
dato1.Enabled = True
dato1.SetFocus
MODIFI = 0
no:
Grid1.Rows = 1
If cierrect = "S" Then cierrect = "": Unload Me
    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    DV.Caption = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
    chk2.Value = "0"
End Sub



Sub DATOSSALDOS()
Dim debe As Double
Dim haber As Double

LEERSALDOS
saldoglobal = LEERSALDOSCTACTEmovi(dato1.text, dato2.text + DV.Caption, empresaactiva)
sumador = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
SALDOS.TextMatrix(1, 1) = Format(sqlconta.response(3, 3), "###,###,##0")
SALDOS.TextMatrix(1, 2) = Format(sqlconta.response(4, 3), "###,###,##0")
SALDOS.TextMatrix(1, 3) = Format(sumador, "###,###,##0")
debe = 0
haber = 0

For k = 1 To 12

SALDOS.TextMatrix(k + 1, 1) = Format(ctadebe2(k), "###,###,##0")
SALDOS.TextMatrix(k + 1, 2) = Format(ctahaber2(k), "###,###,##0")
sumador = sumador + ctadebe2(k) - ctahaber2(k)
SALDOS.TextMatrix(k + 1, 3) = Format(sumador, "###,###,##0")
debe = debe + ctadebe2(k)
haber = haber + ctahaber2(k)

Next k
saldoglobal = debe + haber

End Sub
Sub grillasaldos()
SALDOS.Cols = 4
SALDOS.Rows = 14
SALDOS.ColWidth(0) = 120 * 12
SALDOS.ColWidth(1) = 120 * 8
SALDOS.ColWidth(2) = 120 * 8
SALDOS.ColWidth(3) = 120 * 8
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

Sub LEERSALDOS()
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = ""
    condicion = "tipo=" + "'" + dato1.text + "' and rut='" + dato2.text + DV.Caption + "' and año='" + Mid(fechasistema, 7, 4) + "'"
    campos(0, 2) = "saldosctacte"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
    
   Rem  If sqlconta.status = 4 Then Stop
grillasaldos
End Sub




Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus

End Sub

Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & dato1.text & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then dato2.SetFocus: GoTo no
    dato4.Enabled = True
    dato2.text = Mid(pivote.text, 1, 9)
    DV.Caption = Mid(pivote.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub

Sub LEETIPOCTACTE()

  If PermisosCuentasDelMayor(USUARIOSISTEMA, Format(dato1.text, "00000000")) = False Then
    MsgBox "USTED NO TIENE PRIVILEGIOS PARA ACCEDER A ESTA CUENTA", vbCritical, "ATENCION"
  dato1.SetFocus
  Exit Sub
End If


    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)

   If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
   GLOSACTACTE.Caption = sqlconta.response(1, 3)

no:
End Sub

Sub movimientos()
Rem cartola = "mayor:" + dato1.text + dato2.text + dato3.text
informa04.ctdato1.text = dato1.text
informa04.ctdato2.text = dato2.text
informa04.DV.text = DV.Caption
informa04.nombrectacte = GLOSACTACTE.Caption



informa04.ctnombre = dato4.text
informa04.sbtab1.Tab = 1
informa04.ctindi = True


informa04.Show

End Sub

Sub CARGAGRILLArelacion()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CUENTA"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "SALDO ANTE."
    formatogrilla2(1, 4) = "DEBE"
    formatogrilla2(1, 5) = "HABER"
    formatogrilla2(1, 6) = "SALDO ACTUAL"
    formatogrilla2(1, 7) = "EMPRESA"
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "6"
    formatogrilla2(2, 2) = "20"
    formatogrilla2(2, 3) = "10"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "17"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "C"
    formatogrilla2(3, 2) = "C"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "S"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 3) = " ###,###,##0"
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " ###,###,##0"
    formatogrilla2(4, 6) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 8
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
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
 
    End Sub

Sub grabar_dte(rut, cuenta)
    campos(0, 0) = "rut"
    campos(1, 0) = "contable"
    campos(2, 0) = ""
    campos(0, 1) = rut
    campos(1, 1) = cuenta
    
    campos(0, 2) = "proveedores_cuenta"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    condicion = "rut='" + rut + "' "
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub
