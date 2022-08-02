VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form guiaMolienda 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guia Molienda"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp cabeza 
      Height          =   3765
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   6641
      BackColor       =   16744576
      Caption         =   "Datos Molienda"
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
      Begin VB.TextBox dato5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         CausesValidation=   0   'False
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
         Left            =   5520
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   840
         Width           =   660
      End
      Begin VB.TextBox dato4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         CausesValidation=   0   'False
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
         Left            =   5100
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   840
         Width           =   360
      End
      Begin VB.TextBox dato3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         CausesValidation=   0   'False
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
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   840
         Width           =   360
      End
      Begin XPFrame.FrameXp frmTipo 
         Height          =   1275
         Left            =   1140
         TabIndex        =   42
         Top             =   420
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2249
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
            Caption         =   " 1 - Guia Molienda"
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
            TabIndex        =   44
            Top             =   480
            Width           =   2475
         End
         Begin VB.Label lbl24 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 2 - Factura Molienda"
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
            TabIndex        =   43
            Top             =   840
            Width           =   2475
         End
      End
      Begin VB.TextBox dato0 
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
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   360
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
         Left            =   4680
         MaxLength       =   9
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   2400
         Width           =   1050
      End
      Begin VB.TextBox dato7 
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
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   2820
         Width           =   1050
      End
      Begin VB.TextBox dato1 
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
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   420
         Width           =   1320
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
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   840
         Width           =   1320
      End
      Begin VB.TextBox dato8 
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
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   3300
         Width           =   1320
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
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   3300
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblTipo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   1140
         TabIndex        =   41
         Top             =   420
         Width           =   2010
      End
      Begin VB.Label lbl0 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   40
         Top             =   420
         Width           =   525
      End
      Begin VB.Label lblSucursal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   4680
         TabIndex        =   37
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label lblRut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   1500
         TabIndex        =   36
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label lbl15 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL TRIGO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3300
         TabIndex        =   34
         Top             =   2820
         Width           =   1305
      End
      Begin VB.Label lblTrigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   4680
         TabIndex        =   33
         Top             =   2820
         Width           =   1050
      End
      Begin VB.Label lbl16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KG."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   32
         Top             =   2820
         Width           =   435
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO GUIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3300
         TabIndex        =   31
         Top             =   420
         Width           =   1305
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   105
         TabIndex        =   30
         Top             =   1560
         Width           =   6330
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FOLIO INTERNO"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   29
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblDV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   2880
         TabIndex        =   28
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3300
         TabIndex        =   27
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " RUT CLIENTE"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   26
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SUCURSAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3300
         TabIndex        =   25
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label lbl17 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DOCUMENTO"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   24
         Top             =   3300
         Width           =   1305
      End
      Begin VB.Label lbl18 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FACTURA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3300
         TabIndex        =   23
         Top             =   3300
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HARINA FLOR"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   22
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " AFRECHO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3300
         TabIndex        =   21
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label lbl10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HARINILLA"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   20
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label lbl12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IMPUREZAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3300
         TabIndex        =   19
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label lbl14 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VALOR $"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   18
         Top             =   2820
         Width           =   1305
      End
      Begin VB.Label lblHarina 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   1500
         TabIndex        =   17
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label lbl7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KG."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2580
         TabIndex        =   16
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lblAfrecho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   4680
         TabIndex        =   15
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label lbl9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KG."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   14
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lblHarinilla 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Left            =   1500
         TabIndex        =   13
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Label lbl11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KG."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2580
         TabIndex        =   12
         Top             =   2400
         Width           =   435
      End
      Begin VB.Label lbl13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KG."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   11
         Top             =   2400
         Width           =   435
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   38
      Top             =   0
      Width           =   855
   End
   Begin FlexCell.Grid impresion 
      Height          =   735
      Left            =   60
      TabIndex        =   39
      Top             =   3180
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   1380
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1335
      Left            =   360
      TabIndex        =   35
      Top             =   3960
      Width           =   6375
      _cx             =   11245
      _cy             =   2355
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
End
Attribute VB_Name = "guiaMolienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private m As Moliendas
    Private gm As GuiaMoliendas
    Private modifica As Boolean
    Private estado As Integer

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato0_GotFocus()
        frmTipo.Visible = True
        Call selecciona(dato0)
    End Sub

    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        dato1.text = leer_Ultimo_Folio("numero", "sv_guiasmolienda", dato1.MaxLength, ventasRubro, "local = '" & empresaActiva & "' AND tipo = '" & dato0.text & "'")
        dato1.SelStart = Len(dato1.text) - 1
        dato1.SelLength = 1
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
        dato7.text = Format(dato7.text, "########0")
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
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato0_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 97
                dato0.text = "GM"
            Case 98
                dato0.text = "FM"
            Case Else
                Call Flechas(KeyCode, dato0)
        End Select
    End Sub
    
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato0)
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
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato0_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            lblTipo.Caption = leerNombreDocumento(dato0.text)
            If lblTipo.Caption <> "" Then
                Select Case dato0.text
                    Case "GM"
                        'dato8.Enabled = True
                        'dato9.Enabled = False
                        gm.tipo = "GM"
                    Case "FM"
                        'dato8.Enabled = False
                        'dato9.Enabled = True
                        gm.tipo = "FM"
                End Select
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato1.text <> "" Then
            dato1.text = ceros(dato1)
            If Val(dato1.text) = 0 Then
                dato1.text = leer_Ultimo_Folio("numero", "sv_guiasmolienda", dato1.MaxLength, ventasRubro, "local = '" & empresaActiva & "' AND tipo = 'GM'")
                SendKeys "{Tab}"
            Else
                If leerGuiaMolienda(gm, dato1.text, "=") = True Then
                    Call structtoctrl
                    opciones.Visible = True
                    opciones.SetFocus
                Else
                    Call HabilitarCajas(Me, modifica)
                    'If dato1.text = "GM" Then
                    '    dato9.Enabled = False
                    'Else
                    '    dato8.Enabled = False
                    'End If
                    SendKeys "{Tab}"
                End If
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If leerMolienda(m, dato2.text, "=") = True Then
                Call structtoctrlMolienda
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "00" Then
                dato3.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
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
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato8.text = ceros(dato8)
            If Val(dato8.text) <> 0 Then
                'dato9.text = "0000000000"
                'If estado = 0 Then
                ctrltostruct
                estado = 1
                'End If
            Else
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato9.text = ceros(dato9)
            If Val(dato9.text) <> 0 Then
                dato8.text = "0000000000"
                'If estado = 0 Then
                    ctrltostruct
                    estado = 1
                'End If
            Else
                dato8.SetFocus
            End If
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
        Call seleccionaUno(KeyCode, dato2)
    End Sub
    
    Private Sub dato6_KeyUp(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 37
                If dato6.SelStart > 0 And dato6.SelStart < dato6.MaxLength Then
                    dato6.SelStart = dato6.SelStart - 1
                End If
                dato6.SelLength = 1
                If dato6.SelText = "," Then
                    dato6.SelStart = dato6.SelStart - 1
                    dato6.SelLength = 1
                End If
            Case 39
                If dato6.SelStart > 0 And dato6.SelStart <= dato6.MaxLength Then
                    dato6.SelStart = dato6.SelStart - 1
                End If
                dato6.SelLength = 1
                If dato6.SelText = "," Then
                    dato6.SelStart = dato6.SelStart + 1
                    dato6.SelLength = 1
                End If
            Case 96, 97, 98, 99, 100, 101, 102, 103, 104, 105
                dato6.SelLength = 1
                If dato6.SelText = "," Then
                    dato6.SelStart = dato6.SelStart + 1
                    dato6.SelLength = 1
                End If
        End Select
    End Sub
    
    Private Sub dato7_KeyUp(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 37
                If dato7.SelStart > 0 And dato7.SelStart < dato7.MaxLength Then
                    dato7.SelStart = dato7.SelStart - 1
                End If
                dato7.SelLength = 1
            Case 39
                If dato7.SelStart > 0 And dato7.SelStart <= dato7.MaxLength Then
                    dato7.SelStart = dato7.SelStart - 1
                End If
                dato7.SelLength = 1
            Case 96, 97, 98, 99, 100, 101, 102, 103, 104, 105
                dato7.SelLength = 1
        End Select
    End Sub
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato0_LostFocus()
        frmTipo.Visible = False
    End Sub
    
    Private Sub dato7_LostFocus()
        dato7.text = Format(dato7.text, "$ ###,###,##0")
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        gm.codLoc = empresaActiva
        gm.tipo = dato0.text
        gm.numero = dato1.text
        gm.folio = dato2.text
        gm.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        gm.rut = lblRut.Caption & lblDV.Caption
        gm.sucursal = lblSucursal.Caption
        
        gm.harina = Replace(Format(lblHarina.Caption, "########0.00"), ",", ".")
        gm.afrecho = Replace(Format(lblAfrecho.Caption, "########0.00"), ",", ".")
        gm.harinilla = Replace(Format(lblHarinilla.Caption, "########0.00"), ",", ".")
        gm.impurezas = Replace(Format(dato6.text, "########0.00"), ",", ".")
        gm.valor = Replace(Format(dato7.text, "########0.00"), ",", ".")
        gm.trigo = Replace(Format(lblTrigo.Caption, "########0.00"), ",", ".")
        
        If dato0.text = "GM" Then
            gm.tipodocumento = "BV"
        Else
            gm.tipodocumento = "FV"
        End If
        gm.numeroDocumento = dato8.text
        
        Call grabarGuiaMolienda(gm, modifica)
        If MsgBox("DESEA IMPRIMIR COMPROBANTE ", vbYesNo) = vbYes Then
            Call imprimeGuia(dato0.text, dato1.text, Impresion, data)
        End If
        Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrlMolienda()
        dato2.text = m.folio
        dato3.text = Format(m.fecha, "dd")
        dato4.text = Format(m.fecha, "mm")
        dato5.text = Format(m.fecha, "yyyy")
        lblRut.Caption = Left(m.rut, Len(m.rut) - 1)
        lblDV.Caption = Right(m.rut, 1)
        lblSucursal.Caption = m.sucursal
        lblNombre.Caption = leerNombreClienteSucursal(m.rut, m.sucursal)
        lblHarina.Caption = Format(m.harina, "###,###,##0.00")
        lblAfrecho.Caption = Format(m.afrecho, "###,###,##0.00")
        lblHarinilla.Caption = Format(m.harinilla, "###,###,##0.00")
        dato6.text = Format(m.impurezas, "###,###,##0.00")
        dato7.text = Format(m.valor, " $ ###,###,##0")
        lblTrigo.Caption = Format(m.trigo, "###,###,##0.00")
    End Sub
    
    Private Sub structtoctrl()
        dato1.text = gm.numero
        dato2.text = gm.folio
        dato3.text = Format(gm.fecha, "dd")
        dato4.text = Format(gm.fecha, "mm")
        dato5.text = Format(gm.fecha, "yyyy")
        lblRut.Caption = Left(gm.rut, Len(gm.rut) - 1)
        lblDV.Caption = Right(gm.rut, 1)
        lblSucursal.Caption = gm.sucursal
        lblNombre.Caption = leerNombreClienteSucursal(gm.rut, gm.sucursal)
        lblHarina.Caption = Format(gm.harina, "###,###,##0.00")
        lblAfrecho.Caption = Format(gm.afrecho, "###,###,##0.00")
        lblHarinilla.Caption = Format(gm.harinilla, "###,###,##0.00")
        dato6.text = Format(gm.impurezas, "###,###,##0.00")
        dato7.text = Format(gm.valor, " $ ###,###,##0")
        lblTrigo.Caption = Format(gm.trigo, "###,###,##0.00")
        If gm.tipodocumento = "BV" Then
            dato8.text = gm.numeroDocumento
            dato9.text = "0000000000"
        End If
        If gm.tipodocumento = "FV" Then
            dato8.text = "0000000000"
            dato9.text = gm.numeroDocumento
        End If
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
            Unload Me
            PVentas.Show
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato0" Then
                Unload Me
            End If
        End If
End Sub

Private Sub Form_Load()
    modifica = False
    estado = 0
End Sub

Private Sub manual_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27, Asc("r"), Asc("R")
            Call retorno
        Case Asc("a"), Asc("A"), 37
            Call anterior
        Case Asc("s"), Asc("S"), 39
            Call siguiente
        Case Asc("m"), Asc("M")
            Call modificar
        Case Asc("e"), Asc("E"), 46
            Call eliminar
    End Select
End Sub

'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
                Call modificar
            Case "elimina"
                If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                Call eliminar
                End If
            Case "imprime"
                Call imprimeGuia(dato0.text, dato1.text, Impresion, data)
            Case "movimientos"
            Case "historico"
            Case "retorno"
                Call retorno
            Case "anterior"
                Call anterior
            Case "siguiente"
                Call siguiente
        End Select
    End Sub
    
    Private Sub modificar()
        modifica = True
        Call HabilitarCajas(Me, modifica)
        dato1.Enabled = False
        dato2.Enabled = False
        dato6.SetFocus
    End Sub
    
    Private Sub eliminar()
        Call eliminarGuiaMolienda(gm)
        Call retorno
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
    End Sub

    Private Sub retorno()
        Call LimpiarCajas(Me)
        Call LimpiarLabels(Me)
        modifica = False
        Call DeshabilitarCajas(Me)
        opciones.Visible = False
        estado = 0
        dato0.Enabled = True
        dato0.SetFocus
        Call selecciona(dato0)
    End Sub
        
    Private Sub anterior()
        If leerGuiaMolienda(gm, dato1.text, "<") = True Then
            structtoctrl
        End If
    End Sub
    
    Private Sub siguiente()
        If leerGuiaMolienda(gm, dato1.text, ">") = True Then
            structtoctrl
        End If
    End Sub

Private Sub opciones_GotFocus()
    manual.SetFocus
End Sub
'=============================================================================
'OPCIONES
'=============================================================================

