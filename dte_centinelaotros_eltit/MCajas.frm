VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form MCajas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Cajas"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   6945
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6675
      Left            =   45
      TabIndex        =   5
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11774
      BackColor       =   16744576
      Caption         =   "Datos de la Caja"
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
      Begin VB.TextBox Text1 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   37
         Tag             =   "proveedor"
         Top             =   5160
         Width           =   1965
      End
      Begin VB.TextBox Text4 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   35
         Tag             =   "proveedor"
         Top             =   6240
         Width           =   1965
      End
      Begin VB.TextBox Text3 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   33
         Tag             =   "proveedor"
         Top             =   5880
         Width           =   1965
      End
      Begin VB.TextBox Text2 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   31
         Tag             =   "proveedor"
         Top             =   5520
         Width           =   1965
      End
      Begin VB.TextBox DATO13 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   29
         Tag             =   "proveedor"
         Top             =   4740
         Width           =   1965
      End
      Begin VB.TextBox DATO9 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   28
         Tag             =   "proveedor"
         Top             =   3300
         Width           =   1965
      End
      Begin VB.TextBox DATO10 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   27
         Tag             =   "proveedor"
         Top             =   3660
         Width           =   1965
      End
      Begin VB.TextBox DATO11 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   26
         Tag             =   "proveedor"
         Top             =   4020
         Width           =   1965
      End
      Begin VB.TextBox DATO12 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   25
         Tag             =   "proveedor"
         Top             =   4380
         Width           =   1965
      End
      Begin VB.TextBox DATO8 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   24
         Tag             =   "proveedor"
         Top             =   2940
         Width           =   1965
      End
      Begin VB.TextBox DATO7 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   23
         Tag             =   "proveedor"
         Top             =   2580
         Width           =   1965
      End
      Begin VB.TextBox DATO6 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   22
         Tag             =   "proveedor"
         Top             =   2220
         Width           =   1965
      End
      Begin VB.TextBox DATO5 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   21
         Tag             =   "proveedor"
         Top             =   1860
         Width           =   1965
      End
      Begin VB.TextBox DATO4 
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
         Left            =   3825
         MaxLength       =   8
         TabIndex        =   20
         Tag             =   "proveedor"
         Top             =   1500
         Width           =   1965
      End
      Begin VB.TextBox dato3 
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
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   1080
         Width           =   4875
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
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
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
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE DE EQUIPO"
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
         Left            =   120
         TabIndex        =   38
         Top             =   5160
         Width           =   3585
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6840
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAC ADRESS DEL EQUIPO"
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
         Left            =   120
         TabIndex        =   36
         Top             =   6240
         Width           =   3585
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IP SECUNDARIA DEL EQUIPO"
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
         Left            =   120
         TabIndex        =   34
         Top             =   5880
         Width           =   3585
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IP PRINCIPAL DEL EQUIPO"
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
         Left            =   120
         TabIndex        =   32
         Top             =   5520
         Width           =   3585
      End
      Begin VB.Label lbllocal 
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
         Left            =   2400
         TabIndex        =   30
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO FACTURA MANUAL"
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
         Left            =   90
         TabIndex        =   18
         Top             =   1860
         Width           =   3585
      End
      Begin VB.Label lbl9 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO NOTA CREDITO MANUAL"
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
         Left            =   90
         TabIndex        =   17
         Top             =   2580
         Width           =   3585
      End
      Begin VB.Label lbl13 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO BOLETA ELECTRONICA"
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
         Left            =   90
         TabIndex        =   16
         Top             =   3300
         Width           =   3585
      End
      Begin VB.Label lbl17 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO NOTA DEBITO ELECTRONICA"
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
         Left            =   120
         TabIndex        =   15
         Top             =   4020
         Width           =   3585
      End
      Begin VB.Label lbl21 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO COMPROBANTE PAGO"
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
         Left            =   120
         TabIndex        =   14
         Top             =   4740
         Width           =   3585
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descripción"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbl20 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO NOTA CREDITO ELECTRONICA"
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
         Left            =   120
         TabIndex        =   12
         Top             =   4380
         Width           =   3585
      End
      Begin VB.Label lbl16 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO FACTURA ELECTRONICA"
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
         Left            =   90
         TabIndex        =   11
         Top             =   3660
         Width           =   3630
      End
      Begin VB.Label lbl12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO BOLETA FISCAL"
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
         Left            =   90
         TabIndex        =   10
         Top             =   2940
         Width           =   3585
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO NOTA DEBITO MANUAL"
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
         Left            =   90
         TabIndex        =   9
         Top             =   2220
         Width           =   3585
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO BOLETA MANUAL"
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
         Left            =   90
         TabIndex        =   8
         Top             =   1500
         Width           =   3585
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Local"
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código"
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
   End
   Begin XPFrame.FrameXp frmGrabar 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   6720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "G   R   A   B   A   R"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.PictureBox manual 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   0
      Width           =   555
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   6960
      Width           =   6195
      _cx             =   10927
      _cy             =   2355
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
Attribute VB_Name = "MCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Impresora As Printer
    Private c As caja
    Private cargo As Boolean
    Private modifica As Boolean
    Private TIPO As Integer
    Public ancho As Double
    Public alto As Double
    'Private segurity As Boolean



'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'Click
    '========================================================
'    Private Sub cmbFacturas_Click()
'        cmbFacturas.ToolTipText = cmbFacturas.List(cmbFacturas.ListIndex, 0)
'        Set Impresora = Printers(cmbFacturas.ListIndex)
'        lblContFacturas.Caption = Impresora.DriverName
'        lblPortFacturas.Caption = Impresora.Port
'        lblIdFacturas.Caption = cmbFacturas.ListIndex
'        TIPO = 1
'        SendKeys "{Tab}"
'    End Sub
'
'    Private Sub cmbBoletas_Click()
'        cmbBoletas.ToolTipText = cmbBoletas.List(cmbBoletas.ListIndex, 0)
'        Set Impresora = Printers(cmbBoletas.ListIndex)
'        lblContBoletas.Caption = Impresora.DriverName
'        lblPortBoletas.Caption = Impresora.Port
'        lblIdBoletas.Caption = cmbBoletas.ListIndex
'        TIPO = 2
'        SendKeys "{Tab}"
'    End Sub
'
'    Private Sub cmbGuias_Click()
'        cmbGuias.ToolTipText = cmbGuias.List(cmbGuias.ListIndex, 0)
'        Set Impresora = Printers(cmbGuias.ListIndex)
'        lblContGuias.Caption = Impresora.DriverName
'        lblPortGuias.Caption = Impresora.Port
'        lblIdGuias.Caption = cmbGuias.ListIndex
'        TIPO = 3
'        SendKeys "{Tab}"
'    End Sub
'
'    Private Sub cmbNCredito_Click()
'        cmbNCredito.ToolTipText = cmbNCredito.List(cmbNCredito.ListIndex, 0)
'        Set Impresora = Printers(cmbNCredito.ListIndex)
'        lblContNCredito.Caption = Impresora.DriverName
'        lblPortNCredito.Caption = Impresora.Port
'        lblIdNCredito.Caption = cmbNCredito.ListIndex
'        TIPO = 4
'        SendKeys "{Tab}"
'    End Sub
'
'    Private Sub cmbOtros_Click()
'        cmbOtros.ToolTipText = cmbOtros.List(cmbOtros.ListIndex, 0)
'        Set Impresora = Printers(cmbOtros.ListIndex)
'        lblContOtros.Caption = Impresora.DriverName
'        lblPortOtros.Caption = Impresora.Port
'        lblIdOtros.Caption = cmbOtros.ListIndex
'        TIPO = 5
'        SendKeys "{Tab}"
'    End Sub
'    '========================================================
    'Click
    '========================================================

    '========================================================
    'GotFocus
    '========================================================
'    Private Sub cmbFacturas_GotFocus()
'        Call VerificarCajas(Me, cmbFacturas)
'    End Sub
'
'    Private Sub cmbBoletas_GotFocus()
'        Call VerificarCajas(Me, cmbBoletas)
'    End Sub
'
'    Private Sub cmbGuias_GotFocus()
'        Call VerificarCajas(Me, cmbGuias)
'    End Sub
'
'    Private Sub cmbNCredito_GotFocus()
'        Call VerificarCajas(Me, cmbNCredito)
'    End Sub
'
'    Private Sub cmbOtros_GotFocus()
'        Call VerificarCajas(Me, cmbOtros)
'    End Sub
'

    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Local"
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Caja"
    End Sub
    
    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaEmpresa(dato1)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaCaja(dato2, dato1.text)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            lbllocal.Caption = leerNombreEmpresa(dato1.text)
            
            If lbllocal.Caption <> "" Then
                SendKeys "{Tab}"
            Else
                Call selecciona(dato1)
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If leerCaja(c, dato1.text, dato2.text, "=") = True Then
                Call structtoctrl
                cargo = True
            Else
             If Verifica_Permiso(Me.Caption, "agrega") = True Then
               cargo = False
                Call HabilitarCajas(Me, modifica)
                SendKeys "{Tab}"
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                dato1.SelStart = 0
                dato1.SelLength = Len(dato1.text)
                dato1.SetFocus
            End If
                
            End If
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
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
    
    Private Sub dato2_LostFocus()
        Call limpiaBarra(2)
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato5.SetFocus
End If
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato6.SetFocus
End If
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato7.SetFocus
End If
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato8.SetFocus
End If
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato9.SetFocus
End If
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato10.SetFocus
End If
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato11.SetFocus
End If
End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato12.SetFocus
End If
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato13.SetFocus
End If
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato13.SetFocus
End If
End Sub

    Private Sub Form_Activate()
        If segurity = True Then
            seguridad.Show vbModal
            segurity = False
        End If
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub
    
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
        titCaption = Me.Caption
        'segurity = Not Verificar(usuarioSistema, passwordSistema)
        cargo = False
        modifica = False
       
    End Sub
    
    Private Sub frmGrabar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmGrabar)
        frmGrabar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmGrabar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmGrabar)
        frmGrabar.CaptionEstilo3D = Inserted
        If cargo = True Then
            If modifica = True Then
                Call ctrltostruct
            End If
        Else
            Call ctrltostruct
        End If
    End Sub

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        c.codLoc = dato1.text
        c.NUMERO = dato2.text
        c.descripcion = dato3.text
        c.folioboletas = dato4.text
        c.foliofacturas = dato5.text
        c.folionotadebito = dato6.text
        c.folionotacredito = dato7.text
        c.folioboletafiscal = dato8.text
        c.folioboletaelectronica = dato9.text
        c.foliofacturaelectronica = dato10.text
        c.folionotadebitoelectronica = dato11.text
        c.folionotacreditoelectronica = dato12.text
        c.foliocomprobantepagos = dato13.text
        
        
        Call grabarCaja(c, modifica)
        Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        Dim cad As String
        Dim cadena As String
        
        dato1.text = c.codLoc
        lbllocal.Caption = leerNombreEmpresa(dato1.text)
        dato2.text = c.NUMERO
        dato3.text = c.descripcion
        dato4.text = c.folioboletas
        dato5.text = c.foliofacturas
        dato6.text = c.folionotadebito
        dato7.text = c.folionotacredito
        dato8.text = c.folioboletafiscal
        dato9.text = c.folioboletaelectronica
        dato10.text = c.foliofacturaelectronica
        dato11.text = c.folionotadebitoelectronica
        dato12.text = c.folionotacreditoelectronica
        dato13.text = c.foliocomprobantepagos
        
        Call DeshabilitarCajas(Me)
      
        dato2.Enabled = True
        dato2.SetFocus
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

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
            Call ELIMINAR
    End Select
End Sub

'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
               If Verifica_Permiso(Me.Caption, "modifica") = True Then
                    Call modificar
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                End If
            Case "elimina"
            If Verifica_Permiso(Me.Caption, "elimina") = True Then
                If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                   Call ELIMINAR
                End If
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
            End If
            Case "imprime"
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
       
        dato3.SetFocus
    End Sub
    
    Private Sub ELIMINAR()
        frmglosaeliminacion.Show vbModal
        Call eliminarCaja(c)
        Call retorno
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
    End Sub

    Private Sub retorno()
        Call LimpiarCajas(Me)
        Call LimpiarLabels(Me)
        modifica = False
        cargo = False
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
    End Sub
    
    
    Private Sub anterior()
        If leerCaja(c, dato1.text, dato2.text, "<") = True Then
            structtoctrl
        End If
    End Sub
    
    Private Sub siguiente()
        If leerCaja(c, dato1.text, dato2.text, ">") = True Then
            structtoctrl
        End If
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================








