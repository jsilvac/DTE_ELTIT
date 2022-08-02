VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9d.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form arriendo05 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contratos De Arriendo De Propiedades"
   ClientHeight    =   10275
   ClientLeft      =   2040
   ClientTop       =   1425
   ClientWidth     =   15240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   685
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   4875
      Left            =   7830
      TabIndex        =   33
      Top             =   180
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   8599
      BackColor       =   16761024
      Caption         =   "Detalle de Arriendos"
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4860
         TabIndex        =   46
         Top             =   4545
         Width           =   1860
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4860
         TabIndex        =   45
         Top             =   4230
         Value           =   -1  'True
         Width           =   1860
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Genera Meses"
         Height          =   330
         Left            =   2835
         TabIndex        =   40
         Top             =   4410
         Width           =   1455
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3930
         Left            =   45
         TabIndex        =   34
         Top             =   270
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   6932
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5085
      Left            =   180
      TabIndex        =   16
      Top             =   180
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8969
      BackColor       =   16744576
      Caption         =   "DATOS  "
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
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   44
         Tag             =   "monedaarriendo"
         Top             =   4725
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "pagoarrienos.frx":0000
         Left            =   1800
         List            =   "pagoarrienos.frx":0002
         TabIndex        =   43
         Text            =   "Combo1"
         Top             =   4320
         Visible         =   0   'False
         Width           =   4470
      End
      Begin VB.TextBox dato15 
         Alignment       =   2  'Center
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
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   35
         Tag             =   "monedaarriendo"
         Top             =   3915
         Width           =   375
      End
      Begin VB.TextBox dato14 
         Alignment       =   2  'Center
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
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "gastoscomunes"
         Top             =   3555
         Width           =   1950
      End
      Begin VB.TextBox dato13 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "diapago"
         Top             =   3195
         Width           =   375
      End
      Begin VB.TextBox dato12 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   11
         Tag             =   "rutarrendatario"
         Top             =   2835
         Width           =   1215
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "numero"
         Top             =   315
         Width           =   1215
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "fechainicio"
         Top             =   1035
         Width           =   375
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "propiedad"
         Top             =   675
         Width           =   1215
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "Comuna"
         Top             =   1035
         Width           =   375
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
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
         Left            =   2610
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "ciudad"
         Top             =   1035
         Width           =   975
      End
      Begin VB.TextBox dato11 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   10
         Tag             =   "rutarrendador"
         Top             =   2475
         Width           =   1215
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "celular"
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
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
         Left            =   2610
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "fax"
         Top             =   1395
         Width           =   975
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fechatermino"
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox dato10 
         Alignment       =   2  'Center
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
         MaxLength       =   1
         TabIndex        =   9
         Tag             =   "tipoarriendo"
         Top             =   2115
         Width           =   375
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   8
         Tag             =   "montoarriendo"
         Top             =   1755
         Width           =   1395
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reajuste %"
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
         Left            =   135
         TabIndex        =   42
         Top             =   4725
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reajuste I.P.C"
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
         Left            =   135
         TabIndex        =   41
         Top             =   4320
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblmoneda 
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
         Left            =   2295
         TabIndex        =   37
         Top             =   3960
         Width           =   3930
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "moneda Arriendo"
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
         Left            =   135
         TabIndex        =   36
         Top             =   3915
         Width           =   1530
      End
      Begin VB.Label lblpropiedad 
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
         Left            =   3000
         TabIndex        =   32
         Top             =   675
         Width           =   4290
      End
      Begin VB.Label lbltipo 
         BackColor       =   &H00FF8080&
         Caption         =   " M - Mensual  -  A - Anual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2205
         TabIndex        =   31
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
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
         Left            =   100
         TabIndex        =   30
         Top             =   1755
         Width           =   1530
      End
      Begin VB.Label lblarrendatario 
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
         Left            =   3360
         TabIndex        =   29
         Top             =   2835
         Width           =   3930
      End
      Begin VB.Label lblarrendador 
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
         Left            =   3360
         TabIndex        =   28
         Top             =   2475
         Width           =   3930
      End
      Begin VB.Label dv2 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   27
         Top             =   2835
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
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
         Left            =   100
         TabIndex        =   26
         Top             =   315
         Width           =   1530
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Propiedad"
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
         Left            =   100
         TabIndex        =   25
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
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
         Left            =   100
         TabIndex        =   24
         Top             =   1035
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Termino"
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
         Left            =   100
         TabIndex        =   23
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Pago"
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
         TabIndex        =   22
         Top             =   2115
         Width           =   1530
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut Arrendatario"
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
         Left            =   105
         TabIndex        =   21
         Top             =   2835
         Width           =   1530
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dia Pago"
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
         Left            =   105
         TabIndex        =   20
         Top             =   3195
         Width           =   1530
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut Arrendador"
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
         Left            =   105
         TabIndex        =   19
         Top             =   2475
         Width           =   1530
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gastos Comunes"
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
         TabIndex        =   18
         Top             =   3555
         Width           =   1530
      End
      Begin VB.Label dv 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   17
         Top             =   2475
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
      ScaleWidth      =   15210
      TabIndex        =   15
      Top             =   10275
      Width           =   15240
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   3435
      Left            =   180
      TabIndex        =   38
      Top             =   5490
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   6059
      BackColor       =   16761024
      Caption         =   "Propiedades en Arriendo"
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
      Begin FlexCell.Grid Grid2 
         Height          =   3030
         Left            =   90
         TabIndex        =   39
         Top             =   270
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5345
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   630
      TabIndex        =   14
      Top             =   9000
      Width           =   6735
      _cx             =   11880
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "arriendo05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private moneda As String
Private rutpropi As String

Private MODIFI As Integer

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Public Sub grabarclick()

Dim mesin As Double
Dim añoin As Double
Dim mesfin As Double
Dim añofin As Double
Dim dia As String
Dim final As Boolean
final = False

dia = dato13.text
añofin = CDbl(dato8.text)
mesfin = CDbl(dato7.text)
añoin = CDbl(dato5.text)
mesin = CDbl(dato4.text)
While final = False

Call grabarmeses(dato1.text, Format(añoin, "0000") & "-" & Format(mesin, "00") & "-" & dato13.text, dato9.text, dato15.text, dato14.text)
mesin = mesin + 1
If mesin = 13 Then añoin = añoin + 1: mesin = 1
If añoin = añofin And mesin = mesfin Then
final = True
End If

Wend
If Format(añoin, "0000") + "-" + Format(mesin, "00") + "-" + dato13.text <= dato8.text + "-" + dato7.text + "-" + dato6.text Then
Call grabarmeses(dato1.text, Format(añoin, "0000") & "-" & Format(mesin, "00") & "-" & dato13.text, dato9.text, dato15.text, dato14.text)
End If
Call leermensualidades(dato1.text)

End Sub
Private Sub dato1_GotFocus()
dato1.text = LEERULTIMOFOLIOcontrato

Call cargatexto(dato1)
End Sub

Private Sub dato2_GotFocus()
If MODIFI = 0 Then Call leer
Call cargatexto(dato2)
End Sub
Private Sub dato4_GotFocus()
Call cargatexto(dato4)
End Sub
Private Sub dato5_GotFocus()
Call cargatexto(dato5)
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
Call cargatexto(dato11)
End Sub
 
Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then Call ayudapropiedades(dato2)
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
        Call flechas(dato5, dato7, KeyCode)
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
  If KeyCode = vbKeyF2 Then Call ayudaarrendadores(dato11)
        Call flechas(dato10, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then Call ayudaarrendatarios(dato12)
        Call flechas(dato11, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato13, dato14, KeyCode)
End Sub


Private Sub Grid1_DblClick()
Dim pagado As String

If Grid1.ActiveCell.col = 5 Then
If Grid1.Cell(Grid1.ActiveCell.row, 5).text = "1" Then
pagado = "0"
Else
pagado = "1"
End If

Call modificapago(dato1.text, Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "yyyy-mm-dd"), pagado)
Grid1.Cell(Grid1.ActiveCell.row, 5).text = pagado

End If

End Sub

Private Sub Grid2_DblClick()
dato1.text = Grid2.Cell(Grid2.ActiveCell.row, 4).text
Call dato1_KeyPress(13)
End Sub

Private Sub Label9_Click()

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
 
Rem Call RECUPERAFECHA

Call CARGAPERMISO(Me.Name)
Call CARGAGRILLA
Call CARGAGRILLA2
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato1.text) <> 0 Then
        Call ceros(dato1)
        Call Pregunta(dato1, dato2)
    CARGAGRILLA2
    End If
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
   snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato2.text) <> 0 Then
        If leepropiedad(dato2.text) <> "" Then
            Call Pregunta(dato2, dato3)
        lblpropiedad.Caption = leepropiedad(dato2.text)
        dato15.text = moneda
        lblmoneda.Caption = leemonedas(dato15.text)
        dato11.text = Mid(rutpropi, 1, 9)
        dv.Caption = Mid(rutpropi, 10, 1)
        lblarrendador.Caption = leearrendador(dato11.text & dv.Caption)
        Else
            MsgBox "PROPIEDAD NO EXISTE O CODIGO MAL INGRESADO", vbCritical, "ATENCION"
        End If
    End If
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
        snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato3)
        If dato3.text = "00" Then dato3.text = Format(fechasistema, "dd")
        Call Pregunta(dato3, dato4)
    End If
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
        snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato4)
        If dato4.text = "00" Then dato4.text = Format(fechasistema, "mm")
        Call Pregunta(dato4, dato5)
    End If
    End Sub
    
Private Sub dato5_KeyPress(KeyAscii As Integer)
         snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato5)
        If dato5.text = "0000" Then dato5.text = Format(fechasistema, "yyyy")
        
          If IsDate(dato5.text & "/" & dato4.text & "/" & dato3.text) Then
            Call Pregunta(dato5, dato6)
          Else
            MsgBox ("fecha digitada incorrecta")
          End If
        
        
       
    End If
    End Sub
    
Private Sub dato6_KeyPress(KeyAscii As Integer)
         snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato6)
        If dato6.text = "00" Then dato6.text = Format(fechasistema, "dd")
        Call Pregunta(dato6, dato7)
    End If
    End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
        snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato7)
        If dato7.text = "00" Then dato7.text = Format(fechasistema, "mm")
        Call Pregunta(dato7, dato8)
    End If
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
      snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato8)
        If dato8.text = "0000" Then dato8.text = Format(fechasistema, "yyyy")
        
               If IsDate(dato8.text & "/" & dato7.text & "/" & dato6.text) Then
            Call Pregunta(dato8, dato9)
          Else
            MsgBox ("fecha digitada incorrecta")
          End If
    End If
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
     snum = 1: KeyAscii = esNumero(KeyAscii)
     
    If KeyAscii = 13 And Val(dato9.text) <> 0 Then
        Call Pregunta(dato9, dato10)
    End If
    End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And LTrim(dato10.text) <> "" And (dato10.text = "M" Or dato10.text = "A") Then Call Pregunta(dato10, dato11)

End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
   snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato11.text) <> 0 Then
        Call ceros(dato11)
        dv.Caption = rut(dato11.text)
        If leearrendador(dato11.text & dv.Caption) <> "" Then
            Call Pregunta(dato11, dato12)
        lblarrendador.Caption = leearrendador(dato11.text & dv.Caption)
        Else
        MsgBox "ARRENDADOR NO EXISTE O RUT MAL INGRESADO", vbCritical, "ATENCION"
        End If
    End If
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
   snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato11.text) <> 0 Then
        Call ceros(dato12)
        dv2.Caption = rut(dato12.text)
        If leearrendatario(dato12.text & dv2.Caption) <> "" Then
            Call Pregunta(dato12, dato13)
         lblarrendatario.Caption = leearrendatario(dato12.text & dv2.Caption)
       
        Else
        MsgBox "ARRENDATARIO NO EXISTE O RUT MAL INGRESADO", vbCritical, "ATENCION"
        End If
    End If
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
   snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato13.text) <> 0 Then
        Call ceros(dato13)
        Call Pregunta(dato13, dato14)
    End If
End Sub
Private Sub dato14_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
      
      If KeyAscii = 13 Then
        dato15.Enabled = True
        
        dato15.SetFocus
         
         
     End If
End Sub
Private Sub dato15_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
      
      If KeyAscii = 13 Then
         Call ceros(dato15)
         If leemonedas(dato15.text) <> "" Then
         lblmoneda.Caption = leemonedas(dato15.text)
     
        If Verifica_Permiso(Me.Caption, "agrega") = True Then
            grabar
            grabarclick
        End If
         retorno
         Else
         MsgBox "TIPO DE MONEDA DE ARRIENDO NO EXISTE "
         dato15.SetFocus
         End If
         
     End If
End Sub

Sub leer()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = dato3.Tag
    CAMPOS(3, 0) = dato6.Tag
    CAMPOS(4, 0) = dato9.Tag
    CAMPOS(5, 0) = dato10.Tag
    CAMPOS(6, 0) = dato11.Tag
    CAMPOS(7, 0) = dato12.Tag
    CAMPOS(8, 0) = dato13.Tag
    CAMPOS(9, 0) = dato14.Tag
    CAMPOS(10, 0) = dato15.Tag
    CAMPOS(11, 0) = ""
    
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".contratos_arriendo"
    condicion = "numero= '" & dato1.text & "' "

    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato2.SetFocus: GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
        
no:
End Sub
Sub leersiguiente()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = dato3.Tag
    CAMPOS(3, 0) = dato6.Tag
    CAMPOS(4, 0) = dato9.Tag
    CAMPOS(5, 0) = dato10.Tag
    CAMPOS(6, 0) = dato11.Tag
    CAMPOS(7, 0) = dato12.Tag
    CAMPOS(8, 0) = dato13.Tag
    CAMPOS(9, 0) = dato14.Tag
    CAMPOS(10, 0) = dato15.Tag
    CAMPOS(11, 0) = ""
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".contratos_arriendo"
    condicion = " " & dato1.Tag & " > '" & dato1.text & "' order by " & dato1.Tag & " asc "

    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    
no:
   
    
End Sub
Sub leeranterior()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = dato3.Tag
    CAMPOS(3, 0) = dato6.Tag
    CAMPOS(4, 0) = dato9.Tag
    CAMPOS(5, 0) = dato10.Tag
    CAMPOS(6, 0) = dato11.Tag
    CAMPOS(7, 0) = dato12.Tag
    CAMPOS(8, 0) = dato13.Tag
    CAMPOS(9, 0) = dato14.Tag
    CAMPOS(10, 0) = dato15.Tag
    CAMPOS(11, 0) = ""
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".contratos_arriendo"
    condicion = " " & dato1.Tag & " < '" & dato1.text & "' order by " & dato1.Tag & " desc "
    
    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    
no:
   
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    
    dato3.text = Mid(sqlconta.response(2, 3), 1, 2)
    dato4.text = Mid(sqlconta.response(2, 3), 4, 2)
    dato5.text = Mid(sqlconta.response(2, 3), 7, 4)
    
    dato6.text = Mid(sqlconta.response(3, 3), 1, 2)
    dato7.text = Mid(sqlconta.response(3, 3), 4, 2)
    dato8.text = Mid(sqlconta.response(3, 3), 7, 4)
    dato15.text = sqlconta.response(10, 3)
    
    If dato15.text = "02" Then
    dato9.text = Format(sqlconta.response(4, 3), "###,###,###.00")
    Else
    dato9.text = Format(sqlconta.response(4, 3), "###,###,###")
    End If
    
    dato10.text = sqlconta.response(5, 3)
    
    dato11.text = Mid(sqlconta.response(6, 3), 1, 9)
    dv.Caption = Mid(sqlconta.response(6, 3), 10, 1)
    
    dato12.text = Mid(sqlconta.response(7, 3), 1, 9)
    dv2.Caption = Mid(sqlconta.response(7, 3), 10, 1)
    
    dato13.text = sqlconta.response(8, 3)
    dato14.text = Format(sqlconta.response(9, 3), "$ ###,###,##0")
    lblmoneda.Caption = leemonedas(dato15.text)
        
    lblpropiedad.Caption = leepropiedad(sqlconta.response(1, 3))
    lblarrendador.Caption = leearrendador(sqlconta.response(6, 3))
    lblarrendatario.Caption = leearrendatario(sqlconta.response(7, 3))
Call leermensualidades(dato1.text)

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
 
 
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudaarrendadores(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("rut", "nombre")
    largo = Array("11s", "40s")
    cfijo = "rut like '%%'"
    cabezas = Array("Rut", "Nombre")
    mensajeAyuda = "Ayuda de Arrendadores"
       
    Call cargaAyudaT(servidor, clientesistema & "arriendos", usuario, password, ".maestro_arrendadores", caja, CAMPOS, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudaarrendatarios(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("rut", "nombre")
    largo = Array("11s", "40s")
    cfijo = "rut like '%%'"
    cabezas = Array("Rut", "Nombre")
    mensajeAyuda = "Ayuda de Arrendatarios"
       
    Call cargaAyudaT(servidor, clientesistema & "arriendos", usuario, password, ".maestro_arrendatarios", caja, CAMPOS, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudapropiedades(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("codigopropiedad", "direccion")
    largo = Array("11s", "40s")
    cfijo = "codigopropiedad like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Propiedades"
       
    Call cargaAyudaT(servidor, clientesistema & "arriendos", usuario, password, ".maestro_propiedades", caja, CAMPOS, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = dato3.Tag
    CAMPOS(3, 0) = dato6.Tag
    CAMPOS(4, 0) = dato9.Tag
    CAMPOS(5, 0) = dato10.Tag
    CAMPOS(6, 0) = dato11.Tag
    CAMPOS(7, 0) = dato12.Tag
    CAMPOS(8, 0) = dato13.Tag
    CAMPOS(9, 0) = dato14.Tag
    CAMPOS(10, 0) = "fecha"
    CAMPOS(11, 0) = dato15.Tag
    CAMPOS(12, 0) = ""
    CAMPOS(0, 1) = dato1.text
    CAMPOS(1, 1) = dato2.text
    CAMPOS(2, 1) = dato5.text & "-" & dato4.text & "-" & dato3.text
    CAMPOS(3, 1) = dato8.text & "-" & dato7.text & "-" & dato6.text
    CAMPOS(4, 1) = Replace(dato9.text, ",", ".")
    CAMPOS(5, 1) = dato10.text
    CAMPOS(6, 1) = dato11.text & dv.Caption
    CAMPOS(7, 1) = dato12.text & dv2.Caption
    CAMPOS(8, 1) = dato13.text
    If dato14.text = "" Then dato14.text = "0"
    CAMPOS(9, 1) = CDbl(dato14.text)
    
    CAMPOS(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    CAMPOS(11, 1) = dato15.text
    
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".contratos_arriendo"
    If MODIFI = 1 Then condicion = "numero ='" & dato1.text & "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
 
Sub ELIMINAR()
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".contratos_arriendo"
    condicion = "numero=" + "'" + dato1.text + "' "
    op = 4
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)


End Sub
  

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then
    If Verifica_Permiso(Me.Caption, "modifica") = True Then
        modifica
    End If
End If

If command = "elimina" Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        elimina
    End If
End If


If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
 



End Sub
Sub elimina()
 
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
dato2.SetFocus
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
 
 
    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    lblpropiedad.Caption = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    lblarrendador.Caption = ""
    dv.Caption = ""
    dato12.text = ""
    dv2.Caption = ""
    lblarrendatario.Caption = ""
    dato13.text = ""
    dato14.text = ""
 
End Sub
 
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub

 Private Function leearrendador(rutarrendador) As String
 Dim cSql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set cSql.ActiveConnection = db
 cSql.sql = "select nombre from " & clientesistema & "arriendos" & ".maestro_arrendadores "
 cSql.sql = cSql.sql & "where rut='" & rutarrendador & "' "
 cSql.Execute
 leearrendador = ""
 
 If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
    leearrendador = resultados(0)
 Else
    lblarrendador = ""
 End If
 
 cSql.Close
 Set cSql = Nothing
 Set resultados = Nothing
 
 End Function
 Private Function leearrendatario(rutarrendatario) As String
 Dim cSql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set cSql.ActiveConnection = db
 cSql.sql = "select nombre from " & clientesistema & "arriendos" & ".maestro_arrendatarios "
 cSql.sql = cSql.sql & "where rut='" & rutarrendatario & "' "
 cSql.Execute
 leearrendatario = ""
 
 If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
    leearrendatario = resultados(0)
    
 Else
    leearrendatario = ""
 End If
 
 cSql.Close
 Set cSql = Nothing
 Set resultados = Nothing
 
 End Function
 
 Private Function leepropiedad(codigopropiedad) As String
 Dim cSql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set cSql.ActiveConnection = db
 cSql.sql = "select direccion,monedaarriendo,rutpropietario from " & clientesistema & "arriendos" & ".maestro_propiedades "
 cSql.sql = cSql.sql & "where codigopropiedad='" & codigopropiedad & "' "
 cSql.Execute
 leepropiedad = ""
 moneda = ""
 If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
    leepropiedad = resultados(0)
    moneda = resultados(1)
    rutpropi = resultados(2)
 Else
    leepropiedad = ""

 End If
 
 cSql.Close
 Set cSql = Nothing
 Set resultados = Nothing
 
 End Function

Sub CARGAGRILLA()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "FECHA"
    formatogrilla2(1, 2) = "MONTO"
    formatogrilla2(1, 3) = "MONEDA"
    formatogrilla2(1, 4) = "G/COMUNES "
    formatogrilla2(1, 5) = "PAGADO"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    formatogrilla2(2, 2) = "10"
    formatogrilla2(2, 3) = "10"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "17"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "D"
    formatogrilla2(3, 2) = "N"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 2) = " ###,###,##0.00"
    formatogrilla2(4, 3) = " ###,###,##0.00"
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 6
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
     Grid1.Column(5).CellType = cellCheckBox
     
     
    End Sub

Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "PROPIEDAD"
    formatogrilla2(1, 3) = "DIRECCION"
    formatogrilla2(1, 4) = "CONTRATO"
    formatogrilla2(1, 5) = "ARRENDATARIO"
    formatogrilla2(1, 6) = "DESDE"
    formatogrilla2(1, 7) = "HASTA"
    formatogrilla2(1, 8) = "MONTO"
    formatogrilla2(1, 9) = "MONEDA"
    formatogrilla2(1, 10) = "G/COMUNES"
    formatogrilla2(1, 11) = "MOROSO"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "5"
    formatogrilla2(2, 2) = "10"
    formatogrilla2(2, 3) = "20"
    formatogrilla2(2, 4) = "8"
    formatogrilla2(2, 5) = "20"
    formatogrilla2(2, 6) = "7"
    formatogrilla2(2, 7) = "7"
    formatogrilla2(2, 8) = "10"
    formatogrilla2(2, 9) = "5"
    formatogrilla2(2, 10) = "8"
    formatogrilla2(2, 11) = "5"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "S"
    formatogrilla2(3, 6) = "D"
    formatogrilla2(3, 7) = "D"
    formatogrilla2(3, 8) = "N"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 8) = " ###,###,##0.00"
    formatogrilla2(4, 10) = " ###,###,##0.00"
    
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    formatogrilla2(5, 10) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 12
    Grid2.Rows = 1
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
    Grid2.BackColorFixedSel = RGB(110, 180, 230)
    Grid2.BackColorBkg = RGB(90, 158, 214)
    Grid2.BackColorScrollBar = RGB(231, 235, 247)
    Grid2.BackColor1 = RGB(231, 235, 247)
    Grid2.BackColor2 = RGB(239, 243, 255)
    Grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
    Grid2.Column(11).CellType = cellCheckBox
    
    
    leerpropiedades
    
    End Sub


 Public Sub leerpropiedades()
 Dim cSql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set cSql.ActiveConnection = db
 cSql.sql = "select mp.codigopropiedad,mp.nombrepropiedad,mp.direccion,ca.numero,ca.rutarrendatario,ca.fechainicio,ca.fechatermino,ca.montoarriendo,ca.monedaarriendo,ca.gastoscomunes  from " & clientesistema & "arriendos" & ".maestro_propiedades as mp left join " + clientesistema + "arriendos" & ".contratos_arriendo as ca on (mp.codigopropiedad = ca.propiedad) order by mp.direccion "
 cSql.Execute
 Grid2.Rows = 1
 If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    For k = 1 To 3
    Grid2.Cell(Grid2.Rows - 1, k).text = resultados(k - 1)
    Next k
    If IsNull(resultados(3)) = False Then
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
    Grid2.Cell(Grid2.Rows - 1, 5).text = leearrendatario(resultados(4))
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
    Grid2.Cell(Grid2.Rows - 1, 9).text = leemonedas(resultados(8))
    Grid2.Cell(Grid2.Rows - 1, 10).text = resultados(9)
    Grid2.Cell(Grid2.Rows - 1, 11).text = arriendoatrasado(resultados(3), Format(fechasistema, "yyyy-mm-dd"))
    If Format(resultados(6), "yyyy-mm-dd") < Format(fechasistema, "yyyy-mm-dd") Then
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 11).BackColor = &HFF&
    
    End If
    
    
    
    Else
    
    Grid2.Cell(Grid2.Rows - 1, 5).text = "*** DISPONIBLE **"
    
    End If
    
    resultados.MoveNext
    
    
    
    Wend
    
    
  End If
 cSql.Close
 Set cSql = Nothing
 Set resultados = Nothing
 
 End Sub

Private Function leemonedas(codigo) As String
Dim cSql As New rdoQuery
Dim resultados As rdoResultset

Set cSql.ActiveConnection = conta

cSql.sql = "select nombremoneda from " & clientesistema & "arriendos" & ".maestro_monedas where codigomoneda='" & codigo & "'"
cSql.Execute
leemonedas = ""
If cSql.RowsAffected > 0 Then
Set resultados = cSql.OpenResultset
leemonedas = resultados(0)
End If
Set resultados = Nothing
cSql.Close
Set cSql = Nothing

End Function

Public Function LEERULTIMOFOLIOcontrato() As String

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = db

            cSql.sql = "select IFNULL(max(numero),0) from " + clientesistema + "arriendos.contratos_arriendo"
            
            cSql.Execute
    If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
    
    
        LEERULTIMOFOLIOcontrato = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function

Sub grabarmeses(numero, fecha, monto, tipoarriendo, gastoscomunes)
    CAMPOS(0, 0) = "numero"
    CAMPOS(1, 0) = "fecha"
    CAMPOS(2, 0) = "monto"
    CAMPOS(3, 0) = "tipoarriendo"
    CAMPOS(4, 0) = "gastoscomunes"
    CAMPOS(5, 0) = ""
    CAMPOS(0, 1) = numero
    CAMPOS(1, 1) = fecha
    If dato10.text = "A" Then monto = Round((monto / 12), 0)
    If dato15.text = "01" Then
    CAMPOS(2, 1) = Replace(monto, ".", "")
    Else
    CAMPOS(2, 1) = Replace(monto, ",", ".")
    End If
    CAMPOS(3, 1) = tipoarriendo
    CAMPOS(4, 1) = CDbl(gastoscomunes)
    
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".arriendos_mensuales"
    op = 2
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub

Public Sub leermensualidades(numero)
 Dim cSql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set cSql.ActiveConnection = db
 cSql.sql = "select * from " & clientesistema & "arriendos" & ".arriendos_mensuales as mp where numero='" + numero + "' "
 If Option1.Value = True Then
 cSql.sql = cSql.sql + " and pagado='0' "
 End If
 cSql.sql = cSql.sql + "order by fecha "
 
 cSql.Execute
 Grid1.Rows = 1
 If cSql.RowsAffected > 0 Then
    Set resultados = cSql.OpenResultset
    While resultados.EOF = False
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(1)
    Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(2)
    Grid1.Cell(Grid1.Rows - 1, 3).text = leemonedas(resultados(3))
    Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(4)
    Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(5)
    resultados.MoveNext
    
    Wend
    
    
  End If
 cSql.Close
 Set cSql = Nothing
 Set resultados = Nothing
 
 End Sub

Sub modificapago(numero, fecha, pagado)
    CAMPOS(0, 0) = "pagado"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 1) = pagado
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".arriendos_mensuales"
    condicion = "numero='" + numero + "' and fecha='" + fecha + "' "
    op = 3
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub

 
Private Sub Option1_Click()
Call leermensualidades(dato1.text)

End Sub

Private Sub Option2_Click()
Call leermensualidades(dato1.text)

End Sub

