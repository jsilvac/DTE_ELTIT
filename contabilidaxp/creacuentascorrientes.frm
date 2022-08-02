VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{2B5A7812-71D1-4C51-B59B-AA38CD8D6BA3}#6.0#0"; "VB2_SkinControlLt.ocx"
Begin VB.Form maestro02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Cuentas Corrientes"
   ClientHeight    =   8700
   ClientLeft      =   2040
   ClientTop       =   1425
   ClientWidth     =   13350
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   890
   Begin VB2_SkinControlLt.VB2_SkinCtrlLt VB2_SkinCtrlLt1 
      Left            =   480
      Top             =   0
      _ExtentX        =   1111
      _ExtentY        =   953
      SkinPicture     =   "creacuentascorrientes.frx":0000
      ChangeControlColor=   0   'False
      Skin            =   2
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   13320
      TabIndex        =   45
      Top             =   8565
      Width           =   13350
   End
   Begin VB.Frame FECHAS 
      Height          =   1935
      Left            =   6600
      TabIndex        =   21
      Top             =   6360
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton SIIMPRIME 
         BackColor       =   &H00F5C9B1&
         Caption         =   "IMPRIMIR"
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton NOIMPRIME 
         BackColor       =   &H00F5C9B1&
         Caption         =   "NO IMPRIMIR"
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox HASTA3 
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
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   27
         Tag             =   "fecha"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox HASTA2 
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   26
         Tag             =   "fecha"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox HASTA1 
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   25
         Tag             =   "fechavencimiento"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox DESDE1 
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
         Left            =   240
         MaxLength       =   2
         TabIndex        =   24
         Tag             =   "fecha"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox DESDE2 
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
         Left            =   600
         MaxLength       =   2
         TabIndex        =   23
         Tag             =   "fecha"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox DESDE3 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   22
         Tag             =   "fecha"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESDE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HASTA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000040C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0080C0FF&
         BorderWidth     =   3
         Height          =   1935
         Left            =   0
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3735
      Left            =   8400
      TabIndex        =   4
      Top             =   240
      Width           =   4695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SALDOS 
         Height          =   3495
         Left            =   120
         TabIndex        =   12
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
   Begin VB.Frame datospersonales 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7935
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
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   44
         Tag             =   "rut"
         Top             =   840
         Width           =   255
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   43
         Tag             =   "email"
         Top             =   4080
         Width           =   5895
      End
      Begin VB.TextBox dato13 
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
         TabIndex        =   42
         Tag             =   "contacto"
         Top             =   4440
         Width           =   5895
      End
      Begin VB.TextBox dato9 
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
         MaxLength       =   15
         TabIndex        =   41
         Tag             =   "fono"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox dato10 
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
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "fax"
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox dato11 
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
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "celular"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox dato14 
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
         TabIndex        =   18
         Tag             =   "dest_cheque"
         Top             =   4800
         Width           =   255
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
         MaxLength       =   50
         TabIndex        =   15
         Tag             =   "giro"
         Top             =   2640
         Width           =   6015
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   14
         Tag             =   "ciudad"
         Top             =   2280
         Width           =   3255
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   13
         Tag             =   "Comuna"
         Top             =   1920
         Width           =   3255
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
         TabIndex        =   8
         Tag             =   "nombre"
         Top             =   1200
         Width           =   6015
      End
      Begin VB.TextBox dato2 
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
         MaxLength       =   9
         TabIndex        =   7
         Tag             =   "rut"
         Top             =   840
         Width           =   1095
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "direccion"
         Top             =   1560
         Width           =   6015
      End
      Begin VB.TextBox dato1 
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
         TabIndex        =   5
         Tag             =   "tipo"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label glosactacte 
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
         Left            =   2040
         TabIndex        =   46
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Destino Cheque"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contacto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Celular"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Email"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fax"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fono"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Giro"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ciudad"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comuna"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label NOMBRETIPO2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label NOMBRETIPO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Cuenta"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   6135
         Left            =   0
         Top             =   0
         Width           =   7935
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   7200
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      Height          =   6135
      Left            =   360
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "maestro02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub Command1_Click()
IMPRIMIR
End Sub


Private Sub DATO1_GotFocus()
grillasaldos
Call cargatexto(dato1)
End Sub

Private Sub dato2_GotFocus()

Call cargatexto(dato2)
End Sub
Private Sub dato3_GotFocus()
Call cargatexto(dato3)
End Sub
Private Sub dato4_GotFocus()

If modifi = 0 Then leer
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
Call cargatexto(dato14)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo NO:
    If KeyCode = vbKeyF2 Then Call ayudatipocuenta(dato2)
    Call flechas(dato1, dato2, KeyCode)
NO:
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudactacte(dato4)
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

























Private Sub DESDE1_GotFocus()
Call cargatexto(DESDE1)
End Sub
Private Sub DESDE2_GotFocus()
Call cargatexto(DESDE2)
End Sub
Private Sub DESDE3_GotFocus()
Call cargatexto(DESDE3)
End Sub
Private Sub HASTA1_GotFocus()
Call cargatexto(HASTA1)
End Sub
Private Sub HASTA2_GotFocus()
Call cargatexto(HASTA2)
End Sub
Private Sub HASTA3_GotFocus()
Call cargatexto(HASTA3)
End Sub


Private Sub DESDE1_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then Call ceros(DESDE1): Call Pregunta(DESDE1, DESDE2)
End Sub
Private Sub DESDE2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then Call ceros(DESDE2): Call Pregunta(DESDE1, DESDE3)
End Sub
Private Sub DESDE3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DESDE3): Call Pregunta(DESDE2, HASTA1)
End Sub
Private Sub HASTA1_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(HASTA1): Call Pregunta(DESDE3, HASTA2)
End Sub
Private Sub HASTA2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(HASTA2): Call Pregunta(HASTA1, HASTA3)
End Sub
Private Sub HASTA3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(HASTA2): Call Pregunta(HASTA2, DESDE1)
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
VB2_SkinCtrlLt1.ActivateSkin
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
DOCU(1) = "ACTIVO"
DOCU(2) = "PASIVO"
DOCU(3) = "RESULTADO"
CANDO = 3

Call RECUPERAFECHA
Call CARGAPERMISO(Me.Name)

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
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato4, dato5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato6, dato7)
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato7, dato8)
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato8, dato9)
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
    If KeyAscii = 13 Then GRABAR: retorno
End Sub



Sub leer()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato4.Tag
    campos(3, 0) = dato5.Tag
    campos(4, 0) = dato6.Tag
    campos(5, 0) = dato7.Tag
    campos(6, 0) = dato8.Tag
    campos(7, 0) = dato9.Tag
    campos(8, 0) = dato10.Tag
    campos(9, 0) = dato11.Tag
    campos(10, 0) = dato12.Tag
    campos(11, 0) = dato13.Tag
    campos(12, 0) = dato14.Tag
    campos(13, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut=" + "'" + dato2.text + dato3.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 And PERMISOPROGRAMA(2) = "S" Then dato3.SetFocus: GoTo NO:
    If SQLUTIL.ESTADO = 4 And PERMISOPROGRAMA(2) = "N" Then Call NOPERMISO(2): dato2.SetFocus: GoTo NO:

    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
        
NO:
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato4.Tag
    campos(3, 0) = dato5.Tag
    campos(4, 0) = dato6.Tag
    campos(5, 0) = dato7.Tag
    campos(6, 0) = dato8.Tag
    campos(7, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut>" + "'" + dato2.text + dato3.text + "' order by tipo,rut"

    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then GoTo NO:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
    
    
NO:
   
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato4.Tag
    campos(3, 0) = dato5.Tag
    campos(4, 0) = dato6.Tag
    campos(5, 0) = dato7.Tag
    campos(6, 0) = dato8.Tag
    campos(7, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo<" + "'" + dato1.text + "' and rut<" + "'" + dato2.text + dato3.text + "' order by tipo,rut desc"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then GoTo NO:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    DATOSSALDOS
    
    
NO:
   
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = SQLUTIL.datos(0, 3)
    dato2.text = Mid(SQLUTIL.datos(1, 3), 1, 9)
    dato3.text = Mid(SQLUTIL.datos(1, 3), 10, 1)
    dato4.text = SQLUTIL.datos(2, 3)
    dato5.text = SQLUTIL.datos(3, 3)
    dato6.text = SQLUTIL.datos(4, 3)
    dato7.text = SQLUTIL.datos(5, 3)
    dato8.text = SQLUTIL.datos(6, 3)
    dato9.text = SQLUTIL.datos(7, 3)
    dato10.text = SQLUTIL.datos(8, 3)
    dato11.text = SQLUTIL.datos(9, 3)
    dato12.text = SQLUTIL.datos(10, 3)
    dato13.text = SQLUTIL.datos(11, 3)
    dato14.text = SQLUTIL.datos(12, 3)

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


Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "no"
    
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentasdelmayor", PIVOTE, campos, cfijo, largo, 2)
    If Val(PIVOTE.text) = 0 Then dato1.SetFocus: GoTo NO
    dato2.Enabled = True
    dato3.Enabled = True
    dato1.text = Mid(PIVOTE.text, 1, 2)
    dato2.text = Mid(PIVOTE.text, 3, 2)
    dato3.text = Mid(PIVOTE.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
    If sl = 0 Then leer
sl = 0
NO:
End Sub
Sub ayudatipocuenta(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("ctacte", "glosa")
    largo = Array("8s", "40s")
    cfijo = "CTACTE > '00'"
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentasdelmayor", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub GRABAR()
    
    
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato4.Tag
    campos(3, 0) = dato5.Tag
    campos(4, 0) = dato6.Tag
    campos(5, 0) = dato7.Tag
    campos(6, 0) = dato8.Tag
    campos(7, 0) = dato9.Tag
    campos(8, 0) = dato10.Tag
    campos(9, 0) = dato11.Tag
    campos(10, 0) = dato12.Tag
    campos(11, 0) = dato13.Tag
    campos(12, 0) = dato14.Tag
    
    
    campos(13, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text + dato3.text
    campos(2, 1) = dato4.text
    campos(3, 1) = dato5.text
    campos(4, 1) = dato6.text
    campos(5, 1) = dato7.text
    campos(6, 1) = dato8.text
    campos(7, 1) = dato9.text
    campos(8, 1) = dato10.text
    campos(9, 1) = dato11.text
    campos(10, 1) = dato12.text
    campos(11, 1) = dato13.text
    campos(12, 1) = dato14.text
    campos(0, 2) = "cuentascorrientes"
    If modifi = 1 Then condicion = "tipo=" + "'" + dato1.text + "' and rut ='" + dato2.text + dato3.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If modifi = 0 Then GRABAR2
        
End Sub
Sub GRABAR2()
      
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
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text + dato3.text
    campos(2, 1) = año

    For K = 3 To 28
    campos(K, 1) = "0"
    Next K
    campos(0, 2) = "saldosctacte"
    op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    

End Sub

Sub ELIMINAR()
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + dato1.text + "' and rut=" + "'" + dato2.text + dato3.text + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub


Private Sub NOIMPRIME_Click()
FECHAS.Visible = False

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" And PERMISOPROGRAMA(3) = "N" Then Call NOPERMISO(3)
If command = "modifica" And PERMISOPROGRAMA(3) = "S" Then disponible (True): habilita (False): dato1.Enabled = False: dato2.Enabled = False: dato3.SetFocus: modifi = 1
If command = "elimina" And PERMISOPROGRAMA(4) = "N" Then Call NOPERMISO(4)
If command = "elimina" And PERMISOPROGRAMA(4) = "S" Then disponible (True): habilita (False): ELIMINAR: limpia: opciones.Visible = False: dato1.SetFocus
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "imprime" Then IMPRIMIR
If command = "movimientos" Then CARTOLA

End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
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
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
End Sub

Sub IMPRIMIR()
    informes.info.Clear
    largopagina = 65
    tituloinforme = "plan de cuentas"
    titu(1) = "CODIGO"
    titu(2) = "NOMBRE DE CUENTA"
    titu(3) = "TIPO"
    titu(4) = "CTACTE"
    titu(5) = "NOMBRE CTACTE"
    titu(6) = "AUXILIAR"
    lineas = 70
    Consulta_Informe
    informes.Show
    
End Sub
Sub grilla()
    palabra = ""
      
    For K = 1 To cancolu
    If tipodato(K) = "s" Or tipodato(K) = "S" Then dato(K) = dato(K) & String(colu(K) - Len(dato(K)), 32)
    If tipodato(K) = "n" Or tipodato(K) = "N" Then dato(K) = String(colu(K) - Len(dato(K)), 32) & dato(K)
    palabra = palabra & dato(K)
    Next K
    If lineas > largopagina Then Call cabeza
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (" ")
    If Mid(dato(1), 7, 4) <> "0000" Then informes.info.AddItem ("    " + palabra)
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (Mid(palabra, 1, 40))
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (" ")
    lineas = lineas + 1

End Sub
Sub cabeza()
    informes.info.AddItem ("")
    informes.info.AddItem ("")
    pagina = pagina + 1
    


    informes.info.AddItem ("NOMBRE EMPRESA          " + tituloinforme + "                                   PAGINA " + Str$(pagina))
    informes.info.AddItem ("DIRECCION EMPRESA                                                                              " + Mid(Date$, 4, 2) + "-" + Mid(Date$, 1, 2) + "-" + Mid(Date$, 7, 4))
    informes.info.AddItem ("RUT EMPRESA                                                                                    " + Time$)
    informes.info.AddItem ("                                " + tituloinforme)
    informes.info.AddItem String(132, "_")
    TITULOS = ""
    For K = 1 To cancolu
    titu(K) = titu(K) & String(colu(K) - Len(titu(K)), 32)
    TITULOS = TITULOS & titu(K)
    Next K
    informes.info.AddItem (TITULOS)
    informes.info.AddItem String(132, "_")

lineas = 8

End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                
                dato(1) = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4): colu(1) = 15: tipodato(1) = "s"
                dato(2) = resultados(1): colu(2) = 52: tipodato(2) = "s"
                dato(3) = resultados(2) + " " + DOCU$(Val(resultados(2)))
                dato(4) = resultados(3)
                dato(5) = resultados(4)
                dato(6) = resultados(5) + " " + DOCU2$(Val(resultados(5)))
                colu(3) = 10: tipodato(3) = "s"
                colu(4) = 10: tipodato(4) = "s"
                colu(5) = 20: tipodato(5) = "s"
                colu(6) = 10: tipodato(6) = "s"
                 cancolu = 6
                grilla
                resultados.MoveNext
            Wend
            resultados.Close
            
            Set resultados = Nothing

        End If
    End With

End Sub

Sub DATOSSALDOS()

LEERSALDOS
SUMADOR = Val(SQLUTIL.datos(3, 3)) - Val(SQLUTIL.datos(4, 3))
SALDOS.TextMatrix(1, 1) = Format(SQLUTIL.datos(3, 3), "###,###,##0")
SALDOS.TextMatrix(1, 2) = Format(SQLUTIL.datos(4, 3), "###,###,##0")
SALDOS.TextMatrix(1, 3) = Format(SUMADOR, "###,###,##0")
For K = 5 To 16
SALDOS.TextMatrix(K - 3, 1) = Format(SQLUTIL.datos(K, 3), "###,###,##0")
SALDOS.TextMatrix(K - 3, 2) = Format(SQLUTIL.datos(K + 12, 3), "###,###,##0")
SUMADOR = SUMADOR + Val(SQLUTIL.datos(K, 3)) - Val(SQLUTIL.datos(K + 12, 3))
SALDOS.TextMatrix(K - 3, 3) = Format(SUMADOR, "###,###,##0")
Next K

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
For K = 1 To 13
SALDOS.TextMatrix(K, 1) = "0"
SALDOS.TextMatrix(K, 2) = "0"
SALDOS.TextMatrix(K, 3) = "0"
Next K
End Sub

Sub LEERSALDOS()
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
    condicion = "tipo=" + "'" + dato1.text + "' and rut='" + dato2.text + dato3.text + "' and año='" + año + "'"
    campos(0, 2) = "saldosctacte"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 4 Then Stop
grillasaldos
End Sub

Sub CARTOLA()
FECHAS.Visible = True
DESDE1.text = "01"
DESDE2.text = mes
DESDE3.text = año
HASTA1.text = "31"
HASTA2.text = mes
HASTA3.text = año

DESDE1.SetFocus
End Sub

Sub movimientos()



cartolas.Caption = "CARTOLA CUENTA DEL MAYOR"
cartolas.titulocartola = dato1.text + "." + dato2.text + "." + dato3.text + "   " + dato4.text

cartolas.grilla.Cols = 15
cartolas.grilla.Rows = 2
cartolas.grilla.ColWidth(0) = 120 * 8
cartolas.grilla.ColWidth(1) = 120 * 3
cartolas.grilla.ColWidth(2) = 120 * 10
cartolas.grilla.ColWidth(3) = 120 * 4
cartolas.grilla.ColWidth(4) = 120 * 8
cartolas.grilla.ColWidth(5) = 120 * 25
cartolas.grilla.ColWidth(6) = 120 * 3
cartolas.grilla.ColWidth(7) = 120 * 10
cartolas.grilla.ColWidth(8) = 120 * 8
cartolas.grilla.ColWidth(9) = 120 * 10
cartolas.grilla.ColWidth(10) = 120 * 10
cartolas.grilla.ColWidth(11) = 120 * 10
cartolas.Show


Rem TITULOS
cartolas.grilla.TextMatrix(0, 0) = "FECHA"
cartolas.grilla.TextMatrix(0, 1) = "TIPO"
cartolas.grilla.TextMatrix(0, 2) = "NUMERO"
cartolas.grilla.TextMatrix(0, 3) = "LINEA"
cartolas.grilla.TextMatrix(0, 4) = "CUENTA"
cartolas.grilla.TextMatrix(0, 5) = "GLOSA"
cartolas.grilla.TextMatrix(0, 6) = "TD"
cartolas.grilla.TextMatrix(0, 7) = "NUMERO"
cartolas.grilla.TextMatrix(0, 8) = "VENCIMIENTO"
cartolas.grilla.TextMatrix(0, 9) = "DEBE"
cartolas.grilla.TextMatrix(0, 10) = "HABER"
cartolas.grilla.TextMatrix(0, 11) = "SALDO"
LEERMOVIMIENTOS
GoTo NO:
SUMADOR = Val(SQLUTIL.datos(2, 3)) - Val(SQLUTIL.datos(3, 3))
SALDOS.TextMatrix(1, 1) = Format(SQLUTIL.datos(2, 3), "###,###,##0")
SALDOS.TextMatrix(1, 2) = Format(SQLUTIL.datos(3, 3), "###,###,##0")
SALDOS.TextMatrix(1, 3) = Format(SUMADOR, "###,###,##0")
For K = 4 To 15
SALDOS.TextMatrix(K - 2, 1) = Format(SQLUTIL.datos(K, 3), "###,###,##0")
SALDOS.TextMatrix(K - 2, 2) = Format(SQLUTIL.datos(K + 12, 3), "###,###,##0")
SUMADOR = SUMADOR + Val(SQLUTIL.datos(K, 3)) - Val(SQLUTIL.datos(K + 12, 3))
SALDOS.TextMatrix(K - 2, 3) = Format(SUMADOR, "###,###,##0")
Next K
NO:
End Sub

Sub LEERMOVIMIENTOS()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh "
        cSql.SQL = cSql.SQL + "FROM movimientoscontables"
        PIVOTE.text = dato1.text + dato2.text + dato3.text
        cSql.SQL = cSql.SQL + " where codigocuenta = " + "'" + PIVOTE.text + "' AND FECHA>=" + "'" + DESDE3.text + DESDE2.text + DESDE1.text + "' and fecha<=" + "'" + HASTA3.text + HASTA2.text + HASTA1.text + "'  ORDER BY FECHA "
    
        cSql.Execute
        linea = 0: SUMADOR = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                linea = linea + 1
                cartolas.grilla.Rows = linea + 2
                For K = 0 To 8
                cartolas.grilla.TextMatrix(linea, K) = resultados(K)
                Next K
                If resultados(10) = "D" Then cartolas.grilla.TextMatrix(linea, 9) = Format(resultados(K), "###,###,##0"): SUMADOR = SUMADOR + Val(resultados(K))
                If resultados(10) = "H" Then cartolas.grilla.TextMatrix(linea, 10) = Format(resultados(K), "###,###,##0"): SUMADOR = SUMADOR - Val(resultados(K))
                cartolas.grilla.TextMatrix(linea, 11) = Format(SUMADOR, "###,###,##0"):
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing

        End If
    End With

End Sub

Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus

End Sub

Private Sub SIIMPRIME_Click()
FECHAS.Visible = False

movimientos

End Sub
Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & dato1.text & "'"

    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentascorrientes", PIVOTE, campos, cfijo, largo, 2)

    If Val(PIVOTE.text) = 0 Then dato2.SetFocus: GoTo NO
    dato3.Enabled = True
    dato2.text = Mid(PIVOTE.text, 1, 9)
    dato3.text = Mid(PIVOTE.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus

NO:

End Sub

Sub LEETIPOCTACTE()
    campos(0, 0) = "ctacte"
    campos(1, 0) = "glosa"
    campos(2, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "ctacte=" + "'" + dato1.text + "' order by ctacte "
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    
    Call SQLUTIL.SQLUTIL(op, condicion)

   If SQLUTIL.ESTADO = 4 Then MENSAJE ("TIPO DE CUENTA CORRIENTE NO ESTA CREADO "): dato1.SetFocus: GoTo NO:
   glosactacte.Caption = SQLUTIL.datos(1, 3)
   
NO:
End Sub

