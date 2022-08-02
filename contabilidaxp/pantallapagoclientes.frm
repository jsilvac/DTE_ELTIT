VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form venta006 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Productos"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14130
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   664
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   942
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3255
      Left            =   9480
      TabIndex        =   47
      Top             =   240
      Width           =   4215
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFF2F7&
         BorderStyle     =   0  'None
         Caption         =   "Datos personales"
         Height          =   735
         Left            =   -120
         TabIndex        =   53
         Top             =   0
         Width           =   4215
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FBEDE6&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   55
            Tag             =   "codigoproducto"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FBEDE6&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   240
            MaxLength       =   15
            TabIndex        =   54
            Tag             =   "codigoproducto"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MONTO"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   60
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHEQUE"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   120
            Width           =   1335
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
            Index           =   15
            Left            =   2640
            TabIndex        =   58
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
            Index           =   14
            Left            =   2040
            TabIndex        =   57
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
            Index           =   13
            Left            =   2040
            TabIndex        =   56
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            FillColor       =   &H00FFC0C0&
            Height          =   735
            Left            =   -480
            Top             =   720
            Width           =   13455
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   360
         TabIndex        =   52
         Top             =   -240
         Width           =   255
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   4080
         TabIndex        =   51
         Top             =   -240
         Width           =   135
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   2400
         TabIndex        =   50
         Top             =   480
         Width           =   135
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   840
         TabIndex        =   49
         Top             =   480
         Width           =   135
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   13320
         TabIndex        =   48
         Top             =   -240
         Width           =   135
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   735
      Left            =   6360
      TabIndex        =   22
      Top             =   8760
      Width           =   7215
      Begin VB.Label totalfactura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   35
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label iha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label neto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label iva 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label exento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
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
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IHA"
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
         Height          =   255
         Left            =   3840
         TabIndex        =   29
         Top             =   120
         Width           =   1095
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   7215
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
         TabIndex        =   28
         Top             =   1560
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
         TabIndex        =   27
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
         Index           =   5
         Left            =   2640
         TabIndex        =   26
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NETO"
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
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I.V.A"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXENTO"
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
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame datospersonales 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   2415
      Left            =   480
      TabIndex        =   12
      Top             =   120
      Width           =   8535
      Begin VB.Frame TIPOS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TIPO DE DOCUMENTOS"
         Height          =   1575
         Left            =   1680
         TabIndex        =   39
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
         Begin VB.ListBox LISTATIPOS 
            BackColor       =   &H00FDFBE3&
            Height          =   1035
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   37
         Tag             =   "codigoproducto"
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   36
         Tag             =   "codigoproducto"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox DATO4 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   13
         Tag             =   "codigoproducto"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE     :"
         Height          =   255
         Left            =   2520
         TabIndex        =   42
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00EFFDDF&
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
         Height          =   255
         Left            =   3480
         TabIndex        =   41
         Top             =   840
         Width           =   5895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de pago"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2415
         Left            =   -240
         Top             =   240
         Width           =   13695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "RUT :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   5055
      Left            =   3360
      TabIndex        =   9
      Top             =   3480
      Width           =   4215
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TIPO DE DOCUMENTOS"
         Height          =   1575
         Left            =   600
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
         Begin VB.ListBox List2 
            BackColor       =   &H00FDFBE3&
            Height          =   1035
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   13320
         TabIndex        =   20
         Top             =   -240
         Width           =   135
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   720
         TabIndex        =   19
         Top             =   -240
         Width           =   135
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   2400
         TabIndex        =   18
         Top             =   -240
         Width           =   135
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   4080
         TabIndex        =   17
         Top             =   -480
         Width           =   135
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFF2F7&
         Height          =   5175
         Left            =   0
         TabIndex        =   16
         Top             =   -240
         Width           =   255
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FDFBE3&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4680
         Left            =   240
         TabIndex        =   21
         Top             =   0
         Width           =   4335
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
         Index           =   4
         Left            =   2040
         TabIndex        =   11
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
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Top             =   1320
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   2760
      Width           =   4215
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   43
         Tag             =   "codigoproducto"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox DATO12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "codigoproducto"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox DATO11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "codigoproducto"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIPO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MONTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   120
         Width           =   1335
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
         Index           =   3
         Left            =   2640
         TabIndex        =   4
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
         Index           =   2
         Left            =   2040
         TabIndex        =   3
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
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   735
         Left            =   -360
         Top             =   120
         Width           =   13455
      End
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   5775
      Left            =   3480
      Top             =   2880
      Width           =   4215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   8640
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "\\servidor\e\gestion comercial\barra_opciones.swf"
      Src             =   "\\servidor\e\gestion comercial\barra_opciones.swf"
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   2415
      Left            =   2280
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "venta006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub DATO1_GotFocus()
TIPOS.Visible = True

End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
 
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub dato1_LostFocus()
If Val(dato1.text) < 1 Or Val(dato1.text) > CANDO Then dato1.SetFocus: GoTo NO:

DOCUMENTO.Caption = DOCU(Val(dato1.text))
TIPOS.Visible = False

NO:

End Sub

Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then Call ayudaproducto(dato11)
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Rem Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaclientes(dato5)
    Call flechas(dato3, dato5, KeyCode)
End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudadepto(dato6)
    Call flechas(dato4, dato6, KeyCode)
End Sub

Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudalinea(dato7)
    Call flechas(dato5, dato7, KeyCode)
End Sub

Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then Call ayudaimpuesto(dato8)
    Call flechas(dato6, dato8, KeyCode)

End Sub

Private Sub Form_Load()
    Call Conectar_BD
    Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
DOCU(1) = "EFECTIVO"
DOCU(2) = "CHEQUES"
DOCU(3) = "DEPOSITO"
CANDO = 3
For K = 1 To CANDO
LISTATIPOS.AddItem (Str$(K) + "=" + DOCU(K))
Next K

End Sub

Private Sub DATO1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato3, dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call ceros(dato4): Call Pregunta(dato4, dato5)
End Sub

Private Sub dato10_LostFocus()
        
    campos(0, 0) = "codigoproducto"
    campos(1, 0) = "descripcion"
    campos(2, 0) = "pventadetalle"
    campos(3, 0) = ""
    campos(0, 2) = "maestroproductos"
    
    condicion = "codigoproducto = '" & dato10.text & "'"
    op = 5
    Set SQLUTIL.conexion = db
    SQLUTIL.datos = campos
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.ESTADO
    DESCRI.Caption = SQLUTIL.datos(1, 3)
    dato12.text = SQLUTIL.datos(2, 3)
    If status <> 0 Then dato10.SetFocus

End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call ceros(dato5): Call Pregunta(dato5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Pregunta(dato6, dato7)
End Sub
Private Sub dato6_LostFocus()
    campos(0, 0) = "codigolinea"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestrolineas"
    condicion = "codigodepto = '" & dato5.text & "' AND codigolinea = '" & dato6.text & "'"
    op = 5
    Set SQLUTIL.conexion = db
    SQLUTIL.datos = campos
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.ESTADO
    label(2) = SQLUTIL.datos(1, 3)
    If status <> 0 Then dato6.SetFocus
End Sub
Private Sub dato7_LostFocus()
    campos(0, 0) = "codigoimpuesto"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroimpuestos"
    condicion = "codigoimpuesto = '" & dato7.text & "'"
    op = 5
    Set SQLUTIL.conexion = db
    SQLUTIL.datos = campos
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.ESTADO
    label(3) = SQLUTIL.datos(1, 3)
    If status <> 0 Then dato7.SetFocus
End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call ceros(dato7): Call Pregunta(dato7, dato8)

End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato8): Call Pregunta(dato8, dato9)
End Sub

Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato9): Call Pregunta(dato9, dato10)
End Sub

Private Sub dato10_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Call ceros(dato10): Call Pregunta(dato10, dato11)
End Sub

Private Sub dato11_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato11, 0): Call Pregunta(dato11, dato12)
End Sub

Private Sub dato12_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato12, 0): Call Pregunta(dato12, dato13)
End Sub



Private Sub dato13_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ' If KeyAscii = 13 Then GRABAR: leer:
      If KeyAscii = 13 Then Call formato(dato13, 0): calcula: GRABAR2
End Sub

Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub

Sub leer()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = ""
    campos(0, 2) = "maestroproductos"
    condicion = "codigoproducto=" + "'" + dato1.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'UNIDAD MEDIDA
    campos(3, 0) = dato4.Tag 'SECCION
    campos(4, 0) = dato5.Tag 'DEPARTAMENTO
    campos(5, 0) = dato6.Tag 'LINEA
    campos(6, 0) = dato7.Tag 'IMPUESTO
    campos(7, 0) = dato8.Tag 'PRECIO COSTO IVA
    campos(8, 0) = dato9.Tag 'PRECIO VENTA MAYOR
    campos(9, 0) = dato10.Tag 'PRECIO VENTA DETALLE
    campos(10, 0) = dato11.Tag 'STOCK CRITICO
    campos(11, 0) = dato12.Tag 'DESCUENTO
    campos(12, 0) = dato13.Tag 'DATO EXTRA
    campos(13, 0) = dato14.Tag 'UBICACION
    campos(0, 2) = "maestroproductos"
    condicion = "codigoproducto>" + "'" + dato1.text + "' order by codigoproducto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'UNIDAD MEDIDA
    campos(3, 0) = dato4.Tag 'SECCION
    campos(4, 0) = dato5.Tag 'DEPARTAMENTO
    campos(5, 0) = dato6.Tag 'LINEA
    campos(6, 0) = dato7.Tag 'IMPUESTO
    campos(7, 0) = dato8.Tag 'PRECIO COSTO IVA
    campos(8, 0) = dato9.Tag 'PRECIO VENTA MAYOR
    campos(9, 0) = dato10.Tag 'PRECIO VENTA DETALLE
    campos(10, 0) = dato11.Tag 'STOCK CRITICO
    campos(11, 0) = dato12.Tag 'DESCUENTO
    campos(12, 0) = dato13.Tag 'DATO EXTRA
    campos(13, 0) = dato14.Tag 'UBICACION
    campos(0, 2) = "maestroproductos"
    condicion = "codigoproducto<" + "'" + dato1.text + "' order by codigoproducto"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = SQLUTIL.datos(0, 3)
    dato2.text = SQLUTIL.datos(1, 3)
    dato3.text = SQLUTIL.datos(2, 3)
    dato4.text = SQLUTIL.datos(3, 3)
    dato5.text = SQLUTIL.datos(4, 3)
    dato6.text = SQLUTIL.datos(5, 3)
    dato7.text = SQLUTIL.datos(6, 3)
    dato8.text = SQLUTIL.datos(7, 3)
    dato9.text = SQLUTIL.datos(8, 3)
    dato10.text = SQLUTIL.datos(9, 3)
    dato11.text = SQLUTIL.datos(10, 3)
    dato12.text = SQLUTIL.datos(11, 3)
    dato13.text = SQLUTIL.datos(12, 3)
    dato14.text = SQLUTIL.datos(13, 3)

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
Sub Conecta_Maestro_Productos()
    'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro01
        .mp.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=conta01"
    End With
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub ayudaseccion(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigoseccion", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestrosecciones", dato4, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

Sub ayudaimpuesto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigoimpuesto", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestroimpuestos", dato7, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudadepto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigodepto", "nombre")
    cfijo = Array("codigoseccion", dato4.text)
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestrodepartamentos", dato5, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
    

End Sub

Sub ayudalinea(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigolinea", "nombre")
    cfijo = Array("codigodepto", dato5.text)
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestrolineas", dato6, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

Sub ayudaproducto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigoproducto", "descripcion", "pventadetalle")
    cfijo = Array("no")
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestroproductos", dato10, campos, cfijo, 3)
    caja.Enabled = True
    caja.SetFocus

End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub GRABAR()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'UNIDAD MEDIDA
    campos(3, 0) = dato4.Tag 'SECCION
    campos(4, 0) = dato5.Tag 'DEPARTAMENTO
    campos(5, 0) = dato6.Tag 'LINEA
    campos(6, 0) = dato7.Tag 'IMPUESTO
    campos(7, 0) = dato8.Tag 'PRECIO COSTO IVA
    campos(8, 0) = dato9.Tag 'PRECIO VENTA MAYOR
    campos(9, 0) = dato10.Tag 'PRECIO VENTA DETALLE
    campos(10, 0) = dato11.Tag 'STOCK CRITICO
    campos(11, 0) = dato12.Tag 'DESCUENTO
    campos(12, 0) = dato13.Tag 'DATO EXTRA
    campos(13, 0) = dato14.Tag 'UBICACION
    campos(0, 1) = dato1.text 'CODIGO
    campos(1, 1) = dato2.text 'DESCRIPCION
    campos(2, 1) = dato3.text 'UNIDAD MEDIDA
    campos(3, 1) = dato4.text 'SECCION
    campos(4, 1) = dato5.text 'DEPARTAMENTO
    campos(5, 1) = dato6.text 'LINEA
    campos(6, 1) = dato7.text 'IMPUESTO
    campos(7, 1) = dato8.text 'PRECIO COSTO IVA
    campos(8, 1) = dato9.text 'PRECIO VENTA MAYOR
    campos(9, 1) = dato10.text 'PRECIO VENTA DETALLE
    campos(10, 1) = dato11.text 'STOCK CRITICO
    campos(11, 1) = dato12.text 'DESCUENTO
    campos(12, 1) = dato13.text 'DATO EXTRA
    campos(13, 1) = dato14.text 'UBICACION
    
    campos(0, 2) = "maestroproductos"
    If modifi = 1 Then condicion = "codigoproducto=" + "'" + dato1.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
modifi = 0

End Sub
Sub GRABAR2()
Rem GRABAR
dato(1) = dato10.text: colu(1) = 15: tipodato(1) = "s"
dato(2) = DESCRI.Caption: colu(2) = 52: tipodato(2) = "s"
dato(3) = dato11.text: colu(3) = 13: tipodato(3) = "n"
dato(4) = dato12.text: colu(4) = 15: tipodato(4) = "n"
dato(5) = dato13.text: colu(5) = 9: tipodato(5) = "n"
dato(6) = total.Caption: colu(6) = 16: tipodato(6) = "n"
cancolu = 6
Call grilla(List1)
suma
End Sub
Sub ELIMINAR()
    
    campos(0, 2) = "maestroproductos"
    condicion = "codigoproducto=" + "'" + dato1.text + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
End Sub


Private Sub LISTATIPOS_DblClick()
dato1.text = Mid(LISTATIPOS.text, 2, 1)
End Sub

Private Sub Label14_Click()
End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then disponible (True): habilita (False): limpia: opciones.Visible = False: dato1.SetFocus
If command = "modifica" Then disponible (True): habilita (False): dato1.Enabled = False: dato2.SetFocus: modifi = 1
If command = "elimina" Then disponible (True): habilita (False): ELIMINAR: limpia: opciones.Visible = False: dato1.SetFocus
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
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
Sub ayudaclientes(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("rutcliente", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestroclientes", dato1, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
    
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

TOTALES = Int((CDbl(dato11.text) * CDbl(dato12.text)) + 0.5)
descu = Int((TOTALES * CDbl(dato13.text) / 100) + 0.5)
TOTALES = TOTALES - descu
total.Caption = Format(TOTALES, "#,###,###,##0")


End Sub
Sub suma()

End Sub

Private Sub total_Click()
End Sub

