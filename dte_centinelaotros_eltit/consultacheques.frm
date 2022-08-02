VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form consultacheques 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Agregar Pre-Venta"
   ClientHeight    =   9375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin XPFrame.FrameXp FRMCHEQUES 
      Height          =   9165
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   16166
      BackColor       =   8438015
      Caption         =   "Datos Cheque"
      CaptionEstilo3D =   1
      BackColor       =   8438015
      ColorBarraArriba=   12648447
      ColorBarraAbajo =   33023
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
      Begin XPFrame.FrameXp muestradetalle 
         Height          =   2055
         Left            =   2280
         TabIndex        =   70
         Top             =   3240
         Visible         =   0   'False
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3625
         BackColor       =   33023
         Caption         =   "Respuesta Cheque"
         CaptionEstilo3D =   1
         BackColor       =   33023
         ForeColor       =   8438015
         ColorBarraArriba=   12648447
         ColorBarraAbajo =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSForms.CommandButton CommandButton1 
            Height          =   375
            Left            =   3480
            TabIndex        =   72
            Top             =   1560
            Width           =   2775
            Caption         =   "cerrar"
            Size            =   "4895;661"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label glosarespuesta 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Width           =   9255
         End
      End
      Begin VB.CommandButton detalle 
         Caption         =   "Ver detalle Respuesta"
         Height          =   495
         Left            =   4920
         TabIndex        =   69
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin XPFrame.FrameXp leecheques 
         Height          =   2055
         Left            =   2280
         TabIndex        =   65
         Top             =   1200
         Visible         =   0   'False
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3625
         BackColor       =   33023
         Caption         =   "Ingrese cheque"
         CaptionEstilo3D =   1
         BackColor       =   33023
         ForeColor       =   8438015
         ColorBarraArriba=   12648447
         ColorBarraAbajo =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox datoscheques 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   360
            TabIndex        =   66
            Top             =   1320
            Width           =   8895
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "El ingreso del cheque debe ser sólo a través del lector de cheques"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   9255
         End
      End
      Begin VB.TextBox nombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   60
         Tag             =   "proveedor"
         Top             =   3060
         Width           =   5220
      End
      Begin VB.CommandButton leecheque 
         BackColor       =   &H000080FF&
         Caption         =   "F5 Leer Cheque "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   8655
         Left            =   7290
         TabIndex        =   37
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   15266
         BackColor       =   8438015
         Caption         =   "HISTORICO CHEQUES RECIBIDOS"
         CaptionEstilo3D =   1
         BackColor       =   8438015
         ColorBarraArriba=   8438015
         ColorBarraAbajo =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.TextBox GLOSA 
            BackColor       =   &H00C0FFFF&
            Height          =   555
            Left            =   135
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   7500
            Width           =   5550
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H000080FF&
            Caption         =   "HISTORICO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3330
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   8160
            Width           =   2085
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H000080FF&
            Caption         =   "PENDIENTES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   675
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   8160
            Width           =   2085
         End
         Begin FlexCell.Grid Grid1 
            Height          =   4425
            Left            =   90
            TabIndex        =   38
            Top             =   270
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   7805
            AllowUserResizing=   0   'False
            Appearance      =   0
            BackColor1      =   12648447
            BackColor2      =   12648447
            BackColorBkg    =   12648447
            BackColorFixed  =   8438015
            BackColorFixedSel=   8438015
            BackColorScrollBar=   8438015
            BackColorSel    =   33023
            Cols            =   5
            DefaultFontSize =   8.25
            FixedRowColStyle=   0
            ForeColorFixed  =   12582912
            GridColor       =   33023
            Rows            =   30
         End
         Begin VB.Label lblprotestos 
            Alignment       =   2  'Center
            BackColor       =   &H00FFF2F7&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   975
            Left            =   120
            TabIndex        =   62
            Top             =   6240
            Width           =   5535
         End
         Begin VB.Label CHEQUEPROMEDIO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   3915
            TabIndex        =   54
            Top             =   5850
            Width           =   1815
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHEQUE PROMEDIO"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3915
            TabIndex        =   53
            Top             =   5625
            Width           =   1815
         End
         Begin VB.Label CHEQUESRECIBIDOS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2070
            TabIndex        =   52
            Top             =   5850
            Width           =   1815
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHEQUES RECIBIDOS"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2070
            TabIndex        =   51
            Top             =   5625
            Width           =   1815
         End
         Begin VB.Label TOTALHISTORICO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   135
            TabIndex        =   50
            Top             =   5850
            Width           =   1815
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL HISTORICO"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   135
            TabIndex        =   49
            Top             =   5625
            Width           =   1815
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "COMENTARIOS DE LA CUENTA"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   48
            Top             =   7275
            Width           =   5550
         End
         Begin VB.Label CUPODISPONIBLE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   3915
            TabIndex        =   47
            Top             =   4995
            Width           =   1815
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CUPO DISPONIBLE"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3915
            TabIndex        =   46
            Top             =   4770
            Width           =   1815
         End
         Begin VB.Label CUPOUTILIZADO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2040
            TabIndex        =   45
            Top             =   4995
            Width           =   1815
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CUPO UTILIZADO"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2025
            TabIndex        =   44
            Top             =   4770
            Width           =   1815
         End
         Begin VB.Label CUPOAUTORIZADO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   90
            TabIndex        =   43
            Top             =   4995
            Width           =   1815
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CUPO AUTORIZADO"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   135
            TabIndex        =   41
            Top             =   4770
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdf4 
         BackColor       =   &H000080FF&
         Caption         =   "F4"
         Height          =   285
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6090
         Width           =   495
      End
      Begin VB.CommandButton cmdf3 
         BackColor       =   &H000080FF&
         Caption         =   "F3"
         Height          =   285
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   5610
         Width           =   495
      End
      Begin VB.TextBox dato10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1665
         MaxLength       =   7
         TabIndex        =   26
         Tag             =   "proveedor"
         Top             =   1485
         Width           =   1575
      End
      Begin VB.TextBox dato9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   24
         Tag             =   "proveedor"
         Top             =   4560
         Width           =   2700
      End
      Begin XPFrame.FrameXp frmmanual 
         Height          =   2505
         Left            =   90
         TabIndex        =   19
         Top             =   6480
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   4419
         BackColor       =   8438015
         Caption         =   "Respuesta"
         CaptionEstilo3D =   1
         BackColor       =   8438015
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComctlLib.ProgressBar barra 
            Height          =   195
            Left            =   0
            TabIndex        =   63
            Top             =   1935
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.TextBox NOMBRECLIENTE 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   2295
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   59
            Top             =   1575
            Width           =   4560
         End
         Begin VB.CommandButton cmdF2 
            BackColor       =   &H000080FF&
            Caption         =   "F2"
            Height          =   285
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   2160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox GLOSARECEPCION 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   2835
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   36
            Top             =   315
            Width           =   3975
         End
         Begin VB.TextBox codigorecepcion 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   2295
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   32
            Top             =   315
            Width           =   375
         End
         Begin VB.TextBox lblcodigoautorizacion 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   2295
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   31
            Top             =   1170
            Width           =   4560
         End
         Begin VB.TextBox lblrespuesta 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   30
            Top             =   720
            Width           =   4560
         End
         Begin VB.Label Label10 
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   " Nombre Cliente"
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
            Left            =   45
            TabIndex        =   58
            Top             =   1575
            Width           =   2235
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Imprimir Cheque"
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
            Left            =   2610
            TabIndex        =   57
            Top             =   2160
            Visible         =   0   'False
            Width           =   3705
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   " Codigo Autorizacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            TabIndex        =   33
            Top             =   360
            Width           =   2250
         End
         Begin VB.Label Label2 
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   " Codigo Autorizacion"
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
            Left            =   45
            TabIndex        =   21
            Top             =   1215
            Width           =   2235
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F5C9B1&
            BackStyle       =   0  'Transparent
            Caption         =   " Glosa Autorizacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            TabIndex        =   20
            Top             =   765
            Width           =   2250
         End
      End
      Begin VB.CommandButton cmdF1 
         BackColor       =   &H000080FF&
         Caption         =   " F1"
         Height          =   285
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5130
         Width           =   495
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "proveedor"
         Top             =   4095
         Width           =   720
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   2745
         MaxLength       =   2
         TabIndex        =   15
         Tag             =   "proveedor"
         Top             =   4095
         Width           =   450
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   960
         Width           =   2700
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   480
         Width           =   1740
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1665
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   1980
         Width           =   630
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   2250
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   4095
         Width           =   450
      End
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1665
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   2475
         Width           =   810
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   3615
         Width           =   1665
      End
      Begin VB.Label Label18 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheques mayores a  $ 250.000 deben ser autorizados por supervisor "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4800
         TabIndex        =   67
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Caption         =   "       * * * I M P O R T A N T E  * * *    INGRESAR RUT DUEÑO DE LA CUENTA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   64
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label13 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   61
         Top             =   3105
         Width           =   1335
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Autorizacion Sin Seguro"
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
         Left            =   765
         TabIndex        =   35
         Top             =   6090
         Width           =   3660
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo Autorizado por Internet"
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
         Left            =   765
         TabIndex        =   29
         Top             =   5610
         Width           =   3660
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Numero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   27
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Fono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   25
         Top             =   4590
         Width           =   1335
      End
      Begin VB.Label lblplaza 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   2565
         TabIndex        =   23
         Top             =   2475
         Width           =   4470
      End
      Begin VB.Label lblbanco 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   2430
         TabIndex        =   22
         Top             =   1980
         Width           =   4470
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Consulta Cheques Automaticos"
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
         Left            =   735
         TabIndex        =   18
         Top             =   5130
         Width           =   3705
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   3480
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lbl7 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   4095
         Width           =   1935
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Rut"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl4 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   9
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Plaza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   3615
         Width           =   1215
      End
   End
   Begin VB.TextBox pivote 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "consultacheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private PASADA As Double
    Private segundos As Double
    Private horainicio As String
    Private nombreclienteconsulta As String

    
    
    Private sc As Integer
    Private TABLAIMPUESTOS(10, 3) As Double
    Private fintablaimpuestos
    Public FOLIOFACTURA  As String
    Private imprimio As Boolean
    Private campos(30, 5) As String
    Private matrix(1000, 10) As String
    Private matrixaux(1000, 10) As String
    Private NUMERO As String
    Public numero1 As String
    Private modificar As Boolean
    Public datoscheque As String
    Public IMPRESORABOLETA As String
    Public rutCliente  As String
    Public consultacheque As String
    Public caja As String
    

Private Sub cmdF1_Click()
    Call Form_KeyDown(vbKeyF1, 0)
End Sub

Private Sub cmdF8_Click()
    
    Call Form_KeyDown(vbKeyF8, 0)
    
End Sub

Private Sub cmdF9_Click()
Call Form_KeyDown(vbKeyF9, 0)
End Sub

Private Sub cmdF2_Click()
Dim MONTO As String
MONTO = WORDNUM(Format(CDbl(dato5.text), "#.###.###0"), "PESO", "PESOS", 0)


'estadokc4 = PuntoVenta.KC4.EscribirCheque(dato5.text, 2, _
                                 "4      Septiembre             2007", 2, _
                                 "FERRETERIA ELTIT LIMITADA", 1, _
                                 MONTO, 1, _
                                 "-------------------", 1)
End Sub

Private Sub cmdf3_Click()
    Call Form_KeyDown(vbKeyF3, 0)
    
End Sub
Private Sub cmdF4_Click()
    Call Form_KeyDown(vbKeyF4, 0)

End Sub

Private Sub Command1_Click()
  Call LEERCHEQUES(dato2.text, dato3.text, "n")
dato4.SetFocus

End Sub

Private Sub Command2_Click()
  Call LEERCHEQUES(dato2.text, dato3.text, "s")
dato4.SetFocus

End Sub

Private Sub CUENTASVIGENTES_Click()

End Sub

Private Sub CommandButton1_Click()
muestradetalle.Visible = False

End Sub

Private Sub dato1_GotFocus()
    Call VerificarCajas(Me, dato1)
    
    Call selecciona(dato1)
End Sub

Private Sub dato1_LostFocus()

leecheque.Visible = True
End Sub

Private Sub dato2_GotFocus()
    'Call selecciona(dato2)
    Call leecheque_Click
End Sub

Private Sub dato3_GotFocus()
    Call selecciona(dato3)
End Sub

Private Sub dato4_GotFocus()
    Call selecciona(dato4)
End Sub

Private Sub dato5_GotFocus()
    Call selecciona(dato5)
End Sub

Private Sub dato5_LostFocus()
If CDbl(dato5.text) > MONTOMAXIMOCHEQUE Then
            If MsgBox("EL MONTO DEL CHEQUE SUPERA LA CANTIDAD PERMITIDA , NECESITA AUTORIZACION DE SUPERVISOR DE CAJA ", vbYesNo, "ATENCION") = vbYes Then
                    autorizador = False
                    permiso.Show vbModal
                    If autorizador = False Then
                     dato5.SetFocus
                    Exit Sub
                    End If
             End If
        End If
      
        If CUPOUTILIZADO.Caption = "" Then CUPOUTILIZADO.Caption = "0"
If CDbl(CUPOUTILIZADO.Caption) + CDbl(dato5.text) > 800000 Then
            If MsgBox("CLIENTE EXCEDE CUPO , NECESITA AUTORIZACION DE SUPERVISOR DE CAJAS", vbYesNo, "ATENCION") = vbYes Then
                    autorizador = False
                    permiso.Show vbModal
                    If autorizador = False Then
                     dato5.SetFocus
                    Exit Sub
                    End If
             End If
        End If


End Sub

Private Sub dato6_GotFocus()
'        If CDbl(dato5.text) > CDbl(CUPODISPONIBLE.Caption) Then
'        Call mensaje.mostrarMensaje("CUPO EXCEDIDO", "EL MONTO DEL CHEQUE ES MAYOR QUE EL CUPO DISPONIBLE", "VERIFIQUE CON TESORERIA")
'        dato5.SetFocus
'        End If
        
    Call selecciona(dato6)
End Sub

Private Sub dato7_GotFocus()
    Call selecciona(dato7)
End Sub

Private Sub dato8_GotFocus()
    Call selecciona(dato8)
End Sub


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, dato1)
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, dato1)
    If KeyCode = vbKeyF5 Then
    Call leecheque_Click
    End If
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, dato2)
End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, dato3)
End Sub

 

Private Sub FrameXp2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

End Sub

Private Sub datoscheques_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And Val(datoscheques.text) <> 0 Then
        dato10.text = Format(Mid(datoscheques.text, 1, 7), "0000000")
        dato3.text = Mid(datoscheques.text, 8, 3)
        dato4.text = Mid(datoscheques.text, 11, 4)
        dato2.text = Mid(datoscheques.text, 15, 11)
        
        Call dato3_KeyPress(13)
        Call dato4_KeyPress(13)
        leecheques.Visible = False
        nombre.SetFocus
        datoscheques.text = ""
        
        End If



End Sub



Private Sub detalle_Click()
muestradetalle.Visible = True

End Sub

Private Sub NOMBRE_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, dato4)
End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, nombre)
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


Private Sub dato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 And dato1.text <> "" Then
     Call Pregunta(dato1, dato2)
        dato1.text = ceros(dato1)
        lblDV.Caption = rut(dato1.text)
        rutCliente = dato1.text & lblDV.Caption
    End If
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            dato10.SetFocus
        End If
    
    
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
            dato10.text = ceros(dato10)
            
            dato3.SetFocus
        End If
        
    
    
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        dato3.text = ceros(dato3)
        If leebanco(dato3.text) <> "" Then
        lblbanco.Caption = leebanco(dato3.text)
        Call LEERCHEQUES(dato2.text, dato3.text, "n")
        nombre.text = leernombrecuenta(dato2.text, dato3.text)
        If lblprotestos.Caption = "" Then
            Call Pregunta(dato3, dato4)
        End If
        Else
        dato3.SetFocus
        
    
     End If
    
    End If
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then
        dato4.text = ceros(dato4)
    If leeplaza(dato4.text) <> "" Then
        lblplaza.Caption = leeplaza(dato4.text)
        
        Call Pregunta(dato4, nombre)
        Else
        dato4.SetFocus
        
    End If
    
    End If
    
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato5.text <> "" And dato5.text <> "0" Then
        dato5.text = Format(CDbl(dato5.text), "###,###,###")
        
  
       
             Call Pregunta(dato5, dato6)
  
    
    
    End If
    
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then
    dato6.text = ceros(dato6)
    If dato6.text > "00" And dato6.text < "32" Then
    Call Pregunta(dato6, dato7)
    Else
    dato6.SetFocus
    End If
    
    End If
    
End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    dato7.text = ceros(dato7)
    If dato7.text > "00" And dato7.text < "13" Then
    Call Pregunta(dato7, dato8)
    Else
    dato7.SetFocus
    End If
    End If
    
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    
    dato8.text = ceros(dato8)
    If dato8.text >= Format(fechasistema, "yyyy") And dato8.text < "2100" Then
    Call Pregunta(dato8, dato9)
    Else
    dato8.SetFocus
    End If
    
    
    End If
    
End Sub


Private Sub dato9_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If leernombrecuenta(dato2.text, dato3.text) = "" Then
            Call grabarcuenta(dato2.text, dato3.text, dato4.text, dato5.text, dato1.text + lblDV.Caption, dato9.text, cajera)
        End If
    
        If esvip(dato1.text & lblDV.Caption) = False Then
            
            grabacheque
            consultacheque = "S"
            If consultacheque = "S" Then
                horainicio = Format(Time, "HH:MM:SS")
                barra.Max = segundosespera
                barra.Value = 1
                Timer1.Enabled = True
            End If
        Else
            codigorecepcion.text = "99"
            codigorespuesta = "99"
            lblrespuesta.text = "APROBADO"
            lblcodigoautorizacion = "V.I.P."
             CHEQUEAPROBADO = True
        End If

    End If

End Sub
Sub temporizador()

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    imprimio = False
    Select Case KeyCode
        Case vbKeyF1
            Call dato9_KeyPress(13)
            
        Case vbKeyF2
            
        Case vbKeyF3
        Rem Call DATO9_KeyPress(13)
        frmmanual.Enabled = True
        Timer1.Enabled = False
        
        If codigorecepcion.text = "" Or codigorecepcion.text = "01" Then
        codigorecepcion.text = "02"
        GLOSARECEPCION.text = "AUTORIZADO ASEGURADORA"
        lblrespuesta.text = "AUTORIZACION MANUAL"
        lblcodigoautorizacion.Locked = False
        lblcodigoautorizacion.SetFocus
        Else
        MsgBox ("USTED NO PUEDE AUTORIZAR UN CHEQUE QUE EL SISTEMA RECHAZO")
        
        End If
        
        Case vbKeyF4
        Call dato9_KeyPress(13)
        frmmanual.Enabled = True
        Timer1.Enabled = False
        
        
        codigorecepcion.text = "01"
        GLOSARECEPCION.text = "AUTORIZADO SIN ASEGURADORA"
        lblrespuesta.text = "AUTORIZADO JEFA DE CAJA"
             permiso.Show vbModal
             If autorizador = True Then
             lblcodigoautorizacion.Enabled = True
             
             lblcodigoautorizacion.text = leerjefaautorizacion(claveautorizadora)
             lblcodigoautorizacion.Enabled = False
             
             Else
             lblcodigoautorizacion.Enabled = True
             
             lblcodigoautorizacion.text = ""
             lblcodigoautorizacion.Enabled = False
             
             End If
        
        Case vbKeyF5
        If IMPRESORABOLETA = "KC4" Then
        leecheque_Click
        End If
                      
        Case 27
            If lblcodigoautorizacion.text <> "" Then
            With detallePagos
'            .Pagos.Cell(.Pagos.Rows - 1, 3).text = dato5.text
'            .Pagos.Cell(.Pagos.Rows - 1, 4).text = dato10.text
'            .Pagos.Cell(.Pagos.Rows - 1, 5).text = dato3.text
'            .Pagos.Cell(.Pagos.Rows - 1, 6).text = dato2.text
'            .Pagos.Cell(.Pagos.Rows - 1, 7).text = dato6.text + "-" + dato7.text + "-" + dato8.text
'
            .pagos.Cell(.pagos.ActiveCell.row, 3).text = dato5.text
            .pagos.Cell(.pagos.ActiveCell.row, 4).text = dato10.text
            .pagos.Cell(.pagos.ActiveCell.row, 5).text = dato3.text
            .pagos.Cell(.pagos.ActiveCell.row, 6).text = dato2.text
            .pagos.Cell(.pagos.ActiveCell.row, 7).text = dato6.text + "-" + dato7.text + "-" + dato8.text
         
            
            
            Call modificaautorizacion(dato1.text + lblDV.Caption, dato2.text, dato3.text, dato4.text, CDbl(dato5.text), dato10.text, codigorecepcion.text, lblrespuesta.text, lblcodigoautorizacion.text)
            CHEQUEAPROBADO = True
            Unload Me
            End With
           Else
        
        CHEQUEAPROBADO = False
        Unload Me
           End If
        
    End Select
    
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub grabacheque()
        Dim condicion As String
        Dim campos(20, 3) As String
        Dim op As Integer
        Dim K As Integer
        Dim TIPO As String
        Dim NUMERO As String
        Dim cuota As String
        Dim abono As Double
        
        campos(0, 0) = "fecha"
        campos(1, 0) = "hora"
        campos(2, 0) = "local"
        campos(3, 0) = "caja"
        campos(4, 0) = "cajera"
        campos(5, 0) = "cuenta"
        campos(6, 0) = "banco"
        campos(7, 0) = "plaza"
        campos(8, 0) = "monto"
        campos(9, 0) = "vencimiento"
        campos(10, 0) = "rut"
        campos(11, 0) = "operador"
        campos(12, 0) = "fono"
        campos(13, 0) = "codigorespuesta"
        campos(14, 0) = "glosa"
        campos(15, 0) = "envioback"
        campos(16, 0) = "recibopos"
        campos(17, 0) = "numero"
        campos(18, 0) = "ip"
        campos(19, 0) = ""
        campos(0, 1) = Format(fechasistema, "yyyy-mm-dd")
        campos(1, 1) = Time
        campos(2, 1) = empresaActiva
        campos(3, 1) = caja
        campos(4, 1) = cajera
        campos(5, 1) = dato2.text
        campos(6, 1) = dato3.text
        campos(7, 1) = dato4.text
        campos(8, 1) = CDbl(dato5.text)
        campos(9, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
        campos(10, 1) = dato1.text + lblDV.Caption
        campos(11, 1) = "00"
        campos(12, 1) = dato9.text
        campos(13, 1) = ""
        campos(14, 1) = ""
        campos(15, 1) = "1"
        campos(16, 1) = "0"
        campos(17, 1) = dato10.text
        campos(18, 1) = servidor
        horainicio = Format(Time, "HH:MM:SS")
        campos(0, 2) = "sv_consultacheques"
        condicion = ""
        op = 2
        
        sqlventas.response = campos
        
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
       
     
        
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
lblrespuesta.text = ""
lblcodigoautorizacion.text = ""
frmmanual.Enabled = False
CARGAGRILLA
CHEQUEAPROBADO = False
End Sub

Private Sub lblcodigoautorizacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Form_KeyDown(27, 0)
End If

End Sub

Private Sub leecheque_Click()
'If chequeskc4 = True Then
Dim CORRE As Integer

leecheques.Visible = True
datoscheques.SetFocus




End Sub

Private Sub nombre_GotFocus()
  Call selecciona(nombre)
End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And nombre.text <> "" Then
    dato5.SetFocus
End If

End Sub




Private Sub Timer1_Timer()
Dim HH As Double
Dim MM As Double
Dim SS As Double
Dim HH1 As Double
Dim MM1 As Double
Dim SS1 As Double

If PASADA > 22 Then PASADA = 0: lblrespuesta.text = ""
    PASADA = PASADA + 1
    lblrespuesta.text = String(PASADA, 32) + "ESPERANDO CONFIRMACION"
    HH = Mid(horainicio, 1, 2) * 360
    MM = Mid(horainicio, 4, 2) * 60
    SS = Mid(horainicio, 7, 2) * 1
    HH1 = Mid(Format(Time, "HH:MM:SS"), 1, 2) * 360
    MM1 = Mid(Format(Time, "HH:MM:SS"), 4, 2) * 60
    SS1 = Mid(Format(Time, "HH:MM:SS"), 7, 2) * 1
    segundos = (HH1 + MM1 + SS1) - (HH + MM + SS)
    If segundos > segundosespera Then
        Timer1.Enabled = False
        lblrespuesta.text = "TIEMPO DE ESPERA AGOTADO REINTENTAR"
    End If
    If segundos < segundosespera Then
        barra.Value = segundos
    End If
    If leeconfirmacion(dato1.text + lblDV.Caption, dato2.text, dato3.text, dato4.text, CDbl(dato5.text), dato10.text) = True Then
        
        If codigorespuesta = "01" Then
            lblrespuesta.text = "ABROBADO"
            lblcodigoautorizacion = codigoautorizacion
            Timer1.Enabled = False
            barra.Value = 0
        End If
        If codigorespuesta <> "01" And codigorespuesta <> "80" Then
            lblrespuesta.text = "NO APROBADO"
            lblcodigoautorizacion = ""
            barra.Value = 0
            Timer1.Enabled = False
            detalle.Visible = True
            
        End If
        If codigorespuesta = "80" Then
            lblrespuesta.text = "CONSULTA NO DISPONIBLE, REINTENTE"
            lblcodigoautorizacion = ""
            barra.Value = 0
            Timer1.Enabled = False
        End If
        
    End If



End Sub
Public Function leeconfirmacion(rut, CUENTA, Banco, PLAZA, MONTO, NUMERO) As Boolean

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    

        Set csql.ActiveConnection = ventas
        csql.sql = "SELECT codigorespuesta,codigoautorizacion,codigorecepcion,glosarespuesta,nombrecliente,glosainstacheck "
        csql.sql = csql.sql + "FROM sv_consultacheques "
        csql.sql = csql.sql + "where rut='" + rut + "' and cuenta='" + CUENTA + "' and banco='" + Banco + "' and plaza='" + PLAZA + "' and monto='" & MONTO & "' and numero='" + NUMERO + "' "
        
        csql.Execute
        codigorespuesta = ""
        codigoautorizacion = ""
        
        leeconfirmacion = False
      
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            If resultados(0) <> "" Then
            codigorespuesta = resultados(0)
            codigoautorizacion = resultados(1)
          'GLOSARECEPCION.text = resultados(3)
            codigorecepcion.text = resultados(2)
            NOMBRECLIENTE.text = resultados(4)
            glosarespuesta.Caption = resultados(5)
            
            leeconfirmacion = True
        Else
            codigorespuesta = ""
            codigoautorizacion = ""
            GLOSARECEPCION.text = ""
            codigorecepcion.text = ""
            NOMBRECLIENTE.text = ""
            leeconfirmacion = False
        End If
                    
            resultados.Close
        Set resultados = Nothing
       
        End If
    
End Function
Sub modificaautorizacion(rut, CUENTA, Banco, PLAZA, MONTO, NUMERO, codigorecepcion, glosarespuesta, codigoautorizacion)

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    

        Set csql.ActiveConnection = ventas
        csql.sql = "update sv_consultacheques "
        csql.sql = csql.sql + "set codigorecepcion='" + codigorecepcion + "',glosarespuesta='" + glosarespuesta + "',codigoautorizacion='" + codigoautorizacion + "',recibopos='1' "
        csql.sql = csql.sql + "where rut='" + rut + "' and cuenta='" + CUENTA + "' and banco='" + Banco + "' and plaza='" + PLAZA + "' and monto='" & MONTO & "' and numero='" + NUMERO + "' "
        
        csql.Execute
       Call sincronizadatos(csql.sql, ventasRubro)
    
End Sub

Private Sub CARGAGRILLA()
        Dim i As Integer
        Dim formatogrilla(20, 20) As String
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "NUMERO"
        formatogrilla(1, 2) = "MONTO"
        formatogrilla(1, 3) = "VENCIMIENTO"
        
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "D"
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = "$ ###,###,##0"
        formatogrilla(4, 3) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        
        
        Rem VALOR MINIMO
        'formatoGrilla(6, 1) = ""
        'formatoGrilla(6, 2) = ""
        'formatoGrilla(6, 3) = ""
        'formatoGrilla(6, 4) = ""
        'formatoGrilla(6, 5) = ""
        'formatoGrilla(6, 6) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "10"
        formatogrilla(8, 3) = "10"
        
        Grid1.DefaultFont.Size = 10
        
        Grid1.DefaultFont.Bold = True
        Grid1.ExtendLastCol = True
        
        
        
        Grid1.Cols = 4
        Grid1.Rows = 1
        Grid1.Column(0).Width = 0
        For i = 1 To Grid1.Cols - 1
            Grid1.Cell(0, i).text = formatogrilla(1, i)
            Grid1.Column(i).Width = Val(formatogrilla(8, i)) * (Grid1.Cell(0, i).Font.Size + 1.25)
            Grid1.Column(i).MaxLength = Val(formatogrilla(2, i))
            Grid1.Column(i).FormatString = formatogrilla(4, i)
            Grid1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Grid1.Column(i).Alignment = cellRightCenter
            Else
                Grid1.Column(i).Alignment = cellLeftCenter
            End If
        
        Next i
        Grid1.Column(3).Alignment = cellCenterCenter
        
        
        
        
    End Sub

Sub LEERCHEQUES(CUENTA, Banco, todo)
    Dim total1 As Double
    Dim total2 As Double
    Dim Cheques As Double
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    

        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT numero,monto,vencimiento from " + baseteso + ".rc_cartera "
        csql.sql = csql.sql + "where cuenta='" + CUENTA + "' and banco='" + Banco + "'  order by vencimiento"
        csql.Execute
        total1 = 0
        total2 = 0
        Cheques = 0
        GLOSA.text = ""
        Grid1.Rows = 1
       If csql.RowsAffected > 0 Then
       Set resultados = csql.OpenResultset
       While resultados.EOF = False
       If todo = "n" Then
       If Format(resultados(2), "yyyy-mm-dd") >= Format(fechasistema, "yyyy-mm-dd") Then
       Grid1.Rows = Grid1.Rows + 1
       Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
       Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
       Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
       total1 = total1 + resultados(1)
       total2 = total2 + resultados(1)
        Cheques = Cheques + 1
       Else
       total2 = total2 + resultados(1)
       Cheques = Cheques + 1
       End If
       End If
        
       If todo = "s" Then
       Grid1.Rows = Grid1.Rows + 1
       Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
       Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
       Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
       If Format(resultados(2), "yyyy-mm-dd") >= Format(fechasistema, "yyyy-mm-dd") Then
       total1 = total1 + resultados(1)
       End If
       total2 = total2 + resultados(1)
       Cheques = Cheques + 1
       End If
        
        
       resultados.MoveNext
       
       Wend
    
    CUPOUTILIZADO.Caption = Format(total1, "###,###,##0")
    TOTALHISTORICO.Caption = Format(total2, "###,###,##0")
    CUPOAUTORIZADO.Caption = Format(leecupocuenta(dato2.text, Banco), "###,###,##0")
    CUPODISPONIBLE.Caption = Format(CDbl(CUPOAUTORIZADO.Caption) - CDbl(CUPOUTILIZADO.Caption), "###,###,##0")
    CHEQUESRECIBIDOS.Caption = Format(Cheques, "###,###,##0")
    If Cheques > 0 Then
    CHEQUEPROMEDIO.Caption = Format(CDbl(TOTALHISTORICO.Caption) / CDbl(CHEQUESRECIBIDOS.Caption), "###,###,##0")
    End If
    
    End If
    
    If csql.RowsAffected > 1 Then
    
    If CDbl(CUPOAUTORIZADO.Caption) < 0 Then
    GLOSA.text = "CLIENTE BLOQUEADO CONSULTAR TESORERIA "
      Call mensaje.mostrarMensaje("CLIENTE BLOQUEADO", "LA CUENTA SE ENCUENTRA BLOQUEADA COMUNICARSE CON TESORERIA", "")
    dato3.text = ""
    dato3.SetFocus
    End If
    End If
    
     lblprotestos.Caption = leerprotestos(CUENTA, Banco)
       
       
End Sub
'Public Function leerprotestos(CUENTA, Banco) As String
'    Dim resultados As rdoResultset
'    Dim cSql As New rdoQuery
'    Dim linea As Double
'    Dim suma As Double
'
'    leerprotestos = ""
'    Set cSql.ActiveConnection = ventas
'    cSql.sql = "SELECT sum(monto) "
'    cSql.sql = cSql.sql & "from " & baseteso & ".cheques_protestados "
'    cSql.sql = cSql.sql & "where cuenta='" & CUENTA & "' and banco='" & Banco & "' and recuperado='0' "
'    cSql.sql = cSql.sql + "group BY cuenta"
'    cSql.Execute
'
'    If cSql.RowsAffected > 0 Then
'        Set resultados = cSql.OpenResultset
'        leerprotestos = "  CLIENTE  CON  PROTESTOS  " & Format(resultados(0), "$ ###,###,###")
'        Call mensaje.mostrarMensaje("CLIENTE CON PROTESTOS", "NO PUEDE RECIBIR CHEQUES, COMUNICARSE CON TESORERIA", "")
'        resultados.Close
'        Set resultados = Nothing
'    Else
'        leerprotestos = ""
'    End If
'
'End Function

Public Function leerprotestos(CUENTA, Banco) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim linea As Double
    Dim suma As Double
    Dim totalprotestos As Double
    Dim totalpagados As Double
    
    leerprotestos = ""
    Set csql.ActiveConnection = gestion
    csql.sql = "select cp.monto,cp.folio "
    csql.sql = csql.sql & "from " & baseteso & ".cheques_protestados as cp "
    csql.sql = csql.sql & "Where "
    csql.sql = csql.sql & "cp.cuenta='" & CUENTA & "' and cp.banco='" & Banco & "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            If leerevento(resultados(1)) = "1" Then
                totalpagados = totalpagados + resultados(0)
            Else
                totalprotestos = totalprotestos + resultados(0)
            End If
            resultados.MoveNext
        Wend
        
        resultados.Close
        Set resultados = Nothing
    End If
    csql.Close
    Set csql = Nothing
    
    If totalprotestos > 0 Then
        leerprotestos = "  CLIENTE  CON  PROTESTOS  " & Format(totalprotestos, "$ ###,###,###")
        Call mensaje.mostrarMensaje("CLIENTE CON PROTESTOS", "NO PUEDE RECIBIR CHEQUES, COMUNICARSE CON TESORERIA", "")
    Else
        leerprotestos = ""
    End If
End Function

Private Function leerevento(FOLIO) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = gestion

csql.sql = "select me.cancelacheque from " & baseteso & ".cheques_protestados_detalle as cp inner join " & baseteso & ".maestroeventos as me on "
csql.sql = csql.sql & "cp.evento=me.codigo where cp.folio='" & FOLIO & "' and me.cancelacheque='1' "
csql.Execute
leerevento = "0"
If csql.RowsAffected > 0 Then
leerevento = "1"
Else
leerevento = "0"
End If

End Function
