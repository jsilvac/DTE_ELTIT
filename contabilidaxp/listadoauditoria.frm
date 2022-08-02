VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ingreso01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso Cajas Diarias"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   14445
   Begin XPFrame.FrameXp OPCIONES 
      Height          =   495
      Left            =   9360
      TabIndex        =   30
      Top             =   4455
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      BackColor       =   49344
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   49344
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
      Begin VB.CommandButton Command4 
         Caption         =   "&Retorno"
         Height          =   255
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Imprimir"
         Height          =   255
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton cmdmodifica 
         Caption         =   "&Modificar"
         Height          =   255
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton cmdelimina 
         Caption         =   "&Eliminar"
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1230
      End
   End
   Begin XPFrame.FrameXp PANTALLA 
      Height          =   9135
      Left            =   90
      TabIndex        =   12
      Top             =   765
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   16113
      BackColor       =   16744576
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
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   3885
         Left            =   8595
         TabIndex        =   21
         Top             =   315
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   6853
         BackColor       =   16773879
         Caption         =   "RESUMEN DE CAJA"
         CaptionEstilo3D =   1
         BackColor       =   16773879
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox TOTALPAGOS 
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
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   44
            Tag             =   "fecha"
            Top             =   1080
            Width           =   2505
         End
         Begin VB.TextBox TOTALDONACIONES 
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
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   43
            Tag             =   "fecha"
            Top             =   720
            Width           =   2505
         End
         Begin VB.TextBox TOTALVENTAS 
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
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   42
            Tag             =   "fecha"
            Top             =   360
            Width           =   2505
         End
         Begin XPFrame.FrameXp FrameXp1 
            Height          =   855
            Left            =   60
            TabIndex        =   37
            Top             =   2430
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   1508
            BackColor       =   16773879
            Caption         =   "Glosa Diferencia"
            CaptionEstilo3D =   1
            BackColor       =   16773879
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox glosa 
               BackColor       =   &H00C0FFFF&
               Height          =   525
               Left            =   90
               MultiLine       =   -1  'True
               TabIndex        =   38
               Top             =   270
               Width           =   5205
            End
         End
         Begin XPFrame.FrameXp GRABAR 
            Height          =   495
            Left            =   45
            TabIndex        =   34
            Top             =   3330
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   873
            BackColor       =   49344
            Caption         =   ""
            CaptionEstilo3D =   1
            BackColor       =   49344
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
            Begin VB.CommandButton cmdGrabar 
               Caption         =   "&Grabar"
               Height          =   255
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   225
               Width           =   1410
            End
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PAGOS RECIBIDOS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   225
            TabIndex        =   41
            Top             =   1080
            Width           =   2490
         End
         Begin VB.Label TOTALDIFERENCIACAJA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   2880
            TabIndex        =   29
            Top             =   2115
            Width           =   2490
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DIFERENCIA CAJA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   240
            TabIndex        =   28
            Top             =   2115
            Width           =   2490
         End
         Begin VB.Label TOTALRENDIDO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   2880
            TabIndex        =   27
            Top             =   1755
            Width           =   2490
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TOTAL RENDIDO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   240
            TabIndex        =   26
            Top             =   1800
            Width           =   2490
         End
         Begin VB.Label TOTALARENDIR 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   2880
            TabIndex        =   25
            Top             =   1440
            Width           =   2490
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TOTAL RENDIR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   240
            TabIndex        =   24
            Top             =   1440
            Width           =   2490
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DONACIONES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   2490
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VENTAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   2490
         End
      End
      Begin XPFrame.FrameXp FRMOTROS 
         Height          =   3930
         Left            =   4320
         TabIndex        =   18
         Top             =   315
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   6932
         BackColor       =   16761024
         Caption         =   "OTROS VALORES"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Otros 
            Height          =   3270
            Left            =   45
            TabIndex        =   19
            Top             =   390
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   5768
            Appearance      =   0
            BackColor1      =   16761024
            BackColor2      =   16761024
            Cols            =   3
            DefaultFontSize =   8.25
            Rows            =   12
            DateFormat      =   2
         End
      End
      Begin XPFrame.FrameXp FRMEFECTIVO 
         Height          =   3930
         Left            =   180
         TabIndex        =   16
         Top             =   315
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   6932
         BackColor       =   16761024
         Caption         =   "EFECTIVO A RENDIR"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Dineros 
            Height          =   3120
            Left            =   135
            TabIndex        =   17
            Top             =   315
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   5503
            Cols            =   3
            DefaultFontSize =   8.25
            Rows            =   12
            DateFormat      =   2
         End
         Begin VB.Label EFECTIVO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   1575
            TabIndex        =   20
            Tag             =   "$ 0"
            Top             =   3510
            Width           =   2295
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   4515
         Left            =   180
         TabIndex        =   14
         Top             =   4410
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   7964
         BackColor       =   16773879
         Caption         =   "Resumen de Cajas Digitadas"
         CaptionEstilo3D =   1
         BackColor       =   16773879
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin XPFrame.FrameXp FrameXp2 
            Height          =   645
            Left            =   9045
            TabIndex        =   46
            Top             =   3825
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   1138
            BackColor       =   16761024
            Caption         =   "CORTE CHEQUE A FECHA"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.CheckBox CIERRE 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Cierre Cheques"
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
               Left            =   2790
               TabIndex        =   51
               Top             =   315
               Width           =   2040
            End
            Begin VB.TextBox DIA1 
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
               Left            =   765
               MaxLength       =   2
               TabIndex        =   49
               Tag             =   "fecha"
               Top             =   315
               Width           =   375
            End
            Begin VB.TextBox MES1 
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
               Left            =   1170
               MaxLength       =   2
               TabIndex        =   48
               Tag             =   "fecha"
               Top             =   315
               Width           =   375
            End
            Begin VB.TextBox AÑO1 
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
               Left            =   1575
               MaxLength       =   4
               TabIndex        =   47
               Tag             =   "fecha"
               Top             =   315
               Width           =   615
            End
            Begin VB.Label Label3 
               BackColor       =   &H00F5C9B1&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Fecha"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   90
               TabIndex        =   50
               Top             =   315
               Width           =   615
            End
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FF8080&
            Caption         =   "Imprimir"
            Height          =   330
            Left            =   6435
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   3960
            Width           =   1725
         End
         Begin FlexCell.Grid Ingresadas 
            Height          =   3495
            Left            =   90
            TabIndex        =   15
            Top             =   315
            Width           =   13860
            _ExtentX        =   24448
            _ExtentY        =   6165
            Cols            =   3
            DefaultFontSize =   8.25
            Rows            =   7
            DateFormat      =   2
         End
         Begin XPFrame.FrameXp FrameXp3 
            Height          =   600
            Left            =   135
            TabIndex        =   52
            Top             =   3825
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   1058
            BackColor       =   12648384
            Caption         =   "LOCALES"
            CaptionEstilo3D =   1
            BackColor       =   12648384
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
            Begin VB.OptionButton local3 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ambos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3285
               TabIndex        =   55
               Top             =   315
               Width           =   1860
            End
            Begin VB.OptionButton local2 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Villarrica"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1305
               TabIndex        =   54
               Top             =   315
               Width           =   1860
            End
            Begin VB.OptionButton local1 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Pucon"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   53
               Top             =   315
               Value           =   -1  'True
               Width           =   1860
            End
         End
      End
      Begin MSAdodcLib.Adodc Ingresados 
         Height          =   330
         Left            =   0
         Top             =   0
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
      Begin VB.Label lblVCompra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FBEDE6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   1980
         TabIndex        =   13
         Tag             =   "$ 0"
         Top             =   7785
         Width           =   2295
      End
   End
   Begin XPFrame.FrameXp CABEZA 
      Height          =   735
      Left            =   90
      TabIndex        =   2
      Top             =   -45
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   1296
      BackColor       =   16773879
      Caption         =   "Datos"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato5 
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
         Left            =   8415
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "locali"
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox dato4 
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
         Left            =   3150
         MaxLength       =   9
         TabIndex        =   6
         Tag             =   "caja"
         Top             =   315
         Width           =   1455
      End
      Begin VB.TextBox dato3 
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
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   315
         Width           =   615
      End
      Begin VB.TextBox dato2 
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
         Left            =   1215
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   315
         Width           =   375
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
         Left            =   810
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   315
         Width           =   375
      End
      Begin VB.Label lbllocal 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8820
         TabIndex        =   40
         Top             =   315
         Width           =   2130
         WordWrap        =   -1  'True
      End
      Begin VB.Label dv 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4680
         TabIndex        =   39
         Top             =   315
         Width           =   285
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2385
         TabIndex        =   11
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblcajera 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Top             =   315
         Width           =   2550
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LOCAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7695
         TabIndex        =   9
         Top             =   315
         Width           =   675
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc p 
      Height          =   330
      Left            =   -45
      Top             =   0
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
   Begin FlexCell.Grid impresion 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Cols            =   1
      DefaultFontSize =   8.25
      Rows            =   1
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   0
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "ingreso01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private totalretiro As Double
    Public totalnc As Double
    Public chequesdia As Double
    Public chequesfecha As Double
    
    Private formatogrilla(10, 45) As String
    Private largoCeros(2, 10) As Long
    Private campos(30, 5) As String
    Private fila As Long
    Private columna As Long
    Private vacio As Boolean
    Private sumaDineros As Double
    Private sumacheques As Double
    Private sumaDebitos As Double
    Private sumacreditos As Double
    Private sumaExtranjeras As Double
    Private sumaOtorgados As Double
    Private sumaPagados As Double
    Private sumaVCompra As Double
    Private sumaOCredito As Double
    Private sumaNCredito As Double
    Private sumaChDevueltos As Double
    Private sumaVarios As Double
    Private sumaIngresadas As Double
    Private sumaTotal As Double
    Private modificar As Boolean
    Private TVENTAS As Double
    Private TDONACIONES As Double
    Private TPAGOS As Double
    Private TARENDIR As Double
    Private TRENDIDO As Double
    Private TDIFERENCIAS As Double
    Private loc As String

Private Sub cmdRescatar_Click()
    Call leerVentasDia
End Sub

Private Sub CHEQUES_AfterReorderColumn(ByVal OriginalPosition As Long, ByVal NewPosition As Long)

End Sub

Private Sub Cmdelimina_Click()
If MsgBox("ESTA SEGURO DE ELIMINAR", vbYesNo, "ELIMINAR CAJA") = vbYes Then

elimina
retorno
End If
End Sub

Private Sub cmdmodifica_click()
                
                modificar = True
                
                elimina
                
                OPCIONES.Visible = False
                GRABAR.Visible = True
                FRMEFECTIVO.Enabled = True
                FRMOTROS.Enabled = True
                CABEZA.Enabled = True
                Dineros.Cell(1, 1).SetFocus
                
End Sub

Sub cabezas2(titulo, titulo2)
Dim objReportTitle As FlexCell.ReportTitle
Ingresadas.ReportTitles.Clear

DATOSEMPRESA(1) = "EMPRESAS ELTIT "

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Ingresadas.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo2
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Ingresadas.ReportTitles.Add objReportTitle
    
    'Report Title 1
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = cellLeft
        Ingresadas.ReportTitles.Add objReportTitle
    Next k
    
    Ingresadas.Range(0, 0, 0, Ingresadas.Cols - 1).Borders(cellEdgeTop) = cellThin
    Ingresadas.Range(0, 0, 0, Ingresadas.Cols - 1).Borders(cellEdgeBottom) = cellThin
    Ingresadas.Range(0, 0, 0, Ingresadas.Cols - 1).Borders(cellEdgeLeft) = cellThin
    Ingresadas.Range(0, 0, 0, Ingresadas.Cols - 1).Borders(cellEdgeRight) = cellThin
    Ingresadas.Range(0, 0, 0, Ingresadas.Cols - 1).Borders(cellInsideHorizontal) = cellThin
    Ingresadas.Range(0, 0, 0, Ingresadas.Cols - 1).Borders(cellInsideVertical) = cellThin
    
    
    
    
    
    
With Ingresadas.PageSetup
        
        .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        .LeftMargin = 0.5
        .RightMargin = 0.5
        .BlackAndWhite = True
        .PrintFixedRow = True
        
        
        
        
        
        
        
End With

End Sub



Sub imprimerendicion()
Dim i As Integer
Call cabezas2("LISTADO RESUMEN DE CAJAS ", "DEL DIA " + dato1.text + "-" + dato2.text + "-" + dato3.text)

Ingresadas.PageSetup.Orientation = cellLandscape
Ingresadas.PageSetup.LeftMargin = 2.5
Ingresadas.PageSetup.RightMargin = 0.5
Ingresadas.PageSetup.TopMargin = 0.5
Ingresadas.PageSetup.BottomMargin = 0.5
Ingresadas.PageSetup.PrintFixedRow = True






        For i = 1 To Ingresadas.PageSetup.PaperSizes.Count
            If UCase(Ingresadas.PageSetup.PaperSizes.Item(i).PaperName) = "OFICIO" Then
                Ingresadas.PageSetup.PaperSize = Ingresadas.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        





Ingresadas.Column(21).Width = 50

Ingresadas.PrintPreview
Ingresadas.Column(21).Width = 500












End Sub

Private Sub Command1_Click()
Call CargaGrillaIngresadas(1, 22)
cargaIngresadas

imprimerendicion

End Sub

Private Sub Command4_Click()
retorno

End Sub

'****************************************************************************
'Manejo de los Controles
'****************************************************************************
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    Private Sub DATO1_GotFocus()
        Call cargatexto(dato1)
    End Sub

    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        fechacajas = dato3.text + "-" + dato2.text + "-" + dato1.text
        DIA1.text = dato1.text
        MES1.text = dato2.text
        AÑO1.text = dato3.text
        
        Call cargatexto(dato4)
    End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaCajera(dato4)
    Call Flechas(dato3, dato5, KeyCode)
End Sub
Private Sub DATO1_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Flechas(dato2, dato4, KeyCode)
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato5)
            
            If leerLocal(dato5.text) <> "" Then
                lbllocal.Caption = leerLocal(dato5.text)
                If leerCajaDiaria = False Then
                TVENTAS = LeerTotalVenta
                Call sumaGrilla
                Dineros.Cell(1, 1).SetFocus
                Call leeradmin(dato3.text + "-" + dato2.text + "-" + dato1.text, dato5.text, dato4.text)
                Else
                    OPCIONES.Visible = True
                    GRABAR.Visible = False
                    FRMEFECTIVO.Enabled = False
                    FRMOTROS.Enabled = False
                     CABEZA.Enabled = False
                
                End If

                
                
                Else
               
                
                
                End If
        End If
        
    End Sub
        
  


Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Flechas(dato5, dato5, KeyCode)
End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudalocales(dato5)
    Call Flechas(dato3, dato5, KeyCode)
End Sub
    
    Private Sub dato5_GotFocus()
        Call cargatexto(dato5)
    End Sub
    
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    
    '****************************************************************************
    'KEYPRESS
    '****************************************************************************
    Private Sub DATO1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            
            Call cargaIngresadas
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato4)
            dv.Caption = rut(dato4)
            If leerCajera(dato4.text + dv.Caption) <> "" Then
                lblcajera.Caption = leerCajera(dato4.text + dv.Caption)
                
                dato5.SetFocus
            Else
               MsgBox ("Cajera No existe")
               
                dato4.SetFocus
            
            End If
        End If
    End Sub
    
    


    '****************************************************************************
    'KEYPRESS
    '****************************************************************************
    
Private Sub Dineros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Dineros.ActiveCell.row = 7 And Dineros.ActiveCell.col = 1 Then Otros.Cell(0, 1).SetFocus


End Sub

Private Sub DONACIONES_Click()

End Sub

Private Sub Dineros_LostFocus()
 Call sumaGrilla
End Sub

    Private Sub Form_Load()
    Dim s As String
    creandorut = False
    Centrar Me
    GRABAR.Visible = False
      modificar = False
        Call Centrar(Me)
        Call CargaGrillaDineros(8, 2)
        Call CargaGrillaOtros(11, 2)
        Call CargaGrillaIngresadas(1, 22)
        dato1.text = Format(fechasistema, "dd")
        dato2.text = Format(fechasistema, "mm")
        dato3.text = Format(fechasistema, "yyyy")
        DIA1.text = dato1.text
        MES1.text = dato2.text
        AÑO1.text = dato3.text
        
        Call cargaIngresadas
        FRMEFECTIVO.Enabled = False
        FRMOTROS.Enabled = False
    Call Conectar_BD
        
        
    End Sub
'****************************************************************************
'Manejo de los Controles
'****************************************************************************


'****************************************************************************
'Formato de la Grilla Dineros
'****************************************************************************
    Private Sub CargaGrillaDineros(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 0) = "BILL / MON"
        formatogrilla(1, 1) = "$ retiros "
        formatogrilla(1, 2) = "$ 20.000"
        formatogrilla(1, 3) = "$ 10.000"
        formatogrilla(1, 4) = "$ 5.000"
        formatogrilla(1, 5) = "$ 2.000"
        formatogrilla(1, 6) = "$ 1.000"
        formatogrilla(1, 7) = "$ sencillo"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "20"
        formatogrilla(2, 2) = "9"
        
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = "$ ###,###,##0"
        formatogrilla(4, 3) = "###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "7"
        formatogrilla(8, 2) = "12"
        formatogrilla(8, 3) = "5"
                
        Dineros.Cols = col
        Dineros.Rows = row
        Dineros.AllowUserResizing = False
        Dineros.DisplayFocusRect = False
        Dineros.ExtendLastCol = True
        Dineros.BoldFixedCell = False
        Dineros.DrawMode = cellOwnerDraw
        Dineros.Appearance = Flat
        Dineros.ScrollBarStyle = Flat
        Dineros.FixedRowColStyle = Flat
        Dineros.BackColorFixed = RGB(90, 158, 214)
        Dineros.BackColorFixedSel = RGB(110, 180, 230)
        Dineros.BackColorBkg = RGB(90, 158, 214)
        Dineros.BackColorScrollBar = RGB(231, 235, 247)
        Dineros.BackColor1 = RGB(231, 235, 247)
        Dineros.BackColor2 = RGB(239, 243, 255)
        Dineros.GridColor = RGB(148, 190, 231)
        
        'Dineros.RowHeight(0) = 0
        Dineros.Column(0).Width = Val(formatogrilla(8, 1)) * (Dineros.Cell(0, 0).Font.Size + 1.25)
        Dineros.Column(1).Width = Val(formatogrilla(8, 2)) * (Dineros.Cell(0, 1).Font.Size + 1.25)
        
        Dineros.Column(0).MaxLength = Val(formatogrilla(2, 1))
        Dineros.Column(1).MaxLength = Val(formatogrilla(2, 2))
        
        Dineros.Column(0).FormatString = formatogrilla(4, 1)
        Dineros.Column(1).FormatString = formatogrilla(4, 2)
        
        Dineros.Column(0).Locked = formatogrilla(5, 1)
        Dineros.Column(1).Locked = formatogrilla(5, 2)
        
        
        Dineros.Column(1).Mask = cellNumeric
        
        
        If formatogrilla(3, 1) = "N" Then
            Dineros.Column(0).Alignment = cellRightCenter
        Else
            Dineros.Column(0).Alignment = cellLeftCenter
        End If
        If formatogrilla(3, 2) = "N" Then
            Dineros.Column(1).Alignment = cellRightCenter
        Else
            Dineros.Column(1).Alignment = cellLeftCenter
        End If
        
        For i = 0 To row - 1
            Dineros.Cell(i, 0).text = formatogrilla(1, i)
        Next i
        Dineros.Cell(0, 1).text = "MONTO"

        Dineros.Range(0, 0, 0, col - 1).Alignment = cellCenterCenter
        Dineros.Range(0, 0, row - 1, 0).Alignment = cellCenterCenter
        
        Dineros.Enabled = True
    End Sub
    
    
    Private Sub Dineros_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        Dim i As Integer
        fila = Dineros.ActiveCell.row
        columna = Dineros.ActiveCell.col
        If KeyCode = 13 And Dineros.ActiveCell.row = 7 Then KeyCode = 187
        Select Case KeyCode
                
                
                
                
                Case 187, 106
               
                Otros.Cell(1, 1).SetFocus
        
        End Select
    End Sub
    
    Private Sub Dineros_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
        If vacio = True Then
            NewRow = fila
            NewCol = columna
        Else
            If row < NewRow Then
                If col = 1 And row <> 1 Then
                    
                    If Val(Dineros.Cell(row, col).text) Mod CDbl(Dineros.Cell(row, 0).text) = 0 Then
                        Rem Dineros.Cell(row, newcol).text = CDbl(Dineros.Cell(row, col).text) / CDbl(Dineros.Cell(row, 0).text)
                    Else
                        MsgBox "Ingrese una cantidad correcta", vbOKOnly, "ERROR"
                        NewCol = 1
                        NewRow = row
                    End If
                End If
                
            Else
'                If col = 1 And CDbl(Dineros.Cell(row, 1).text) <> 0 Then
'
'                    If CDbl(Dineros.Cell(row, col).text) Mod CDbl(Dineros.Cell(row, 0).text) = 0 Then
'
'                    Else
'                        MsgBox "Ingrese una cantidad correcta", vbOKOnly, "ERROR"
'                        newcol = col
'                        newrow = row
'                    End If
'                End If
'
            End If
            Call sumaGrilla
        End If
    End Sub
    
  
'****************************************************************************
'Formato de la Grilla Dineros
'****************************************************************************


'****************************************************************************
'Formato de la Grilla OTROS
'****************************************************************************
    Private Sub CargaGrillaOtros(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 0) = "TIPO DOC."
        formatogrilla(1, 1) = "CHEQUES"
        formatogrilla(1, 2) = "T.CREDITO"
        formatogrilla(1, 3) = "T.DEBITO"
        formatogrilla(1, 4) = "MONEDA EXT."
        formatogrilla(1, 5) = "CREDITO TMP"
        formatogrilla(1, 6) = "V.COMPRA"
        formatogrilla(1, 7) = "O.CREDITO"
        formatogrilla(1, 8) = "N.CREDITO"
        formatogrilla(1, 9) = "T.CASAS.COM"
        
        formatogrilla(1, 10) = "VARIOS"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "15"
        formatogrilla(2, 2) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "15"
                
        Otros.Cols = col
        Otros.Rows = row
        Otros.AllowUserResizing = False
        Otros.DisplayFocusRect = False
        Otros.ExtendLastCol = True
        Otros.BoldFixedCell = False
        Otros.DrawMode = cellOwnerDraw
        Otros.Appearance = Flat
        Otros.ScrollBarStyle = Flat
        Otros.FixedRowColStyle = Flat
        Otros.BackColorFixed = RGB(90, 158, 214)
        Otros.BackColorFixedSel = RGB(110, 180, 230)
        Otros.BackColorBkg = RGB(90, 158, 214)
        Otros.BackColorScrollBar = RGB(231, 235, 247)
        Otros.BackColor1 = RGB(231, 235, 247)
        Otros.BackColor2 = RGB(239, 243, 255)
        Otros.GridColor = RGB(148, 190, 231)
        
        'otros.RowHeight(0) = 0
        Otros.Column(0).Width = Val(formatogrilla(8, 1)) * (Otros.Cell(0, 0).Font.Size + 1.25)
        Otros.Column(1).Width = Val(formatogrilla(8, 2)) * (Otros.Cell(0, 1).Font.Size + 1.25)
        
        Otros.Column(0).MaxLength = Val(formatogrilla(2, 1))
        Otros.Column(1).MaxLength = Val(formatogrilla(2, 2))
        
        Otros.Column(0).FormatString = formatogrilla(4, 1)
        Otros.Column(1).FormatString = formatogrilla(4, 2)
        
        Otros.Column(0).Locked = True
        Otros.Column(1).Locked = False
        
        Otros.Column(1).Mask = cellNumeric
        
        If formatogrilla(3, 1) = "N" Then
            Otros.Column(0).Alignment = cellRightCenter
        Else
            Otros.Column(0).Alignment = cellLeftCenter
        End If
        If formatogrilla(3, 2) = "N" Then
            Otros.Column(1).Alignment = cellRightCenter
        Else
            Otros.Column(1).Alignment = cellLeftCenter
        End If
        
        For i = 0 To row - 1
            Otros.Cell(i, 0).text = formatogrilla(1, i)
        Next i
        Otros.Cell(0, 1).text = "MONTO"
        Otros.Range(0, 0, 0, col - 1).Alignment = cellCenterCenter
        'Otros.Range(0, 0, row - 1, 0).Alignment = cellCenterCenter
        
        Otros.Enabled = True
    End Sub
    
Private Sub ingresos_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
If modificar = True Then MsgBox ("imposible salir en estado de modificacion grabe antes de salor "): Cancel = 1: cmdGrabar.SetFocus

End Sub

Private Sub GLOSA_GotFocus()
For k = 1 To 10
If Otros.Cell(k, 1).text = "" Then
Otros.Cell(k, 1).text = "0"
End If


Next k

End Sub

Private Sub glosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
GRABAR.Visible = True: cmdGrabar.SetFocus
End If

End Sub

Private Sub lblCheques_Click()

End Sub

Private Sub Ingresadas_AfterReorderColumn(ByVal OriginalPosition As Long, ByVal NewPosition As Long)

End Sub

Private Sub local1_Click()
Call cargaIngresadas
End Sub

Private Sub local2_Click()
Call cargaIngresadas
End Sub

Private Sub local3_Click()
Call cargaIngresadas
End Sub

    Private Sub Otros_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        fila = Otros.ActiveCell.row
        columna = Otros.ActiveCell.col
     If KeyCode = 13 And Otros.ActiveCell.row = 10 Then KeyCode = 200
        Select Case KeyCode
            Case 13, 37, 38, 39, 40
                If Otros.ActiveCell.text = "" Then
                    Otros.ActiveCell.text = "0"
                End If
                vacio = False
            Case 187, 106
                Select Case fila
                    Case 1
                        ingresach.Show vbModal
                    Case 2
                        ingresatc.Show vbModal
                    Case 3
                        ingresatd.Show vbModal
                    Case 4
                        ingresame.Show vbModal
                    Case 5
                        ingresacre.Show vbModal
                    Case 6
                        ingresavc.Show vbModal
                    Case 7
                        ingresaoc.Show vbModal
                    Case 8
                        ingresanc.Show vbModal
                    Case 9
                        ingresacc.Show vbModal
                    Case 10
                        ingresaconta.Show vbModal
                                
                        
                End Select
        Case 200
                        glosa.SetFocus
                        
        End Select
    End Sub
    
    Private Sub Otros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Otros.ActiveCell.row = 11 Then GRABAR.Visible = True: cmdGrabar.SetFocus
    
    
        KeyAscii = 0
    
    End Sub
    
    Private Sub Otros_LostFocus()
        Call sumaGrilla
    End Sub
    
    Private Sub CargaGrillaIngresadas(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
                
        formatogrilla(1, 1) = "LO": formatogrilla(8, 1) = "2"
        formatogrilla(1, 2) = "CAJERA": formatogrilla(8, 2) = "10"
        formatogrilla(1, 3) = "EFECTIVO": formatogrilla(8, 3) = "8"
        formatogrilla(1, 4) = "CHE.DIAS": formatogrilla(8, 4) = "8"
        formatogrilla(1, 5) = "CHE.FEC ": formatogrilla(8, 5) = "8"
        formatogrilla(1, 6) = "T.CREDITO": formatogrilla(8, 6) = "6"
        formatogrilla(1, 7) = "T.DEBITO": formatogrilla(8, 7) = "6"
        formatogrilla(1, 8) = "M.EXT": formatogrilla(8, 8) = "6"
        formatogrilla(1, 9) = "CREDITOS": formatogrilla(8, 9) = "6"
        formatogrilla(1, 10) = "V.COMPRA": formatogrilla(8, 10) = "6"
        formatogrilla(1, 11) = "O.CRED.": formatogrilla(8, 11) = "6"
        formatogrilla(1, 12) = "N.CRED.": formatogrilla(8, 12) = "6"
        formatogrilla(1, 13) = "VARIOS": formatogrilla(8, 13) = "6"
        formatogrilla(1, 14) = "T.RENDIDO": formatogrilla(8, 14) = "8"
        formatogrilla(1, 15) = "T. VENTA": formatogrilla(8, 15) = "8"
        formatogrilla(1, 16) = "DONAC.": formatogrilla(8, 16) = "0"
        formatogrilla(1, 17) = "T.PAGOS": formatogrilla(8, 17) = "6"
        formatogrilla(1, 18) = "T.A RENDIR": formatogrilla(8, 18) = "8"
        formatogrilla(1, 19) = "T.RENDIDO": formatogrilla(8, 19) = "0"
        formatogrilla(1, 20) = "D.CAJA": formatogrilla(8, 20) = "6"
        formatogrilla(1, 21) = "GLOSA": formatogrilla(8, 21) = "50"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "15"
        formatogrilla(2, 2) = "20"
        For k = 3 To 22
        formatogrilla(2, k) = "8"
        Next k
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "S"
        For k = 3 To 20
        formatogrilla(3, k) = "N"
        Next k
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        For k = 3 To 21
        formatogrilla(4, k) = "###,###,##0"
        Next k
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        For k = 3 To 22
        formatogrilla(5, k) = "TRUE"
        Next k
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        
        
        Ingresadas.Cols = col
        Ingresadas.Rows = row
        Ingresadas.AllowUserResizing = False
        Ingresadas.DisplayFocusRect = False
        Ingresadas.ExtendLastCol = True
        Ingresadas.BoldFixedCell = False
        Ingresadas.DrawMode = cellOwnerDraw
        Ingresadas.Appearance = Flat
        Ingresadas.ScrollBarStyle = Flat
        Ingresadas.FixedRowColStyle = Flat
        Ingresadas.BackColorFixed = RGB(90, 158, 214)
        Ingresadas.BackColorFixedSel = RGB(110, 180, 230)
        Ingresadas.BackColorBkg = RGB(90, 158, 214)
        Ingresadas.BackColorScrollBar = RGB(231, 235, 247)
        Ingresadas.BackColor1 = RGB(231, 235, 247)
        Ingresadas.BackColor2 = RGB(239, 243, 255)
        Ingresadas.GridColor = RGB(148, 190, 231)
        Ingresadas.Column(0).Width = 0
        
        For i = 1 To col - 1
            Ingresadas.Cell(0, i).text = formatogrilla(1, i)
            Ingresadas.Column(i).Width = Val(formatogrilla(8, i)) * (Ingresadas.Cell(0, i).Font.Size + 1.25)
            Ingresadas.Column(i).MaxLength = Val(formatogrilla(2, i))
            Ingresadas.Column(i).FormatString = formatogrilla(4, i)
            Ingresadas.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Ingresadas.Column(i).Alignment = cellRightCenter
                Ingresadas.Column(i).Mask = cellNumeric
            Else
                Ingresadas.Column(i).Alignment = cellLeftCenter
                Ingresadas.Column(i).Mask = cellUpper
            End If
        Next i
        Ingresadas.Range(0, 0, 0, col - 1).Alignment = cellCenterCenter
        
        Ingresadas.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Ingresadas
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Impresion
'****************************************************************************
    Private Sub CargaGrillaImpresionDetalle(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem ANCHO
        formatogrilla(1, 1) = "7"   '$
        formatogrilla(1, 2) = "7"   '20000
        formatogrilla(1, 3) = "7"   '10000
        formatogrilla(1, 4) = "6"   '5000
        formatogrilla(1, 5) = "6"   '2000
        formatogrilla(1, 6) = "6"   '1000
        formatogrilla(1, 7) = "5"   '500
        formatogrilla(1, 8) = "5"   '100
        formatogrilla(1, 9) = "4"   '50
        formatogrilla(1, 10) = "4"  '10
        formatogrilla(1, 11) = "3"  '5
        formatogrilla(1, 12) = "3"  '1
        
                
        impresion.Cols = col
        impresion.Rows = row
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        impresion.Column(0).Width = 0
        
        For i = 1 To col - 1
            impresion.Column(i).Width = Val(formatogrilla(1, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
        Next i
    End Sub
    

Private Sub sumaGrilla()
    Dim i As Integer
    Dim suma As Double
    
        
            suma = 0
            For i = 1 To Dineros.Rows - 1
                    suma = suma + Val(Dineros.Cell(i, 1).text)
                           Next i
            sumaDineros = suma
            EFECTIVO.Caption = Format(suma, "$ ###,###,##0")
            suma = 0
            For i = 1 To Otros.Rows - 1
                    suma = suma + Val(Otros.Cell(i, 1).text)
                
            Next i
            
    TOTALRENDIDO.Caption = Format(CDbl(EFECTIVO.Caption) + suma, "$ ###,###,##0")
       
    MUESTRATOTAL
    
End Sub
Sub MUESTRATOTAL()
    TARENDIR = TVENTAS + TDONACIONES + TPAGOS
    TDIFERENCIAS = CDbl(TOTALRENDIDO.Caption) - TARENDIR
    EFECTIVO.Caption = Format(sumaDineros, " ###,###,##0")
    TOTALVENTAS.text = Format(TVENTAS, " ###,###,##0")
    TOTALPAGOS.text = Format(TPAGOS, " ###,###,##0")
    TOTALDONACIONES.text = Format(TDONACIONES, " ###,###,##0")
    TOTALARENDIR.Caption = Format(TARENDIR, " ###,###,##0")
    TOTALDIFERENCIACAJA.Caption = Format(TDIFERENCIAS, " ###,###,##0")
End Sub

Private Sub cmdGrabar_Click()
    If dato1.text = "" Or dato2.text = "" Or dato3.text = "" Or dato4.text = "" Or dato5.text = "" Then
        MsgBox "No puede garabar la informacion sin llenar los datos de la cabeza", vbOKOnly, "ERROR"
    Else
            Call grabarcaja
      
    End If
    retorno
    
    End Sub



Private Function leerCliente(ByVal CODIGO As String) As String
    Dim condicion As String
    Dim op As Integer
    'Set sql = New CSQLUtil
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = "maestroclientes"

    condicion = "rut = '" & CODIGO & "'"
    op = 5
    SQLUTIL.datos = campos
'    Set SQLUTIL.conexion = gestion
'    Call SQLUTIL.SQLUTIL(op, condicion)
   If SQLUTIL.estado = 0 Then
        leerCliente = SQLUTIL.datos(0, 3)
    Else
        leerCliente = "NO EXISTE"
    End If
End Function

Private Function leerCajaDiaria() As Boolean
    Dim condicion As String
    Dim op As Integer
    'Set sql = New CSQLUtil
     campos(0, 0) = "fecha"
    campos(1, 0) = "codigocajera"
    campos(2, 0) = "local"
    campos(3, 0) = "caja"
    campos(4, 0) = "montoretiros"
    campos(5, 0) = "billeteveinte"
    campos(6, 0) = "billetediez"
    campos(7, 0) = "billetecinco"
    campos(8, 0) = "billetedos"
    campos(9, 0) = "billetemil"
    campos(10, 0) = "sencillo"
    campos(11, 0) = "montocheques"
    campos(12, 0) = "montotdebito"
    campos(13, 0) = "montotcredito"
    campos(14, 0) = "montoextranjera"
    campos(15, 0) = "montootorgado"
    campos(16, 0) = "montovcompra"
    campos(17, 0) = "montoocredito"
    campos(18, 0) = "montoncredito"
    campos(19, 0) = "montocasascomerciales"
    campos(20, 0) = "montovarios"
    campos(21, 0) = "totalventas"
    campos(22, 0) = "totaldonaciones"
    campos(23, 0) = "totalpagos"
    campos(24, 0) = "totalarendir"
    campos(25, 0) = "totalrendido"
    campos(26, 0) = "totaldiferenciadecaja"
    campos(27, 0) = "glosa"
    campos(28, 0) = ""
    campos(0, 2) = "rc_rendicionesdecaja"

    condicion = "fecha = '" & dato3.text & "-" & dato2.text & "-" & dato1.text & "' AND codigocajera = '" & dato4.text + dv.Caption + "' AND local = '" & dato5.text & "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then
        leerCajaDiaria = True
        Call CARGACAJADIARIA2
        
        TVENTAS = Val(SQLUTIL.datos(21, 3))
        TDONACIONES = Val(SQLUTIL.datos(22, 3))
        TPAGOS = Val(SQLUTIL.datos(23, 3))
        TARENDIR = Val(SQLUTIL.datos(24, 3))
        TRENDIDO = Val(SQLUTIL.datos(25, 3))
        TDIFERENCIAS = Val(SQLUTIL.datos(26, 3))
        glosa.text = SQLUTIL.datos(27, 3)
        
        MUESTRATOTAL
        
        
        Rem Call cargaCajaDiaria
    Else
        leerCajaDiaria = False
        FRMEFECTIVO.Enabled = True
        FRMOTROS.Enabled = True
     
    End If
End Function


Private Sub cargaIngresadas()
Dim TOTALES(20) As Double

 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = db
    chequesdia = 0
    chequesfecha = 0
    Dim suma As Double
    sql.sql = "SELECT local,codigocajera,totalefectivo,montocheques,montochequefecha,montotcredito,montotdebito,montoextranjera,montootorgado,montovcompra,montoocredito,montoncredito,montovarios,totalrendido,totalventas+totaldonaciones,'0',totalpagos,totalarendir,'0',totalrendido-totalarendir,glosa "
    sql.sql = sql.sql + "FROM rc_rendicionesdecaja "
    sql.sql = sql.sql + "WHERE fecha = '" & dato3.text & "-" & dato2.text & "-" & dato1.text & "' "
    If local1.Value = True Then
        sql.sql = sql.sql + "and local<>'43' and local<>'42' "
        End If
        If local2.Value = True Then
        sql.sql = sql.sql + "and (local='43' or local='42') "
        End If

    
    sql.sql = sql.sql + " order by local,codigocajera "
    sql.Execute
    For k = 1 To 20
    TOTALES(k) = 0
    Next k
    chequesdia = 0
    chequesfecha = 0
   Ingresadas.Rows = 1
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        Ingresadas.Rows = 1
        While resultados.EOF = False
        Ingresadas.Rows = Ingresadas.Rows + 1
            Ingresadas.Cell(Ingresadas.Rows - 1, 1).text = resultados(0)
            Ingresadas.Cell(Ingresadas.Rows - 1, 2).text = leerCajera(resultados(1))
            For k = 2 To 19
            Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text = resultados(k)
            If k = 3 Then
            Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text = montocheques(dato3.text + "-" + dato2.text + "-" + dato1.text, resultados(0), Mid(resultados(1), 1, 9), AÑO1.text + "-" + MES1.text + "-" + DIA1.text, "1")
            End If
            If k = 4 Then
            Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text = montocheques(dato3.text + "-" + dato2.text + "-" + dato1.text, resultados(0), Mid(resultados(1), 1, 9), AÑO1.text + "-" + MES1.text + "-" + DIA1.text, "2")
            End If
            Ingresadas.Cell(Ingresadas.Rows - 1, 21).text = resultados(20)
            
            TOTALES(k) = TOTALES(k) + CDbl(Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text)
            Next k
            resultados.MoveNext
        Wend
            Ingresadas.Rows = Ingresadas.Rows + 1
            Ingresadas.Cell(Ingresadas.Rows - 1, 2).text = "TOTALES "
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellEdgeTop) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellEdgeLeft) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellEdgeRight) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellInsideHorizontal) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellInsideVertical) = cellThin
    
            For k = 2 To 19
            Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text = TOTALES(k)
            Next k
            Ingresadas.Rows = Ingresadas.Rows + 1
            Ingresadas.Cell(Ingresadas.Rows - 1, 1).text = sql.RowsAffected
            Ingresadas.Cell(Ingresadas.Rows - 1, 2).text = "CHEQUES "
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, 5).Borders(cellEdgeTop) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, 5).Borders(cellEdgeBottom) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, 5).Borders(cellEdgeLeft) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, 5).Borders(cellEdgeRight) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, 5).Borders(cellInsideHorizontal) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, 5).Borders(cellInsideVertical) = cellThin
            Ingresadas.Cell(Ingresadas.Rows - 1, 4).text = chequesdia
            Ingresadas.Cell(Ingresadas.Rows - 1, 5).text = chequesfecha
            Ingresadas.Cell(Ingresadas.Rows - 1, 7).text = "SISTEMA"
            Ingresadas.Cell(Ingresadas.Rows - 1, 9).text = Leercreditos(dato3.text + "-" + dato2.text + "-" + dato1.text)
            Ingresadas.Cell(Ingresadas.Rows - 1, 17).text = Leerpagosdia(dato3.text + "-" + dato2.text + "-" + dato1.text)
            Ingresadas.Rows = Ingresadas.Rows + 1
            Ingresadas.Cell(Ingresadas.Rows - 1, 7).text = "DIFE"
            Ingresadas.Cell(Ingresadas.Rows - 1, 9).text = TOTALES(8) - Leercreditos(dato3.text + "-" + dato2.text + "-" + dato1.text)
            Ingresadas.Cell(Ingresadas.Rows - 1, 17).text = TOTALES(16) - Leerpagosdia(dato3.text + "-" + dato2.text + "-" + dato1.text)
            
            
    End If



End Sub

Private Sub grabarcaja()
    'Dim sql As New csqlutil
    Dim i As Long
    Dim cad As String
    Dim op As Integer
    Dim condicion As String
    
    campos(0, 0) = "fecha"
    campos(1, 0) = "codigocajera"
    campos(2, 0) = "local"
    campos(3, 0) = "caja"
    campos(4, 0) = "montoretiros"
    campos(5, 0) = "billeteveinte"
    campos(6, 0) = "billetediez"
    campos(7, 0) = "billetecinco"
    campos(8, 0) = "billetedos"
    campos(9, 0) = "billetemil"
    campos(10, 0) = "sencillo"
    campos(11, 0) = "montocheques"
    campos(12, 0) = "montotdebito"
    campos(13, 0) = "montotcredito"
    campos(14, 0) = "montoextranjera"
    campos(15, 0) = "montootorgado"
    campos(16, 0) = "montovcompra"
    campos(17, 0) = "montoocredito"
    campos(18, 0) = "montoncredito"
    campos(19, 0) = "montocasascomerciales"
    campos(20, 0) = "montovarios"
    campos(21, 0) = "totalventas"
    campos(22, 0) = "totaldonaciones"
    campos(23, 0) = "totalpagos"
    campos(24, 0) = "totalarendir"
    campos(25, 0) = "totalrendido"
    campos(26, 0) = "totaldiferenciadecaja"
    campos(27, 0) = "totalefectivo"
    campos(28, 0) = "glosa"
    campos(29, 0) = ""
    campos(0, 1) = dato3.text & "-" & dato2.text & "-" & dato1.text
    campos(1, 1) = dato4.text + dv.Caption
    campos(2, 1) = dato5.text
    campos(3, 1) = "00"
    
    
    For i = 4 To 10
        campos(i, 1) = Val(Dineros.Cell(i - 3, 1).text)
    Next i
    For i = 11 To 20
        campos(i, 1) = Val(Otros.Cell(i - 10, 1).text)
    Next i
    campos(21, 1) = Str(TVENTAS)
    campos(22, 1) = Str(TDONACIONES)
    campos(23, 1) = Str(TPAGOS)
    campos(24, 1) = Str(TARENDIR)
    campos(25, 1) = Str(TOTALRENDIDO.Caption)
    campos(26, 1) = Str(TDIFERENCIAS)
    campos(27, 1) = CDbl(EFECTIVO.Caption)
    campos(28, 1) = glosa.text
    
    campos(0, 2) = "rc_rendicionesdecaja"
    
    condicion = ""
    op = 2
    SQLUTIL.datos = campos
    'Call Auditoria(sql)
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
End Sub


Private Sub eliminacaja()
    'Dim sql As New csqlutil
    Dim i As Long
    Dim cad As String
    Dim op As Integer
    Dim condicion As String
    
Rem elimina caja
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND codigocajera = '" & dato4.text & dv.Caption & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_rendicionesdecaja"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If modificar = False Then
Rem elimina contables
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_contabilidad"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina creditos otorgados
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_creditosotorgados"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina creditospagados
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_creditospagados"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina monedaextranjera
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_monedaextranjera"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina notasdecredito
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_notasdecreditos"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina tarjetas de credito
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_tarjetacredito"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina tarjeta de debito
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_tarjetadebito"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina tarjetascasascomerciales
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_tarjetascasascomerciales"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina vales de compra
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_valesdecompra"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina otros creditos
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_otroscreditos"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem elimina cartera
    cad = dato3.text & "-" & dato2.text & "-" & dato1.text
    condicion = "fecha = '" & cad & "' AND cajera = '" & dato4.text & "' AND local = '" & dato5.text & "'"
    campos(0, 2) = "rc_cartera"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

End If






End Sub



Private Sub modifica()
    modificar = True
    loc = dato5.text
    dato1.Locked = True
    dato2.Locked = True
    dato3.Locked = True
    dato4.Locked = True
    dato5.SetFocus
End Sub

Private Sub elimina()
    Call eliminacaja
 
End Sub

Private Sub retorno()
    
    modificar = False
    
    CABEZA.Enabled = True
    FRMEFECTIVO.Enabled = True
    FRMOTROS.Enabled = True
    OPCIONES.Visible = False
    GRABAR.Visible = False
    
    dato1.Locked = False
    dato2.Locked = False
    dato3.Locked = False
    dato4.Locked = False
    Call LimpiarCajas(Me)
            dato1.text = Format(fechasistema, "dd")
        dato2.text = Format(fechasistema, "mm")
        dato3.text = Format(fechasistema, "yyyy")
        DIA1.text = dato1.text
        MES1.text = dato2.text
        AÑO1.text = dato3.text

    Call LimpiarLabels(Me)
    Dineros.Range(1, 1, Dineros.Rows - 1, Dineros.Cols - 1).ClearText
    Otros.Range(1, 1, Otros.Rows - 1, Otros.Cols - 1).ClearText
    
    TVENTAS = 0
    TDONACIONES = 0
    TPAGOS = 0
    TARENDIR = 0
    TRENDIDO = 0
    TDIFERENCIAS = 0
    MUESTRATOTAL
    
    TOTALDIFERENCIACAJA.Caption = "0"
    TOTALDONACIONES.text = "0"
    TOTALVENTAS.text = "0"
    TOTALARENDIR.Caption = "0"
    TOTALRENDIDO.Caption = "0"
    TOTALPAGOS.text = "0"
    EFECTIVO.Caption = "0"
    
    
    dato1.SetFocus
End Sub



Private Function LeerTotalVenta() As String
    Dim condicion As String
    Dim op As Integer
    'Set sql = New CSQLUtil
    campos(0, 0) = "sum(total)"
    campos(1, 0) = ""
    
    campos(0, 2) = "sv_documento_cabeza_" + dato5.text

    condicion = "fecha = '" & dato3.text & "-" & dato2.text & "-" & dato1.text & "' AND local = '" & dato5.text & "' AND cajera = '" & dato4.text & "' AND (tipo = 'BV' OR tipo = 'FV')"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = ventasRubro
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then
        LeerTotalVenta = SQLUTIL.datos(0, 3)
       
        
    Else
        LeerTotalVenta = "0"
    End If
End Function

Private Sub leerVentasDia()

End Sub
Public Sub ayudaCajera(ByRef txt As TextBox)
    servidorAyuda = servidor
    basedatosAyuda = baseventas
    usuarioAyuda = usuario
    passAyuda = password
    tablaAyuda = "sv_maestrocajeras"
    mensajeAyuda = "Ayuda de Cajeras"
    camposAyuda = Array("rut", "nombre")
    cabezasAyuda = Array("rut", "nombre")
    largoAyuda = Array("10n", "30s")
    condicionAyuda = "no"
    cantidadAyuda = 2
    Call Mayuda.cargaAyuda(txt)
End Sub

Public Sub ayudalocales(ByRef txt As TextBox)
    
    servidorAyuda = servidor
    basedatosAyuda = clientesistema + "gestion"
    usuarioAyuda = usuario
    passAyuda = password
    tablaAyuda = "g_maestroempresas"
    mensajeAyuda = "Ayuda de Locales"
    camposAyuda = Array("codigo", "nombre")
    cabezasAyuda = Array("codigo", "nombre")
    largoAyuda = Array("3n", "30s")
    condicionAyuda = "no"
    cantidadAyuda = 2
    Call Mayuda.cargaAyuda(txt)
End Sub

Sub Flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef CODIGO As Integer)
    If CODIGO = 38 And caja.Enabled = True Then caja.SetFocus
    If CODIGO = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub


Sub CARGACAJADIARIA2()
        
        Dineros.Cell(1, 1).text = SQLUTIL.datos(4, 3)
        Dineros.Cell(2, 1).text = SQLUTIL.datos(5, 3)
        Dineros.Cell(3, 1).text = SQLUTIL.datos(6, 3)
        Dineros.Cell(4, 1).text = SQLUTIL.datos(7, 3)
        Dineros.Cell(5, 1).text = SQLUTIL.datos(8, 3)
        Dineros.Cell(6, 1).text = SQLUTIL.datos(9, 3)
        Dineros.Cell(7, 1).text = SQLUTIL.datos(10, 3)
        
        Otros.Cell(1, 1).text = SQLUTIL.datos(11, 3)
        Otros.Cell(2, 1).text = SQLUTIL.datos(12, 3)
        Otros.Cell(3, 1).text = SQLUTIL.datos(13, 3)
        Otros.Cell(4, 1).text = SQLUTIL.datos(14, 3)
        Otros.Cell(5, 1).text = SQLUTIL.datos(15, 3)
        Otros.Cell(6, 1).text = SQLUTIL.datos(16, 3)
        Otros.Cell(7, 1).text = SQLUTIL.datos(17, 3)
        Otros.Cell(8, 1).text = SQLUTIL.datos(18, 3)
        Otros.Cell(9, 1).text = SQLUTIL.datos(19, 3)
        Otros.Cell(10, 1).text = SQLUTIL.datos(20, 3)
        
sumaGrilla
End Sub



'Private Function notadecreditoexiste(ByVal codigoLocal As String, ByVal NUMERO As String) As Boolean
'    Dim condicion As String
'    Dim op As Integer
'    'Set sql = New CSQLUtil
'
'    campos(0, 0) = "local"
'    campos(1, 0) = "tipo"
'    campos(2, 0) = "numero"
'    campos(3, 0) = "total"
'    campos(4, 0) = ""
'
'    campos(0, 2) = "sv_documento_cabeza_" + dato5.text
'
'
'    condicion = "local = '" & codigoLocal & "' and tipo = 'NB' and numero='" & NUMERO & "'"
'    op = 5
'    SQLUTIL.datos = campos
'    Set SQLUTIL.conexion = ventasRubro
'    Call SQLUTIL.SQLUTIL(op, condicion)
'    If SQLUTIL.estado = 0 Then
'        notadecreditoexiste = True
'        totalnc = CDbl(SQLUTIL.datos(3, 3))
'    Else
'        notadecreditoexiste = False
'    End If
'End Function
'Sub notarendida(ByVal codigoLocal As String, ByVal NUMERO As String)
'    Dim condicion As String
'    Dim op As Integer
'    'Set sql = New CSQLUtil
'
'    campos(0, 0) = "fecharendicion"
'    campos(1, 0) = ""
'
'    campos(0, 2) = "sv_documento_cabeza_" + dato5.text
'
'    campos(0, 1) = dato3.text + "-" + dato2.text + "-" + dato1.text
'
'    condicion = "local = '" & codigoLocal & "' and tipo = 'NB' and numero='" & NUMERO & "'"
'    op = 3
'    SQLUTIL.datos = campos
'    Set SQLUTIL.conexion = ventasRubro
'    Call SQLUTIL.SQLUTIL(op, condicion)
'    If SQLUTIL.estado <> 0 Then
'        MsgBox ("error en modificacion nota de credito")
'    End If
'
'End Sub
'
'

Private Sub TOTALDONACIONES_Click()

End Sub

        

Public Function montocheques(fecha, loc, rut, vencimiento, tipo) As Double

 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
        Set sql.ActiveConnection = db
        If tipo = "1" Then
        sql.sql = "SELECT sum(monto),count(monto) "
        sql.sql = sql.sql + "FROM rc_cartera "
        sql.sql = sql.sql + "WHERE fecha = '" & fecha & "' and local='" + loc + "' and cajera='" + rut + "' "
        If CIERRE.Value = 1 Then
        sql.sql = sql.sql + " and vencimiento<='" + vencimiento + "' "
        End If
        
        
        If CIERRE.Value <> 1 Then
        sql.sql = sql.sql + "and cartera='N' "
        End If
        
        sql.sql = sql.sql + "group by fecha "
        
        
        sql.Execute
        montocheques = 0
            If sql.RowsAffected > 0 Then
            Set resultados = sql.OpenResultset
            montocheques = resultados(0)
            chequesdia = chequesdia + resultados(1)
                
            End If
        
        If CIERRE.Value = 1 Then
        Call modificacartera(fecha, loc, rut, vencimiento, "N")
        End If
        
        
        End If
        
        If tipo = "2" Then
        sql.sql = "SELECT sum(monto),count(monto) "
        sql.sql = sql.sql + "FROM rc_cartera "
        sql.sql = sql.sql + "WHERE fecha = '" & fecha & "' and local='" + loc + "' and cajera='" + rut + "' "
        If CIERRE.Value = 1 Then
        sql.sql = sql.sql + " and vencimiento>'" + vencimiento + "' "
        End If
        
        
        If CIERRE.Value <> 1 Then
        sql.sql = sql.sql + "and cartera='S' "
        End If
        
        
        
        sql.sql = sql.sql + "group by fecha "
        sql.Execute
        montocheques = 0
    
            If sql.RowsAffected > 0 Then
            Set resultados = sql.OpenResultset
            montocheques = resultados(0)
            chequesfecha = chequesfecha + resultados(1)
            
            End If
        If CIERRE.Value = "1" Then
         Call modificacartera(fecha, loc, rut, vencimiento, "S")
        End If
        
        
        End If
        
    
End Function

Sub modificacartera(fecha, loc, rut, vencimiento, cartera)
 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
        Set sql.ActiveConnection = db
        
        If cartera = "N" Then
        sql.sql = "UPDATE rc_cartera set cartera='N' "
        sql.sql = sql.sql + "WHERE fecha = '" & fecha & "' and local='" + loc + "' and cajera='" + rut + "' and vencimiento<='" + vencimiento + "' "
        
        sql.Execute
        Else
        sql.sql = "UPDATE rc_cartera set cartera='S' "
        sql.sql = sql.sql + "WHERE fecha = '" & fecha & "' and local='" + loc + "' and cajera='" + rut + "' and vencimiento>'" + vencimiento + "' "
        
        sql.Execute
        End If
        
      
End Sub

Sub leeradmin(fecha, loc, rut)
Dim sumacheques As Double
Dim sumatcredito As Double
Dim sumatdebito As Double
Dim sumacreditos As Double
Dim capitalcredito As Double

Dim lineatc As Double
Dim lineatd As Double
Dim lineacreditos As Double

 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
        Set sql.ActiveConnection = ventasRubro
        
        
 
        sql.sql = "select dc.cajera,dc.fecha,dp.tipopago,dp.cuentacorriente,dp.banco,dp.plaza,dp.numerodocumento,dp.monto,dp.vencimiento,dp.rutcredito,dc.tipo,dc.foliosii,dp.cuotas,dp.montocuotas "
        sql.sql = sql.sql + "from sv_documento_pagos_" + loc + " as dp inner join sv_documento_cabeza_" + loc + " as dc on (dp.tipo=dc.tipo and dp.numero=dc.numero and dp.fecha=dc.fecha) "
        sql.sql = sql.sql + "where dc.fecha='" + Format(fecha, "yyyy-mm-dd") + "' and dp.tipopago<>'1' and dc.nula<>'S' and dc.cajera='" + rut + "' "
        sql.sql = sql.sql + "order by dc.cajera "
        sql.Execute
        sumacheques = 0
        sumatcredito = 0
        sumatdebito = 0
        sumacreditos = 0
        lineatc = 0
        lineatd = 0
        lineacreditos = 0
        If sql.RowsAffected > 1 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
        If resultados(2) = "2" Then
        sumacheques = sumacheques + resultados(7)
        Call GRABARcheques(fecha, rut, loc, resultados(8), resultados(6), resultados(7), "N", resultados(3), resultados(4), resultados(5), "00")
        End If
        If resultados(2) = "3" Then
        lineatc = lineatc + 1
        sumatcredito = sumatcredito + resultados(7)
        Call GRABARtarjetas(fecha, rut, loc, "00", lineatc, "TB", resultados(7), "C")
        End If
        If resultados(2) = "4" Then
        lineatd = lineatd + 1
        sumatdebito = sumatdebito + resultados(7)
        Call GRABARtarjetas(fecha, rut, loc, "00", lineatc, "TB", resultados(7), "D")
        End If
        If resultados(2) = "6" Then
        lineacreditos = lineacreditos + 1
        capitalcredito = Leercredito(loc, resultados(9), resultados(10), resultados(11), resultados(1))
        sumacreditos = sumacreditos + capitalcredito
        
        
        Call GRABARcreditos(fecha, rut, loc, "00", lineacreditos, loc, resultados(9), resultados(12), resultados(13), capitalcredito)
        End If
        
        
        resultados.MoveNext
        
        Wend
        
        
        
        End If
        
        ingreso01.Otros.Cell(1, 1).text = sumacheques
        ingreso01.Otros.Cell(2, 1).text = sumatcredito
        ingreso01.Otros.Cell(3, 1).text = sumatdebito
        ingreso01.Otros.Cell(5, 1).text = sumacreditos

        
      
End Sub


Public Sub GRABARcheques(fecha, cajera, loc, vencimiento, numero, MONTO, cartera, cuenta, banco, PLAZA, caja)
    'Dim sql As New csqlutil
    
    Dim i As Long
    Dim cad As String
    Dim op As Integer
    Dim condicion As String
    Dim TOTAL As Double
    campos(0, 0) = "fecha"
    campos(1, 0) = "cajera"
    campos(2, 0) = "local"
    campos(3, 0) = "vencimiento"
    campos(4, 0) = "numero"
    campos(5, 0) = "monto"
    campos(6, 0) = "cartera"
    campos(7, 0) = "cuenta"
    campos(8, 0) = "banco"
    campos(9, 0) = "plaza"
    campos(10, 0) = "caja"
    campos(11, 0) = ""
    campos(0, 2) = "rc_cartera"
    
    condicion = ""
    op = 2
    TOTAL = 0
    campos(0, 1) = Format(fecha, "yyyy-mm-dd")
    campos(1, 1) = cajera
    campos(2, 1) = loc
    campos(3, 1) = Format(vencimiento, "yyyy-mm-dd")
    campos(4, 1) = Format(numero, "0000000")
    campos(5, 1) = MONTO
    campos(6, 1) = cartera
    campos(7, 1) = Format(cuenta, "000000000000")
    campos(8, 1) = Format(banco, "000")
    campos(9, 1) = Format(PLAZA, "0000")
    campos(10, 1) = caja
    
    SQLUTIL.datos = campos
            
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
End Sub


Sub GRABARtarjetas(fecha, cajera, loc, caja, linea, tipo, MONTO, tipoTA)
    'Dim sql As New csqlutil
    
    Dim i As Long
    Dim cad As String
    Dim op As Integer
    Dim condicion As String
    Dim TOTAL As Double
    campos(0, 0) = "fecha"
    campos(1, 0) = "cajera"
    campos(2, 0) = "local"
    campos(3, 0) = "caja"
    campos(4, 0) = "linea"
    campos(5, 0) = "tipo"
    campos(6, 0) = "monto"
    campos(7, 0) = ""
    If tipoTA = "C" Then
    campos(0, 2) = "rc_tarjetacredito"
    Else
    campos(0, 2) = "rc_tarjetadebito"
    End If
    
    condicion = ""
    op = 2
    campos(0, 1) = Format(fecha, "yyyy-mm-dd")
    campos(1, 1) = cajera
    campos(2, 1) = loc
    campos(3, 1) = caja
    campos(4, 1) = linea
    campos(5, 1) = tipo
    campos(6, 1) = MONTO
    
    SQLUTIL.datos = campos
            'Call Auditoria(sql)
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
End Sub

Private Sub GRABARcreditos(fecha, cajera, loc, caja, linea, localcredito, rut, CUOTAS, VALOR, TOTAL)
    'Dim sql As New csqlutil
    
    Dim i As Long
    Dim cad As String
    Dim op As Integer
    Dim condicion As String
   
    campos(0, 0) = "fecha"
    campos(1, 0) = "cajera"
    campos(2, 0) = "local"
    campos(3, 0) = "caja"
    campos(4, 0) = "linea"
    campos(5, 0) = "localcredito"
    campos(6, 0) = "rut"
    campos(7, 0) = "cuotas"
    campos(8, 0) = "valor"
    campos(9, 0) = "total"
    campos(10, 0) = ""
    
    campos(0, 2) = "rc_creditosotorgados"
    
    condicion = ""
    op = 2
  
    campos(0, 1) = fecha
    campos(1, 1) = cajera
    campos(2, 1) = loc
    campos(3, 1) = caja
    campos(4, 1) = linea
    campos(5, 1) = localcredito
    campos(6, 1) = rut
    campos(7, 1) = CUOTAS
    campos(8, 1) = VALOR
    campos(9, 1) = TOTAL
        
        
        SQLUTIL.datos = campos
            'Call Auditoria(sql)
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
End Sub

Private Function Leercredito(loc, rut, tipo, numero, fecha) As Double
    Dim condicion As String
    Dim op As Integer
    'Set sql = New CSQLUtil
    campos(0, 0) = "montocredito"
    campos(1, 0) = ""
    campos(0, 2) = "sv_cuotas_detalle"
    condicion = "fechacompra = '" & Format(fecha, "yyyy-mm-dd") & "' AND local = '" & loc & "' AND rut = '" & rut & "' and tipo='" + tipo + "' and numero='" + numero + "' limit 0,1"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = VENTAS
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then
        Leercredito = SQLUTIL.datos(0, 3)
       
        
    Else
        Leercredito = 0
    End If
End Function

Private Function Leercreditos(fecha) As Double
 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
        Set sql.ActiveConnection = VENTAS
        sql.sql = "SELECT sum(montocredito) "
        sql.sql = sql.sql + "FROM sv_cuotas_detalle "
        sql.sql = sql.sql + "WHERE fechacompra = '" & fecha & "' and numerocuota='1' "
        If local1.Value = True Then
        sql.sql = sql.sql + "and local<>'43' and local<>'42' "
        End If
        If local2.Value = True Then
        sql.sql = sql.sql + "and (local='43' or local='42') "
        End If
        
        sql.Execute
            If sql.RowsAffected > 0 Then
            Set resultados = sql.OpenResultset
            Leercreditos = resultados(0)
            Else
            Leercreditos = 0
            End If
        
End Function
Private Function Leerpagosdia(fecha) As Double
 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
        Set sql.ActiveConnection = VENTAS
        sql.sql = "SELECT sum(monto) "
        sql.sql = sql.sql + "FROM sv_cuotas_pago_cabeza "
        sql.sql = sql.sql + "WHERE fecha = '" & fecha & "' "
        If local1.Value = True Then
        sql.sql = sql.sql + "and local<>'43' and local<>'42' and local<>'77' "
        End If
        If local2.Value = True Then
        sql.sql = sql.sql + "and (local='43' or local='42' or local='77') "
        End If
        
        sql.Execute
            If sql.RowsAffected > 0 Then
            Set resultados = sql.OpenResultset
            Leerpagosdia = resultados(0)
            Else
            Leerpagosdia = 0
            End If
        
End Function

