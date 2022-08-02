VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form multi01 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEMA CONTROL MULTICAJA"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14475
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   662
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   965
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11400
      TabIndex        =   32
      Top             =   0
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
         TabIndex        =   34
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   33
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   17171
      BackColor       =   8454016
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   8454016
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
      Begin VB.OptionButton Option12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Creditos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   50
         Top             =   1480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6960
         TabIndex        =   49
         Top             =   1480
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "Debitos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   48
         Top             =   1480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   615
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         BackColor       =   49344
         Caption         =   "FILTRO POR MONTO"
         CaptionEstilo3D =   1
         BackColor       =   49344
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
         Begin VB.TextBox MONTO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   3495
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   1455
         Left            =   4080
         TabIndex        =   35
         Top             =   0
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2566
         BackColor       =   12640511
         Caption         =   "Ventas Multicard                         Abonos Multicard"
         CaptionEstilo3D =   1
         BackColor       =   12640511
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
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Ventas Resumen"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1200
            Width           =   2175
         End
         Begin VB.OptionButton Option10 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Abonos Resumen"
            Height          =   255
            Left            =   2400
            TabIndex        =   46
            Top             =   1200
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Abonos detallado"
            Height          =   255
            Left            =   2400
            TabIndex        =   45
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Abonos conciliado"
            Height          =   255
            Left            =   2400
            TabIndex        =   44
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Abonos Inconsistencias"
            Height          =   255
            Left            =   2400
            TabIndex        =   43
            Top             =   960
            Width           =   2175
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Abonos Totalizado"
            Height          =   255
            Left            =   2400
            TabIndex        =   42
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Venta inconsistencias"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Venta Totalizado"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Venta conciliado"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Venta detallado"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Generar Informacion"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CARGAR CARTOLAS"
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
         Left            =   120
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   2355
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   8025
         Left            =   90
         TabIndex        =   1
         Top             =   2070
         Width           =   14145
         _ExtentX        =   24950
         _ExtentY        =   14155
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
         Begin XPFrame.FrameXp frameelimina 
            Height          =   1665
            Left            =   4560
            TabIndex        =   13
            Top             =   2880
            Visible         =   0   'False
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   2937
            BackColor       =   14737632
            Caption         =   "Elimina cartolas X Rangos de Fecha"
            CaptionEstilo3D =   1
            BackColor       =   14737632
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
            Begin VB.CommandButton Command9 
               Caption         =   "Cancela"
               Height          =   375
               Left            =   2520
               TabIndex        =   28
               Top             =   1200
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker elihasta 
               Height          =   255
               Left            =   2520
               TabIndex        =   18
               Top             =   840
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   450
               _Version        =   393216
               Format          =   160563201
               CurrentDate     =   40274
            End
            Begin MSComCtl2.DTPicker elidesde 
               Height          =   255
               Left            =   360
               TabIndex        =   17
               Top             =   840
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   450
               _Version        =   393216
               Format          =   160563201
               CurrentDate     =   40274
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Elimina cartolas"
               Height          =   375
               Left            =   240
               TabIndex        =   16
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Hasta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   2520
               TabIndex        =   15
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Desde "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   360
               TabIndex        =   14
               Top             =   360
               Width           =   1935
            End
         End
         Begin XPFrame.FrameXp CARGATXT 
            Height          =   4200
            Left            =   2880
            TabIndex        =   19
            Top             =   1320
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   7408
            BackColor       =   16761024
            Caption         =   "BUSCAR "
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
            Begin VB.FileListBox File1 
               Height          =   2235
               Left            =   4230
               TabIndex        =   25
               Top             =   315
               Width           =   4275
            End
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   180
               TabIndex        =   24
               Top             =   315
               Width           =   3855
            End
            Begin VB.TextBox ARCHIVO 
               Height          =   285
               Left            =   4230
               TabIndex        =   23
               Top             =   3060
               Width           =   4275
            End
            Begin VB.DirListBox Dir1 
               Height          =   2565
               Left            =   180
               TabIndex        =   22
               Top             =   765
               Width           =   3855
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H00FF8080&
               Caption         =   "PROCESAR"
               Height          =   465
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   3465
               Width           =   2625
            End
            Begin VB.CommandButton Command3 
               BackColor       =   &H00FF8080&
               Caption         =   "RETORNO"
               Height          =   465
               Left            =   4635
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   3465
               Width           =   2625
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ARCHIVO SELECCIONADO"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   4230
               TabIndex        =   26
               Top             =   2790
               Width           =   4290
            End
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5625
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   7290
            Width           =   2625
         End
         Begin MSComctlLib.ProgressBar barra 
            Height          =   195
            Left            =   135
            TabIndex        =   11
            Top             =   6975
            Width           =   13920
            _ExtentX        =   24553
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Conciliar Tarjetas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2385
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   7290
            Visible         =   0   'False
            Width           =   2625
         End
         Begin FlexCell.Grid Grid1 
            Height          =   6630
            Left            =   -30
            TabIndex        =   2
            Top             =   225
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   11695
            AllowUserSort   =   -1  'True
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1665
         Left            =   9120
         TabIndex        =   4
         Top             =   240
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   2937
         BackColor       =   14737632
         Caption         =   "Rangos de Fecha"
         CaptionEstilo3D =   1
         BackColor       =   14737632
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
         Begin CoolButtons.cool_Button command8 
            Height          =   375
            Left            =   1440
            TabIndex        =   5
            Top             =   1260
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            SkinId          =   "13"
            Caption         =   "Cambia Fecha"
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desde Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hasta Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label desdefecha 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label hastafecha 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   6
            Top             =   720
            Width           =   1935
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   675
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   "TERMINALES"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Alignment       =   1
         Begin VB.ComboBox ComboLOCAL 
            Height          =   315
            Left            =   90
            TabIndex        =   31
            Top             =   270
            Width           =   3675
         End
      End
      Begin VB.CommandButton ELIMINA 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ELIMINA CARTOLAS"
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
         Height          =   255
         Left            =   2520
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   2355
      End
   End
End
Attribute VB_Name = "multi01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BENEFICIARIO As String
Dim banco_cuenta As String
Dim banco_glosa As String
Dim banco_dh As String
Dim conta_glosa As String
Dim conta_fecha As String
Dim conta_monto As String
Dim banco_glosa2 As String
Dim total(20) As Double
Dim saldo As Double






 

Private Sub Command1_Click()
CARGATXT.Visible = True

End Sub

Private Sub COMMAND2_Click()
Dim o As Double
Dim origen As String
Dim destino As String
Dim original As String


For o = 0 To File1.ListCount - 1
If Mid(File1.List(0), 1, 7) = "transac" Then
        ARCHIVO.text = File1.List(o)
        original = ARCHIVO.text
                If UCase(Right(ARCHIVO.text, 3)) = "XLS" Then
                CARGATXT.Visible = True
                Call saveExcelAsCsv("u:\multicaja_files\" + ARCHIVO.text)
                Rem ARCHIVO.text = "transacciones_Mes_Aprobadas_47850.xls"
                ARCHIVO.text = Replace(ARCHIVO.text, "xls", "txt")
                TRASPASADATOS
            
                CARGATXT.Visible = False
                origen = "u:\MULTICAJA_FILES\" + original
                destino = "u:\MULTICAJA_FILES_usados\" + original
'                FileCopy origen, destino
'                Kill origen
                origen = "u:\MULTICAJA_FILES\" + ARCHIVO.text
                destino = "u:\MULTICAJA_FILES_usados\" + ARCHIVO.text
                FileCopy origen, destino
                Kill origen
'
                
                Else
                MsgBox ("ESTE ARCHIVO NO ES UN ETBK")
                End If

End If

If Mid(File1.List(0), 1, 7) = "7757534" Then
                ARCHIVO.text = File1.List(o)
                If UCase(Right(ARCHIVO.text, 3)) = "CSV" Then
                CARGATXT.Visible = True

                TRASPASADATOS2

                CARGATXT.Visible = False
                origen = "u:\MULTICAJA_FILES\" + ARCHIVO.text
                destino = "u:\MULTICAJA_FILES_usados\" + ARCHIVO.text
                FileCopy origen, destino

                Kill origen
                Else
                MsgBox ("ESTE ARCHIVO NO ES UN ETBK")
                End If

End If

Next o

End Sub

Sub saveExcelAsCsv(ExcelFileName As String)
'Dim objXlApp As New Excel.Application
'Dim objXlBook As Excel.Workbook
'' Setup the Excel Workbook to save
'Set objXlBook = Excel.Workbooks.Open(ExcelFileName)
'' Save the Excel file as a CSV file in a temp location
'objXlBook.SaveAs Replace(ExcelFileName, "xls", "txt"), xlCSVWindows
'' Clean up
'objXlBook.Close True
'' Save changes
'objXlApp.Quit
'' Close Excel
'Set objXlBook = Nothing
'Set objXlApp = Nothing
'Kill ExcelFileName
End Sub
Sub TRASPASADATOS()
Dim lin As Double
Dim palabras() As String
Dim palabras2() As String
Close 20


Open File1.path + "\" + ARCHIVO.text For Input As #20
lin = 0

While EOF(20) = False
 
   
Line Input #20, varipaso

palabras() = Split(varipaso, ",")
'Rem palabras2() = Split(varipaso, ",")
lin = UBound(palabras())
barra.Max = lin + 1
barra.Value = 0
'



If palabras(0) <> "Fecha" Then
Call GRABACARTOLA(palabras(0), palabras(1), palabras(2), palabras(3), palabras(4), palabras(5), palabras(6), palabras(7), palabras(8), palabras(9))
barra.Value = barra.Value + 1
barra.Refresh

End If



'Next k
Wend

Rem Call GRABACARTOLA(lin)
Close 20

End Sub

Sub TRASPASADATOS2()
Dim lin As Double
Dim palabras() As String
Dim palabras2() As String
Close 20
Open File1.path + "\" + ARCHIVO.text For Input As #20
lin = 0
While EOF(20) = False
 
   
Line Input #20, varipaso

palabras() = Split(varipaso, Chr(10))
Rem palabras2() = Split(varipaso, ",")
lin = UBound(palabras())
barra.Max = lin + 1
barra.Value = 0

For k = 0 To lin - 1
palabras2() = Split(palabras(k), ";")
If UBound(palabras2()) = 14 Then
If palabras2(0) <> "emisor" Then

Call GRABACARTOLA_ABONOS_td(palabras2(0), palabras2(1), palabras2(2), palabras2(3), palabras2(4), palabras2(5), palabras2(6), palabras2(7), palabras2(8), palabras2(9), palabras2(10), palabras2(11), palabras2(12), palabras2(13), palabras2(14), Right(palabras2(3), 8))

End If

barra.Value = barra.Value + 1
barra.Refresh

End If


Next k
Wend

Rem Call GRABACARTOLA(lin)
Close 20

End Sub
 

Sub GRABACARTOLA(fecha, hora, terminal, canal, transacciones, estado, respuesta, detalle, codigo_mc, monto)
    campos(0, 0) = "fecha"
    campos(1, 0) = "hora"
    campos(2, 0) = "terminal"
    campos(3, 0) = "canal"
    campos(4, 0) = "transacciones"
    campos(5, 0) = "estado"
    campos(6, 0) = "respuesta"
    campos(7, 0) = "detalle"
    campos(8, 0) = "codigo_mc"
    campos(9, 0) = "monto"
    campos(10, 0) = ""
    campos(11, 0) = ""
    campos(0, 1) = Format(fecha, "yyyy-mm-dd")
    campos(1, 1) = hora
    campos(2, 1) = terminal
    campos(3, 1) = canal
    campos(4, 1) = transacciones
    campos(5, 1) = estado
    campos(6, 1) = respuesta
    campos(7, 1) = detalle
    campos(8, 1) = codigo_mc
    campos(9, 1) = Replace(monto, ".", "")
    
    campos(0, 2) = "rc_multicaja"
           
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   
End Sub
Sub GRABACARTOLA_ABONOS_tc(tipo_transaccion, fecha_venta, tipo_tarjeta, identificador, tipocuota, monto_original, codigo_autorizacion, ncuota, monto_para_abono, comision_iva, comision_adicional, boleta, monto_anulacion, devolucion_comision, devolucion_adicional, monto_retencion, periodo_de_cobro, motivo, detalle_cobros, MONTO_2, iva, fecha_abono, cuenta_abono, loc)
    campos(0, 0) = "tipo_transaccion"
    campos(1, 0) = "fecha_venta"
    campos(2, 0) = "tipo_tarjeta"
    campos(3, 0) = "identificador"
    campos(4, 0) = "tipocuota"
    campos(5, 0) = "monto_original"
    campos(6, 0) = "codigo_autorizacion"
    campos(7, 0) = "ncuota"
    campos(8, 0) = "monto_para_abono"
    campos(9, 0) = "comision_iva"
    campos(10, 0) = "comision_adicional"
    campos(11, 0) = "boleta"
    campos(12, 0) = "monto_anulacion"
    campos(13, 0) = "devolucion_comision"
    campos(14, 0) = "devolucion_adicional"
    campos(15, 0) = "monto_retencion"
    campos(16, 0) = "periodo_de_cobro"
    campos(17, 0) = "motivo"
    campos(18, 0) = "detalle_cobros"
    campos(19, 0) = "monto"
    campos(20, 0) = "iva"
    campos(21, 0) = "fecha_abono"
    campos(22, 0) = "cuenta_abono"
    campos(23, 0) = "loc"
    campos(24, 0) = ""
    campos(0, 1) = tipo_transaccion
    campos(1, 1) = Format(fecha_venta, "yyyy-mm-dd")
    campos(2, 1) = tipo_tarjeta
    campos(3, 1) = identificador
    campos(4, 1) = tipocuota
    campos(5, 1) = Replace(monto_original, ".", "")
    campos(6, 1) = codigo_autorizacion
    campos(7, 1) = ncuota
    campos(8, 1) = Replace(monto_para_abono, ".", "")
    campos(9, 1) = Replace(comision_iva, ".", "")
    campos(10, 1) = Replace(comision_adicional, ".", "")
    campos(11, 1) = boleta
    campos(12, 1) = Replace(monto_anulacion, ".", "")
    campos(13, 1) = Replace(devolucion_comision, ".", "")
    campos(14, 1) = Replace(devolucion_adicional, ".", "")
    campos(15, 1) = Replace(monto_retencion, ".", "")
    campos(16, 1) = Replace(periodo_de_cobro, ".", "")
    campos(17, 1) = motivo
    campos(18, 1) = detalle_cobros
    campos(19, 1) = Replace(MONTO_2, ".", "")
    campos(20, 1) = Replace(iva, ".", "")
    campos(21, 1) = Format(fecha_abono, "yyyy-mm-dd")
    campos(22, 1) = cuenta_abono
    campos(23, 1) = Mid(loc, 1, 8)
    If Len(fecha_venta) > 12 Then
    campos(24, 1) = Right(fecha_venta, 5)
    Else
    campos(24, 1) = "00:00"
    
    End If
    
    
    
    campos(0, 2) = "rc_transbank_abonos_tc"
           
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
   
End Sub
Sub GRABACARTOLA_ABONOS_td(emisor, rut_prestador, lugar, fecha, codigo_mc, rut_titular, Aplicacion, monto, cuotas, respuesta, comision_emisor, comision_operador, costo_transaccion, monto_pagado, terminal, hora)
    campos(0, 0) = "emisor"
    campos(1, 0) = "rut_prestador"
    campos(2, 0) = "lugar"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigo_mc"
    campos(5, 0) = "rut_titular"
    campos(6, 0) = "aplicacion"
    campos(7, 0) = "monto"
    campos(8, 0) = "cuotas"
    campos(9, 0) = "respuesta"
    campos(10, 0) = "comision_emisor"
    campos(11, 0) = "comision_operador"
    campos(12, 0) = "costo_transaccion"
    campos(13, 0) = "monto_pagado"
    campos(14, 0) = "terminal"
    campos(15, 0) = "hora"
    campos(16, 0) = ""
    campos(0, 1) = emisor
    campos(1, 1) = rut_prestador
    campos(2, 1) = lugar
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigo_mc
    campos(5, 1) = rut_titular
    campos(6, 1) = Aplicacion
    campos(7, 1) = monto
    campos(8, 1) = cuotas
    campos(9, 1) = respuesta
    campos(10, 1) = comision_emisor
    campos(11, 1) = comision_operador
    campos(12, 1) = costo_transaccion
    campos(13, 1) = monto_pagado
    campos(14, 1) = terminal
    campos(15, 1) = hora
    
    campos(0, 2) = "rc_multicaja_abonos"
           
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   
End Sub



Private Sub Command3_Click()
CARGATXT.Visible = False


End Sub


Private Sub Command4_Click()
Dim DIFE As Double
Dim fechapaso As String
Dim fecharevisa As String

Call CARGAGRILLA2
total(1) = 0
total(2) = 0
total(3) = 0
total(4) = 0
total(5) = 0
total(6) = 0

DIFE = DateDiff("d", Format(desdefecha.Caption, "yyyy-mm-dd"), Format(hastafecha.Caption, "yyyy-mm-dd")) + 1
fechapaso = Format(desdefecha.Caption, "yyyy-mm-dd")
saldo = 0
For k = 1 To DIFE

Call revisa_transbank(fechapaso, Mid(ComboLOCAL.text, 1, 8))
fecharevisa = fechapaso
fechapaso = DateAdd("d", 1, Format(fechapaso, "yyyy-mm-dd"))

        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 7).BackColor = vbGreen
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(fecharevisa, "dd-mm-yyyy")
        
        Grid1.Cell(Grid1.Rows - 1, 4).text = total(1)
        Grid1.Cell(Grid1.Rows - 1, 5).text = total(2)
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 5).text = "DIFERENCIA"
        
        Grid1.Cell(Grid1.Rows - 1, 6).text = total(1) - total(2)
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0
        

Next k
End Sub
Private Sub resumen_credito()
Dim DIFE As Double
Dim fechapaso As String
Dim fecharevisa As String

Call CARGAGRILLA2
total(1) = 0
total(2) = 0
total(3) = 0
total(4) = 0
total(5) = 0
total(6) = 0


Call revisa_transbank_resumen(desdefecha.Caption, hastafecha.Caption, Mid(ComboLOCAL.text, 1, 8))

Rem fecharevisa = fechapaso
Rem fechapaso = DateAdd("d", 1, Format(fechapaso, "yyyy-mm-dd"))
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 7).BackColor = vbGreen
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(fecharevisa, "dd-mm-yyyy")
        
        Grid1.Cell(Grid1.Rows - 1, 4).text = total(1)
        Grid1.Cell(Grid1.Rows - 1, 5).text = total(2)
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 5).text = "DIFERENCIA"
        
        Grid1.Cell(Grid1.Rows - 1, 6).text = total(1) - total(2)
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0


End Sub
Private Sub Command5_Click()
Dim titulo As String
If Option1.Value = True Then titulo = "TODAS"
If Option2.Value = True Then titulo = "INCONSISTENCIAS"

Call CABEZAS2("LISTA COMPARACION TRANSBANK " + ComboLOCAL.text, titulo)

Grid1.PageSetup.BlackAndWhite = True




Grid1.PrintPreview


End Sub

Private Sub Command6_Click()
If Option1.Value = True Or Option4.Value = True Then
LEERCARTOLAS
End If
If Option2.Value = True Or Option3.Value = True Then
Call Command4_Click
End If
If Option5.Value = True Then
Call resumen_credito
End If
If Option6.Value = True Then
LEERCARTOLAS_pagos
End If
If Option9.Value = True Then
LEERCARTOLAS_pagos_inconsistencias

End If
If Option7.Value = True Or Option8.Value = True Then
LEERCARTOLAS_pagos_conciliados

End If

End Sub

Private Sub Command7_Click()

If Verifica_Permiso(Me.Caption, "elimina") = True Then
Rem     Call eliminacartolas(Format(elidesde, "yyyy-mm-dd"), Format(elihasta, "yyyy-mm-dd"), dato1.text + dato2.text + dato3.text)
End If


frameelimina.Visible = False

End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub Command9_Click()
frameelimina.Visible = False
End Sub



Private Sub Dir1_Change()
Dir1.path = "U:"
File1.path = "U:\MULTICAJA_FILES\"
File1.Pattern = "*.XLS"


End Sub

Private Sub Drive1_Change()
Dir1.path = "U:"
File1.path = "U:\MULTICAJA_FILES\"
File1.Pattern = "*.XLS;*.csv"

End Sub

Private Sub ELIMINA_Click()
frameelimina.Visible = True

End Sub

Private Sub File1_Click()
k = File1.ListIndex

ARCHIVO.text = File1.List(k)

End Sub

Private Sub Form_Load()

Dim cadena As String

    cadena = "net use u: \\192.168.4.6\c /delete"
    Shell cadena
    cadena = "NET START " & Chr(34) & "SERVIDOR" & Chr(34)
    Shell cadena
    cadena = "NET START " & Chr(34) & "examinador de equipos" & Chr(34)
    Shell cadena
    cadena = "net use u: \\192.168.4.6\c "
    Shell cadena
    cadena = "net use /persistent: yes"
    Shell cadena

CENTRAR Me
    Call Conectar_BD
    sc = 0
CARGAGRILLA
CARGATXT.Visible = False

desdefecha.Caption = Format(DateAdd("d", -3, fechasistema), "dd-mm-yyyy")
hastafecha.Caption = desdefecha.Caption


LEErlocales

    

End Sub



Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT terminal,local "
        csql.sql = csql.sql + "FROM " + clientesistema + "conta.rc_multicaja_terminales "
        csql.sql = csql.sql + "ORDER BY local,terminal "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + "-" + leernombrelocal(resultados(1)))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.AddItem ("99     - todos ")
        ComboLOCAL.text = "99      - todos "
        End If
      
        
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


Private Sub lblhistorico_Click(Index As Integer)

End Sub







Sub imprimir()
    
   
End Sub
Sub grilla()
    
End Sub
Sub CABEZA()
    

End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        csql.sql = csql.sql + "FROM cuentasdelmayor where  año='" + Format(fechasistema, "yyyy") + "' "
        csql.sql = csql.sql + " order by codigo"
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                
                dato(1) = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4): colu(1) = 15: tipodato(1) = "s"
                dato(2) = resultados(1): colu(2) = 52: tipodato(2) = "s"
                dato(3) = resultados(2) + " " + DOCU$(Val(resultados(2)))
                dato(4) = resultados(3)
                dato(5) = resultados(4)
                dato(6) = resultados(5) + " " + DOCU2$(Val(resultados(5)))
                colu(3) = 15: tipodato(3) = "s"
                colu(4) = 3: tipodato(4) = "s"
                colu(5) = 20: tipodato(5) = "s"
                colu(6) = 20: tipodato(6) = "s"
                 cancolu = 6
                grilla
                resultados.MoveNext
            Wend
            resultados.Close
            
            Set resultados = Nothing

        End If
    

End Sub

Sub revisa_transbank(fecha, comercio)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim DIFE As Double
    
'select me.codigo,me.nombre,'2014-01-02',
'
'ifnull((select sum(monto) from eltit_conta.rc_multicaja as rc inner join eltit_conta.rc_multicaja_terminales as mte on mte.terminal=rc.terminal where  fecha like '2014-01-02' and te.local=mte.local and (respuesta='Venta Confirmada' or respuesta='Venta Aprobada' or (transacciones='Confirmacion Corona' and respuesta='')) group by te.local ),0) as multicard,
'ifnull((select sum(monto) from eltit_teso.rc_tarjetascasascomerciales as tb where fecha = '2014-01-02' and local=te.local group by fecha ),0) as tesoreria from eltit_gestion.g_maestroempresas as me inner join eltit_conta.rc_multicaja_terminales as te on te.local=me.codigo
'group by te.local;

Set csql.ActiveConnection = contadb
csql.sql = "select me.codigo,me.nombre,'" + Format(fecha, "yyyy-mm-dd") + "', "
csql.sql = csql.sql + "ifnull((select sum(monto) from eltit_conta.rc_multicaja as rc inner join eltit_conta.rc_multicaja_terminales as mte on mte.terminal=rc.terminal where  fecha like '" & Format(fecha, "yyyy-mm-dd") & "' and te.local=mte.local and (respuesta='Venta Confirmada' or respuesta='Venta Aprobada' or (transacciones='Confirmacion Corona' and respuesta='')) group by te.local ),0) as multicard,"
csql.sql = csql.sql + "ifnull((select sum(monto) from eltit_teso.rc_tarjetascasascomerciales as tb where fecha = '" & Format(fecha, "yyyy-mm-dd") & "' and local=te.local group by fecha ),0) as tesoreria "
csql.sql = csql.sql + "from eltit_gestion.g_maestroempresas as me inner join eltit_conta.rc_multicaja_terminales as te on te.local=me.codigo "

'csql.sql = csql.sql + " ifnull((select sum(monto) from eltit_conta.rc_multicaja as rc  where  fecha like '" + Format(FECHA, "yyyy-mm-dd") + "'  group by fecha ),0) as tbk_debito,"
'csql.sql = csql.sql + "ifnull((select sum(monto) from eltit_teso.rc_tarjetascasascomerciales as tb where fecha = '" + Format(FECHA, "yyyy-mm-dd") + "' group by fecha ),0) as teso_debito"
''
''/*+ifnull((select sum(monto) from eltit_teso.rc_tarjetacredito as tc left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tc.local where me2.codigocomerciotbk=rc.loc and fecha=rc.fecha_venta group by me2.codigocomerciotbk ),0) as total */
'csql.sql = csql.sql + " from eltit_gestion.g_maestroempresas as me where codigo='08' "

If Mid(comercio, 1, 2) <> "99" Then

csql.sql = csql.sql + " and terminal='" + comercio + "' "
End If
csql.sql = csql.sql + "group by te.local "

'    select rc.loc,
'mid((select nombre from eltit_gestion.g_maestroempresas where codigocomerciotbk=rc.loc limit 0,1),1,21) AS nombre ,
'rc.fecha_venta,
'sum(monto_afecto+monto_exento),
'ifnull((select sum(monto) from eltit_teso.rc_tarjetadebito as tb left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tb.local where me2.codigocomerciotbk=loc and fecha = rc.fecha_venta group by me2.codigocomerciotbk ),0)+ifnull((select sum(monto) from eltit_teso.rc_tarjetacredito as tc left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tc.local where me2.codigocomerciotbk=rc.loc and fecha=rc.fecha_venta group by me2.codigocomerciotbk ),0) as totalteso
'from eltit_conta.rc_transbank as rc
'where fecha_venta like '2014-01-02' group by /*loc,*/fecha_venta order by loc;

        
        
        Grid1.AutoRedraw = False
        
        
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                                
                If Option3.Value = False Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
                Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
                Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
                Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
                End If
                
                total(1) = total(1) + resultados(3)
                total(2) = total(2) + resultados(4)
               
                
                DIFE = (resultados(4) - resultados(3))
                saldo = saldo + DIFE
'
                If resultados(4) > resultados(3) Then
                If Option3.Value = False Then
                Grid1.Cell(Grid1.Rows - 1, 6).text = (resultados(4) - resultados(3))
                End If
                total(3) = total(3) + (resultados(4) - resultados(3))

                Else
                If Option3.Value = False Then
                Grid1.Cell(Grid1.Rows - 1, 7).text = (resultados(3) - resultados(4))
                End If
                total(4) = total(4) + (resultados(3) - resultados(4))
'
                End If
'
'
                If Option3.Value = False Then
                Grid1.Cell(Grid1.Rows - 1, 8).text = saldo * -1
                End If
                
                'End If
                resultados.MoveNext
            Wend
            
            resultados.Close
            
            Set resultados = Nothing
        


        End If
Grid1.AutoRedraw = True
Grid1.Refresh
    

End Sub

Sub revisa_transbank_resumen(desde, hasta, comercio)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim DIFE As Double
    
'select me.codigo,me.nombre,'2014-01-02',
'
'ifnull((select sum(monto) from eltit_conta.rc_multicaja as rc inner join eltit_conta.rc_multicaja_terminales as mte on mte.terminal=rc.terminal where  fecha like '2014-01-02' and te.local=mte.local and (respuesta='Venta Confirmada' or respuesta='Venta Aprobada' or (transacciones='Confirmacion Corona' and respuesta='')) group by te.local ),0) as multicard,
'ifnull((select sum(monto) from eltit_teso.rc_tarjetascasascomerciales as tb where fecha = '2014-01-02' and local=te.local group by fecha ),0) as tesoreria from eltit_gestion.g_maestroempresas as me inner join eltit_conta.rc_multicaja_terminales as te on te.local=me.codigo
'group by te.local;

Set csql.ActiveConnection = contadb
csql.sql = "select me.codigo,me.nombre,'" + Format(hasta, "yyyy-mm-dd") + "', "
csql.sql = csql.sql + "ifnull((select sum(monto) from eltit_conta.rc_multicaja as rc inner join eltit_conta.rc_multicaja_terminales as mte on mte.terminal=rc.terminal where  fecha between '" & Format(desde, "yyyy-mm-dd") & "' and '" & Format(hasta, "yyyy-mm-dd") & "' and te.local=mte.local and (respuesta='Venta Confirmada' or respuesta='Venta Aprobada' or (transacciones='Confirmacion Corona' and respuesta='')) group by te.local ),0) as multicard,"
csql.sql = csql.sql + "ifnull((select sum(monto) from eltit_teso.rc_tarjetascasascomerciales as tb where fecha between '" & Format(desde, "yyyy-mm-dd") & "' and '" & Format(hasta, "yyyy-mm-dd") & "' and local=te.local group by te.local ),0) as tesoreria "
csql.sql = csql.sql + "from eltit_gestion.g_maestroempresas as me inner join eltit_conta.rc_multicaja_terminales as te on te.local=me.codigo "

'csql.sql = csql.sql + " ifnull((select sum(monto) from eltit_conta.rc_multicaja as rc  where  fecha like '" + Format(FECHA, "yyyy-mm-dd") + "'  group by fecha ),0) as tbk_debito,"
'csql.sql = csql.sql + "ifnull((select sum(monto) from eltit_teso.rc_tarjetascasascomerciales as tb where fecha = '" + Format(FECHA, "yyyy-mm-dd") + "' group by fecha ),0) as teso_debito"
''
''/*+ifnull((select sum(monto) from eltit_teso.rc_tarjetacredito as tc left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tc.local where me2.codigocomerciotbk=rc.loc and fecha=rc.fecha_venta group by me2.codigocomerciotbk ),0) as total */
'csql.sql = csql.sql + " from eltit_gestion.g_maestroempresas as me where codigo='08' "

If Mid(comercio, 1, 2) <> "99" Then

csql.sql = csql.sql + " and terminal='" + comercio + "' "
End If
csql.sql = csql.sql + "group by te.local "

'    select rc.loc,
'mid((select nombre from eltit_gestion.g_maestroempresas where codigocomerciotbk=rc.loc limit 0,1),1,21) AS nombre ,
'rc.fecha_venta,
'sum(monto_afecto+monto_exento),
'ifnull((select sum(monto) from eltit_teso.rc_tarjetadebito as tb left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tb.local where me2.codigocomerciotbk=loc and fecha = rc.fecha_venta group by me2.codigocomerciotbk ),0)+ifnull((select sum(monto) from eltit_teso.rc_tarjetacredito as tc left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tc.local where me2.codigocomerciotbk=rc.loc and fecha=rc.fecha_venta group by me2.codigocomerciotbk ),0) as totalteso
'from eltit_conta.rc_transbank as rc
'where fecha_venta like '2014-01-02' group by /*loc,*/fecha_venta order by loc;

        saldo = 0
        
        Grid1.AutoRedraw = False
        
        
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                                
                If Option3.Value = False Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
                Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
                Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
                Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
                End If
                
                total(1) = total(1) + resultados(3)
                total(2) = total(2) + resultados(4)
               
                
                DIFE = (resultados(4) - resultados(3))
                saldo = saldo + DIFE
'
                If resultados(4) > resultados(3) Then
                If Option3.Value = False Then
                Grid1.Cell(Grid1.Rows - 1, 6).text = (resultados(4) - resultados(3))
                End If
                total(3) = total(3) + (resultados(4) - resultados(3))

                Else
                If Option3.Value = False Then
                Grid1.Cell(Grid1.Rows - 1, 7).text = (resultados(3) - resultados(4))
                End If
                total(4) = total(4) + (resultados(3) - resultados(4))
'
                End If
'
'
                If Option3.Value = False Then
                Grid1.Cell(Grid1.Rows - 1, 8).text = saldo * -1
                End If
                
                'End If
                resultados.MoveNext
            Wend
            
            resultados.Close
            
            Set resultados = Nothing
        


        End If
Grid1.AutoRedraw = True
Grid1.Refresh
    

End Sub


'Sub revisa_transbank_resumen(desde, hasta, comercio)
'    Dim resultados As rdoResultset
'    Dim csql As New rdoQuery
'    Dim rut As String
'    Dim dife As Double
'
'
'        Set csql.ActiveConnection = contadb
'csql.sql = "select me.codigo,me.nombre,'" + Format(hasta, "yyyy-mm-dd") + "', "
'csql.sql = csql.sql + " ifnull((select sum(monto_afecto)+sum(monto_exento) from eltit_conta.rc_transbank as rc  where  loc=codigocomerciotbk and fecha_venta between '" + Format(desde, "yyyy-mm-dd") + "' and '" + Format(hasta, "yyyy-mm-dd") + "' and tipo_tarjeta='DB' ),0) as tbk_debito,"
'csql.sql = csql.sql + "ifnull((select sum(monto) from eltit_teso.rc_tarjetadebito as tb left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tb.local where me2.codigocomerciotbk=me.codigocomerciotbk and fecha between '" + Format(desde, "yyyy-mm-dd") + "' and '" + Format(hasta, "yyyy-mm-dd") + "' group by me2.codigocomerciotbk ),0) as teso_debito,"
'csql.sql = csql.sql + "ifnull((select sum(monto_afecto)+sum(monto_exento) from eltit_conta.rc_transbank as rc  where  loc=codigocomerciotbk and fecha_venta between '" + Format(desde, "yyyy-mm-dd") + "' and '" + Format(hasta, "yyyy-mm-dd") + "'  and tipo_tarjeta<>'DB' ),0) as tbk_credito,"
'csql.sql = csql.sql + "ifnull((select sum(monto) from eltit_teso.rc_tarjetacredito as tb left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tb.local where me2.codigocomerciotbk=me.codigocomerciotbk and fecha between '" + Format(desde, "yyyy-mm-dd") + "' and '" + Format(hasta, "yyyy-mm-dd") + "'  group by me2.codigocomerciotbk ),0) as teso_credito"
''
''/*+ifnull((select sum(monto) from eltit_teso.rc_tarjetacredito as tc left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tc.local where me2.codigocomerciotbk=rc.loc and fecha=rc.fecha_venta group by me2.codigocomerciotbk ),0) as total */
'csql.sql = csql.sql + " from eltit_gestion.g_maestroempresas as me where codigocomerciotbk<>'' "
'
'If Mid(comercio, 1, 2) <> "99" Then
'
'csql.sql = csql.sql + " and codigocomerciotbk='" + comercio + "' "
'End If
'csql.sql = csql.sql + "group by codigocomerciotbk "
'
''    select rc.loc,
''mid((select nombre from eltit_gestion.g_maestroempresas where codigocomerciotbk=rc.loc limit 0,1),1,21) AS nombre ,
''rc.fecha_venta,
''sum(monto_afecto+monto_exento),
''ifnull((select sum(monto) from eltit_teso.rc_tarjetadebito as tb left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tb.local where me2.codigocomerciotbk=loc and fecha = rc.fecha_venta group by me2.codigocomerciotbk ),0)+ifnull((select sum(monto) from eltit_teso.rc_tarjetacredito as tc left join eltit_gestion.g_maestroempresas as me2 on me2.codigo=tc.local where me2.codigocomerciotbk=rc.loc and fecha=rc.fecha_venta group by me2.codigocomerciotbk ),0) as totalteso
''from eltit_conta.rc_transbank as rc
''where fecha_venta like '2014-01-02' group by /*loc,*/fecha_venta order by loc;
'
'
'
'        Grid1.AutoRedraw = False
'
'
'        csql.Execute
'        If csql.RowsAffected > 0 Then
'            Set resultados = csql.OpenResultset
'            While Not resultados.EOF
'
'                If Option3.Value = False Then
'                Grid1.Rows = Grid1.Rows + 1
'                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
'                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
'                Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
'                Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
'                Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
'                Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
'                Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(6)
'                End If
'
'                total(1) = total(1) + resultados(3)
'                total(2) = total(2) + resultados(4)
'                total(3) = total(3) + resultados(5)
'                total(4) = total(4) + resultados(6)
'
'
'                dife = (resultados(4) + resultados(6)) - (resultados(3) + resultados(5))
'                saldo = saldo + dife
''
'                If resultados(4) + resultados(6) > resultados(3) + resultados(5) Then
'                If Option3.Value = False Then
'                Grid1.Cell(Grid1.Rows - 1, 8).text = (resultados(4) + resultados(6)) - (resultados(3) + resultados(5))
'                End If
'                total(5) = total(5) + (resultados(4) + resultados(6)) - (resultados(3) + resultados(5))
'
'                Else
'                If Option3.Value = False Then
'                Grid1.Cell(Grid1.Rows - 1, 9).text = (resultados(3) + resultados(5)) - (resultados(4) + resultados(6))
'                End If
'                total(6) = total(6) + (resultados(3) + resultados(5)) - (resultados(4) + resultados(6))
''
'                End If
''
''
'                If Option3.Value = False Then
'                Grid1.Cell(Grid1.Rows - 1, 10).text = saldo * -1
'                End If
'                'End If
'                resultados.MoveNext
'            Wend
'
'            resultados.Close
'
'            Set resultados = Nothing
'
'
'
'        End If
'Grid1.AutoRedraw = True
'Grid1.Refresh
'
'
'End Sub


Private Sub opciones_GotFocus()



End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 14)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "HORA"
    FORMATOGRILLA(1, 3) = "TERMINAL"
    FORMATOGRILLA(1, 4) = "CANAL"
    FORMATOGRILLA(1, 5) = "TRANSACCION"
    FORMATOGRILLA(1, 6) = "ESTADO"
    FORMATOGRILLA(1, 7) = "RESPUESTA"
    FORMATOGRILLA(1, 8) = "DETALLE"
    FORMATOGRILLA(1, 9) = "CODIGO_MC"
    FORMATOGRILLA(1, 10) = "MONTO"
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "4"
    FORMATOGRILLA(2, 3) = "30"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "15"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "15"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "8"
    FORMATOGRILLA(2, 10) = "8"
    FORMATOGRILLA(2, 11) = "8"
    FORMATOGRILLA(2, 12) = "8"
    FORMATOGRILLA(2, 13) = "8"
    FORMATOGRILLA(2, 14) = "8"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "D"
    FORMATOGRILLA(3, 14) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 10) = "###,###,###,##0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    
    Grid1.Cols = 11
    Grid1.Rows = 1
    
     'infogrilla.grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
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
    Dim o As Double
    
    For o = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, o).text = FORMATOGRILLA(1, o)
        Grid1.Column(o).Width = Val(FORMATOGRILLA(2, o)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(o).MaxLength = Val(FORMATOGRILLA(2, o))
        Grid1.Column(o).FormatString = FORMATOGRILLA(4, o)
        Grid1.Column(o).Locked = FORMATOGRILLA(5, o)
        If FORMATOGRILLA(3, o) = "N" Then Grid1.Column(o).Alignment = cellRightCenter
        If FORMATOGRILLA(3, o) = "D" Then Grid1.Column(o).CellType = cellCalendar
        If FORMATOGRILLA(3, o) = "S" Then Grid1.Column(o).CellType = cellTextBox
        
    Next o
End Sub

Sub CARGAGRILLA2()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 12)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "LOCAL"
    FORMATOGRILLA(1, 3) = "FECHA"
    FORMATOGRILLA(1, 4) = "MULTICAJA"
    FORMATOGRILLA(1, 5) = "TESORERIA"
    FORMATOGRILLA(1, 6) = "DIF.MULTI"
    FORMATOGRILLA(1, 7) = "DIF.CAJA"
    FORMATOGRILLA(1, 8) = "SALDO"
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "20"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"
    FORMATOGRILLA(2, 10) = "10"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "D"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 4) = "###,###,###,##0"
    FORMATOGRILLA(4, 5) = "###,###,###,##0"
    FORMATOGRILLA(4, 6) = "###,###,###,##0"
    FORMATOGRILLA(4, 7) = "###,###,###,##0"
    FORMATOGRILLA(4, 8) = "###,###,###,##0"
    FORMATOGRILLA(4, 9) = "###,###,###,##0"
    FORMATOGRILLA(4, 10) = "###,###,###,##0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    
    Grid1.Cols = 9
    Grid1.Rows = 1
    
     'infogrilla.grid1.AllowUserResizing = False
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
    Dim o As Double
    
    For o = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, o).text = FORMATOGRILLA(1, o)
        Grid1.Column(o).Width = Val(FORMATOGRILLA(2, o)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(o).MaxLength = Val(FORMATOGRILLA(2, o))
        Grid1.Column(o).FormatString = FORMATOGRILLA(4, o)
        Grid1.Column(o).Locked = FORMATOGRILLA(5, o)
        If FORMATOGRILLA(3, o) = "N" Then Grid1.Column(o).Alignment = cellRightCenter: Grid1.Column(o).CellType = cellDefault
        If FORMATOGRILLA(3, o) = "D" Then Grid1.Column(o).CellType = cellCalendar
        If FORMATOGRILLA(3, o) = "S" Then Grid1.Column(o).CellType = cellTextBox
    Next o
End Sub



Private Sub LEERCARTOLAS()

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim total2 As Double
    Dim estado As String
    CARGAGRILLA
    LINEA = 0
 Call consolidatarjetas(desdefecha.Caption, hastafecha.Caption)
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT * "
        Rem - select * from rc_multicaja where (respuesta='Venta Confirmada' or respuesta='Venta Aprobada') and FECHA='2014-01-02';
        csql.sql = csql.sql + "FROM rc_multicaja where  fecha>='" + Format(desdefecha.Caption, "yyyy-mm-dd") + "' and fecha <='" + Format(hastafecha.Caption, "yyyy-mm-dd") + "' "
        csql.sql = csql.sql + " and (respuesta='Venta Confirmada' or respuesta='Venta Aprobada' or (transacciones='Confirmacion Corona' and respuesta='')) "
        If Val(monto.text) <> 0 Then
        csql.sql = csql.sql + " and monto=" & Val(monto.text)
        
        End If
        If Mid(ComboLOCAL.text, 1, 2) <> "99" Then
        csql.sql = csql.sql + " and terminal='" + Mid(ComboLOCAL.text, 1, 8) + "' "
        End If
                
        
        csql.sql = csql.sql + " order by monto "
        csql.Execute
        total = 0
        total2 = 0
        
        Grid1.Rows = 1
Grid1.AutoRedraw = False

LINEA = 1
        
        If csql.RowsAffected > 0 Then
barra.Max = csql.RowsAffected + 1
barra.Value = 0

        Dim siesta As Boolean
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
                    
                    
                    If LEERTARJETAENTESO(resultados(0), resultados(9)) = True Then
                    siesta = True
                    Else
                    siesta = False
                    End If
                    If Option4.Value = True And siesta = True Then GoTo PASO:
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
                    If siesta = True Then
                    Grid1.Range(LINEA, 1, LINEA, 10).BackColor = vbGreen
                    Else
                    Grid1.Range(LINEA, 1, LINEA, 10).BackColor = vbRed
                    
                    End If
                    
                    Grid1.Cell(LINEA, 1).text = resultados(0)
                    Grid1.Cell(LINEA, 2).text = resultados(1)
                    Grid1.Cell(LINEA, 3).text = resultados(2) + " " + leernombrelocal(Leerlocalterminal(resultados(2)))
                    Grid1.Cell(LINEA, 4).text = resultados(3)
                    Grid1.Cell(LINEA, 5).text = resultados(4)
                    Grid1.Cell(LINEA, 6).text = resultados(5)
                    Grid1.Cell(LINEA, 7).text = resultados(6)
                    Grid1.Cell(LINEA, 8).text = resultados(7)
                    Grid1.Cell(LINEA, 9).text = resultados(8)
                    Grid1.Cell(LINEA, 10).text = resultados(9)
                    
             
             
        
       
             
PASO:
          
          barra.Value = barra.Value + 1
          barra.Refresh

          resultados.MoveNext

         Wend
End If
Grid1.AutoRedraw = True
Grid1.Refresh
If Val(monto.text) = 0 And Mid(ComboLOCAL.text, 1, 2) = "99" Then
LEERsobrantes_teso
End If


End Sub
Private Sub LEERsobrantes_teso()

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim total2 As Double
    Dim estado As String
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT * "
        csql.sql = csql.sql + "FROM eltit_teso.rc_tarjetasbancarias where  auditoria='' "
        csql.Execute
        total = 0
        total2 = 0
        
Grid1.AutoRedraw = False

        If csql.RowsAffected > 0 Then
barra.Max = csql.RowsAffected + 1
barra.Value = 0

        Dim siesta As Boolean
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
                    
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
                    Grid1.Range(LINEA, 1, LINEA, 10).BackColor = vbYellow
                    Grid1.Cell(LINEA, 1).text = resultados(0)
                    Grid1.Column(2).Locked = False
                    Grid1.Column(3).Locked = False
                    Grid1.Range(LINEA, 2, LINEA, 3).Merge
                    Grid1.Cell(LINEA, 2).text = LeerNombrecajera(resultados(1))
                    Grid1.Column(2).Locked = True
                    Grid1.Column(3).Locked = True
                    
                    Grid1.Cell(LINEA, 5).text = leernombrelocal(resultados(2))
                    Grid1.Cell(LINEA, 4).text = resultados(3)
                    Rem Grid1.Cell(linea, 5).text = resultados(4)
                    Grid1.Cell(LINEA, 7).text = LeerNombreTARJETA(resultados(5))
                    Grid1.Cell(LINEA, 10).text = resultados(6)
                    Rem Grid1.Cell(linea, 8).text = resultados(7)
                    
        
       
             
          
          barra.Value = barra.Value + 1
          barra.Refresh

          resultados.MoveNext

         Wend
End If
Grid1.AutoRedraw = True
Grid1.Refresh



End Sub



Private Sub monto_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    

End Sub

Private Sub Option1_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If
End Sub

Private Sub Option2_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If
    
End Sub
Private Sub Option3_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If
    
End Sub
Private Sub Option4_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If

End Sub

Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = MES
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub

Public Function LEERULTIMOFOLIO(tipo) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from movimientoscontables where mes = '" & Format(MES, "00") & "' AND año = '" & año & "' and tipo='" + tipo + "' "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIO = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function

Public Function LEERcuentacontable(codigo) As Boolean


    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = conta

            csql.sql = "select * from cartolasbancarias_codigoscontables where codigo='" + codigo + "' and codigocontable<>'00000000' "
            
            csql.Execute
    LEERcuentacontable = False
    If csql.RowsAffected > 0 Then
  
    
    Set resultados = csql.OpenResultset
        banco_cuenta = resultados(3)
        banco_dh = resultados(4)
        banco_glosa = resultados(2)
        banco_glosa2 = resultados(5)
        LEERcuentacontable = True
            
    End If
    
End Function

Public Sub LEERcuentacontabilizada(codigo, fecha, monto, DH)


    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    If DH = "A" Then DH = "H" Else DH = "D"
        Set csql.ActiveConnection = contadb

            csql.sql = "select fecha,monto from movimientoscontables where codigocuenta='" & codigo & "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' and monto='" & monto & "' and dh='" & DH & "' "

            csql.Execute
    conta_fecha = ""
    conta_monto = ""
    conta_glosa = "NO CONTABILIZADO"
    If csql.RowsAffected > 0 Then


    Set resultados = csql.OpenResultset
        conta_fecha = resultados(0)
        conta_monto = resultados(1)
        conta_glosa = "CONTABILIZADO"
        
    End If

End Sub


Sub CABEZAS2(titulo, titulo2)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear
Grid1.PageSetup.Orientation = cellLandscape



Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo2
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    
With Grid1.PageSetup
        
        .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub

Sub eliminacartolas(fe1, fe2, cta)
    campos(0, 2) = "cartolasbancarias"
    condicion = "fecha BETWEEN '" & fe1 & "' AND '" & fe2 & "' AND cuenta='" & cta & "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
End Sub


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Public Function LEERTARJETAENTESO(fecha, monto) As Boolean



    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim csql2 As New rdoQuery
        
        Set csql.ActiveConnection = contadb

            csql.sql = "select * from eltit_teso.rc_tarjetasbancarias where monto='" & monto & "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' and auditoria='' limit 0,1"

            csql.Execute
            
    LEERTARJETAENTESO = False
    If csql.RowsAffected > 0 Then


    Set resultados = csql.OpenResultset
         LEERTARJETAENTESO = True
            Set csql2.ActiveConnection = contadb
            csql2.sql = "update eltit_teso.rc_tarjetasbancarias set auditoria='1' where fecha='" & Format(resultados(0), "yyyy-mm-dd") & "' and cajera='" & resultados(1) & "' and local='" & resultados(2) & "' and  caja='" & resultados(3) & "' and linea ='" & resultados(4) & "' and monto='" & resultados(6) & "' " '"

            csql2.Execute
            
    End If

End Function

Public Sub consolidatarjetas(desde, hasta)



    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = contadb
            csql.sql = "truncate table eltit_teso.rc_tarjetasbancarias "
            csql.Execute
            csql.sql = "insert ignore into eltit_teso.rc_tarjetasbancarias (fecha,cajera,local,caja,linea,tipo,monto,auditoria)"
            csql.sql = csql.sql + "select fecha,cajera,local,caja,linea,tipo,monto,'' from eltit_teso.rc_tarjetascasascomerciales where fecha between '" & Format(desde, "yyyy-mm-dd") & "' and '" & Format(hasta, "yyyy-mm-dd") & "' "
            

            csql.Execute
    
End Sub

Private Sub Option5_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If

End Sub
Private Sub Option6_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If

End Sub


Private Sub Option7_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If

End Sub

Private Sub Option8_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If

End Sub
Private Sub Option9_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If

End Sub
Private Sub Option10_Click()
If MsgBox("desea regenerar informes ", vbYesNo) = vbYes Then
Command6_Click
End If

End Sub

Private Sub LEERCARTOLAS_pagos()

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim total2 As Double
    Dim estado As String
    Dim TOTAL1 As Double
    
    CARGAGRILLA_3
    LINEA = 0
Rem Call consolidatarjetas(desdefecha.Caption, hastafecha.Caption)
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT * "
        csql.sql = csql.sql + "FROM rc_multicaja_abonos where  fecha>='" + Format(desdefecha.Caption, "yyyy-mm-dd") + "' and fecha <='" + Format(hastafecha.Caption, "yyyy-mm-dd") + "' "
        If Val(monto.text) <> 0 Then
        csql.sql = csql.sql + " and monto=" & Val(monto.text)
        
        End If
        If Mid(ComboLOCAL.text, 1, 2) <> "99" Then
        csql.sql = csql.sql + " and terminal='" + Format(Mid(ComboLOCAL.text, 3, 6), "00000000") + "' "
        End If
        csql.sql = csql.sql + " order by emisor,fecha,monto "
        
        
        csql.Execute
        total = 0
        total2 = 0
        TOTAL1 = 0
        Grid1.Rows = 1
Grid1.AutoRedraw = False

LINEA = 1
        
        If csql.RowsAffected > 0 Then
barra.Max = csql.RowsAffected + 1
barra.Value = 0

        Dim siesta As Boolean
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
                    
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
                    
                    Grid1.Cell(LINEA, 1).text = resultados(0)
                    Grid1.Cell(LINEA, 2).text = resultados(1)
                    Grid1.Cell(LINEA, 3).text = resultados(2)
                    Grid1.Cell(LINEA, 4).text = resultados(3)
                    Grid1.Cell(LINEA, 5).text = resultados(4)
                    Grid1.Cell(LINEA, 6).text = resultados(5)
                    Grid1.Cell(LINEA, 7).text = resultados(6)
                    Grid1.Cell(LINEA, 8).text = resultados(7)
                    Grid1.Cell(LINEA, 9).text = resultados(8)
                    Grid1.Cell(LINEA, 10).text = resultados(9)
                    Grid1.Cell(LINEA, 11).text = resultados(10)
                    Grid1.Cell(LINEA, 12).text = resultados(11)
                    Grid1.Cell(LINEA, 13).text = resultados(12)
                    Grid1.Cell(LINEA, 14).text = resultados(13)
                    Grid1.Cell(LINEA, 15).text = resultados(14)
                    Grid1.Cell(LINEA, 16).text = resultados(15)
                    
             
        total = total + resultados(7)
        total2 = total2 + resultados(10)
        
        
       
             
PASO:
          
          barra.Value = barra.Value + 1
          barra.Refresh

          resultados.MoveNext

         Wend
End If
Grid1.AutoRedraw = True
Grid1.Refresh
Grid1.Rows = Grid1.Rows + 1
LINEA = Grid1.Rows - 1
                    
                    Grid1.Cell(LINEA, 8).text = total
                    Grid1.Cell(LINEA, 9).text = TOTAL1
                    
                    Grid1.Cell(LINEA, 11).text = total2
                    If total <> 0 Then
                    Grid1.Cell(LINEA, 12).text = Format((total2 / total * 100) / 1.19, "%##0.00")
                    End If
                    total = 0
                    total2 = 0
                    
                    


End Sub
Private Sub LEERCARTOLAS_pagos_inconsistencias()

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim total2 As Double
    Dim estado As String
    Dim TOTAL1 As Double
    Dim comer As String
    
    CARGAGRILLA_4
    LINEA = 0
        
        Set csql.ActiveConnection = contadb
        If Option11.Value = True Then
        csql.sql = "SELECT mt.local,tb.fecha,tb.hora,tb.respuesta,tb.transacciones,tb.codigo_mc,tb.terminal,sum(tb.monto) as venta,ifnull(sum(tba.monto),0) as abono "
'        csql.sql = csql.sql + "sum(tb.monto)-ifnull(Sum(tba.monto),0) as pendi, "
'        csql.sql = csql.sql + " ifnull(sum(tba.comision_emisor),0) ,"
'        csql.sql = csql.sql + "ifnull(max(tba.fecha),'0000-00-00') as ultimo_abono "
        csql.sql = csql.sql + "from eltit_conta.rc_multicaja as tb left join eltit_conta.rc_multicaja_abonos as tba "
        csql.sql = csql.sql + "on tb.monto=tba.monto and tb.fecha=tba.fecha and mid(tb.terminal,5,4)=mid(tba.terminal,5,4) "
        csql.sql = csql.sql + "inner join rc_multicaja_terminales as mt on mt.terminal=tb.terminal "
        csql.sql = csql.sql + " where  tb.fecha>='" + Format(desdefecha.Caption, "yyyy-mm-dd") + "' and tb.fecha <='" + Format(hastafecha.Caption, "yyyy-mm-dd") + "' "
        csql.sql = csql.sql + " and (tb.respuesta='Venta Confirmada' or tb.respuesta='Venta Aprobada' or (tb.transacciones='Confirmacion Corona' and tb.respuesta='')) "
        
        If Val(monto.text) <> 0 Then
        csql.sql = csql.sql + " and tb.monto_afecto+tb.monto_exento=" & Val(monto.text)
        End If
        If Mid(ComboLOCAL.text, 1, 2) <> "99" Then
        csql.sql = csql.sql + " and tb.loc='" + Mid(ComboLOCAL.text, 1, 8) + "' "
        End If
        If Option8.Value = False Then
        csql.sql = csql.sql + " group by mt.terminal,tb.fecha "
        End If
        
        csql.sql = csql.sql + "HAVING abono='0' "
        
        
        csql.sql = csql.sql + "order by mt.local,mt.terminal,tb.fecha "
        
        End If
        
        csql.Execute
        total = 0
        total2 = 0
        TOTAL1 = 0
        Grid1.Rows = 1
Grid1.AutoRedraw = False

LINEA = 1
        
        If csql.RowsAffected > 0 Then
barra.Max = csql.RowsAffected + 1
barra.Value = 0

        Dim siesta As Boolean
        Set resultados = csql.OpenResultset
        comer = resultados(0)
         While Not resultados.EOF
                            
                    If comer <> resultados(0) Then
                    
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
                    Grid1.Range(LINEA, 1, LINEA, 9).BackColor = vbGreen
                    Grid1.Cell(LINEA, 1).text = "TOTAL COMERCIO"
                    Grid1.Cell(LINEA, 8).text = total
                    total = 0
                    comer = resultados(0)
                    End If
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
                    
                    Grid1.Cell(LINEA, 1).text = resultados(0) + " " + leernombrelocal(resultados(0))
                    Grid1.Cell(LINEA, 2).text = resultados(1)
                    Grid1.Cell(LINEA, 3).text = resultados(2)
                    Grid1.Cell(LINEA, 4).text = resultados(3)
                    Grid1.Cell(LINEA, 5).text = resultados(4)
                    Grid1.Cell(LINEA, 6).text = resultados(5)
                    Grid1.Cell(LINEA, 7).text = resultados(6)
                    Grid1.Cell(LINEA, 8).text = resultados(7)
                    Grid1.Cell(LINEA, 9).text = resultados(8)

             
        total = total + resultados(7)
        total2 = total2 + resultados(7)
       
             
PASO:
          
          barra.Value = barra.Value + 1
          barra.Refresh

          resultados.MoveNext

         Wend
End If
Grid1.AutoRedraw = True
Grid1.Refresh
Grid1.Rows = Grid1.Rows + 1
LINEA = Grid1.Rows - 1
Grid1.Range(LINEA, 1, LINEA, 9).BackColor = vbGreen
                    Grid1.Cell(LINEA, 1).text = "TOTAL COMERCIO"
                    Grid1.Cell(LINEA, 8).text = total
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
                    Grid1.Cell(LINEA, 1).text = "TOTAL EMPRESA"
                    Grid1.Cell(LINEA, 8).text = total2
                    total = 0
                    total2 = 0
                    
                    


End Sub

Sub CARGAGRILLA_3()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 30)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "EMISOR"
    FORMATOGRILLA(1, 2) = "RUT_PRESTADOR"
    FORMATOGRILLA(1, 3) = "LUGAR"
    FORMATOGRILLA(1, 4) = "FECHA"
    FORMATOGRILLA(1, 5) = "CODIGO_MC"
    FORMATOGRILLA(1, 6) = "RUT_TITULAR"
    FORMATOGRILLA(1, 7) = "APLICACION"
    FORMATOGRILLA(1, 8) = "MONTO"
    FORMATOGRILLA(1, 9) = "CUOTAS"
    FORMATOGRILLA(1, 10) = "RESPUESTA"
    FORMATOGRILLA(1, 11) = "COM.EMISOR"
    FORMATOGRILLA(1, 12) = "COM.OPERADOR"
    FORMATOGRILLA(1, 13) = "COSTO_TRANSACCION"
    FORMATOGRILLA(1, 14) = "MONTO_PAGADO"
    FORMATOGRILLA(1, 15) = "TERMINAL"
    FORMATOGRILLA(1, 16) = "HORA"
    
    
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "4"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "3"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "5"
    FORMATOGRILLA(2, 10) = "0"
    FORMATOGRILLA(2, 11) = "8"
    FORMATOGRILLA(2, 12) = "8"
    FORMATOGRILLA(2, 13) = "8"
    FORMATOGRILLA(2, 14) = "8"
    FORMATOGRILLA(2, 15) = "8"
    FORMATOGRILLA(2, 16) = "8"
    FORMATOGRILLA(2, 17) = "8"
    FORMATOGRILLA(2, 18) = "8"
    FORMATOGRILLA(2, 19) = "8"
    FORMATOGRILLA(2, 20) = "8"
    FORMATOGRILLA(2, 21) = "8"
    FORMATOGRILLA(2, 22) = "8"
    FORMATOGRILLA(2, 23) = "8"
    FORMATOGRILLA(2, 24) = "8"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "D"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "N"
    FORMATOGRILLA(3, 15) = "S"
    FORMATOGRILLA(3, 16) = "S"
    FORMATOGRILLA(3, 17) = "N"
    FORMATOGRILLA(3, 18) = "N"
    FORMATOGRILLA(3, 19) = "N"
    FORMATOGRILLA(3, 20) = "N"
    FORMATOGRILLA(3, 21) = "N"
    FORMATOGRILLA(3, 22) = "S"
    FORMATOGRILLA(3, 23) = "N"
    FORMATOGRILLA(3, 24) = "N"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 8) = "###,###,###,##0"
    FORMATOGRILLA(4, 9) = "###,###,###,##0"
    FORMATOGRILLA(4, 11) = "###,###,###,##0"
    
    
    FORMATOGRILLA(4, 17) = "###,###,###,##0"
    FORMATOGRILLA(4, 18) = "###,###,###,##0"
    FORMATOGRILLA(4, 19) = "###,###,###,##0"
    FORMATOGRILLA(4, 20) = "###,###,###,##0"
    FORMATOGRILLA(4, 21) = "###,###,###,##0"
    FORMATOGRILLA(4, 22) = ""
    FORMATOGRILLA(4, 23) = ""
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    FORMATOGRILLA(5, 16) = "TRUE"
    FORMATOGRILLA(5, 17) = "TRUE"
    FORMATOGRILLA(5, 18) = "TRUE"
    FORMATOGRILLA(5, 19) = "TRUE"
    FORMATOGRILLA(5, 20) = "TRUE"
    FORMATOGRILLA(5, 21) = "TRUE"
    FORMATOGRILLA(5, 22) = "TRUE"
    FORMATOGRILLA(5, 23) = "TRUE"
    FORMATOGRILLA(5, 24) = "TRUE"
    
    
    Grid1.Cols = 17
    Grid1.Rows = 1
    
     'infogrilla.grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
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
    Dim o As Double
    
    For o = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, o).text = FORMATOGRILLA(1, o)
        Grid1.Column(o).Width = Val(FORMATOGRILLA(2, o)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(o).MaxLength = Val(FORMATOGRILLA(2, o))
        Grid1.Column(o).FormatString = FORMATOGRILLA(4, o)
        Grid1.Column(o).Locked = FORMATOGRILLA(5, o)
        If FORMATOGRILLA(3, o) = "N" Then Grid1.Column(o).Alignment = cellRightCenter
        If FORMATOGRILLA(3, o) = "D" Then Grid1.Column(o).CellType = cellCalendar
        If FORMATOGRILLA(3, o) = "S" Then Grid1.Column(o).CellType = cellDefault
        
    Next o
End Sub

Sub CARGAGRILLA_4()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 30)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "LOCAL"
    FORMATOGRILLA(1, 2) = "FECHA"
    FORMATOGRILLA(1, 3) = "HORA"
    FORMATOGRILLA(1, 4) = "TIPO TARJETA"
    FORMATOGRILLA(1, 5) = "TIPO VENTA"
    FORMATOGRILLA(1, 6) = "CODIGO"
    FORMATOGRILLA(1, 7) = "IDENTIFICADOR"
    FORMATOGRILLA(1, 8) = "MONTO"
    FORMATOGRILLA(1, 9) = "ABONO"
    
    
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "25"
    FORMATOGRILLA(2, 2) = "8"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "15"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "8"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "D"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 8) = "###,###,###,##0"
    FORMATOGRILLA(4, 9) = "###,###,###,##0"
    
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    
    
    Grid1.Cols = 10
    Grid1.Rows = 1
    
     'infogrilla.grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
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
    Dim o As Double
    
    For o = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, o).text = FORMATOGRILLA(1, o)
        Grid1.Column(o).Width = Val(FORMATOGRILLA(2, o)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(o).MaxLength = Val(FORMATOGRILLA(2, o))
        Grid1.Column(o).FormatString = FORMATOGRILLA(4, o)
        Grid1.Column(o).Locked = FORMATOGRILLA(5, o)
        If FORMATOGRILLA(3, o) = "N" Then Grid1.Column(o).Alignment = cellRightCenter
        If FORMATOGRILLA(3, o) = "D" Then Grid1.Column(o).CellType = cellCalendar
        If FORMATOGRILLA(3, o) = "S" Then Grid1.Column(o).CellType = cellTextBox
        If FORMATOGRILLA(3, o) = "N" Then Grid1.Column(o).CellType = cellTextBox
        
    Next o
End Sub


Public Function leernombrecomercio(comercio) As String

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = contadb
        
        tabla = "SELECT nombre "
        tabla = tabla & "FROM " & clientesistema & "gestion" & ".g_maestroempresas "
        tabla = tabla & "WHERE codigocomerciotbk= '" & comercio & "' "
        sql.sql = tabla
        sql.Execute
        
        leernombrecomercio = ""
        If sql.RowsAffected > 0 Then
        
            Set resultados = sql.OpenResultset
            leernombrecomercio = resultados(0)
            
            
        End If
    
    End Function
    
    

Private Sub LEERCARTOLAS_pagos_conciliados()

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim total2 As Double
    Dim estado As String
    Dim TOTAL1 As Double
    Dim comer As String
    Dim total4 As Double
    
    
    CARGAGRILLA_5
    LINEA = 0

        Set csql.ActiveConnection = contadb
        If Option11.Value = True Then
        csql.sql = "SELECT mt.local,'','',tb.fecha,sum(tb.monto) as venta,ifnull(sum(tba.monto),0) as abono,"
        csql.sql = csql.sql + "sum(tb.monto)-ifnull(Sum(tba.monto),0) as pendi, "
        csql.sql = csql.sql + " ifnull(sum(tba.comision_emisor),0) ,"
        csql.sql = csql.sql + "ifnull(max(tba.fecha),'0000-00-00') as ultimo_abono "
        csql.sql = csql.sql + "from eltit_conta.rc_multicaja as tb left join eltit_conta.rc_multicaja_abonos as tba "
        csql.sql = csql.sql + "on tb.monto=tba.monto and tb.fecha=tba.fecha and mid(tb.terminal,5,4)=mid(tba.terminal,5,4) "
        csql.sql = csql.sql + "inner join rc_multicaja_terminales as mt on mt.terminal=tb.terminal "
        csql.sql = csql.sql + " where  tb.fecha>='" + Format(desdefecha.Caption, "yyyy-mm-dd") + "' and tb.fecha <='" + Format(hastafecha.Caption, "yyyy-mm-dd") + "' "
        csql.sql = csql.sql + " and (tb.respuesta='Venta Confirmada' or tb.respuesta='Venta Aprobada' or (tb.transacciones='Confirmacion Corona' and tb.respuesta='')) "
        
        If Val(monto.text) <> 0 Then
        csql.sql = csql.sql + " and tb.monto_afecto+tb.monto_exento=" & Val(monto.text)
        End If
        If Mid(ComboLOCAL.text, 1, 2) <> "99" Then
        csql.sql = csql.sql + " and tb.loc='" + Mid(ComboLOCAL.text, 1, 8) + "' "
        End If
        If Option8.Value = False Then
        csql.sql = csql.sql + " group by mt.local,tb.fecha "
        End If
        
        
        
        
        csql.sql = csql.sql + "order by mt.local,tb.fecha "
        End If
        
        
        
        
        csql.Execute
        total = 0
        total2 = 0
        TOTAL1 = 0
        Grid1.Rows = 1
Grid1.AutoRedraw = False

LINEA = 1
        
        If csql.RowsAffected > 0 Then
barra.Max = csql.RowsAffected + 1
barra.Value = 0

        Dim siesta As Boolean
        Set resultados = csql.OpenResultset
        comer = resultados(0)
         While Not resultados.EOF
                            
'                    If comer <> resultados(0) Then
'                    Grid1.Range(linea, 1, linea, 8).BackColor = vbGreen
'                    Grid1.Cell(linea, 1).text = "TOTAL COMERCIO"
'                    Grid1.Cell(linea, 8).text = TOTAL
'                    TOTAL = 0
'                    comer = resultados(0)
'                    End If
'
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
                    Grid1.Cell(LINEA, 1).text = resultados(0)
                    If Option8.Value = True Then
                    Grid1.Cell(LINEA, 2).text = "TODOS DEL PERIODO"
                    
                    Else
                    Grid1.Cell(LINEA, 2).text = leernombrelocal(resultados(0))
                    
                    End If
                    
                    Grid1.Cell(LINEA, 3).text = resultados(2)
                    If Option8.Value = True Then
                    Grid1.Cell(LINEA, 4).CellType = cellDefault
                    Grid1.Cell(LINEA, 4).text = "periodo"
                    
                    Else
                    Grid1.Cell(LINEA, 4).CellType = cellCalendar
                    
                    Grid1.Cell(LINEA, 4).text = resultados(3)
                    End If
                    
                    Grid1.Cell(LINEA, 5).text = resultados(4)
                    Grid1.Cell(LINEA, 6).text = resultados(5)
                    Grid1.Cell(LINEA, 7).text = resultados(6)
                    Grid1.Cell(LINEA, 8).text = resultados(7)
                    Grid1.Cell(LINEA, 9).text = resultados(8)
                    
                    
             
        total = total + resultados(4)
        TOTAL1 = TOTAL1 + resultados(5)
       
        total2 = total2 + resultados(6)
       total4 = total4 + resultados(7)
             
PASO:
          
          barra.Value = barra.Value + 1
          barra.Refresh

          resultados.MoveNext

         Wend
End If
Grid1.AutoRedraw = True
Grid1.Refresh
'Grid1.Rows = Grid1.Rows + 1
'linea = Grid1.Rows - 1
'                    Grid1.Cell(linea, 1).text = "TOTAL COMERCIO"
'                    Grid1.Cell(linea, 8).text = TOTAL
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
                    Grid1.Range(LINEA, 1, LINEA, 8).BackColor = vbGreen
                    
                    Grid1.Cell(LINEA, 5).text = total
                    Grid1.Cell(LINEA, 6).text = TOTAL1
                    Grid1.Cell(LINEA, 7).text = total2
                    Grid1.Cell(LINEA, 8).text = total4

                    Grid1.Cell(LINEA, 1).text = "TOTAL EMPRESA"
'                    Grid1.Cell(linea, 8).text = total2
                    total = 0
                    total2 = 0
                    TOTAL1 = 0
                    


End Sub

Sub CARGAGRILLA_5()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 30)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "LOCAL"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = "TIPO"
    FORMATOGRILLA(1, 4) = "FECHA"
    FORMATOGRILLA(1, 5) = "VENTA"
    FORMATOGRILLA(1, 6) = "ABONO"
    FORMATOGRILLA(1, 7) = "PENDIENTE"
    FORMATOGRILLA(1, 8) = "COMISION"
    FORMATOGRILLA(1, 9) = "F.U.ABONO"
    
    
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "20"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "12"
    FORMATOGRILLA(2, 6) = "12"
    FORMATOGRILLA(2, 7) = "12"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "8"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "D"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "D"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 5) = "###,###,###,##0"
    FORMATOGRILLA(4, 6) = "###,###,###,##0"
    FORMATOGRILLA(4, 7) = "###,###,###,##0"
    FORMATOGRILLA(4, 8) = "###,###,###,##0"
    
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    
    
    Grid1.Cols = 10
    Grid1.Rows = 1
    
     'infogrilla.grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
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
    Dim o As Double
    
    For o = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, o).text = FORMATOGRILLA(1, o)
        Grid1.Column(o).Width = Val(FORMATOGRILLA(2, o)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(o).MaxLength = Val(FORMATOGRILLA(2, o))
        Grid1.Column(o).FormatString = FORMATOGRILLA(4, o)
        Grid1.Column(o).Locked = FORMATOGRILLA(5, o)
        If FORMATOGRILLA(3, o) = "N" Then Grid1.Column(o).Alignment = cellRightCenter
        If FORMATOGRILLA(3, o) = "D" Then Grid1.Column(o).CellType = cellCalendar
        If FORMATOGRILLA(3, o) = "S" Then Grid1.Column(o).CellType = cellTextBox
        If FORMATOGRILLA(3, o) = "N" Then Grid1.Column(o).CellType = cellDefault
        
    Next o
End Sub


Private Function LeerNombreTARJETA(codigo) As String
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "ventas.sv_tiposdepagoclientes "
    condicion = "codigo='" & codigo & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LeerNombreTARJETA = sqlconta.response(1, 3)
    Else
    LeerNombreTARJETA = ""
    End If
    
End Function

Private Function LeerNombrecajera(codigo) As String
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "ventas.sv_maestrocajeras "
    condicion = "rut like '%" & codigo & "%' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LeerNombrecajera = sqlconta.response(1, 3)
    Else
    LeerNombrecajera = ""
    End If
    
End Function

Private Function Leerlocalterminal(terminal) As String
    campos(0, 0) = "local"
    
    campos(1, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.rc_multicaja_terminales "
    condicion = "terminal like '%" & terminal & "%' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    Leerlocalterminal = sqlconta.response(0, 3)
    Else
    Leerlocalterminal = ""
    End If
    
End Function



