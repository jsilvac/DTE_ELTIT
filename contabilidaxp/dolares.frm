VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form banco09 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEMA AUTOMATICO BANCO SANTANDER"
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9855
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   17383
      BackColor       =   16744576
      Caption         =   "DATOS DEL CHEQUE"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FF8080&
         Caption         =   "Solo 943"
         Height          =   285
         Left            =   3600
         TabIndex        =   45
         Top             =   1080
         Width           =   1680
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Genera Informe"
         Height          =   375
         Left            =   3600
         TabIndex        =   43
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF8080&
         Caption         =   "Solo Depositos"
         Height          =   195
         Left            =   3600
         TabIndex        =   42
         Top             =   1800
         Width           =   1695
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
         Height          =   375
         Left            =   120
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1560
         Width           =   2355
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "No Contabilizados"
         Height          =   285
         Left            =   3600
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Inconsistencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   5760
         TabIndex        =   18
         Top             =   1560
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos Lo Movimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   5760
         TabIndex        =   17
         Top             =   1200
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TRASPASA CARTOLAS"
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
         Left            =   120
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1170
         Width           =   2355
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   8025
         Left            =   90
         TabIndex        =   5
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
         Begin XPFrame.FrameXp FrameQuickMenu 
            Height          =   615
            Left            =   10920
            TabIndex        =   39
            Top             =   7080
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
            Begin VB.CommandButton botonmisfavoritos 
               Caption         =   "Mis Favoritos"
               Height          =   255
               Left            =   1680
               TabIndex        =   41
               Top             =   280
               Width           =   1335
            End
            Begin VB.CommandButton botonmisaccesos 
               Caption         =   "Permisos Modulo"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   280
               Width           =   1455
            End
         End
         Begin XPFrame.FrameXp frameelimina 
            Height          =   1665
            Left            =   4560
            TabIndex        =   23
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
               TabIndex        =   38
               Top             =   1200
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker elihasta 
               Height          =   255
               Left            =   2520
               TabIndex        =   28
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
               TabIndex        =   27
               Top             =   840
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   450
               _Version        =   393216
               Format          =   75366401
               CurrentDate     =   40274
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Elimina cartolas"
               Height          =   375
               Left            =   240
               TabIndex        =   26
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
               TabIndex        =   25
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
               TabIndex        =   24
               Top             =   360
               Width           =   1935
            End
         End
         Begin XPFrame.FrameXp CARGATXT 
            Height          =   4200
            Left            =   2760
            TabIndex        =   29
            Top             =   1440
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
            Begin VB.CommandButton Command11 
               BackColor       =   &H00FF8080&
               Caption         =   "PROCESAR MASIVO"
               Height          =   465
               Left            =   3000
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   3480
               Width           =   2625
            End
            Begin VB.FileListBox File1 
               Height          =   2235
               Left            =   4230
               TabIndex        =   35
               Top             =   315
               Width           =   4275
            End
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   180
               TabIndex        =   34
               Top             =   315
               Width           =   3855
            End
            Begin VB.TextBox ARCHIVO 
               Height          =   285
               Left            =   4230
               TabIndex        =   33
               Top             =   3060
               Width           =   4275
            End
            Begin VB.DirListBox Dir1 
               Height          =   2565
               Left            =   180
               TabIndex        =   32
               Top             =   765
               Width           =   3855
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H00FF8080&
               Caption         =   "PROCESAR"
               Height          =   465
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   3465
               Width           =   2625
            End
            Begin VB.CommandButton Command3 
               BackColor       =   &H00FF8080&
               Caption         =   "RETORNO"
               Height          =   465
               Left            =   5880
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   3480
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
               TabIndex        =   36
               Top             =   2790
               Width           =   4290
            End
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Genera Asientos Contables"
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
            Left            =   7920
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   7320
            Width           =   2625
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
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   7290
            Width           =   2625
         End
         Begin MSComctlLib.ProgressBar barra 
            Height          =   195
            Left            =   0
            TabIndex        =   19
            Top             =   6840
            Width           =   13920
            _ExtentX        =   24553
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Conciliar Cheques"
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
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   7290
            Width           =   2625
         End
         Begin FlexCell.Grid Grid1 
            Height          =   6630
            Left            =   -30
            TabIndex        =   6
            Top             =   240
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   11695
            AllowUserSort   =   -1  'True
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
            DateFormat      =   2
         End
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00FFFFFF&
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
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "codigo"
         Text            =   "11"
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "12"
         Top             =   270
         Width           =   375
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1665
         Left            =   8775
         TabIndex        =   9
         Top             =   270
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
         Alignment       =   1
         Begin CoolButtons.cool_Button command8 
            Height          =   375
            Left            =   1440
            TabIndex        =   10
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "0002"
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox PIVOTE 
         Height          =   285
         Left            =   225
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   810
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LBLBANCO 
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
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   8385
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Banco"
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
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   1455
      End
   End
End
Attribute VB_Name = "banco09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BORRARARCHIVO As Boolean

Dim BENEFICIARIO As String
Dim banco_cuenta As String
Dim banco_glosa As String
Dim banco_dh As String
Dim conta_glosa As String
Dim conta_fecha As String
Dim conta_monto As String
Dim banco_glosa2 As String

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub




Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Private Sub Check1_Click()
 CARGAGRILLA
    
    Call LEERCARTOLAS(dato1.text + dato2.text + dato3.text)

End Sub

Private Sub Command1_Click()
CARGATXT.Visible = True

End Sub

Private Sub Command10_Click()
'    Call dato3_KeyPress(13)

   lblBanco.Caption = leerNombreMayor(dato1.text + dato2.text + Format(dato3.text, "0000"))
   CARGAGRILLA
    
    Call generacheques(empresaactiva, dato1.text + dato2.text + dato3.text, DateAdd("d", -30, Format(Date, "yyyy-mm-dd")))
    Call LEERCARTOLAS(dato1.text + dato2.text + dato3.text)
    
    If Verifica_Permiso(Me.Caption, "autoriza") = True Then
        ELIMINA.Enabled = True
    Else
        ELIMINA.Enabled = False
    End If

End Sub

Private Sub Command11_Click()
Dim o As Double
Dim origen As String
Dim destino As String


For o = 0 To File1.ListCount - 1
ARCHIVO.text = File1.List(o)
If UCase(Right(ARCHIVO.text, 3)) = "TXT" Then
        ARCHIVO.text = File1.List(o)
                
                CARGATXT.Visible = True

                TRASPASADATOS2
                If BORRARARCHIVO = True Then
                CARGATXT.Visible = False
                origen = "u:\cartolas\" + ARCHIVO.text
                destino = "u:\cartolas_usadas\" + ARCHIVO.text
                FileCopy origen, destino

                Kill origen
                End If
                
                Else
                MsgBox ("ESTE ARCHIVO NO ES UN ETBK")
                

End If
 
Next o


End Sub

Private Sub COMMAND2_Click()
If UCase(Right(ARCHIVO.text, 3)) = "TXT" Then
TRASPASADATOS
CARGATXT.Visible = False

Else
MsgBox ("ESTE ARCHIVO NO ES UN TXT")
End If

End Sub
Sub TRASPASADATOS()
Dim lin As Double

Close 20
Open File1.path + "\" + ARCHIVO.text For Input As #20
lin = 0
While EOF(20) = False
 
   
Line Input #20, varipaso
'If lin = 0 And Mid(varipaso, 56, 9) <> "SANTANDER" And Mid(varipaso, 56, 5) <> "BANCO" Then
'MsgBox ("ARCHIVO NO CORRESPONDE A CARTOLA BANCARIA")
'Exit Sub
'
'End If

If lin = 0 And InStr(1, varipaso, "BCO. SANTANDER") = 0 Then BORRARARCHIVO = False: Exit Sub
If lin = 0 And InStr(1, varipaso, "EXTRANJERADOLAR") = 0 Then BORRARARCHIVO = False: Exit Sub


If Mid(varipaso, 1, 2) = "00" Then
If Mid(varipaso, 4, 9) <> Mid(rutempresa, 1, 8) + Mid(rutempresa, 10, 1) Then
MsgBox ("ARCHIVO NO CORRESPONDE A DATOS DE LA EMPRESA ACTIVA " + Mid(varipaso, 3, 9) + "-" + Mid(varipaso, 11, 1) + " " + rutempresa)
Exit Sub
End If

End If

If Mid(varipaso, 1, 2) <> "00" Then

Call GRABACARTOLA(lin)

End If
lin = lin + 1
Wend


End Sub

'Sub TRASPASADATOS()
'Dim lin As Double
'
'Close 20
'Open File1.path + "\" + ARCHIVO.text For Input As #20
'lin = 0
'While EOF(20) = False
'
'
'Line Input #20, varipaso
'If lin = 0 And InStr(1, varipaso, "BCO. SANTANDER") = 0 And InStr(1, varipaso, "BANCO SANTIAGO") = 0 Then BORRARARCHIVO = False: Exit Sub
'If lin = 0 And InStr(1, varipaso, "MONEDA LOCALPESOS") = 0 Then BORRARARCHIVO = False: Exit Sub
'
'If Mid(varipaso, 1, 2) = "00" Then
'If Mid(varipaso, 4, 9) <> Mid(rutempresa, 1, 8) + Mid(rutempresa, 10, 1) Then
'MsgBox ("ARCHIVO NO CORRESPONDE A DATOS DE LA EMPRESA ACTIVA " + Mid(varipaso, 3, 9) + "-" + Mid(varipaso, 11, 1) + " " + rutempresa)
'Exit Sub
'End If
'
'End If
'
'If Mid(varipaso, 1, 2) <> "00" Then
'
'Call GRABACARTOLA(lin)
'
'End If
'lin = lin + 1
'Wend
'
'
'End Sub

Sub TRASPASADATOS2()
Dim lin As Double
Dim CodEmpresa As String
Close 20
BORRARARCHIVO = False

Open File1.path + "\" + ARCHIVO.text For Input As #20
lin = 0
While EOF(20) = False
 
   
Line Input #20, varipaso
If lin = 0 And InStr(1, varipaso, "BCO. SANTANDER") = 0 And InStr(1, varipaso, "BANCO SANTIAGO") = 0 Then BORRARARCHIVO = False: Exit Sub
If lin = 0 And InStr(1, varipaso, "MONEDA LOCALPESOS") <> 0 Then BORRARARCHIVO = False: Exit Sub

If Mid(varipaso, 1, 2) = "00" Then
'If Mid(varipaso, 4, 9) <> Mid(rutempresa, 1, 8) + Mid(rutempresa, 10, 1) Then
'MsgBox ("ARCHIVO NO CORRESPONDE A DATOS DE LA EMPRESA ACTIVA " + Mid(varipaso, 3, 9) + "-" + Mid(varipaso, 11, 1) + " " + rutempresa)
'Exit Sub
'End If
CodEmpresa = LeerCodigoEmpresa(Format(Mid(varipaso, 4, 8), "000000000"))
 
If CodEmpresa = "" Then
    MsgBox "RUT NO ENCONTRADO " & Mid(varipaso, 4, 9)
    Exit Sub
End If
End If

If Mid(varipaso, 1, 2) <> "00" Then
BORRARARCHIVO = True

Call GRABACARTOLA2(CodEmpresa, lin)

End If
lin = lin + 1
Wend
Close 20

End Sub

Sub GRABACARTOLA(LINEA)
    Dim dolardeldia As Double
    Dim monto As Double
    
    dolardeldia = buscadolar(Format(fechasistema, "YYYY") + "-" + Mid(varipaso, 3, 2) + "-" + Mid(varipaso, 1, 2))
    monto = Mid(varipaso, 49, 15)
    
    
    campos(0, 0) = "cuenta"
    campos(1, 0) = "cartola"
    campos(2, 0) = "tac"
    campos(3, 0) = "fecha"
    campos(4, 0) = "tipo"
    campos(5, 0) = "glosa"
    campos(6, 0) = "numero"
    campos(7, 0) = "monto"
    campos(8, 0) = "sucursal"
    campos(9, 0) = "glosasucursal"
    campos(10, 0) = "fecha_comparacion"
    campos(11, 0) = "comparable"
    campos(12, 0) = ""
    
    campos(0, 1) = "11120002"
    campos(1, 1) = "PROVI"
    campos(2, 1) = Mid(varipaso, 5, 1)
    campos(3, 1) = Format(fechasistema, "YYYY") + "-" + Mid(varipaso, 3, 2) + "-" + Mid(varipaso, 1, 2)
    
    If Mid(varipaso, 6, 5) = "00000" And Mid(varipaso, 11, 3) = "Dep" Then
    Exit Sub
    Else
    campos(4, 1) = Mid(varipaso, 6, 5)
    End If
    
    campos(5, 1) = Mid(Replace(varipaso, "ó", "o"), 11, 30) & " US " & monto & " X $ " & dolardeldia
    campos(6, 1) = Mid(varipaso, 42, 7)
    campos(7, 1) = monto * dolardeldia
    
    campos(8, 1) = Mid(varipaso, 64, 5)
    campos(9, 1) = Mid(varipaso, 69, 12)
    campos(10, 1) = campos(3, 1)
    campos(11, 1) = "0"
    
    campos(0, 2) = "cartolasbancarias"
           
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    k = sqlconta.status
End Sub

Function buscadolar(FECHACONSULTA) As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select tipocambio from " & clientesistema & "teso.cambio_dolares "
    csql.sql = csql.sql & "where fecha='" & Format(FECHACONSULTA, "yyyy-mm-dd") & "' and tipomoneda='01' limit 0,1"
    csql.Execute
        buscadolar = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        buscadolar = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
End Function
Sub GRABACARTOLA2(empresa, LINEA)
    campos(0, 0) = "cuenta"
    campos(1, 0) = "cartola"
    campos(2, 0) = "tac"
    campos(3, 0) = "fecha"
    campos(4, 0) = "tipo"
    campos(5, 0) = "glosa"
    campos(6, 0) = "numero"
    campos(7, 0) = "monto"
    campos(8, 0) = "sucursal"
    campos(9, 0) = "glosasucursal"
    campos(10, 0) = "fecha_comparacion"
    campos(11, 0) = "comparable"
    campos(12, 0) = ""
    campos(0, 1) = "11120001"
    campos(1, 1) = "PROVI"
    campos(2, 1) = Mid(varipaso, 5, 1)
    campos(3, 1) = Format(fechasistema, "YYYY") + "-" + Mid(varipaso, 3, 2) + "-" + Mid(varipaso, 1, 2)
    If Mid(varipaso, 6, 5) = "00000" And Mid(varipaso, 11, 3) = "Dep" Then
    Exit Sub
    Else
    campos(4, 1) = Mid(varipaso, 6, 5)
    End If
    campos(5, 1) = Mid(varipaso, 11, 30)
    campos(6, 1) = Mid(varipaso, 42, 7)
    campos(7, 1) = Mid(varipaso, 49, 15)
    campos(8, 1) = Mid(varipaso, 64, 5)
    campos(9, 1) = Mid(varipaso, 69, 12)
    campos(10, 1) = campos(3, 1)
    campos(11, 1) = "0"
    
    campos(0, 2) = clientesistema & "conta" & empresa & ".cartolasbancarias"
           
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    k = sqlconta.status
End Sub

Sub modificacartola(cuenta, cartola, fecha, tac, tipo, glosa, numero, monto, SUCURSAL, glosasucursal, fecha_comparacion, comparable)
    campos(0, 0) = "fecha_comparacion"
    campos(1, 0) = "comparable"
    campos(2, 0) = ""
    campos(0, 1) = Format(fecha_comparacion, "yyyy-mm-dd")
    campos(1, 1) = comparable
    campos(0, 2) = "cartolasbancarias"
    condicion = "cuenta='" + cuenta + "' and cartola='" + cartola + "' and fecha='" + Format(fecha, "yyyy-mm-dd") + "' and tac='" + tac + "' and tipo='" + tipo + "' and glosa='" + glosa + "' and numero='" + numero + "' and monto='" + monto + "' and sucursal='" + SUCURSAL + "' and glosasucursal='" + glosasucursal + "'"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    k = sqlconta.status
End Sub



Private Sub Command3_Click()
CARGATXT.Visible = False


End Sub

Private Sub Command4_Click()
Dim k As Double
barra.Max = Grid1.Rows + 1
barra.Value = 0

For k = 1 To Grid1.Rows - 1
barra.Value = barra.Value + 1
If Grid1.Cell(k, 9).text = "NO CONCILIADO" Then
Call conciliacheque(dato1.text + dato2.text + dato3.text, Format(Grid1.Cell(k, 5).text, "0000000000"), Grid1.Cell(k, 1).text)
barra.Refresh

End If
Next k
Call dato3_KeyPress(13)
End Sub

Private Sub Command5_Click()
Dim titulo As String
If Option1.Value = True Then titulo = "TODAS"
If Option2.Value = True Then titulo = "INCONSISTENCIAS"

Call CABEZAS2("LISTA CARTOLAS BANCARIAS", titulo)

Grid1.PageSetup.BlackAndWhite = True




Grid1.PrintPreview


End Sub

Private Sub Command6_Click()
Dim k As Double
Dim fecha As String

Dim numero As String

Dim LINEA As Double
Dim DH As String
Dim DH2 As String
If Format(desdefecha.Caption, "yyyy-mm") > "2008-08" Then
If Grid1.Rows > 1 Then
barra.Max = Grid1.Rows + 1
barra.Value = 0
fecha = Format(Grid1.Cell(1, 1).text, "yyyy-mm-dd")
numero = LEERULTIMOFOLIO("CB")
LINEA = 0
For k = 1 To Grid1.Rows - 1

barra.Value = barra.Value + 1
If Grid1.Cell(k, 9).text <> "CONTABILIZADO" Then
    If fecha <> Format(Grid1.Cell(k, 1).text, "yyyy-mm-dd") Then
    LINEA = 0
    numero = LEERULTIMOFOLIO("CB")
    fecha = Format(Grid1.Cell(k, 1).text, "yyyy-mm-dd")
    
    End If
    
    If Grid1.Cell(k, 14).text <> "" Then
    
    banco_cuenta = Grid1.Cell(k, 14).text
    banco_glosa2 = UCase(Grid1.Cell(k, 4).text)
    banco_dh = Grid1.Cell(k, 2).text
    If banco_dh = "A" Then DH = "D": DH2 = "H"
    If banco_dh = "C" Then DH = "H": DH2 = "D"
    
    LINEA = LINEA + 1
    Call grabarcomprobante_lineas("CB", numero, LINEA, fecha, dato1.text + dato2.text + dato3.text, "", "", "", banco_glosa2, "CB", numero, fecha, fecha, Grid1.Cell(k, 6).text, DH, USUARIOSISTEMA, Format(fecha, "mm"), Format(fecha, "yyyy"), Date, Time, "")
    LINEA = LINEA + 1
    Call grabarcomprobante_lineas("CB", numero, LINEA, fecha, banco_cuenta, "", "", "", banco_glosa2, "CB", numero, fecha, fecha, Grid1.Cell(k, 6).text, DH2, USUARIOSISTEMA, Format(fecha, "mm"), Format(fecha, "yyyy"), Date, Time, "")
    End If
    
    
    
   

barra.Refresh
End If

Next k
Call dato3_KeyPress(13)
End If
End If

End Sub

Private Sub Command7_Click()

If Verifica_Permiso(Me.Caption, "elimina") = True Then
    Call eliminacartolas(Format(elidesde, "yyyy-mm-dd"), Format(elihasta, "yyyy-mm-dd"), dato1.text + dato2.text + dato3.text)
End If

Call dato3_KeyPress(13)
frameelimina.Visible = False

End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub Command9_Click()
frameelimina.Visible = False
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub Dir1_Change()
Dir1.path = Drive1.Drive
File1.path = Dir1.path
File1.Pattern = "*.TXT"


End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
File1.path = Dir1.path
File1.Pattern = "*.TXT"

End Sub

Private Sub ELIMINA_Click()
elidesde.Value = Format(fechasistema, "yyyy-mm") + "-01"
elihasta.Value = Format(fechasistema, "yyyy-mm-dd")

frameelimina.Visible = True

End Sub

Private Sub File1_Click()
k = File1.ListIndex

ARCHIVO.text = File1.List(k)

End Sub

Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD
    sc = 0
CARGAGRILLA
CARGATXT.Visible = False

desdefecha.Caption = "01-" + Format(fechasistema, "mm-yyyy")
hastafecha.Caption = fechasistema



    

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
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
    
    If KeyAscii = 13 Then
    dato3.text = Format(dato3.text, "0000")
    lblBanco.Caption = leerNombreMayor(dato1.text + dato2.text + Format(dato3.text, "0000"))
    Call ceros(dato3)
    Call Pregunta(dato3, dato3)
    Call Command10_Click
    
    
End If

End Sub

Sub leer()
    Rem lee cuenta madre
  
lee2:    Rem lee cuenta madre
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "' año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    
    
no:
   
End Sub
   
    


Sub carga()
    habilita (True)
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    
    
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    
    

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
    cfijo = "banco='1'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato3.Enabled = True
    dato1.text = Mid(pivote.text, 1, 2)
    dato2.text = Mid(pivote.text, 3, 2)
    dato3.text = Mid(pivote.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
    
no:
End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub ELIMINAR()
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub


Private Sub lblhistorico_Click(Index As Integer)

End Sub






Sub retorno()
disponible (True)
habilita (False)
limpia

dato1.Enabled = True
dato1.SetFocus
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
        
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


Private Sub opciones_GotFocus()



End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 15)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TIPO"
    FORMATOGRILLA(1, 3) = "CODIGO"
    FORMATOGRILLA(1, 4) = "GLOSA"
    FORMATOGRILLA(1, 5) = "NUMERO"
    FORMATOGRILLA(1, 6) = "MONTO"
    FORMATOGRILLA(1, 7) = "COD"
    FORMATOGRILLA(1, 8) = "SUCURSAL"
    FORMATOGRILLA(1, 9) = "COBRADO"
    FORMATOGRILLA(1, 10) = "VENCIA"
    FORMATOGRILLA(1, 11) = "MONTO"
    FORMATOGRILLA(1, 12) = "F.COMPARA"
    FORMATOGRILLA(1, 13) = "NO VENTA"
    FORMATOGRILLA(1, 14) = "CUENTA"
    FORMATOGRILLA(1, 15) = "NOMBRE"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "3"
    FORMATOGRILLA(2, 3) = "6"
    FORMATOGRILLA(2, 4) = "30"
    FORMATOGRILLA(2, 5) = "6"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "0"
    FORMATOGRILLA(2, 8) = "0"
    FORMATOGRILLA(2, 9) = "10"
    FORMATOGRILLA(2, 10) = "8"
    FORMATOGRILLA(2, 11) = "8"
    FORMATOGRILLA(2, 12) = "8"
    FORMATOGRILLA(2, 13) = "8"
    FORMATOGRILLA(2, 14) = "8"
    FORMATOGRILLA(2, 15) = "20"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
        
    FORMATOGRILLA(3, 10) = "D"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "D"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,###,##0"
    FORMATOGRILLA(4, 11) = "###,###,###,##0"
    
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
    FORMATOGRILLA(5, 12) = "FALSE"
    FORMATOGRILLA(5, 13) = "FALSE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    
    Grid1.Cols = 16
    Grid1.Rows = 2
    
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
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
    Grid1.Column(13).CellType = cellCheckBox
    Grid1.Column(12).CellType = cellCalendar
    
    
End Sub



Private Sub LEERCARTOLAS(cuenta)

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
    
    LINEA = 0
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT * "
        csql.sql = csql.sql + "FROM cartolasbancarias where cuenta='" + cuenta + "' and fecha>='" + Format(desdefecha.Caption, "yyyy-mm-dd") + "' and fecha <='" + Format(hastafecha.Caption, "yyyy-mm-dd") + "' "
        If Check2.Value = 1 Then
        csql.sql = csql.sql + "and (tipo='00451' or tipo='00452' or tipo='00700' or tipo='01586' or tipo='00453' ) "
        End If
        If Check3.Value = 1 Then
        csql.sql = csql.sql + "and (tipo='00943' ) "
        End If
        
        
        csql.Execute
        total = 0
        total2 = 0
        
        Grid1.Rows = 1
Grid1.AutoRedraw = False


        
        If csql.RowsAffected > 0 Then
barra.Max = csql.RowsAffected + 1
barra.Value = 0

        
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
             
             If Option1.Value = True And Check1.Value = 0 Then
                    Grid1.Rows = Grid1.Rows + 1
                    LINEA = Grid1.Rows - 1
             
                     Grid1.Cell(LINEA, 1).text = resultados(2)
                     Grid1.Cell(LINEA, 2).text = resultados(3)
                     Grid1.Cell(LINEA, 3).text = resultados(4)
                    Grid1.Cell(LINEA, 4).text = resultados(5)
                    Grid1.Cell(LINEA, 5).text = resultados(6)
                    Grid1.Cell(LINEA, 6).text = resultados(7)
                    Grid1.Cell(LINEA, 7).text = resultados(8)
                    Grid1.Cell(LINEA, 8).text = resultados(9)
                    If IsNull(resultados("fecha_comparacion")) = False Then
                    Grid1.Cell(LINEA, 12).text = resultados("fecha_comparacion")
                    End If
                    Grid1.Cell(LINEA, 13).text = resultados("comparable")
                    If LEERcuentacontable2(resultados(4), resultados(5)) <> "" Then
                              Grid1.Cell(LINEA, 14).text = LEERcuentacontable2(resultados(4), resultados(5))
                              Grid1.Cell(LINEA, 15).text = leerNombreMayor(Grid1.Cell(LINEA, 14).text)
                              Call LEERcuentacontabilizada(Grid1.Cell(LINEA, 14).text, resultados(2), resultados(7), banco_dh)
                              Grid1.Cell(LINEA, 9).text = conta_glosa
                              Grid1.Cell(LINEA, 10).text = Format(conta_fecha, "dd-mm-yyyy")
                              Grid1.Cell(LINEA, 11).text = conta_monto
                            End If
             
             
            
             If resultados(4) = "00510" Or resultados(4) = "00505" Or resultados(4) = "00891" Or resultados(4) = "00600" Then
                 estado = leerestadocheque(cuenta, Format(resultados(6), "0000000000"), resultados(7), resultados(2))
                If estado = "0" Then Grid1.Cell(LINEA, 9).text = "NO CONCILIADO"
                If estado = "1" Then Grid1.Cell(LINEA, 9).text = "COBRADO"
                If estado = "2" Then Grid1.Cell(LINEA, 9).text = "COBRADO ANTES"
                If estado = "3" Then Grid1.Cell(LINEA, 9).text = "COBRADO POR OTRO MONTO"
                If estado = "3" Then Grid1.Range(LINEA, 1, LINEA, 10).BackColor = vbYellow
                If estado = "4" Then Grid1.Cell(LINEA, 9).text = "NO CONTABILIDAD"
                If estado = "4" Then Grid1.Range(LINEA, 1, LINEA, 10).BackColor = &HFF&
             
                    If estado <> "4" Then
                     Grid1.Cell(LINEA, 4).text = nombregirador
                    Grid1.Cell(LINEA, 10).text = VENCIMIENTOREAL
                    Grid1.Cell(LINEA, 11).text = MONTOREAL
                           
             
                    End If
             
             
             
            End If
             End If
             
             If Option1.Value = True And Check1.Value = 1 Then
                If LEERcuentacontable2(resultados(4), resultados(5)) <> "" Then
                              Grid1.Cell(LINEA, 14).text = LEERcuentacontable2(resultados(4), resultados(5))
                              Grid1.Cell(LINEA, 15).text = leerNombreMayor(Grid1.Cell(LINEA, 14).text)
                    
                      Call LEERcuentacontabilizada(Grid1.Cell(LINEA, 14).text, resultados(2), resultados(7), banco_dh)
                        If conta_glosa = "NO CONTABILIZADO" Then
                            Grid1.Rows = Grid1.Rows + 1
                            LINEA = Grid1.Rows - 1
             
                            Grid1.Cell(LINEA, 1).text = resultados(2)
                            Grid1.Cell(LINEA, 2).text = resultados(3)
                            Grid1.Cell(LINEA, 3).text = resultados(4)
                           Grid1.Cell(LINEA, 4).text = resultados(5)
                           Grid1.Cell(LINEA, 5).text = resultados(6)
                           Grid1.Cell(LINEA, 6).text = resultados(7)
                           Grid1.Cell(LINEA, 7).text = resultados(8)
                           Grid1.Cell(LINEA, 8).text = resultados(9)
                            Grid1.Cell(LINEA, 12).text = resultados("fecha_comparacion")
                            Grid1.Cell(LINEA, 13).text = resultados("comparable")
                    
                             Grid1.Cell(LINEA, 9).text = conta_glosa
                             Grid1.Cell(LINEA, 10).text = Format(conta_fecha, "dd-mm-yyyy")
                             Grid1.Cell(LINEA, 11).text = conta_monto
                           If LEERcuentacontable2(resultados(4), resultados(5)) <> "" Then
                              Grid1.Cell(LINEA, 14).text = LEERcuentacontable2(resultados(4), resultados(5))
                              Grid1.Cell(LINEA, 15).text = leerNombreMayor(Grid1.Cell(LINEA, 14).text)
                            End If
                 
                        
                        End If
            
             End If
             End If
             
             
             If Option2.Value = True Then
                If resultados(4) = "00510" Or resultados(4) = "00505" Or resultados(4) = "00891" Or resultados(4) = "00600" Then
                    estado = leerestadocheque(cuenta, Format(resultados(6), "0000000000"), resultados(7), resultados(2))
             
                     If estado = "3" Or estado = "0" Or estado = "4" Then
                     Grid1.Rows = Grid1.Rows + 1
                     LINEA = Grid1.Rows - 1
                        Grid1.Cell(LINEA, 1).text = resultados(2)
                        Grid1.Cell(LINEA, 2).text = resultados(3)
                        Grid1.Cell(LINEA, 3).text = resultados(4)
                        Grid1.Cell(LINEA, 4).text = resultados(5)
                        Grid1.Cell(LINEA, 5).text = resultados(6)
                        Grid1.Cell(LINEA, 6).text = resultados(7)
                        Grid1.Cell(LINEA, 7).text = resultados(8)
                        Grid1.Cell(LINEA, 8).text = resultados(9)
                    Grid1.Cell(LINEA, 12).text = resultados("fecha_comparacion")
                    Grid1.Cell(LINEA, 13).text = resultados("comparable")
                    
                        If estado = "0" Then Grid1.Cell(LINEA, 9).text = "NO CONCILIADO"
                        If estado = "3" Then Grid1.Cell(LINEA, 9).text = "COBRADO POR OTRO MONTO"
                        If estado = "4" Then Grid1.Cell(LINEA, 9).text = "NO CONTABILIDAD"
                        Grid1.Cell(LINEA, 4).text = nombregirador
                        Grid1.Cell(LINEA, 10).text = VENCIMIENTOREAL
                        Grid1.Cell(LINEA, 11).text = MONTOREAL
                           
             
                    End If
                End If
          End If
        
        
        
       
             
         
          
          barra.Value = barra.Value + 1
          barra.Refresh
          
          resultados.MoveNext

         Wend
End If
Grid1.AutoRedraw = True
Grid1.Refresh



End Sub


Private Sub Grid1_CellChange(ByVal row As Long, ByVal col As Long)
If Grid1.Cell(row, 3).text = "00451" Or Grid1.Cell(row, 3).text = "00453" Or Grid1.Cell(row, 3).text = "00452" Or Grid1.Cell(row, 3).text = "00700" Or Grid1.Cell(row, 3).text = "01586" Then
Call modificacartola(dato1.text + dato2.text + dato3.text, "PROVI", Grid1.Cell(row, 1).text, Grid1.Cell(row, 2).text, Grid1.Cell(row, 3).text, Grid1.Cell(row, 4).text, Grid1.Cell(row, 5).text, Grid1.Cell(row, 6).text, Grid1.Cell(row, 7).text, Grid1.Cell(row, 8).text, Grid1.Cell(row, 12).text, Grid1.Cell(row, 13).text)
End If

End Sub

Private Sub Grid1_DblClick()
If Grid1.Cell(Grid1.ActiveCell.row, 2).text = "C" And (Grid1.Cell(Grid1.ActiveCell.row, 3).text = "00505" Or Grid1.Cell(Grid1.ActiveCell.row, 3).text = "00510" Or Grid1.Cell(Grid1.ActiveCell.row, 3).text = "00891") Then
banco01.dato1.text = dato1.text
banco01.dato2.text = dato2.text
banco01.dato3.text = dato3.text
banco01.dato4.text = Grid1.Cell(Grid1.ActiveCell.row, 5).text
Rem -- Call banco01.dato4_KeyPress(13)
banco01.Show

End If

End Sub


Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
If Grid1.Cell(row, 3).text = "00451" Or Grid1.Cell(row, 3).text = "00453" Or Grid1.Cell(row, 3).text = "00452" Or Grid1.Cell(row, 3).text = "00700" Or Grid1.Cell(row, 3).text = "01586" Then
Call modificacartola(dato1.text + dato2.text + dato3.text, "PROVI", Grid1.Cell(row, 1).text, Grid1.Cell(row, 2).text, Grid1.Cell(row, 3).text, Grid1.Cell(row, 4).text, Grid1.Cell(row, 5).text, Grid1.Cell(row, 6).text, Grid1.Cell(row, 7).text, Grid1.Cell(row, 8).text, Grid1.Cell(row, 12).text, Grid1.Cell(row, 13).text)
End If


End Sub

Private Sub Option1_Click()
 CARGAGRILLA
    
    Call LEERCARTOLAS(dato1.text + dato2.text + dato3.text)
    
End Sub

Private Sub Option2_Click()
 CARGAGRILLA
    
    Call LEERCARTOLAS(dato1.text + dato2.text + dato3.text)
    
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
Public Function LEERcuentacontable2(codigo, glosa) As String


    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = conta

            csql.sql = "select * from cartolasbancarias_codigoscontables where codigo='" + codigo + "' and codigocontable<>'00000000' "
            
            csql.Execute
    LEERcuentacontable2 = ""
    If csql.RowsAffected > 0 Then
  
    
    Set resultados = csql.OpenResultset
        
        LEERcuentacontable2 = resultados(3)
        If codigo = "00943" Then
            LEERcuentacontable2 = "11500160"
            If InStr(1, UCase(glosa), "TRANSBAN", vbBinaryCompare) > 0 Or InStr(1, UCase(glosa), "REDCOMPRA", vbBinaryCompare) > 0 Then
            LEERcuentacontable2 = "11100008"
            End If
            If InStr(1, UCase(glosa), "ISW", vbBinaryCompare) > 0 Then
            LEERcuentacontable2 = "11100011"
            End If
            
            
        End If
        If LEERcuentacontable2 = "11500110" Then
            LEERcuentacontable2 = "11100005"
        End If
        banco_dh = resultados(4)
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
    
'    ' agregado el 20-03-2014
'    campos(0, 2) = "movimientoscontables"
'    condicion = "fecha BETWEEN '" & fe1 & "' AND '" & fe2 & "' AND tipo='CB'"
'    op = 4
'    sqlconta.response = campos
'    Set sqlconta.conexion = contadb
'    Call sqlconta.sqlconta(op, condicion)
    
End Sub



Public Function LeerCodigoEmpresa(rutempresa) As String
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    campos(0, 2) = clientesistema & "conta.maestroempresas"
    condicion = "rut='" & Val(rutempresa) & "-" & rut(Format(rutempresa, "000000000")) & "' "
    op = 5
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LeerCodigoEmpresa = sqlconta.response(0, 3)
    End If
  
End Function

