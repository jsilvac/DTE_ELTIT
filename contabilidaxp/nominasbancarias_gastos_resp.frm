VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0016 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Facturas de Compras Relacionadas"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17415
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   659
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1161
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   14280
      TabIndex        =   27
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
         TabIndex        =   29
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   280
         Width           =   1335
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8865
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9825
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   17340
      _ExtentX        =   30586
      _ExtentY        =   17330
      BackColor       =   16761024
      Caption         =   "Genera Nominas Pago Proveedores"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   17160
         _ExtentX        =   30268
         _ExtentY        =   1852
         BackColor       =   16761024
         Caption         =   "DATOS DE FILTRADO"
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
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "LISTAR"
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
            Left            =   14040
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   1695
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   7
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
            BackColor       =   16761024
            Caption         =   "DESDE EL MES"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   65535
            ColorBarraArriba=   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOMES 
               Height          =   315
               Left            =   45
               TabIndex        =   8
               Top             =   270
               Width           =   3180
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   3510
            TabIndex        =   9
            Top             =   270
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1191
            BackColor       =   16761024
            Caption         =   "AÑO"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   65535
            ColorBarraArriba=   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOAÑO 
               Height          =   315
               Left            =   90
               TabIndex        =   10
               Top             =   270
               Width           =   2865
            End
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   675
            Left            =   6705
            TabIndex        =   11
            Top             =   270
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   1191
            BackColor       =   16761024
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   65535
            ColorBarraArriba=   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox ComboLOCAL 
               Height          =   315
               Left            =   90
               TabIndex        =   12
               Top             =   270
               Width           =   4395
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   7440
         Left            =   90
         TabIndex        =   3
         Top             =   1485
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   13123
         BackColor       =   16761024
         Caption         =   "LISTADO DE NOMINAS VIGENTES"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   4194304
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
            BackColor       =   &H00FF8080&
            Caption         =   "IMPRIMIR"
            Height          =   330
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   6840
            Width           =   1725
         End
         Begin FlexCell.Grid Grid1 
            Height          =   6360
            Left            =   90
            TabIndex        =   4
            Top             =   270
            Width           =   6225
            _ExtentX        =   10980
            _ExtentY        =   11218
            BackColorFixed  =   16761024
            BackColorSel    =   16761024
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   12640511
            Rows            =   30
            DateFormat      =   2
         End
         Begin VB.Label acumulado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000007&
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
            Height          =   465
            Left            =   4080
            TabIndex        =   19
            Top             =   6795
            Width           =   1950
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   8250
         Left            =   6600
         TabIndex        =   13
         Top             =   1485
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   14552
         BackColor       =   16761024
         Caption         =   "DETALLE NOMINA DE PAGOS"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin XPFrame.FrameXp FrameXp8 
            Height          =   1590
            Left            =   8325
            TabIndex        =   22
            Top             =   6660
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   2805
            BackColor       =   16761024
            Caption         =   "Envios de Mail"
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
            Begin VB.TextBox datoprove 
               Height          =   330
               Left            =   225
               TabIndex        =   26
               Top             =   855
               Width           =   1905
            End
            Begin VB.OptionButton oprut 
               Caption         =   "Rut Individual"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   315
               TabIndex        =   25
               Top             =   585
               Width           =   1770
            End
            Begin VB.OptionButton opto 
               Caption         =   "Todos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   315
               TabIndex        =   24
               Top             =   315
               Value           =   -1  'True
               Width           =   1770
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00FF8080&
               Caption         =   "Enviar Email @"
               Height          =   330
               Left            =   90
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   1215
               Width           =   2130
            End
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FF8080&
            Caption         =   "Genera Archivo banco"
            Height          =   330
            Left            =   5640
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   7320
            Width           =   2130
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FF8080&
            Caption         =   "IMPRIMIR"
            Height          =   330
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   7320
            Width           =   2130
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "NOMINA DETALLADA"
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
            Left            =   4455
            TabIndex        =   16
            Top             =   6840
            Width           =   2400
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "AGRUPADA POR RUT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1755
            TabIndex        =   15
            Top             =   6885
            Value           =   -1  'True
            Width           =   2310
         End
         Begin FlexCell.Grid Grid2 
            Height          =   6360
            Left            =   90
            TabIndex        =   14
            Top             =   270
            Width           =   10545
            _ExtentX        =   18600
            _ExtentY        =   11218
            BackColorFixed  =   16761024
            BackColorSel    =   16761024
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   12640511
            Rows            =   30
            SelectionMode   =   1
            DateFormat      =   2
         End
         Begin FlexCell.Grid Grid3 
            Height          =   6360
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   10545
            _ExtentX        =   18600
            _ExtentY        =   11218
            BackColorFixed  =   16761024
            BackColorSel    =   16761024
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   12640511
            Rows            =   30
            SelectionMode   =   1
            DateFormat      =   2
         End
      End
   End
End
Attribute VB_Name = "prove0016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String
Private montonomina As Double
Private glosapago As String
Private montopago As Double
Private numeropago As String
Private LINEAS As Double
Private banco As String
Private SUCURSAL As String
Private cuentacorriente As String
Private email As String
Private rutproveedor As String
Private RUTRETIRA2 As String
Private NOMBRERETIRA2 As String
Private FECHANOMINA As String
Private revisanumero As Boolean
Private modopago As String
Private lineacom As Double
Private GENERANOMINA As Boolean





Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1

leer


End Sub



Private Sub Command3_Click()
IMPRIMIR2


End Sub



Private Sub Command4_Click()
If GENERANOMINA = True Then
If Option2.Value = True Then
GENERATXT
End If
Else
MsgBox "IMPOSIBLE GENERAR NOMINA COMPROBANTES CON MAS DE 11 LINEAS - ELIMINE COMPROBANTE Y RECTIFIQUE"
End If

End Sub
'anterior
'Private Sub Command5_Click()
'Dim rutproveedor As String
'
'
'
'Dim s As Integer
'
'If Option2.Value = True Then
'Shell "C:\Archivos de programa\Microsoft Office\OFFICE11\OUTLOOK.EXE"
'
'CARGAGRILLA3
'Grid3.Rows = 1
'Grid3.Rows = 1
'For k = 1 To Grid2.Rows - 1
'
'Grid3.Rows = Grid3.Rows + 1
'
'For s = 1 To Grid2.Cols - 1
'Grid3.Cell(Grid3.Rows - 1, s).text = Grid2.Cell(k, s).text
'Next s
'If Mid(Grid2.Cell(k - 1, 1).text, 1, 5) = "PROVE" Then
'rutproveedor = Grid2.Cell(k - 1, 2).text
'Grid3.Rows = Grid3.Rows + 1
'Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Merge
'Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Alignment = cellCenterCenter
'Grid3.Cell(Grid3.Rows - 1, 1).text = "CANCELACION DE FACTURAS VIA TRANSFERENCIA BANCARIA "
'Grid3.Rows = Grid3.Rows + 1
'Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Merge
'Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Alignment = cellCenterCenter
'Grid3.Cell(Grid3.Rows - 1, 1).text = "DEPARTAMENTO PAGO DE PROVEEDORES"
'Grid3.Rows = Grid3.Rows + 1
'Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Merge
'Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Alignment = cellCenterCenter
'Grid3.Cell(Grid3.Rows - 1, 1).text = nombreempresa
'
'
'
'Rem Call cabezas(Grid2.Cell(K - 1, 3).text)
'If oprut.Value = False Then
'Grid3.ExportToHTML ("c:\comprobantedepago.htm")
'Call SendOutlookMail("CANCELACION DE FACTURAS " + nombreempresa, Grid3.Cell(Grid3.Rows - 4, 4).text, "SEÑORES:" + vbCrLf + Grid3.Cell(Grid3.Rows - 5, 3).text + vbCrLf + vbCrLf + " Por medio de la presente informamos a ud. que ha sido abonado en su cuenta" + vbCrLf + "el siguiente monto:" + Format(Grid3.Cell(Grid3.Rows - 5, 6).text, "$ ###,###,###") + vbCrLf + "Atentamente " + vbCrLf + "Departamento Pago de Proveedores" + vbCrLf + nombreempresa)
'End If
'If oprut.Value = True And rutproveedor = datoprove.text Then
'Grid3.ExportToHTML ("c:\comprobantedepago.htm")
'Call SendOutlookMail("CANCELACION DE FACTURAS " + nombreempresa, Grid3.Cell(Grid3.Rows - 4, 4).text, "SEÑORES:" + vbCrLf + Grid3.Cell(Grid3.Rows - 5, 3).text + vbCrLf + vbCrLf + " Por medio de la presente informamos a ud. que ha sido abonado en su cuenta" + vbCrLf + "el siguiente monto:" + Format(Grid3.Cell(Grid3.Rows - 5, 6).text, "$ ###,###,###") + vbCrLf + "Atentamente " + vbCrLf + "Departamento Pago de Proveedores" + vbCrLf + nombreempresa)
'End If
'
'Grid3.Rows = 1
'End If
'
'Next k
'End If
'End Sub

Private Sub Command5_Click()
Dim rutproveedor As String
Dim s As Integer
Dim archivo As String
Dim MENSAJE As String
If UsuarioCorreo = "" Then
MsgBox "SU CORREO NO ESTA CONFIGURADO CORRECTAMENTE" & vbCr & _
                            "SOLICITE AYUDA A COMPUTACION " & vbCr & _
                            "SU CORREO NO SE ENVIARÁ ", vbInformation, "ENVIO DE CORREOS"
GoTo fin

Else
'    Call VerificaAplicacion("admin_comunicaciones.exe")
If Option2.Value = True Then
'    EnviarEmail.Show
    CARGAGRILLA3
    Grid3.Rows = 1
    Grid3.Rows = 1
    For k = 1 To Grid2.Rows - 1
    
    Grid3.Rows = Grid3.Rows + 1
    
    For s = 1 To Grid2.Cols - 1
    Grid3.Cell(Grid3.Rows - 1, s).text = Grid2.Cell(k, s).text
    Next s
    If Mid(Grid2.Cell(k - 1, 1).text, 1, 5) = "PROVE" Then
    rutproveedor = Grid2.Cell(k - 1, 2).text
    Grid3.Rows = Grid3.Rows + 1
    Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Merge
    Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Alignment = cellCenterCenter
    Grid3.Cell(Grid3.Rows - 1, 1).text = "CANCELACION DE FACTURAS VIA TRANSFERENCIA BANCARIA "
    Grid3.Rows = Grid3.Rows + 1
    Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Merge
    Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Alignment = cellCenterCenter
    Grid3.Cell(Grid3.Rows - 1, 1).text = "DEPARTAMENTO PAGO DE PROVEEDORES"
    Grid3.Rows = Grid3.Rows + 1
    Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Merge
    Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, Grid3.Cols - 1).Alignment = cellCenterCenter
    Grid3.Cell(Grid3.Rows - 1, 1).text = nombreempresa
    
    Rem Call cabezas(Grid2.Cell(K - 1, 3).text)
    'lineas originales 31-08-2016 rz
'    If oprut.Value = False Then
'    Grid3.ExportToHTML ("c:\comprobantedepago.htm")
'    Call EnviarEmail.ENVIARMAIL(nombreempresa, usuariocorreo, clavecorreo, "CANCELACION DE FACTURAS", _
'                                "SEÑORES:" & vbCrLf & Grid3.Cell(Grid3.Rows - 5, 3).text & vbCrLf & vbCrLf & _
'                            " Por medio de la presente informamos a ud. que ha sido abonado en su cuenta" & _
'                            vbCrLf & "el siguiente monto:" & Format(Grid3.Cell(Grid3.Rows - 5, 6).text, "$ ###,###,###") & _
'                            vbCrLf & "Atentamente " & vbCrLf + "Departamento Pago de Proveedores" + vbCrLf + nombreempresa, _
'                            ServerCorreo, Grid3.Cell(Grid3.Rows - 4, 4).text, Grid3.Cell(Grid3.Rows - 5, 3).text, "c:\comprobantedepago.htm")
'    End If
'    If oprut.Value = True And rutproveedor = datoprove.text Then
'    Grid3.ExportToHTML ("c:\comprobantedepago.htm")
'    'Call SendOutlookMail("CANCELACION DE FACTURAS " + nombreempresa, Grid3.Cell(Grid3.Rows - 4, 4).text, "SEÑORES:" + vbCrLf + Grid3.Cell(Grid3.Rows - 5, 3).text + vbCrLf + vbCrLf + " Por medio de la presente informamos a ud. que ha sido abonado en su cuenta" + vbCrLf + "el siguiente monto:" + Format(Grid3.Cell(Grid3.Rows - 5, 6).text, "$ ###,###,###") + vbCrLf + "Atentamente " + vbCrLf + "Departamento Pago de Proveedores" + vbCrLf + nombreempresa)
'    Call EnviarEmail.ENVIARMAIL(nombreempresa, usuariocorreo, clavecorreo, "CANCELACION DE FACTURAS", _
'                                "SEÑORES:" & vbCrLf & Grid3.Cell(Grid3.Rows - 5, 3).text & vbCrLf & vbCrLf & _
'                                " Por medio de la presente informamos a ud. que ha sido abonado en su cuenta" & _
'                                vbCrLf & "el siguiente monto:" & Format(Grid3.Cell(Grid3.Rows - 5, 6).text, "$ ###,###,###") & _
'                                vbCrLf & "Atentamente " & vbCrLf + "Departamento Pago de Proveedores" + vbCrLf + nombreempresa, _
'                                ServerCorreo, Grid3.Cell(Grid3.Rows - 4, 4).text, Grid3.Cell(Grid3.Rows - 5, 3).text, "c:\comprobantedepago.htm")
'    End If
'
archivo = "c:\comprobantedepago.htm"
    If oprut.Value = False Then
    MENSAJE = "SEÑORES:" & vbCrLf & Grid3.Cell(Grid3.Rows - 5, 3).text & vbCrLf & vbCrLf & _
                            " Por medio de la presente informamos a ud. que ha sido abonado en su cuenta" & _
                            vbCrLf & "el siguiente monto:" & Format(Grid3.Cell(Grid3.Rows - 5, 6).text, "$ ###,###,###") & _
                            vbCrLf & "Atentamente " & vbCrLf + "Departamento Pago de Proveedores" + vbCrLf + nombreempresa

    Grid3.ExportToHTML ("c:\comprobantedepago.htm")
'    Call enviaremail2(ServerCorreo, UsuarioCorreo, ClaveCorreo, 25, nombreempresa, Grid3.Cell(Grid3.Rows - 4, 4).text, "rauls@eltit.cl", "", "CANCELACION DE FACTURAS", MENSAJE, ARCHIVO)
   
    Call EnviarMail2("CANCELACION DE FACTURAS", MENSAJE, ServerCorreo, Grid3.Cell(Grid3.Rows - 4, 4).text, Grid3.Cell(Grid3.Rows - 5, 3).text, archivo, "")

  
   
   
    End If
    If oprut.Value = True And rutproveedor = datoprove.text Then
    Grid3.ExportToHTML ("c:\comprobantedepago.htm")
      MENSAJE = "SEÑORES:" & vbNewLine & Grid3.Cell(Grid3.Rows - 5, 3).text & vbNewLine & vbNewLine & _
                            " Por medio de la presente informamos a ud. que ha sido abonado en su cuenta" & _
                            vbNewLine & "el siguiente monto:" & Format(Grid3.Cell(Grid3.Rows - 5, 6).text, "$ ###,###,###") & _
                            vbNewLine & "Atentamente " & vbNewLine + "Departamento Pago de Proveedores" + vbNewLine + nombreempresa

'    Call ENVIARMAIL2("CANCELACION DE FACTURAS", mensaje, ServerCorreo, Grid3.Cell(Grid3.Rows - 4, 4).text, Grid3.Cell(Grid3.Rows - 5, 3).text, archivo, "", usuariocorreo, clavecorreo, 25)
'   Call enviaremail2(ServerCorreo, UsuarioCorreo, ClaveCorreo, 25, nombreempresa, Grid3.Cell(Grid3.Rows - 4, 4).text, "rauls@eltit.cl", "", "CANCELACION DE FACTURAS", MENSAJE, ARCHIVO)
   Call EnviarMail2("CANCELACION DE FACTURAS", MENSAJE, ServerCorreo, Grid3.Cell(Grid3.Rows - 4, 4).text, Grid3.Cell(Grid3.Rows - 5, 3).text, archivo, "")

    End If


    Grid3.Rows = 1
    End If
    
    Next k
'    Unload EnviarEmail
    End If
End If
fin:
End Sub

'Public Sub enviaremail2(servidor_correo, usuario_correo, clave_correo, puertosalida, _
'                        NOMBRE, destinatario, copia, copiaoculta, Asunto, MENSAJE, _
'                        Optional adjunto As String)
'    Dim iMsg As Object
'   Dim iConf As Object
'Dim strbody  As String
'    Dim Flds As Variant
'     Dim ssl As Double
'     Dim smtp As Boolean
'Set iMsg = CreateObject("CDO.Message")
'Set iConf = CreateObject("CDO.Configuration")
'smtp = False
'ssl = 1
'puertosalida = 25
'If InStr(Servidor, "gmail") > 0 Then
'    ssl = 2
'    puertosalida = 465
'End If
'
'
'If InStr(Servidor, "@eltit") > 0 Then
'    ssl = 2
'    puertosalida = 465
'    smtp = True
'End If
'
'
'
'iConf.Load -1
'Set Flds = iConf.Fields
'    With Flds
'        .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = smtp
'        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = ssl
'        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = usuario_correo
'        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = clave_correo
'        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = servidor_correo
'        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = puertosalida
'        .Update
'    End With
'copiaoculta = "" ' "soporte@adminerp.cl"
'With iMsg
'    Set .Configuration = iConf
'
'    .To = destinatario
'    .cc = copia
'    .BCC = copiaoculta
' Rem   .From = usuario_correo
'    .From = """Proveedores Eltit"" <" & usuario_correo & ">"
'    .Subject = Asunto
'    .TextBody = MENSAJE
'
'    If ExisteArchivo(adjunto) = True Then
'        .AddAttachment (adjunto)
'    End If
'    .Send
'    MsgBox "CORREO ENVIADO A " & destinatario & " " & Now, vbInformation, "ATENCION"
'    End With
'End Sub

Public Sub EnviarMail2(ByRef Asunto, ByRef MENSAJE, ByRef Servidor, ByRef MailDestinatario, _
                        ByVal NombreDestinatario As String, ByRef ArchivAdjunto, _
                        ByVal archivadjunto2 As String)
Dim enviados As String
On Error GoTo error
Screen.MousePointer = vbHourglass
MailDestinatario = LCase(MailDestinatario)
'MailDestinatario = "cesarsandoval@adminerp.cl"
'Call empresadte(empresaactiva)
'rz 05-12-2017 '
'funcion adaptada para enviar con cuenta de eltit

If LeerCuentaAlternativa = True Then
        confi_mailsalida = email_cuenta_usuario
        confi_clavemail = email_cuenta_clave
        confi_servermail = email_cuenta_server
End If
 
If ArchivAdjunto <> "" And archivadjunto2 = "" Then enviados = ArchivAdjunto
If archivadjunto2 <> "" And ArchivAdjunto = "" Then enviados = archivadjunto2
If archivadjunto2 <> "" And ArchivAdjunto <> "" Then enviados = ArchivAdjunto + ";" + archivadjunto2



Dim iMsg As Object
Dim iConf As Object
Dim strbody  As String
Dim Flds As Variant
Dim comi As String
Dim puertosalida As Double
Dim destinatario As String
'destinatario = "aalarcon@eltit.cl; rlzurita@gmail.com"
 
comi = Chr(34)
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
destinatario = Replace(destinatario, "<", "")
destinatario = Replace(destinatario, ">", "")
iConf.Load -1
 puertosalida = 465

Set Flds = iConf.Fields
With Flds
    .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = email_cuenta_usuario
    .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = email_cuenta_clave
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = email_cuenta_server
    .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = puertosalida
    .Update
End With
'MailDestinatario = "sarcesnot@hotmail.com"
With iMsg
Set .Configuration = iConf
.To = MailDestinatario
.cc = "rauls@eltit.cl"
.BCC = "" 'aalarcon@eltit.cl; raulzurita@adminerp.cl"
 
.From = comi & comi & comi & nombreempresa & comi & " <" & email_cuenta_usuario & ">"
.Subject = Asunto
.TextBody = MENSAJE

If enviados <> "" Then .AddAttachment enviados


.Send

End With

MsgBox "CORREO ENVIADO A " & MailDestinatario & " " & Now, vbInformation, "ATENCION"
 
Screen.MousePointer = vbDefault
Exit Sub


error:
MsgBox "NO SE PUDO ENVIAR EL CORREO" & vbNewLine & err.Description
Screen.MousePointer = vbDefault
End Sub


Public Sub SendOutlookMail(Subject As String, Recipient As _
String, Message As String)

On Error GoTo errorHandler
Dim oLapp As Object
Dim oItem As Object

Set oLapp = CreateObject("Outlook.application")
Rem - Set oLapp = CreateObject("cdo.mesagges")
Set oItem = oLapp.CreateItem(0)




With oItem
   .Subject = Subject
   .To = Recipient
    .body = Message
   .attachments.Add ("C:\comprobantedepago.htm")
  
   .Save
   .Close
    
    
End With
'
Set oLapp = Nothing
Set oItem = Nothing
'

Exit Sub

errorHandler:
Set oLapp = Nothing
Set oItem = Nothing
Exit Sub
End Sub


Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD
    sc = 0
CARGAGRILLA
CARGAGRILLA2

Call Conectarventas(Servidor, clientesistema + "ventas00", Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion00", Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
LEErlocales
Call COMMAND2_Click


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




Private Sub Label16_Click()
End Sub

Sub limpia()
    
    
End Sub

Sub imprimir()
Dim titulo As String
titulo = "LISTADO DE NOMINAS BANCARIAS DESDE EL MES " + COMBOMES.text + " " + COMBOAÑO.text
Call CABEZAS2(titulo, "N", "000000000")
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub
Sub IMPRIMIR2()
Dim titulo As String
titulo = "NOMINA REVISION PAGO PROVEEDORES A BANCO FECHA " + Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "dd-mm-yyyy")
Call cabezas3(titulo, "N", "000000000")
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid2.DefaultFont.Size = 8
Grid2.PageSetup.Orientation = cellPortrait


Grid2.PageSetup.PrintFixedRow = True
Grid2.PageSetup.BottomMargin = 2
Grid2.PageSetup.TopMargin = 1
Grid2.PageSetup.LeftMargin = 1
Grid2.PageSetup.RightMargin = 0
Grid2.PageSetup.BlackAndWhite = False
Grid2.PageSetup.PrintGridlines = False
Grid2.PrintPreview 100

   
End Sub

Sub grilla()
    
End Sub

Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
'Sub CARGAGRILLA()
'Rem DATOS DE LA COLUMNA
'    Dim FormatoGrilla(10, 20)
'    Grid1.DefaultFont.Size = 10
'    Grid1.DefaultFont.Bold = False
'
'
'    FormatoGrilla(1, 1) = "FECHA"
'    FormatoGrilla(1, 2) = "MONTO"
'    FormatoGrilla(1, 3) = "ESTADO"
'    FormatoGrilla(1, 4) = "AUTORIZADA"
'    FormatoGrilla(1, 5) = "ENVIADA"
'
'    Rem LARGO DE LOS DATOS
'    FormatoGrilla(2, 1) = "9"
'    FormatoGrilla(2, 2) = "10"
'    FormatoGrilla(2, 3) = "8"
'    FormatoGrilla(2, 4) = "10"
'    FormatoGrilla(2, 5) = "5"
'
'
'    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
'    FormatoGrilla(3, 1) = "D"
'    FormatoGrilla(3, 2) = "N"
'    FormatoGrilla(3, 3) = "N"
'    FormatoGrilla(3, 4) = "N"
'    FormatoGrilla(3, 5) = "N"
'
'    Rem FORMATO GRILLA
'    FormatoGrilla(4, 2) = "###,###,##0"
'    FormatoGrilla(4, 3) = "###,###,##0"
'    FormatoGrilla(4, 4) = "###,###,##0"
'    FormatoGrilla(4, 5) = "###,###,##0"
'
'    Rem LOCCKED
'    For k = 1 To 5
'    FormatoGrilla(5, k) = "TRUE"
'    Next k
'
'
'    Grid1.Cols = 4
'    Grid1.Rows = 2
'
'    Grid1.AllowUserResizing = False
'    Grid1.DisplayFocusRect = False
'    Grid1.ExtendLastCol = True
'    Grid1.BoldFixedCell = False
'    Grid1.DrawMode = cellOwnerDraw
'
'    Grid1.Appearance = Flat
'    Grid1.ScrollBarStyle = Flat
'    Grid1.FixedRowColStyle = Flat
'
''   Grid1.BackColorFixed = RGB(90, 158, 214)
''   Grid1.BackColorFixedSel = RGB(110, 180, 230)
''   Grid1.BackColorBkg = RGB(90, 158, 214)
''   Grid1.BackColorScrollBar = RGB(231, 235, 247)
''   Grid1.BackColor1 = RGB(231, 235, 247)
''   Grid1.BackColor2 = RGB(239, 243, 255)
''   Grid1.GridColor = RGB(148, 190, 231)
'   Grid1.Column(0).Width = 0
'
'    For k = 1 To Grid1.Cols - 1
'
'        Grid1.Cell(0, k).text = FormatoGrilla(1, k)
'        Grid1.Column(k).Width = Val(FormatoGrilla(2, k)) * (Grid1.DefaultFont.Size - 1)
'        Grid1.Column(k).MaxLength = Val(FormatoGrilla(2, k))
'        Grid1.Column(k).FormatString = FormatoGrilla(4, k)
'        Grid1.Column(k).Locked = FormatoGrilla(5, k)
'        If FormatoGrilla(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
'        If FormatoGrilla(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
'
'    Next k
'   Grid1.Column(3).CellType = cellCheckBox
'
'
'End Sub

Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 10
    Grid1.DefaultFont.Bold = False
    
    
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "MONTO"
    FORMATOGRILLA(1, 3) = "ESTADO"
    FORMATOGRILLA(1, 4) = "AUTORIZADA"
    FORMATOGRILLA(1, 5) = "F.LIBERACION"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "9"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "5"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "D"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 2) = "###,###,##0"
    FORMATOGRILLA(4, 3) = "###,###,##0"
    FORMATOGRILLA(4, 4) = "###,###,##0"
    FORMATOGRILLA(4, 5) = ""
    
    Rem LOCCKED
    For k = 1 To 5
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    
    Grid1.Cols = 6
    Grid1.Rows = 2
    
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
   Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * (Grid1.DefaultFont.Size - 1)
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
   Grid1.Column(3).CellType = cellCheckBox
  
    
End Sub

Sub CARGAGRILLA2()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid2.DefaultFont.Size = 10
    Grid2.DefaultFont.Bold = False
    
    
    FORMATOGRILLA(1, 1) = "EGRESO"
    FORMATOGRILLA(1, 2) = "PROVEEDOR"
    FORMATOGRILLA(1, 3) = "GLOSA"
    FORMATOGRILLA(1, 4) = "TD"
    FORMATOGRILLA(1, 5) = "NUMERO"
    FORMATOGRILLA(1, 6) = "MONTO"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "12"
    FORMATOGRILLA(2, 3) = "30"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,##0"
    
    Rem LOCCKED
    For k = 1 To 6
    FORMATOGRILLA(5, k) = "FALSE"
    Next k
    
    
    Grid2.Cols = 7
    Grid2.Rows = 1
    
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
   Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        
        Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid2.Column(k).Width = Val(FORMATOGRILLA(2, k)) * (Grid1.DefaultFont.Size - 1)
        Grid2.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid2.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid2.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
   
   
    
End Sub


Sub CARGAGRILLA3()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid3.DefaultFont.Size = 10
    Grid3.DefaultFont.Bold = False
    
    
    FORMATOGRILLA(1, 1) = "EGRESO"
    FORMATOGRILLA(1, 2) = "PROVEEDOR"
    FORMATOGRILLA(1, 3) = "GLOSA"
    FORMATOGRILLA(1, 4) = "TD"
    FORMATOGRILLA(1, 5) = "NUMERO"
    FORMATOGRILLA(1, 6) = "MONTO"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "12"
    FORMATOGRILLA(2, 3) = "30"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,##0"
    
    Rem LOCCKED
    For k = 1 To 6
    FORMATOGRILLA(5, k) = "FALSE"
    Next k
    
    
    Grid3.Cols = 7
    Grid3.Rows = 1
    
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
   Grid3.Column(0).Width = 0
    
    For k = 1 To Grid3.Cols - 1
        
        Grid3.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid3.Column(k).Width = Val(FORMATOGRILLA(2, k)) * (Grid3.DefaultFont.Size - 1)
        Grid3.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid3.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid3.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid3.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid3.Column(k).CellType = cellCalendar
        
    Next k
   Grid3.Column(6).Mask = cellNumeric
   
    
End Sub

Private Sub monto_Click()
End Sub

Private Sub leer()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim MESCONTABLE As Double
    Dim totalpago As Double
    
    Dim AÑOCONTABLE As Double
    
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + Format(MES, "00") + "-" + "01"
    fecha2 = "2100-12-01"
    
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fechavencimiento,sum(monto)"
        csql.sql = csql.sql + "FROM movimientoscontables "
        csql.sql = csql.sql + "where fechavencimiento>='" + fecha1 + "' AND fechavencimiento<='" + fecha2 + "' and (tipo='NG' ) and codigocuenta='11130001' "
        csql.sql = csql.sql + "group by fechavencimiento order by fechavencimiento"
        csql.Execute
        
                    
            
        total = 0
        total2 = 0
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         
         While Not resultados.EOF
         Grid1.Rows = Grid1.Rows + 1
         Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
         Grid1.Cell(Grid1.Rows - 1, 2).text = Format(resultados(1), "###,###,###,###")
          Grid1.Cell(Grid1.Rows - 1, 5).text = LEERfechaLiberacion(resultados(0), resultados(1))
        totalpago = totalpago + resultados(1)
        resultados.MoveNext
        
            Wend
End If
      Grid1.AutoRedraw = True
      Grid1.Refresh
      
      acumulado.Caption = Format(totalpago, "$###,###,###")
      
End Sub
Sub limpiar()


End Sub
Function LEERfechaLiberacion(vencimiento, monto) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select fecha from cartolasbancarias "
    csql.sql = csql.sql & " where (tipo='00945' or tipo='00600') and fecha>='" & Format(DateAdd("d", -15, vencimiento), "yyyy-mm-dd") & "' and monto='" & monto & "' "
    csql.Execute
    LEERfechaLiberacion = ""
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LEERfechaLiberacion = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
End Function

Sub CABEZAS2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    'Report Title 1
    If tipo = "N" Then
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
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
    End If
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        .Orientation = cellLandscape
        
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub

Sub cabezas3(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid2.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle


    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid2.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid2.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid2.ReportTitles.Add objReportTitle
        
    End If
    
With Grid2.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        .Orientation = cellLandscape
        
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub


Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub
Sub eliminafactura(tipo, numero)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_pagos_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM l_movimientos_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, gestionrubro, "")
        
        
End Sub


Private Sub Grid1_DblClick()
If Option1.Value = True Then
Call LEERNOMINA(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
Else
Call LEERNOMINA2(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'End If
'leer
End Sub
Public Function leefactura(tipo, numero, rut) As String

    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    If tipo = "FA" Then tipo = "1"
    If tipo = "ND" Then tipo = "2"
    If tipo = "NC" Then tipo = "3"
    If tipo = "FAE" Then tipo = "4"
    If tipo = "NDE" Then tipo = "5"
    If tipo = "NCE" Then tipo = "6"
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
    campos(0, 2) = "facturasdecompras"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leefactura = "1"
    
    Else
    leefactura = "0"
    
    End If
    
    

End Function

Sub crearcuentacorriente(rut)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion

            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,mc.giro,mc.fono1 "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            


End Sub
'cSql.SQL = "INSERT INTO l_movimientos_detalle_" & empresaactiva & " "
'            cSql.SQL = cSql.SQL & "(tipo, numero, linea, fecha, rut, codigo, descripcion, cantidad, unidades, precio, total, costoventa, bodega, bodegatraspaso, uxc) "
'            cSql.SQL = cSql.SQL & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.unidades, dd.precio, dd.total, dd.pcosto, dd.bodega, dd.bodega, ROUND(dd.unidades / dd.cantidad, 0) "
'            cSql.SQL = cSql.SQL & "FROM " & baseVentas & rubro & ".sv_documento_detalle_" + empresaactiva + " as dd "
'            cSql.SQL = cSql.SQL & "WHERE dd.local = '" & empresaactiva & "' AND dd.tipo = '" & v.detalle.tipo & "' AND dd.numero = '" & v.detalle.numero & "'"
'            cSql.Execute

Public Function LEERULTIMOFOLIO(mesconta, añoconta) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select max(folio) from facturasdecompras where mescontable = '" & Format(mesconta, "00") & "' AND añocontable = '" & añoconta & "' "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
        If resultados(0) <> "NULO" Then
        LEERULTIMOFOLIO = resultados(0) + 1
        Else
        LEERULTIMOFOLIO = "0000000001"
        End If
        
    End If
    
End Function
Public Sub LEERNOMINA(fecha)
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
             
            
            csql.sql = "select mc.numero,mc.rutproveedor, (select cc.nombre from cuentascorrientes as cc where mc.rutproveedor=cc.rut and "
            csql.sql = csql.sql & "(cc.tipo='" + CUENTAPROVEEDOR + "' or cc.tipo='23100029' or cc.tipo='47150022') and cc.año='" + Mid(COMBOAÑO.text, 1, 4) + "' limit 0,1) as nombre,"
            csql.sql = csql.sql & "sum(mc.monto),mc.mes,mc.año from movimientoscontables as mc  "
            csql.sql = csql.sql & "Where mc.codigocuenta='11130001' and fechavencimiento='" & Format(fecha, "yyyy-mm-dd") & "' "
            csql.sql = csql.sql & "and (mc.tipo='NG')  group by rutproveedor order by nombre "
            csql.Execute
            
            
    Grid2.Rows = 1
    montonomina = 0
    If csql.RowsAffected > 0 Then
       
    Set resultados = csql.OpenResultset
          While resultados.EOF = False
           montonomina = montonomina + resultados(3)
           Grid2.Rows = Grid2.Rows + 1
           Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0)
           Grid2.Cell(Grid2.Rows - 1, 2).text = Mid(resultados(1), 1, 9) + "-" + Mid(resultados(1), 10, 1)
           Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
           Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(3)
           
           resultados.MoveNext
           Wend
           
           
    End If
           
           Grid2.Rows = Grid2.Rows + 1
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).FontBold = True
           Grid2.Cell(Grid2.Rows - 1, 3).text = "TOTAL NOMINA " + Format(fecha, "dd-mm-yyyy")
           Grid2.Cell(Grid2.Rows - 1, 6).text = Format(montonomina, "###,###,###,###")
          
Grid2.Rows = Grid2.Rows + 4
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).FontBold = True
           Grid2.Cell(Grid2.Rows - 1, 3).text = "VB GERENCIA"
                    
          FECHANOMINA = fecha
          
End Sub
Public Sub LEERNOMINA2(fecha)
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim prove As String
    Dim numero As String
    Dim tipodoc As String
    
    GENERANOMINA = True
    
        Set csql.ActiveConnection = contadb
            csql.sql = "select mc.numero,mc.rutproveedor,cc.nombre, mc.monto,mc.mes,mc.año,mc.tipo from movimientoscontables as mc,cuentascorrientes as cc "
            csql.sql = csql.sql & "Where mc.codigocuenta='11130001' and mc.rutproveedor=cc.rut and (cc.tipo='" + CUENTAPROVEEDOR + "' or cc.tipo='23100029' or cc.tipo='47150022' ) and fechavencimiento='" & Format(fecha, "yyyy-mm-dd") & "' and (mc.tipo='NG' )  group by mc.rutproveedor,mc.numero order by cc.nombre "
            csql.Execute
            
    Grid2.Rows = 1
    montonomina = 0
    montopago = 0
    LINEAS = 0
       lineacom = 0
    If csql.RowsAffected > 0 Then
    
    Set resultados = csql.OpenResultset
    prove = resultados(1)
    numero = resultados(0)
    tipodoc = resultados(6)
          While resultados.EOF = False
          
           montonomina = montonomina + resultados(3)
           If prove <> resultados(1) Or numero <> resultados(0) Or LINEA > 12 Then
           Grid2.Rows = Grid2.Rows + 1
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).FontBold = True
           Grid2.Cell(Grid2.Rows - 1, 1).text = "PROVEEDOR"
           Grid2.Cell(Grid2.Rows - 1, 2).text = Mid(prove, 1, 9) + "-" + Mid(prove, 10, 1)
           Grid2.Cell(Grid2.Rows - 1, 3).text = glosapago
           
           Grid2.Cell(Grid2.Rows - 1, 6).text = montopago * -1
           
           Grid2.Rows = Grid2.Rows + 1
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).Borders(cellEdgeBottom) = cellThin
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).FontBold = True
                       
           If resultados(6) = "PS" Then
                Call leedatoscuentaTrabajador(prove, resultados(2))
           Else
               Call leedatoscuenta(prove)
'                Call leedatoscuentaTrabajador
           End If
           
           
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 1).Merge
           Grid2.Range(Grid2.Rows - 1, 4, Grid2.Rows - 1, 6).Merge
           
           Grid2.Cell(Grid2.Rows - 1, 1).text = banco
           Grid2.Cell(Grid2.Rows - 1, 2).text = lineacom
           Grid2.Cell(Grid2.Rows - 1, 3).text = cuentacorriente
           Grid2.Cell(Grid2.Rows - 1, 4).text = email
           If lineacom > 11 Then
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 4).BackColor = &HFF&
           GENERANOMINA = False
           End If
           
           
           numeropago = ""
            montopago = 0
            prove = resultados(1)
          numero = resultados(0)
          tipodoc = resultados(6)
          LINEAS = 0
   lineacom = 0
           End If
            
            Call LEERdetalle(resultados(0), resultados(4), resultados(5), resultados(6))
            
            
           
           resultados.MoveNext
           If resultados.EOF = False Then
            LINEAS = LINEAS + cuentalineas(resultados(0), resultados(4), resultados(5), resultados(6))
            If LINEAS > 11 Then
            revisanumero = True
            Else
            revisanumero = False
            End If
            
            If revisanumero = False Then
            numero = resultados(0)
            End If
           
           End If
           
           Wend
           
           
    End If
           
           Grid2.Rows = Grid2.Rows + 1
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).FontBold = True
           Grid2.Cell(Grid2.Rows - 1, 1).text = "PROVEEDOR"
           Grid2.Cell(Grid2.Rows - 1, 2).text = Mid(prove, 1, 9) + "-" + Mid(prove, 10, 1)
           Grid2.Cell(Grid2.Rows - 1, 3).text = glosapago
           Grid2.Cell(Grid2.Rows - 1, 6).text = montopago * -1
           Grid2.Rows = Grid2.Rows + 1
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).Borders(cellEdgeBottom) = cellThin
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).FontBold = True
                       
'           Call leedatoscuenta(prove)
           
            If tipodoc = "PS" Then
                Call leedatoscuentaTrabajador(prove, glosapago)
           Else
               Call leedatoscuenta(prove)
'                Call leedatoscuentaTrabajador
           End If
           
           
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 2).Merge
           Grid2.Range(Grid2.Rows - 1, 4, Grid2.Rows - 1, 6).Merge
           
           Grid2.Cell(Grid2.Rows - 1, 1).text = banco
           Grid2.Cell(Grid2.Rows - 1, 2).text = lineacom
           Grid2.Cell(Grid2.Rows - 1, 3).text = cuentacorriente
           Grid2.Cell(Grid2.Rows - 1, 4).text = email
            
            montopago = 0
           
            
           Grid2.Rows = Grid2.Rows + 1
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).FontBold = True
           Grid2.Cell(Grid2.Rows - 1, 3).text = "TOTAL NOMINA " + Format(fecha, "dd-mm-yyyy")
           Grid2.Cell(Grid2.Rows - 1, 6).text = Format(montonomina, "###,###,###,###")
          
Grid2.Rows = Grid2.Rows + 4
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
           Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 6).FontBold = True
           Grid2.Cell(Grid2.Rows - 1, 3).text = "VB GERENCIA"
          FECHANOMINA = fecha
                    
          
End Sub


Public Sub LEERdetalle(numero, MES, año, tipo)
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
            
            csql.sql = "select mc.numero,mc.rutproveedor,mc.glosacontable,mc.tipodocumento,mc.numerodocumento,if (mc.dh='D',mc.monto,mc.monto*-1),codigocuenta from movimientoscontables as mc "
            csql.sql = csql.sql & "Where mc.tipo='" & tipo & "' and mc.numero='" + numero + "' and mes='" + MES + "' and año='" + año + "'  order by linea "
            csql.Execute
            
    If csql.RowsAffected > 0 Then
   
    Set resultados = csql.OpenResultset
          While resultados.EOF = False
           
           
           If resultados(6) = "11130001" Then
           glosapago = resultados(2)
           montopago = montopago + CDbl(resultados(5))
           numeropago = ""
           Else
           lineacom = lineacom + 1
           LINEAS = LINEAS + 1
           Grid2.Rows = Grid2.Rows + 1
           
           Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0)
           Grid2.Cell(Grid2.Rows - 1, 2).text = Mid(resultados(1), 1, 9) + "-" + Mid(resultados(1), 10, 1)
           Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
           Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
           Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
           Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
           
           End If
           
           
           resultados.MoveNext
           Wend
           
           
    End If
    
End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Static palabra As String
    Dim i As Integer
    Dim largo As Integer
    If KeyAscii = 13 Then
        palabra = ""
    Else
        palabra = palabra + UCase(Chr(KeyAscii))
        largo = Len(palabra)
        For i = 1 To Grid1.Rows - 1
            If Mid(Grid1.Cell(i, 16).text, 1, largo) = palabra Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                Exit For
            End If
        Next i
    End If
    
End Sub

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
Call LEERNOMINA(Grid1.Cell(NewRow, 1).text)
End Sub

Private Sub Grid2_DblClick()
If oprut.Value = True Then
If Grid2.Cell(Grid2.ActiveCell.row, 1).text = "PROVEEDOR" Then
datoprove.text = Grid2.Cell(Grid2.ActiveCell.row, 2).text
End If



End If

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Call LEERNOMINA(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
Else
Call LEERNOMINA2(Grid1.Cell(Grid1.ActiveCell.row, 1).text)

End If

End Sub

Private Sub Option2_Click()
If Option1.Value = True Then
Call LEERNOMINA(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
Else
Call LEERNOMINA2(Grid1.Cell(Grid1.ActiveCell.row, 1).text)

End If

End Sub
Public Sub leedatoscuenta(ByVal rut As String)

        Dim op As Integer
        Dim condicion As String
        
        campos(0, 0) = "banco"
        campos(1, 0) = "sucursal"
        campos(2, 0) = "cuentacorriente"
        campos(3, 0) = "email"
        campos(4, 0) = "modopago"
        campos(5, 0) = "rutretira"
        campos(6, 0) = "nombreretira"
        campos(7, 0) = ""
        campos(0, 2) = "cuentascorrientes_datos_pago"

        condicion = "rut = '" & rut & "' "

        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
        banco = sqlconta.response(0, 3)
        SUCURSAL = sqlconta.response(1, 3)
        cuentacorriente = sqlconta.response(2, 3)
        email = sqlconta.response(3, 3)
        RUTRETIRA2 = sqlconta.response(5, 3)
        NOMBRERETIRA2 = sqlconta.response(6, 3)
        modopago = sqlconta.response(4, 3)
        Else
        banco = ""
        SUCURSAL = ""
        cuentacorriente = ""
        email = ""
        RUTRETIRA2 = ""
        NOMBRERETIRA2 = ""
        modopago = ""
        End If
End Sub

Public Sub leedatoscuentaTrabajador(ByVal rut As String, NOMBRE)

        Dim op As Integer
        Dim condicion As String
        
        campos(0, 0) = "banco"
        campos(1, 0) = "cuenta"
        campos(2, 0) = "rut"
        campos(3, 0) = ""

        campos(0, 2) = clientesistema & "remu.cuentasbancarias"

        condicion = "rut = '" & rut & "' "

        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
            banco = sqlconta.response(0, 3)
            SUCURSAL = ""
            cuentacorriente = sqlconta.response(1, 3)
            email = ""
            RUTRETIRA2 = sqlconta.response(2, 3)
            NOMBRERETIRA2 = NOMBRE
            modopago = ""
        Else
            banco = ""
            SUCURSAL = ""
            cuentacorriente = ""
            email = ""
            RUTRETIRA2 = ""
            NOMBRERETIRA2 = ""
            modopago = ""
        End If
End Sub


Sub GENERATXT()
Dim k As Double
Dim s As Double
Dim L As Double

Dim MATRIX(11, 3) As String
Dim contador As Double
Dim signo As String
Dim VARIABLE As String
Dim BLANCO As String * 9
Dim rut As String * 12
Dim NOMBRE As String * 40
Dim direccion As String * 40
Dim comuna As String * 15
Dim ciudad As String * 15

Dim MODALIDAD As String * 2
Dim SUCURSAL As String * 3
Dim cuenta As String * 18
Dim banco As String * 4
Dim RUTRETIRA As String * 12
Dim NOMBRERETIRA As String * 40
Dim numerodo As String * 8
Dim MONTODO As String * 11
Dim SIGNODO As String * 1
Dim TOTALGENERAL As String * 13
Dim sumatotal As Double

 On Error GoTo no:
 FECHANOMINA = Replace(FECHANOMINA, "/", "-")
Close 20
Open "c:\nominas\NOMINA_" + empresaactiva + "_" + FECHANOMINA + ".TXT" For Output As #20

contador = 0
For k = 1 To Grid2.Rows - 1
If Grid2.Cell(k, 1).text = "PROVEEDOR" Then
VARIABLE = ""
rut = "00" + Mid(Grid2.Cell(k, 2).text, 1, 9) + Mid(Grid2.Cell(k, 2).text, 11, 1)
NOMBRE = Grid2.Cell(k, 3).text
TOTALGENERAL = Format(CDbl(Grid2.Cell(k, 6).text), "0000000000000")
banco = Format(CDbl(Grid2.Cell(k + 1, 1).text), "000 ")
SUCURSAL = "999"
cuenta = Format(CDbl(Grid2.Cell(k + 1, 3).text), "000000000000000")

direccion = String(40, 32)
comuna = String(40, 32)
ciudad = String(40, 32)
Call leedatoscuenta(Mid(Grid2.Cell(k, 2).text, 1, 9) + Mid(Grid2.Cell(k, 2).text, 11, 1))
If modopago = "3" Then
RUTRETIRA = String(12, 32)
NOMBRERETIRA = String(40, 32)
Else
RUTRETIRA = RUTRETIRA2
NOMBRERETIRA = NOMBRERETIRA2


End If
MODALIDAD = "0" + modopago


VARIABLE = rut + NOMBRE + direccion + comuna + ciudad + SUCURSAL + MODALIDAD + banco + cuenta + RUTRETIRA + NOMBRERETIRA
If rut = "000827834009" Then
Print "hola"

End If
For s = 1 To 11
If MATRIX(s, 1) <> "" Then
numerodo = Format(MATRIX(s, 1), "00000000")
MONTODO = Format(MATRIX(s, 2), "00000000000") + "  "
SIGNODO = MATRIX(s, 3)
Else
numerodo = "00000000"
MONTODO = "00000000000"
SIGNODO = "0"
End If
VARIABLE = VARIABLE + numerodo + SIGNODO + MONTODO
Next s
VARIABLE = VARIABLE + TOTALGENERAL
Print #20, VARIABLE

contador = 0
For L = 1 To 11
MATRIX(L, 1) = ""
MATRIX(L, 2) = ""
MATRIX(L, 3) = ""

Next L

End If

signo = "+"
If Mid(Grid2.Cell(k, 1).text, 1, 3) = "000" Then

    contador = contador + 1
    
    MATRIX(contador, 1) = Mid(Grid2.Cell(k, 5).text, 3, 8)
    signo = "+"
    MATRIX(contador, 2) = Grid2.Cell(k, 6).text
    MATRIX(contador, 3) = signo

    If Grid2.Cell(k, 6).text < "0" Then
    signo = "-"
    MATRIX(contador, 2) = CDbl(Grid2.Cell(k, 6).text) * -1
    MATRIX(contador, 3) = signo
    End If
End If

Next k
Close 20

Shell "NOTEPAD c:\nominas\NOMINA_" + empresaactiva + "_" + FECHANOMINA + ".TXT"
Exit Sub
no:
If error = "No se ha encontrado la ruta de acceso" Then
    MsgBox " DEBE CREAR CARPETA c:\nominas"
Else
    MsgBox "ERROR: " & error
End If

End Sub
Function cuentalineas(numero, MES, año, tipo) As Double

    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
            
            csql.sql = "select mc.numero,mc.rutproveedor,mc.glosacontable,mc.tipodocumento,mc.numerodocumento,if (mc.dh='D',mc.monto,mc.monto*-1),codigocuenta from movimientoscontables as mc "
            csql.sql = csql.sql & "Where mc.tipo='" & tipo & "' and mc.numero='" + numero + "' and mes='" + MES + "' and año='" + año + "' and mc.codigocuenta<>'11130001'  order by linea "
            csql.Execute
    cuentalineas = 0
    If csql.RowsAffected > 0 Then
    cuentalineas = csql.RowsAffected
    
   
    End If
    
End Function

Sub cabezas(cliente)
Dim objReportTitle As FlexCell.ReportTitle
Dim o As Integer

Grid3.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "COMPROBANTE CANCELACION DE FACTURA "
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid3.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = cliente
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid3.ReportTitles.Add objReportTitle
    
    'Report Title 1
        For o = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(o)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid3.ReportTitles.Add objReportTitle
    Next o
    
With Grid3.PageSetup
        
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

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
