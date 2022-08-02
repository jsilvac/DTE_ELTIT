VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LibroVentasvendedores 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadisticas de Ventas por Vendedores"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   14565
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   120
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   7140
      Left            =   90
      TabIndex        =   1
      Top             =   2100
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   12594
      BackColor       =   16744576
      Caption         =   "Informe"
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
      Begin FlexCell.Grid impresion 
         Height          =   6780
         Left            =   45
         TabIndex        =   2
         Top             =   360
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   11959
         AllowUserReorderColumn=   -1  'True
         AllowUserSort   =   -1  'True
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   1
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1950
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   3440
      BackColor       =   16744576
      Caption         =   "Ingreso de Información"
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EXPORTAR EXCEL"
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
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1620
         Width           =   2535
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   960
         Left            =   4770
         TabIndex        =   5
         Top             =   360
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1693
         BackColor       =   16711680
         Caption         =   "FECHA DE CONSULTA"
         CaptionEstilo3D =   1
         BackColor       =   16711680
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
            Height          =   315
            Left            =   180
            MaxLength       =   2
            TabIndex        =   11
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
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
            Height          =   315
            Left            =   630
            MaxLength       =   2
            TabIndex        =   10
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
         Begin VB.TextBox dato3 
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
            Height          =   315
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   9
            Tag             =   "proveedor"
            Top             =   540
            Width           =   705
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
            Height          =   315
            Left            =   2790
            MaxLength       =   4
            TabIndex        =   8
            Tag             =   "proveedor"
            Top             =   540
            Width           =   705
         End
         Begin VB.TextBox dato4 
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
            Height          =   315
            Left            =   1890
            MaxLength       =   2
            TabIndex        =   7
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
         Begin VB.TextBox dato5 
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
            Height          =   315
            Left            =   2340
            MaxLength       =   2
            TabIndex        =   6
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
         Begin VB.Label lbl3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Hasta"
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
            Height          =   315
            Left            =   1890
            TabIndex        =   13
            Top             =   270
            Width           =   1605
         End
         Begin VB.Label lbl2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Desde"
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
            Height          =   315
            Left            =   180
            TabIndex        =   12
            Top             =   270
            Width           =   1605
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   1230
         Left            =   8505
         TabIndex        =   4
         Top             =   360
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   2170
         BackColor       =   16761024
         Caption         =   "TIPO CONSULTA"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.OptionButton TIPO2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "x Articulos"
            Height          =   375
            Left            =   135
            TabIndex        =   19
            Top             =   675
            Width           =   1635
         End
         Begin VB.OptionButton TIPO1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "x Documentos"
            Height          =   330
            Left            =   135
            TabIndex        =   18
            Top             =   270
            Value           =   -1  'True
            Width           =   1365
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   660
         Left            =   45
         TabIndex        =   14
         Top             =   1200
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   1164
         BackColor       =   16744576
         Caption         =   "VENDEDOR"
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
         Begin VB.ComboBox COMBOVENDEDOR 
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   270
            Width           =   4485
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   675
         Left            =   45
         TabIndex        =   16
         Top             =   360
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   "LOCAL"
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
         Begin VB.ComboBox combolocal 
            Height          =   315
            Left            =   45
            TabIndex        =   17
            Top             =   270
            Width           =   4485
         End
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1230
         Left            =   10170
         TabIndex        =   20
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   2170
         BackColor       =   16761024
         Caption         =   "VISTA CONSULTA"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.OptionButton VISTA1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Detallada"
            Height          =   330
            Left            =   90
            TabIndex        =   22
            Top             =   270
            Width           =   1950
         End
         Begin VB.OptionButton VISTA2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Acumulada"
            Height          =   375
            Left            =   90
            TabIndex        =   21
            Top             =   675
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Genera Informes"
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1620
         Width           =   2535
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   660
         Left            =   4750
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   1164
         BackColor       =   16744576
         Caption         =   "CODIGO PRODUCTO"
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
         Begin VB.TextBox CODIGO 
            Alignment       =   1  'Right Justify
            Height          =   365
            Left            =   720
            MaxLength       =   13
            TabIndex        =   25
            Top             =   240
            Width           =   2400
         End
      End
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   1230
         Left            =   13080
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   2170
         BackColor       =   16761024
         Caption         =   "RANKING"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Atenciones"
            Height          =   375
            Left            =   90
            TabIndex        =   29
            Top             =   675
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Venta"
            Height          =   330
            Left            =   90
            TabIndex        =   28
            Top             =   270
            Value           =   -1  'True
            Width           =   1950
         End
      End
      Begin XPFrame.FrameXp FrameXp9 
         Height          =   1230
         Left            =   11760
         TabIndex        =   30
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   2170
         BackColor       =   16761024
         Caption         =   "VALORIZADO"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.OptionButton IVACON 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Con Iva"
            Height          =   330
            Left            =   90
            TabIndex        =   32
            Top             =   240
            Value           =   -1  'True
            Width           =   1950
         End
         Begin VB.OptionButton IVACON1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Sin I.V.A"
            Height          =   375
            Left            =   90
            TabIndex        =   31
            Top             =   675
            Width           =   2175
         End
      End
      Begin VB.Label lblayuda 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 AYUDA PRODUCTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   5715
      TabIndex        =   3
      Top             =   9315
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "I   M   P   R   I   M   I   R"
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
End
Attribute VB_Name = "LibroVentasvendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private TIPO As String
    Private detalle As Boolean
    Private fecha1 As String
    Private fecha2 As String
    Private codigoempresa As String
    Private codigovendedor As String
    



Private Sub CODIGO_GotFocus()
lblayuda.Visible = True

End Sub

Private Sub CODIGO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
 Call ayudaProductotxt(CODIGO)
 End If
 
End Sub

Private Sub CODIGO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
CODIGO.text = ceros(CODIGO)
Command1.SetFocus
End If

End Sub

Private Sub CODIGO_LostFocus()
lblayuda.Visible = False

End Sub

Private Sub combolocal_Click()

Call conecntarVentasAuditoria(servidor, baseVentas & Mid(combolocal.text, 1, 2), usuario, password)

LEErVENDEDORES
End Sub



Private Sub COMBOVENDEDOR_Click()
Command1_Click
End Sub

Private Sub Command1_Click()
codigoempresa = Mid(combolocal.text, 1, 2)
codigovendedor = Mid(COMBOVENDEDOR.text, 1, 10)

            If TIPO1.Value = True Then
            Call CargaGrillaInforme(1, 13)
            'Call CargaGrillaInformeventasxvendedor(1, 7)
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            Call generaInformeLV(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
            End If

            If TIPO2.Value = True Then
            Call CargaGrillaInforme2(1, 8)
            
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            Call generaInformevp(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
           End If
           
End Sub

Private Sub Command2_Click()
impresion.ExportToExcel ("")
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
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
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
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
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
           dato2.SetFocus
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
           dato3.SetFocus
        End If
    End Sub
        
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
           dato4.SetFocus
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            dato5.SetFocus
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            dato6.SetFocus
        End If
    End Sub
        
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
        Command1.SetFocus
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
'    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato1.text) = dato1.MaxLength Then
'            Call dato1_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato2.text) = dato2.MaxLength Then
'            Call dato2_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato3.text) = dato3.MaxLength Then
'            Call dato3_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato4_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato4.text) = dato4.MaxLength Then
'            Call dato4_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato5_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato5.text) = dato5.MaxLength Then
'            Call dato5_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato6_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato6.text) = dato6.MaxLength Then
'            Call dato6_KeyPress(13)
'        End If
'    End Sub
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
 
    Private Sub dato1_LostFocus()
    Call limpiaBarra(2)
    Call esfecha(dato1, dato2, dato3, "dd")
    End Sub
    Private Sub dato2_LostFocus()
    Call esfecha(dato1, dato2, dato3, "mm")
    End Sub
    Private Sub dato3_LostFocus()
    Call esfecha(dato1, dato2, dato3, "yyyy")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato4, dato5, dato6, "dd")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato4, dato5, dato6, "mm")
    End Sub
    Private Sub dato6_LostFocus()
    Call esfecha(dato4, dato5, dato6, "yyyy")
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 27
                Unload Me
            Case 38
                If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                    Unload Me
                End If
        End Select
    End Sub
    
    Private Sub Form_Load()
        Call Centrar(Me)
       
        
        TIPO = "(dc.tipo = 'FV')"
        detalle = False
        dato1.text = Format(fechasistema, "dd")
        dato2.text = Format(fechasistema, "mm")
        dato3.text = Format(fechasistema, "yyyy")
        dato4.text = Format(fechasistema, "dd")
        dato5.text = Format(fechasistema, "mm")
        dato6.text = Format(fechasistema, "yyyy")
    LEErlocales
    Call conecntarVentasAuditoria(servidor, baseVentas & Mid(combolocal.text, 1, 2), usuario, password)

    LEErVENDEDORES
    Call CargaGrillaInforme(1, 13)
    'Call CargaGrillaInformeventasxvendedor(1, 7)
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatogrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "TIPO"
        formatogrilla(1, 2) = "NUMERO"
        formatogrilla(1, 3) = "FECHA"
        formatogrilla(1, 4) = "RUT"
        formatogrilla(1, 5) = "CLIENTE"
        formatogrilla(1, 6) = "TOTAL"
        formatogrilla(1, 7) = "N/CREDITO"
        formatogrilla(1, 8) = "V.LIQUIDA"
        formatogrilla(1, 9) = "DESC"
        formatogrilla(1, 10) = "Nº DOC"
        formatogrilla(1, 11) = "CANT PRODUCTOS"
        formatogrilla(1, 12) = "% PARTI"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "4"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "8"
        formatogrilla(2, 4) = "10"
        formatogrilla(2, 5) = "25"
        formatogrilla(2, 6) = "9"
        formatogrilla(2, 7) = "9"
        formatogrilla(2, 8) = "9"
        formatogrilla(2, 9) = "9"
        formatogrilla(2, 10) = "6"
        formatogrilla(2, 11) = "9"
        formatogrilla(2, 12) = "9"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "C"
        formatogrilla(3, 3) = "D"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        formatogrilla(3, 11) = "N"
        formatogrilla(3, 12) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = "###,###,##0"
        formatogrilla(4, 7) = "###,###,##0"
        formatogrilla(4, 8) = "###,###,##0"
        
        formatogrilla(4, 9) = "% ###,##0.00"
        formatogrilla(4, 10) = "###,###,##0"
        formatogrilla(4, 11) = "###,###,##0"
        formatogrilla(4, 12) = "% #,###,##0.00"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        formatogrilla(6, 10) = ""
        formatogrilla(6, 11) = ""
        formatogrilla(6, 12) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        Rem ANCHO
        formatogrilla(8, 1) = "3"
        formatogrilla(8, 2) = "8"
        formatogrilla(8, 3) = "8"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "20"
        formatogrilla(8, 6) = "8"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
        
        formatogrilla(8, 10) = "6"
        formatogrilla(8, 11) = "8"
        formatogrilla(8, 12) = "6"
        
                
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.AllowUserSort = True
        
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
        impresion.SelectionMode = cellSelectionNone
        impresion.Column(0).Width = 0
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1.75
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
  
     
                    
    End Sub
'    Private Sub impresion_CellChange(ByVal Row As Long, ByVal Col As Long)
'    If Row >= impresion.FixedRows Then
'        impresion.Cell(Row, 6).Refresh
'    End If
'End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
'formato grilla ventas por vendedor
Private Sub CargaGrillaInformeventasxvendedor(ByVal row As Integer, ByVal col As Integer)
        Dim formatogrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "VENDEDOR"
        formatogrilla(1, 3) = "NETO"
        formatogrilla(1, 4) = "I.V.A"
        formatogrilla(1, 5) = "EXENTO"
        formatogrilla(1, 6) = "TOTAL"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "5"
        formatogrilla(2, 2) = "20"
        formatogrilla(2, 3) = "9"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "9"
        formatogrilla(2, 6) = "9"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "$ ###,###,##0"
        formatogrilla(4, 4) = "$ ###,###,##0"
        formatogrilla(4, 5) = "$ ###,###,##0"
        formatogrilla(4, 6) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        Rem ANCHO
        formatogrilla(8, 1) = "5"
        formatogrilla(8, 2) = "20"
        formatogrilla(8, 3) = "8"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "8"
        formatogrilla(8, 6) = "8"
        
                
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
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
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1.75
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub
    
    'fin configura grilla ventas x vendedor

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        If impresion.Rows > 1 Then
        Call imprimir
        End If
        
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        Dim glosaiva As String
        
        If IVACON.Value = True Then
        glosaiva = " CON IVA"
        Else
        glosaiva = " SIN IVA"
        
        End If
        
        Call Titulos("Ventas x Vendedores " + glosaiva)
        impresion.AutoRedraw = False
        impresion.Range(1, 1, 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThick
        impresion.PageSetup.HeaderMargin = 2
        impresion.PageSetup.TopMargin = 3
        impresion.PageSetup.LeftMargin = 0.5
        impresion.PageSetup.RightMargin = 0.5
        impresion.PageSetup.BottomMargin = 3
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.Orientation = cellLandscape
        
        impresion.PageSetup.PrintFixedRow = True
        Call verificaImpresora(5, impresion)
        impresion.AutoRedraw = True
    End Sub
    
    
Sub LEErVENDEDORES()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventas
        COMBOVENDEDOR.Clear
        csql.sql = "SELECT rut,nombre "
        csql.sql = csql.sql + "FROM sv_maestrovendedores "
        csql.sql = csql.sql + "ORDER BY nombre "
        csql.Execute
        COMBOVENDEDOR.AddItem ("99" + "  TODOS LOS VENDEDORES")
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                COMBOVENDEDOR.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            
            COMBOVENDEDOR.text = COMBOVENDEDOR.List(0)
            
        End If
        
End Sub

Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                combolocal.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
           combolocal.AddItem ("99" + "  TODOS LOS LOCALES")
                
        combolocal.text = combolocal.List(CDbl(empresaActiva))
        End If
        
End Sub

Sub LEErventas()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                combolocal.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
           combolocal.AddItem ("99" + "  TODOS LOS LOCALES")
                
        combolocal.text = combolocal.List(CDbl(empresaActiva))
        End If
        
End Sub

Public Sub generaInformeLV(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Dim documento As String
    Dim glosaiva As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    If IVACON.Value = True Then
    glosaiva = " CON IVA"
    Else
    glosaiva = " SIN IVA"
    
    End If
    
    
    Call cargaCabeza("LISTADO VENTAS X VENDEDORES DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy") & glosaiva, empresaActiva, impresion)
    Call resumenVentas(data, impresion, TIPO, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Function resumenVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim FILTRO As String
    
    Dim tabla As String
    Dim rubAux As String
    Dim harinas As Double
    Dim subproductos As Double
    Dim envases As Double
    Dim trigo As Double
    Dim maquila As Double
    Dim otros As Double
    Dim cadena As String
    Dim tipoDoc As String
    Dim numeroDoc As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim i As Integer
    Dim totales(10) As Double
    Dim totales2(10) As Double
    Dim conta As Double
    Dim totalventa As Double
    Dim cSql2 As New rdoQuery
    Dim resultados As rdoResultset
        Set csql.ActiveConnection = ventasAuditoria
    Dim iva As String
    If IVACON.Value = True Then
    iva = "1"
    Else
    iva = "1.19"
    End If
    
    
    
    Rem calcula total venta
        csql.sql = "SELECT sum(dd.total)/" + iva + " "
        csql.sql = csql.sql + "FROM sv_documento_detalle_" + codigoempresa + " as dd "
        csql.sql = csql.sql + "where dd.fecha BETWEEN '" + fecha1 + "' AND '" + fecha2 + "' "
        csql.sql = csql.sql + "AND ( dd.tipo<>'PV' AND dd.tipo<>'NP' AND dd.tipo<>'CO' and dd.tipo<>'NB' and dd.tipo<>'NF') and caja<'90' "
        csql.Execute
        impresion.AutoRedraw = False
        
        
     If csql.RowsAffected > 0 Then
     Set resultado = csql.OpenResultset
    If Not IsNull(resultado(0)) Then
    totalventa = resultado(0)
    End If
    End If
    Call LEErnotasdecredito(fecha1, fecha2, Mid(combolocal.text, 1, 2))
    
    
    If VISTA1.Value = True Then
    
    Set cSql2.ActiveConnection = ventasAuditoria
    rubAux = rubro
    tabla = "SELECT dd.tipo, dd.numero , dd.fecha, dd.rut, mc.nombre,sum(if(tipo='FV' or tipo='BV',dd.total,0)),sum(if(tipo='NB' or tipo='NF',dd.total,0)),sum(if(tipo='FV' or tipo='BV',dd.total,0))-sum(if(tipo='NB' or tipo='NF',dd.total,0)), dd.descuento, sum(dd.cantidad), dd.vendedor,count(dd.numero) "
    tabla = tabla & "FROM sv_documento_detalle_" & codigoempresa & " as dd INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON dd.rut = mc.rut AND mc.sucursal = '0'"
    tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and (dd.tipo='FV' or dd.tipo='BV' or dd.tipo='NB' or dd.tipo='NF') AND dd.local='" + codigoempresa + "' and caja<'90' "
    
    If codigovendedor <> "99  TODOS " Then
    tabla = tabla & "and dd.vendedor='" + codigovendedor + "' "
    End If
   
    
     tabla = tabla & "group by dd.numero,dd.vendedor ORDER BY dd.vendedor,mc.nombre,dd.numero "
   
    cSql2.sql = tabla
    cSql2.Execute
    Else
        Set cSql2.ActiveConnection = ventasAuditoria
        rubAux = rubro
        
        tabla = "SELECT dd.tipo, dd.numero , dd.fecha, dd.rut, '',"
        tabla = tabla & "sum(if(dd.tipo='BV' OR dd.tipo='FV',dd.total,0))/1,"
        tabla = tabla & "sum(if(dd.tipo='NB' OR dd.tipo='NF',dd.total,0))/1,"
        tabla = tabla & "(sum(if(dd.tipo='BV' OR dd.tipo='FV',dd.total,0))/1 - sum(if(dd.tipo='NB' OR dd.tipo='NF',dd.total,0))/1) as total,"
        tabla = tabla & "sum(dd.precio*(1+dd.descuento)/100) as descuento, "
        tabla = tabla & "sum(dd.cantidad)as cantidad, dd.vendedor,count(dc.numero) "
        tabla = tabla & "FROM sv_documento_cabeza_" & codigoempresa & " as dc "
        tabla = tabla & "INNER JOIN sv_documento_detalle_" & codigoempresa & " AS dd "
        tabla = tabla & "ON dc.local=dd.local and dc.numero=dd.numero and "
        tabla = tabla & "dc.fecha=dd.fecha and dc.caja=dd.caja and dc.tipo=dd.tipo "
        tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
        tabla = tabla & "and (dd.tipo='FV' or dd.tipo='BV' or dd.tipo='NB' or dd.tipo='NF' ) AND dd.local='" & codigoempresa & "' "
        tabla = tabla & "and dc.nula<>'S' and dc.caja<'90' group by dd.vendedor ORDER BY  "
        tabla = tabla & "total desc"
        cSql2.sql = tabla
        cSql2.Execute
    End If
    
    'Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    On Local Error Resume Next
    If cSql2.RowsAffected > 0 Then
    Set resultados = cSql2.OpenResultset
    
       FILTRO = resultados("vendedor")
       impresion.Rows = 2
        While Not resultados.EOF
           If FILTRO <> resultados("vendedor") Then
           linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           impresion.Column(4).Locked = False
           impresion.Column(5).Locked = False
''           impresion.Column(1).Locked = False
           
           
          
           impresion.Cell(linea, 4).text = FILTRO
           impresion.Cell(linea, 5).text = leerNombreVendedor(FILTRO)
           
           impresion.Cell(linea, 6).text = totales(1)
           impresion.Cell(linea, 7).text = totales(2)
           impresion.Cell(linea, 8).text = totales(3)
           
           impresion.Cell(linea, 9).text = (totales(4) / totales(1)) * 100
           impresion.Cell(linea, 10).text = totales(5)
           impresion.Cell(linea, 11).text = totales(6)
           impresion.Cell(linea, 12).text = ((totales(1) / totalventa) * 100)
        For i = 1 To 6
        totales(i) = 0
        Next i
        FILTRO = resultados("vendedor")
           End If
           If VISTA1.Value = True Then
           
            linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            impresion.Cell(linea, 1).text = resultados(0)
            impresion.Cell(linea, 2).text = resultados(1)
            impresion.Cell(linea, 3).text = resultados(2)
            impresion.Cell(linea, 4).text = resultados(3)
            impresion.Cell(linea, 5).text = resultados(4)
            impresion.Cell(linea, 6).text = resultados(5)
            impresion.Cell(linea, 7).text = resultados(6)
            impresion.Cell(linea, 8).text = resultados(7)
            
            
            impresion.Cell(linea, 9).text = resultados(8)
            impresion.Cell(linea, 10).text = 1
            impresion.Cell(linea, 11).text = resultados(9)
            impresion.Cell(linea, 12).text = ((resultados(5) / totalventa) * 100)
           End If
           
            conta = 1
            totales(1) = totales(1) + CDbl(resultados(5))
            totales(2) = totales(2) + CDbl(resultados(6))
            totales(3) = totales(3) + CDbl(resultados(7))
            totales(4) = totales(4) + CDbl(resultados(8))
            totales(5) = totales(5) + CDbl(resultados(11))
            totales(6) = totales(6) + CDbl(resultados(9))
            
            totales2(1) = totales2(1) + CDbl(resultados(5))
            totales2(2) = totales2(2) + CDbl(resultados(6))
            totales2(3) = totales2(3) + CDbl(resultados(7))
            totales2(4) = totales2(4) + CDbl(resultados(8))
            totales2(5) = totales2(5) + CDbl(resultados(11))
            totales2(6) = totales2(6) + CDbl(resultados(9))
            
            
            resultados.MoveNext
        Wend
        cSql2.Close
        Set cSql2 = Nothing
        Set resultados = Nothing
        
    ' total vendedor
    
    linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           
           impresion.Cell(linea, 4).text = FILTRO
           impresion.Cell(linea, 5).text = leerNombreVendedor(FILTRO)
            
           impresion.Cell(linea, 6).text = totales(1)
           impresion.Cell(linea, 7).text = totales(2)
           impresion.Cell(linea, 8).text = totales(3)
           If totales(1) <> 0 Then
           
           impresion.Cell(linea, 9).text = (totales(4) / totales(1)) * 100
           End If
           impresion.Cell(linea, 10).text = totales(5)
           impresion.Cell(linea, 11).text = totales(6)
           impresion.Cell(linea, 12).text = ((totales(1) / totalventa) * 100)
            
        For i = 1 To 4
        totales(i) = 0
        Next i
        
    ' total todos
    
    
    linea = linea + 1
            impresion.Rows = impresion.Rows + 1
             impresion.Column(4).Locked = False
           impresion.Column(5).Locked = False
            
           
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 4, linea, 5).Merge
           
           
           
           impresion.Cell(linea, 4).text = "TOTAL GENERAL VENTAS"
            
            
            impresion.Cell(linea, 6).text = totales2(1)
            impresion.Cell(linea, 7).text = totales2(2)
            impresion.Cell(linea, 8).text = totales2(3)
            
            impresion.Cell(linea, 9).text = totales2(4) / totales2(1)
            impresion.Cell(linea, 10).text = totales2(5)
            impresion.Cell(linea, 11).text = totales2(6)
        
    End If
           impresion.Column(4).Locked = True
           impresion.Column(5).Locked = True
impresion.AutoRedraw = True
impresion.Refresh

        
        
    'Call sumaGrilla(impresion)
End Function



 

Private Sub TIPO1_Click()
FrameXp5.Visible = False
lblayuda.Visible = False
Command1_Click
End Sub

Private Sub TIPO2_Click()
FrameXp5.Visible = True
lblayuda.Visible = True

CODIGO.SetFocus
'Command1_Click
End Sub

Private Sub TIPO2_LostFocus()
lblayuda.Visible = False

End Sub

Private Sub VISTA1_Click()
Command1_Click

End Sub

Private Sub VISTA2_Click()
Command1_Click
End Sub
Private Function resumenVentasproductos(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim FILTRO As String
    
    Dim tabla As String
    Dim rubAux As String
    Dim harinas As Double
    Dim subproductos As Double
    Dim envases As Double
    Dim trigo As Double
    Dim maquila As Double
    Dim otros As Double
    Dim cadena As String
    Dim tipoDoc As String
    Dim numeroDoc As String
    Dim csql As rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim i As Integer
    Dim totales(10) As Double
    Dim totales2(10) As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim T3 As Double
    Dim PORCE As Double
    Dim cSql2 As New rdoQuery
    Dim resultados As rdoResultset
    
    Set cSql2.ActiveConnection = ventasAuditoria
    
    
    rubAux = rubro
 Rem IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(SUM(dd.total / " & Replace((1 + iva / 100), ",", ".") & "),0)
   ' ROUND(SUM(dd.total - dd.total * dd.descuento / 100),0)
    tabla = "SELECT dv.codigo, dv.descripcion, sum(dv.cantidad) as cantidad, IF (dv.tipo = 'FV' ,ROUND(SUM((dv.total - (dv.total * dv.descuento2 / 100)) ),0),ROUND(SUM(dv.total - (dv.total * dv.descuento2 / 100)/1.19),0)), dv.vendedor, "
    tabla = tabla + "IF (dv.tipo = 'FV' ,ROUND(SUM(((dv.total * dv.descuento2 / 100)) ),0),ROUND(SUM((dv.total * dv.descuento2 / 100)/1.19),0)) "
    tabla = tabla & "FROM sv_documento_detalle_" & codigoempresa & " AS dv "
    If codigoempresa = "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO'  "
    End If
    If codigoempresa <> "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO' AND dv.local='" + codigoempresa + "' "
    End If
    If Mid(codigovendedor, 1, 2) <> "99" Then
    tabla = tabla & "and dv.vendedor='" + Mid(codigovendedor, 1, 2) + "' and dv.caja < '90' "
    End If
    If CODIGO.text <> "" Then
    tabla = tabla & "and codigo='" & CODIGO.text & "'  "
    End If
    tabla = tabla & "GROUP BY dv.vendedor,dv.codigo "
    tabla = tabla & "ORDER BY dv.vendedor,total desc "
    cSql2.sql = tabla
    cSql2.Execute
    
   ' Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    
    If cSql2.RowsAffected > 0 Then
    Set resultados = cSql2.OpenResultset
    
       FILTRO = resultados("vendedor")
       impresion.Rows = 2
        While Not resultados.EOF
           If FILTRO <> resultados("vendedor") Then
           linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
         Rem   impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).text = FILTRO + " " + leerNombreVendedor(FILTRO)
            PORCE = 0
            If totales(1) <> 0 Then
            PORCE = totales(2) / (totales(3) + totales(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).text = totales(1)
            impresion.Cell(linea, 4).text = totales(2)
            impresion.Cell(linea, 5).text = PORCE
            impresion.Cell(linea, 6).text = totales(3)
        
            For i = 1 To 4
            totales(i) = 0
            Next i
        FILTRO = resultados("vendedor")
           End If
           If VISTA1.Value = True Then
           
           linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            impresion.Cell(linea, 1).text = resultados(0)
            impresion.Cell(linea, 2).text = resultados(1)
            t1 = resultados(2)
            t2 = resultados(3)
            
            If t1 = 0 Then t1 = 1
            PORCE = 0
            If resultados(5) <> 0 Then
            PORCE = resultados(5) / (t2 + resultados(5)) * 100
            
            End If
            
            impresion.Cell(linea, 3).text = resultados(2)
            impresion.Cell(linea, 4).text = resultados(5)
            impresion.Cell(linea, 5).text = PORCE
            impresion.Cell(linea, 6).text = resultados(3)
            impresion.Cell(linea, 7).text = resultados(3) / t1
            
           End If
            
            totales(1) = totales(1) + CDbl(resultados(2))
            totales(2) = totales(2) + CDbl(resultados(5))
            totales(3) = totales(3) + CDbl(resultados(3))
            
            
            totales2(1) = totales2(1) + CDbl(resultados(2))
            totales2(2) = totales2(2) + CDbl(resultados(5))
            totales2(3) = totales2(3) + CDbl(resultados(3))
            
            resultados.MoveNext
        Wend
        cSql2.Close
        Set cSql2 = Nothing
        Set resultados = Nothing
        
    ' total vendedor
    
    linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
           Rem impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).text = leerNombreVendedor(FILTRO)
           If totales(1) <> 0 Then
            PORCE = totales(2) / (totales(3) + totales(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).text = totales(1)
            impresion.Cell(linea, 4).text = totales(2)
            impresion.Cell(linea, 5).text = PORCE
            impresion.Cell(linea, 6).text = totales(3)
        
        For i = 1 To 4
        totales(i) = 0
        Next i
        
    ' total todos
    
    
    linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            
            
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
          Rem  impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).text = "TOTAL GENERAL VENTAS"
            
           If totales2(1) <> 0 Then
            PORCE = totales2(2) / (totales2(3) + totales2(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).text = totales2(1)
            impresion.Cell(linea, 4).text = totales2(2)
            impresion.Cell(linea, 5).text = PORCE
            impresion.Cell(linea, 6).text = totales2(3)
            
    End If

    'Call sumaGrilla(impresion)
End Function

Public Sub generaInformevp(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Dim documento As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("LISTADO VENTAS X VENDEDORES DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
    
    Call resumenVentasproductos(data, impresion, TIPO, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Sub CargaGrillaInforme2(ByVal row As Integer, ByVal col As Integer)
        Dim formatogrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "DESCRIPCION"
        formatogrilla(1, 3) = "VENDIDAS"
        formatogrilla(1, 4) = "DESCUENTO"
        formatogrilla(1, 5) = "DCTO(%)"
        formatogrilla(1, 6) = "VENTA NETA "
        formatogrilla(1, 7) = "P.PROMEDIO"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "12"
        formatogrilla(2, 2) = "45"
        formatogrilla(2, 3) = "9"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "9"
        formatogrilla(2, 6) = "9"
        formatogrilla(2, 7) = "9"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "C"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = "$ ###,###,##0"
        formatogrilla(4, 5) = "% #0.00"
        formatogrilla(4, 6) = "$ ###,###,##0"
        formatogrilla(4, 7) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "30"
        formatogrilla(8, 3) = "8"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "8"
        formatogrilla(8, 6) = "8"
        formatogrilla(8, 7) = "8"
        
                
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
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
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1.75
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub

  Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    impresion.FixedRowColStyle = Fixed3D
    impresion.CellBorderColorFixed = vbButtonShadow
    impresion.ShowResizeTips = False
    impresion.ReportTitles.Clear
    impresion.PageSetup.CenterHorizontally = True
    impresion.PageSetup.Orientation = cellLandscape
    
      
    impresion.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    impresion.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    impresion.PageSetup.HeaderAlignment = cellLeft
    impresion.PageSetup.HeaderFont.Name = "Verdana"
    impresion.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "ENTRE EL DIA  :  " & dato1.text + "-" + dato2.text + "-" + dato3.text & " y " & dato4.text + "-" + dato5.text + "-" + dato6.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle
    
    If TIPO2.Value = True Then
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "VENDEDOR: " & COMBOVENDEDOR.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle

    End If

    
    'PIE DE PAGINA
    impresion.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & usuarioSistema
    impresion.PageSetup.FooterAlignment = cellRight
    impresion.PageSetup.FooterFont.Name = "Verdana"
    impresion.PageSetup.FooterFont.Size = 7
    End Sub
    


Public Function LEErnotasdecredito(DESDE, HASTA, empresa) As Double
        

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventasRubro
        
        csql.sql = "SELECT tipodocumento,numerodocumento,codigo,tipo,numero,caja,fecha "
        csql.sql = csql.sql + "FROM " + clientesistema + "ventas" + empresa + ".sv_documento_detalle_" + empresa + " "
        csql.sql = csql.sql + "where fecha between '" + DESDE + "' and '" + HASTA + "' and (tipo='NB' or tipo='NF') and vendedor='' "
        csql.Execute
            
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
            Call modificadocumentos(resultados(0), resultados(1), resultados(2), resultados(3), resultados(4), resultados(5), resultados(6), empresa)
            resultados.MoveNext
            
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        
End Function
Sub modificadocumentos(TIPO, NUMERO, CODIGO, TIPO1, numero1, caja1, fecha1, empresa)
        

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventasRubro
        
        csql.sql = "select dd.vendedor from " + clientesistema + "ventas" + empresa + ".sv_documento_detalle_" + empresa + " as dd left join " + clientesistema + "ventas" + empresa + ".sv_documento_cabeza_" + empresa + " as dc on dd.tipo=dc.tipo and dd.numero=dc.numero and dd.fecha=dc.fecha and dd.caja=dc.caja "
        csql.sql = csql.sql + "where dc.tipo='" + TIPO + "' and dc.foliosii='" + NUMERO + "' and dd.codigo='" + CODIGO + "' "
        csql.Execute
            
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
            If resultados(0) <> "" Then
            Call modificadocumentos2(TIPO1, numero1, caja1, fecha1, resultados(0), empresa)
            End If
            resultados.MoveNext
            
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        
        
        
End Sub

Sub modificadocumentos2(TIPO1, numero1, caja1, fecha1, vendedor, empresa)
        

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventasRubro
        
        csql.sql = "update " + clientesistema + "ventas" + empresa + ".sv_documento_detalle_" + empresa + " set vendedor='" + vendedor + "' "
        csql.sql = csql.sql + "where tipo='" + TIPO1 + "' and numero='" + numero1 + "' and fecha='" + Format(fecha1, "yyyy-mm-dd") + "' and caja='" + caja1 + "'  "
        csql.Execute
        
        
        
End Sub


