VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LibroVentasclientes 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadistica de Ventas por Clientes"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   14565
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
      LockType        =   -1
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
      Height          =   6990
      Left            =   90
      TabIndex        =   1
      Top             =   2205
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   12330
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
         Height          =   6555
         Left            =   45
         TabIndex        =   2
         Top             =   360
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   11562
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   1
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2100
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   3704
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
      Begin XPFrame.FrameXp fechas 
         Height          =   915
         Left            =   4770
         TabIndex        =   5
         Top             =   360
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1614
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
         Left            =   8550
         TabIndex        =   4
         Top             =   360
         Width           =   1950
         _ExtentX        =   3440
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
            Width           =   1365
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
         Left            =   10530
         TabIndex        =   20
         Top             =   360
         Width           =   1680
         _ExtentX        =   2963
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
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Generar Informe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   12690
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   675
         Width           =   1545
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   660
         Left            =   0
         TabIndex        =   24
         Top             =   1260
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   1164
         BackColor       =   16744576
         Caption         =   "CLIENTES X RUT"
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
         Begin VB.TextBox rut1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   90
            MaxLength       =   9
            TabIndex        =   25
            Tag             =   "proveedor"
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label lblnombre 
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
            Left            =   2025
            TabIndex        =   27
            Top             =   270
            Width           =   6300
         End
         Begin VB.Label lblDV 
            Alignment       =   1  'Right Justify
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
            Left            =   1530
            TabIndex        =   26
            Top             =   270
            Width           =   375
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   660
         Left            =   45
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   1164
         BackColor       =   16744576
         Caption         =   "CLIENTES"
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
         Begin VB.ComboBox COMBOCLIENTES 
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   270
            Width           =   4485
         End
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
   Begin VB.TextBox sucu 
      Height          =   285
      Left            =   360
      MaxLength       =   1
      TabIndex        =   28
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "LibroVentasclientes"
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
    Private codigoCLIENTE As String
    Private rut_cliente As String
    




Private Sub Command1_Click()
codigoempresa = Mid(combolocal.text, 1, 2)
'codigoCLIENTE = Mid(COMBOCLIENTES.text, 1, 10)
codigoCLIENTE = rut_cliente

            If TIPO1.Value = True Then
            Call CargaGrillaInforme(1, 10)
            
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
        rut1.SetFocus
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
        dato1.text = "01"
        dato2.text = Format(fechasistema, "mm")
        dato3.text = Format(fechasistema, "yyyy")
        dato4.text = Format(fechasistema, "dd")
        dato5.text = Format(fechasistema, "mm")
        dato6.text = Format(fechasistema, "yyyy")
    LEErlocales
'    LEErclientes
    TIPO1.Value = True
    VISTA1.Value = True
    Call CargaGrillaInforme(1, 10)
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
        formatogrilla(1, 6) = "NETO"
        formatogrilla(1, 7) = "I.V.A"
        formatogrilla(1, 8) = "EXENTO"
        formatogrilla(1, 9) = "TOTAL"
        
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
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = "$ ###,###,##0"
        formatogrilla(4, 7) = "$ ###,###,##0"
        formatogrilla(4, 8) = "$ ###,###,##0"
        formatogrilla(4, 9) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "FALSE"
        formatogrilla(5, 8) = "FALSE"
        formatogrilla(5, 9) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        
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
        formatogrilla(8, 5) = "30"
        formatogrilla(8, 6) = "8"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
        
                
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
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        
        impresion.AutoRedraw = False
        If TIPO2.Value = 0 Then
        impresion.Range(1, 1, 1, 9).Borders(cellEdgeTop) = cellThick
        Else
        impresion.Range(1, 1, 1, 7).Borders(cellEdgeTop) = cellThick
        End If
        
        impresion.PageSetup.HeaderMargin = 2
        impresion.PageSetup.TopMargin = 1
        impresion.PageSetup.LeftMargin = 0.5
        impresion.PageSetup.RightMargin = 0
        impresion.PageSetup.BottomMargin = 1
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.Orientation = cellLandscape
        impresion.PageSetup.PrintFixedRow = True
        Call verificaImpresora(5, impresion)
        impresion.AutoRedraw = True
    End Sub
    
    
Sub LEErclientes()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventas
        
        csql.sql = "SELECT rut,nombre "
        csql.sql = csql.sql + "FROM sv_maestroclientes "
        csql.sql = csql.sql + "ORDER BY nombre "
        csql.Execute
        COMBOCLIENTES.AddItem ("9999999999" + "  TODOS LOS CLIENTES")
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                COMBOCLIENTES.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            
            COMBOCLIENTES.text = COMBOCLIENTES.List(0)
            
        End If
        
End Sub

Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas where rubro='" + rubro + "' "
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
      
                
        combolocal.text = combolocal.List(0)
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
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("LISTADO VENTAS X CLIENTES DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
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
    Dim csql As rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim i As Integer
    Dim totales(10) As Double
    Dim totales2(10) As Double

    rubAux = rubro
    tabla = "SELECT dc.tipo, dc.numero , dc.fecha, dc.rut, mc.nombre, dc.neto, dc.iva, dc.exento, dc.total, dc.rut "
                                                                        'ARIEL CAMBIA INNER POR LEFT
    tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc LEFT JOIN " & baseVentas & ".sv_maestroclientes AS mc ON dc.rut = mc.rut AND mc.sucursal = '0'"
    If codigoempresa = "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO'  "
    End If
    If codigoempresa <> "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO' AND dc.local='" + codigoempresa + "' "
    End If
    If codigoCLIENTE <> "9999999999" Then
    tabla = tabla & "and dc.rut='" + codigoCLIENTE + "' "
    End If
    
    tabla = tabla & "ORDER BY dc.vendedor,mc.nombre,dc.numero "
    Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    
    If data.Recordset.RecordCount > 0 Then
       FILTRO = data.Recordset.Fields("rut")
       impresion.Rows = 2
        While Not data.Recordset.EOF
           If FILTRO <> data.Recordset.Fields("rut") Then
           linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           impresion.Range(linea, 4, linea, 5).Merge
           impresion.Cell(linea, 4).text = leerNombreCliente(FILTRO)
           impresion.Cell(linea, 6).text = totales(1)
           impresion.Cell(linea, 7).text = totales(2)
           impresion.Cell(linea, 8).text = totales(3)
           impresion.Cell(linea, 9).text = totales(4)
        For i = 1 To 4
        totales(i) = 0
        Next i
        FILTRO = data.Recordset.Fields("rut")
           End If
           If VISTA1.Value = True Then
           
           linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            impresion.Cell(linea, 1).text = data.Recordset.Fields(0)
            impresion.Cell(linea, 2).text = data.Recordset.Fields(1)
            impresion.Cell(linea, 3).text = data.Recordset.Fields(2)
            impresion.Cell(linea, 4).text = data.Recordset.Fields(3)
            impresion.Cell(linea, 5).text = data.Recordset.Fields(4)
            impresion.Cell(linea, 6).text = data.Recordset.Fields(5)
            impresion.Cell(linea, 7).text = data.Recordset.Fields(6)
            impresion.Cell(linea, 8).text = data.Recordset.Fields(7)
            impresion.Cell(linea, 9).text = data.Recordset.Fields(8)
           End If
            
            totales(1) = totales(1) + CDbl(data.Recordset.Fields(5))
            totales(2) = totales(2) + CDbl(data.Recordset.Fields(6))
            totales(3) = totales(3) + CDbl(data.Recordset.Fields(7))
            totales(4) = totales(4) + CDbl(data.Recordset.Fields(8))
            totales2(1) = totales2(1) + CDbl(data.Recordset.Fields(5))
            totales2(2) = totales2(2) + CDbl(data.Recordset.Fields(6))
            totales2(3) = totales2(3) + CDbl(data.Recordset.Fields(7))
            totales2(4) = totales2(4) + CDbl(data.Recordset.Fields(8))
           
            data.Recordset.MoveNext
        Wend
    ' total vendedor
    
    linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 4, linea, 5).Merge
           impresion.Cell(linea, 4).text = leerNombreCliente(FILTRO)
           
            impresion.Cell(linea, 6).text = totales(1)
            impresion.Cell(linea, 7).text = totales(2)
            impresion.Cell(linea, 8).text = totales(3)
            impresion.Cell(linea, 9).text = totales(4)
        For i = 1 To 4
        totales(i) = 0
        Next i
        
    ' total todos
    
    
    linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            
            
           
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 4, linea, 5).Merge
           
           
           
           impresion.Cell(linea, 4).text = "TOTAL GENERAL VENTAS"
            
            
            impresion.Cell(linea, 6).text = totales2(1)
            impresion.Cell(linea, 7).text = totales2(2)
            impresion.Cell(linea, 8).text = totales2(3)
            impresion.Cell(linea, 9).text = totales2(4)
        
    End If

    'Call sumaGrilla(impresion)
End Function








Private Sub rut1_GotFocus()
        Call VerificarCajas(Me, rut1)
        Call selecciona(rut1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
End Sub

Private Sub rut1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
            Call ayudaCliente(rut1, SUCU, lblDV)
        Else
            Call Flechas(KeyCode, rut1)
        End If
End Sub

Private Sub rut1_KeyPress(KeyAscii As Integer)

 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And rut1.text <> "" And Val(rut1.text) <> 0 Then
            rut1.text = ceros(rut1)
            lblDV.Caption = rut(rut1.text)
            rut_cliente = rut1.text + lblDV.Caption
            lblNombre.Caption = LEERCLIENTE2(rut_cliente)
           Command1.SetFocus
           
        End If
End Sub

Private Sub rut_LostFocus()
   Call limpiaBarra(2)
End Sub

Private Sub TIPO1_Click()
Command1_Click
End Sub

Private Sub TIPO2_Click()
Command1_Click
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
    Dim porce As Double
    
    rubAux = rubro
Rem IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(SUM(dd.total / " & Replace((1 + iva / 100), ",", ".") & "),0)
   ' ROUND(SUM(dd.total - dd.total * dd.descuento / 100),0)
    tabla = "SELECT dv.codigo, mpf.descripcion, sum(dv.cantidad) as cantidad, IF (dv.tipo = 'FV' ,ROUND(SUM((dv.total - (dv.total * dv.descuento2 / 100)) ),0),ROUND(SUM(dv.total - (dv.total * dv.descuento2 / 100)/1.19),0)), dv.rut, "
    tabla = tabla + "IF (dv.tipo = 'FV' ,ROUND(SUM(((dv.total * dv.descuento2 / 100)) ),0),ROUND(SUM((dv.total * dv.descuento2 / 100)/1.19),0)) "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " AS dv INNER JOIN " & basedatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON mpf.codigobarra = dv.codigo "
    If codigoempresa = "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO'  "
    End If
    If codigoempresa <> "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO' AND dv.local='" + codigoempresa + "' "
    End If
    If codigoCLIENTE <> "9999999999" Then
    tabla = tabla & "and dv.rut='" + codigoCLIENTE + "' "
    End If
    
    tabla = tabla & "GROUP BY dv.codigo "
    tabla = tabla & "ORDER BY dv.rut,total desc "
    
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    
    If data.Recordset.RecordCount > 0 Then
       FILTRO = data.Recordset.Fields("rut")
       impresion.Rows = 2
        While Not data.Recordset.EOF
           If FILTRO <> data.Recordset.Fields("rut") Then
           linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).text = leerNombreCliente(FILTRO)
            porce = 0
            If totales(1) <> 0 Then
            porce = totales(2) / (totales(3) + totales(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).text = totales(1)
            impresion.Cell(linea, 4).text = totales(2)
            impresion.Cell(linea, 5).text = porce
            impresion.Cell(linea, 6).text = totales(3)
        
            For i = 1 To 4
            totales(i) = 0
            Next i
        FILTRO = data.Recordset.Fields("rut")
           End If
           If VISTA1.Value = True Then
           
           linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            impresion.Cell(linea, 1).text = data.Recordset.Fields(0)
            impresion.Cell(linea, 2).text = data.Recordset.Fields(1)
            t1 = data.Recordset.Fields(2)
            t2 = data.Recordset.Fields(3)
            
            If t1 = 0 Then t1 = 1
            porce = 0
            If data.Recordset.Fields(5) <> 0 Then
            porce = data.Recordset.Fields(5) / (t2 + data.Recordset.Fields(5)) * 100
            
            End If
            
            impresion.Cell(linea, 3).text = data.Recordset.Fields(2)
            impresion.Cell(linea, 4).text = data.Recordset.Fields(5)
            impresion.Cell(linea, 5).text = porce
            impresion.Cell(linea, 6).text = data.Recordset.Fields(3)
            impresion.Cell(linea, 7).text = data.Recordset.Fields(3) / t1
            
           End If
            
            totales(1) = totales(1) + CDbl(data.Recordset.Fields(2))
            totales(2) = totales(2) + CDbl(data.Recordset.Fields(5))
            totales(3) = totales(3) + CDbl(data.Recordset.Fields(3))
            
            
            totales2(1) = totales2(1) + CDbl(data.Recordset.Fields(2))
            totales2(2) = totales2(2) + CDbl(data.Recordset.Fields(5))
            totales2(3) = totales2(3) + CDbl(data.Recordset.Fields(3))
            
            data.Recordset.MoveNext
        Wend
    ' total vendedor
    
    linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).text = leerNombreCliente(FILTRO)
           If totales(1) <> 0 Then
            porce = totales(2) / (totales(3) + totales(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).text = totales(1)
            impresion.Cell(linea, 4).text = totales(2)
            impresion.Cell(linea, 5).text = porce
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
           
           impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).text = "TOTAL GENERAL VENTAS"
            
           If totales2(1) <> 0 Then
            porce = totales2(2) / (totales2(3) + totales2(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).text = totales2(1)
            impresion.Cell(linea, 4).text = totales2(2)
            impresion.Cell(linea, 5).text = porce
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
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "FALSE"
        
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

Public Function LEERCLIENTE2(rut) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        
        csql.sql = "SELECT mc.nombre "
        csql.sql = csql.sql & "FROM sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE mc.rut='" + rut + "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
        
          LEERCLIENTE2 = resultado(0)
      
        
     
        
            resultado.MoveNext
          Wend
        
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function
Sub CARGARDESDEAFUERA()
rut1_KeyPress (13)
Command1_Click
End Sub

