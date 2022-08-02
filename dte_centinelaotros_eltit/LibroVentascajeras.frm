VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LibroVentascajera 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadisticas de Ventas por Caja y Cajera"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   720
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
      Height          =   7260
      Left            =   90
      TabIndex        =   0
      Top             =   1980
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   12806
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
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   13260
         _ExtentX        =   23389
         _ExtentY        =   11959
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   1
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   5715
      TabIndex        =   2
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1830
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   3228
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
         Height          =   645
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Width           =   1545
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1440
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   2540
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
            TabIndex        =   10
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
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
            TabIndex        =   9
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
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
            TabIndex        =   7
            Tag             =   "proveedor"
            Top             =   540
            Width           =   705
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
            TabIndex        =   6
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
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
            TabIndex        =   5
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
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
            TabIndex        =   11
            Top             =   240
            Width           =   1605
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   660
         Left            =   0
         TabIndex        =   16
         Top             =   1080
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   1164
         BackColor       =   16744576
         Caption         =   "TODAS LAS CAJAS                       TODAS LAS CAJERAS"
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
         Begin VB.OptionButton cajeras 
            BackColor       =   &H00FF8080&
            Caption         =   "CAJERAS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2565
            TabIndex        =   24
            Top             =   225
            Width           =   1275
         End
         Begin VB.OptionButton cajas 
            BackColor       =   &H00FF8080&
            Caption         =   "CAJAS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   675
         Left            =   90
         TabIndex        =   17
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
            TabIndex        =   18
            Top             =   270
            Width           =   4485
         End
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1230
         Left            =   10920
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
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
         Begin VB.OptionButton VISTA2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Acumulada"
            Height          =   375
            Left            =   90
            TabIndex        =   21
            Top             =   675
            Width           =   2175
         End
         Begin VB.OptionButton VISTA1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Detallada"
            Height          =   330
            Left            =   90
            TabIndex        =   20
            Top             =   270
            Width           =   1950
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   1230
         Left            =   8760
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
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
         Begin VB.OptionButton TIPO1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "x Documentos"
            Height          =   330
            Left            =   135
            TabIndex        =   15
            Top             =   270
            Width           =   1365
         End
         Begin VB.OptionButton TIPO2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "x Articulos"
            Height          =   375
            Left            =   135
            TabIndex        =   14
            Top             =   630
            Width           =   1635
         End
      End
   End
End
Attribute VB_Name = "LibroVentascajera"
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

    
Private Sub COMBOVENDEDOR_Click()
Command1_Click



End Sub

Private Sub cajas_Click()
If combolocal.text <> "" Then
Call Command1_Click
End If


End Sub

Private Sub cajeras_Click()
If combolocal.text <> "" Then
Call Command1_Click
End If

End Sub

Private Sub Command1_Click()
codigoempresa = Mid(combolocal.text, 1, 2)



            If TIPO1.Value = True Then
            Call CargaGrillaInforme(1, 7)
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
'    LEErVENDEDORES
    Call CargaGrillaInforme(1, 6)
    TIPO1.Value = True
    VISTA1.Value = True
    
    'Call CargaGrillaInformeventasxvendedor(1, 7)
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        col = 6
        Rem DATOS DE LA COLUMNA
        If cajas.Value = True Then
        formatoGrilla(1, 1) = "CAJA"
        formatoGrilla(1, 2) = "NOMBRE"
        End If
        If cajeras.Value = True Then
        formatoGrilla(1, 1) = "CAJERA"
        formatoGrilla(1, 2) = "NOMBRE"
         End If
'        formatoGrilla(1, 3) = "NETO"
'        formatoGrilla(1, 4) = "I.V.A"
'        formatoGrilla(1, 5) = "EXENTO"
        formatoGrilla(1, 3) = "TOTAL"
        formatoGrilla(1, 4) = "DONACION"
        formatoGrilla(1, 5) = "LOCAL"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "8"
        formatoGrilla(2, 2) = "10"
        formatoGrilla(2, 3) = "8"
        formatoGrilla(2, 4) = "10"
        formatoGrilla(2, 5) = "25"
'        formatoGrilla(2, 6) = "9"
'        formatoGrilla(2, 7) = "9"
'        formatoGrilla(2, 8) = "6"
'        formatoGrilla(2, 9) = "9"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "D"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "S"
'        formatoGrilla(3, 6) = "N"
'        formatoGrilla(3, 7) = "N"
'        formatoGrilla(3, 8) = "N"
'        formatoGrilla(3, 9) = "N"
        
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = "$ ###,###,##0"
        formatoGrilla(4, 4) = "$ ###,###,##0"
'        formatoGrilla(4, 5) = "$ ###,###,##0"
'        formatoGrilla(4, 6) = "$ ###,###,##0"
'        formatoGrilla(4, 7) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
'        formatoGrilla(5, 6) = "FALSE"
'        formatoGrilla(5, 7) = "FALSE"
       
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
'        formatoGrilla(6, 5) = ""
'        formatoGrilla(6, 6) = ""
       
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
'        formatoGrilla(7, 5) = ""
'        formatoGrilla(7, 6) = ""
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
      
        formatoGrilla(8, 2) = "20"
     
      
      
        
        formatoGrilla(8, 3) = "8"
        formatoGrilla(8, 4) = "10"
        formatoGrilla(8, 5) = "20"

'        formatoGrilla(8, 5) = "10"
'        formatoGrilla(8, 6) = "8"
'        formatoGrilla(8, 6) = "7"
          
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
            impresion.Cell(0, i).text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.Column(5).Width = 200
        
        
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
'formato grilla ventas por vendedor
Private Sub CargaGrillaInformeventasxvendedor(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "VENDEDOR"
        formatoGrilla(1, 3) = "NETO"
        formatoGrilla(1, 4) = "I.V.A"
        formatoGrilla(1, 5) = "EXENTO"
        formatoGrilla(1, 6) = "TOTAL"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "5"
        formatoGrilla(2, 2) = "20"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "9"
        formatoGrilla(2, 5) = "9"
        formatoGrilla(2, 6) = "9"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = "$ ###,###,##0"
        formatoGrilla(4, 4) = "$ ###,###,##0"
        formatoGrilla(4, 5) = "$ ###,###,##0"
        formatoGrilla(4, 6) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        Rem ANCHO
        formatoGrilla(8, 1) = "5"
        formatoGrilla(8, 2) = "20"
        formatoGrilla(8, 3) = "8"
        formatoGrilla(8, 4) = "8"
        formatoGrilla(8, 5) = "8"
        formatoGrilla(8, 6) = "8"
        
                
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
            impresion.Cell(0, i).text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
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
        
        impresion.AutoRedraw = False
        impresion.Range(1, 1, 1, 5).Borders(cellEdgeTop) = cellThick
        impresion.PageSetup.HeaderMargin = 2
        impresion.PageSetup.TopMargin = 1
        impresion.PageSetup.LeftMargin = 0.5
        impresion.PageSetup.RightMargin = 0
        impresion.PageSetup.BottomMargin = 1
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.Orientation = cellPortrait
        impresion.Cols = 6
                
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeTop) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeRight) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellInsideHorizontal) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellInsideVertical) = cellThick

        impresion.PageSetup.PrintFixedRow = True
        Call verificaImpresora(5, impresion)
        impresion.AutoRedraw = True
    End Sub
    
    
Sub LEErVENDEDORES()
'    Dim resultados As rdoResultset
'    Dim cSql As New rdoQuery
'
'        Set cSql.ActiveConnection = ventas
'
'        cSql.sql = "SELECT rut,nombre "
'        cSql.sql = cSql.sql + "FROM sv_maestrovendedores "
'        cSql.sql = cSql.sql + "ORDER BY codigo "
'        cSql.Execute
'        COMBOVENDEDOR.AddItem ("99" + "  TODOS LOS VENDEDORES")
'        If cSql.RowsAffected > 0 Then
'            Set resultados = cSql.OpenResultset
'            While Not resultados.EOF
'                COMBOVENDEDOR.AddItem (resultados(0) + " " + resultados(1))
'                resultados.MoveNext
'            Wend
'            resultados.Close
'            Set resultados = Nothing
'
'            COMBOVENDEDOR.text = COMBOVENDEDOR.List(0)
'
'        End If
'
End Sub
Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = gestion
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM g_maestroempresas "
        ' original cSql.sql = cSql.sql + "ORDER BY codigo "
        ' ariel agrega condicion local < 50 para que no liste locales 50 y 51
        cSql.sql = cSql.sql + "  WHERE CODIGO < '50' ORDER BY codigo "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
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
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM g_maestroempresas "
        cSql.sql = cSql.sql + "ORDER BY codigo "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
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
    If cajas.Value = True Then
    Call cargaCabeza("LISTADO VENTAS X CAJA DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
   End If
   If cajeras.Value = True Then
   Call cargaCabeza("LISTADO VENTAS X CAJERAS DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
    End If
   
    Call resumenVentas(data, impresion, TIPO, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Function resumenVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim FILTRO As String
    Dim Q As Integer
    
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
    Dim cSql As rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim i As Integer
    Dim totales(10) As Double
    Dim totales2(10) As Double

    rubAux = rubro
    If codigoempresa <> "99" Then
    tabla = "SELECT  dc.caja,dc.cajera,dc.fecha,sum(dc.neto), sum(dc.iva), sum(dc.exento), sum(dc.total),SUM(DONACION) "
    tabla = tabla & "FROM sv_documento_cabeza_" + codigoempresa + " AS dc "
    If codigoempresa <> "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO' AND  tipo<>'NF' and tipo<>'NB' and caja <'90' and dc.local='" + codigoempresa + "' "
    End If
    If cajas.Value = True Then
    tabla = tabla & " group by dc.caja "
    End If
    
    If cajeras.Value = True Then
    tabla = tabla & " group by dc.cajera "
    End If
    
    Call ConectarControlData(data, servidor, baseVentas & codigoempresa, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    impresion.Cols = 8
    If data.Recordset.RecordCount > 0 Then
       
       impresion.Rows = 2
        While Not data.Recordset.EOF
           
           
           If VISTA1.Value = True Then
           
           linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            If cajas.Value = True Then
            impresion.Cell(linea, 1).text = data.Recordset.Fields(0)
            impresion.Cell(linea, 2).text = leerNombreCaja(impresion.Cell(linea, 1).text)
         
            End If
            If cajeras.Value = True Then
            impresion.Cell(linea, 1).text = data.Recordset.Fields(1) & "-" & rut(data.Recordset.Fields(1))
            impresion.Cell(linea, 2).text = leerNombreCajera(Replace(impresion.Cell(linea, 1).text, "-", ""))
             End If
            impresion.Cell(linea, 3).text = data.Recordset.Fields(6)
            impresion.Cell(linea, 4).text = data.Recordset.Fields(7)
'            impresion.Cell(linea, 5).text = data.Recordset.Fields(5)
'            impresion.Cell(linea, 6).text = data.Recordset.Fields(6)
'            impresion.Cell(linea, 7).text = data.Recordset.Fields(7)


           End If
            
            totales(1) = totales(1) + CDbl(data.Recordset.Fields(6))
            totales(2) = totales(2) + CDbl(data.Recordset.Fields(7))
'            totales(3) = totales(3) + CDbl(data.Recordset.Fields(7))
'            totales(4) = totales(4) + CDbl(data.Recordset.Fields(8))
'            totales2(1) = totales2(1) + CDbl(data.Recordset.Fields(5))
'            totales2(2) = totales2(2) + CDbl(data.Recordset.Fields(6))
'            totales2(3) = totales2(3) + CDbl(data.Recordset.Fields(7))
'            totales2(4) = totales2(4) + CDbl(data.Recordset.Fields(8))
           
            data.Recordset.MoveNext
        Wend
    ' total vendedor
    
    linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           
'           impresion.Range(linea, 4, linea, 5).Merge
'           impresion.Cell(linea, 4).text = leerNombreVendedor(FILTRO)
           
'            impresion.Cell(linea, 6).text = totales(1)
'            impresion.Cell(linea, 7).text = totales(2)
'            impresion.Cell(linea, 8).text = totales(3)
'            impresion.Cell(linea, 9).text = totales(4)
        totales2(1) = totales(1)
        totales2(2) = totales(2)
        For i = 1 To 4
        totales(i) = 0
        Next i
        
    ' total todos
    
    
    linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            
            
           
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
'
           impresion.Range(linea, 1, linea, 2).Merge
           
           
           
           impresion.Cell(linea, 1).text = "TOTAL GENERAL VENTAS"
            
            
            impresion.Cell(linea, 3).text = totales2(1)
            impresion.Cell(linea, 4).text = totales2(2)
'            impresion.Cell(linea, 8).text = totales2(3)
'            impresion.Cell(linea, 9).text = totales2(4)
'
    End If
Else
impresion.Rows = 2
linea = 0
For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    
For Q = 0 To (combolocal.ListCount - 2)
    
    tabla = "SELECT  dc.caja,dc.cajera,dc.fecha,sum(dc.neto), sum(dc.iva), sum(dc.exento), sum(dc.total),SUM(DONACION) "
    tabla = tabla & "FROM sv_documento_cabeza_" + Mid(combolocal.List(Q), 1, 2) + " AS dc "
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO' AND  tipo<>'NF' and tipo<>'NB' and caja <'90' "
    If cajas.Value = True Then
    tabla = tabla & " group by dc.caja "
    End If
    If cajeras.Value = True Then
    tabla = tabla & " group by dc.cajera "
    End If
    
        Call ConectarControlData(data, servidor, baseVentas & Mid(combolocal.List(Q), 1, 2), usuario, password, tabla)
   
    
    impresion.Cols = 8
    If data.Recordset.RecordCount > 0 Then
       
      
        While Not data.Recordset.EOF
           
           
           If VISTA1.Value = True Then
           
           linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            If cajas.Value = True Then
            impresion.Cell(linea, 1).text = data.Recordset.Fields(0)
            impresion.Cell(linea, 2).text = leerNombreCaja(impresion.Cell(linea, 1).text)
         
            End If
            If cajeras.Value = True Then
            impresion.Cell(linea, 1).text = data.Recordset.Fields(1) & "-" & rut(data.Recordset.Fields(1))
            impresion.Cell(linea, 2).text = leerNombreCajera(Replace(impresion.Cell(linea, 1).text, "-", ""))
             End If
            impresion.Cell(linea, 3).text = data.Recordset.Fields(6)
            impresion.Cell(linea, 4).text = data.Recordset.Fields(7)
            impresion.Cell(linea, 5).text = leerNombreEmpresa(Mid(combolocal.List(Q), 1, 2))
'            impresion.Cell(linea, 6).text = data.Recordset.Fields(6)
'            impresion.Cell(linea, 7).text = data.Recordset.Fields(7)


           End If
            
            totales(1) = totales(1) + CDbl(data.Recordset.Fields(6))
            totales(2) = totales(2) + CDbl(data.Recordset.Fields(7))
'            totales(3) = totales(3) + CDbl(data.Recordset.Fields(7))
'            totales(4) = totales(4) + CDbl(data.Recordset.Fields(8))
            totales2(1) = totales2(1) + CDbl(data.Recordset.Fields(6))
            totales2(2) = totales2(2) + CDbl(data.Recordset.Fields(7))
'            totales2(3) = totales2(3) + CDbl(data.Recordset.Fields(7))
'            totales2(4) = totales2(4) + CDbl(data.Recordset.Fields(8))
           
            data.Recordset.MoveNext
        Wend
    ' total vendedor
    
    linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
'           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           
'           impresion.Range(linea, 4, linea, 5).Merge
'           impresion.Cell(linea, 4).text = leerNombreVendedor(FILTRO)
           
'            impresion.Cell(linea, 6).text = totales(1)
'            impresion.Cell(linea, 7).text = totales(2)
'            impresion.Cell(linea, 8).text = totales(3)
'            impresion.Cell(linea, 9).text = totales(4)
     

            
           
           impresion.Range(linea, 1, linea, 5).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 5).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 5).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 5).Borders(cellEdgeRight) = cellThin

           impresion.Range(linea, 1, linea, 2).Merge
           
           
           
           impresion.Cell(linea, 1).text = "TOTAL GENERAL LOCAL"
            
            
            impresion.Cell(linea, 3).text = totales(1)
            impresion.Cell(linea, 4).text = totales(2)
linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            
   For i = 1 To 4
        totales(i) = 0
        Next i
        
'            impresion.Cell(linea, 8).text = totales2(3)
'            impresion.Cell(linea, 9).text = totales2(4)
'
    
    
    ' total todos
    
    
    
    End If

Next Q
linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            
            
           
           impresion.Range(linea, 1, linea, 5).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 5).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 5).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 5).Borders(cellEdgeRight) = cellThin
'
           impresion.Range(linea, 1, linea, 2).Merge
           
           
           
           impresion.Cell(linea, 1).text = "TOTAL GENERAL VENTAS"
            
            
            impresion.Cell(linea, 3).text = totales2(1)
            impresion.Cell(linea, 4).text = totales2(2)
'            impresion.Cell(linea, 8).text = totales2(3)
'            impresion.Cell(linea, 9).text = totales2(4)
'
End If

    'Call sumaGrilla(impresion)
End Function



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
    Dim cSql As rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim i As Integer
    Dim totales(10) As Double
    Dim totales2(10) As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim T3 As Double
    Dim porce As Double
    Dim forinicio As Integer
    Dim forfinal As Integer
    
    rubAux = rubro
Rem IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(SUM(dd.total / " & Replace((1 + iva / 100), ",", ".") & "),0)
   ' ROUND(SUM(dd.total - dd.total * dd.descuento / 100),0)
    codigoempresa = Mid(combolocal.text, 1, 2)
   
    tabla = "SELECT dv.codigo, mpf.descripcion, sum(dv.cantidad) as cantidad, IF (dv.tipo = 'FV' ,ROUND(SUM((dv.total - (dv.total * dv.descuento2 / 100)) ),0),ROUND(SUM(dv.total - (dv.total * dv.descuento2 / 100)/1.19),0)), dv.vendedor, "
    tabla = tabla + "IF (dv.tipo = 'FV' ,ROUND(SUM(((dv.total * dv.descuento2 / 100)) ),0),ROUND(SUM((dv.total * dv.descuento2 / 100)/1.19),0)) "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " AS dv INNER JOIN " & basedatos & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf ON mpf.codigobarra = dv.codigo "
    If codigoempresa = "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO'  "
    End If
    If codigoempresa <> "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO' AND dv.local='" + codigoempresa + "' "
    End If
    
    tabla = tabla & "GROUP BY dv.codigo "
    tabla = tabla & "ORDER BY dv.vendedor,total desc "
    
    Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    
    If data.Recordset.RecordCount > 0 Then
       FILTRO = data.Recordset.Fields("vendedor")
       impresion.Rows = 2
        While Not data.Recordset.EOF
           If FILTRO <> data.Recordset.Fields("vendedor") Then
           linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).text = leerNombreVendedor(FILTRO)
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
        FILTRO = data.Recordset.Fields("vendedor")
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
           impresion.Cell(linea, 2).text = leerNombreVendedor(FILTRO)
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
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "DESCRIPCION"
        formatoGrilla(1, 3) = "VENDIDAS"
        formatoGrilla(1, 4) = "DESCUENTO"
        formatoGrilla(1, 5) = "DCTO(%)"
        formatoGrilla(1, 6) = "VENTA NETA "
        formatoGrilla(1, 7) = "P.PROMEDIO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "12"
        formatoGrilla(2, 2) = "45"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "9"
        formatoGrilla(2, 5) = "9"
        formatoGrilla(2, 6) = "9"
        formatoGrilla(2, 7) = "9"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "C"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "N"
        formatoGrilla(3, 7) = "N"
        
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = "$ ###,###,##0"
        formatoGrilla(4, 5) = "% #0.00"
        formatoGrilla(4, 6) = "$ ###,###,##0"
        formatoGrilla(4, 7) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        formatoGrilla(5, 7) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        formatoGrilla(7, 7) = ""
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "30"
        formatoGrilla(8, 3) = "8"
        formatoGrilla(8, 4) = "8"
        formatoGrilla(8, 5) = "8"
        formatoGrilla(8, 6) = "8"
        formatoGrilla(8, 7) = "8"
        
                
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
            impresion.Cell(0, i).text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub

