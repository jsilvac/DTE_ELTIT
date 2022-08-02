VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form electro02 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Ventas"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   15150
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Height          =   7260
      Left            =   120
      TabIndex        =   9
      Top             =   1500
      Width           =   15000
      _ExtentX        =   26458
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   11959
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1335
      Left            =   75
      TabIndex        =   6
      Top             =   90
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   2355
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "impresas"
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "x Imprimir"
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   960
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   0
         Top             =   0
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "Impresion Automatica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.OptionButton opt5 
         BackColor       =   &H00FF8080&
         Caption         =   "NC Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6960
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.OptionButton opt4 
         BackColor       =   &H00FF8080&
         Caption         =   "NC Boleta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4560
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Informe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   15
         Top             =   480
         Width           =   1635
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7440
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Boletas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7320
         TabIndex        =   13
         Top             =   1560
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton opt3 
         BackColor       =   &H00FF8080&
         Caption         =   "Zetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   9240
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   1245
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
         Left            =   6540
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   420
         Visible         =   0   'False
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
         Left            =   6060
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   420
         Visible         =   0   'False
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
         Left            =   7020
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   420
         Visible         =   0   'False
         Width           =   795
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
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   420
         Width           =   795
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
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   420
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
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin VB.Label LBLTIPO 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   21
         Top             =   480
         Width           =   2775
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
         Left            =   660
         TabIndex        =   8
         Top             =   420
         Width           =   1335
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
         Left            =   4620
         TabIndex        =   7
         Top             =   420
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   8415
      TabIndex        =   11
      Top             =   8865
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
Attribute VB_Name = "electro02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private contador As Double
    
    Private detalle2 As String
    Private tipo As String
    Private detalle As Boolean
    Private fecha1 As String
    Private fecha2 As String
    Public impresora As Printer
    
Private Sub Check1_Click()
Option1.Value = True

End Sub

Private Sub Command1_Click()
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, tipo, detalle, dato1.text, fecha1, fecha2)
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
            SendKeys "{Tab}"
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            SendKeys "{Tab}"
            Call generaInformeLV(data, impresion, tipo, detalle, dato1.text, fecha1, fecha2)
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
Call Command1_Click

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
        Call CargaGrillaInforme(1, 16)
        
        tipo = "(dc.tipo = 'FV')"
        detalle = False
       
        dato1.text = Format(fechasistema, "dd")
        dato2.text = Format(fechasistema, "mm")
        dato3.text = Format(fechasistema, "yyyy")
        dato4.text = Format(fechasistema, "dd")
        dato5.text = Format(fechasistema, "mm")
        dato6.text = Format(fechasistema, "yyyy")
    contador = 0
LBLTIPO.Caption = imprIMETIPO

    
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatogrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "TD"
        formatogrilla(1, 2) = "NUMERO"
        formatogrilla(1, 3) = "FECHA"
        formatogrilla(1, 4) = "RUT"
        formatogrilla(1, 5) = "CLIENTE"
        formatogrilla(1, 6) = "NETO"
        formatogrilla(1, 7) = "I.V.A"
        formatogrilla(1, 8) = "I.REF"
        formatogrilla(1, 9) = "I.VINOS"
        formatogrilla(1, 10) = "I.LIC"
        formatogrilla(1, 11) = "IHA "
        formatogrilla(1, 12) = "ICA "
        formatogrilla(1, 13) = "EXENTO"
        formatogrilla(1, 14) = "TOTAL"
        formatogrilla(1, 15) = "Nº SISTEMA"
        
        Rem LARGO DE LOS DATOS
        
        formatogrilla(2, 1) = "4"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "9"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "30"
        formatogrilla(2, 6) = "9"
        formatogrilla(2, 7) = "9"
        formatogrilla(2, 8) = "8"
        formatogrilla(2, 9) = "8"
        formatogrilla(2, 10) = "8"
        formatogrilla(2, 11) = "8"
        formatogrilla(2, 12) = "8"
        formatogrilla(2, 13) = "0"
        formatogrilla(2, 14) = "9"
        formatogrilla(2, 15) = "9"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
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
        formatogrilla(3, 13) = "N"
        formatogrilla(3, 14) = "N"
        formatogrilla(3, 15) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = "0000000000"
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = "###,###,##0"
        formatogrilla(4, 7) = "##,###,##0"
        formatogrilla(4, 8) = "##,###,##0"
        formatogrilla(4, 9) = "##,###,##0"
        formatogrilla(4, 10) = "##,###,##0"
        formatogrilla(4, 11) = "##,###,##0"
        formatogrilla(4, 12) = "##,###,##0"
        formatogrilla(4, 13) = "##,###,##0"
        formatogrilla(4, 14) = "###,###,##0"
        formatogrilla(4, 15) = "0000000000"
        
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
        formatogrilla(5, 13) = "TRUE"
        formatogrilla(5, 14) = "TRUE"
        formatogrilla(5, 15) = "TRUE"
        
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
        formatogrilla(6, 13) = ""
        formatogrilla(6, 14) = ""
        formatogrilla(6, 15) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        Rem ANCHO
        formatogrilla(8, 1) = "2"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "7"
        formatogrilla(8, 4) = "7"
        formatogrilla(8, 5) = "24"
       If opt2.Value = False Then
        formatogrilla(8, 6) = "7"
        formatogrilla(8, 7) = "5"
        formatogrilla(8, 8) = "5"
        formatogrilla(8, 9) = "5"
        formatogrilla(8, 10) = "5"
        formatogrilla(8, 11) = "5"
        formatogrilla(8, 12) = "5"
        Else
        formatogrilla(8, 6) = "0"
        formatogrilla(8, 7) = "0"
        formatogrilla(8, 8) = "0"
        formatogrilla(8, 9) = "0"
        formatogrilla(8, 10) = "0"
        formatogrilla(8, 11) = "0"
        formatogrilla(8, 12) = "0"
       End If
        formatogrilla(8, 13) = "0"
        
        formatogrilla(8, 14) = "7"
         formatogrilla(8, 15) = "7"
        
'        formatoGrilla(1, 7) = "I.V.A"
'        formatoGrilla(1, 8) = "I.REF"
'        formatoGrilla(1, 9) = "I.VINOS"
'        formatoGrilla(1, 10) = "I.LICORES"
'        formatoGrilla(1, 11) = "IHA "
'        formatoGrilla(1, 12) = "ICA "
        
                
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
        impresion.Column(0).Alignment = cellLeftGeneral
        
        
        impresion.Column(0).Width = 16
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
  impresion.Column(2).Mask = cellNumeric
  
        
        
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
        If impresion.Rows > 1 Then
        Call imprimir
        End If
        
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        impresion.AutoRedraw = False
        impresion.Range(1, 1, 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThick
       
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


Private Sub impresion_DblClick()
Dim xml As New ChilkatXml
Dim entradapdf As String
Dim salidapdf As String
Dim pdf As String
Dim tipodocu As String
If impresion.Cell(impresion.ActiveCell.row, 1).text = "FV" Then tipodocu = "33"
If impresion.Cell(impresion.ActiveCell.row, 1).text = "ND" Then tipodocu = "56"
If impresion.Cell(impresion.ActiveCell.row, 1).text = "NF" Then tipodocu = "61"


Call imprimelectronica(tipodocu, impresion.Cell(impresion.ActiveCell.row, 2).text)

Command1_Click

Rem MsgBox "presione enter para imprimir "
End Sub

    Private Sub opt1_Click()
        If opt1.Value = True Then
            tipo = "(dc.tipo = 'FV')"
           Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, tipo, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    
    Private Sub opt2_Click()
        If opt2.Value = True Then
            tipo = "(dc.tipo = 'BV')"
            Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, tipo, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    Private Sub opt3_Click()
        If opt3.Value = True Then
            tipo = "(dc.tipo = 'ZE')"
            Call generaInformeLV(data, impresion, tipo, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    
    
Private Sub opt4_Click()
      If opt4.Value = True Then
            tipo = "(dc.tipo = 'NB')"
             Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, tipo, detalle, dato1.text, fecha1, fecha2)
        End If
      End Sub

    Private Sub opt5_Click()
        If opt5.Value = True Then
            tipo = "(dc.tipo = 'NF')"
             Call CargaGrillaInforme(1, 16)
            Call generaInformeLV(data, impresion, tipo, detalle, dato1.text, fecha1, fecha2)
        End If

End Sub

Private Function listadte(ByRef data As Adodc, ByRef impresion As Grid, ByVal tipo As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
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
    Dim resultados As rdoResultset
        
    Dim i As Integer

    rubAux = rubro
    Rem Call Conectarventas(servidor, baseVentas + empresaActiva, usuario, password)
    
    
    Set csql.ActiveConnection = ventasRubro


    csql.sql = "SELECT dc.tipo,dt.numero,dc.fecha,dc.rut,IFNULL(mc.nombre,'') "
    csql.sql = csql.sql & "as nombre,dc.neto,dc.iva,dc.exento,dc.total, "
    csql.sql = csql.sql & "dc.impuestoilarefrescos,dc.impuestoilavinos, "
    csql.sql = csql.sql & "dc.impuestoilalicores,dc.impuestoharina, "
    csql.sql = csql.sql & "dc.impuestocarne,dc.foliosii,dc.caja "
    csql.sql = csql.sql & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc left JOIN " & baseVentas & ".sv_maestroclientes AS mc ON dc.rut = mc.rut AND mc.sucursal = '0'"
    csql.sql = csql.sql & "inner join " + clientesistema + "fae" + empresaActiva + ".sv_dte" + empresaActiva + " as dt on dt.tipodocumento = dc.TIPO and dt.numerodocumento=dc.numero and dt.cajadocumento=dc.caja and dt.fechadocumento=dc.fecha "
    csql.sql = csql.sql & "WHERE dc.fecha = '" & fecha1 & "' AND (dc.tipo='FV' OR dc.tipo='NF' OR dc.tipo='ND') "
    If Option1.Value = True Then
    csql.sql = csql.sql + "and dt.impresa='0' "
    Else
    csql.sql = csql.sql + "and dt.impresa='1' "
    
    End If
    If imprIMETIPO = "VENTAS" Then
    csql.sql = csql.sql + " and dc.caja<'90'"
    End If
    If imprIMETIPO = "CAJA99" Then
    csql.sql = csql.sql + " and dc.caja='99'"
    End If
    If imprIMETIPO = "CAJA98" Then
    csql.sql = csql.sql + " and dc.caja='98'"
    End If
    
    
    csql.sql = csql.sql & "ORDER BY dt.numero "
    'Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    csql.Execute
  
    linea = 0
    If csql.RowsAffected > 0 Then
       impresion.Rows = 1
       Set resultados = csql.OpenResultset
        While Not resultados.EOF
           impresion.Rows = impresion.Rows + 1
           linea = linea + 1
            impresion.Cell(linea, 0).text = resultados("caja")
            impresion.Cell(linea, 1).text = resultados("tipo")
            impresion.Cell(linea, 2).text = resultados("numero")
            impresion.Cell(linea, 3).text = resultados("fecha")
            impresion.Cell(linea, 4).text = resultados("rut")
            impresion.Cell(linea, 5).text = resultados("nombre")
            impresion.Cell(linea, 6).text = resultados("neto")
            impresion.Cell(linea, 7).text = resultados("iva")
            impresion.Cell(linea, 8).text = resultados("impuestoilarefrescos")
            impresion.Cell(linea, 9).text = resultados("impuestoilavinos")
            impresion.Cell(linea, 10).text = resultados("impuestoilalicores")
            impresion.Cell(linea, 11).text = resultados("impuestoharina")
            impresion.Cell(linea, 12).text = resultados("impuestocarne")
            impresion.Cell(linea, 13).text = resultados("exento")
            impresion.Cell(linea, 14).text = resultados("total")
            impresion.Cell(linea, 15).text = resultados("foliosii")
            
            resultados.MoveNext
        Wend
    
    End If
Set csql = Nothing
csql.Close
Set resultados = Nothing

    'Call sumaGrilla(impresion)
End Function

Public Sub generaInformeLV(ByRef data As Adodc, ByRef impresion As Grid, ByVal tipo As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Dim documento As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    If tipo = "FV" Then documento = "FACTURAS"
    If tipo = "BV" Then documento = "BOLETAS "
    If tipo = "ZE" Then documento = "ZETAS   "
    
    Call cargaCabeza("LISTADO DOCUMENTOS EMITIDOS  " + documento + " DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
    Call listadte(data, impresion, tipo, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Sub Option2_Click()
Check1.Value = 0
End Sub

Private Sub Timer1_Timer()

If Check1.Value = 1 Then
Command1_Click
If impresion.Rows > 1 Then
impresion.Cell(1, 1).SetFocus
Call impresion_DblClick
End If
End If

End Sub
