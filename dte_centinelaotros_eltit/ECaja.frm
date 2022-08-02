VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ECaja 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Egresos de Caja"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   6750
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   5400
      Top             =   6720
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
   Begin XPFrame.FrameXp frm1 
      Height          =   2235
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3942
      BackColor       =   16744576
      Caption         =   "Datos Egreso"
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
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   1575
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
         Height          =   285
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   780
         Width           =   495
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
         Height          =   285
         Left            =   5100
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   780
         Width           =   495
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
         Height          =   285
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   780
         Width           =   735
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
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   780
         Width           =   1575
      End
      Begin VB.TextBox dato8 
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
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   1860
         Width           =   4935
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
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   1140
         Width           =   375
      End
      Begin VB.TextBox dato7 
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
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Folio Interno"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   17
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Glosa"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Número"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label lblTipo 
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
         Left            =   1860
         TabIndex        =   12
         Top             =   1140
         Width           =   4515
      End
   End
   Begin XPFrame.FrameXp frm2 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2460
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
      BackColor       =   16744576
      Caption         =   "Recibido Por..."
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
      Begin VB.TextBox dato9 
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
         Height          =   285
         Left            =   120
         MaxLength       =   9
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblNombre 
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
         Left            =   2100
         TabIndex        =   19
         Top             =   360
         Width           =   4275
      End
      Begin VB.Label lblDv 
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
         Left            =   1620
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox manual 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   0
      Width           =   555
   End
   Begin XPFrame.FrameXp frm3 
      Height          =   3795
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6694
      BackColor       =   12648384
      Caption         =   "Listado de Egresos del Día"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
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
      Begin FlexCell.Grid Lista 
         Height          =   3255
         Left            =   120
         TabIndex        =   22
         Top             =   420
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5741
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   7140
      Width           =   6015
      _cx             =   10610
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "ECaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private e As egreso
    Private modifica As Boolean
    Private fecha As String
    'Private segurity As Boolean

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
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Tipo Egreso"
    End Sub
    
    Private Sub dato7_GotFocus()
        Call VerificarCajas(Me, DATO7)
        Call selecciona(DATO7)
    End Sub
    
    Private Sub dato8_GotFocus()
        Call VerificarCajas(Me, dato8)
        Call selecciona(dato8)
    End Sub
    
    Private Sub dato9_GotFocus()
        Call VerificarCajas(Me, dato9)
        Call selecciona(dato9)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Persona"
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
        If KeyCode = vbKeyF2 Then
            Call ayudaTipoEgreso(dato6)
        Else
            Call Flechas(KeyCode, dato5)
        End If
    End Sub
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, DATO7)
    End Sub
    
    Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaVendedores(dato9)
        Else
            Call Flechas(KeyCode, dato8)
        End If
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
            If Val(dato1.text) = 0 Then
                dato1.text = leer_Ultimo_Folio("numero", "sv_egresoscaja_" + rubro, dato1.MaxLength, ventasRubro, "local = '" & empresaActiva & "'")
            Else
                If leerEgreso(e, dato1.text, "=", fecha) = True Then
                    Call structtoctrl
                    opciones.SetFocus
                Else
                    Call HabilitarCajas(Me, modifica)
                    SendKeys "{Tab}"
                End If
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And Val(dato2.text) <> 0 Then
            dato2.text = ceros(dato2)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "00" Then
                dato3.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "0000" Then
                dato5.text = Format(fechasistema, "yyyy")
            End If
            fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
            Call leerEgresos
            'If leerEgreso(e, dato2.text, "=", fecha) = True Then
            '    Call structtoctrl
            'Else
            '    Call HabilitarCajas(Me, modifica)
                SendKeys "{Tab}"
            'End If
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            lblTipo.Caption = leerTipoEgreso(dato6.text)
            If lblTipo.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato9.text = ceros(dato9)
            lblDV.Caption = rut(dato9.text)
            'lblNombre.Caption = leerNombreVendedor(dato9.text)
            'If lblNombre.Caption <> "" Then
                Call ctrltostruct
                Call leerEgresos
            'End If
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
        Call seleccionaUno(KeyCode, dato1)
    End Sub
    
    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
        Call seleccionaUno(KeyCode, dato2)
    End Sub
    
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
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
     
    Private Sub dato3_LostFocus()
    Call esfecha(dato3, dato4, dato5, "dd")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato3, dato4, dato5, "mm")
    End Sub
    Private Sub dato5_LostFocus()
        fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        Call esfecha(dato3, dato4, dato5, "yyyy")
    End Sub
    Private Sub dato6_LostFocus()
        Call limpiaBarra(2)
    End Sub
    Private Sub dato9_LostFocus()
        Call limpiaBarra(2)
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_Activate()
        If segurity = True Then
            Seguridad.Show vbModal
            segurity = False
        End If
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub
    
    Private Sub Form_Load()
        titCaption = Me.Caption
        'segurity = Not Verificar(usuarioSistema, passwordSistema)
        modifica = False
        Call Centrar(Me)
        fecha = Format(fechasistema, "yyyy-mm-dd")
        dato3.text = Format(fecha, "dd")
        dato4.text = Format(fecha, "mm")
        dato5.text = Format(fecha, "yyyy")
        dato1.text = leer_Ultimo_Folio("id", "sv_egresoscaja_" + rubro, dato1.MaxLength, ventasRubro, "local = '" & empresaActiva & "'")
        Call CARGAGRILLA(1, 4)
        Call leerEgresos
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        e.loc = empresaActiva
        e.folio = dato1.text
        e.NUMERO = dato2.text
        e.fecha = dato5.text & dato4.text & dato3.text
        e.TIPO = dato6.text
        e.MONTO = DATO7.text
        e.glosa = dato8.text
        e.Recibido = dato9.text & lblDV.Caption
        Call grabarEgreso(e, modifica)
        Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        Dim cad As String
        dato1.text = e.folio
        dato2.text = e.NUMERO
        dato3.text = Format(e.fecha, "dd")
        dato4.text = Format(e.fecha, "mm")
        dato5.text = Format(e.fecha, "yyyy")
        dato6.text = e.TIPO
        DATO7.text = e.MONTO
        dato8.text = e.glosa
        dato9.text = e.Recibido
        lblDV.Caption = rut(dato9.text)
        LBLNOMBRE.Caption = leerNombreVendedor(e.Recibido)
        Call dato6_KeyPress(13)
        Call DeshabilitarCajas(Me)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

Private Sub manual_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27, Asc("r"), Asc("R")
            Call retorno
        Case Asc("a"), Asc("A"), 37
            Call anterior
        Case Asc("s"), Asc("S"), 39
            Call siguiente
        Case Asc("m"), Asc("M")
            Call modificar
        Case Asc("e"), Asc("E"), 46
            Call eliminar
    End Select
End Sub

'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
                Call modificar
            Case "elimina"
                If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                Call eliminar
                End If
            Case "imprime"
            Case "movimientos"
            Case "historico"
            Case "retorno"
                Call retorno
            Case "anterior"
                Call anterior
            Case "siguiente"
                Call siguiente
        End Select
    End Sub
    
    Private Sub modificar()
        modifica = True
        Call HabilitarCajas(Me, modifica)
        dato2.Enabled = False
        dato3.SetFocus
    End Sub
    
    Private Sub eliminar()
        Call eliminarEgreso(e)
        Call retorno
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
    End Sub

    Private Sub retorno()
        Call LimpiarCajas(ECaja)
        Call LimpiarLabels(ECaja)
        modifica = False
        fecha = Format(fechasistema, "yyyy-mm-dd")
        dato3.text = Format(fecha, "dd")
        dato4.text = Format(fecha, "mm")
        dato5.text = Format(fecha, "yyyy")
        Call DeshabilitarCajas(Me)
        dato1.text = leer_Ultimo_Folio("id", "sv_egresoscaja_" + rubro, dato1.MaxLength, ventasRubro, "local = '" & empresaActiva & "'")
        Call leerEgresos
        dato1.SetFocus
    End Sub
        
    Private Sub anterior()
        If leerEgreso(e, dato2.text, "<", fecha) = True Then
            structtoctrl
        End If
    End Sub
    
    Private Sub siguiente()
        If leerEgreso(e, dato2.text, ">", fecha) = True Then
            structtoctrl
        End If
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

Private Sub opciones_GotFocus()
    manual.SetFocus
End Sub

    Private Sub leerEgresos()
        Dim tabla As String
        tabla = "SELECT CONCAT(ec.numero, '" & vbTab & "', ec.tipo, ' ', mec.nombre, '" & vbTab & "', ec.monto) AS item, ec.glosa "
        tabla = tabla & "FROM sv_egresoscaja_" + rubro + " AS ec INNER JOIN " & baseVentas & ".sv_maestrotipoegresoscaja AS mec ON ec.tipo = mec.codigo "
        tabla = tabla & "WHERE ec.local = '" & empresaActiva & "' AND ec.fecha = '" & fecha & "' "
        tabla = tabla & "ORDER BY ec.monto ASC "
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        lista.Rows = 1
        lista.AutoRedraw = False
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                lista.AddItem data.Recordset.Fields("item"), True
                lista.AddItem "GLOSA" & vbTab & data.Recordset.Fields("glosa"), True
                lista.Cell(lista.Rows - 1, 1).Font.Bold = True
                lista.Range(lista.Rows - 1, 2, lista.Rows - 1, lista.Cols - 1).Merge
                lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellLeftCenter
                'Lista.Range(Lista.Rows - 1, 1, Lista.Rows - 1, Lista.Cols - 1).Locked = True
                data.Recordset.MoveNext
            Wend
        End If
        lista.AutoRedraw = True
        lista.Refresh
    End Sub

'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CARGAGRILLA(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Dim formatoGrilla(10, 5) As String
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "NUMERO"
        formatoGrilla(1, 2) = "TIPO EGRESO"
        formatoGrilla(1, 3) = "MONTO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "13"
        formatoGrilla(2, 2) = "50"
        formatoGrilla(2, 3) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = "0000000000"
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "20"
        formatoGrilla(8, 3) = "10"
            
        lista.Cols = col
        lista.Rows = row
        lista.AllowUserResizing = False
        lista.DisplayFocusRect = False
        lista.ExtendLastCol = True
        lista.BoldFixedCell = False
        lista.DrawMode = cellOwnerDraw
        lista.Appearance = Flat
        lista.ScrollBarStyle = Flat
        lista.FixedRowColStyle = Flat
        lista.BackColorFixed = RGB(90, 214, 158)
        lista.BackColorFixedSel = RGB(110, 230, 180)
        lista.BackColorBkg = RGB(90, 214, 158)
        lista.BackColorScrollBar = RGB(231, 247, 235)
        lista.BackColor1 = RGB(231, 247, 235)
        lista.BackColor2 = RGB(239, 255, 243)
        lista.GridColor = RGB(148, 231, 190)
        
        lista.Column(0).Width = 0
        For i = 1 To col - 1
            lista.Cell(0, i).text = formatoGrilla(1, i)
            lista.Column(i).Width = Val(formatoGrilla(8, i)) * (lista.Cell(0, i).Font.Size + 1.25)
            lista.Column(i).MaxLength = Val(formatoGrilla(2, i))
            lista.Column(i).FormatString = formatoGrilla(4, i)
            lista.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                lista.Column(i).Alignment = cellRightCenter
            Else
                lista.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        lista.Range(0, 1, 0, lista.Cols - 1).Alignment = cellCenterCenter
        lista.Range(0, 1, 0, lista.Cols - 1).FontBold = True
        lista.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

