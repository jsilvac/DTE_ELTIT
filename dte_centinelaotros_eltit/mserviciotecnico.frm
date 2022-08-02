VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form MServicio 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro Servicio Tecnico"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5670
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
      BackColor       =   16744576
      Caption         =   "Datos Tecnico(a)"
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
      Begin VB.TextBox dato8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox dato5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         MaxLength       =   20
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Celular"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comuna"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fono"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dirección"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
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
         Left            =   2940
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   5775
      _cx             =   10186
      _cy             =   2355
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
Attribute VB_Name = "MServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private c As tipoServicio
    Private modifica As Boolean

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Tecnico"
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
        Call VerificarCajas(Me, DATO4)
        Call selecciona(DATO4)
    End Sub

    Private Sub dato5_GotFocus()
        Call VerificarCajas(Me, DATO5)
        Call selecciona(DATO5)
    End Sub

    Private Sub dato6_GotFocus()
        Call VerificarCajas(Me, DATO6)
        Call selecciona(DATO6)
    End Sub
    
    Private Sub dato7_GotFocus()
        Call VerificarCajas(Me, DATO7)
        Call selecciona(DATO7)
    End Sub

    Private Sub dato8_GotFocus()
        Call VerificarCajas(Me, DATO8)
        Call selecciona(DATO8)
    End Sub
   
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaServicio(dato1)
        Else
            Call Flechas(KeyCode, dato1)
        End If
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
        Call Flechas(KeyCode, DATO4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, DATO5)
    End Sub
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, DATO6)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, DATO7)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And Val(dato1.text) <> 0 Then
            'Call Pregunta(dato1, dato1)
            dato1.text = ceros(dato1)
            lblDV.Caption = rut(dato1.text)
            If leerTecnico(c, dato1.text & lblDV.Caption, "=") = True Then
                Call structtoctrl
            Else
                If Verifica_Permiso(Me.Caption, "agrega") = True Then
                Call HabilitarCajas(Me, modifica)
                 SendKeys "{Tab}"
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                dato1.SelStart = 0
                dato1.SelLength = Len(dato1.text)
                dato1.SetFocus
            End If

            End If
            
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And LTrim(dato2.text) <> "" Then
            'Call Pregunta(dato2, dato3)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And LTrim(dato3.text) <> "" Then
            'Call Pregunta(dato3, dato4)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And LTrim(DATO4.text) <> "" Then
            'Call Pregunta(dato4, dato5)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And LTrim(DATO5.text) <> "" Then
            'Call Pregunta(dato5, dato6)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And DATO6.text <> "" Then
            'Call Pregunta(dato6, dato7)
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And DATO7.text <> "" Then
            'Call Pregunta(dato7, dato8)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And LTrim(DATO8.text) <> "" Then
            Call ctrltostruct
        End If
    End Sub
    
    '========================================================
    'KeyPress
    '========================================================

    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        Call limpiaBarra(2)
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
        modifica = False
        Call Centrar(Me)
    End Sub
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub
    
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        c.rut = dato1.text & lblDV.Caption
        c.nombre = dato2.text
        c.direccion = dato3.text
        c.comuna = DATO4.text
        c.ciudad = DATO5.text
        c.fono = DATO6.text
        c.celular = DATO7.text
        c.CODIGO = DATO8.text
        Call grabarTecnico(c, modifica)
        Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        dato1.text = c.rut
        lblDV.Caption = rut(dato1.text)
        dato2.text = c.nombre
        dato3.text = c.direccion
        DATO4.text = c.comuna
        DATO5.text = c.ciudad
        DATO6.text = c.fono
        DATO7.text = c.celular
        DATO8.text = c.CODIGO
        Call DeshabilitarCajas(Me)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    
'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
               If Verifica_Permiso(Me.Caption, "modifica") = True Then
                    Call modificar
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                End If
            Case "elimina"
                If Verifica_Permiso(Me.Caption, "elimina") = True Then
                    If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                    Call ELIMINAR
                    End If
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
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
        dato1.Enabled = False
        dato2.SetFocus
    End Sub
    
    Private Sub ELIMINAR()
        frmglosaeliminacion.Show vbModal
        Call eliminarTecnico(c)
        Call retorno
        Call HabilitarCajas(Me, modifica)
        dato1.Enabled = True
        dato1.SetFocus
    End Sub

    Private Sub retorno()
        Call LimpiarCajas(Me)
        Rem Call LimpiarLabels(Me)
        modifica = False
        Call DeshabilitarCajas(Me)
        dato1.SetFocus
        lblDV.Caption = ""
        
    End Sub
        
    Private Sub anterior()
        If leerTecnico(c, dato1.text & lblDV.Caption, "<") = True Then
            structtoctrl
        End If
    End Sub
    
    Private Sub siguiente()
        If leerTecnico(c, dato1.text & lblDV.Caption, ">") = True Then
            structtoctrl
        End If
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

