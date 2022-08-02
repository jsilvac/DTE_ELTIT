VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FF9A514-A943-11D2-8D43-F90F0D71B6F6}#1.0#0"; "changeres.ocx"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Ventas"
   ClientHeight    =   8385
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Principal.frx":1ADB4
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ChangeResProject.ChangeRes ChangeRes1 
      Left            =   60
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1085
   End
   Begin MSComctlLib.StatusBar barraEstado 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8010
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   14446
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Mventas 
      Caption         =   "FACTURA ELECTRONICA"
      Index           =   99
      Begin VB.Menu elec 
         Caption         =   "Administrador de Folios (CAF)"
         Index           =   1
      End
      Begin VB.Menu elec 
         Caption         =   "Impresion de Facturas Electronicas"
         Index           =   2
      End
      Begin VB.Menu elec 
         Caption         =   "Motor Genera Facturas Electronicas"
         Index           =   3
      End
      Begin VB.Menu elec 
         Caption         =   "Envio D.T.E SII"
         Index           =   4
      End
      Begin VB.Menu elec 
         Caption         =   "Envio Libro de Compras  SII"
         Index           =   5
      End
      Begin VB.Menu elec 
         Caption         =   "Envio Libro de Ventas SII"
         Index           =   6
      End
      Begin VB.Menu elec 
         Caption         =   "Recibo D.T.E Proveedores"
         Index           =   7
      End
   End
   Begin VB.Menu MConfiguracion 
      Caption         =   "&CONFIGURACION"
      Index           =   99
      Begin VB.Menu MCLocal 
         Caption         =   "Cambiar &Local Activo"
         Shortcut        =   ^L
      End
      Begin VB.Menu MCPermisos 
         Caption         =   "&Permisos de Usuario"
      End
      Begin VB.Menu Mauditoria 
         Caption         =   "Modulo de Auditoria de Usuarios"
      End
      Begin VB.Menu MCClave 
         Caption         =   "&Cambio Clave"
      End
   End
   Begin VB.Menu MSalir 
      Caption         =   "&SALIR"
      Index           =   99
      Begin VB.Menu MSSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Public PASO As Boolean

Private Sub ELEC_Click(Index As Integer)
If Index = 1 Then electro01.Show: electro01.Caption = elec(1).Caption
If Index = 2 Then electro02.Show: electro02.Caption = elec(2).Caption
If Index = 3 Then electro04.Show: electro04.Caption = elec(3).Caption
If Index = 4 Then electro05.Show: electro05.Caption = elec(4).Caption
Rem If Index = 5 Then electro06.Show: electro05.Caption = elec(5).Caption

Call grabaprincipal(elec(Index).Caption)

End Sub

Private Sub MCLocal_Click()
        cambioLocal.Show vbModal
        cambioLocal.Caption = Replace(MCLocal.Caption, "&", "")
        Call grabaprincipal(cambioLocal.Caption)
End Sub

Private Sub MCPermisos_Click()
'    permisosUsuario.Show
'    permisosUsuario.Caption = Replace(MCPermisos.Caption, "&", "")
seguridad2.Show
Call grabaprincipal(seguridad2.Caption)

End Sub






Private Sub MDIForm_Load()
'    Dim saveTitle$
'    If App.PrevInstance Then
'        saveTitle$ = App.Title
'        App.Title = "... duplicate instance."
'        Me.Caption = "... duplicate instance. "
'        AppActivate saveTitle$
'        SendKeys "% R", True
'        End
'    End If
  Me.barraEstado.Panels(1).text = UCase(Me.Caption)
    If PASO = False Then
        barraEstado.Panels(3).text = ""
'        If segu = True Then
'        Seguridad.Show vbModal, Me

             Call revisarmenus(Principal)
            'Call cargaMenuPermisos
'        End If
'        segurity = False
    End If
    sqlventas.audit = False

    cajera = "0000"
'    ChangeRes1.GetMonitorInfo = True
'    resX = ChangeRes1.Xpixels
'    resY = ChangeRes1.Ypixels
'    ChangeRes1.Xpixels = 1024
'    ChangeRes1.Ypixels = 768
'    ChangeRes1.ChangeResolution = True
'    If ChangeRes1.Error = True Then
'        MsgBox "La resolucion de su monitor no puede cambiarse a " & ChangeRes1.Xpixels & " X " & ChangeRes1.Ypixels
'    End If
    Rem Me.Picture = LoadPicture(App.Path & "\trigo.jpg")
    'Call TranslucentForm(Me, 200)
   
End Sub

Private Sub MSCerrar_Click()
    PASO = False
    usuarioSistema = ""
    passwordSistema = ""
    Call UnloadHijos(Me)
     
End Sub

