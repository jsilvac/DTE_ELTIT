VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form maestro09 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Proveedores"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc mprov 
      Height          =   330
      Left            =   7320
      Top             =   6720
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
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   5655
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   7455
      Begin MSMask.MaskEdBox rutproveedores 
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         MaxLength       =   10
         Mask            =   "99.999.999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   15
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   16
         Top             =   1680
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   17
         Top             =   2040
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   18
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   5
         Left            =   5400
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   22
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   23
         Top             =   3480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   8
         Left            =   2520
         TabIndex        =   24
         Top             =   3840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskproveedores 
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   25
         Top             =   4200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14213628
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Fono (2)"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRATO DE PUBLICIDAD"
         Height          =   495
         Left            =   6000
         TabIndex        =   13
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. Visitas Mes"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Convenio"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comuna"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fono (1)"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creacion de Proveedores"
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
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RUT"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   5655
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   2400
         Width           =   855
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1215
      Left            =   720
      TabIndex        =   12
      Top             =   6720
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\remuxp\barra_opciones.swf"
      Src             =   "c:\remuxp\barra_opciones.swf"
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
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   5655
      Left            =   720
      Top             =   600
      Width           =   7455
   End
End
Attribute VB_Name = "maestro09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then ayudamaestrodeproveedores.Show


End Sub


Private Sub Form_Activate()

    rutproveedores.SetFocus
    
End Sub

Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 685
    Me.ScaleHeight = 578
    'CARGA LA BARRA DE TITULO
    Rem swfBarra.Width = Me.ScaleWidth
    Rem swfBarra.Height = Me.ScaleHeight
    Rem Call swfBarra.LoadMovie(0, Interfaces.path + "\Data\Barra_Titulo.swf")
    'CARGA EL BOTON NUEVO
    Rem Call swfNuevo.LoadMovie(0, Interfaces.path + "\Data\btn_nuevo.swf")
    'OBTENER POSICION DEL FORMULARIO
    posx2 = Me.ScaleWidth
    posy2 = Me.ScaleHeight
    posx1 = (Interfaces.equiAncho(Screen.Width) - posx2) \ 2
    posy1 = (Interfaces.equiAlto(Screen.Height) - posy2) \ 2
    'CARGADO DE LA IMAGEN DEGRADADA
    apis.Degradado Me, Principal, posx1, posx2, posy1, posy2, 150
       
    'FLAG = 0 SE GRABA/MODIFICA  FLAG = 1 YA SE GUARDO EN BD
    
    Call Conecta_Maestro_Proveedores
    'Call Lista_Cuentas_Mayor
    Call Conectar_BD
    
End Sub

Private Sub Label12_Click()

contratopublicidad.Show

End Sub

Private Sub mskproveedores_GotFocus(Index As Integer)
'----------------------------------------------------
'   CAMBIA EL COLOR DE FONDO DEL LOS CUADRO DE
'   TEXTO CADA VEZ QUE OBTENGAN O PIERDAN EL FOCO.
'----------------------------------------------------

    mskproveedores(Index).BackColor = &HFFFF&

End Sub

Private Sub mskproveedores_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'VALIDA LA ENTRADA DE CADA CUADRO DE TEXTO
    
    Select Case KeyCode
    
        
        Case 38:    '>> TECLA ARRIBA
        
            If Index > 0 Then
                SendKeys "+{tab}"
            Else
                If rutproveedores.Enabled = True Then
                    rutproveedores.SetFocus
                End If
            End If
        
                
        Case 40:    '>> TECLA ABAJO
        
            If Index < 9 Then
                SendKeys "{tab}"
                If Index = 0 Then
                
                End If
            Else
                'RUTINA DE GRABADO O ACTUALIZADO DE DATOS.
                Call Manejo_Datos_Proveedores
                Call Limpia_Formulario_Proveedores
                rutproveedores.Enabled = True
                rutproveedores.SetFocus
            End If
            
        Case 13:    '>> TECLA ENTER
        
            If Index <> 9 Then 'SI NO ES EL ULTIMO PASA AL SIGUIENTE, SINO.......
                SendKeys "{tab}"
            Else
                'RUTINA DE GRABADO O ACTUALIZADO DE DATOS.
                Call Manejo_Datos_Proveedores
                Call Limpia_Formulario_Proveedores
                rutproveedores.Enabled = True
                rutproveedores.SetFocus
            End If
            
    End Select
        
    If Not IsNumeric(Chr(KeyCode)) Then
         KeyCode = 0
    End If
    
End Sub

Private Sub rutproveedores_GotFocus()

'-----------------------------------------------------------------
'      COLOR CUANDO TOMA EL FOCO CODIGO
'-----------------------------------------------------------------

    rutproveedores.BackColor = &HFFFF&

'-----------------------------------------------------------------

End Sub

Private Sub rutproveedores_KeyDown(KeyCode As Integer, Shift As Integer)
' Uso de ENTER y Flecha Abajo para moverse en Formulario

    ' tecla F2 para llamar ayuda
    
    If KeyCode = 113 Then ayudamaestrodeproveedores.Show
    
    '--------------------------------------------------
     Select Case KeyCode
                Case 38:    'TECLA ARRIBA
                        'SendKeys "+{tab}"
                        
                Case 40:    'TECLA ABAJO
                        If Val(Mid(rutproveedores.Text, 1, 2)) < 1 Or Val(Mid(rutproveedores.Text, 4, 3)) < 1 Or InStr(Mid(rutproveedores.Text, 8, 3), " ") > 0 Then
                            rutproveedores.SetFocus
                        Else
                            Carga_Cuenta_Proveedores (Mid(rutproveedores.Text, 1, 2) & Mid(rutproveedores.Text, 4, 3) & Mid(rutproveedores.Text, 8, 3))
                            'SendKeys "{tab}"
                            mskproveedores(0).SetFocus
                        End If
                        
                Case 13:    'TECLA ENTER
                        If Val(Mid(rutproveedores.Text, 1, 2)) < 1 Or Val(Mid(rutproveedores.Text, 4, 3)) < 1 Or InStr(Mid(rutproveedores.Text, 8, 3), " ") > 0 Then
                            rutproveedores.SetFocus
                        Else
                            Carga_Cuenta_Proveedores (Mid(rutproveedores.Text, 1, 2) & Mid(rutproveedores.Text, 4, 3) & Mid(rutproveedores.Text, 8, 3))
                            'SendKeys "{tab}"
                            mskproveedores(0).SetFocus
                        End If
            End Select
End Sub

Private Sub rutproveedores_KeyPress(KeyAscii As Integer)

' tecla enter continua
    If KeyAscii = 13 And rutproveedores.Text <> "" Then
        mskproveedores(0).SetFocus
             
    End If
    
' tecla ESC sale del formulario
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub rutproveedores_LostFocus()
 
    rutproveedores.BackColor = &HD8E1FC

End Sub
