VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form maestro05 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Impuestos"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc mi 
      Height          =   495
      Left            =   7680
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Height          =   2895
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   6735
      Begin VB.Label codigo 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 ( ? )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3240
         TabIndex        =   0
         Top             =   720
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2895
         Left            =   0
         Top             =   0
         Width           =   6735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creacion de Impuestos"
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
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   3015
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1215
      Left            =   1200
      TabIndex        =   2
      Top             =   3840
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
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   2895
      Left            =   1080
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "maestro05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub codimpuesto_GotFocus()

'-----------------------------------------------------------------
'      COLOR CUANDO TOMA EL FOCO CODIGO
'-----------------------------------------------------------------

    codimpuesto.BackColor = &HFFFF&

'-----------------------------------------------------------------
End Sub

Private Sub codimpuesto_KeyDown(KeyCode As Integer, Shift As Integer)

' Uso de ENTER y Flecha Abajo para moverse en Formulario

    ' tecla F2 para llamar ayuda
    
    If KeyCode = 113 Then ayudamaestrodeimpuestos.Show
    
    '--------------------------------------------------
         
     Select Case KeyCode
                Case 38:    'TECLA ARRIBA
                        'SendKeys "+{tab}"
                        
                Case 40:    'TECLA ABAJO
                        If codimpuesto.text = "" Then
                            codimpuesto.SetFocus
                        Else
                            
                            Carga_Cuenta_Impuestos (codimpuesto.text)
                            'SendKeys "{tab}"
                            txtimpuesto(0).SetFocus
                        End If
                        
                Case 13:    'TECLA ENTER
                        If codimpuesto.text = "" Then
                            codimpuesto.SetFocus
                        Else
                            Carga_Cuenta_Impuestos (codimpuesto.text)
                            'SendKeys "{tab}"
                            txtimpuesto(0).SetFocus
                        End If
                        
      End Select
    
End Sub

Private Sub codimpuesto_KeyPress(KeyAscii As Integer)
' tecla enter continua
    If KeyAscii = 13 And codimpuesto.text <> "" Then
        txtimpuesto(0).SetFocus
             
    End If
    
' tecla ESC sale del formulario
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub codimpuesto_LostFocus()
' perdida de color

    codimpuesto.BackColor = &HD8E1FC

End Sub

Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 616
    Me.ScaleHeight = 380
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

    
    Call Conecta_Maestro_Impuestos
    Call Conectar_BD



End Sub

Private Sub txtimpuesto_GotFocus(Index As Integer)
'----------------------------------------------------
'   CAMBIA EL COLOR DE FONDO DEL LOS CUADRO DE
'   TEXTO CADA VEZ QUE OBTENGAN O PIERDAN EL FOCO.
'----------------------------------------------------

    txtimpuesto(Index).BackColor = &HFFFF&

End Sub

Private Sub txtimpuesto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'VALIDA LA ENTRADA DE CADA CUADRO DE TEXTO

    Select Case KeyCode
    
        Case 38:    '>> TECLA ARRIBA
        
            If Index > 0 Then
                SendKeys "+{tab}"
            Else
                If codimpuesto.Enabled = True Then
                    codimpuesto.SetFocus
                End If
            End If
        
        Case 40:    '>> TECLA ABAJO
        
            If Index < 1 Then
                SendKeys "{tab}"
                If Index = 0 Then
                
                End If
            Else
                'RUTINA DE GRABADO O ACTUALIZADO DE DATOS.
                Rem Call Manejo_Datos_Impuestos
                'Call Lista_Cuentas_Mayor
                Call Limpia_Formulario_Impuestos
                codimpuesto.Enabled = True
                codimpuesto.SetFocus
            End If
            
        Case 13:    '>> TECLA ENTER
        
            If Index <> 1 Then 'SI NO ES EL ULTIMO PASA AL SIGUIENTE, SINO.......
                SendKeys "{tab}"
            Else
                'RUTINA DE GRABADO O ACTUALIZADO DE DATOS.
                Call Manejo_Datos_Impuestos
                'Call Lista_Cuentas_Mayor
                Call Limpia_Formulario_Impuestos
                codimpuesto.Enabled = True
                codimpuesto.SetFocus
            End If
            
    End Select
        
    If Not IsNumeric(Chr(KeyCode)) Then
         KeyCode = 0
    End If
End Sub

Private Sub txtimpuesto_KeyPress(Index As Integer, KeyAscii As Integer)
'VALIDA ENTRADAS EN LOS CUADROS DE TEXTO.

    Select Case Index
        Case 0: 'Nombre
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            
        Case 1: 'Porcentaje
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
                
    End Select
End Sub

Private Sub txtimpuesto_LostFocus(Index As Integer)
    
    txtimpuesto(Index).BackColor = &HD8E1FC

End Sub
