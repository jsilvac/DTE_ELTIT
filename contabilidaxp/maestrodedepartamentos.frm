VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form maestro03 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Departamentos"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc md 
      Height          =   495
      Left            =   7920
      Top             =   4560
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
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   6735
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   4
         Tag             =   "margenteorico"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   3
         Tag             =   "descuentoventa"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   2
         Tag             =   "codigoseccion"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Tag             =   "nombre"
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   0
         Tag             =   "codigodepto"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDF5E6&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "(%) Margen Teorico"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(%) Descuento en Venta"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
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
         Caption         =   "Seccion"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creacion de Departamentos"
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
         TabIndex        =   6
         Top             =   120
         Width           =   2415
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1215
      Left            =   1320
      TabIndex        =   12
      Top             =   3960
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
      Left            =   1320
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "maestro03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dato_GotFocus(Index As Integer)
'   CAMBIA EL COLOR DE FONDO DEL LOS CUADRO DE
'   TEXTO CADA VEZ QUE OBTENGAN O PIERDAN EL FOCO.
'----------------------------------------------------
    
        dato(Index).BackColor = &HFFFF&
        dato(Index).SelStart = 0
        dato(Index).SelLength = dato(Index).MaxLength
    
End Sub

Private Sub dato_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'VALIDA LA ENTRADA DE CADA CUADRO DE TEXTO
   
    Dim condicion As String
    Dim op As Integer
        
    Select Case KeyCode
    
        Case 38:    '>> TECLA ARRIBA
        
                
                If Index > 0 Then
                    evento = 1
                    If dato(Index - 1).Enabled = True Then
                        SendKeys "+{tab}"
                    End If
                End If
            
        Case 40:    '>> TECLA ABAJO
        
                If Index <> 4 Then
                    If Index = 0 And dato(0).text <> "" Then
                                'NOMBRE DE LOS CAMPOS POR LOS QUE SE PREGUNTA.
                                campos(0, 0) = dato(0).Tag 'CODIGO
                                campos(1, 0) = dato(1).Tag 'NOMBRE
                                campos(2, 0) = dato(2).Tag 'SECCION
                                campos(3, 0) = dato(3).Tag 'DESCUENTO DE VENTA
                                campos(4, 0) = dato(4).Tag 'MARGEN TEORICO
'                                campos(5, 0) = dato(5).Tag '...
'                                campos(6, 0) = dato(6).Tag '...
'                                campos(7, 0) = dato(7).Tag '...
'                                campos(8, 0) = dato(8).Tag '...
                                'TABLA DE LA BASE DE DATOS.
                                campos(0, 2) = "maestrodepartamentos"
                                'CONDICION DE LA CONSULTA.
                                condicion = "codigodepto=" + "'" + dato(0).text + "'"
                                'OPCION CON QUE SE LLAMA LA FUNCION (5=LEE, OSEA RESCATA DATOS).
                                op = 5
                                SQLUTIL.datos = campos
                                Set SQLUTIL.conexion = db
                                Call SQLUTIL.SQLUTIL(op, condicion)
                                If SQLUTIL.estado = 0 Then 'ENCONTRO DATOS
                                    'dato(0).Text = campos(0, 3)
                                    dato(1).text = SQLUTIL.datos(1, 3)
                                    dato(2).text = SQLUTIL.datos(2, 3)
                                    dato(3).text = SQLUTIL.datos(3, 3)
                                    dato(4).text = SQLUTIL.datos(4, 3)
'                                    dato(5).Text = SQLUTIL.datos(5, 3)
'                                    dato(6).Text = SQLUTIL.datos(6, 3)
'                                    dato(7).Text = SQLUTIL.datos(7, 3)
'                                    dato(8).Text = SQLUTIL.datos(8, 3)
                                    'flash_opciones.Visible = True
                                End If
                    End If
                    SendKeys "{tab}"
                Else
                    'RUTINA DE GRABADO O ACTUALIZADO DE DATOS.
                    If dato(0).text <> "" And dato(1).text <> "" And dato(2).text <> "" And dato(3).text <> "" And dato(4).text <> "" Then
                        Call Funciones_Forms_M_Departamentos.Manejo_Datos_Departamentos
                        'Call funciones_CuentasCorrientes.Lista_Cuentas_Corrientes
                        Call Funciones_Forms_M_Departamentos.Limpia_Formulario_Departamentos
                        dato(0).SetFocus
                    End If
                End If
        
        Case 13:    '>> TECLA ENTER
                Call dato_KeyDown(Index, 40, Shift)
                
                
    End Select
        
    'ACEPTA LAS TECLAS <<BACKSPACE>>,<<FLECHA IZQUIERDA>>,<<FLECHA DERECHA>> Y <<SUPRIMIR>>.
    If KeyCode <> 37 And KeyCode <> 39 And KeyCode <> 8 And KeyCode <> 46 Then
        If Not IsNumeric(Chr(KeyCode)) Then
             KeyCode = 0
        End If
    End If
End Sub

Private Sub dato_KeyPress(Index As Integer, KeyAscii As Integer)
'VALIDA ENTRADAS EN LOS CUADROS DE TEXTO.
    Select Case Index
        Case 0: 'codigo
                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
                    KeyAscii = 0
                End If
                
        Case 1: 'nombre
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                
        Case 2: 'seccion
                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
                    KeyAscii = 0
                End If
                
        Case 3: 'descuento venta
                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
                    KeyAscii = 0
                End If
                
        Case 4: 'margen teorico
                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
                    KeyAscii = 0
                End If
                
'        Case 5: 'linea
'                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
'                    KeyAscii = 0
'                End If
'
'        Case 6: 'impuesto
'                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
'                    KeyAscii = 0
'                End If
                
    End Select
End Sub

Private Sub dato_LostFocus(Index As Integer)
        dato(Index).BackColor = &HD8E1FC
        dato(Index).SelStart = 0
        Pregunta (Index)
End Sub
Sub Pregunta(Index As Integer)

    If evento = 0 Then
        If dato(Index).text = "" And dato(Index).ToolTipText = "Obligatorio" Then
            If dato(Index).Enabled = True Then
                dato(Index).SetFocus
                evento = 1
                Exit Sub
            End If
        End If
    End If
    evento = 0
End Sub

Private Sub Form_Load()

Dim posx1, posx2, posy1, posy2 As Long
    'TAMAÑO Y POSICION DEL FORMULARIO
    Me.ScaleWidth = 674
    Me.ScaleHeight = 586
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

   Call Funciones_Forms_M_Departamentos.Conecta_Maestro_Departamentos
    Call Conectar_BD

End Sub


