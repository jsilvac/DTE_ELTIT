VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form maestro06 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Locales"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc ml 
      Height          =   330
      Left            =   7320
      Top             =   6360
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
      Height          =   4935
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   7335
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   7
         Tag             =   "auditoria"
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   6
         Tag             =   "rut"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   5
         Tag             =   "tlocal"
         Top             =   2760
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   4
         Tag             =   "ciudad"
         Top             =   2400
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   3
         Tag             =   "comuna"
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   2
         Tag             =   "direccion"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Tag             =   "nombre"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox dato 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E1FC&
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "codigolocal"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Parametros Financieros"
         Height          =   495
         Left            =   6240
         TabIndex        =   19
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "T. Local"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "RUT"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Auditoria"
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creacion de Locales"
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
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   4935
         Left            =   0
         Top             =   0
         Width           =   7335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comuna"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   2400
         Width           =   855
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1215
      Left            =   600
      TabIndex        =   18
      Top             =   5760
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
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   4935
      Left            =   600
      Top             =   480
      Width           =   7335
   End
End
Attribute VB_Name = "maestro06"
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
        
                If Index <> 7 Then
                    If Index = 0 And dato(0).text <> "" Then
                                'NOMBRE DE LOS CAMPOS POR LOS QUE SE PREGUNTA.
                                campos(0, 0) = dato(0).Tag 'CODIGO
                                campos(1, 0) = dato(1).Tag 'NOMBRE
                                campos(2, 0) = dato(2).Tag 'DIRECCION
                                campos(3, 0) = dato(3).Tag 'COMUNA
                                campos(4, 0) = dato(4).Tag 'CIUDAD
                                campos(5, 0) = dato(5).Tag 'T.LOCAL
                                campos(6, 0) = dato(6).Tag 'RUT
                                campos(7, 0) = dato(7).Tag 'AUDITORIA
'                                campos(8, 0) = dato(8).Tag 'PRECIO VENTA MAYOR
'                                campos(9, 0) = dato(9).Tag 'PRECIO VENTA DETALLE
'                                campos(10, 0) = dato(10).Tag 'STOCK CRITICO
                                'TABLA DE LA BASE DE DATOS.
                                campos(0, 2) = "maestrolocales"
                                'CONDICION DE LA CONSULTA.
                                condicion = "codigolocal=" + "'" + dato(0).text + "'"
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
                                    dato(5).text = SQLUTIL.datos(5, 3)
                                    dato(6).text = SQLUTIL.datos(6, 3)
                                    dato(7).text = SQLUTIL.datos(7, 3)
'                                    dato(8).text = SQLUTIL.datos(8, 3)
'                                    dato(9).text = SQLUTIL.datos(9, 3)
'                                    dato(10).text = SQLUTIL.datos(10, 3)
                                    'flash_opciones.Visible = True
                                End If
                    End If
                    SendKeys "{tab}"
                Else
                    'RUTINA DE GRABADO O ACTUALIZADO DE DATOS.
                    If dato(0).text <> "" And dato(1).text <> "" And dato(2).text <> "" And dato(3).text <> "" And dato(4).text <> "" And dato(5).text <> "" And dato(6).text <> "" And dato(7).text <> "" Then
                        Call Funciones_Forms_M_Locales.Manejo_Datos_Locales
                        'Call funciones_CuentasCorrientes.Lista_Cuentas_Corrientes
                        Call Funciones_Forms_M_Locales.Limpia_Formulario_Locales
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
                
        Case 2: 'direccion
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                
        Case 3: 'comuna
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                
        Case 4: 'ciudad
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                
        Case 5: 'tlocal
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                
        Case 6: 'rut
                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
                    KeyAscii = 0
                End If

        Case 7: 'auditoria
                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
                    KeyAscii = 0
                End If

'        Case 8: 'precio venta mayorista
'                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
'                    KeyAscii = 0
'                End If
'
'        Case 9: 'precio venta detalle
'                If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
'                    KeyAscii = 0
'                End If
'
'        Case 10: 'stock critico
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
    Me.ScaleHeight = 533
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

    Call Funciones_Forms_M_Locales.Conecta_Maestro_Locales
    Call Conectar_BD
    'Call Desabilita_Locales  ' desabilito cajas de textos de locales

End Sub

