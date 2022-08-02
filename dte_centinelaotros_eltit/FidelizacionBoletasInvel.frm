VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form FidelizacionBoletasInvel 
   BackColor       =   &H00FF8080&
   Caption         =   "Ingreso de clientes Fidelizacion Invel"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Fidelizacion de Clientes"
   ScaleHeight     =   6060
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton CMDIMPRIMIR 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5640
      Width           =   2775
   End
   Begin XPFrame.FrameXp ingresaprecio 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4048
      BackColor       =   16761024
      Caption         =   "DATOS DE LA BOLETA"
      BackColor       =   16761024
      ColorBarraAbajo =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   16777215
      Begin VB.TextBox TxtMonto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1200
         Width           =   2610
      End
      Begin VB.TextBox DatoFecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   6
         Top             =   600
         Width           =   2610
      End
      Begin VB.TextBox TxtFolio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1080
         Width           =   1650
      End
      Begin VB.TextBox TxtCaja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   810
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "F2 IMPRIME VACIOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4920
         Width           =   2775
      End
      Begin VB.Label LblMonto 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MONTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   4440
         TabIndex        =   9
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label lblFecha 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   4440
         TabIndex        =   7
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label lblNumero 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nº BOLETA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label lblcaja 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CAJA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1560
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   3135
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5530
      BackColor       =   16761024
      Caption         =   "DATOS DEL CLIENTE"
      BackColor       =   16761024
      ColorBarraAbajo =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   16777215
      Begin VB.TextBox rut2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   16
         Top             =   600
         Width           =   1530
      End
      Begin VB.TextBox nombrecliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1080
         Width           =   6450
      End
      Begin VB.TextBox ciudad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2040
         Width           =   2490
      End
      Begin VB.TextBox direccion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1560
         Width           =   6450
      End
      Begin VB.TextBox celular 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2520
         Width           =   2490
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "F2 IMPRIME VACIOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4920
         Width           =   2775
      End
      Begin VB.Label lbldv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3720
         TabIndex        =   23
         Top             =   600
         Width           =   285
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Rut Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   1560
      End
      Begin VB.Label lbltipocliente 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   840
         Width           =   3735
      End
   End
End
Attribute VB_Name = "FidelizacionBoletasInvel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub celular_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then rut2.SetFocus
End Sub

Private Sub ciudad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then celular.SetFocus
End Sub
Private Sub CMDIMPRIMIR_Click()
    Dim tipofinal As String
        If nombrecliente.text <> "" Then
            RUT_PUNTOS = rut2.text + lbldv.Caption
            NOMBRE_PUNTOS = nombrecliente.text
            DIRECCION_PUNTOS = direccion.text
            CIUDAD_PUNTOS = ciudad.text
            CELULAR_PUNTOS = celular.text
            Call modificacliente(RUT_PUNTOS, NOMBRE_PUNTOS, DIRECCION_PUNTOS, CIUDAD_PUNTOS, CELULAR_PUNTOS)
            Call grabarPUNTOS(empresaActiva, TxtCaja, "BV", Format(TxtFolio, "0000000000"), Format(fechasistema, "yyyy-mm-dd"), 0, RUT_PUNTOS, TxtMonto)
            
        Else
            RUT_PUNTOS = Empty
            NOMBRE_PUNTOS = Empty
            DIRECCION_PUNTOS = Empty
            CIUDAD_PUNTOS = Empty
            CELULAR_PUNTOS = Empty
            
        End If
End Sub

Private Sub direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And direccion.text <> "" Then ciudad.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        If nombrecliente.text <> "" Then
            CMDIMPRIMIR_Click
        End If
    Case 27
        Unload Me
End Select
End Sub
Private Sub Form_Load()
    rut2.text = Mid(rutcredito, 1, 9)
    lbldv.Caption = Mid(rutcredito, 10, 1)
End Sub
Private Sub localenvio_GotFocus()
    nombrelocal.Caption = ""
End Sub
Sub modificacliente(rut, nombre, direccion, ciudad, celular)
    Dim condicion As String
    Dim op As Integer
    Dim cad As String
    Dim i As Long
    Dim campos(40, 5) As String
                campos(0, 0) = "rut"
                campos(1, 0) = "sucursal"
                campos(2, 0) = "nombre"
                campos(3, 0) = "direccion"
                campos(4, 0) = "ciudad"
                campos(5, 0) = "celular"
                campos(6, 0) = ""
                campos(0, 1) = rut
                campos(1, 1) = "0"
                campos(2, 1) = nombre
                campos(3, 1) = direccion
                campos(4, 1) = ciudad
                campos(5, 1) = celular
                campos(0, 2) = "sv_maestroclientes"
                condicion = "rut='" + rut + "' and sucursal='0' "
                op = 5
                sqlventas.response = campos
                Set sqlventas.conexion = ventas
                
                Call sqlventas.sqlventas(op, condicion)
                If sqlventas.Status = 0 Then
                op = 3
                Else
                op = 2
                condicion = ""
                End If
                sqlventas.response = campos
                Set sqlventas.conexion = ventas
                
                Call sqlventas.sqlventas(op, condicion)
    
End Sub

Sub LEERCLIENTEPUNTOS(rut)
      
    Dim condicion As String
    Dim op As Integer
    Dim cad As String
    Set sql = New sqlventas.sqlventa
    Dim i As Long
    Dim campos(40, 5) As String
    
                campos(0, 0) = "rut"
                campos(1, 0) = "sucursal"
                campos(2, 0) = "nombre"
                campos(3, 0) = "direccion"
                campos(4, 0) = "ciudad"
                campos(5, 0) = "celular"
                campos(6, 0) = "tipocliente"
                campos(0, 2) = "sv_maestroclientes"
                condicion = "rut='" + rut + "' and sucursal='0' "
                
                op = 5
                sqlventas.response = campos
                Set sqlventas.conexion = ventas
               
                Call sqlventas.sqlventas(op, condicion)
                If sqlventas.Status = 0 Then
                
                    nombrecliente.text = sqlventas.response(2, 3)
                    direccion.text = sqlventas.response(3, 3)
                    ciudad.text = sqlventas.response(4, 3)
                    celular.text = sqlventas.response(5, 3)
                    lbltipocliente.Caption = sqlventas.response(6, 3)
                    TxtCaja.SetFocus
                Else
                    nombrecliente.text = Empty
                    direccion.text = Empty
                    ciudad.text = Empty
                    celular.text = Empty
                    lbltipocliente.Caption = "01"
                
                End If
End Sub
Private Sub nombrecliente_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If KeyAscii = 13 And nombrecliente.text <> "" Then
       direccion.SetFocus
        
     End If
End Sub
Private Sub rut2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And rut2.text <> "" Then
        rut2.text = ceros(rut2)
        lbldv.Caption = rut(rut2.text)
      
     Call LEERCLIENTEPUNTOS(rut2.text & lbldv.Caption)
      
    End If
End Sub

Sub grabarPUNTOS(loc, caja, TIPO, NUMERO, fecha, puntos, rut, MONTO)
    Dim condicion As String
    Dim op As Integer
    Dim cad As String
    Dim i As Long
    Dim campos(40, 5) As String
    
    ''''''''''''''''''''''''''''''''''
    'Graba el detalle del documento
    ''''''''''''''''''''''''''''''''''
    'boletas(0) = leerUltimaBoleta
     
                campos(0, 0) = "rut"
                campos(1, 0) = "local"
                campos(2, 0) = "tipo"
                campos(3, 0) = "numero"
                campos(4, 0) = "fecha"
                campos(5, 0) = "caja"
                campos(6, 0) = "monto"
                campos(7, 0) = "puntos"
                campos(8, 0) = ""
                
                campos(0, 1) = rut
                campos(1, 1) = loc
                campos(2, 1) = TIPO
                campos(3, 1) = NUMERO
                campos(4, 1) = Format(fecha, "yyyy-mm-dd")
                campos(5, 1) = caja
                campos(6, 1) = Replace(MONTO, ",", "")
                campos(7, 1) = puntos
                
                campos(0, 2) = "sv_puntos"
                op = 2
                sqlventas.response = campos
                Set sqlventas.conexion = ventas
                Call sqlventas.sqlventas(op, condicion)
        
End Sub
