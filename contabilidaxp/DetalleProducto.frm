VERSION 5.00
Begin VB.Form DetalleProducto 
   Caption         =   "Detalle Producto"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6015
      Begin VB.CommandButton Volver 
         Cancel          =   -1  'True
         Caption         =   "VOLVER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Aceptar 
         Caption         =   "ACEPTAR"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Cantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1940
         Width           =   1575
      End
      Begin VB.TextBox Precio 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1480
         Width           =   1575
      End
      Begin VB.TextBox Precio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1000
         Width           =   1575
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "PRECIO MAYORISTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   3
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton Opcion 
         Caption         =   "PRECIO DETALLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Label Producto 
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label2 
         Caption         =   "CANTIDAD"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "COMPRUEBE DATOS DE VENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "DetalleProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Aceptar_Click()
'GUARDA UNA NUEVA LINEA EN LA NOTA DE VENTA.

    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Dim codigo_producto As String
    Dim total As Double
    Dim precio_venta As Double
    Dim nota_venta As String
    
    With ventas13
        Set cSql.ActiveConnection = db
        nota_venta = .NumeroNota.Caption
        'PREGUNTA SI EXISTE UNA NOTA DE VENTA CON ESTE NUMERO.
        cSql.SQL = "SELECT numero "
        cSql.SQL = cSql.SQL + "FROM notaventa "
        cSql.SQL = cSql.SQL + "WHERE numero='" & nota_venta & "'"
        cSql.Execute
        If cSql.RowsAffected = 0 Then   'SI NO EXISTE LA CREA.
            cSql.SQL = "INSERT INTO notaventa (numero,fecha_creacion,hora_creacion,valides,autor) VALUES ('" & nota_venta & "','" & Format(Now(), "yyyy-mm-dd") & "','" & Format(Now(), "hh:mm:ss") & "','1','DENNIS')"
            cSql.Execute
        End If
        If Precio(0).Enabled = True Then
            precio_venta = CDbl(Precio(0).Text)
            total = precio_venta * CDbl(Cantidad.Text)
        Else
            precio_venta = CDbl(Precio(1).Text)
            total = precio_venta * CDbl(Cantidad.Text)
        End If
        cSql.SQL = "INSERT INTO detallenotaventa (notaventa,producto,cantidad,precio_venta,total) VALUES ('" & nota_venta & "','" & Producto.Caption & "','" & Cantidad.Text & "',float(" & precio_venta & "),'" & total & "')"
        cSql.Execute
        .ControlData.Refresh
        Volver.Value = True
    End With

End Sub

Private Sub Form_Activate()
    Precio(0).SetFocus
End Sub

Private Sub Form_Load()

    With ventas13
        'NOMBRE PRODUCTO
        Producto.Caption = .Productos.TextMatrix(.Productos.Row, 0) & Space(10) & .Productos.TextMatrix(.Productos.Row, 1)
        'PRECIO DETALLE
        Precio(0).Text = Format(.Productos.TextMatrix(.Productos.Row, 4), "###,###,###.00")
        'PRECIO MAYORISTA
        Precio(1).Text = Format(.Productos.TextMatrix(.Productos.Row, 3), "###,###,###.00")
        'CANTIDAD
        Cantidad.Text = "1"
    End With
    
End Sub



Private Sub Opcion_Click(Index As Integer)
    
        Select Case Index
        
            Case 0  'PRECIO DETALLE
                    Precio(Index).Enabled = True
                    Precio(1).Enabled = False
                    Precio(Index).SetFocus
            Case 1  'PRECIO MAYORISTA
                    Precio(Index).Enabled = True
                    Precio(0).Enabled = False
                    Precio(Index).SetFocus
                    
        End Select
End Sub

Private Sub Precio_GotFocus(Index As Integer)
    
    Precio(Index).SelStart = 0
    Precio(Index).SelLength = Len(Precio(Index).Text)
    
End Sub

Private Sub Precio_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode <> 37 And KeyCode <> 39 And KeyCode <> 8 And KeyCode <> 46 Then
        If Not IsNumeric(Chr(KeyCode)) Then
             KeyCode = 0
        End If
    End If

End Sub

Private Sub precio_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Cantidad_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 37 And KeyCode <> 39 And KeyCode <> 8 And KeyCode <> 46 Then
        If Not IsNumeric(Chr(KeyCode)) Then
             KeyCode = 0
        End If
    End If

End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)

    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub


Private Sub Volver_Click()
    Unload Me
End Sub
