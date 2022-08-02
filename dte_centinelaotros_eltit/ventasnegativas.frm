VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ventasnegativas 
   BackColor       =   &H00FF8080&
   Caption         =   "PANTALLA HISTORICO VENTAS NEGATIVAS"
   ClientHeight    =   7185
   ClientLeft      =   480
   ClientTop       =   1650
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11400
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "RETORNO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   2850
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4485
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7911
      BackColor       =   16761024
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid1 
         Height          =   4470
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   7885
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2990
      BackColor       =   16744576
      Caption         =   "Periodo"
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
      Begin VB.CheckBox locales 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos los Locales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FF8080&
         TabIndex        =   16
         Top             =   1080
         Width           =   2415
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
         Left            =   10080
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "proveedor"
         Top             =   360
         Width           =   375
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
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   360
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
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   360
         Width           =   735
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
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   360
         Width           =   735
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
         Left            =   7080
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   360
         Width           =   375
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
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   360
         Width           =   375
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
         Height          =   285
         Left            =   1500
         MaxLength       =   2
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5280
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8280
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Local"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   60
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblLocal 
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
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   5775
      End
   End
End
Attribute VB_Name = "ventasnegativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call cabezaInforme("", Grid1, "LISTADO DE VENTAS NEGATIVAS", 1)
        Grid1.PageSetup.HeaderMargin = 1
        Grid1.PageSetup.TopMargin = 1
        Grid1.PageSetup.LeftMargin = 1.5
        Grid1.PageSetup.RightMargin = 1
        Grid1.PageSetup.PrintFixedRow = True
        Grid1.PageSetup.BlackAndWhite = True
        Grid1.PageSetup.Orientation = cellLandscape
        Grid1.Range(0, 0, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Grid1.PageSetup.FooterMargin = 2
        Grid1.PageSetup.PrintFixedRow = True
        Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
        Grid1.PageSetup.FooterAlignment = cellRight
        Grid1.PageSetup.FooterFont.Name = "Verdana"
        Grid1.PageSetup.FooterFont.Size = 7
        
        Call verificaImpresora(5, Grid1)

End Sub
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
    
    Private Sub dato7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
    End Sub
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
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
 Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            lblLocal.Caption = leerNombreEmpresa(dato1.text)
            If lblLocal.Caption <> "" Then
                rubroAuditoria = leerRubro(dato1.text)
                Call ConectarAuditoria(servidor, rubroAuditoria, usuario, password, dato1.text)
                SendKeys "{Tab}"
            End If
        End If
    End Sub
      
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "00" Then
                dato3.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "0000" Then
                dato4.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "00" Then
                dato6.text = Format(fechasistema, "mm")
            End If
            dato7.SetFocus
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)

        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato7.text = ceros(dato7)
            If dato7.text = "0000" Then
                dato7.text = Format(fechasistema, "yyyy")
            End If
            Lee_negativos
        End If
    End Sub
Private Sub Form_Load()
CARGAGRILLA

End Sub
Sub CARGAGRILLA()
    Grid1.Cols = 8
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 50
    Grid1.Column(2).Width = 50
    Grid1.Column(3).Width = 90
    Grid1.Column(4).Width = 90
    Grid1.Column(5).Width = 90
    Grid1.Column(6).Width = 200
    Grid1.Column(7).Width = 200
   
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
    
    Grid1.Cell(0, 1).text = "LOCAL"
    Grid1.Cell(0, 2).text = "TP"
    Grid1.Cell(0, 3).text = "FECHA"
    Grid1.Cell(0, 4).text = "NUMERO"
    Grid1.Cell(0, 5).text = "CODIGO"
    Grid1.Cell(0, 6).text = "DESCRIPCION"
    Grid1.Cell(0, 7).text = "GLOSA"
    
    
    Grid1.Range(0, 1, 0, 7).Alignment = cellLeftGeneral
    Grid1.Range(0, 1, 0, 7).FontSize = 7
    Grid1.Range(0, 1, 0, 7).FontBold = True
    Grid1.Range(0, 1, 0, 7).Borders(cellEdgeBottom) = cellThick
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
Grid1.Rows = 1
End Sub

Sub Lee_negativos()
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
   
    Dim linea As Double
    Dim desde1 As String
    Dim hasta1 As String
    
            desde1 = dato4.text & "-" & dato3.text & "-" & dato2.text
            hasta1 = dato7.text & "-" & dato6.text & "-" & dato5.text
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT local,tipo,fecha,numero,codigo,descripcion,glosa FROM sv_documento_detalle_" + empresaActiva + " "
        If locales.Value = "0" Then
        csql.sql = csql.sql + "WHERE local='" & dato1.text & "' and fecha between '" & desde1 & "' and '" & hasta1 & "' and glosa <> '' "
        Else
        csql.sql = csql.sql + "WHERE  fecha between '" & desde1 & "' and '" & hasta1 & "' and glosa <> '' "
       
        End If
        csql.Execute
          Grid1.Rows = csql.RowsAffected + 1
        Grid1.AutoRedraw = False
        linea = 0
        If csql.RowsAffected > 0 Then
            
            Set resultados = csql.OpenResultset
           
            While Not resultados.EOF
           linea = linea + 1
                Grid1.Cell(linea, 1).text = resultados(0)
                Grid1.Cell(linea, 2).text = resultados(1)
                Grid1.Cell(linea, 3).text = resultados(2)
                Grid1.Cell(linea, 4).text = resultados(3)
                Grid1.Cell(linea, 5).text = resultados(4)
                Grid1.Cell(linea, 6).text = resultados(5)
                Grid1.Cell(linea, 7).text = resultados(6)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        Else
            
        End If
        Grid1.AutoRedraw = True
        Grid1.Refresh
End Sub

Private Sub locales_Click()
If locales.Value = 1 Then
dato1.Enabled = False
lblLocal.Caption = ""
dato2.SetFocus
Else
dato1.Enabled = True
dato1.SetFocus
End If

End Sub
