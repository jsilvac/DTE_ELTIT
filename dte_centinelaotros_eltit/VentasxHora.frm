VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form VentasxHoras 
   BackColor       =   &H00FF8080&
   Caption         =   "PANTALLA HISTORICO VENTAS X HORAS"
   ClientHeight    =   9510
   ClientLeft      =   480
   ClientTop       =   1650
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   12675
   StartUpPosition =   2  'CenterScreen
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
      Height          =   345
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   2295
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
      Height          =   345
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      Width           =   2370
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7605
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13414
      BackColor       =   16761024
      Caption         =   "DETALLES"
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
         Height          =   7215
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   12726
         Cols            =   7
         DefaultFontSize =   9
         DefaultFontBold =   -1  'True
         GridColor       =   16761024
         Rows            =   26
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   2143
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Ventas X Sectores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton todas 
         BackColor       =   &H00FF8080&
         Caption         =   "TODAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8880
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton facturas 
         BackColor       =   &H00FF8080&
         Caption         =   "FACTURAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton boletas 
         BackColor       =   &H00FF8080&
         Caption         =   "BOLETAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox detallados 
         BackColor       =   &H00FF8080&
         Caption         =   "Detallado"
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
         Left            =   2640
         MaskColor       =   &H00FF8080&
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
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
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   5310
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1665
         Visible         =   0   'False
         Width           =   330
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
Attribute VB_Name = "VentasxHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub boletas_Click()
CARGAGRILLA
If lbllocal.Caption <> "" And dato1.text <> "" And dato2.text <> "" And dato3.text <> "" And dato4.text <> "" And dato5.text <> "" And dato6.text <> "" And dato7.text <> "" Then
Lee_horas
End If

End Sub
Private Sub facturas_Click()
CARGAGRILLA
If lbllocal.Caption <> "" And dato1.text <> "" And dato2.text <> "" And dato3.text <> "" And dato4.text <> "" And dato5.text <> "" And dato6.text <> "" And dato7.text <> "" Then
Lee_horas
End If
End Sub









Private Sub Option1_Click()
CARGAGRILLA2
Lee_horasXsector
End Sub

Private Sub todas_Click()
CARGAGRILLA
If lbllocal.Caption <> "" And dato1.text <> "" And dato2.text <> "" And dato3.text <> "" And dato4.text <> "" And dato5.text <> "" And dato6.text <> "" And dato7.text <> "" Then
Lee_horas
End If
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call cabezaInforme("", Grid1, "LISTADO DE VENTAS X HORAS", 1)
        Grid1.DefaultFont.Size = 9
        Grid1.PageSetup.HeaderMargin = 1
        Grid1.PageSetup.TopMargin = 1
        Grid1.PageSetup.LeftMargin = 1.5
        Grid1.PageSetup.RightMargin = 1
        Grid1.PageSetup.PrintFixedRow = True
        Grid1.PageSetup.BlackAndWhite = True
        Grid1.PageSetup.Orientation = cellPortrait
        Grid1.Range(0, 0, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Grid1.PageSetup.FooterMargin = 2
        Grid1.PageSetup.PrintFixedRow = True
        Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " + usuarioSistema
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
         If KeyCode = vbKeyF2 Then
         Call ayudaEmpresa(dato1)
         
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
            lbllocal.Caption = leerNombreEmpresa(dato1.text)
            If lbllocal.Caption <> "" Then
                rubroAuditoria = leerRubro(dato1.text)
                Call ConectarAuditoria(servidor, rubroAuditoria, usuario, password, dato1.text)
                SendKeys "{Tab}"
                If dato1.text = "20" Then Option1.Visible = True
                
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
            Call CARGAGRILLA
            
            Call Lee_horas
        End If
    End Sub

Private Sub detallados_Click()
CARGAGRILLA
If dato1.text <> "" Then
Call Lee_horas
Else
If locales.Value = "0" Then
dato1.SetFocus
Else
dato2.SetFocus
 End If
End If
End Sub



Private Sub Form_Load()

CARGAGRILLA

End Sub
Sub CARGAGRILLA()
    Grid1.Cols = 7
    Grid1.Column(2).Alignment = cellRightCenter
    Grid1.Column(3).Alignment = cellRightCenter
    Grid1.Column(4).Alignment = cellRightCenter
    Grid1.Column(5).Alignment = cellRightCenter
    Grid1.Column(6).Alignment = cellRightCenter
    
   
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 83
    Grid1.Column(2).Width = 100
    Grid1.Column(3).Width = 130
    Grid1.Column(4).Width = 120
    Grid1.Column(5).Width = 120
    Grid1.Column(6).Width = 120
  
   If facturas.Value = True Then
   Grid1.Cell(0, 2).text = "Nº FACTURAS"
     End If
   If boletas.Value = True Then
    Grid1.Cell(0, 2).text = "Nº BOLETAS"
    End If
   If todas.Value = True Then
    Grid1.Cell(0, 2).text = "Nº BOl y FAC"
   End If
   
   
    Grid1.Cell(0, 1).text = "HORA"
    Grid1.Cell(0, 3).text = "TOTAL X HORA"
    Grid1.Cell(0, 4).text = "DESCUENTO"
    Grid1.Cell(0, 5).text = "NULAS"
    Grid1.Cell(0, 6).text = "CREDITO"
   
   
    Grid1.Range(0, 1, 0, 6).Alignment = cellLeftGeneral
    Grid1.Range(0, 1, 0, 6).FontSize = 9
    Grid1.Range(0, 1, 0, 6).FontBold = True
    Grid1.Range(0, 1, 0, 6).Borders(cellEdgeBottom) = cellThick
    Grid1.SelectionMode = cellSelectionByRow
    Grid1.ExtendLastCol = True
    
    Grid1.Rows = 26
End Sub

Sub Lee_horas()
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
   
    Dim linea As Double
    Dim DESDE1 As String
    Dim HASTA1 As String
    Dim total_doc As Double
    Dim total_total As Double
    Dim total_descuento As Double
    Dim total_credito As Double
    Dim total_nulas As Double
    Dim K As Integer
           
        DESDE1 = dato4.text & "-" & dato3.text & "-" & dato2.text
        HASTA1 = dato7.text & "-" & dato6.text & "-" & dato5.text
        
          Set csql.ActiveConnection = ventasAuditoria
        
            csql.sql = "SELECT local,fecha,mid(horaventas,1,2),sum(total),tipo,numero,sum(descuento) as descuento,count(numero) as contados FROM sv_documento_cabeza_" + dato1.text + " "
            csql.sql = csql.sql + "WHERE local='" & dato1.text & "' and fecha between '" & DESDE1 & "' and '" & HASTA1 & "' and caja <'90' "
            If facturas.Value = True Then
            csql.sql = csql.sql & "and tipo='FV' "
            End If
            If boletas.Value = True Then
            csql.sql = csql.sql & "and tipo='BV' "
            End If
            If todas.Value = True Then
            csql.sql = csql.sql & "and (tipo='BV' or tipo='FV')"
            End If
            
            csql.sql = csql.sql + " group by MID(horaventas,1,2) "
            csql.sql = csql.sql + "order by horaventas,fecha asc "
            csql.Execute
       
       For K = 1 To 24
       Grid1.Cell(K, 1).text = Format(K, "00") + ":00"
       Next K
       
        Grid1.AutoRedraw = False
        linea = 0
        If csql.RowsAffected > 0 Then
            
            Set resultados = csql.OpenResultset
            While resultados.EOF = False
            On Error Resume Next
            linea = CDbl(Replace(resultados(2), ":", ""))
               If linea = 0 Then linea = 1
              
                
                Grid1.Cell(linea, 2).text = resultados("contados")
                Grid1.Cell(linea, 3).text = Format(resultados(3), "$ ###,###,##0")
                Grid1.Cell(linea, 4).text = Format(resultados("descuento"), "$ ###,###,##0")
                Grid1.Cell(linea, 5).text = leernulas(DESDE1, HASTA1, pivote.text)
                Grid1.Cell(linea, 6).text = Format(leercreditoventas(DESDE1, HASTA1, pivote.text), "$ ###,###,##0")
                
                total_doc = total_doc + resultados("contados")
                total_total = total_total + resultados(3)
                total_descuento = total_descuento + resultados("descuento")
                total_nulas = total_nulas + CDbl(Grid1.Cell(linea, 5).text)
                total_credito = total_credito + CDbl(Grid1.Cell(linea, 6).text)
               resultados.MoveNext
               
           Wend
           
            resultados.Close
            Set resultados = Nothing
        
        End If
        
        
        
        Grid1.AutoRedraw = True
        Grid1.Refresh
     
                     
        linea = 25
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).FontSize = 12
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).FontBold = True
       
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).BackColor = RGB(90, 158, 214)
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).ForeColor = RGB(0, 0, 0)
        Grid1.Cell(linea, 1).text = "TOTALES"
        Grid1.Cell(linea, 2).text = total_doc
        Grid1.Cell(linea, 3).text = Format(total_total, "$ ###,###,##0")
        Grid1.Cell(linea, 4).text = Format(total_descuento, "$ ###,###,##0")
        Grid1.Cell(linea, 5).text = Format(total_nulas, "###,###,##0")
        Grid1.Cell(linea, 6).text = Format(total_credito, "$ ###,###,##0")
        

End Sub
Function leercreditoventas(DESDE, HASTA, HORA) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tabla As String
Set csql.ActiveConnection = ventasAuditoria
tabla = "select sum(dp.monto)as total,replace(mid(horaventas,1,2),':','') as hora "
tabla = tabla & "from sv_documento_pagos_" & dato1.text & " as dp,sv_documento_cabeza_" & dato1.text & " as dc "
tabla = tabla & "where replace(mid(horaventas,1,2),':','')='" & HORA & "' and  dc.foliosii=dp.foliofiscal and dc.numero=dp.numero and  dp.tipo=dc.tipo and dp.fecha between '" & DESDE & "' and '" & HASTA & "'and dc.fecha between '" & DESDE1 & "' and '" & HASTA & "' and dp.tipopago='4' "
If boletas.Value = True Then
tabla = tabla & "and dc.tipo='BV' "
End If
If facturas.Value = True Then
tabla = tabla & "and dc.tipo='FV' "
End If
tabla = tabla & "group by dc.tipo,hora"
csql.sql = tabla
csql.Execute
leercreditoventas = "0"
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leercreditoventas = resultados("total")
End If
csql.Close
Set csql = Nothing
Set resultados = Nothing
End Function


Function leernulas(DESDE, HASTA, HORA) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tabla As String
Set csql.ActiveConnection = ventasAuditoria
tabla = "select sum(dc.total)as total,replace(mid(horaventas,1,2),':','') as hora,count(numero) as emitidas "
tabla = tabla & "from sv_documento_cabeza_" & dato1.text & " as dc "
tabla = tabla & "where replace(mid(horaventas,1,2),':','')='" & HORA & "' and dc.fecha between '" & DESDE & "' and '" & HASTA & "' and nula='S' "
If boletas.Value = True Then
tabla = tabla & "and dc.tipo='BV' "
End If
If facturas.Value = True Then
tabla = tabla & "and dc.tipo='FV' "
End If
tabla = tabla & "group by hora"
csql.sql = tabla
csql.Execute
leernulas = "0"
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leernulas = resultados("emitidas")
End If
csql.Close
Set csql = Nothing
Set resultados = Nothing
End Function


Sub CARGAGRILLA2()
    Grid1.Cols = 6
    Grid1.Column(2).Alignment = cellRightCenter
    Grid1.Column(3).Alignment = cellRightCenter
    Grid1.Column(4).Alignment = cellRightCenter
    Grid1.Column(5).Alignment = cellRightCenter
    
    
   
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 83
    Grid1.Column(2).Width = 100
    Grid1.Column(3).Width = 130
    Grid1.Column(4).Width = 120
    Grid1.Column(5).Width = 120
    
  
   
    
   
   
    Grid1.Cell(0, 1).text = "HORA"
    Grid1.Cell(0, 2).text = "FECHA"
    Grid1.Cell(0, 3).text = "VENTAS SALA"
    Grid1.Cell(0, 4).text = "VENTAS PATIO"
    Grid1.Cell(0, 5).text = "TOTAL"
    
   
   
    Grid1.Range(0, 1, 0, 5).Alignment = cellLeftGeneral
    Grid1.Range(0, 1, 0, 5).FontSize = 9
    Grid1.Range(0, 1, 0, 5).FontBold = True
    Grid1.Range(0, 1, 0, 5).Borders(cellEdgeBottom) = cellThick
    Grid1.SelectionMode = cellSelectionByRow
    Grid1.ExtendLastCol = True
    
    Grid1.Rows = 26
End Sub
Sub Lee_horasXsector()
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
   
    Dim linea As Double
    Dim DESDE1 As String
    Dim HASTA1 As String
    Dim total_doc As Double
    Dim total_total As Double
    Dim total_descuento As Double
    Dim total_credito As Double
    Dim total_nulas As Double
    Dim K As Integer
           
           
           total_doc = 0
           total_total = 0
           total_descuento = 0
           
        DESDE1 = dato4.text & "-" & dato3.text & "-" & dato2.text
        HASTA1 = dato7.text & "-" & dato6.text & "-" & dato5.text
        
          Set csql.ActiveConnection = gestionRubro
        
         csql.sql = "SELECT CONCAT(hora,':00') AS hora,fecha,SUM(IF(gon.sector= '001', total,0)) AS VentasSala, SUM(IF(gon.sector= '002', total,0)) AS VentasPatio, " & _
                    "SUM(total) As total FROM (SELECT dc.tipo   AS tipo, dc.fecha  AS fecha, dd.codigo AS codigo, dd.total  AS total, LPAD(HOUR(dc.horaventas),2,'0') AS hora " & _
                    "FROM eltit_ventas20.sv_documento_cabeza_20 AS dc INNER JOIN eltit_ventas20.sv_documento_detalle_20 AS dd ON (dd.tipo = dc.tipo AND dd.numero = dc.numero " & _
                    "AND dd.caja = dc.caja AND dd.fecha = dc.fecha) WHERE dc.fecha between '" & DESDE1 & "' and '" & HASTA1 & "' AND (dc.tipo = 'BV' OR dc.tipo = 'FV') ORDER BY LPAD(HOUR(dc.horaventas),2,'0')) AS ventas " & _
                    "INNER JOIN l_ubicacion AS ubi ON (ventas.codigo = ubi.codigo AND ubi.local = '20')INNER JOIN l_gondolas AS gon ON (ubi.gondola = gon.codigo) GROUP BY hora "
          csql.Execute
       
       For K = 1 To 24
       Grid1.Cell(K, 1).text = Format(K, "00") + ":00"
       Grid1.Cell(K, 2).text = ""
       Grid1.Cell(K, 3).text = ""
       Grid1.Cell(K, 4).text = ""
       Grid1.Cell(K, 5).text = ""
       Next K
       
        Grid1.AutoRedraw = False
        linea = 0
        If csql.RowsAffected > 0 Then
            
            Set resultados = csql.OpenResultset
            While resultados.EOF = False
            'On Error Resume Next
            linea = CDbl(Mid(resultados(0), 1, 2))
               If linea = 0 Then linea = 1
                Grid1.Cell(linea, 2).text = resultados(1)
                Grid1.Cell(linea, 3).text = Format(resultados(2), "$ ###,###,##0")
                Grid1.Cell(linea, 4).text = Format(resultados(3), "$ ###,###,##0")
                Grid1.Cell(linea, 5).text = Format(resultados(4), "$ ###,###,##0")
                
                total_doc = total_doc + resultados(2)
                total_total = total_total + resultados(3)
                total_descuento = total_descuento + resultados(4)
                
                
                resultados.MoveNext
               
           Wend
           
            resultados.Close
            Set resultados = Nothing
        
        End If
        
        
        
        Grid1.AutoRedraw = True
        Grid1.Refresh
     
                     
        linea = 25
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).FontSize = 12
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).FontBold = True
       
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).BackColor = RGB(90, 158, 214)
        Grid1.Range(linea, 1, linea, Grid1.Cols - 1).ForeColor = RGB(0, 0, 0)
        Grid1.Cell(linea, 1).text = "TOTALES"
        Grid1.Cell(linea, 2).text = ""
        Grid1.Cell(linea, 3).text = Format(total_doc, "$ ###,###,##0")
        Grid1.Cell(linea, 4).text = Format(total_total, "$ ###,###,##0")
        Grid1.Cell(linea, 5).text = Format(total_descuento, "$ ###,###,##0")
'        Grid1.Cell(linea, 6).text = ""
        
        

End Sub

