VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ListadoRetiros 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Retiros de Dinero por Caja y Cajera"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   13695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar progreso 
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   8880
      Visible         =   0   'False
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdimprimir 
      BackColor       =   &H00FF8080&
      Caption         =   "I M P R I M I R"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9360
      Width           =   3735
   End
   Begin XPFrame.FrameXp FrameXp5 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   12726
      BackColor       =   16744576
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin FlexCell.Grid impresion 
         Height          =   6975
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   12303
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   -1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
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
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   3413
      BackColor       =   16744576
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   975
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1720
         BackColor       =   16744576
         Caption         =   "Rango de Fechas a Procesar"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         BordeEstilo     =   2
         Alignment       =   1
         Begin VB.TextBox dato1 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   120
            MaxLength       =   2
            TabIndex        =   15
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
         Begin VB.TextBox dato2 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   570
            MaxLength       =   2
            TabIndex        =   14
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
         Begin VB.TextBox dato3 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   1020
            MaxLength       =   4
            TabIndex        =   13
            Tag             =   "proveedor"
            Top             =   540
            Width           =   705
         End
         Begin VB.TextBox dato6 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   3210
            MaxLength       =   4
            TabIndex        =   12
            Tag             =   "proveedor"
            Top             =   540
            Width           =   825
         End
         Begin VB.TextBox dato4 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   11
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
         Begin VB.TextBox dato5 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   2760
            MaxLength       =   2
            TabIndex        =   10
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
         Begin VB.Label lbl3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Hasta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2280
            TabIndex        =   17
            Top             =   240
            Width           =   1725
         End
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Desde"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   270
            Width           =   1605
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "Seleccion del Local a Procresar"
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
         Begin VB.ComboBox combolocal 
            Height          =   315
            Left            =   0
            TabIndex        =   8
            Top             =   240
            Width           =   4485
         End
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         BackColor       =   16744576
         Caption         =   "Detallar informe Por:"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin VB.OptionButton cajas 
            BackColor       =   &H00FF8080&
            Caption         =   "Agrupado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1665
         End
         Begin VB.OptionButton cajeras 
            BackColor       =   &H00FF8080&
            Caption         =   "Detallado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   5
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Genera Informes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1545
      End
   End
End
Attribute VB_Name = "ListadoRetiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private TIPO As String
    Private detalle As Boolean
    Private fecha1 As String
    Private fecha2 As String
    Private codigoempresa As String
    Private codigovendedor As String

Private Sub cajas_Click()
'If combolocal.text <> "" Then Call Command1_Click
End Sub

Private Sub cajeras_Click()
'If combolocal.text <> "" Then Call Command1_Click
End Sub

Private Sub cmdImprimir_Click()
If impresion.Rows > 1 Then Call imprimir
End Sub

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
progreso.Visible = True
codigoempresa = Mid(combolocal.text, 1, 2)
    
 If cajeras.Value = True Then Call CargaGrillaInforme(1, 7)
 If Cajas.Value = True Then Call CargaGrillaInforme(1, 8)
 
    fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
    fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
    Call generaInformeLV(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
    Call RetirosCaja(fecha1, fecha2)
Screen.MousePointer = vbNormal
progreso.Visible = False
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

    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
           dato2.SetFocus
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
           dato3.SetFocus
        End If
    End Sub
        
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
           dato4.SetFocus
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            dato5.SetFocus
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            dato6.SetFocus
        End If
    End Sub
        
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
        Command1.SetFocus
        End If
    End Sub
   
    Private Sub dato1_LostFocus()
    Call limpiaBarra(2)
    Call esfecha(dato1, dato2, dato3, "dd")
    End Sub
    Private Sub dato2_LostFocus()
    Call esfecha(dato1, dato2, dato3, "mm")
    End Sub
    Private Sub dato3_LostFocus()
    Call esfecha(dato1, dato2, dato3, "yyyy")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato4, dato5, dato6, "dd")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato4, dato5, dato6, "mm")
    End Sub
    Private Sub dato6_LostFocus()
    Call esfecha(dato4, dato5, dato6, "yyyy")
    End Sub

    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 27
                Unload Me
            Case 38
                If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                    Unload Me
                End If
        End Select
    End Sub
    
    Private Sub Form_Load()
        Call Centrar(Me)
       
        
        TIPO = "(dc.tipo = 'FV')"
        detalle = False
        dato1.text = Format(fechasistema, "dd")
        dato2.text = Format(fechasistema, "mm")
        dato3.text = Format(fechasistema, "yyyy")
        dato4.text = Format(fechasistema, "dd")
        dato5.text = Format(fechasistema, "mm")
        dato6.text = Format(fechasistema, "yyyy")
    LEErlocales
 Command1_Click
End Sub
 
    Private Sub imprimir()
        Dim i As Long
        
        impresion.AutoRedraw = False
        If cajeras.Value = True Then
        impresion.Cols = 7
            impresion.Range(1, 1, 1, 6).Borders(cellEdgeTop) = cellThick
        Else
       impresion.Cols = 8
                impresion.Range(1, 1, 1, 7).Borders(cellEdgeTop) = cellThick
        End If
        impresion.PageSetup.HeaderMargin = 2
        impresion.PageSetup.TopMargin = 1
        impresion.PageSetup.LeftMargin = 0.5
        impresion.PageSetup.RightMargin = 0
        impresion.PageSetup.BottomMargin = 1
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.Orientation = cellPortrait
       'impresion.Cols = 8
                
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeTop) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeRight) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellInsideHorizontal) = cellThick
impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellInsideVertical) = cellThick

        impresion.PageSetup.PrintFixedRow = True
        Call verificaImpresora(5, impresion)
        impresion.AutoRedraw = True
    End Sub

Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas "
        ' original cSql.sql = cSql.sql + "ORDER BY codigo "
        ' ariel agrega condicion local < 50 para que no liste locales 50 y 51
        csql.sql = csql.sql + "  WHERE CODIGO < '50' ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                combolocal.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
           combolocal.AddItem ("99" + "  TODOS LOS LOCALES")
                
        combolocal.text = combolocal.List(CDbl(empresaActiva))
        End If
        
End Sub
Public Sub generaInformeLV(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Dim documento As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    If Cajas.Value = True Then
    Call cargaCabeza("LISTADO DE RETIROS DE DINERO X CAJA DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), Mid(combolocal, 1, 2), impresion)
   End If
   If cajeras.Value = True Then
   Call cargaCabeza("LISTADO DE RETIROS DE DINERO X CAJERAS DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), Mid(combolocal, 1, 2), impresion)
    End If
   
Rem    Call resumenVentas(data, impresion, "", Mid(combolocal, 1, 3), fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatogrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
If cajeras.Value = True Then
    formatogrilla(1, 1) = "FECHA"
    formatogrilla(1, 2) = "CAJA"
    formatogrilla(1, 3) = "RUT CAJERA"
    formatogrilla(1, 4) = "NOMBRE CAJERA"
    formatogrilla(1, 5) = "MONTO"
    formatogrilla(1, 6) = "RETIRADO POR"
    
    Rem ANCHO DE LAS CELDAS
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "5"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "20"
        formatogrilla(8, 5) = "10"
        formatogrilla(8, 6) = "20"
    
    Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "15"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "15"
        formatogrilla(2, 4) = "15"
        formatogrilla(2, 5) = "9"
        formatogrilla(2, 6) = "20"
        formatogrilla(2, 7) = "9"
      
      
        formatogrilla(4, 5) = "$ ###,###,##0"
        formatogrilla(4, 6) = "$ ###,###,##0"
End If
    
If Cajas.Value = True Then
    formatogrilla(1, 1) = "FECHA"
    formatogrilla(1, 2) = "HORA"
    formatogrilla(1, 3) = "CAJA"
    formatogrilla(1, 4) = "RUT CAJERA"
    formatogrilla(1, 5) = "NOMBRE CAJERA"
    formatogrilla(1, 6) = "CANTIDAD RETIROS"
    formatogrilla(1, 7) = "TOTAL RETIROS"
     
       
    Rem ANCHO DE LAS CELDAS
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "5"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "10"
        formatogrilla(8, 5) = "15"
        formatogrilla(8, 6) = "10"
        formatogrilla(8, 7) = "10"
        
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = "$ ###,###,##0"
            
End If
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "C"
        formatogrilla(3, 2) = "C"
        formatogrilla(3, 3) = "C"
        formatogrilla(3, 4) = "C"
        formatogrilla(3, 5) = "C"
        formatogrilla(3, 6) = "C"
        formatogrilla(3, 7) = "C"
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        
        
        
        Rem LOCCKED

        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        Rem ANCHO

        
                
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        
        impresion.Column(0).Width = 0
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1.75
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            impresion.Column(i).Alignment = cellRightCenter
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub

Sub RetirosCaja(fecha1, fecha2)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
Set csql.ActiveConnection = ventasRubro
    csql.sql = "select "
If Cajas.Value = True Then csql.sql = csql.sql & "fecha,hora,caja,cajera,count(cajera),SUM(monto) "
If cajeras.Value = True Then csql.sql = csql.sql & "fecha,caja,cajera,monto,rutretiradopor "
    csql.sql = csql.sql & "from " & cliente_sql & "ventas" & empresaActiva & ".sv_retirosdecaja_" & empresaActiva
    csql.sql = csql.sql & "  where fecha between '" & fecha1 & "' and '" & fecha2 & "'"
If Cajas.Value = True Then csql.sql = csql.sql & " group by cajera,caja order by fecha,caja asc"
If cajeras.Value = True Then csql.sql = csql.sql & " order by fecha,caja asc"


csql.Execute
        
If csql.RowsAffected > 0 Then
progreso.Max = csql.RowsAffected
   Set resultados = csql.OpenResultset
   Dim r As Long
   While Not resultados.EOF
   With impresion
   .Rows = .Rows + 1
   
   r = .Rows - 1
   progreso.Value = r
    .Cell(r, 1).text = resultados(0)
    .Cell(r, 2).text = resultados(1)
    .Cell(r, 3).text = resultados(2)
If cajeras.Value = True Then
    .Cell(r, 3).text = Mid(resultados(2), 1, 9) & "-" & Mid(resultados(2), 10, 11)
    .Cell(r, 4).text = leerNombreCajera(resultados(2))
    
    .Cell(r, 5).text = resultados(3)
    .Cell(r, 6).text = leerNombreCajera(resultados(4))
Else
    .Cell(r, 4).text = Mid(resultados(3), 1, 9) & "-" & Mid(resultados(3), 10, 11)
    .Cell(r, 5).text = leerNombreCajera(resultados(3))
    .Cell(r, 6).text = resultados(4)
    .Cell(r, 7).text = resultados(5)
   
End If
    
    End With
   
resultados.MoveNext
Wend
resultados.Close
End If
End Sub

