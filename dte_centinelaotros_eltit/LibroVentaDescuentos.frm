VERSION 5.00
Begin VB.Form descuentos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadisticas de Ventas por Vendedores"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   14565
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox data 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   30
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox FrameXp2 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7140
      Left            =   90
      ScaleHeight     =   7080
      ScaleWidth      =   14400
      TabIndex        =   1
      Top             =   2100
      Width           =   14460
      Begin VB.PictureBox impresion 
         Height          =   6780
         Left            =   45
         ScaleHeight     =   6720
         ScaleWidth      =   14280
         TabIndex        =   2
         Top             =   360
         Width           =   14340
      End
   End
   Begin VB.PictureBox FrameXp1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   45
      ScaleHeight     =   1890
      ScaleWidth      =   14370
      TabIndex        =   0
      Top             =   90
      Width           =   14430
      Begin VB.PictureBox fechas 
         BackColor       =   &H00FF0000&
         Height          =   960
         Left            =   4770
         ScaleHeight     =   900
         ScaleWidth      =   3600
         TabIndex        =   5
         Top             =   360
         Width           =   3660
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
            Height          =   315
            Left            =   180
            MaxLength       =   2
            TabIndex        =   11
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
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
            Height          =   315
            Left            =   630
            MaxLength       =   2
            TabIndex        =   10
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
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
            Height          =   315
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   9
            Tag             =   "proveedor"
            Top             =   540
            Width           =   705
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
            Height          =   315
            Left            =   2790
            MaxLength       =   4
            TabIndex        =   8
            Tag             =   "proveedor"
            Top             =   540
            Width           =   705
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
            Height          =   315
            Left            =   1890
            MaxLength       =   2
            TabIndex        =   7
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
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
            Height          =   315
            Left            =   2340
            MaxLength       =   2
            TabIndex        =   6
            Tag             =   "proveedor"
            Top             =   540
            Width           =   435
         End
         Begin VB.Label lbl3 
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
            Left            =   1890
            TabIndex        =   13
            Top             =   270
            Width           =   1605
         End
         Begin VB.Label lbl2 
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
            Left            =   180
            TabIndex        =   12
            Top             =   270
            Width           =   1605
         End
      End
      Begin VB.PictureBox FrameXp4 
         BackColor       =   &H00FFC0C0&
         Height          =   1230
         Left            =   8505
         ScaleHeight     =   1170
         ScaleWidth      =   1530
         TabIndex        =   4
         Top             =   360
         Width           =   1590
         Begin VB.OptionButton TIPO2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "x Articulos"
            Height          =   375
            Left            =   135
            TabIndex        =   19
            Top             =   675
            Width           =   1635
         End
         Begin VB.OptionButton TIPO1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "x Documentos"
            Height          =   330
            Left            =   135
            TabIndex        =   18
            Top             =   270
            Width           =   1365
         End
      End
      Begin VB.PictureBox FrameXp6 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H0000FFFF&
         Height          =   660
         Left            =   45
         ScaleHeight     =   600
         ScaleWidth      =   4590
         TabIndex        =   14
         Top             =   1200
         Width           =   4650
         Begin VB.ComboBox COMBOVENDEDOR 
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   270
            Width           =   4485
         End
      End
      Begin VB.PictureBox FrameXp7 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H0000FFFF&
         Height          =   675
         Left            =   45
         ScaleHeight     =   615
         ScaleWidth      =   4590
         TabIndex        =   16
         Top             =   360
         Width           =   4650
         Begin VB.ComboBox combolocal 
            Height          =   315
            Left            =   45
            TabIndex        =   17
            Top             =   270
            Width           =   4485
         End
      End
      Begin VB.PictureBox FrameXp3 
         BackColor       =   &H00FFC0C0&
         Height          =   1230
         Left            =   10170
         ScaleHeight     =   1170
         ScaleWidth      =   1440
         TabIndex        =   20
         Top             =   360
         Width           =   1500
         Begin VB.OptionButton VISTA1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Detallada"
            Height          =   330
            Left            =   90
            TabIndex        =   22
            Top             =   270
            Width           =   1950
         End
         Begin VB.OptionButton VISTA2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Acumulada"
            Height          =   375
            Left            =   90
            TabIndex        =   21
            Top             =   675
            Width           =   2175
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
         Height          =   285
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1620
         Width           =   2535
      End
      Begin VB.PictureBox FrameXp5 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H0000FFFF&
         Height          =   660
         Left            =   4750
         ScaleHeight     =   600
         ScaleWidth      =   3630
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   3690
         Begin VB.TextBox CODIGO 
            Alignment       =   1  'Right Justify
            Height          =   365
            Left            =   720
            MaxLength       =   13
            TabIndex        =   25
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.PictureBox FrameXp8 
         BackColor       =   &H00FFC0C0&
         Height          =   1230
         Left            =   11790
         ScaleHeight     =   1170
         ScaleWidth      =   2025
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   2085
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Atenciones"
            Height          =   375
            Left            =   90
            TabIndex        =   29
            Top             =   675
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Venta"
            Height          =   330
            Left            =   90
            TabIndex        =   28
            Top             =   270
            Width           =   1950
         End
      End
      Begin VB.Label lblayuda 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 AYUDA PRODUCTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.PictureBox frmImprimir 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5715
      ScaleHeight     =   315
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   9315
      Width           =   3375
   End
End
Attribute VB_Name = "descuentos"
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
    



Private Sub CODIGO_GotFocus()
lblayuda.Visible = True

End Sub

Private Sub CODIGO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
 Call ayudaProductotxt(CODIGO)
 End If
 
End Sub

Private Sub CODIGO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
CODIGO.Text = ceros(CODIGO)
Command1.SetFocus
End If

End Sub

Private Sub CODIGO_LostFocus()
lblayuda.Visible = False

End Sub

Private Sub combolocal_Click()

Call conecntarVentasAuditoria(servidor, baseVentas & Mid(combolocal.Text, 1, 2), usuario, password)

LEErVENDEDORES
End Sub



Private Sub COMBOVENDEDOR_Click()
Command1_Click
End Sub

Private Sub Command1_Click()
codigoempresa = Mid(combolocal.Text, 1, 2)
codigovendedor = Mid(COMBOVENDEDOR.Text, 1, 10)

            If TIPO1.Value = True Then
            Call CargaGrillaInforme(1, 11)
            'Call CargaGrillaInformeventasxvendedor(1, 7)
            fecha1 = dato3.Text & "-" & dato2.Text & "-" & dato1.Text
            fecha2 = dato6.Text & "-" & dato5.Text & "-" & dato4.Text
            Call generaInformeLV(data, impresion, TIPO, detalle, dato1.Text, fecha1, fecha2)
            End If

            If TIPO2.Value = True Then
            Call CargaGrillaInforme2(1, 8)
            
            fecha1 = dato3.Text & "-" & dato2.Text & "-" & dato1.Text
            fecha2 = dato6.Text & "-" & dato5.Text & "-" & dato4.Text
            Call generaInformevp(data, impresion, TIPO, detalle, dato1.Text, fecha1, fecha2)
           End If
           
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
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
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
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
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.Text = ceros(dato1)
            If dato1.Text = "00" Then
                dato1.Text = Format(fechasistema, "dd")
            End If
           dato2.SetFocus
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.Text = ceros(dato2)
            If dato2.Text = "00" Then
                dato2.Text = Format(fechasistema, "mm")
            End If
           dato3.SetFocus
        End If
    End Sub
        
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.Text = ceros(dato3)
            If dato3.Text = "0000" Then
                dato3.Text = Format(fechasistema, "yyyy")
            End If
           dato4.SetFocus
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.Text = ceros(dato4)
            If dato4.Text = "00" Then
                dato4.Text = Format(fechasistema, "dd")
            End If
            dato5.SetFocus
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.Text = ceros(dato5)
            If dato5.Text = "00" Then
                dato5.Text = Format(fechasistema, "mm")
            End If
            dato6.SetFocus
        End If
    End Sub
        
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.Text = ceros(dato6)
            If dato6.Text = "0000" Then
                dato6.Text = Format(fechasistema, "yyyy")
            End If
        Command1.SetFocus
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
'    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato1.text) = dato1.MaxLength Then
'            Call dato1_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato2.text) = dato2.MaxLength Then
'            Call dato2_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato3.text) = dato3.MaxLength Then
'            Call dato3_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato4_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato4.text) = dato4.MaxLength Then
'            Call dato4_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato5_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato5.text) = dato5.MaxLength Then
'            Call dato5_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato6_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato6.text) = dato6.MaxLength Then
'            Call dato6_KeyPress(13)
'        End If
'    End Sub
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
 
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
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).Text = UCase(Me.Caption)
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
        dato1.Text = Format(fechasistema, "dd")
        dato2.Text = Format(fechasistema, "mm")
        dato3.Text = Format(fechasistema, "yyyy")
        dato4.Text = Format(fechasistema, "dd")
        dato5.Text = Format(fechasistema, "mm")
        dato6.Text = Format(fechasistema, "yyyy")
    LEErlocales
    Call conecntarVentasAuditoria(servidor, baseVentas & Mid(combolocal.Text, 1, 2), usuario, password)

    LEErVENDEDORES
    TIPO1.Value = True
    VISTA1.Value = True
    Call CargaGrillaInforme(1, 11)
    'Call CargaGrillaInformeventasxvendedor(1, 7)
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "TIPO"
        formatoGrilla(1, 2) = "NUMERO"
        formatoGrilla(1, 3) = "FECHA"
        formatoGrilla(1, 4) = "RUT"
        formatoGrilla(1, 5) = "CLIENTE"
        formatoGrilla(1, 6) = "TOTAL"
        formatoGrilla(1, 7) = "DESC"
        formatoGrilla(1, 8) = "Nº DOC"
        formatoGrilla(1, 9) = "CANT PRODUCTOS"
        formatoGrilla(1, 10) = "% PARTICIPACION"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "4"
        formatoGrilla(2, 2) = "10"
        formatoGrilla(2, 3) = "8"
        formatoGrilla(2, 4) = "10"
        formatoGrilla(2, 5) = "25"
        formatoGrilla(2, 6) = "9"
        formatoGrilla(2, 7) = "9"
        formatoGrilla(2, 8) = "6"
        formatoGrilla(2, 9) = "9"
        formatoGrilla(2, 10) = "9"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "C"
        formatoGrilla(3, 3) = "D"
        formatoGrilla(3, 4) = "S"
        formatoGrilla(3, 5) = "S"
        formatoGrilla(3, 6) = "N"
        formatoGrilla(3, 7) = "N"
        formatoGrilla(3, 8) = "N"
        formatoGrilla(3, 9) = "N"
         formatoGrilla(3, 10) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = ""
        formatoGrilla(4, 5) = ""
        formatoGrilla(4, 6) = "###,###,##0"
        formatoGrilla(4, 7) = "% ###,##0.000"
        formatoGrilla(4, 8) = "###,###,##0"
        formatoGrilla(4, 9) = "###,###,##0"
        formatoGrilla(4, 10) = "% #,###,##0.00"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        formatoGrilla(5, 7) = "FALSE"
        formatoGrilla(5, 8) = "FALSE"
        formatoGrilla(5, 9) = "FALSE"
         formatoGrilla(5, 10) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        formatoGrilla(6, 8) = ""
        formatoGrilla(6, 9) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        Rem ANCHO
        formatoGrilla(8, 1) = "3"
        formatoGrilla(8, 2) = "8"
        formatoGrilla(8, 3) = "8"
        formatoGrilla(8, 4) = "10"
        formatoGrilla(8, 5) = "30"
        formatoGrilla(8, 6) = "8"
        formatoGrilla(8, 7) = "8"
        formatoGrilla(8, 8) = "6"
        formatoGrilla(8, 9) = "8"
        formatoGrilla(8, 10) = "8"
        
                
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
            impresion.Cell(0, i).Text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
'formato grilla ventas por vendedor
Private Sub CargaGrillaInformeventasxvendedor(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "VENDEDOR"
        formatoGrilla(1, 3) = "NETO"
        formatoGrilla(1, 4) = "I.V.A"
        formatoGrilla(1, 5) = "EXENTO"
        formatoGrilla(1, 6) = "TOTAL"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "5"
        formatoGrilla(2, 2) = "20"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "9"
        formatoGrilla(2, 5) = "9"
        formatoGrilla(2, 6) = "9"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = "$ ###,###,##0"
        formatoGrilla(4, 4) = "$ ###,###,##0"
        formatoGrilla(4, 5) = "$ ###,###,##0"
        formatoGrilla(4, 6) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        Rem ANCHO
        formatoGrilla(8, 1) = "5"
        formatoGrilla(8, 2) = "20"
        formatoGrilla(8, 3) = "8"
        formatoGrilla(8, 4) = "8"
        formatoGrilla(8, 5) = "8"
        formatoGrilla(8, 6) = "8"
        
                
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
            impresion.Cell(0, i).Text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub
    
    'fin configura grilla ventas x vendedor

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        If impresion.Rows > 1 Then
        Call imprimir
        End If
        
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        Call Titulos("Ventas x Vendedores")
        impresion.AutoRedraw = False
        impresion.Range(1, 1, 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThick
        impresion.PageSetup.HeaderMargin = 2
        impresion.PageSetup.TopMargin = 3
        impresion.PageSetup.LeftMargin = 0.5
        impresion.PageSetup.RightMargin = 0.5
        impresion.PageSetup.BottomMargin = 3
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.Orientation = cellLandscape
        
        impresion.PageSetup.PrintFixedRow = True
        Call verificaImpresora(5, impresion)
        impresion.AutoRedraw = True
    End Sub
    
    
Sub LEErVENDEDORES()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = ventas
        COMBOVENDEDOR.Clear
        cSql.sql = "SELECT rut,nombre "
        cSql.sql = cSql.sql + "FROM sv_maestrovendedores "
        cSql.sql = cSql.sql + "where local='" & Mid(combolocal.Text, 1, 2) & "' ORDER BY rut "
        cSql.Execute
        COMBOVENDEDOR.AddItem ("99" + "  TODOS LOS VENDEDORES")
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                COMBOVENDEDOR.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            
            COMBOVENDEDOR.Text = COMBOVENDEDOR.List(0)
            
        End If
        
End Sub

Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = gestion
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM g_maestroempresas "
        cSql.sql = cSql.sql + "ORDER BY codigo "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                combolocal.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
           combolocal.AddItem ("99" + "  TODOS LOS LOCALES")
                
        combolocal.Text = combolocal.List(CDbl(empresaActiva))
        End If
        
End Sub

Sub LEErventas()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM g_maestroempresas "
        cSql.sql = cSql.sql + "ORDER BY codigo "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                combolocal.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
           combolocal.AddItem ("99" + "  TODOS LOS LOCALES")
                
        combolocal.Text = combolocal.List(CDbl(empresaActiva))
        End If
        
End Sub

Public Sub generaInformeLV(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Dim documento As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("LISTADO VENTAS X VENDEDORES DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
    Call resumenVentas(data, impresion, TIPO, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Function resumenVentas(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim FILTRO As String
    
    Dim tabla As String
    Dim rubAux As String
    Dim harinas As Double
    Dim subproductos As Double
    Dim envases As Double
    Dim trigo As Double
    Dim maquila As Double
    Dim otros As Double
    Dim cadena As String
    Dim tipoDoc As String
    Dim numeroDoc As String
    Dim cSql As New rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim i As Integer
    Dim totales(10) As Double
    Dim totales2(10) As Double
    Dim conta As Double
    Dim totalventa As Double
    Dim cSql2 As New rdoQuery
    Dim resultados As rdoResultset
        Set cSql.ActiveConnection = ventasAuditoria
    Rem calcula total venta
        cSql.sql = "SELECT sum(dd.total) "
        cSql.sql = cSql.sql + "FROM sv_documento_detalle_" + codigoempresa + " as dd "
        cSql.sql = cSql.sql + "where dd.fecha BETWEEN '" + fecha1 + "' AND '" + fecha2 + "' "
        cSql.sql = cSql.sql + "AND ( dd.tipo<>'PV' AND dd.tipo<>'NP' AND dd.tipo<>'CO' and dd.tipo<>'NB' and dd.tipo<>'NF') and caja<'90' "
        cSql.Execute
        
     If cSql.RowsAffected > 0 Then
     Set resultado = cSql.OpenResultset
    If Not IsNull(resultado(0)) Then
    totalventa = resultado(0)
    End If
    End If
    
    Set cSql2.ActiveConnection = ventasAuditoria
    rubAux = rubro
    tabla = "SELECT dd.tipo, dd.numero , dd.fecha, dd.rut, mc.nombre,sum(dd.total), sum(dd.descuento), sum(dd.cantidad), dd.vendedor "
    tabla = tabla & "FROM sv_documento_detalle_" & codigoempresa & " as dd INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON dd.rut = mc.rut AND mc.sucursal = '0'"
    tabla = tabla & "WHERE dd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and (dd.tipo='FV' or dd.tipo='BV') AND dd.local='" + codigoempresa + "' and caja<'90' "
    
    If codigovendedor <> "99  TODOS " Then
    tabla = tabla & "and dd.vendedor='" + codigovendedor + "' "
    End If
   
    
    tabla = tabla & "group by dd.numero,dd.vendedor ORDER BY dd.vendedor,mc.nombre,dd.numero "
    cSql2.sql = tabla
    cSql2.Execute
    
    'Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    
    If cSql2.RowsAffected > 0 Then
    Set resultados = cSql2.OpenResultset
    
       FILTRO = resultados("vendedor")
       impresion.Rows = 2
        While Not resultados.EOF
           If FILTRO <> resultados("vendedor") Then
           linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           impresion.Range(linea, 4, linea, 5).Merge
           impresion.Cell(linea, 4).Text = leerNombreVendedor(FILTRO)
           impresion.Cell(linea, 6).Text = totales(1)
           impresion.Cell(linea, 7).Text = totales(2) / totales(3)
           impresion.Cell(linea, 8).Text = totales(3)
           impresion.Cell(linea, 9).Text = totales(4)
           impresion.Cell(linea, 10).Text = ((totales(1) / totalventa) * 100)
        For i = 1 To 4
        totales(i) = 0
        Next i
        FILTRO = resultados("vendedor")
           End If
           If VISTA1.Value = True Then
           
            linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            impresion.Cell(linea, 1).Text = resultados(0)
            impresion.Cell(linea, 2).Text = resultados(1)
            impresion.Cell(linea, 3).Text = resultados(2)
            impresion.Cell(linea, 4).Text = resultados(3)
            impresion.Cell(linea, 5).Text = resultados(4)
            impresion.Cell(linea, 6).Text = resultados(5)
            impresion.Cell(linea, 7).Text = resultados(6)
            impresion.Cell(linea, 8).Text = 1
            impresion.Cell(linea, 9).Text = resultados(7)
            impresion.Cell(linea, 10).Text = ((resultados(5) / totalventa) * 100)
           End If
           
            conta = 1
            totales(1) = totales(1) + CDbl(resultados(5))
            totales(2) = totales(2) + CDbl(resultados(6))
            totales(3) = totales(3) + conta
            totales(4) = totales(4) + CDbl(resultados(7))
            
            totales2(1) = totales2(1) + CDbl(resultados(5))
            totales2(2) = totales2(2) + CDbl(resultados(6))
            totales2(3) = totales2(3) + conta
            totales2(4) = totales2(4) + CDbl(resultados(7))
            
            resultados.MoveNext
        Wend
        cSql2.Close
        Set cSql2 = Nothing
        Set resultados = Nothing
        
    ' total vendedor
    
    linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 4, linea, 5).Merge
           impresion.Cell(linea, 4).Text = leerNombreVendedor(FILTRO)
            
           impresion.Cell(linea, 6).Text = totales(1)
           impresion.Cell(linea, 7).Text = totales(2) / totales(3)
           impresion.Cell(linea, 8).Text = totales(3)
           impresion.Cell(linea, 9).Text = totales(4)
           impresion.Cell(linea, 10).Text = ((totales(1) / totalventa) * 100)
            
        For i = 1 To 4
        totales(i) = 0
        Next i
        
    ' total todos
    
    
    linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            
            
           
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 9).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 4, linea, 5).Merge
           
           
           
           impresion.Cell(linea, 4).Text = "TOTAL GENERAL VENTAS"
            
            
            impresion.Cell(linea, 6).Text = totales2(1)
            impresion.Cell(linea, 7).Text = totales2(2) / totales2(3)
            impresion.Cell(linea, 8).Text = totales2(3)
            impresion.Cell(linea, 9).Text = totales2(4)
        
    End If

    'Call sumaGrilla(impresion)
End Function



Private Sub TIPO1_Click()
FrameXp5.Visible = False
lblayuda.Visible = False
Command1_Click
End Sub

Private Sub TIPO2_Click()
FrameXp5.Visible = True
lblayuda.Visible = True

CODIGO.SetFocus
'Command1_Click
End Sub

Private Sub TIPO2_LostFocus()
lblayuda.Visible = False

End Sub

Private Sub VISTA1_Click()
Command1_Click

End Sub

Private Sub VISTA2_Click()
Command1_Click
End Sub
Private Function resumenVentasproductos(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String) As Long
    Dim FILTRO As String
    
    Dim tabla As String
    Dim rubAux As String
    Dim harinas As Double
    Dim subproductos As Double
    Dim envases As Double
    Dim trigo As Double
    Dim maquila As Double
    Dim otros As Double
    Dim cadena As String
    Dim tipoDoc As String
    Dim numeroDoc As String
    Dim cSql As rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim i As Integer
    Dim totales(10) As Double
    Dim totales2(10) As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim T3 As Double
    Dim porce As Double
    Dim cSql2 As New rdoQuery
    Dim resultados As rdoResultset
    
    Set cSql2.ActiveConnection = ventasAuditoria
    
    
    rubAux = rubro
 Rem IF(dd.tipo = 'BV' OR dd.tipo = 'ZE', ROUND(SUM(dd.total / " & Replace((1 + iva / 100), ",", ".") & "),0)
   ' ROUND(SUM(dd.total - dd.total * dd.descuento / 100),0)
    tabla = "SELECT dv.codigo, dv.descripcion, sum(dv.cantidad) as cantidad, IF (dv.tipo = 'FV' ,ROUND(SUM((dv.total - (dv.total * dv.descuento2 / 100)) ),0),ROUND(SUM(dv.total - (dv.total * dv.descuento2 / 100)/1.19),0)), dv.vendedor, "
    tabla = tabla + "IF (dv.tipo = 'FV' ,ROUND(SUM(((dv.total * dv.descuento2 / 100)) ),0),ROUND(SUM((dv.total * dv.descuento2 / 100)/1.19),0)) "
    tabla = tabla & "FROM sv_documento_detalle_" & codigoempresa & " AS dv "
    If codigoempresa = "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO'  "
    End If
    If codigoempresa <> "99" Then
    tabla = tabla & "WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and tipo<>'PV' AND TIPO<>'NP' AND TIPO<>'CO' AND dv.local='" + codigoempresa + "' "
    End If
    If Mid(codigovendedor, 1, 2) <> "99" Then
    tabla = tabla & "and dv.vendedor='" + Mid(codigovendedor, 1, 2) + "' and dv.caja < '90' "
    End If
    If CODIGO.Text <> "" Then
    tabla = tabla & "and codigo='" & CODIGO.Text & "'  "
    End If
    tabla = tabla & "GROUP BY dv.vendedor,dv.codigo "
    tabla = tabla & "ORDER BY dv.vendedor,total desc "
    cSql2.sql = tabla
    cSql2.Execute
    
   ' Call ConectarControlData(data, servidor, baseVentas & rubAux, usuario, password, tabla)
    
    For i = 0 To 10
        totales(i) = 0
        totales2(i) = 0
    Next i
    linea = 0
    
    If cSql2.RowsAffected > 0 Then
    Set resultados = cSql2.OpenResultset
    
       FILTRO = resultados("vendedor")
       impresion.Rows = 2
        While Not resultados.EOF
           If FILTRO <> resultados("vendedor") Then
           linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).Text = leerNombreVendedor(FILTRO)
            porce = 0
            If totales(1) <> 0 Then
            porce = totales(2) / (totales(3) + totales(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).Text = totales(1)
            impresion.Cell(linea, 4).Text = totales(2)
            impresion.Cell(linea, 5).Text = porce
            impresion.Cell(linea, 6).Text = totales(3)
        
            For i = 1 To 4
            totales(i) = 0
            Next i
        FILTRO = resultados("vendedor")
           End If
           If VISTA1.Value = True Then
           
           linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            impresion.Cell(linea, 1).Text = resultados(0)
            impresion.Cell(linea, 2).Text = resultados(1)
            t1 = resultados(2)
            t2 = resultados(3)
            
            If t1 = 0 Then t1 = 1
            porce = 0
            If resultados(5) <> 0 Then
            porce = resultados(5) / (t2 + resultados(5)) * 100
            
            End If
            
            impresion.Cell(linea, 3).Text = resultados(2)
            impresion.Cell(linea, 4).Text = resultados(5)
            impresion.Cell(linea, 5).Text = porce
            impresion.Cell(linea, 6).Text = resultados(3)
            impresion.Cell(linea, 7).Text = resultados(3) / t1
            
           End If
            
            totales(1) = totales(1) + CDbl(resultados(2))
            totales(2) = totales(2) + CDbl(resultados(5))
            totales(3) = totales(3) + CDbl(resultados(3))
            
            
            totales2(1) = totales2(1) + CDbl(resultados(2))
            totales2(2) = totales2(2) + CDbl(resultados(5))
            totales2(3) = totales2(3) + CDbl(resultados(3))
            
            resultados.MoveNext
        Wend
        cSql2.Close
        Set cSql2 = Nothing
        Set resultados = Nothing
        
    ' total vendedor
    
    linea = linea + 1
           impresion.Rows = impresion.Rows + 1
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).Text = leerNombreVendedor(FILTRO)
           If totales(1) <> 0 Then
            porce = totales(2) / (totales(3) + totales(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).Text = totales(1)
            impresion.Cell(linea, 4).Text = totales(2)
            impresion.Cell(linea, 5).Text = porce
            impresion.Cell(linea, 6).Text = totales(3)
        
        For i = 1 To 4
        totales(i) = 0
        Next i
        
    ' total todos
    
    
    linea = linea + 1
            impresion.Rows = impresion.Rows + 1
            
            
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeTop) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeLeft) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeBottom) = cellThin
           impresion.Range(linea, 1, linea, 7).Borders(cellEdgeRight) = cellThin
           
           impresion.Range(linea, 2, linea, 2).Merge
           impresion.Cell(linea, 2).Text = "TOTAL GENERAL VENTAS"
            
           If totales2(1) <> 0 Then
            porce = totales2(2) / (totales2(3) + totales2(2)) * 100
            
            End If
         
            impresion.Cell(linea, 3).Text = totales2(1)
            impresion.Cell(linea, 4).Text = totales2(2)
            impresion.Cell(linea, 5).Text = porce
            impresion.Cell(linea, 6).Text = totales2(3)
            
    End If

    'Call sumaGrilla(impresion)
End Function

Public Sub generaInformevp(ByRef data As Adodc, ByRef impresion As Grid, ByVal TIPO As String, ByVal detalle As Boolean, ByVal codLoc As String, ByVal fecha1 As String, ByVal fecha2 As String)
    Dim i As Long
    Dim documento As String
    
   
    impresion.Rows = 1
    impresion.AutoRedraw = False
    
    Call cargaCabeza("LISTADO VENTAS X VENDEDORES DESDE " & Format(fecha1, "dd-mm-yyyy") & " HASTA " & Format(fecha2, "dd-mm-yyyy"), empresaActiva, impresion)
    
    Call resumenVentasproductos(data, impresion, TIPO, codLoc, fecha1, fecha2)
    
    impresion.AutoRedraw = True
    impresion.Refresh
End Sub

Private Sub CargaGrillaInforme2(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "DESCRIPCION"
        formatoGrilla(1, 3) = "VENDIDAS"
        formatoGrilla(1, 4) = "DESCUENTO"
        formatoGrilla(1, 5) = "DCTO(%)"
        formatoGrilla(1, 6) = "VENTA NETA "
        formatoGrilla(1, 7) = "P.PROMEDIO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "12"
        formatoGrilla(2, 2) = "45"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "9"
        formatoGrilla(2, 5) = "9"
        formatoGrilla(2, 6) = "9"
        formatoGrilla(2, 7) = "9"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "C"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "N"
        formatoGrilla(3, 7) = "N"
        
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = "$ ###,###,##0"
        formatoGrilla(4, 5) = "% #0.00"
        formatoGrilla(4, 6) = "$ ###,###,##0"
        formatoGrilla(4, 7) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        formatoGrilla(5, 7) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        formatoGrilla(7, 7) = ""
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "30"
        formatoGrilla(8, 3) = "8"
        formatoGrilla(8, 4) = "8"
        formatoGrilla(8, 5) = "8"
        formatoGrilla(8, 6) = "8"
        formatoGrilla(8, 7) = "8"
        
                
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
            impresion.Cell(0, i).Text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub

  Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    impresion.FixedRowColStyle = Fixed3D
    impresion.CellBorderColorFixed = vbButtonShadow
    impresion.ShowResizeTips = False
    impresion.ReportTitles.Clear
    impresion.PageSetup.CenterHorizontally = True
    impresion.PageSetup.Orientation = cellLandscape
    
      
    impresion.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    impresion.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    impresion.PageSetup.HeaderAlignment = cellLeft
    impresion.PageSetup.HeaderFont.Name = "Verdana"
    impresion.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Text = titulo1 & "  |  " & "ENTRE EL DIA  :  " & dato1.Text + "-" + dato2.Text + "-" + dato3.Text & " y " & dato4.Text + "-" + dato5.Text + "-" + dato6.Text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle
    
    If TIPO2.Value = True Then
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Text = "VENDEDOR: " & COMBOVENDEDOR.Text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle

    End If

    
    'PIE DE PAGINA
    impresion.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & usuarioSistema
    impresion.PageSetup.FooterAlignment = cellRight
    impresion.PageSetup.FooterFont.Name = "Verdana"
    impresion.PageSetup.FooterFont.Size = 7
    End Sub
    

