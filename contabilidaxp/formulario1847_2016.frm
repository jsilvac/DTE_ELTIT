VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form1847_2016 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BALANCE TRIBUTARIO"
   ClientHeight    =   7155
   ClientLeft      =   240
   ClientTop       =   1290
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   4560
      TabIndex        =   20
      Top             =   6480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      BackColor       =   16744576
      Caption         =   " Mis Datos"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   855
      Left            =   1320
      TabIndex        =   10
      Top             =   4680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      BackColor       =   16761024
      Caption         =   "INGRESO FOLIOS TRIMBRADOS"
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
      Begin VB.TextBox foliofin 
         Height          =   285
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox FOLIOini 
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton timbrado 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprime Timbrado"
         Height          =   255
         Left            =   -1920
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton original 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprime Original"
         Height          =   255
         Left            =   -2520
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "HASTA"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DESDE"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
   End
   Begin CoolButtons.cool_Button GENERA 
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   6600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "GENERA INFORME"
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   4215
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7435
      BackColor       =   16761024
      Caption         =   "Configuracion"
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cuentas Madres"
         Height          =   195
         Left            =   3720
         TabIndex        =   18
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Detallado"
         Height          =   195
         Left            =   720
         TabIndex        =   17
         Top             =   3840
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   855
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "EMPRESA"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox DATO1 
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
            Left            =   240
            TabIndex        =   3
            Text            =   "01"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label empresanombre 
            BackStyle       =   0  'Transparent
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
            Left            =   840
            TabIndex        =   4
            Top             =   360
            Width           =   3255
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   855
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         BackColor       =   16744576
         Caption         =   "MES"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox COMBOMES 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   3855
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   1095
         Left            =   1320
         TabIndex        =   7
         Top             =   2280
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1931
         BackColor       =   16744576
         Caption         =   "AÑO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox COMBOAÑO 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   3855
         End
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   855
      Left            =   2160
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1508
      BackColor       =   16761024
      Caption         =   "OPCION DE IMPRESION"
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
      Begin VB.OptionButton ccodigo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Con Codigo"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton scodigo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sin Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "form1847_2016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CAN22 As Double

Private f22(100, 3) As Variant
Private FORMATOGRILLA(10, 20) As String
Private sumas(10) As Double
Private suma(10) As Double
Private sumas2(10) As Double
Private sumas3(10) As Double
Private sumade As Double
Private sumaha As Double

 

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato1)
    
End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub


Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    COMBOMES.SetFocus
    empresanombre.Caption = sqlconta.response(1, 3)
no:
End Sub

Private Sub Form_Load()
CENTRAR Me

 Call Conectar_BD
 Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
dato1.text = empresaactiva
empresanombre.Caption = nombreempresa
original.Value = True
CARGAf22
End Sub
Sub ACEPTA(opcion)
Dim TIMBRA As String

Dim infogrilla As grilla1847_2016
Set infogrilla = New grilla1847_2016
If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"
    If opcion = 1 Then
    infogrilla.Caption = "BALANCE TRIBUTARIO"
    infogrilla.CABEZA.Caption = "BALANCE TRIBUTARIO desde el 1 de Enero del " & COMBOAÑO.text & " al " & Mid(DateSerial(CDbl(COMBOAÑO.text), (COMBOMES.ListIndex + 1) + 1, 0), 1, 2) & " de " & COMBOMES.text & " del " & COMBOAÑO.text + " de la empresa " + empresanombre.Caption
    End If


Call CARGAGRILLA(infogrilla)

Call CARGABALANCE(infogrilla)

infogrilla.Visible = True

infogrilla.Show

End Sub




    Sub DIFERENCIA(infogrilla As grilla1847_2016, row)
    infogrilla.Grid1.Rows = row + 1
     With infogrilla.Grid1.Range(row, 1, row, 10)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
    
    infogrilla.Grid1.Cell(row, 2).text = "RESULTADOS"
   
    For k = 1 To 8
    infogrilla.Grid1.Cell(row, k + 2).text = difer(k - 1)
  
    Next k
    End Sub
    Sub totales(infogrilla As grilla1847_2016, row)
    Dim DIFE As Double
    Dim H As Double
    Dim j As Double
    
    For j = 1 To CAN22
    f22(j, 0) = 0
    Next j
    
 
    For H = 1 To infogrilla.Grid1.Rows - 1
    For j = 1 To CAN22
    If f22(j, 1) = infogrilla.Grid1.Cell(H, 12).text Then
    If f22(j, 3) = "+" Then f22(j, 0) = f22(j, 0) + (Val(infogrilla.Grid1.Cell(H, 11).text) - Val(infogrilla.Grid1.Cell(H, 10).text))
    If f22(j, 3) = "-" Then f22(j, 0) = f22(j, 0) - (Val(infogrilla.Grid1.Cell(H, 10).text) - Val(infogrilla.Grid1.Cell(H, 11).text))
    
    End If
    
    Next j
    
    Next H
    
    
    
    
    infogrilla.Grid1.Rows = row + 1
    
     With infogrilla.Grid1.Range(row, 1, row, 11)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
    
    infogrilla.Grid1.Cell(row, 1).text = ""
    infogrilla.Grid1.Cell(row, 2).text = "TOTALES"
                 
    For k = 1 To 8
    infogrilla.Grid1.Cell(row, k + 3).text = sumas(k)
    sumas2(k) = 0
    Next k
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
    
     With infogrilla.Grid1.Range(row + 1, 1, row + 1, 11)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
    
    DIFE = sumas(1) - sumas(2)
    If DIFE < 0 Then sumas2(1) = DIFE * -1 Else sumas2(2) = DIFE
    DIFE = sumas(3) - sumas(4)
    If DIFE < 0 Then sumas2(3) = DIFE * -1 Else sumas2(4) = DIFE
    DIFE = sumas(5) - sumas(6)
    If DIFE < 0 Then sumas2(5) = DIFE * -1 Else sumas2(6) = DIFE
    DIFE = sumas(7) - sumas(8)
    If DIFE < 0 Then sumas2(7) = DIFE * -1 Else sumas2(8) = DIFE
    
    infogrilla.Grid1.Cell(row + 1, 1).text = ""
    infogrilla.Grid1.Cell(row + 1, 2).text = "RESULTADOS EJERCICIOS"
     infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
     With infogrilla.Grid1.Range(row + 2, 1, row + 2, 11)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
                 
                 
    For k = 1 To 8
    infogrilla.Grid1.Cell(row + 1, k + 3).text = sumas2(k)
    sumas3(k) = sumas2(k) + sumas(k)
    
    Next k
    
    infogrilla.Grid1.Cell(row + 2, 1).text = ""
    infogrilla.Grid1.Cell(row + 2, 2).text = "SUMAS IGUALES"
                 
    For k = 1 To 8
    infogrilla.Grid1.Cell(row + 2, k + 3).text = sumas3(k)
    
    Next k
 
 Dim totalf22 As Double
 
 For k = 1 To CAN22
 If f22(k, 0) <> 0 Then
 
 infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
 infogrilla.Grid1.Column(2).Locked = False
 infogrilla.Grid1.Column(3).Locked = False
 
 
 infogrilla.Grid1.Range(infogrilla.Grid1.Rows - 1, 2, infogrilla.Grid1.Rows - 1, 3).Merge
 
 infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 1).text = f22(k, 1)
 infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 2).text = f22(k, 2)
 infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 4).text = Format(f22(k, 0), "###,###,###,##0")
 infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 5).text = f22(k, 3)
 totalf22 = totalf22 + f22(k, 0)
 Rem If F22(k, 3) = "-" Then totalf22 = totalf22 - F22(k, 0)
 End If
 Next k
 infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
 infogrilla.Grid1.Column(2).Locked = False
 infogrilla.Grid1.Column(3).Locked = False
 
 
 infogrilla.Grid1.Range(infogrilla.Grid1.Rows - 1, 2, infogrilla.Grid1.Rows - 1, 3).Merge
 
 infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 2).Font.Bold = True
 
 
 infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 2).text = "TOTAL BASE IMPONIBLE"
 infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 4).Font.Bold = True
 
 
 infogrilla.Grid1.Cell(infogrilla.Grid1.Rows - 1, 4).text = Format(totalf22, "###,###,###,##0")
 
 infogrilla.Grid1.Column(2).Locked = True
 infogrilla.Grid1.Column(3).Locked = True
 
 
 
 
    
    
    End Sub
    




Sub total()
    
End Sub
Sub TOTAL1()
    difer(0) = 0: difer(1) = 0: difer(2) = 0: difer(3) = 0
    If sumas(5) > sumas(4) Then difer(4) = sumas(5) - sumas(4): difer(5) = 0
    If sumas(4) > sumas(5) Then difer(5) = sumas(4) - sumas(5): difer(4) = 0
    
    If sumas(7) > sumas(6) Then difer(6) = sumas(7) - sumas(6): difer(7) = 0
    If sumas(6) > sumas(7) Then difer(7) = sumas(6) - sumas(7): difer(6) = 0
    
    
    sumast(0) = sumas(0) + difer(0)
    sumast(1) = sumas(1) + difer(1)
    sumast(2) = sumas(2) + difer(2)
    sumast(3) = sumas(3) + difer(3)
    sumast(4) = sumas(4) + difer(4)
    sumast(5) = sumas(5) + difer(5)
    sumast(6) = sumas(6) + difer(6)
    sumast(7) = sumas(7) + difer(7)

    suma(0) = 0: suma(1) = 0: suma(2) = 0: suma(3) = 0: suma(4) = 0: suma(5) = 0: suma(6) = 0: suma(7) = 0
    
                
End Sub
Sub LEERSALDOS(LLAVE, tipo)
Dim SUMD As Double
Dim SUMH As Double
Dim anted As Double
Dim anteh As Double
Dim DIFE As Double
Dim fechaproceso As String


    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = ""
    
    condicion = "codigo=" + "'" + LLAVE + "' and año ='" + COMBOAÑO.text + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    
    Call sqlconta.sqlconta(op, condicion)
 '   If sqlconta.status = 4 Then Stop
    anted = sqlconta.response(2, 3)
    anteh = sqlconta.response(3, 3)
    Rem anted = 0
     Rem anteh = 0
    fechaproceso = DateSerial(año, MES + 1, 0)
    
    
    
   Call LEERSALDOSMAYOR(LLAVE, Format(fechaproceso, "yyyy-mm-dd"))
   Rem  sumade = 0
    Rem sumaha = 0
    SUMD = sumade: SUMH = sumaha

For k = 1 To 8
suma(k) = 0
sumas2(k) = 0
sumas3(k) = 0
Next k

suma(1) = anted + SUMD
suma(2) = anteh + SUMH
DIFE = suma(1) - suma(2)

If DIFE > 0 Then suma(3) = DIFE
If DIFE < 0 Then suma(4) = DIFE * -1


If tipo = "1" Or tipo = "2" Then suma(5) = suma(3): suma(6) = suma(4)

If tipo <> "1" And tipo <> "2" Then suma(7) = suma(3): suma(8) = suma(4)
For k = 1 To 8
sumas(k) = sumas(k) + suma(k)
Next k

End Sub
Sub LEERSALDOSMAYOR(codigo, fecha)
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim fecha1 As String
    Dim fecha2 As String
    Dim resultados As rdoResultset
  Dim NIVEL As String
  
    fecha1 = Format(fecha, "yyyy") + "-01-01"
    fecha2 = Format(fecha, "yyyy-mm-dd")
        Set csql2.ActiveConnection = contadb
       NIVEL = "3"
        If Mid(codigo, 5, 5) = "0000" Then NIVEL = "2"
        If Mid(codigo, 3, 6) = "000000" Then NIVEL = "1"
        csql2.sql = "SELECT fecha,sum(monto),dh "
        csql2.sql = csql2.sql + "FROM movimientoscontables WHERE fecha between '" + fecha1 + "' and '" + fecha2 + "' "
        If NIVEL = "1" Then
        csql2.sql = csql2.sql + "and mid(codigocuenta,1,2)='" + Mid(codigo, 1, 2) + "' "
        End If
        If NIVEL = "2" Then
        csql2.sql = csql2.sql + "and mid(codigocuenta,1,4)='" + Mid(codigo, 1, 4) + "' "
        End If
        If NIVEL = "3" Then
        csql2.sql = csql2.sql + "and codigocuenta='" + codigo + "' "
        End If
        
        
        csql2.sql = csql2.sql + " group by dh "
        csql2.Execute
        LINEAS = 0
        sumade = 0: sumaha = 0
        If csql2.RowsAffected > 0 Then
         
        Set resultados = csql2.OpenResultset
        While Not resultados.EOF
        If resultados(2) = "D" Then
        sumade = resultados(1)
        Else
        sumaha = resultados(1)
        End If
        
        
        
        resultados.MoveNext
        Wend
          
          resultados.Close
            Set resultados = Nothing
        End If
  

  
End Sub




Sub CARGAGRILLA(infogrilla As grilla1847_2016)
Rem DATOS DE LA COLUMNA
    
    
    FORMATOGRILLA(1, 1) = " CODIGO "
    FORMATOGRILLA(1, 2) = " SII "
    FORMATOGRILLA(1, 3) = " CUENTA         "
    FORMATOGRILLA(1, 4) = "DEBITOS"
    FORMATOGRILLA(1, 5) = "CREDITOS"
    FORMATOGRILLA(1, 6) = "DEUDOR"
    FORMATOGRILLA(1, 7) = "ACREEDOR"
    FORMATOGRILLA(1, 8) = " ACTIVO"
    FORMATOGRILLA(1, 9) = "PASIVO"
    FORMATOGRILLA(1, 10) = "PERDIDA"
    FORMATOGRILLA(1, 11) = "GANANCIA"
    FORMATOGRILLA(1, 12) = " F 22 "
    FORMATOGRILLA(1, 13) = "V.TRIBUTARIO "
    Rem LARGO DE LOS DATOS
    If scodigo.Value = True Then
        FORMATOGRILLA(2, 1) = "0"
    Else
        FORMATOGRILLA(2, 1) = "8"
    End If
    FORMATOGRILLA(2, 2) = "12"
    
    FORMATOGRILLA(2, 3) = "28"
    
    FORMATOGRILLA(2, 4) = "12"
    FORMATOGRILLA(2, 5) = "11"
    FORMATOGRILLA(2, 6) = "11"
    FORMATOGRILLA(2, 7) = "11"
    FORMATOGRILLA(2, 8) = "11"
    FORMATOGRILLA(2, 9) = "11"
    FORMATOGRILLA(2, 10) = "11"
    FORMATOGRILLA(2, 11) = "11"
    FORMATOGRILLA(2, 12) = "11"
    FORMATOGRILLA(2, 13) = "11"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "N"
    
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "S"
    FORMATOGRILLA(3, 13) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = "###,###,###,###"
    FORMATOGRILLA(4, 5) = "###,###,###,###"
    FORMATOGRILLA(4, 6) = "###,###,###,###"
    FORMATOGRILLA(4, 7) = "###,###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###,###"
    FORMATOGRILLA(4, 10) = "###,###,###,###"
    FORMATOGRILLA(4, 11) = "###,###,###,###"
    FORMATOGRILLA(4, 12) = ""
    FORMATOGRILLA(4, 13) = "###,###,###,###"
    
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    
    infogrilla.Grid1.Cols = 14
    infogrilla.Grid1.Rows = 2
    
     'INFOGRILLA.GRID1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'INFOGRILLA.GRID1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'INFOGRILLA.GRID1.BackColorFixed = RGB(90, 158, 214)
   ' INFOGRILLA.GRID1.BackColorFixedSel = RGB(110, 180, 230)
   ' INFOGRILLA.GRID1.BackColorBkg = RGB(90, 158, 214)
   ' INFOGRILLA.GRID1.BackColorScrollBar = RGB(231, 235, 247)
   ' INFOGRILLA.GRID1.BackColor1 = RGB(231, 235, 247)
   ' INFOGRILLA.GRID1.BackColor2 = RGB(239, 243, 255)
   ' INFOGRILLA.GRID1.GridColor = RGB(148, 190, 231)
    infogrilla.Grid1.Column(0).Width = 0
    
    For k = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        infogrilla.Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        infogrilla.Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub
   

Sub CARGABALANCE(infogrilla As grilla1847_2016)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim lin As Double
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,f_1847,nombre,tipo,IFNULL(f_1847_2016,''),patrimonio "
        csql.sql = csql.sql + "FROM cuentasdelmayor "
        csql.sql = csql.sql + "WHERE año='" + COMBOAÑO.text + "' "
        csql.sql = csql.sql + " and mid(codigo,5,4)<>'0000' "
        csql.sql = csql.sql + "order by codigo,año "
        csql.Execute
        lin = 0
        For k = 1 To 8
        sumas(k) = 0
        sumas2(k) = 0
        sumas3(k) = 0
        sumast(k) = 0
        
        Next k
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             barra.Max = csql.RowsAffected + 1
             barra.Value = 0
                While Not resultados.EOF
                    
                    Call LEERSALDOS(resultados(0), resultados(3))
                            If suma(1) + suma(2) <> 0 Then
                            lin = lin + 1
                            infogrilla.Grid1.Rows = lin + 1
                            
                            barra.Value = barra.Value + 1
                            infogrilla.Grid1.Cell(lin, 1).text = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4)
                            infogrilla.Grid1.Cell(lin, 2).text = resultados(1)
                            infogrilla.Grid1.Cell(lin, 3).text = resultados(2)
                            For k = 1 To 8
                            infogrilla.Grid1.Cell(lin, k + 3).text = suma(k)
                            Next k
                            infogrilla.Grid1.Cell(lin, 12).text = resultados(4)
                            If resultados(3) < 3 Then
                            If resultados("patrimonio") = 0 Then
                            infogrilla.Grid1.Cell(lin, 13).text = suma(1) - suma(2)
                            End If
                            End If
                            
                            
                            
                            
                            End If
                        
                resultados.MoveNext
                Wend
            Call totales(infogrilla, infogrilla.Grid1.Rows)
            
            
            resultados.Close
            
            Set resultados = Nothing
' datos finales
            
            infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 5
            lin = lin + 5
            For k = 1 To 10
            infogrilla.Grid1.Column(k).Locked = False
            
            
            Next k
            
            
            
        End If
    


End Sub

Sub CARGAf22()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Double
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT * "
        csql.sql = csql.sql + "FROM sii_1846_2016 "
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
                While Not resultados.EOF
                    LINEA = LINEA + 1
                    f22(LINEA, 1) = resultados(0)
                    f22(LINEA, 2) = resultados(1)
                    f22(LINEA, 3) = resultados(2)
                    
                resultados.MoveNext
                Wend
            
            
            resultados.Close
            
            Set resultados = Nothing
' datos fina
        End If
    

CAN22 = LINEA
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Private Sub NOIMPRIME_Click()
Unload Me
End Sub

Private Sub GENERA_Click()

Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
If MES <> 12 Then MsgBox "NO PUEDE GENERARA EN OTRA FECHA QUE NO SEA DICIEMBRE ": Exit Sub
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)
FOLIOini2 = FOLIOini.text
foliofin2 = foliofin.text
If Val(FOLIOini2) = 0 Or Val(foliofin2) = 0 Then
MsgBox "PROBLEMA CON LOS FOLIOS TIMBRADOS "
Exit Sub
End If


Call ACEPTA(1)
Unload Me

End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
