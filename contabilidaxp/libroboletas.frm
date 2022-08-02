VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form auxiliar07 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Boletas"
   ClientHeight    =   8460
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   8325
   Begin XPFrame.FrameXp fechas 
      Height          =   1935
      Left            =   1755
      TabIndex        =   17
      Top             =   6480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
      BackColor       =   14737632
      Caption         =   "Rangos de Fecha"
      CaptionEstilo3D =   1
      BackColor       =   14737632
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
      Alignment       =   1
      Begin CoolButtons.cool_Button command8 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         SkinId          =   "13"
         Caption         =   "Cambia Fecha"
      End
      Begin VB.Label hastafecha 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label desdefecha 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hasta Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
   End
   Begin XPFrame.FrameXp OPCIONES 
      Height          =   6285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11086
      BackColor       =   16761024
      Caption         =   "Lista Libro de Ventas"
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
      Begin CoolButtons.cool_Button COMMAND2 
         Height          =   495
         Left            =   450
         TabIndex        =   12
         Top             =   4140
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Genera Informe"
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1095
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
         BackColor       =   16761024
         Caption         =   "Datos"
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
         Begin VB.OptionButton datos2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Rango Fecha"
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton datos1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Mensual"
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   5850
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
         BackColor       =   16761024
         Caption         =   "Resumen"
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
         Begin VB.OptionButton RESUMEN2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Resumido"
            Height          =   375
            Left            =   480
            TabIndex        =   5
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton RESUMEN1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Detallado"
            Height          =   375
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   4455
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7858
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
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   855
            Left            =   120
            TabIndex        =   7
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
               TabIndex        =   14
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
               TabIndex        =   13
               Top             =   360
               Width           =   3255
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   855
            Left            =   120
            TabIndex        =   8
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
               TabIndex        =   10
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   855
            Left            =   120
            TabIndex        =   9
            Top             =   2280
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
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
               TabIndex        =   11
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp9 
            Height          =   855
            Left            =   135
            TabIndex        =   27
            Top             =   3285
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "Centros de Costo"
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
            Begin VB.ComboBox Combocrcc 
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
               TabIndex        =   28
               Top             =   360
               Width           =   3855
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   855
         Left            =   1260
         TabIndex        =   23
         Top             =   4815
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "TIPO DE IMPRESION"
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
         Begin VB.OptionButton original 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Original"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton timbrado 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Timbrado"
            Height          =   255
            Left            =   2160
            TabIndex        =   25
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox FOLIO 
            Height          =   285
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   24
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   5040
      TabIndex        =   29
      Top             =   6480
      Width           =   3135
      _ExtentX        =   5530
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   31
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "auxiliar07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(20, 20)
Private lin As Double
Private tipoprove As String
Private plan(2000, 3) As Variant
Private canplan As Double
Private total(10) As Double
Private detalle(10, 10) As Double
Private TIPOS(6) As String
Private MES As String
Private año As String
Private centro As String








Private Sub COMMAND2_Click()
Dim TIMBRA As String

If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"

Dim infogrilla As grillainformes
Set infogrilla = New grillainformes

Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)
centro = Mid(Combocrcc.text, 1, 2) + Mid(Combocrcc.text, 4, 2)

año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)


Call CARGAGRILLA(infogrilla)
Call Consulta_Informe(infogrilla)


infogrilla.Visible = True
infogrilla.Caption = "LIBRO DE VENTAS": grillainformes.Tag = "auxiliar07" & TIMBRA & FOLIO.text

infogrilla.Show


End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)


End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato1)
    
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

Private Sub datos1_Click()
If datos2.Value = True Then fechas.Visible = True
If datos2.Value = False Then fechas.Visible = False

End Sub

Private Sub datos2_Click()
If datos2.Value = True Then fechas.Visible = True
If datos2.Value = False Then fechas.Visible = False

End Sub

Private Sub Form_Load()

CENTRAR Me

Dim i As Integer
Dim k As Integer

    
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
opciones.Visible = True

original.Value = True

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
datos1.Value = True
RESUMEN1.Value = True

desdefecha.Caption = fechasistema
hastafecha.Caption = fechasistema

fechas.Visible = False
CARGAcrcc

End Sub


    
Sub Consulta_Informe(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim TIPODO As String
    Dim total2(10) As Double
    Dim PASATBK As Boolean
    Dim NETO As Double
    Dim iva As Double
    Dim TOTALge As Double
    

    Dim PASO As String
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fecha,caja,tipodocumento,boletainicial,boletafinal,"
        csql.sql = csql.sql & "(boletafinal-boletainicial+1) as diferencia,round(monto/1.19),"
        csql.sql = csql.sql & "monto-round(monto/1.19),exento,total,estbk,cigarro "
        csql.sql = csql.sql + "FROM boletasdeventa "
        csql.sql = csql.sql + "where mid(fecha,1,7) = '" + año + "-" + MES + "'  "
        If centro <> "9999" Then csql.sql = csql.sql + "and centrocosto='" + centro + "' "
        If centro = "9999" Then csql.sql = csql.sql + "and centrocosto<>'' "
        csql.sql = csql.sql + "order by estbk,fecha "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0
        
           For k = 0 To 3
            detalle(k, 0) = 0
            detalle(k, 1) = 0
            detalle(k, 2) = 0
            detalle(k, 3) = 0
            detalle(k, 4) = 0
            detalle(k, 5) = 0
            
        Next k
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        Set resultados = csql.OpenResultset
        TIPODO = resultados(10)
        PASATBK = False
        lin = 0
         While Not resultados.EOF
        If TIPODO <> resultados(10) Then
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 6, lin, 10).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Range(lin, 6, lin, 10).FontBold = False
        
        infogrilla.Grid1.Cell(lin, 5).text = "BOLETAS"
        infogrilla.Grid1.Cell(lin, 6).text = total2(1)
        infogrilla.Grid1.Cell(lin, 7).text = total2(2)
        infogrilla.Grid1.Cell(lin, 8).text = total2(3)
        infogrilla.Grid1.Cell(lin, 9).text = total2(4)
        infogrilla.Grid1.Cell(lin, 10).text = total2(5)
        TIPODO = resultados(10)
        For k = 1 To 5
        total2(k) = 0
        Next k
        PASATBK = True
        End If
        
     
        
        If datos2.Value = True And resultados(0) < desdefecha.Caption Then GoTo PASO:
        If datos2.Value = True And resultados(0) > hastafecha.Caption Then GoTo PASO:
    
         If RESUMEN1.Value = True Then
             barra.Value = lin
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             For k = 0 To 9
             infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             
            If resultados("cigarro") = "0" Then
                infogrilla.Grid1.Cell(lin, 6).text = resultados(5)
                infogrilla.Grid1.Cell(lin, 7).text = resultados(6)
                infogrilla.Grid1.Cell(lin, 8).text = resultados(7)
                infogrilla.Grid1.Cell(lin, 9).text = resultados(8)
                infogrilla.Grid1.Cell(lin, 10).text = resultados(9)
            Else

                If resultados(8) = 0 Then
                    infogrilla.Grid1.Cell(lin, 6).text = resultados(5)
                    infogrilla.Grid1.Cell(lin, 7).text = resultados(6)
                    infogrilla.Grid1.Cell(lin, 8).text = resultados(7)
                    infogrilla.Grid1.Cell(lin, 9).text = resultados(8)
                    infogrilla.Grid1.Cell(lin, 10).text = resultados(9)
                Else
                    infogrilla.Grid1.Cell(lin, 6).text = resultados(5)
                    infogrilla.Grid1.Cell(lin, 3).text = "EXENTAS"
                    NETO = Round(resultados(8) / 1.19)
                    iva = resultados(8) - NETO
                    infogrilla.Grid1.Cell(lin, 7).text = NETO * -1
                    infogrilla.Grid1.Cell(lin, 8).text = iva * -1
                    infogrilla.Grid1.Cell(lin, 9).text = NETO + iva
                    infogrilla.Grid1.Cell(lin, 10).text = resultados(8)
                
                End If
           End If
          

         
         End If
         
         If resultados("cigarro") = "0" Then
         
             total(1) = total(1) + resultados(5)
             total(2) = total(2) + resultados(6)
             total(3) = total(3) + resultados(7)
             total(4) = total(4) + resultados(8)
             total(5) = total(5) + resultados(9)
             total2(1) = total2(1) + resultados(5)
             total2(2) = total2(2) + resultados(6)
             total2(3) = total2(3) + resultados(7)
             total2(4) = total2(4) + resultados(8)
             total2(5) = total2(5) + resultados(9)
         Else
            
             total(1) = total(1) + CDbl(infogrilla.Grid1.Cell(lin, 6).text)
             total(2) = total(2) + CDbl(infogrilla.Grid1.Cell(lin, 7).text)
             total(3) = total(3) + CDbl(infogrilla.Grid1.Cell(lin, 8).text)
             total(4) = total(4) + CDbl(infogrilla.Grid1.Cell(lin, 9).text)
             total(5) = total(5) + CDbl(infogrilla.Grid1.Cell(lin, 10).text)
             total2(1) = total2(1) + CDbl(infogrilla.Grid1.Cell(lin, 6).text)
             total2(2) = total2(2) + CDbl(infogrilla.Grid1.Cell(lin, 7).text)
             total2(3) = total2(3) + CDbl(infogrilla.Grid1.Cell(lin, 8).text)
             total2(4) = total2(4) + CDbl(infogrilla.Grid1.Cell(lin, 9).text)
             total2(5) = total2(5) + CDbl(infogrilla.Grid1.Cell(lin, 10).text)
         End If
             

             'bfiscal
If resultados("tipodocumento") = "" Then
             detalle(1, 1) = detalle(1, 1) + resultados(5)
             detalle(1, 2) = detalle(1, 2) + resultados(6)
             detalle(1, 3) = detalle(1, 3) + resultados(7)
             detalle(1, 4) = detalle(1, 4) + resultados(8)
             detalle(1, 5) = detalle(1, 5) + resultados(9)

End If
             
             'belectronica
If resultados("tipodocumento") = "E" Then
             detalle(2, 1) = detalle(2, 1) + resultados(5)
             detalle(2, 2) = detalle(2, 2) + resultados(6)
             detalle(2, 3) = detalle(2, 3) + resultados(7)
             detalle(2, 4) = detalle(2, 4) + resultados(8)
             detalle(2, 5) = detalle(2, 5) + resultados(9)
End If
             'b.exenta
If resultados("tipodocumento") = "X" Then
             detalle(3, 1) = detalle(3, 1) + resultados(5)
             detalle(3, 2) = detalle(3, 2) + resultados(6)
             detalle(3, 3) = detalle(3, 3) + resultados(7)
             detalle(3, 4) = detalle(3, 4) + resultados(8)
             detalle(3, 5) = detalle(3, 5) + resultados(9)
End If
             
             
             
        
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
If PASATBK = True Then
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 6, lin, 10).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 5).text = "TRANSBANK"
        infogrilla.Grid1.Cell(lin, 6).text = total2(1)
        infogrilla.Grid1.Cell(lin, 7).text = total2(2)
        infogrilla.Grid1.Cell(lin, 8).text = total2(3)
        infogrilla.Grid1.Cell(lin, 9).text = total2(4)
        infogrilla.Grid1.Cell(lin, 10).text = total2(5)
        infogrilla.Grid1.Range(lin, 6, lin, 10).FontBold = False
        
        For k = 1 To 5
        total2(k) = 0
        Next k
        
        End If
             
     
Call totallibro(infogrilla)
barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub

Sub totallibro(infogrilla As grillainformes)
    
    Dim TOTALge As Double
        
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 6, lin, 10).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 5).text = "TOTAL"
        infogrilla.Grid1.Cell(lin, 5).Alignment = cellLeftCenter
        infogrilla.Grid1.Cell(lin, 6).text = total(1)
        infogrilla.Grid1.Cell(lin, 7).text = total(2)
        infogrilla.Grid1.Cell(lin, 8).text = total(3)
        infogrilla.Grid1.Cell(lin, 9).text = total(4)
        infogrilla.Grid1.Cell(lin, 10).text = total(5)
        
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 6, lin, 10).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 5).text = "NORMAL"
        infogrilla.Grid1.Cell(lin, 5).Alignment = cellLeftCenter
        infogrilla.Grid1.Cell(lin, 6).text = detalle(1, 1)
        infogrilla.Grid1.Cell(lin, 7).text = detalle(1, 2)
        infogrilla.Grid1.Cell(lin, 8).text = detalle(1, 3)
        infogrilla.Grid1.Cell(lin, 9).text = detalle(1, 4)
        infogrilla.Grid1.Cell(lin, 10).text = detalle(1, 5)
    
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 6, lin, 10).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 5).text = "ELECTRONICA"
        infogrilla.Grid1.Cell(lin, 5).Alignment = cellLeftCenter
        infogrilla.Grid1.Cell(lin, 6).text = detalle(2, 1)
        infogrilla.Grid1.Cell(lin, 7).text = detalle(2, 2)
        infogrilla.Grid1.Cell(lin, 8).text = detalle(2, 3)
        infogrilla.Grid1.Cell(lin, 9).text = detalle(2, 4)
        infogrilla.Grid1.Cell(lin, 10).text = detalle(2, 5)
    
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 6, lin, 10).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 5).text = "EXENTA"
        infogrilla.Grid1.Cell(lin, 5).Alignment = cellLeftCenter
        infogrilla.Grid1.Cell(lin, 6).text = detalle(3, 1)
        infogrilla.Grid1.Cell(lin, 7).text = detalle(3, 2)
        infogrilla.Grid1.Cell(lin, 8).text = detalle(3, 3)
        infogrilla.Grid1.Cell(lin, 9).text = detalle(3, 4)
        infogrilla.Grid1.Cell(lin, 10).text = detalle(3, 5)
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
    TOTALge = 0
    
    
    End Sub
    





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 8
        
    
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "CAJA"
    FORMATOGRILLA(1, 3) = "TIPO"
    FORMATOGRILLA(1, 4) = "B.INICIAL"
    FORMATOGRILLA(1, 5) = "B.FINAL"
    FORMATOGRILLA(1, 6) = "B.EMITIDAS"
    FORMATOGRILLA(1, 7) = "NETO"
    FORMATOGRILLA(1, 8) = "I.V.A"
    FORMATOGRILLA(1, 9) = "EXENTO"
    FORMATOGRILLA(1, 10) = "TOTAL"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "6"
    FORMATOGRILLA(2, 3) = "6"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"
    FORMATOGRILLA(2, 10) = "10"
    FORMATOGRILLA(2, 11) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,##0"
    FORMATOGRILLA(4, 7) = "###,###,##0"
    FORMATOGRILLA(4, 8) = "###,###,##0"
    FORMATOGRILLA(4, 9) = "###,###,##0"
    FORMATOGRILLA(4, 10) = "###,###,##0"
    FORMATOGRILLA(4, 11) = "###,###,##0"
    
    Rem LOCCKED
    For k = 1 To 10
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 11
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
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

Sub leermayor()
    tipoprove = CUENTAPROVEEDOR
    

    
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

Sub CARGAcrcc()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM centrosdecosto where año='" + COMBOAÑO.text + "' "
        csql.sql = csql.sql + "order by codigo"
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             Combocrcc.AddItem (Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + " " + resultados(1))
             
            resultados.MoveNext
            Wend
        End If
        Combocrcc.AddItem ("99.99" + " " + "TODOS")
            
        Combocrcc.text = Combocrcc.List(LINEA)
        
   

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

