VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Modificarcuotas 
   BackColor       =   &H00FF8080&
   Caption         =   "MODIFICACION DE CUOTAS"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   14340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox SUCU 
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   11139
      BackColor       =   12582912
      Caption         =   "CUOTAS PENDIENTES"
      CaptionEstilo3D =   1
      BackColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton AGREGAR 
         BackColor       =   &H0080FF80&
         Caption         =   "AGREGAR CUOTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5760
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.CommandButton modificar 
         BackColor       =   &H0080FF80&
         Caption         =   "MODIFICAR CUOTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5400
         Width           =   2085
      End
      Begin VB.CommandButton RETORNAR 
         BackColor       =   &H000000FF&
         Caption         =   "RETORNO"
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
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5355
         Width           =   2085
      End
      Begin FlexCell.Grid Grid1 
         Height          =   5055
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   14145
         _ExtentX        =   24950
         _ExtentY        =   8916
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         DateFormat      =   2
      End
      Begin VB.Label lbl11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total a Pagar"
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
         Height          =   390
         Index           =   11
         Left            =   9765
         TabIndex        =   10
         Top             =   5445
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label total3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   9765
         TabIndex        =   9
         Top             =   5805
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label total2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   7425
         TabIndex        =   8
         Top             =   5805
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Interes"
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
         Height          =   390
         Index           =   2
         Left            =   7425
         TabIndex        =   7
         Top             =   5445
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label total1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4860
         TabIndex        =   6
         Top             =   5805
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Cuotas"
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
         Height          =   390
         Index           =   1
         Left            =   4815
         TabIndex        =   5
         Top             =   5445
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total General"
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
         Height          =   390
         Index           =   12
         Left            =   12015
         TabIndex        =   4
         Top             =   5445
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label total4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   12015
         TabIndex        =   3
         Top             =   5805
         Visible         =   0   'False
         Width           =   2130
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   3201
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Elimina Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1260
         Width           =   1725
      End
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
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   16
         Top             =   360
         Width           =   1410
      End
      Begin VB.TextBox NUMERO 
         Height          =   375
         Left            =   10485
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1350
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero Documento"
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
         Height          =   390
         Index           =   6
         Left            =   7830
         TabIndex        =   28
         Top             =   1350
         Width           =   2595
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
         Left            =   3060
         TabIndex        =   27
         Top             =   360
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
         Left            =   30
         TabIndex        =   26
         Top             =   390
         Width           =   1680
      End
      Begin VB.Label lblNombre 
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
         Left            =   3600
         TabIndex        =   25
         Top             =   360
         Width           =   8580
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Autorizado"
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
         Height          =   390
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   945
         Width           =   2325
      End
      Begin VB.Label lblCupo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1305
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Utilizado"
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
         Height          =   390
         Index           =   4
         Left            =   2520
         TabIndex        =   22
         Top             =   960
         Width           =   2310
      End
      Begin VB.Label lblUtilizado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   1320
         Width           =   2310
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Disponible"
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
         Height          =   390
         Index           =   5
         Left            =   5040
         TabIndex        =   20
         Top             =   945
         Width           =   2460
      End
      Begin VB.Label lblDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   5040
         TabIndex        =   19
         Top             =   1305
         Width           =   2460
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SE PUEDEN MODIFICAR LA FECHA,EL ABONO Y EL DETALLE DE COMPRA"
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
         Left            =   7800
         TabIndex        =   18
         Top             =   840
         Width           =   6255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PARA ELIMINAR UNA CUOTA  SUPR SOBRE LA CUOTA"
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
         Left            =   7800
         TabIndex        =   17
         Top             =   1080
         Width           =   6255
      End
   End
End
Attribute VB_Name = "Modificarcuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AGREGAR_Click()
 Load creditoTMPmanual
 creditoTMPmanual.rut2.text = rut2.text
 creditoTMPmanual.Show
End Sub

Private Sub Command1_Click()
Call eliminarboleta(rut2.text + lbldv.Caption, NUMERO.text)

End Sub

Private Sub Form_Activate()
rut2.SetFocus

End Sub

Private Sub Form_Load()
Call CARGAGRILLA
End Sub



Private Sub GRID1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
Select Case KeyCode
            Case 46
                If Grid1.ActiveCell.row > 0 Then
                    frmglosaeliminacion.Show vbModal
                    Call eliminarCuota(rut2.text & lbldv.Caption, Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 2), Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 4, 10), Grid1.Cell(Grid1.ActiveCell.row, 2).text, Grid1.Cell(Grid1.ActiveCell.row, 3).text)
                    
                    Grid1.RemoveItem (Grid1.ActiveCell.row)
                End If
        End Select
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub modificar_Click()
If modificar.Caption = "MODIFICAR CUOTAS" Then
Grid1.Column(3).Locked = False
Grid1.Column(5).Locked = False
Grid1.Column(11).Locked = False
Grid1.SelectionMode = cellSelectionFree
modificar.Caption = "FINALIZAR"
Else
grabarcambios
Grid1.Column(3).Locked = True
Grid1.Column(5).Locked = True
Grid1.Column(11).Locked = True
modificar.Caption = "MODIFICAR CUOTAS"
Call rut2_KeyPress(13)
End If


End Sub

Private Sub NUMERO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
NUMERO.text = ceros(NUMERO)
Call LEErcuotas(rut2.text + lbldv.Caption)
           
End If
End Sub

Private Sub RETORNAR_Click()
Grid1.Rows = 1
rut2.text = ""
lbldv.Caption = ""
lblnombre.Caption = ""
lblCupo.Caption = "$ 0"
lblUtilizado.Caption = "$ 0"
lblDisponible.Caption = "$ 0"
total1.Caption = "$ 0"
total2.Caption = "$ 0"
total3.Caption = "$ 0"
total4.Caption = "$ 0"
rut2.SetFocus


End Sub

Private Sub rut2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    Call ayudaCliente(rut2, SUCU, lbldv)
    
  End If
End Sub
Private Sub rut2_KeyPress(KeyAscii As Integer)
          KeyAscii = esNumero(KeyAscii)
           If KeyAscii = 13 And rut2.text <> "" And Val(rut2.text) <> 0 Then
             rut2.text = ceros(rut2)
             lbldv.Caption = rut(rut2.text)
             If LEERCLIENTE(rut2.text + lbldv.Caption) = True Then
        
             Call LEErcuotas(rut2.text + lbldv.Caption)
           
             Else
             MsgBox ("CLIENTE NO CORRESPONDE A CLIENTE A CREDITO O NO TIENE CUPO ASIGNADO")
             rut2.SetFocus
             End If
        End If
End Sub

Public Function LEERCLIENTE(rut) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        
        csql.sql = "SELECT mc.diapago,mc.cupodirecto,mc.cupoutilizadodirecto,mc.nombre,mc.direccion "
        csql.sql = csql.sql & "FROM sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE mc.rut='" + rut + "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
        
        lblCupo.Caption = Format(resultado(1), "###,###,##0")
        
        lblnombre.Caption = resultado(3)

        
        If resultado(1) > 0 Then
                LEERCLIENTE = True
        Else
                LEERCLIENTE = True
        
        End If
        
            resultado.MoveNext
            Wend
        Else
        LEERCLIENTE = False
        
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function
'
'    Sub LEErcuotas(rut)
'
'        Dim cSql As rdoQuery
'        Dim resultado As rdoResultset
'        Dim i As Integer
'        Dim diasmora As Double
'        Dim saldo As Double
'        Dim interes As Double
'        Dim t1 As Double
'        Dim t2 As Double
'        Dim total As Double
'        Dim porcecondo1 As Double
'        Dim porcecondo2 As Double
'        Dim cuota As Double
'        Dim interescuota As Double
'        Dim capital As Double
'        Dim cuotabase As Double
'
'        Set cSql = New rdoQuery
'        Set cSql.ActiveConnection = ventas
'
'        cSql.sql = "SELECT *  "
'        cSql.sql = cSql.sql & "FROM sv_cuotas_detalle "
'        cSql.sql = cSql.sql & "WHERE rut='" + rut + "' " ' and ( (montocuota-abono)>0 or ((interesmora+montocuota)-abono)>0)
'        cSql.sql = cSql.sql & "order by vencimientoactual "
'        cSql.Execute
'        totalusado = 0
'        moratotal = 0
'        If cSql.RowsAffected > 0 Then
'
'            Set resultado = cSql.OpenResultset
'
'        Grid1.Rows = 1
'        Grid1.AutoRedraw = False
'
'        totalusado = 0
'        moratotal = 0
'        While Not resultado.EOF
'        Grid1.Rows = Grid1.Rows + 1
'        Grid1.Cell(Grid1.Rows - 1, 0).text = resultado(0)
'        Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(1) & " " & resultado(2)
'        Grid1.Cell(Grid1.Rows - 1, 2).text = Format(resultado(4), "00") & " / " & Format(resultado(12), "00")
'        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultado(6), "dd/mm/yyyy")
'        cuotabase = resultado(7)
'        cuota = resultado(7)
'
'        interescuota = resultado(7) - resultado("capitalcuota")
'
'
'        capital = resultado("capitalcuota")
'
'        cuota = cuotabase
'
'
'
'        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(cuota, "###,###,###")
'        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(8), "###,###,###")
'
'        saldo = (cuota + resultado("interesmora")) - resultado(8)
'
'
'        tazainteresmora = leerInteresMora("00")
'        porcecondo2 = 1 - (CDbl(0) / 100)
'        tazainteresmora = tazainteresmora * porcecondo2
'        If resultado(1) <> "CA" Then
'        diasmora = DateDiff("d", resultado(6), fechasistema)
'        Else
'        diasmora = 0
'        End If
'        If diasmora <= diasgracia Then
'        diasmora = 0
'        Else
'
'        End If
'
'        interes = Round(saldo * ((tazainteresmora * diasmora) / 100), 0)
'
'        total = saldo + interes
'        If saldo = 0 Then
'        Grid1.Cell(Grid1.Rows - 1, 6).text = "0"
'        Else
'         Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
'        End If
'        Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
'        Grid1.Cell(Grid1.Rows - 1, 8).text = interes
'        If total = 0 Then
'        Grid1.Cell(Grid1.Rows - 1, 9).text = "0"
'        Else
'        Grid1.Cell(Grid1.Rows - 1, 9).text = Format(total, "###,###,###")
'        End If
'        Grid1.Cell(Grid1.Rows - 1, 10).text = "0"
'        Grid1.Cell(Grid1.Rows - 1, 11).text = resultado(13)
'        Grid1.Cell(Grid1.Rows - 1, 12).text = resultado("capitalcuota")
'
'
'
'        totalusado = totalusado + total
'        If interes <> 0 Then moratotal = moratotal + total
'        If Format(resultado(6), "yyyy-mm") <= Format(fechasistema, "yyyy-mm") Then
'        Grid1.Cell(Grid1.Rows - 1, 10).text = Grid1.Cell(Grid1.Rows - 1, 9).text
'        t1 = t1 + saldo
'        t2 = t2 + interes
'
'        End If
'
'            resultado.MoveNext
'            Wend
'        Else
'
'        End If
'        Set resultado = Nothing
'        cSql.Close
'        Set cSql = Nothing
'        Grid1.AutoRedraw = True
'        Grid1.Refresh
'        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
'        lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - totalusado, "###,###,##0")
'        total4.Caption = Format(totalusado, "###,###,##0")
'        total1.Caption = Format(t1, "###,###,##0")
'        total2.Caption = Format(t2, "###,###,##0")
'        total3.Caption = Format(t1 + t2, "###,###,##0")
'
'    SUMAPAGOS
'    End Sub
    
    
    Sub LEErcuotas(rut)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim total As Double
        Dim porcecondo1 As Double
        Dim porcecondo2 As Double
        Dim cuota As Double
        Dim interescuota As Double
        Dim capital As Double
        Dim cuotabase As Double
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT *  "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle "
        If NUMERO.text <> "" And NUMERO.text <> "0000000000" Then
        csql.sql = csql.sql & "WHERE rut='" + rut + "' and numero='" + NUMERO.text + "' "
        Else
        csql.sql = csql.sql & "WHERE rut='" + rut + "' "
        End If
        csql.sql = csql.sql & "order by fechacompra,local,tipo,numero asc "
        csql.Execute
        totalusado = 0
        moratotal = 0
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        totalusado = 0
        moratotal = 0
        While Not resultado.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 0).text = resultado(0)
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(1) & " " & resultado(2)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(4) & " / " & resultado(12)
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultado(6), "dd/mm/yyyy")
        cuotabase = resultado(7)
        cuota = resultado(7)
    
        interescuota = resultado(7) - resultado("capitalcuota")
     
        
        capital = resultado("capitalcuota")
'        If CDbl(CONDO1.text) > 0 Then
'        porcecondo1 = 1 - (CDbl(CONDO1.text) / 100)
'        interescuota = Round(interescuota * porcecondo1)
'        cuota = capital + interescuota
'        Else
        cuota = cuotabase
'        End If
        
        
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(cuota, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(8), "###,###,###")
       
        saldo = (cuota) - resultado(8)
      
        Grid1.Cell(Grid1.Rows - 1, 10).text = Format(saldo, "###,###,###")
        
        tazainteresmora = leerInteresMora("00")
'        porcecondo2 = 1 - (CDbl(CONDO2.text) / 100)
'        tazainteresmora = tazainteresmora * porcecondo2
        
        diasmora = DateDiff("d", resultado(6), fechasistema)
      If diasmora > 0 Then
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &HFF&
      End If
      
      
        If diasmora <= diasgracia Then diasmora = 0
      
        
        
        interes = Round(saldo * ((tazainteresmora / 100 / 30) * diasmora), 0)
        
        total = saldo + interes
        If saldo = 0 Then
        Grid1.Cell(Grid1.Rows - 1, 6).text = "0"
        Else
         Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
        End If
        Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
        Grid1.Cell(Grid1.Rows - 1, 8).text = interes
        If total = 0 Then
        Grid1.Cell(Grid1.Rows - 1, 9).text = "0"
        Else
        Grid1.Cell(Grid1.Rows - 1, 9).text = Format(saldo, "###,###,###")
        End If
        If Not IsNull(resultado("fechacompra")) = True Then
        fechacom = Format(resultado("fechacompra"), "dd-mm-yyyy")
        End If
        
        Grid1.Cell(Grid1.Rows - 1, 11).text = fechacom + "  " + resultado(13)
        Grid1.Cell(Grid1.Rows - 1, 12).text = resultado("capitalcuota")
        
        
        
        totalusado = totalusado + total
        If interes <> 0 Then moratotal = moratotal + total
        If Format(resultado(6), "yyyy-mm") <= Format(fechasistema, "yyyy-mm") Then
        Grid1.Cell(Grid1.Rows - 1, 10).text = Grid1.Cell(Grid1.Rows - 1, 9).text
        t1 = t1 + saldo
        t2 = t2 + interes
               
        End If
            
            resultado.MoveNext
            Wend
        Else
       
        End If
        
       
       
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
        
        lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - totalusado, "###,###,##0")
        total4.Caption = Format(totalusado, "###,###,##0")
        total1.Caption = Format(t1, "###,###,##0")
        total2.Caption = Format(t2, "###,###,##0")
        total3.Caption = Format(t1 + t2, "###,###,##0")
    
  
'    Call LEErcuotasACUMULADAS(rut)
    End Sub
    
    
    
 
Sub CARGAGRILLA()
    Grid1.Cols = 13
    
    Grid1.Column(0).Width = 20
    Grid1.Column(1).Width = 100
    Grid1.Column(2).Width = 80
    Grid1.Column(3).Width = 100
    Grid1.Column(4).Width = 60
    Grid1.Column(5).Width = 60
    Grid1.Column(6).Width = 0
    Grid1.Column(7).Width = 0
    Grid1.Column(8).Width = 0
    Grid1.Column(9).Width = 0
    Grid1.Column(10).Width = 90
    Grid1.Column(11).Width = 300
    Grid1.Column(12).Width = 0
   
    
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
    Grid1.Column(8).Locked = True
    Grid1.Column(9).Locked = True
    Grid1.Column(10).Locked = True
    Grid1.Column(11).Locked = True
    Grid1.Column(3).CellType = cellCalendar
    
    Grid1.Cell(0, 0).text = "LO"
    Grid1.Cell(0, 1).text = "DOCUMENTO"
    Grid1.Cell(0, 2).text = "N.CUOTA"
    Grid1.Cell(0, 3).text = "VENCIMIENTO"
    Grid1.Cell(0, 4).text = "CUOTA"
    Grid1.Cell(0, 5).text = "ABONO"
    Grid1.Cell(0, 6).text = "SALDO"
    Grid1.Cell(0, 7).text = "DIAS MORA"
    Grid1.Cell(0, 8).text = "INTERES"
    Grid1.Cell(0, 9).text = "TOTAL"
    Grid1.Cell(0, 10).text = "SALDO"
    

    
    
    
    Grid1.Cell(0, 11).text = "DETALLE COMPRAS "
    Grid1.Cell(0, 12).text = "CAPITAL"
    
    Grid1.Column(4).Alignment = cellRightTop
    Grid1.Column(5).Alignment = cellRightTop
    Grid1.Column(6).Alignment = cellRightTop
    Grid1.Column(7).Alignment = cellRightTop
    Grid1.Column(8).Alignment = cellRightTop
    Grid1.Column(9).Alignment = cellRightTop
    Grid1.Column(10).Alignment = cellRightTop
    
    Grid1.Column(11).Alignment = cellLeftCenter
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
Grid1.Rows = 1
   
End Sub


Sub eliminarCuota(rut, TIPO, NUMERO, cuota, fecha)
Dim CAMPOS(5, 5) As String
Dim op As Integer
Dim numerocuota As String

numerocuota = Mid(cuota, 1, 2)
numerocuota = CDbl(Replace(numerocuota, " ", ""))
    Call eliminarPagos(rut, NUMERO, numerocuota)
    
    CAMPOS(0, 2) = "sv_cuotas_detalle"
    condicion = "rut='" & rut & "' AND tipo='" & TIPO & "' and numero='" & NUMERO & "' and numerocuota='" & numerocuota & "' and vencimientoactual ='" & Format(fecha, "yyyy-mm-dd") & "' "
    op = 4
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
        sqlventas.audit = True:   sqlventas.programaactivo = Me.Caption
        Set sqlventas.conauditoria = conauditoria: sqlventas.usuarioauditoria = usuarioSistema
        sqlventas.glosaeliminacion = glosaeliminacionsistema
        sqlventas.solicitoeliminacion = solicitaeliminacion
    Call sqlventas.sqlventas(op, condicion)
Call LEErcuotas(rut2.text & lbldv.Caption)

End Sub
Sub eliminarPagos(rut, NUMERO, cuota)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim cmps(5, 5) As String
Dim op As Integer

Dim numeropago As String
Dim numerocuota As String
numerocuota = Mid(cuota, 1, 2)
numerocuota = CDbl(Replace(numerocuota, " ", ""))

Set csql.ActiveConnection = ventas
csql.sql = "select numero "
csql.sql = csql.sql & "from sv_cuotas_pago_detalle "
csql.sql = csql.sql & "where rut='" & rut & "' and numerodocumento='" & NUMERO & "' and numerocuota='" & numerocuota & "'"
csql.Execute
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
numeropago = resultados(0)
Call eliminarCabeza(rut, NUMERO, numerocuota, numeropago)

End If

cmps(0, 2) = "sv_cuotas_pago_detalle"
condicion = "rut='" & rut & "' and numero='" & numeropago & "' and numerodocumento='" & NUMERO & "' and numerocuota='" & cuota & "'"
op = 4
sqlventas.response = cmps
Set sqlventas.conexion = ventas
Call sqlventas.sqlventas(op, condicion)

End Sub
Sub eliminarCabeza(rut, NUMERO, cuota, numeropago)
Dim cmp(5, 5) As String
Dim op As Integer


cmp(0, 2) = "sv_cuotas_pago_cabeza"
condicion = "rut='" & rut & "' and numero='" & numeropago & "' "
op = 4
sqlventas.response = cmp
Set sqlventas.conexion = ventas
Call sqlventas.sqlventas(op, condicion)

End Sub
Sub grabarcambios()
Dim CAMPOS(5, 5) As String
Dim op As Integer

Dim K As Double
Dim numerocuota As String

CAMPOS(0, 0) = "vencimientooriginal"
CAMPOS(1, 0) = "vencimientoactual"
CAMPOS(2, 0) = "abono"
CAMPOS(3, 0) = "glosacompra"
CAMPOS(4, 0) = ""

For K = 1 To Grid1.Rows - 1
numerocuota = Mid(Grid1.Cell(K, 2).text, 1, 2)
numerocuota = CDbl(Replace(numerocuota, " ", ""))
CAMPOS(0, 1) = Format(Grid1.Cell(K, 3).text, "yyyy-mm-dd")
CAMPOS(1, 1) = Format(Grid1.Cell(K, 3).text, "yyyy-mm-dd")
CAMPOS(2, 1) = Replace(Grid1.Cell(K, 5).text, ".", "")
CAMPOS(3, 1) = Grid1.Cell(K, 11).text
CAMPOS(4, 1) = ""
CAMPOS(0, 2) = "sv_cuotas_detalle"
condicion = "rut='" & rut2.text & lbldv.Caption & "' and tipo='" & Mid(Grid1.Cell(K, 1).text, 1, 2) & "' and numero='" & Mid(Grid1.Cell(K, 1).text, 4, 10) & "' and numerocuota='" & numerocuota & "' "
        op = 3
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)

'If GRID1.Cell(K, 5).text = "0" Or GRID1.Cell(K, 5).text = "" Then
'Rem Call eliminarPagos(rut2.text & lbldv.Caption, Mid(GRID1.Cell(K, 1).text, 4, 10), GRID1.Cell(K, 2).text)
'End If
Next K
End Sub
Sub eliminarboleta(rut, NUMERO)
Dim CAMPOS(5, 5) As String
Dim op As Integer
Dim numerocuota As String

 
    
    CAMPOS(0, 2) = "sv_cuotas_detalle"
    condicion = "rut='" & rut & "' and numero='" & NUMERO & "' "
    op = 4
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
        sqlventas.audit = True:   sqlventas.programaactivo = Me.Caption
        Set sqlventas.conauditoria = conauditoria: sqlventas.usuarioauditoria = usuarioSistema
        sqlventas.glosaeliminacion = glosaeliminacionsistema
        sqlventas.solicitoeliminacion = solicitaeliminacion
    Call sqlventas.sqlventas(op, condicion)
Call LEErcuotas(rut2.text & lbldv.Caption)

End Sub

