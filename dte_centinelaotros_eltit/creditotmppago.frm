VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form creditotmppago 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPRESION PAGOS"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ELIMINA 
      BackColor       =   &H000000FF&
      Caption         =   "ELIMINA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8100
      Width           =   2400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "RETORNO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6165
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8100
      Width           =   2400
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8100
      Width           =   2400
   End
   Begin XPFrame.FrameXp detalle 
      Height          =   4470
      Left            =   180
      TabIndex        =   0
      Top             =   3465
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7885
      BackColor       =   16761024
      Caption         =   "DETALLE"
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
         Height          =   3435
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   6059
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6300
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   6885
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox PIVOTE3 
      Height          =   285
      Left            =   4815
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   6390
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox PIVOTE2 
      Height          =   285
      Left            =   4860
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   6615
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblcajera 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   375
      Left            =   180
      TabIndex        =   26
      Top             =   2115
      Width           =   8805
   End
   Begin VB.Label repa 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   180
      TabIndex        =   25
      Top             =   90
      Width           =   8790
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3690
      TabIndex        =   19
      Top             =   1035
      Width           =   195
   End
   Begin VB.Label rut2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1845
      TabIndex        =   18
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label año 
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
      Height          =   375
      Left            =   6345
      TabIndex        =   17
      Top             =   630
      Width           =   930
   End
   Begin VB.Label mes 
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
      Height          =   375
      Left            =   5805
      TabIndex        =   16
      Top             =   630
      Width           =   480
   End
   Begin VB.Label dia 
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
      Height          =   375
      Left            =   5265
      TabIndex        =   15
      Top             =   630
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUT"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   180
      TabIndex        =   14
      Top             =   1035
      Width           =   1635
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
      Height          =   375
      Left            =   180
      TabIndex        =   13
      Top             =   1575
      Width           =   8805
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
      Height          =   375
      Left            =   3915
      TabIndex        =   12
      Top             =   1080
      Width           =   285
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interes Mora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   6705
      TabIndex        =   11
      Top             =   2700
      Width           =   1950
   End
   Begin VB.Label interesmora 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6705
      TabIndex        =   10
      Top             =   3015
      Width           =   1950
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FOLIO"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   180
      TabIndex        =   9
      Top             =   630
      Width           =   1635
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Left            =   3825
      TabIndex        =   8
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Cuota"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   3555
      TabIndex        =   7
      Top             =   2700
      Width           =   1950
   End
   Begin VB.Label total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   225
      TabIndex        =   6
      Top             =   3015
      Width           =   1950
   End
   Begin VB.Label montocuota 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3510
      TabIndex        =   5
      Top             =   3015
      Width           =   1950
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total a Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   225
      TabIndex        =   4
      Top             =   2700
      Width           =   1950
   End
   Begin VB.Label folio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1845
      TabIndex        =   3
      Top             =   630
      Width           =   1860
   End
End
Attribute VB_Name = "creditotmppago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click()
If MsgBox("SEGURO QUE DESEA IMPRIMIR COMPROBANTE", vbOKCancel, "ATENCION") = vbOK Then

Rem IMPRIMEPAGOcredito
imprime2
End If
End Sub
Public Sub IMPRIMEPAGOcredito()
    Dim K As Integer
    Dim o As Integer
    ''''''''''''''''''
    numfic = 20
    
If impresoracredito = "0" Then
    Open "impresion.txt" For Output As #numfic
    End If
    If impresoracredito = "1" Then
    Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #numfic
    End If
    If impresoracredito = "2" Then
    Open "LPT1" For Output As #numfic
    End If
        
    ''''''''''''''''''

    '''''''''''''''''''''
    'EMPAQUE
    '''''''''''''''''''''
    For K = 1 To 1
    Print #numfic, Chr$(27); Chr$(64) '
    Print #numfic, "ALMACENES ELTIT   "
    Print #numfic, ""
    Print #numfic, "          COMPROBANTE DE PAGO CREDITO    "
    Print #numfic, "          ===========================    "
    Print #numfic,
    Print #numfic, "NUMERO  :"; FOLIO.Caption
    Print #numfic, "FECHA   :"; Format(fechasistema, "dd-mm-yyyy")
    Print #numfic, "CLIENTE :"; rut2.Caption + "-" + lbldv.Caption
    Print #numfic, "NOMBRE  :"; lblnombre.Caption
    Print #numfic,
    Print #numfic, "PAGO    :"; total.Caption
    If interesmora.Caption <> "" Then
    Print #numfic, "MORA    :"; interesmora.Caption
    End If
    
    Print #numfic, "DETALLE CUOTAS "
    Print #numfic,
    
    Print #numfic, "N° DOC.       N° CUOTA      MONTO       "
    For o = 1 To Grid1.Rows - 1
    pivote.MaxLength = 10
    pivote.text = Grid1.Cell(o, 4).text
    pivote.text = ceros(pivote)
    
    PIVOTE2.MaxLength = 10
    PIVOTE2.text = Grid1.Cell(o, 5).text
    PIVOTE2.text = ceros(PIVOTE2)
    
    PIVOTE3.MaxLength = 10
    PIVOTE3.text = Grid1.Cell(o, 6).text
    PIVOTE3.text = String(PIVOTE3.MaxLength - Len(PIVOTE3.text), " ") & PIVOTE3.text
    
    Print #numfic, pivote.text; "    "; PIVOTE2.text; "    "; PIVOTE3.text
    Next o
    
    Print #numfic, "              _______________            "
    Print #numfic, "               FIRMA Y TIMBRE            "
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    
    Print #numfic, Chr(27); "i"
    Next K
    Close #numfic

 If impresoracredito = "0" Then Shell "notepad impresion.txt"

End Sub
Sub imprime2()
Titulos



Grid1.PrintPreview

End Sub

Private Sub Command1_Click()
Unload Me
creditoPAGOSTMP.rut2.SetFocus
End Sub



Private Sub ELIMINA_Click()
frmglosaeliminacion.Show vbModal
Call ELIMINARTODO(FOLIO.Caption)
Unload Me

End Sub

Private Sub Form_Load()
CARGAGRILLA
'folio.Caption = creditoPAGOSTMP.folio.Caption

rut2.Caption = creditoPAGOSTMP.rut2.text
lbldv.Caption = creditoPAGOSTMP.lbldv.Caption
lblnombre.Caption = creditoPAGOSTMP.lblnombre.Caption
Call LEERCABEZACUOTAS(creditoPAGOSTMP.FOLIO.Caption)
Call LEERPAGOCUOTAS(creditoPAGOSTMP.FOLIO.Caption)
ELIMINA.Visible = False


If Verifica_Permiso("PANTALLA PAGO CUOTAS", "autoriza") = True Then
    ELIMINA.Visible = True

End If

End Sub
Sub CARGAGRILLA()
    Grid1.Cols = 6
    Grid1.DefaultFont.Size = 12
    
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 50
    Grid1.Column(2).Width = 120
    Grid1.Column(3).Width = 100
    Grid1.Column(4).Width = 150
    Grid1.Column(5).Width = 100
   
    
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
   
     
    Grid1.Cell(0, 1).text = "TIPO"
    Grid1.Cell(0, 2).text = "DOCUMENTO"
    Grid1.Cell(0, 3).text = "CUOTA"
    Grid1.Cell(0, 4).text = "VENCIMIENTO"
    Grid1.Cell(0, 5).text = "MONTO"
    Grid1.Column(1).Alignment = cellCenterCenter
    Grid1.Column(2).Alignment = cellCenterCenter
    
    Grid1.Column(3).Alignment = cellCenterCenter
    Grid1.Column(4).Alignment = cellCenterCenter
    
    
    
    Grid1.Column(5).Alignment = cellRightCenter
    
    
    
      
    
    Grid1.ExtendLastCol = True
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
Grid1.Rows = 1
   
End Sub
Sub LEERPAGOCUOTAS(folio1)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
            
            
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT tipodocumento,numerodocumento,totalcuotas,vencimiento,monto  "
        csql.sql = csql.sql & "FROM sv_cuotas_pago_detalle "
        csql.sql = csql.sql & "WHERE numero='" + folio1 + "' and rut='" + rut2.Caption + lbldv.Caption + "' "
        csql.sql = csql.sql & "order by numerocuota "
        csql.Execute
        
         If csql.RowsAffected > 0 Then
         Set resultado = csql.OpenResultset
         Grid1.Rows = 1
         Grid1.AutoRedraw = False
         saldo = 0
         totalusado = 0
         moratotal = 0
         While Not resultado.EOF
         Grid1.Rows = Grid1.Rows + 1
         Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(0)
         Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(1)
         Grid1.Cell(Grid1.Rows - 1, 3).text = resultado(2)
         Grid1.Cell(Grid1.Rows - 1, 4).text = resultado(3)
         Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(4), "$ ###,###,###")
         saldo = saldo + resultado(4)
         resultado.MoveNext
        Wend
    Else
       
    End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid1.Rows = Grid1.Rows + 1
    Grid1.Column(0).Locked = False
    Grid1.Column(1).Locked = False
    Grid1.Column(2).Locked = False
    Grid1.Column(3).Locked = False
    Grid1.Column(4).Locked = False
    Grid1.Column(5).Locked = False
        
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 5).Borders(cellEdgeTop) = cellThin
        Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 4).Merge
        
        
        Grid1.Cell(Grid1.Rows - 1, 3).text = "CUOTAS CANCELADAS"
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(saldo, "$ ###,###,###")
        If CDbl(interesmora.Caption) <> 0 Then
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 4).Merge
        
        Grid1.Cell(Grid1.Rows - 1, 3).text = "INTERES MORA"
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(interesmora.Caption, "$ ###,###,###")
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 4).Merge
        
        Grid1.Cell(Grid1.Rows - 1, 3).text = "TOTAL CANCELADO"
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(CDbl(montocuota.Caption) + CDbl(interesmora.Caption), "$ ###,###,###")
        End If
        For K = 1 To 10
        Grid1.Rows = Grid1.Rows + 1
    
        Next K
        
       
        Grid1.Range(Grid1.Rows - 3, 1, Grid1.Rows - 3, 5).Merge
        
        Grid1.Cell(Grid1.Rows - 3, 1).Alignment = cellCenterCenter
        Grid1.Cell(Grid1.Rows - 3, 1).Font.Bold = True
        Grid1.Cell(Grid1.Rows - 3, 1).Font.Size = 10
        Grid1.Cell(Grid1.Rows - 3, 1).text = "CANCELADO PROMOTORA PALGUIN     "
        
        
        
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 5).Merge
        
        Grid1.Cell(Grid1.Rows - 1, 1).Alignment = cellCenterCenter
        Grid1.Cell(Grid1.Rows - 1, 1).Font.Bold = True
        Grid1.Cell(Grid1.Rows - 1, 1).Font.Size = 10
        Grid1.Cell(Grid1.Rows - 1, 1).text = "FIRMA CAJERA"
        
        
        
        If repa.Caption = "COMPROBANTE DE REPACTACION" Then
        
        CARGAREPACTACION
        
        End If
        
        Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
   
        
        Grid1.AutoRedraw = True
        Grid1.Refresh
        
        FOLIO.Caption = folio1
        
        
        
    End Sub


Sub LEERCABEZACUOTAS(folio1)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim dia1 As String
        Dim mes1 As String
        Dim año1 As String
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT *  "
        csql.sql = csql.sql & "FROM sv_cuotas_pago_cabeza "
        csql.sql = csql.sql & "WHERE numero='" + folio1 + "' and rut='" + rut2.Caption + lbldv.Caption + "' "
        csql.Execute
        
    If csql.RowsAffected > 0 Then

         Set resultado = csql.OpenResultset
         total.Caption = Format(resultado(4), "###,###,###0")
         montocuota.Caption = Format(resultado(6), "###,###,###0")
         interesmora.Caption = Format(resultado(7), "###,###,###0")
         lblcajera.Caption = leerNombreCajera(resultado(5) + rut(resultado(5)))
         
         
         pivote.MaxLength = 2
         pivote.text = Format(resultado(3), "dd")
'         pivote.text = ceros(pivote)
         dia1 = pivote.text
         pivote.text = Format(resultado(3), "mm")
'         pivote.text = ceros(pivote)
         mes1 = pivote.text
         año1 = Format(resultado(3), "yyyy")
         dia.Caption = dia1
         mes.Caption = mes1
         año.Caption = año1
        If resultado(9) = "1" Then
        repa.Caption = "COMPROBANTE DE REPACTACION"
        Else
        repa.Caption = "COMPROBANTE PAGO DE CUOTAS"
        End If
        While Not resultado.EOF
        
         resultado.MoveNext
        Wend
    Else
       
    End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
        
        
        
        
    End Sub

Sub ELIMINARTODO(folio1)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim dia1 As String
        Dim mes1 As String
        Dim año1 As String
        Dim CAMPOS(5, 5) As String
        rebajacuotas
        
         
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 2) = "sv_cuotas_pago_cabeza"
        condicion = "numero='" + folio1 + "' and rut='" + rut2.Caption + lbldv.Caption + "' "
        op = 4
        sql.response = CAMPOS
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 2) = "sv_cuotas_pago_detalle"
        condicion = "numero='" + folio1 + "' and rut='" + rut2.Caption + lbldv.Caption + "' "
        op = 4
        sql.response = CAMPOS
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
         Call sql.sqlventas(op, condicion)
         
         If repa.Caption = "COMPROBANTE DE REPACTACION" Then
       
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 2) = "sv_cuotas_detalle"
        condicion = "numero='" + folio1 + "' and rut='" + rut2.Caption + lbldv.Caption + "' and tipo='VR' "
        op = 4
        sql.response = CAMPOS
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        
         End If
        
         
        
    End Sub



Private Sub Grid1_Click()
Dim TIPO As String
Dim NUMERO As String
Dim fecha As String
'
'        If 0 < GRID1.ActiveCell.Row And GRID1.ActiveCell.Row < GRID1.Rows Then
'            tipo = Left(GRID1.Cell(GRID1.ActiveCell.Row, 3).text, 2)
'            numero = Right(GRID1.Cell(GRID1.ActiveCell.Row, 4).text, 10)
'            Load PVentas
'            PVentas.dato1.text = tipo
'            PVentas.dato2.text = numero
'            PVentas.cargardeafuera
'            PVentas.Show vbModal
'        End If

'            TIPO = Left(Grid1.Cell(Grid1.ActiveCell.row, 3).text, 2)
'            NUMERO = Right(Grid1.Cell(Grid1.ActiveCell.row, 4).text, 10)
'            localAuditoria = empresaActiva
'            pivote.MaxLength = 10
'            pivote.text = NUMERO
'            NUMERO = ceros(pivote)
'            fecha = Grid1.Cell(Grid1.ActiveCell.row, 2).text
'            Load DetalleDocumento
'            DetalleDocumento.TIPO = TIPO
'            DetalleDocumento.NUMERO = NUMERO
'            DetalleDocumento.fechaAudit = Format(fecha, "yyyy-mm-dd")
'            rut_cliente = rut2.Caption & lbldv.Caption
'            DetalleDocumento.Show vbModal
   
End Sub
Sub Titulos()

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    
  
    
    
    
      
    Grid1.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    
    Grid1.PageSetup.Header = "PROMOTORA PALGUIN LTDA" & vbCrLf & "PUCON" & vbCrLf & "045-441349 ANEXO 350 "
    Grid1.PageSetup.HeaderAlignment = cellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = repa.Caption + " " + FOLIO.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CAJERA    :" + lblcajera.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FECHA    :" + dia.Caption + "-" + mes.Caption + "-" + año.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CLIENTE  :" + rut2.Caption + "-" + lbldv.Caption
    
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "NOMBRE   :" + lblnombre.Caption
    
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CREDITO AUTORIZADO   :" + creditoPAGOSTMP.lblCupo.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CREDITO UTILIZADO       :" + creditoPAGOSTMP.lblUtilizado.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CREDITO DISPONIBLE     :" + creditoPAGOSTMP.lblDisponible.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    Grid1.PageSetup.PrintFixedRow = True
    Grid1.PageSetup.BlackAndWhite = True
    
    Grid1.Range(0, 1, 0, 5).Borders(cellEdgeTop) = cellThick
    Grid1.Range(0, 1, 0, 5).Borders(cellEdgeBottom) = cellThick
    Grid1.Range(0, 1, 0, 5).Borders(cellEdgeRight) = cellThick
    Grid1.Range(0, 1, 0, 5).Borders(cellEdgeLeft) = cellThick
    Grid1.Range(0, 1, 0, 5).Borders(cellInsideHorizontal) = cellThick
    Grid1.Range(0, 1, 0, 5).Borders(cellInsideVertical) = cellThick
    
End Sub

Sub rebajacuotas()
  
        Dim CAMPOS(12, 3) As String
        Dim op As Integer
        Dim K As Integer
        Dim TIPO As String
        Dim NUMERO As String
        Dim fecha As String
        
        Set sql = New sqlventas.sqlventa
        For K = 1 To Grid1.Rows - 1
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
                 
         TIPO = Grid1.Cell(K, 1).text
        If TIPO = "VC" Or TIPO = "FV" Or TIPO = "BV" Or TIPO = "CA" Or TIPO = "VR" Then
         NUMERO = Grid1.Cell(K, 2).text
         cuota = Mid(Grid1.Cell(K, 3).text, 1, 2)
         fecha = Format(Grid1.Cell(K, 4).text, "yyyy-mm-dd")
        
        csql.sql = "update sv_cuotas_detalle set abono=abono-'" & CDbl(Grid1.Cell(K, 5).text) & "', vencimientoactual='" & Format(fecha, "yyyy-mm-dd") & "' "
        csql.sql = csql.sql & "WHERE tipo='" & TIPO & "' and numero='" & NUMERO & "' and rut='" & rut2.Caption & lbldv.Caption & "' and numerocuota='" & CDbl(cuota) & "' "
        csql.Execute
            Call sincronizadatos(csql.sql, ventas)
        csql.Close
         End If
        
        
        Next K
       
        
End Sub

Sub CARGAREPACTACION()

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT cantidadcuotas,montocuota,montoventa,vencimientoactual  "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle "
        csql.sql = csql.sql & "WHERE numero='" + creditoPAGOSTMP.FOLIO.Caption + "' and rut='" + rut2.Caption + lbldv.Caption + "' and tipo='VR' and numerocuota='1' "
        csql.Execute
        
    If csql.RowsAffected > 0 Then

         Set resultado = csql.OpenResultset
        While Not resultado.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Rows = Grid1.Rows + 1
        
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 5).Merge
        Grid1.Cell(Grid1.Rows - 1, 1).Alignment = cellLeftCenter
        Grid1.Cell(Grid1.Rows - 1, 1).text = "FIRMA :------------------------------------"
        Grid1.Rows = Grid1.Rows + 1
        
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 5).Merge
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = "Yo " + lblnombre.Caption
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 5).Merge
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = "CI." + rut2.Caption + "-" + lbldv.Caption + " autorizo segun contrato PALGUIN LTDA "
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 5).Merge
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = "Cargar a mi cuenta " & resultado(0) & " cuotas de " + Format(resultado(1), "$ ###,###,###")
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 5).Merge
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = "primer vencimiento :" + Format(resultado(3), "dd-mm-yyyy")
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 5).Merge
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTAL CREDITO :" + Format(resultado(2), "$ ###,###,##0")
        Grid1.Rows = Grid1.Rows + 1
        
        Rem Grid1.Cell(Grid1.Rows - 1, 1).text = "PIE :" + Format(CDbl(montototalventa) - CDbl(montocredito), "$ ###,###,##0")
        
         resultado.MoveNext
        Wend
   
       
    End If
      
        
        
    End Sub

