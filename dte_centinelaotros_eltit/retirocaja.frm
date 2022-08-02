VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form RetirosCaja 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pantalla de Egresos de Caja"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp frmGlosa 
      Height          =   4815
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
      BackColor       =   16744576
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   0
      Begin VB.CommandButton CmdEliminar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "F2 ELIMINAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4200
         Width           =   2775
      End
      Begin VB.CommandButton CMDIMPRIMIR 
         BackColor       =   &H00FFC0C0&
         Caption         =   "F1 IMPRIMIR "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4200
         Width           =   2775
      End
      Begin FlexCell.Grid Grid1 
         Height          =   30
         Left            =   7080
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   53
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin XPFrame.FrameXp ingresaprecio 
         Height          =   3735
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   6588
         BackColor       =   16761024
         Caption         =   ""
         BackColor       =   16761024
         ForeColor       =   65535
         ColorBarraArriba=   4194304
         ColorBarraAbajo =   16761024
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
         Begin VB.TextBox DIATXT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   2415
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   13
            Tag             =   "proveedor"
            Top             =   645
            Width           =   450
         End
         Begin VB.TextBox MESTXT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   2910
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   12
            Tag             =   "proveedor"
            Top             =   645
            Width           =   450
         End
         Begin VB.TextBox AÑOTXT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   3405
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   11
            Tag             =   "proveedor"
            Top             =   645
            Width           =   720
         End
         Begin VB.TextBox CAJATXT 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            MaxLength       =   2
            TabIndex        =   0
            ToolTipText     =   "CAJA AL CUAL SE LE HARA EL RETIRO"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox folio_retiro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6240
            MaxLength       =   10
            TabIndex        =   5
            ToolTipText     =   "FOLIO DEL RETIRO"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox rut2 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2400
            MaxLength       =   9
            TabIndex        =   2
            Top             =   1815
            Width           =   1530
         End
         Begin VB.TextBox nombre_retirador 
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
            Height          =   420
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2400
            Width           =   6100
         End
         Begin VB.TextBox rut_cajera 
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
            Height          =   420
            Left            =   2400
            MaxLength       =   9
            TabIndex        =   1
            ToolTipText     =   "Rut Del Cajero"
            Top             =   1200
            Width           =   1530
         End
         Begin VB.TextBox nombre_cajera 
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
            Height          =   420
            Left            =   4440
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   10
            Top             =   1200
            Width           =   4095
         End
         Begin VB.TextBox monto_retiro 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            TabIndex        =   3
            Top             =   3120
            Width           =   2850
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "              FECHA"
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
            Left            =   2280
            TabIndex        =   23
            Top             =   240
            Width           =   1920
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  CAJA  "
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
            Left            =   135
            TabIndex        =   21
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "              FOLIO RETIRO"
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
            TabIndex        =   20
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lbldv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   3960
            TabIndex        =   19
            Top             =   1815
            Width           =   285
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "RUT RETIRA"
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
            TabIndex        =   18
            Top             =   1800
            Width           =   1920
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "NOMBRE"
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
            TabIndex        =   17
            Top             =   2400
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "CAJERO (A)"
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
            TabIndex        =   16
            Top             =   1200
            Width           =   1920
         End
         Begin VB.Label lbldvcajera 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   3960
            TabIndex        =   15
            Top             =   1215
            Width           =   285
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
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
            Index           =   3
            Left            =   240
            TabIndex        =   14
            Top             =   3120
            Width           =   1320
         End
      End
      Begin VB.Label nombrelocal 
         BackColor       =   &H0080FF80&
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
         Left            =   840
         TabIndex        =   22
         Top             =   720
         Width           =   6015
      End
   End
End
Attribute VB_Name = "RetirosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private existe As Boolean


Private Sub CmdEliminar_Click()
If CAJATXT = Empty Or folio_retiro = Empty Then Exit Sub
If Verifica_Permiso(Me.Caption, "elimina") = True Then

           If MsgBox("REALMENTE DESEA ELIMINAR COMPROBANTE", vbYesNo, "ATENCION") = vbYes Then
            Call eliminarretiro(CAJATXT.text, AÑOTXT.text & "-" & MESTXT.text & "-" & DIATXT.text, folio_retiro.text, rut_cajera.text & lbldvcajera.Caption)
            Call limpiapan
        End If


End If
End Sub

Private Sub CMDIMPRIMIR_Click()
If nombre_retirador.text <> "" Then
    Call grabarretiro(empresaActiva, CAJATXT.text, rut_cajera.text & lbldvcajera.Caption, AÑOTXT.text & "-" & MESTXT.text & "-" & DIATXT.text, Replace(monto_retiro.text, ".", ""), rut2.text & lbldv.Caption, nombre_retirador.text, existe)
    Call ImprimeTicketRetiro(rut_cajera.text & lbldvcajera.Caption, nombre_cajera.text, Format(fechasistema, "dd-mm-yyyy"), Time, folio_retiro.text, monto_retiro.text, Me.CAJATXT, nombre_retirador.text)

MsgBox "RETIRO POR : $" & monto_retiro & vbCr & _
" PARA LA CAJERA " & nombre_cajera & vbCr & _
" RETIRADO POR: " & nombre_retirador, vbInformation, "SE GRABO EL RETIRO DE LA CAJA " & CAJATXT & " Nº " & folio_retiro & " "
End If
Call limpiapan
End Sub



Private Sub folio_retiro_GotFocus()
Call selecciona(folio_retiro)
End Sub

Private Sub folio_retiro_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And folio_retiro.text <> "" Then
    folio_retiro.text = ceros(folio_retiro)
    If leerretiro(folio_retiro.text) = True Then
        existe = True
        CMDIMPRIMIR.SetFocus
    Else
        existe = False
        rut2.SetFocus
    End If
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            Unload Me
        Case vbKeyF1
            If monto_retiro.text <> "" Then
                CMDIMPRIMIR_Click
            Else
                MsgBox "NO PUEDE IMPRIMIR SI NO A INGRESADO UN MONTO ", vbCritical, "ATENCION"
            End If
        Case vbKeyF2
            If existe = True Then
            End If
    End Select
End Sub
Private Sub Form_Load()
    DIATXT.text = Format(fechasistema, "dd")
    MESTXT.text = Format(fechasistema, "mm")
    AÑOTXT.text = Format(fechasistema, "yyyy")
    frmGlosa.Caption = "VALE RETIRO " & leerNombreEmpresa(empresaActiva)
End Sub
Sub grabarretiro(localretiro, cajaretiro, cajeraretiro, fecharetiro, montoretiro, rutretiro, nombreretiro, existe As Boolean)
Dim resultados As rdoResultset
Dim cSql As New rdoQuery
Set cSql.ActiveConnection = ventasRubro
If existe = False Then
cSql.sql = "insert into sv_retirosdecaja_" & empresaActiva & "( `local`, `caja`, `cajera`, `folio`, `fecha`, `hora`, `monto`, `rutretiradopor` ) values ("
cSql.sql = cSql.sql & "'" & localretiro & "','" & cajaretiro & "','" & cajeraretiro & "','" & folio_retiro.text & "','" & fecharetiro & "','" & Time & "','" & montoretiro & "','" & rutretiro & "' )"
cSql.Execute
Call sincronizadatos(cSql.sql, ventasRubro)
Set resultados = Nothing
End If
    
    
 If existe = True Then
 '    Call IMPRIMEPREretiro(folio_retiro.text, GRID1, localretiro, cajaretiro, cajeraretiro, Format(fecharetiro, "dd-mm-yyyy"), campos(5, 1), montoretiro, rutretiro, nombreretiro)
 End If
    End Sub
 
    Function leerultimofolioretiro(cajaconsulta) As String
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set cSql.ActiveConnection = ventasRubro
    cSql.sql = "select ifnull(max(folio)+1,1) "
    cSql.sql = cSql.sql & "from sv_retirosdecaja_" & empresaActiva & " "
    cSql.sql = cSql.sql & "where caja='" & cajaconsulta & "' and local='" & empresaActiva & "' "
    cSql.Execute
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        leerultimofolioretiro = String(10 - Len(resultados(0)), "0") & resultados(0)
    Else
        leerultimofolioretiro = "0000000001"
    End If
    cSql.Close
    Set cSql = Nothing
    Set resultados = Nothing
End Function
 Public Sub IMPRIMEPREretiro(numeroretiro, lista As Grid, localretiro, cajaretiro, cajeraretiro, fecharetiro, horaretiro, montoretiro, rutretiro, nombreretiro)
    Dim tabla As String
    Dim i As Long
    Dim CODIGO As String
    Dim cantidad As String
    Dim precio As String
    Dim total As String
    Dim totalPreventa As Double
    Dim cadena As String
    Dim p As Printer
    Dim numfic As Integer
    Dim ubicacionfisica As String
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Dim despacho As String
    Dim Descuento As String
    Dim tipoempaque As String
    Dim Impresora As String
    Impresora = "0"
    'Call ConectarControlData(PuntoVenta.rollo, servidor, baseventas &empresaactiva, usuario, password, tabla)
    lista.Rows = 1
    lista.Cols = 6
    lista.AutoRedraw = False
    
    lista.PageSetup.HeaderMargin = 1.25
    lista.PageSetup.TopMargin = 1.25
    lista.PageSetup.FooterMargin = 0.5
    lista.PageSetup.BottomMargin = 0.5
    lista.PageSetup.LeftMargin = 0.5
    lista.PageSetup.RightMargin = 0.5
    
    lista.Column(0).Width = 0
    lista.Column(1).Width = 90
    lista.Column(2).Width = 35
    lista.Column(3).Width = 60
    lista.Column(4).Width = 65
    
    'lista.Column(0).Width = 0
    'lista.Column(1).Width = 50
    'lista.Column(2).Width = 30
    'lista.Column(3).Width = 30
    'lista.Column(4).Width = 30
    
   
 
            
        lista.AddItem leerNombreEmpresa(localretiro), True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        
        lista.AddItem leerNombreRubro(rubro), True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        
        lista.AddItem "", True
        
        lista.AddItem "CAJA :  " & cajaretiro & "   NRO RETIRO:  " & numeroretiro
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        lista.AddItem "", True
        lista.AddItem "FECHA: " & Format(fecharetiro, "dd/mm/yyyy") & "     HORA: " & Format(horaretiro, "hh:mm:ss")
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        'lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        lista.AddItem "CAJERO(A): " & leerNombreCajera(cajeraretiro), True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.AddItem "RETIRADO POR: " & Format(Mid(rutretiro, 1, 9), "###,###,###") & "-" & Mid(rutretiro, 9, 1), True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.AddItem "" & nombreretiro, True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        
        
        lista.AddItem "", True
        
        lista.AddItem "==========================================", True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.AddItem "", True
        cadena = Format(montoretiro, "$ ###,###,##0")
        cadena = String(15 - Len(cadena), " ") & cadena
        cadena = "                   TOTAL: " & cadena
        lista.AddItem cadena, True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        'tipopagofinal
        lista.AddItem "", True
        lista.AddItem "==========================================", True
        lista.AddItem "", True
        lista.AddItem "", True
        lista.AddItem "", True
        lista.AddItem "", True
        
     
    Set cSql = Nothing
    cSql.Close
    Set resultados = Nothing
    
    lista.AutoRedraw = True
    lista.Refresh
    
    
    
    'For i = 0 To Printers.Count
    '    If UCase(Printers(i).DeviceName) = "SRP350 PARTIAL CUT" Then
    '        lista.PageSetup.PrinterName = Printers(i).DeviceName
    '        Exit For
    '    End If
    'Next i
    
    For i = 1 To lista.PageSetup.PaperSizes.Count
        If UCase(lista.PageSetup.PaperSizes.Item(i).PaperName) = "A4 LENGTH" Then
            lista.PageSetup.PaperSize = lista.PageSetup.PaperSizes.Item(i).Kind
            Exit For
        End If
    Next i
    ''''''''''''''''''
    numfic = FreeFile
    Close numfic
    If Impresora = 0 Then
    Open "impresion.txt" For Output As #numfic
    End If
    If Impresora = 1 Then
    Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #numfic
    End If
    If Impresora = 2 Then
    Open "LPT1" For Output As #numfic
    End If
    
    ''''''''''''''''''

    '''''''''''''''''''''
    'EMPAQUE
    '''''''''''''''''''''
    lista.Cell(lista.Rows - 3, 1).text = "  RETIRO"
    Print #numfic, Chr$(27); Chr$(64) '
    For i = 1 To lista.Rows - 4
    If lista.Cell(i, 1).text <> "  RETIRO" Then
        If lista.Cell(i, 0).text = "01" Or lista.Cell(i, 0).text = "" Then
            Print #numfic, Mid(lista.Cell(i, 1).text, 1, 42)
        End If
    Else
            Print #numfic, Chr(29); Chr(33); Chr(33); lista.Cell(i, 1).text
            Print #numfic, Chr$(27); Chr$(64)
        End If
    Next i
    Print #numfic, Chr(29); Chr(33); Chr(33); lista.Cell(lista.Rows - 3, 1).text
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic, Chr(27); "i"
    '''''''''''''''''''''
    'EMPAQUE
    '''''''''''''''''''''
    

    '''''''''''''''''''''
    'PRE-VENTA
    '''''''''''''''''''''
    Print #numfic, Chr$(27); Chr$(64)
    Print #numfic, Chr(29); Chr(33); Chr(33); "  RETIRO"
    Print #numfic, Chr$(27) & Chr$(64)
    Print #numfic, Chr$(29) & Chr$(104) & Chr(100)
    Print #numfic, Chr$(29) & Chr$(119) & Chr(2)
    Print #numfic, Chr$(29) & Chr$(72) & Chr(50)
    Print #numfic, Chr$(29) & Chr$(107) & Chr(4) & cajaretiro & numeroretiro & Chr(0)
    Print #numfic, Chr$(29) & Chr$(33) & Chr(12) & "TOTAL $ "; montoretiro
    Print #numfic, Chr$(10) & Chr$(10) & Chr(10) & Chr(10)
    Print #numfic, Chr(27); "i"
    '''''''''''''''''''''
    'PRE-VENTA
    '''''''''''''''''''''
    Close #numfic
    If Impresora = 0 Then Shell "notepad impresion.txt"
'  lista.PrintPreview
End Sub

Private Sub CAJATXT_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        CAJATXT.text = ceros(CAJATXT)
        
        folio_retiro.text = leerultimofolioretiro(CAJATXT.text)
        rut_cajera.SetFocus
    End If
End Sub
Private Sub DIATXT_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
DIATXT.text = ceros(DIATXT)
MESTXT.SetFocus
 Call esfecha(DIATXT, MESTXT, AÑOTXT, "dd")
End If
End Sub

Private Sub mestXT_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
MESTXT.text = ceros(MESTXT)
AÑOTXT.SetFocus
  Call esfecha(DIATXT, MESTXT, AÑOTXT, "mm")
End If

End Sub
Private Sub AÑOTXT_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
AÑOTXT.text = ceros(AÑOTXT)
 
 Call esfecha(DIATXT, MESTXT, AÑOTXT, "yyyy")
End If

End Sub
 
 

Private Sub monto_retiro_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And monto_retiro.text <> "" Then
        If CDbl(monto_retiro.text) > 0 Then
            monto_retiro.text = Format(monto_retiro.text, "###,###,##0")
            CMDIMPRIMIR.SetFocus
        Else
            monto_retiro.SetFocus
        End If
    End If
End Sub

Private Sub nombre_retirador_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And nombre_retirador.text <> "" Then
        monto_retiro.SetFocus
    End If
End Sub

Private Sub rut_cajera_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And rut_cajera.text <> "" Then
        rut_cajera.text = ceros(rut_cajera)
        lbldvcajera.Caption = rut(rut_cajera.text)
        If leerNombreCajera(rut_cajera.text + lbldvcajera.Caption) <> "" Then
            nombre_cajera = leerNombreCajera(rut_cajera.text + lbldvcajera.Caption)
           rut2.SetFocus
        Else
        Call selecciona(rut_cajera)
        End If
        
        End If
End Sub

Private Sub rut2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And rut2.text <> "" Then
        rut2.text = ceros(rut2)
        lbldv.Caption = rut(rut2.text)
        If leerNombreCajera(rut2.text + lbldv.Caption) <> "" Then
        nombre_retirador.text = leerNombreCajera(rut2.text + lbldv.Caption)
        monto_retiro.SetFocus
        Else
        rut2.SetFocus
        End If
        End If
End Sub
Function leerretiro(numeroretiro) As Boolean
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set cSql.ActiveConnection = ventasRubro
    cSql.sql = "select local,caja,cajera,folio,fecha,hora,monto,rutretiradopor "
    cSql.sql = cSql.sql & "from sv_retirosdecaja_" & empresaActiva
    cSql.sql = cSql.sql & " where local='" & empresaActiva & "' and caja='" & CAJATXT & "' and folio='" & numeroretiro & "' "
    cSql.Execute
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        leerretiro = True
        CAJATXT.text = resultados(1)
        DIATXT.text = Format(resultados(4), "dd")
        MESTXT.text = Format(resultados(4), "mm")
        AÑOTXT.text = Format(resultados(4), "yyyy")
        rut_cajera.text = Mid(resultados(2), 1, 9)
        lbldvcajera.Caption = Mid(resultados(2), 10, 1)
        nombre_cajera.text = leerNombreCajera(resultados(2))
        rut2.text = Mid(resultados(7), 1, 9)
        lbldv.Caption = Right(resultados(7), 1)
        nombre_retirador.text = leerNombreCajera(resultados(7))
        monto_retiro.text = Format(resultados(6), "###,###,##0")
        
    End If
    
End Function

Sub limpiapan()
    monto_retiro.text = ""
    nombre_retirador.text = ""
    rut2.text = ""
    lbldv.Caption = ""
    nombre_cajera.text = ""
    lbldvcajera.Caption = ""
    rut_cajera.text = ""
    folio_retiro.text = ""
    AÑOTXT.text = ""
    MESTXT.text = ""
    DIATXT.text = ""
    CAJATXT.text = ""
     CAJATXT.text = ""
    DIATXT.text = Format(fechasistema, "dd")
    MESTXT.text = Format(fechasistema, "mm")
    AÑOTXT.text = Format(fechasistema, "yyyy")
    folio_retiro.text = leerultimofolioretiro(CAJATXT.text)
    rut_cajera.text = Empty
    lbldvcajera.Caption = Empty
    nombre_cajera.text = Empty
    CAJATXT.SetFocus
End Sub

Sub ImprimeTicketRetiro(cajera, nombrecajera, fecha, HORA, NUMERO, MONTO, caja, retiradopor)
 Dim numfic As Integer
 Dim K As Double
 Dim cSql As New rdoQuery
 numfic = 20
 Close numfic
 Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #numfic
 Print #numfic, Chr$(27); Chr$(64)
 Print #numfic, Chr$(27)
 For K = 1 To 3
   
                Print #numfic, "  " & leerNombreEmpresa(empresaActiva)
                Print #numfic, "  " & leerDireccionEmpresa(empresaActiva)
                Print #numfic, "      VALE RETIRO DE CAJA"
  If K = 1 Then Print #numfic, "       Copia Supervisor"
  If K = 2 Then Print #numfic, "         Copia Cajero"
  If K = 3 Then Print #numfic, "        Copia Tesoreria"
                Print #numfic, "  _______________________________________ "
                Print #numfic, "          " & "NUMERO :" & NUMERO
                Print #numfic, "          " & "CAJA   :" & caja
                Print #numfic, "          " & "FECHA  :" & Format(fecha, "dd/mm/yyyy")
                Print #numfic, "          " & "HORA   :" & HORA
                Print #numfic, "          " & "CAJERA :" & cajera
                Print #numfic, "          " & "        " & nombrecajera
                Print #numfic, "  _______________________________________ "
           
If K = 3 Then
                Print #numfic, "    "
                Print #numfic, "    $ 20.000 : $ ________________"
                Print #numfic, "    "
                Print #numfic, "    $ 10.000 : $ ________________"
                Print #numfic, "    "
                Print #numfic, "     $ 5.000 : $ ________________"
                Print #numfic, "    "
                Print #numfic, "         Total $ ________________"
                Print #numfic, "          "
                Print #numfic, "  _______________________________________ "
                Print #numfic, "          "
End If
                Print #numfic, "          "
                Print #numfic, "MONTO DEL RETIRO :" & Format(MONTO, "$ ###,###,###")
                Print #numfic, " "
                Print #numfic, "  _______________________________________ "
                Print #numfic, "        RETIRADO POR " & retiradopor
                Print #numfic, " "
                Print #numfic, " "
                Print #numfic, " "
                Print #numfic, " "
                Print #numfic, "    _______________     ________________"
                Print #numfic, "      FIRMA CAJERA      FIRMA JEFA.CAJA "
                Print #numfic, "  _______________________________________ "
                Print #numfic, " "
                Print #numfic, " "
                Print #numfic, " "; Chr(27)
                Print #numfic, Chr(27); "i"
Next K
                Close #numfic
    
End Sub
Sub eliminarretiro(cajaretiro, fecharetiro, numeroretiro, cajeraretiro)
    Dim i As Integer
    Dim columna As Integer
    'agrega modifica elimina
      
    
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "delete from " & "sv_retirosdecaja_" + empresaActiva
        cSql.sql = cSql.sql & " where local='" & empresaActiva & "' and caja='" & cajaretiro & "' and fecha='" & fecharetiro & "' and folio='" & numeroretiro & "' and cajera='" & cajeraretiro & "' "
        cSql.Execute
                
    End Sub

