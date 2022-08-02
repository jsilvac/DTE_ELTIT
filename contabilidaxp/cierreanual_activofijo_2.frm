VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form activo04 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Anual"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6525
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   3015
      Left            =   7245
      TabIndex        =   3
      Top             =   135
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5318
      BackColor       =   16744576
      Caption         =   "REGISTRO EN PROCESO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   900
      Left            =   120
      TabIndex        =   1
      Top             =   2250
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1588
      BackColor       =   16744576
      Caption         =   "PROCESO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   375
         Left            =   135
         TabIndex        =   2
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3201
      BackColor       =   16744576
      Caption         =   "CIERRE ANUAL"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin CoolButtons.cool_Button command1 
         Height          =   495
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "COMIENZA ACTUALIZACION"
      End
      Begin VB.Label LBLCIERRE 
         BackStyle       =   0  'Transparent
         Caption         =   "2007 A 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1395
         TabIndex        =   7
         Top             =   270
         Width           =   3885
      End
      Begin VB.Label actualiza 
         BackColor       =   &H00FF8080&
         Height          =   465
         Left            =   1350
         TabIndex        =   6
         Top             =   855
         Width           =   3750
      End
   End
   Begin VB.PictureBox CmdFavoritos 
      BackColor       =   &H0000FF00&
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
      Left            =   8520
      ScaleHeight     =   195
      ScaleWidth      =   2715
      TabIndex        =   8
      Top             =   3480
      Width           =   2775
   End
End
Attribute VB_Name = "activo04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim debe(12) As Double
Dim haber(12) As Double
Dim AÑOACTUAL As String
Dim AÑOSIGUIENTE As String
Dim saldocuenta As Double
Dim dedonde As Integer
Dim saldobalance As Double

Private Sub CmdFavoritos_Click()
    Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Private Sub Command1_Click()

dedonde = 1
creamayor
 


no:
End Sub
Sub creamayor()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim resultados3 As rdoResultset
    Dim cSql3 As New rdoQuery
        
        Set csql2.ActiveConnection = contadb
        
        csql2.sql = "insert ignore INTO " + clientesistema + "conta" + empresaactiva + ".activo_fijo_nuevo"
        csql2.sql = csql2.sql & "(  codigo,  nombre,  familia,  fechapuestaenmarcha,  factura,  proveedor,  vidautil,  valorcompra,  depreciacion,  valorreal,  vidausada,  correcionmonetaria,  crcc,  año)"
        csql2.sql = csql2.sql & " select   codigo,  nombre,  familia,  fechapuestaenmarcha,  factura,  proveedor,  vidautil,  valorcompra,  depreciacion,  valorreal,  vidausada,  correcionmonetaria,  crcc, '" & AÑOSIGUIENTE & "' "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta" + empresaactiva + ".activo_fijo_nuevo where año='" + Format(fechasistema, "yyyy") + "' and fechaventa='0000-00-00'   "
        
        csql2.Execute
        Call sincronizadatos(csql2.sql, contadb, "")
        csql2.Close
         
        MsgBox "CIERRE ACTIVO FIJO FINALIZADO", vbExclamation, "ATENCION"
        Unload Me

End Sub
Sub cambiamayor(codigo, año, saldodebe, saldohaber)
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim resultados3 As rdoResultset
    Dim cSql3 As New rdoQuery
       
       Set cSql3.ActiveConnection = contadb
        cSql3.sql = "update saldosdelmayor set debeanterior='" & saldodebe & "',haberanterior='" & saldohaber & "' "
        cSql3.sql = cSql3.sql + " where codigo='" + codigo + "' and año='" + año + "' "
        cSql3.Execute
        Call sincronizadatos(cSql3.sql, contadb, "")
        
        cSql3.Close
        

End Sub
Sub cambiactacte(tipo, rut, año, saldodebe, saldohaber)
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim resultados3 As rdoResultset
    Dim cSql3 As New rdoQuery
       
       Set cSql3.ActiveConnection = contadb
        cSql3.sql = "update saldosctacte set debeanterior='" + saldodebe + "',haberanterior='" + saldohaber + "' "
        cSql3.sql = cSql3.sql + " where tipo='" + tipo + "' and rut='" + rut + "' and año='" + año + "' "
        cSql3.Execute
        Call sincronizadatos(cSql3.sql, contadb, "")
        
        cSql3.Close
        
End Sub


Sub creactacte()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim resultados3 As rdoResultset
    Dim cSql3 As New rdoQuery
        
        Set csql2.ActiveConnection = contadb
        
        csql2.sql = "INSERT INTO " & clientesistema & "conta" & empresaactiva & ".cuentascorrientes (tipo,año,rut,nombre,direccion,comuna,ciudad,giro,fono,fax,celular,email,contacto) "
        csql2.sql = csql2.sql + " SELECT tipo,'" + AÑOSIGUIENTE + "',rut,nombre,direccion,comuna,ciudad,giro,fono,fax,celular,email,contacto "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta" & empresaactiva & ".cuentascorrientes as cc where año ='" & Format(fechasistema, "yyyy") & "' "
        csql2.sql = csql2.sql + "on duplicate key update año='" + AÑOSIGUIENTE + "'"
        csql2.Execute
        Call sincronizadatos(csql2.sql, contadb, "")
        csql2.Close
        
       Set cSql3.ActiveConnection = contadb
        cSql3.sql = "INSERT INTO " & clientesistema & "conta" & empresaactiva & ".saldosctacte (tipo,año,rut) "
        cSql3.sql = cSql3.sql + " SELECT tipo,'" + AÑOSIGUIENTE + "',rut "
        cSql3.sql = cSql3.sql + "FROM " & clientesistema & "conta" & empresaactiva & ".saldosctacte "
        cSql3.sql = cSql3.sql + "on duplicate key update año='" + AÑOSIGUIENTE + "' "
        cSql3.Execute
        Call sincronizadatos(cSql3.sql, contadb, "")
         cSql3.Close
        

End Sub
Sub creacrcc()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim resultados3 As rdoResultset
    Dim cSql3 As New rdoQuery
        
        Set csql2.ActiveConnection = contadb
        
        csql2.sql = "INSERT INTO " & clientesistema & "conta" & empresaactiva & ".centrosdecosto (codigo,año,nombre) "
        csql2.sql = csql2.sql + " SELECT codigo,'" + AÑOSIGUIENTE + "',nombre "
        
        
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta" & empresaactiva & ".centrosdecosto "
        csql2.sql = csql2.sql + "on duplicate key update año='" + AÑOSIGUIENTE + "' "
        csql2.Execute
        Call sincronizadatos(csql2.sql, contadb, "")
        csql2.Close
        
        
       
       Set cSql3.ActiveConnection = contadb
        cSql3.sql = "INSERT INTO " & clientesistema & "conta" & empresaactiva & ".saldoscentrosdecosto (codigo,año,cuenta) "
        cSql3.sql = cSql3.sql + " SELECT codigo,'" + AÑOSIGUIENTE + "',cuenta "
        cSql3.sql = cSql3.sql + "FROM " & clientesistema & "conta" & empresaactiva & ".saldoscentrosdecosto "
        cSql3.sql = cSql3.sql + "on duplicate key update año='" + AÑOSIGUIENTE + "' "
        cSql3.Execute
        Call sincronizadatos(cSql3.sql, contadb, "")
        cSql3.Close
        
        

End Sub

Sub traspasamayor()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim totaldebe As Double
    Dim totalhaber As Double
    Dim LINEAS As Double
    

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT sm.codigo,sm.año,cm.tipo "
        csql2.sql = csql2.sql + "FROM saldosdelmayor as sm left join cuentasdelmayor as cm on sm.codigo=cm.codigo and sm.año=cm.año where sm.año='" + Format(fechasistema, "yyyy") + "' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        saldobalance = 0
        saldodebe = "0"
        saldohaber = "0"
        lin = 0
        Barra.Min = 0.01
        Barra.Max = csql2.RowsAffected + 4
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        saldodebe = "0"
        saldohaber = "0"
        
        Call LEERSALDOS(resultados2(0))
        If saldocuenta < 0 Then
        saldohaber = saldocuenta * -1
        Else
        saldodebe = saldocuenta
        End If
        If resultados2(2) > "2" And Mid(resultados2(0), 5, 4) <> "0000" Then
        totaldebe = totaldebe + CDbl(saldodebe)
        totalhaber = totalhaber + CDbl(saldohaber)
        saldodebe = "0"
        saldohaber = "0"
        
        End If
        
        ' ariel quita stop por insrtruccion de granadino
        'If resultados2(0) = "11500001" Then Stop
        
        Call cambiamayor(resultados2(0), AÑOSIGUIENTE, saldodebe, saldohaber)
        
        Rem
        Barra.Value = LINEAS
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
     saldobalance = totaldebe - totalhaber
        
     
     If saldobalance < 0 Then
        saldohaber = 0
        saldodebe = 0
        Call LEERSALDOS(cuentaganancia)
        If saldocuenta < 0 Then
        saldohaber = saldocuenta * -1
        Else
        saldodebe = saldocuenta
        End If
        Call cambiamayor(cuentaganancia, AÑOSIGUIENTE, saldodebe, saldohaber + saldobalance * -1)
     Else
        saldohaber = 0
        saldodebe = 0
        
        Call LEERSALDOS(cuentaganancia)
        If saldocuenta < 0 Then
        saldohaber = saldocuenta * -1
        Else
        saldodebe = saldocuenta
        End If
       
       Call cambiamayor(cuentaperdida, AÑOSIGUIENTE, saldobalance + saldodebe, saldohaber)
     End If
     
      
  Barra.Visible = False
  Unload Me

End Sub
Sub traspasactacte()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim LINEAS As Double

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT sa.tipo,sa.rut,sa.año "
        csql2.sql = csql2.sql + "FROM saldosctacte as sa inner join cuentascorrientes as cc on cc.rut=sa.rut and cc.tipo=sa.tipo and cc.año=sa.año "
        csql2.sql = csql2.sql + "where cc.año='" + Format(fechasistema, "yyyy") + "' "
        If empresaactiva <> "28" Then
        csql2.sql = csql2.sql + "and cc.tipo<>'11200027' and mid(cc.tipo,1,4)<>'1135' "
        
        End If
        
        csql2.sql = csql2.sql + "order by tipo,rut "
        csql2.Execute
        saldodebe = "0"
        saldohaber = "0"
        lin = 0
        Barra.Value = 1
        Barra.Min = 0.01
        Barra.Max = csql2.RowsAffected + 4
        LINEAS = 0
        
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        saldodebe = "0"
        saldohaber = "0"
        saldocuenta = leersaldoctacte(resultados2(0), resultados2(1), Format(fechasistema, "YYYY-MM-DD"))
        
        If saldocuenta < 0 Then
        saldohaber = saldocuenta * -1
        Else
        saldodebe = saldocuenta
        End If
       Rem  If resultados2(1) = "0158643375" Then Stop
        Call cambiactacte(resultados2(0), resultados2(1), AÑOSIGUIENTE, saldodebe, saldohaber)
        Rem
        Barra.Value = LINEAS
        Barra.Refresh
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
      
  Barra.Visible = False
  

End Sub




Private Sub Form_Load()
         
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
     AÑOACTUAL = Format(fechasistema, "YYYY")
     AÑOSIGUIENTE = Format(fechasistema, "YYYY") + 1
     
     LBLCIERRE.Caption = AÑOACTUAL + " AL " + AÑOSIGUIENTE
    Call CENTRAR(Me)
     
End Sub


Sub LEERSALDOS(cuenta)

Dim resultados3 As rdoResultset
    
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mesante As Integer
    Dim sumade As Double
    Dim sumaha As Double
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = "debe01"
    campos(5, 0) = "debe02"
    campos(6, 0) = "debe03"
    campos(7, 0) = "debe04"
    campos(8, 0) = "debe05"
    campos(9, 0) = "debe06"
    campos(10, 0) = "debe07"
    campos(11, 0) = "debe08"
    campos(12, 0) = "debe09"
    campos(13, 0) = "debe10"
    campos(14, 0) = "debe11"
    campos(15, 0) = "debe12"
    campos(16, 0) = "haber01"
    campos(17, 0) = "haber02"
    campos(18, 0) = "haber03"
    campos(19, 0) = "haber04"
    campos(20, 0) = "haber05"
    campos(21, 0) = "haber06"
    campos(22, 0) = "haber07"
    campos(23, 0) = "haber08"
    campos(24, 0) = "haber09"
    campos(25, 0) = "HABER10"
    campos(26, 0) = "HABER11"
    campos(27, 0) = "HABER12"
    campos(28, 0) = ""
    
    If dedonde = 1 Then condicion = "codigo=" + "'" + cuenta + "' and año='" + Mid(fechasistema, 7, 4) + "' order by codigo"
    If dedonde = 3 Then condicion = "codigo=" + "'" + cuenta + "' and año='" + Mid(fechasistema, 7, 4) + "' order by codigo"
    
    If dedonde = 1 Then campos(0, 2) = "saldosdelmayor"
    If dedonde = 3 Then campos(0, 2) = "saldoscentrosdecosto"
 
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop
    sumador = Val(sqlconta.response(2, 3)) - Val(sqlconta.response(3, 3))
   For k = 1 To 12
   sumade = sumade + CDbl(sqlconta.response(3 + k, 3))
   sumaha = sumaha + CDbl(sqlconta.response(15 + k, 3))
   
   
   Next k
   saldocuenta = sumador + sumade - sumaha
   
End Sub

