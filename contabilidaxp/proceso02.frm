VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form proceso02 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza_Movimientos"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   11190
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   4440
      TabIndex        =   10
      Top             =   4200
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   280
         Width           =   1335
      End
   End
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
      Left            =   720
      TabIndex        =   1
      Top             =   2250
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1588
      BackColor       =   16744576
      Caption         =   "PROCESO MAYOR"
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
      Left            =   720
      TabIndex        =   0
      Top             =   135
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3201
      BackColor       =   16744576
      Caption         =   "ACTUALIZA TODAS LAS EMPRESAS"
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
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "COMIENZA ACTUALIZACION"
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   675
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   1191
         BackColor       =   16744576
         Caption         =   "LOCAL"
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
         Begin VB.ComboBox ComboLOCAL 
            Height          =   315
            Left            =   45
            TabIndex        =   8
            Top             =   270
            Width           =   5715
         End
      End
      Begin VB.Label actualiza 
         BackColor       =   &H00FF8080&
         Height          =   465
         Left            =   1350
         TabIndex        =   6
         Top             =   630
         Width           =   3750
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "DURACION APROXIMADA DE LA ACTUALIZACION 1 A 2 MINUTOS POR FAVOR ESPERAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   3360
      Width           =   8415
   End
End
Attribute VB_Name = "proceso02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim debe(12) As Double
Dim haber(12) As Double

 

Private Sub Command1_Click()
Dim empre As String
Dim o As Integer
  
        Rem limpiactacte
        If Mid(ComboLOCAL.text, 1, 2) = "99" Then
        For o = 0 To ComboLOCAL.ListCount - 1
        empre = Mid(ComboLOCAL.List(o), 1, 2)
        ComboLOCAL.text = empre
        ComboLOCAL.Refresh
        Call blanqueamayor(empre)
        Call blanqueacrcc(empre)
        Call leecuentas(empre)
        Next o
        End If
        
        If Mid(ComboLOCAL.text, 1, 2) <> "99" Then
        empre = Mid(ComboLOCAL.text, 1, 2)
        ComboLOCAL.text = empre
        ComboLOCAL.Refresh
        Call blanqueamayor(empre)
        Call blanqueacrcc(empre)
        Call leecuentas(empre)
        
        End If
        MsgBox "ACTULIZACION REALIZADA "
    
no:
End Sub

Sub limpiactacte()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    
    mes1 = 1
    mes2 = 12
    If mes1 < 10 Then mes1 = "0" + Mid(Str(mes1), 2, 1)
    If mes2 < 10 Then mes2 = "0" + Mid(Str(mes2), 2, 1)
    
    Call blanqueactacte
  

End Sub

Sub blanqueactacte()
Dim i As Integer
Dim pasada As Integer

pasada = 0

    For i = 1 To 12
    If i < 10 Then MES = "0" + Mid(Str(i), 2, 1) Else MES = Mid(Str(i), 2, 2)
    campos(0 + pasada, 0) = "debe" + MES
    campos(1 + pasada, 0) = "haber" + MES
    campos(0 + pasada, 1) = "0"
    campos(1 + pasada, 1) = "0"
    
    pasada = pasada + 2
    Next i
    campos(pasada, 0) = ""
    campos(0, 2) = "saldosctacte"
    condicion = "AÑO='" + Format(fechasistema, "YYYY") + "'"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
    End If
End Sub
Sub blanqueamayor(empre)
Dim i As Integer
Dim pasada As Integer

pasada = 0

    For i = 1 To 12
    If i < 10 Then MES = "0" + Mid(Str(i), 2, 1) Else MES = Mid(Str(i), 2, 2)
    campos(0 + pasada, 0) = "debe" + MES
    campos(1 + pasada, 0) = "haber" + MES
    campos(0 + pasada, 1) = "0"
    campos(1 + pasada, 1) = "0"
    
    pasada = pasada + 2
    Next i
    campos(pasada, 0) = ""
    campos(0, 2) = clientesistema + "conta" + empre + ".saldosdelmayor"
    condicion = "AÑO='" + Format(fechasistema, "YYYY") + "'"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    'If sqlconta.status = 4 Then Stop
End Sub
Sub blanqueacrcc(empre)
Dim i As Integer
Dim pasada As Integer

pasada = 0

    For i = 1 To 12
    If i < 10 Then MES = "0" + Mid(Str(i), 2, 1) Else MES = Mid(Str(i), 2, 2)
    campos(0 + pasada, 0) = "debe" + MES
    campos(1 + pasada, 0) = "haber" + MES
    campos(0 + pasada, 1) = "0"
    campos(1 + pasada, 1) = "0"
    
    pasada = pasada + 2
    Next i
    campos(pasada, 0) = ""
    campos(0, 2) = clientesistema + "conta" + empre + ".saldoscentrosdecosto"
    condicion = "AÑO='" + Format(fechasistema, "YYYY") + "'"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop
End Sub


Private Sub Form_Load()
    cargameses
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    LEErlocales
    
    
End Sub
Sub cargameses()

Dim i As Integer



End Sub
    Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigoempresa,nombre "
        csql.sql = csql.sql + "FROM maestroempresas where codigoempresa<'90' "
        csql.sql = csql.sql + "ORDER BY codigoempresa "
        csql.Execute
        ComboLOCAL.Clear
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            
        End If
        ComboLOCAL.AddItem ("99" + " " + "TODAS LAS EMPRESAS")
        ComboLOCAL.text = "99 TODAS LAS EMPRESAS "
        
        localfiltro = "99"
        
End Sub


Sub leecuentas(empre)
barra.Visible = True
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim fecha1 As String
    
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    
    mes1 = 1
    mes2 = 12
    If mes1 < 10 Then mes1 = "0" + Mid(Str(mes1), 2, 1)
    If mes2 < 10 Then mes2 = "0" + Mid(Str(mes2), 2, 1)
    
    fecha1 = Mid(fechasistema, 7, 4) & "-" & mes1 & "-" & "01"
    fecha2 = Mid(fechasistema, 7, 4) & "-" & mes2 & "-" & "31"
    

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo "
        csql2.sql = csql2.sql + "FROM " + clientesistema + "conta" + empre + ".cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        lin = 0
         barra.Min = 0.01
        barra.Max = csql2.RowsAffected + 20
        barra.Value = 1
      
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        barra.Value = barra.Value + 1
     
        If Mid(resultados2(0), 5, 4) <> "0000" Then Call LEERMOVIMIENTOS(resultados2(0), fecha1, fecha2, empre)
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
      Rem Call LEERMOVIMIENTOScuentascorrientes(fecha1, fecha2)
      Call LEERMOVIMIENTOSCRCC(fecha1, fecha2, empre)
  
  barra.Visible = False
List1.AddItem "PROCESO FINALIZADO CORRECTAMENTE"
Unload Me

End Sub
Sub LEERMOVIMIENTOS(cuenta, fecha1, fecha2, empre)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim COMANDO As String
    Dim MES As String
    Dim año As String
'        consulta = "SELECT bodega, DATE_FORMAT(fecha,'%m') as mes, tipo, codigo, descripcion, SUM(unidades) AS cantidad, SUM(total) AS total "
'    consulta = consulta & "FROM l_movimientos_detalle_" & empresaactiva & " "
'    Rem consulta = consulta & "WHERE fecha LIKE '" & Format(fechasistema, "yyyy") & "%' and codigo='0000078000834' "
'    consulta = consulta & "WHERE fecha LIKE '" & Format(fechasistema, "yyyy") & "%' "
'    consulta = consulta & "GROUP BY bodega, codigo, tipo, mes "
'    consulta = consulta & "ORDER BY codigo, mes, tipo ASC"

        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT mes,codigocuenta,sum(monto),dh "
 
        csql.sql = csql.sql + "FROM " + clientesistema + "conta" + empre + ".movimientoscontables where codigocuenta='" + cuenta + "' "
        csql.sql = csql.sql + "and fecha>='" & fecha1 & "' and fecha<='" & fecha2 & "' "
        
        csql.sql = csql.sql + "group by dh,mes"
        csql.Execute
        For k = 1 To 12
        debe(k) = 0
        haber(k) = 0
        Next k
        
        If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
            
            If resultados(3) = "H" Then haber(resultados(0)) = resultados(2)
            If resultados(3) = "D" Then debe(resultados(0)) = resultados(2)
             resultados.MoveNext
           
         Wend
         Call actualizacuentamayor(cuenta, empre)
          resultados.Close
            Set resultados = Nothing

        End If
        
End Sub

Sub actualizacuentamayor(CUENTAMAYOR, empre)
    Dim SUMAVALOR As Double
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debe01"
    campos(3, 0) = "debe02"
    campos(4, 0) = "debe03"
    campos(5, 0) = "debe04"
    campos(6, 0) = "debe05"
    campos(7, 0) = "debe06"
    campos(8, 0) = "debe07"
    campos(9, 0) = "debe08"
    campos(10, 0) = "debe09"
    campos(11, 0) = "debe10"
    campos(12, 0) = "debe11"
    campos(13, 0) = "debe12"
    campos(14, 0) = "haber01"
    campos(15, 0) = "haber02"
    campos(16, 0) = "haber03"
    campos(17, 0) = "haber04"
    campos(18, 0) = "haber05"
    campos(19, 0) = "haber06"
    campos(20, 0) = "haber07"
    campos(21, 0) = "haber08"
    campos(22, 0) = "haber09"
    campos(23, 0) = "haber10"
    campos(24, 0) = "haber11"
    campos(25, 0) = "haber12"
    campos(26, 0) = ""
    For k = 1 To 12
    
    campos(k + 1, 1) = debe(k)
    campos(k + 13, 1) = haber(k)
    Next k
    campos(0, 1) = CUENTAMAYOR
    campos(1, 1) = año
    condicion = "codigo=" + "'" + CUENTAMAYOR + "' and año ='" + año + "' order by codigo"
    
    campos(0, 2) = clientesistema + "conta" + empre + ".saldosdelmayor"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
Rem actualiza cuenta madre
    condicion = "codigo=" + "'" + Mid(CUENTAMAYOR, 1, 4) + "0000" + "' and año ='" + año + "' order by codigo"

    campos(0, 2) = clientesistema + "conta" + empre + ".saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)


  '  If sqlconta.status = 4 Then Stop


    campos(0, 1) = sqlconta.response(0, 3)
    campos(1, 1) = sqlconta.response(1, 3)
    For k = 1 To 12
    campos(k + 1, 1) = Str(sqlconta.response(k + 1, 3) + debe(k))
    campos(k + 13, 1) = Str(sqlconta.response(k + 13, 3) + haber(k))
    
    Next k
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
 '   If sqlconta.status = 4 Then Stop

Rem actualiza cuenta principal
    condicion = "codigo=" + "'" + Mid(CUENTAMAYOR, 1, 2) + "000000" + "' and año ='" + año + "' order by codigo"

    campos(0, 2) = clientesistema + "conta" + empre + ".saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

'
  '  If sqlconta.status = 4 Then Stop
    campos(0, 1) = sqlconta.response(0, 3)
    campos(1, 1) = sqlconta.response(1, 3)
    For k = 1 To 12
    campos(k + 1, 1) = Str(sqlconta.response(k + 1, 3) + debe(k))
    campos(k + 13, 1) = Str(sqlconta.response(k + 13, 3) + haber(k))
    
    Next k
    

    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop

    
End Sub

Sub LEERMOVIMIENTOScuentascorrientes(fecha1, fecha2)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim COMANDO As String
    Dim MES As String
    Dim año As String
    Dim pasador As String
'  select centrocosto,codigocuenta,sum(monto),date_format(fecha,'%m'),dh as mes
'from movimientoscontables where centrocosto<>"" and codigocuenta<>""
'group by codigocuenta,centrocosto,mes,dh order by codigocuenta,centrocosto


        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT rutctacte,codigocuenta,sum(monto),mes,dh "
 
        csql.sql = csql.sql + "FROM movimientoscontables where rutctacte<>' ' and codigocuenta<>' ' "
        csql.sql = csql.sql + "and fecha>='" & fecha1 & "' and fecha<='" & fecha2 & "' "
        
        csql.sql = csql.sql + "group by codigocuenta,rutctacte,mes,dh order by codigocuenta,rutctacte "
        csql.Execute
        For k = 1 To 12
        debe(k) = 0
        haber(k) = 0
        Next k
        barra.Value = 1
        barra.Max = csql.RowsAffected + 2
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        pasador = resultados(1) + resultados(0)
   
         While Not resultados.EOF
            If pasador <> resultados(1) + resultados(0) Then
            Call actualizacuentacorriente(Mid(pasador, 1, 8), Mid(pasador, 9, 10))
            pasador = resultados(1) + resultados(0)
            For k = 1 To 12
            debe(k) = 0
            haber(k) = 0
            Next k
            
            End If
            barra.Value = barra.Value + 1
            If resultados(4) = "H" Then haber(resultados(3)) = resultados(2)
            If resultados(4) = "D" Then debe(resultados(3)) = resultados(2)
            
             resultados.MoveNext
           
         Wend
         Call actualizacuentacorriente(Mid(pasador, 1, 8), Mid(pasador, 9, 10))
          resultados.Close
            Set resultados = Nothing

        End If

End Sub

Sub actualizacuentacorriente(tipo, rut)
    Dim SUMAVALOR As Double
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debe01"
    campos(4, 0) = "debe02"
    campos(5, 0) = "debe03"
    campos(6, 0) = "debe04"
    campos(7, 0) = "debe05"
    campos(8, 0) = "debe06"
    campos(9, 0) = "debe07"
    campos(10, 0) = "debe08"
    campos(11, 0) = "debe09"
    campos(12, 0) = "debe10"
    campos(13, 0) = "debe11"
    campos(14, 0) = "debe12"
    campos(15, 0) = "haber01"
    campos(16, 0) = "haber02"
    campos(17, 0) = "haber03"
    campos(18, 0) = "haber04"
    campos(19, 0) = "haber05"
    campos(20, 0) = "haber06"
    campos(21, 0) = "haber07"
    campos(22, 0) = "haber08"
    campos(23, 0) = "haber09"
    campos(24, 0) = "haber10"
    campos(25, 0) = "haber11"
    campos(26, 0) = "haber12"
    campos(27, 0) = ""
    For k = 1 To 12
    
    campos(k + 2, 1) = debe(k)
    campos(k + 14, 1) = haber(k)
    Next k
    campos(0, 1) = tipo
    campos(1, 1) = rut
    campos(2, 1) = año
    condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año ='" + año + "' "
    
    campos(0, 2) = "saldosctacte"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub
Sub LEERMOVIMIENTOSCRCC(fecha1, fecha2, empre)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim COMANDO As String
    Dim MES As String
    Dim año As String
    Dim pasador As String
'  select centrocosto,codigocuenta,sum(monto),date_format(fecha,'%m'),dh as mes
'from movimientoscontables where centrocosto<>"" and codigocuenta<>""
'group by codigocuenta,centrocosto,mes,dh order by codigocuenta,centrocosto


        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT centrocosto,codigocuenta,sum(monto),mes,dh "
        csql.sql = csql.sql + "FROM " + clientesistema + "conta" + empre + ".movimientoscontables where centrocosto<>' ' and codigocuenta<>' ' "
        csql.sql = csql.sql + "and fecha>='" & fecha1 & "' and fecha<='" & fecha2 & "' "
        csql.sql = csql.sql + "group by codigocuenta,centrocosto,mes,dh order by codigocuenta,centrocosto"
        csql.Execute
        For k = 1 To 12
        debe(k) = 0
        haber(k) = 0
        Next k
        barra.Value = 1
        barra.Max = csql.RowsAffected + 2
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        pasador = resultados(1) + resultados(0)
   
         While Not resultados.EOF
            If pasador <> resultados(1) + resultados(0) Then
            Call actualizacrcc(Mid(pasador, 9, 4), Mid(pasador, 1, 8), empre)
            
            pasador = resultados(1) + resultados(0)
            For k = 1 To 12
            debe(k) = 0
            haber(k) = 0
            Next k
            
            End If
            barra.Value = barra.Value + 1
            If resultados(4) = "H" Then haber(resultados(3)) = resultados(2)
            If resultados(4) = "D" Then debe(resultados(3)) = resultados(2)
            
             resultados.MoveNext
           
         Wend
         Call actualizacrcc(Mid(pasador, 9, 4), Mid(pasador, 1, 8), empre)
          resultados.Close
            Set resultados = Nothing

        End If

End Sub

Sub actualizacrcc(CRCC, cuenta, empre)
    Dim SUMAVALOR As Double
    campos(0, 0) = "codigo"
    campos(1, 0) = "cuenta"
    campos(2, 0) = "año"
    campos(3, 0) = "debe01"
    campos(4, 0) = "debe02"
    campos(5, 0) = "debe03"
    campos(6, 0) = "debe04"
    campos(7, 0) = "debe05"
    campos(8, 0) = "debe06"
    campos(9, 0) = "debe07"
    campos(10, 0) = "debe08"
    campos(11, 0) = "debe09"
    campos(12, 0) = "debe10"
    campos(13, 0) = "debe11"
    campos(14, 0) = "debe12"
    campos(15, 0) = "haber01"
    campos(16, 0) = "haber02"
    campos(17, 0) = "haber03"
    campos(18, 0) = "haber04"
    campos(19, 0) = "haber05"
    campos(20, 0) = "haber06"
    campos(21, 0) = "haber07"
    campos(22, 0) = "haber08"
    campos(23, 0) = "haber09"
    campos(24, 0) = "haber10"
    campos(25, 0) = "haber11"
    campos(26, 0) = "haber12"
    campos(27, 0) = ""
    
    
    For k = 1 To 12
    campos(k + 2, 1) = debe(k)
    campos(k + 14, 1) = haber(k)
    
    Next k
    
    campos(0, 1) = CRCC
    campos(1, 1) = cuenta
    campos(2, 1) = año
    condicion = "codigo=" + "'" + CRCC + "' and cuenta='" + cuenta + "' and año ='" + año + "' "
    
    campos(0, 2) = clientesistema + "conta" + empre + ".saldoscentrosdecosto"
   
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   
Rem actualiza cuenta madre
    condicion = "codigo='" + CRCC + "' and cuenta=" + "'" + Mid(cuenta, 1, 4) + "0000" + "' and año ='" + año + "' order by codigo"

    campos(0, 2) = clientesistema + "conta" + empre + ".saldoscentrosdecosto"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)


  '  If sqlconta.status = 4 Then Stop


    campos(0, 1) = CRCC
    campos(1, 1) = Mid(cuenta, 1, 4) + "0000"
    campos(2, 1) = año
    For k = 1 To 12
    campos(k + 2, 1) = Str(sqlconta.response(k + 2, 3) + debe(k))
    campos(k + 14, 1) = Str(sqlconta.response(k + 14, 3) + haber(k))
    
    Next k
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

Rem actualiza cuenta madre
    condicion = "codigo='" + CRCC + "' and cuenta=" + "'" + Mid(cuenta, 1, 2) + "000000" + "' and año ='" + año + "' order by codigo"

    campos(0, 2) = clientesistema + "conta" + empre + ".saldoscentrosdecosto"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)


  '  If sqlconta.status = 4 Then Stop


    campos(0, 1) = CRCC
    campos(1, 1) = Mid(cuenta, 1, 2) + "000000"
    campos(2, 1) = año
    For k = 1 To 12
    campos(k + 2, 1) = Str(sqlconta.response(k + 2, 3) + debe(k))
    campos(k + 14, 1) = Str(sqlconta.response(k + 14, 3) + haber(k))
    
    Next k
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show vbModal
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
