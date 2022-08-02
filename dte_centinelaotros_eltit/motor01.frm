VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form motor01 
   Caption         =   "Motor Consulta Cheques"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   10290
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   18150
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
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   600
         Top             =   9480
      End
      Begin VB.Timer Timer2 
         Interval        =   5000
         Left            =   90
         Top             =   9720
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Top             =   9600
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   495
         Left            =   2640
         TabIndex        =   8
         Top             =   9600
         Width           =   1335
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   120
         TabIndex        =   7
         Top             =   4680
         Width           =   13695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Salir"
         Height          =   465
         Left            =   10920
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   9600
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
      Begin XPFrame.FrameXp CONSULTAS 
         Height          =   870
         Left            =   120
         TabIndex        =   4
         Top             =   8520
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   1535
         BackColor       =   16761024
         Caption         =   "ESTADO DE CONSULTAS EN PROGRESO"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
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
         Begin MSComctlLib.ProgressBar barra 
            Height          =   465
            Left            =   45
            TabIndex        =   5
            Top             =   360
            Width           =   13605
            _ExtentX        =   23998
            _ExtentY        =   820
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Detener Consultas"
         Height          =   510
         Left            =   4680
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   9600
         UseMaskColor    =   -1  'True
         Width           =   1995
      End
      Begin FlexCell.Grid grid1 
         Height          =   4305
         Left            =   45
         TabIndex        =   1
         Top             =   240
         Width           =   13770
         _ExtentX        =   24289
         _ExtentY        =   7594
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Continuar Consultas"
         Height          =   510
         Left            =   7320
         MaskColor       =   &H00C0FFC0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   9600
         UseMaskColor    =   -1  'True
         Width           =   1995
      End
      Begin MSWinsockLib.Winsock Ws 
         Left            =   1440
         Top             =   7680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "motor01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private resul(20) As Variant
Private Progreso As Boolean
Private K As Double
Private codigorecepcion As String
Private glosarespuesta As String
Private codigoaurotizacion As String
Private codigoreferencia As String
Private Declare Sub ISOCHEQUE_DLL Lib "c:\orsan\IsoPosUI\IsoPosUI\IsoCheq.dll" (ByVal MIRC As String, ByVal rut As String, ByVal MONTO As String, ByVal operador As String, ByVal FECHAVENCI As String, ByVal rspOUT As String, ByVal scrOUT As String, ByVal ticketOUT As String)
Private PASADA As Boolean

Private Sub Command1_Click()
CONSULTAS.Caption = "ESTADO DE CONSULTAS DETENIDO"
Progreso = False
barra.Value = 0
End Sub

Private Sub Command2_Click()


CONSULTAS.Caption = "ESTADO DE CONSULTAS EN PROGRESO"
Progreso = True
LEERCONSULTAS



End Sub

Private Sub Command3_Click()
If MsgBox("ESTA OPCION DETENDRA EL PROCESO DE CONSULTA CON LA SEGURADORA DE CHEQUES ESTA SEGURO", vbYesNo) = vbYes Then

Unload Me
End If
End Sub

Private Sub Command4_Click()
LEERCONSULTAS
PASADA = True

With Ws
.Close
      ' conecta al servidor en el puerto 25
      .Connect "216.241.20.139", 1001
         
      ' Bucle mientras conecta al Smtp
      Do While .State <> sckConnected
         DoEvents
         If .State = sckClosed Or .State = sckError Then
               MsgBox "Error ", vbCritical
              
               Exit Sub
         End If
      Loop
End With

K = 1
            If Grid1.Rows - 1 > 0 Then
            resul(0) = Grid1.Cell(K, 1).text
            resul(1) = Grid1.Cell(K, 2).text
            resul(2) = Grid1.Cell(K, 3).text
            resul(3) = Grid1.Cell(K, 4).text
            resul(4) = Grid1.Cell(K, 5).text
            resul(5) = Grid1.Cell(K, 6).text
            resul(6) = Grid1.Cell(K, 7).text
            resul(7) = Grid1.Cell(K, 8).text
            resul(8) = Grid1.Cell(K, 9).text
            resul(9) = Grid1.Cell(K, 10).text
            resul(10) = Grid1.Cell(K, 11).text
            resul(11) = Grid1.Cell(K, 12).text
sucursal = "     148482"
PASADA = False

Call send(sucursal, "00", resul(11), resul(5), resul(7), resul(8), resul(9), resul(6), resul(10))

            End If



End Sub

Private Sub Command5_Click()
List1.Clear

End Sub

Private Sub Form_Load()
Call CARGAGRILLA(2, 13)
Progreso = True
barra.Max = 1001
Ws.RemoteHost = "216.241.20.139"
Ws.RemotePort = 1001

'Command2_Click




End Sub

  Private Sub CARGAGRILLA(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Dim formatogrilla(12, 12) As String
        
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "FECHA"
        formatogrilla(1, 2) = "HORA"
        formatogrilla(1, 3) = "LOCAL"
        formatogrilla(1, 4) = "CAJA"
        formatogrilla(1, 5) = "CAJERA"
        formatogrilla(1, 6) = "CUENTA"
        formatogrilla(1, 7) = "NUMERO"
        formatogrilla(1, 8) = "BANCO"
        formatogrilla(1, 9) = "PLAZA"
        formatogrilla(1, 10) = "MONTO"
        formatogrilla(1, 11) = "VENCIMIENTO"
        formatogrilla(1, 12) = "RUT"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "8"
        formatogrilla(2, 2) = "8"
        formatogrilla(2, 3) = "3"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "10"
        formatogrilla(2, 6) = "10"
        formatogrilla(2, 7) = "10"
        formatogrilla(2, 8) = "10"
        formatogrilla(2, 9) = "10"
        formatogrilla(2, 10) = "10"
        formatogrilla(2, 11) = "10"
        formatogrilla(2, 12) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "D"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        formatogrilla(3, 11) = "D"
        formatogrilla(3, 12) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 10) = "$###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "8"
        formatogrilla(8, 2) = "8"
        formatogrilla(8, 3) = "3"
        formatogrilla(8, 4) = "9"
        formatogrilla(8, 5) = "10"
        formatogrilla(8, 6) = "10"
        formatogrilla(8, 7) = "10"
        formatogrilla(8, 8) = "10"
        formatogrilla(8, 9) = "10"
        formatogrilla(8, 10) = "10"
        formatogrilla(8, 11) = "10"
        formatogrilla(8, 12) = "10"
            
        Grid1.Cols = col
        Grid1.Rows = row
        Grid1.AllowUserResizing = False
        Grid1.DisplayFocusRect = False
        Grid1.ExtendLastCol = True
        Grid1.BoldFixedCell = False
        Grid1.DrawMode = cellOwnerDraw
        Grid1.Appearance = Flat
        Grid1.ScrollBarStyle = Flat
        Grid1.FixedRowColStyle = Flat
        Grid1.BackColorFixed = RGB(90, 158, 214)
        Grid1.BackColorFixedSel = RGB(110, 180, 230)
        Grid1.BackColorBkg = RGB(90, 158, 214)
        Grid1.BackColorScrollBar = RGB(231, 235, 247)
        Grid1.BackColor1 = RGB(231, 235, 247)
        Grid1.BackColor2 = RGB(239, 243, 255)
        Grid1.GridColor = RGB(148, 190, 231)
        
        Grid1.Column(0).Width = 0
        
        For i = 1 To col - 1
            Grid1.Cell(0, i).text = formatogrilla(1, i)
            Grid1.Column(i).Width = Val(formatogrilla(8, i)) * (Grid1.Cell(0, i).Font.Size + 1.25)
            Grid1.Column(i).MaxLength = Val(formatogrilla(2, i))
            Grid1.Column(i).FormatString = formatogrilla(4, i)
            Grid1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Grid1.Column(i).Alignment = cellRightCenter
            Else
                Grid1.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        
        Grid1.Enabled = True
    End Sub

Private Sub Timer1_Timer()
'Command4_Click

End Sub
Sub LEERCONSULTAS()

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    

        Set cSql.ActiveConnection = ventas
        cSql.sql = "SELECT fecha,hora,local,caja,cajera,cuenta,numero,banco,plaza,monto,vencimiento,rut "
        cSql.sql = cSql.sql + "FROM sv_consultacheques "
        cSql.sql = cSql.sql + "where codigorecepcion='' order by fecha,hora"
        cSql.Execute
        Grid1.AutoRedraw = False
        Grid1.Rows = 1
        If cSql.RowsAffected > 0 Then
            
            Set resultados = cSql.OpenResultset
            
            While resultados.EOF = False
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 1).text = Format(resultados(0), "yyyy-mm-dd")
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
            Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
            Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
            Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(6)
            Grid1.Cell(Grid1.Rows - 1, 8).text = resultados(7)
            Grid1.Cell(Grid1.Rows - 1, 9).text = resultados(8)
            Grid1.Cell(Grid1.Rows - 1, 10).text = resultados(9)
            Grid1.Cell(Grid1.Rows - 1, 11).text = Format(resultados(10), "yyyy-mm-dd")
            Grid1.Cell(Grid1.Rows - 1, 12).text = resultados(11)
            If ASEGURADORA = "1" Then
            
            resul(0) = resultados(0)
            resul(1) = resultados(1)
            resul(2) = resultados(2)
            resul(3) = resultados(3)
            resul(4) = resultados(4)
            resul(5) = resultados(5)
            resul(6) = resultados(6)
            resul(7) = resultados(7)
            resul(8) = resultados(8)
            resul(9) = resultados(9)
            resul(10) = resultados(10)
            resul(11) = resultados(11)
           
            
            
            End If
            If ASEGURADORA = "2" Then
            Call AUTORIZACHEQUESORSAN(resultados(11), resultados(5), resultados(7), resultados(8), resultados(9), resultados(6), resultados(10))
            End If
            
            
            resultados.MoveNext
            Wend
            
                    
            resultados.Close
        Set resultados = Nothing
       
        End If
    Grid1.AutoRedraw = True
    Grid1.Refresh
    
    
End Sub
Sub AUTORIZACHEQUESORSAN(rut, CUENTA, BANCO, PLAZA, MONTO, NUMERO, vencimiento)
'Dim oMIRC As String * 32
'Dim oRUT As String * 10
'Dim oMONTO As String * 6
'Dim oOPERADOR As String * 2
'Dim oFECHAVENCI As String * 8
'Dim rspOUT As String * 21
'Dim scrOUT As String * 61
'Dim ticketOUT As String * 1200
'oMIRC = NUMERO + "XX" + Banco + plaza + "X" + cuenta + "XX00"
'oRUT = rut
'oMONTO = Format(MONTO, "000000")
'oOPERADOR = "00"
'oFECHAVENCI = Format(vencimiento, "yyyymmdd")
'
' ISOCHEQUE_DLL oMIRC, oRUT, oMONTO, oOPERADOR, oFECHAVENCI, rspOUT, scrOUT, ticketOUT
'
' codigoreferencia = Mid(rspOUT, 1, 12)
' codigoautorizacion = Mid(rspOUT, 13, 6)
' codigorespuesta = Mid(rspOUT, 19, 2)
' codigorecepcion = codigorespuesta
'
' If codigorespuesta = "00" Then
' glosarespuesta = "APROBADO"
' End If
'
' If codigorespuesta = "99" Then
' glosarespuesta = "RECHAZADO"
' codigoreferencia = ""
' codigoautorizacion = ""
' codigorespuesta = "99"
' codigorecepcion = "99"
'
' End If
' If codigorespuesta = "91" Then
' glosarespuesta = "ERROR DE CONEXION"
' codigoreferencia = ""
' codigoautorizacion = ""
' codigorespuesta = "91"
' codigorecepcion = "91"
'
' End If
'
'    Call modificaautorizacion(rut, cuenta, Banco, plaza, MONTO, NUMERO, codigorecepcion, glosarespuesta, codigoautorizacion, codigorespuesta, codigoreferencia, nom)
'
  
End Sub

Sub modificaautorizacion(rut, CUENTA, BANCO, PLAZA, MONTO, NUMERO, codigorecepcion, glosarespuesta, codigoautorizacion, codigorespuesta, codigoreferencia, NOMBRECLIENTE)

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    

        Set cSql.ActiveConnection = ventas
        cSql.sql = "update sv_consultacheques "
        cSql.sql = cSql.sql + "set codigorecepcion='" + codigorecepcion + "',glosarespuesta='" + glosarespuesta + "',codigoautorizacion='" + codigoautorizacion + "',codigorespuesta='" + codigorespuesta + "',numeroreferencia='" + codigoreferencia + "',nombrecliente='" + NOMBRECLIENTE + "' "
        cSql.sql = cSql.sql + "where rut='" + rut + "' and cuenta='" + CUENTA + "' and banco='" + BANCO + "' and plaza='" + PLAZA + "' and monto='" & MONTO & "' and numero='" + NUMERO + "' "
        
        cSql.Execute
            Call sincronizadatos(cSql.sql, ventas)
       
End Sub
Sub AUTORIZACHEQUESINSTACHECK(rut, CUENTA, BANCO, PLAZA, MONTO, NUMERO, vencimiento)
Dim sucursal As String

With Ws
 .Close
      .Connect "216.241.20.139", 1001
         
      Do While .State <> sckConnected
         DoEvents
         If .State = sckClosed Or .State = sckError Then
               
               MsgBox "Error ", vbCritical
              conectainstachek = False
              
               
         End If
      Loop

sucursal = "     148482"

Rem  Call send(SUCURSAL, "00", rut, cuenta, NUMERO, banco, vencimiento, MONTO)


 
End With

End Sub


Private Sub Timer2_Timer()

'Command4_Click


End Sub

Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
Dim datos As String
On Error Resume Next
Ws.GetData datos

List1.AddItem Str(Date) + " " + Str(Time) + " RESPUESTA :" + datos
codigoautorizacion = Mid(datos, 5, 10)
If CDbl(codigoautorizacion) <> 0 Then
 codigorecepcion = "01"
 codigorespuesta = "01"
 codigoautorizacion = Mid(datos, 5, 10)
 glosarespuesta = Mid(datos, 35, 30)
 nombreaprobacion = Mid(datos, 75, 30)
 
 Else
 codigorecepcion = "02"
 codigorespuesta = "02"
 glosarespuesta = "RECHAZADO"
 codigoautorizacion = Mid(datos, 5, 10)
 glosarespuesta = Mid(datos, 35, 30)
 nombreaprobacion = Mid(datos, 75, 30)
 
 End If
 
    Call modificaautorizacion(resul(11), resul(5), resul(7), resul(8), resul(9), resul(6), codigorecepcion, glosarespuesta, codigoautorizacion, codigorespuesta, codigoreferencia, NOMBRECLIENTE)
PASADA = True
Call grabarconsultas2(Date, Time, " RESPUESTA :" + datos)

LEERCONSULTAS

End Sub

Sub send(sucursal, fijo, rut, CUENTA, BANCO, PLAZA, MONTO, NUMERO, vencimiento)
On Error Resume Next
DATO1 = sucursal
dato2 = sucursal
dato3 = fijo
dato4 = " " + rut
dato5 = "    " + CUENTA
dato6 = "   " + NUMERO
dato7 = "         " + BANCO
dato8 = Format(vencimiento, "yyyymmdd")
dato9 = ""
dato10 = Format(MONTO, "0000000000")
Call grabarconsultas2(Str(Date), Str(Time), " CONSULTADO " + DATO1 + dato2 + dato3 + dato4 + dato5 + dato6 + dato7 + dato8 + dato9 + dato10)
Sleep (300)

List1.AddItem (Str(Date) + " " + Str(Time) + " CONSULTADO " + DATO1 + dato2 + dato3 + dato4 + dato5 + dato6 + dato7 + dato8 + dato9 + dato10)

Ws.SendData DATO1 + dato2 + dato3 + dato4 + dato5 + dato6 + dato7 + dato8 + dato9 + dato10
Timer2.Enabled = False



End Sub


Sub grabarconsultas2(fecha, HORA, evento)
Dim CAMPOS(5, 5) As String
Dim op As Integer

Dim K As Double
Dim numerocuota As String

CAMPOS(0, 0) = "fecha"
CAMPOS(1, 0) = "hora"
CAMPOS(2, 0) = "evento"
CAMPOS(3, 0) = ""

CAMPOS(0, 1) = Format(fecha, "yyyy-mm-dd")
CAMPOS(1, 1) = HORA
CAMPOS(2, 1) = evento
CAMPOS(0, 2) = "sv_consultas_instacheck"
condicion = ""
        op = 2
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)

End Sub

