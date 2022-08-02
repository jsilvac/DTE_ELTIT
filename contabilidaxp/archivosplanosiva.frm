VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form planosiva 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivos Planos I.V.A"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11595
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   773
   Begin VB.CommandButton Command2 
      Caption         =   "Libro Compras"
      Height          =   495
      Left            =   9360
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8865
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2010
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   3545
      BackColor       =   16744576
      Caption         =   "Archivos Planos I.V.A"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
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
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   8280
         TabIndex        =   10
         Top             =   0
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
            TabIndex        =   12
            Top             =   280
            Width           =   1335
         End
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   280
            Width           =   1455
         End
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   3
         Top             =   360
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   1852
         BackColor       =   16744576
         Caption         =   "DATOS DE FILTRADO"
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
         Begin VB.CommandButton Command1 
            Caption         =   "Libro Ventas"
            CausesValidation=   0   'False
            Height          =   495
            Left            =   6720
            TabIndex        =   8
            Top             =   360
            Width           =   2055
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   4
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
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
               Height          =   315
               Left            =   45
               TabIndex        =   5
               Top             =   270
               Width           =   3180
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   3510
            TabIndex        =   6
            Top             =   270
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1191
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
               Height          =   315
               Left            =   90
               TabIndex        =   7
               Top             =   270
               Width           =   2865
            End
         End
      End
   End
End
Attribute VB_Name = "planosiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String
Private COSTO1 As Double
Private COSTO2 As Double
Private COSTO3 As Double
Private COSTO10 As Double
Private COSTO20 As Double
Private COSTO30 As Double

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim consu1, consu2, consu3, archivo As String
    Dim mes2, año2, fecha1, fecha2 As String

    
    
mes2 = Format(COMBOMES.ListIndex + 1, "00")
año2 = COMBOAÑO.text

fecha1 = año2 & "-" & mes2 & "-" & "01"
fecha2 = año2 & "-" & mes2 & "-" & "31"

        Set csql.ActiveConnection = contadb
consu1 = "(select lpad('1',9,' ') as '1',lpad('2',6,' ') as '2' ,lpad('3',10,' ')as '3'," & _
    "lpad('4',1,' ')as '4'," & _
    "lpad('5',2,' ')as '5'," & _
    "lpad('6',10,' ')as '6'," & _
    "lpad('7',8,' ')as '7'," & _
    "lpad('8',9,' ')as '8'," & _
    "lpad('9',50,' ')as '9'," & _
    "lpad('10',13,' ')as '10'," & _
    "lpad('11',13,' ')as '11'," & _
    "lpad('12',13,' ')as '12'," & _
    "lpad('13',13,' ')as '13')"
    
    
consu2 = "(select  '775765305'," & _
    "DATE_FORMAT(fv.fecha, '%m%Y')," & _
    "lpad(replace(format(fv.numero,0),',',''),10,' ')," & _
    "'V'," & _
    "case " & _
    " when fv.tipo='1' then '01'" & _
    " when fv.tipo='2' then '03'" & _
    " when fv.tipo='3' then '04'" & _
    " when fv.tipo='4' then '04'" & _
    " when fv.tipo='5' then '06' end as tipo," & _
    "lpad(replace(format(fv.numero,0),',',''),10,' ')," & _
    "DATE_FORMAT(fv.fecha, '%d%m%Y')," & _
    "mid(fv.rut,2,9)," & _
    "lpad(cc.nombre,50,' ')," & _
    "lpad(replace(format(exento,0),',',''),13,' ')," & _
    "lpad(replace(format(neto,0),',',''),13,' ')," & _
    "lpad(replace(format(iva,0),',',''),13,' ')," & _
    "lpad(replace(format(total,0),',',''),13,' ')" & _
    "from facturasdeventas as fv left join cuentascorrientes as cc on (fv.rut=cc.rut)where " & _
    "fv.fecha between '" & fecha1 & "' and '" & fecha2 & "' and cc.tipo='11200027' and fv.rut <>'0888888888')"
    
consu3 = "(select '775765305'," & _
    "DATE_FORMAT(fecha, '%m%Y')," & _
    "lpad(format(7,0),10,' ')," & _
    "'V'," & _
    "'88'," & _
    "lpad('1',10,' ')," & _
    "DATE_FORMAT(fecha, '%d%m%Y')," & _
    "'775765305'," & _
    "lpad('Almacenes eltit limitada',50,' ')," & _
    "lpad( replace(format(sum(exento),0),',',''),13,' ') AS exen," & _
    "lpad(replace(format(sum(monto)-((sum(monto)*19)/100),0),',',''),13,' ') as netoo," & _
    "lpad(replace(format(sum(monto)-(sum(monto)/1.19)   ,0),',',''),13,' ') as ivaa," & _
    "lpad(Replace(Format(Sum(total), 0), ',', ''), 13, ' ') As totall " & _
    " from boletasdeventa where " & _
    "fecha Between '" & fecha1 & "' and '" & fecha2 & "' group by mid(fecha,1,7))"
 csql.sql = consu1 & " union " & consu2 & " union " & consu3
 csql.Execute
 If csql.RowsAffected > 0 Then
      archivo = "c:\sii\V" + año2 + mes2 + ".txt"
      Open archivo For Output As #14
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
            Print #14, resultados(0) & _
                       resultados(1) & _
                       resultados(2) & _
                       resultados(3) & _
                       resultados(4) & _
                       resultados(5) & _
                       resultados(6) & _
                       resultados(7) & _
                       resultados(8) & _
                       resultados(9) & _
                      resultados(10) & _
                      resultados(11) & _
                      resultados(12)
                resultados.MoveNext
            Wend
            Close 14
            resultados.Close
        Set resultados = Nothing
        End If
MsgBox ("Archivo " & archivo & " generado")
End Sub

Private Sub COMMAND2_Click()

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim consu1, consu2, consu3, archivo As String
    Dim mes2, año2, fecha1, fecha2 As String

    
    
mes2 = Format(COMBOMES.ListIndex + 1, "00")
año2 = COMBOAÑO.text

fecha1 = año2 & "-" & mes2 & "-" & "01"
fecha2 = año2 & "-" & mes2 & "-" & "31"

        Set csql.ActiveConnection = contadb
consu1 = "(select lpad('1',9,' ') as '1',lpad('2',6,' ') as '2' ,lpad('3',10,' ')as '3'," & _
    "lpad('4',1,' ')as '4'," & _
    "lpad('5',2,' ')as '5'," & _
    "lpad('6',10,' ')as '6'," & _
    "lpad('7',8,' ')as '7'," & _
    "lpad('8',9,' ')as '8'," & _
    "lpad('9',50,' ')as '9'," & _
    "lpad('10',13,' ')as '10'," & _
    "lpad('11',13,' ')as '11'," & _
    "lpad('12',13,' ')as '12'," & _
    "lpad('13',13,' ')as '13')"
    
    
consu2 = "( select  '775765305'," & _
    "DATE_FORMAT(fc.fechadigitacion, '%m%Y'), " & _
    "lpad(replace(format(fc.numero,0),',',''),10,' '), " & _
    "'C'," & _
    "Case " & _
    "when fc.tipo='1' then '01'" & _
    "when fc.tipo='2' then '03'" & _
    "when fc.tipo='3' then '04'" & _
    "when fc.tipo='4' then '01'" & _
    "when fc.tipo='5' then '03'" & _
    "when fc.tipo='6' then '04'" & _
    "when fc.tipo='7' then '02'" & _
    "end as tipo," & _
    "lpad(replace(format(fc.numero,0),',',''),10,' ')," & _
    "DATE_FORMAT(fc.fecha, '%d%m%Y')," & _
    "mid(fc.rut,2,9)," & _
    "lpad(cc.nombre,50,' ')," & _
    "lpad(replace(format(exento,0),',',''),13,' ')," & _
    "lpad(replace(format(neto,0),',',''),13,' ')," & _
    "lpad(replace(format(iva,0),',',''),13,' ')," & _
    "lpad(replace(format(total,0),',',''),13,' ')" & _
    "from facturasdecompras as fc left join cuentascorrientes as cc " & _
    " on (fc.rut=cc.rut) where " & _
    "fc.fechadigitacion between '" & fecha1 & "' and '" & fecha2 & "' and cc.tipo='23100026' and fc.rut <>'0888888888' and (mid(fc.fecha,1,4) = cc.año ))"
    
    
    
    
 csql.sql = consu1 & " union " & consu2
 csql.Execute
 If csql.RowsAffected > 0 Then
      archivo = "c:\sii\C" + año2 + mes2 + ".txt"
      Open archivo For Output As #14
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
            Print #14, resultados(0) & _
                       resultados(1) & _
                       resultados(2) & _
                       resultados(3) & _
                       resultados(4) & _
                       resultados(5) & _
                       resultados(6) & _
                       resultados(7) & _
                       resultados(8) & _
                       resultados(9) & _
                      resultados(10) & _
                      resultados(11) & _
                      resultados(12)
                resultados.MoveNext
            Wend
            Close 14
            resultados.Close
        Set resultados = Nothing
        End If
MsgBox ("Archivo " & archivo & " generado")
End Sub

Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD
 Call Conectarventas(Servidor, clientesistema + "ventas00", Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion00", Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001



End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub
Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
