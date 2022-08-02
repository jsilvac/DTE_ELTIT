VERSION 5.00
Begin VB.Form Form_pdf 
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox idpdf 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Text            =   "000000001"
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox desde 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Text            =   "c:\33-28987cedi.pdf"
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox salida 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Text            =   "archivosalida.pdf"
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "guadar en Disco"
      Height          =   1095
      Left            =   1800
      TabIndex        =   1
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton guardar 
      Caption         =   "Guardar"
      Height          =   1215
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
   End
End
Attribute VB_Name = "Form_pdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'baja pdf desde mysql
Private Sub Command1_Click()
Dim conn As New ADODB.Connection
conn.ConnectionString = "Driver={Mysql ODBC 3.51 Driver}; Server=192.168.4.9;port=3306; database=eltit_fae00; user=root; password=123; option=3;"
conn.CursorLocation = adUseClient
conn.Open
Dim rs As New ADODB.Recordset
Dim mystream As New ADODB.Stream
mystream.Type = adTypeBinary
rs.Open "SELECT * FROM pdf WHERE pdfid = '" & idpdf.Text & "'", conn
mystream.Open
mystream.Write rs!pdf
mystream.SaveToFile App.Path & "\" & salida.Text, adSaveCreateOverWrite
mystream.Close
rs.Close
conn.Close
End Sub


Private Sub Form_Load()

End Sub

'sube pdf a mysql
Private Sub guardar_Click()
        Dim cn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Set cn = New ADODB.Connection
        cn.ConnectionString = "Driver={Mysql ODBC 3.51 Driver}; Server=192.168.4.9;port=3306; database=eltit_fae00; user=root; password=123; option=3;"
        cn.Open
        Set rs = New ADODB.Recordset
    msgRep = MsgBox("Save the edition?", vbYesNo + vbQuestion, Me.Caption)
    If msgRep = vbYes Then
            rs.Open " select * from pdf", cn, adOpenKeyset, adLockOptimistic
            rs.AddNew
            Set pdffile = New ADODB.Stream
            pdffile.Type = adTypeBinary
            pdffile.Open
            pdffile.LoadFromFile desde.Text
            rs.Fields("pdf") = pdffile.Read
            rs!pdfid = idpdf.Text
            rs!pdfname = idpdf.Text
            pdffile.Close
            Set pdffile = Nothing
            rs.Update
    End If
cn.Close
rs.Close
pdffile.Close
End Sub

'imprime pdf gurdado en HDD
Private Sub Command2_Click()
ShellExecute Me.hwnd, "open", App.Path & "\" & salida.Text, "", "", 4
End Sub

