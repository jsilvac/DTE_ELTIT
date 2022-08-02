VERSION 5.00
Begin VB.Form infoma03 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Centro de Costos"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "LISTADO CENTROS DE COSTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
End
Attribute VB_Name = "infoma03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
IMPRIMIR
End Sub

Private Sub Form_Load()
    
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)
fechasistema = Date
dia = Mid(Date, 1, 2)
mes = Mid(Date, 4, 2)
año = Mid(Date, 7, 4)
End Sub
Sub IMPRIMIR()
    informes.info.Clear
    largopagina = 65
    tituloinforme = "CENTRO DE COSTOS"
    titu(1) = "CODIGO    "
    titu(2) = "NOMBRE DE CUENTA"
    lineas = 70
    Consulta_Informe
    informes.Show
    
End Sub
Sub grilla()
    palabra = ""
      
    For K = 1 To cancolu
    If tipodato(K) = "s" Or tipodato(K) = "S" Then dato(K) = dato(K) & String(colu(K) - Len(dato(K)), 32)
    If tipodato(K) = "n" Or tipodato(K) = "N" Then dato(K) = String(colu(K) - Len(dato(K)), 32) & dato(K)
    palabra = palabra & dato(K)
    Next K
    If lineas > largopagina Then Call cabeza
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (" ")
    If Mid(dato(1), 7, 4) <> "0000" Then informes.info.AddItem ("    " + palabra)
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (Mid(palabra, 1, 40))
    If Mid(dato(1), 7, 4) = "0000" Then informes.info.AddItem (" ")
    lineas = lineas + 1

End Sub
Sub cabeza()
    informes.info.AddItem ("")
    informes.info.AddItem ("")
    pagina = pagina + 1
    


    informes.info.AddItem ("NOMBRE EMPRESA          " + tituloinforme + "                                   PAGINA " + Str$(pagina))
    informes.info.AddItem ("DIRECCION EMPRESA                                                                              " + Mid(Date$, 4, 2) + "-" + Mid(Date$, 1, 2) + "-" + Mid(Date$, 7, 4))
    informes.info.AddItem ("RUT EMPRESA                                                                                    " + Time$)
    informes.info.AddItem ("                                " + tituloinforme)
    informes.info.AddItem String(132, "_")
    TITULOS = ""
    For K = 1 To cancolu
    titu(K) = titu(K) & String(colu(K) - Len(titu(K)), 32)
    TITULOS = TITULOS & titu(K)
    Next K
    informes.info.AddItem (TITULOS)
    informes.info.AddItem String(132, "_")

lineas = 8

End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre "
        cSql.SQL = cSql.SQL + "FROM centrosdecosto"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                
                dato(1) = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2): colu(1) = 11: tipodato(1) = "s"
                dato(2) = resultados(1): colu(2) = 52: tipodato(2) = "s"
                 cancolu = 2
                grilla
                resultados.MoveNext
            Wend
            resultados.Close
            
            Set resultados = Nothing

        End If
    End With

End Sub

