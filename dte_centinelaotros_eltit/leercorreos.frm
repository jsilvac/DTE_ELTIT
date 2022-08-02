VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form electro07 
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "ENVIAR RESPUESTA CORREO"
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
      Left            =   7560
      TabIndex        =   9
      Top             =   9480
      Width           =   4095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Todos"
      Height          =   255
      Left            =   11640
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "RECIBIR DOCUMENTO A D.D."
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
      Left            =   3360
      TabIndex        =   4
      Top             =   9480
      Width           =   3735
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   16325
      BackColor       =   16744576
      Caption         =   "Envio de Correos"
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Por Enviar Respuesta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "GENERA INFORME"
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Por Recibir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin FlexCell.Grid impresion 
         Height          =   7815
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   13785
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin MSComctlLib.ProgressBar RECIBIDOS 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   8880
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LEER DTE CORREOS"
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
      Left            =   480
      TabIndex        =   0
      Top             =   9480
      Width           =   2535
   End
End
Attribute VB_Name = "electro07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
For K = 1 To impresion.Rows - 1
If impresion.Cell(K, 5).text = "1" Then
impresion.Cell(K, 5).text = "0"
Else
impresion.Cell(K, 5).text = "1"

End If

Next K

End Sub

Private Sub Command1_Click()

Dim objOutlook As Object
Dim objMAPI As Object
Dim Item As Object
Dim Attach As Object
Dim RUT As String

'Abrir Outlook
Set objOutlook = CreateObject("Outlook.Application")
'Iniciar sesión MAPI
Set objMAPI = objOutlook.GetNamespace("MAPI")
' Set objOutlook.ActiveExplorer.CurrentFolder = _
' objMAPI.GetDefaultFolder(6)

'objMAPI.Logon _
'"[dte@eltit.cl]", _
'"[123]", _
'False, _
'False
'Obtener la carpeta de mensajes
With objMAPI.GetDefaultFolder(6) '6=olFolderInbox
'Recorrer los elementos de la carpeta

maximo = .items.Count
INICIO = 1
RECIBIDOS.Max = maximo
RECIBIDOS.Value = 0

For K = 1 To maximo
'Recuperar el primer elemento

Set Item = .items(K)


'Guardar los archivos adjuntos
RECIBIDOS.Value = K
For Each Attach In Item.Attachments


If UCase(Right(Attach, 3)) = "XML" Then
If existecorreo(Attach) = False Then
de = .items(K).senderemailaddress
        
        ARCHIVO = "c:\DTE_RECIBIDOS\" + Attach
        Attach.SaveAsFile ARCHIVO
        archivo_leido = leerxmlrecibido(ARCHIVO)

        Call grabar_recepcion(de, Attach, fechasistema, archivo_leido, respuesta_leido, fechasistema)
End If

End If

Next
'Borrar el correo
Rem Item.Delete
Next K
End With
'Cerrar sesión
objMAPI.Logoff
'Destruir referencias
Set objOutlook = Nothing
Set Item = Nothing
Set Attach = Nothing
Set objMAPI = Nothing
Command2_Click

End Sub

Private Sub Grid1_AfterReorderColumn(ByVal OriginalPosition As Long, ByVal NewPosition As Long)

End Sub
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatogrilla(10, 20) As String
        Dim i As Integer
        col = 6
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "RUT"
        formatogrilla(1, 2) = "PROVEEDOR"
        formatogrilla(1, 3) = "CORREO"
        formatogrilla(1, 4) = "ARCHIVO"
        formatogrilla(1, 5) = "OK"
        formatogrilla(1, 6) = "ENVIAR"
        formatogrilla(1, 7) = "ENVIADO"
        formatogrilla(1, 8) = "I.REF"
        formatogrilla(1, 9) = "I.VINOS"
        formatogrilla(1, 10) = "I.LIC"
        formatogrilla(1, 11) = "IHA "
        formatogrilla(1, 12) = "ICA "
        formatogrilla(1, 13) = "EXENTO"
        formatogrilla(1, 14) = "TOTAL"
        formatogrilla(1, 15) = "Nº SISTEMA"
        
        Rem LARGO DE LOS DATOS
        
        formatogrilla(2, 1) = "4"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "9"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "30"
        formatogrilla(2, 6) = "9"
        formatogrilla(2, 7) = "9"
        formatogrilla(2, 8) = "8"
        formatogrilla(2, 9) = "8"
        formatogrilla(2, 10) = "8"
        formatogrilla(2, 11) = "8"
        formatogrilla(2, 12) = "8"
        formatogrilla(2, 13) = "0"
        formatogrilla(2, 14) = "9"
        formatogrilla(2, 15) = "9"
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "D"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        formatogrilla(3, 11) = "N"
        formatogrilla(3, 12) = "N"
        formatogrilla(3, 13) = "N"
        formatogrilla(3, 14) = "N"
        formatogrilla(3, 15) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = "###,###,##0"
        formatogrilla(4, 7) = "##,###,##0"
        formatogrilla(4, 8) = "##,###,##0"
        formatogrilla(4, 9) = "##,###,##0"
        formatogrilla(4, 10) = "##,###,##0"
        formatogrilla(4, 11) = "##,###,##0"
        formatogrilla(4, 12) = "##,###,##0"
        formatogrilla(4, 13) = "##,###,##0"
        formatogrilla(4, 14) = "###,###,##0"
        formatogrilla(4, 15) = "0000000000"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "TRUE"
        formatogrilla(5, 13) = "TRUE"
        formatogrilla(5, 14) = "TRUE"
        formatogrilla(5, 15) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        formatogrilla(6, 10) = ""
        formatogrilla(6, 11) = ""
        formatogrilla(6, 12) = ""
        formatogrilla(6, 13) = ""
        formatogrilla(6, 14) = ""
        formatogrilla(6, 15) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        Rem ANCHO
        formatogrilla(8, 1) = "8"
        formatogrilla(8, 2) = "25"
        formatogrilla(8, 3) = "20"
        formatogrilla(8, 4) = "25"
        formatogrilla(8, 5) = "5"
        formatogrilla(8, 6) = "5"
        
        formatogrilla(8, 13) = "0"
        
        formatogrilla(8, 14) = "7"
         formatogrilla(8, 15) = "7"
        
'        formatoGrilla(1, 7) = "I.V.A"
'        formatoGrilla(1, 8) = "I.REF"
'        formatoGrilla(1, 9) = "I.VINOS"
'        formatoGrilla(1, 10) = "I.LICORES"
'        formatoGrilla(1, 11) = "IHA "
'        formatoGrilla(1, 12) = "ICA "
        
                
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        impresion.Column(0).Alignment = cellLeftGeneral
        
        
        impresion.Column(0).Width = 16
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1.75
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
  impresion.Column(5).CellType = cellCheckBox
 Rem  impresion.Column(6).CellType = cellCheckBox
  
  
  
        
        
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
    End Sub

Private Sub Command2_Click()
Call CargaGrillaInforme(1, 8)

listarecepciones

End Sub

Private Sub Command3_Click()
If Option1.Value = True Then
RECIBIDOS.Max = impresion.Rows - 1
RECIBIDOS.Value = 0
For K = 1 To impresion.Rows - 1
RECIBIDOS.Value = K
If impresion.Cell(K, 5).text = "1" Then
Call recepcionar(impresion.Cell(K, 4).text, impresion.Cell(K, 1).text)
End If
Next K
Command2_Click
Else
MsgBox "NO VALIDO EN ESTE ESTADO DE SELECCION "
End If

End Sub

Private Sub Command4_Click()
Dim archi As String
Dim adjunto As String

If Option2.Value = True Then
For K = 1 To impresion.Rows - 1
If impresion.Cell(K, 5).text = "1" Then
Rem impresion.Cell(K, 3).text = "arielgodoy@gmail.com"

adjunto = "c:\fae\" + empresaActiva + "\acuse\ACUSE_" + impresion.Cell(K, 4).text
Call EnviarMail("respuesta ", "adjuntamos respuesta recepcion de DTE ", "mail.eltit.cl", impresion.Cell(K, 3).text, impresion.Cell(K, 2).text, adjunto)
Call modificaenviocorreo(impresion.Cell(K, 4).text)


End If

Next K

End If
Command2_Click

End Sub

Private Sub Form_Load()
Call CargaGrillaInforme(1, 8)


End Sub
Private Function listarecepciones()
    Dim tabla As String
    Dim rubAux As String
    Dim harinas As Double
    Dim subproductos As Double
    Dim envases As Double
    Dim trigo As Double
    Dim maquila As Double
    Dim otros As Double
    Dim cadena As String
    Dim tipoDoc As String
    Dim numeroDoc As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    Dim linea As Double
    Dim resultados As rdoResultset
        
    Dim i As Integer

    rubAux = rubro
    Rem Call Conectarventas(servidor, baseVentas + empresaActiva, usuario, password)
    
    
    Set csql.ActiveConnection = ventasRubro


    csql.sql = "SELECT fr.rut,pr.razonsocial,fr.correo,fr.archivo,fr.fecha_respuesta "
    csql.sql = csql.sql & "FROM " + clientesistema + "fae" + empresaActiva + ".sv_recepcion_dte" + empresaActiva + " as fr left join " + clientesistema + "fae" + empresaActiva + ".sv_fae_proveedores as pr on pr.rut=fr.rut "
    If Option1.Value = True Then
    csql.sql = csql.sql & "where archivo_respuesta='' "
    End If
    
    If Option2.Value = True Then
    csql.sql = csql.sql & "where fecha_envio='0000-00-00' "
    End If
    
    csql.Execute
  
    linea = 0
    If csql.RowsAffected > 0 Then
       impresion.Rows = 1
       Set resultados = csql.OpenResultset
        While Not resultados.EOF
       If Right(resultados(2), 6) <> "sii.cl" Then
           impresion.Rows = impresion.Rows + 1
           linea = linea + 1
            
            impresion.Cell(linea, 1).text = resultados(0)
            If IsNull(resultados(1)) = False Then
            impresion.Cell(linea, 2).text = resultados(1)
            End If
            impresion.Cell(linea, 3).text = resultados(2)
            impresion.Cell(linea, 4).text = resultados(3)
            Rem impresion.Cell(linea, 5).text = resultados(4)
            End If
            resultados.MoveNext
        Wend
    
    End If
Set csql = Nothing
csql.Close
Set resultados = Nothing

    'Call sumaGrilla(impresion)
    
End Function

Private Sub EnviarMail(ByRef Asunto, ByRef mensaje, ByRef Servidor, ByRef MailDestinatario, ByRef NombreDestinatario, ByRef ArchivAdjunto)
Set email = New clsSendMail
Screen.MousePointer = vbHourglass
MailDestinatario = LCase(MailDestinatario)

 With email

     ' **************************************************************************
        .SMTPHostValidation = VALIDATE_NONE
        .EmailAddressValidation = VALIDATE_SYNTAX
        .Delimiter = ";"
     ' **************************************************************************
        .AsHTML = False                             'ENVIAR COMO HTML O TEXTO PLANO
        .SMTPHost = Servidor     'Servidor                        ' servidor smtp
        .From = "dte@eltit.cl" 'usuario                  'email de quien envia
        .FromDisplayName = NOMBREEMPRESA 'nombreempresa            'NOMBRE DE QUIEN ENVIA
        .Recipient = MailDestinatario           'email para
        .RecipientDisplayName = NombreDestinatario  'NOMBRE PARA
        '.CcRecipient = usuario                           'EMAIL COPIA
        .CcDisplayName = ""                         'NOMBRE COPIA
        .BccRecipient = MailDestinatario                                     'EMAIL COPIA OCULTA
        .ReplyToAddress = ""                        'responder a otro email
        .Subject = Asunto                           'ASUNTO
        .Message = mensaje 'txtemail.text                    'CUERPO MAIL
If ArchivAdjunto <> "*.*" Then .Attachment = ArchivAdjunto 'RUTA ARCHIVO ADJUNTO Trim(txtAttach.Text)
     ' **************************************************************************
        .ContentBase = ""
        .EncodeType = 0 'CODIFICACION ARCHIVOS ADJUNTOS
        .Priority = HIGH_PRIORITY                   'PRIORIDAD
        .Receipt = False                            '
        .UseAuthentication = True                   'servidor requiere autenticacion
        .UsePopAuthentication = False               'usar autenticacion del pop
        .UserName = "dte@eltit.cl" 'usuario                         'usuario cuenta correo
        .password = "123"    'CLAVE                           'clave cuenta correo
        .POP3Host = "mail.eltit.cl" 'Servidor                        'servidor pop
        .MaxRecipients = 100                        '
        .Send                                       'envia el correo

    End With
   
    Screen.MousePointer = vbDefault
End Sub

