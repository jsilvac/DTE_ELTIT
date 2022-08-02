VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MailSocios 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Envio de Correos EltitPlus"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp ESTADO_ENVIOS 
      Height          =   3855
      Left            =   1320
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6800
      BackColor       =   16744576
      Caption         =   "ENVIANDO CORREOS"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BordeEstilo     =   4
      Alignment       =   1
      Begin VB.TextBox txtemail 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         ToolTipText     =   "Vista Previa del mensaje"
         Top             =   720
         Width           =   7575
      End
      Begin MSComctlLib.ProgressBar progresototal 
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   3600
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ProgressBar progresocorreo 
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENVIANDO "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL DE CORREOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CORREOS RESTANTES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label correoactual 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label totalcorreos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label correosrestantes 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   3240
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdEnviarEmail 
      Caption         =   "ENVIAR EMAIL"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CommandButton CmdAdjuntarArchivo 
      Caption         =   "Adjuntar Un Archivo"
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   8280
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog ArchivoAdjunto 
      Left            =   10200
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Adjuntar un Archivo"
      FileName        =   "*.*"
      InitDir         =   "C:"
   End
   Begin XPFrame.FrameXp FrameTextoEmail 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6165
      BackColor       =   16761024
      Caption         =   "TEXTO DEL EMAIL"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      ColorBarraAbajo =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox TxEmail 
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FF00&
         Height          =   3135
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "MailSocios.frx":0000
         Top             =   240
         Width           =   10095
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1720
      BackColor       =   16761024
      Caption         =   "DATOS DEL DESTINATARIO DEL CORREO"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      ColorBarraAbajo =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox TxtAsunto 
         Height          =   285
         Left            =   4560
         TabIndex        =   11
         Text            =   "ASUNTO"
         Top             =   600
         Width           =   5295
      End
      Begin VB.OptionButton Para 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Seleccion"
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Para 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Para Todos"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton Para 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Para"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox TxtPara 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "para@mail.cl"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ASUNTO :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   3495
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6165
      BackColor       =   16761024
      Caption         =   "LISTADO DE SOCIOS"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      ColorBarraAbajo =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin FlexCell.Grid GrillaClientes 
         Height          =   3255
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5741
         Appearance      =   0
         BackColor2      =   12632256
         Cols            =   3
         DefaultFontSize =   8.25
         DefaultRowHeight=   20
         Rows            =   10
      End
   End
End
Attribute VB_Name = "MailSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAdjuntarArchivo_Click()
On Error GoTo error
Dim ArchivosAdj As String
With ArchivoAdjunto
            .FileName = ""
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
        .FilterIndex = 1
        .DialogTitle = "Seleccione el archivo a adjuntar!!"
        .ShowOpen
        ArchivosAdj = .FileName
End With
If ArchivosAdj <> Empty Then CmdAdjuntarArchivo.BackColor = vbGreen

Exit Sub
error:
Resume Next
End Sub

Private Sub CmdEnviarEmail_Click()
totalcorreos = GrillaClientes.Rows - 1
correosrestantes = totalcorreos
ESTADO_ENVIOS.Visible = True
ESTADO_ENVIOS.ZOrder (0)

If Para(0).Value = True Then
 Call EnviarMail(TxtAsunto.text, Me.TxEmail.text, "", TxtPara.text, "", ArchivoAdjunto.FileName)
ESTADO_ENVIOS.Visible = False
ESTADO_ENVIOS.ZOrder (1)
Exit Sub
End If
ESTADO_ENVIOS.ZOrder (0)
ESTADOS_ENVIOS.Visible = True
Dim Y, cont As Integer
cont = 0
For Y = 1 To GrillaClientes.Rows - 1

'With GrillaClientes

   
   If GrillaClientes.Cell(Y, 3).text <> Empty Then
   cont = cont + 1
   correosrestantes = correosrestantes - 1
   Call EnviarMail(TxtAsunto.text, Me.TxEmail.text, "", GrillaClientes.Cell(Y, 2).text, GrillaClientes.Cell(Y, 1).text, ArchivoAdjunto.FileName)
   End If
Next
MsgBox "correos enviados " & cont
ESTADO_ENVIOS.Visible = False
ESTADO_ENVIOS.ZOrder (1)
End Sub
Private Sub EnviarMail(ByRef Asunto, ByRef mensaje, ByRef Servidor, ByRef MailDestinatario, ByRef NombreDestinatario, ByRef ArchivAdjunto)
Set email = New clsSendMail
Screen.MousePointer = vbHourglass
MailDestinatario = LCase(MailDestinatario)
progresocorreo.Value = 5
 With email

     ' **************************************************************************
        .SMTPHostValidation = VALIDATE_NONE
        .EmailAddressValidation = VALIDATE_SYNTAX
        .Delimiter = ";"
     ' **************************************************************************
        .AsHTML = False                             'ENVIAR COMO HTML O TEXTO PLANO
        .SMTPHost = "mail.eltitplus.cl"     'Servidor                        ' servidor smtp
        .From = "web@eltitplus.cl" 'usuario                  'email de quien envia
        .FromDisplayName = "EltitPlus" 'nombreempresa            'NOMBRE DE QUIEN ENVIA
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
        .UserName = "web@eltitplus.cl" 'usuario                         'usuario cuenta correo
        .password = "123456"    'CLAVE                           'clave cuenta correo
        .POP3Host = "mail.eltitplus.cl" 'Servidor                        'servidor pop
        .MaxRecipients = 100                        '
        .Send                                       'envia el correo

    End With
    progresocorreo.Value = 10
    Screen.MousePointer = vbDefault
progresocorreo.Value = 1
correosrestantes = correosrestantes.Caption - 1
End Sub

Private Sub Form_Load()
Formatos
LeeSocios
End Sub
Public Sub LeeSocios()
    Dim consulta As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim total_comprado As Double
    Dim total_recepcionado As Double
    Dim F As String
    Dim color As String
    color = &HFFFFFF
    consulta = "SELECT (CONCAT(nombre,' ',apellidopaterno)),email FROM sv_maestrosocios where email <> '' and email <> '0' ;"
    Set csql.ActiveConnection = ventas
    csql.sql = consulta
    csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            GrillaClientes.AutoRedraw = False
            GrillaClientes.Rows = 1
            GrillaClientes.Cols = 4
            
     While Not resultados.EOF
            GrillaClientes.AddItem resultados(0) & vbTab & resultados(1)
            resultados.MoveNext
     Wend
            resultados.Close
            Set resultados = Nothing
            GrillaClientes.AutoRedraw = True
            GrillaClientes.Refresh
            CmdEnviarEmail.Enabled = True
            CmdAdjuntarArchivo.Enabled = True
        Else
            CmdEnviarEmail.Enabled = False
            CmdAdjuntarArchivo.Enabled = False
            grillasocios.Rows = 1
            MsgBox "No se encontraron resultados para la búsqueda, elija otro criterio e intente nuevamente.", vbInformation + vbOKOnly, "Consulta sin Resultados"
        End If




End Sub

Private Sub Formatos()
With GrillaClientes
.Cols = 4
    .FrozenCols = 1
    .Column(0).Width = 0
    .Column(1).Width = 250
    .Column(2).Width = 250
    .Column(3).Width = 100
    
    .Column(0).Locked = True
    .Column(1).Locked = True
    .Column(2).Locked = True
    
    .Cell(0, 1).text = "NOMBRE DEL SOCIO"
    .Cell(0, 2).text = "EMAIL"
    .Cell(0, 3).text = "ENVIAR EMAIL"
       
    .Column(3).CellType = cellCheckBox
    .Column(3).Alignment = cellLeftBottom
End With
End Sub

Private Sub Para_Click(Index As Integer)
Dim Y, cont As Integer

For Y = 1 To GrillaClientes.Rows - 1
TxtPara.Locked = True
TxtPara.text = Empty
Select Case Index
    Case 0
        GrillaClientes.Cell(Y, 3).text = Empty
        TxtPara.Locked = False
        TxtPara.text = Empty
        TxtPara.SetFocus
        Exit Sub
    Case 1
        
        GrillaClientes.Cell(Y, 3).text = 1
    Case 2
         GrillaClientes.Cell(Y, 3).text = Empty
End Select
Next

End Sub
