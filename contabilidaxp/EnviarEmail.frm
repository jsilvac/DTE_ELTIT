VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EnviarEmail 
   BorderStyle     =   0  'None
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp ESTADO_ENVIOS 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7435
      BackColor       =   16744576
      Caption         =   "ENVIANDO CORREOS"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
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
      Begin VB.TextBox txtemail 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Vista Previa del mensaje"
         Top             =   720
         Width           =   7935
      End
      Begin MSComctlLib.ProgressBar progresocorreo 
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
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
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label correoactual 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "EnviarEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents email As EnviarMail_rz.clsSendMail
Attribute email.VB_VarHelpID = -1

Sub EnviarMail(ByRef NOMBRE, ByRef Usuario, ByRef CLAVE, ByRef Asunto, _
ByRef MENSAJE, ByRef Servidor, ByRef MailDestinatario, ByRef NombreDestinatario, _
ByRef ArchivAdjunto)

If NOMBRE = "" Or Usuario = "" Or CLAVE = "" Or Servidor = "" Or MailDestinatario = "" Then GoTo error

Me.txtemail.text = MENSAJE
Me.correoactual.Caption = MailDestinatario
Set email = New clsSendMail
Screen.MousePointer = vbHourglass
MailDestinatario = LCase(MailDestinatario)
progresocorreo.Value = 50
'MailDestinatario = "rodrigogranadino@hotmail.com"
 With email
    
     ' **************************************************************************
        .SMTPHostValidation = VALIDATE_NONE
        .EmailAddressValidation = VALIDATE_SYNTAX
        .Delimiter = ";"
     ' **************************************************************************
        .AsHTML = False                             'ENVIAR COMO HTML O TEXTO PLANO
        .SMTPHost = Servidor                        ' servidor smtp
        .From = Usuario                             'email de quien envia
        .FromDisplayName = nombreempresa            'NOMBRE DE QUIEN ENVIA
        .Recipient = MailDestinatario               'email para
        .RecipientDisplayName = NombreDestinatario  'NOMBRE PARA
        .CcRecipient = "rauls@adminerp.cl"          'EMAIL COPIA
        .CcDisplayName = ""                         'NOMBRE COPIA
        .BccRecipient = ""                          'EMAIL COPIA OCULTA
        .ReplyToAddress = ""                        'responder a otro email
        .Subject = Asunto                           'ASUNTO
        .Message = MENSAJE                          'CUERPO MAIL
'If AdjuntarOC.Value = 1 And ArchivAdjunto <> "" Then .Attachment = ArchivAdjunto 'RUTA ARCHIVO ADJUNTO Trim(txtAttach.Text)
     ' **************************************************************************
progresocorreo = 70
        .Attachment = ArchivAdjunto
        .ContentBase = ""
        .EncodeType = 0                             'CODIFICACION ARCHIVOS ADJUNTOS
        .Priority = HIGH_PRIORITY                   'PRIORIDAD
        .Receipt = False                            '
        .UseAuthentication = True                   'servidor requiere autenticacion
        .UsePopAuthentication = False               'usar autenticacion del pop
        .UserName = Usuario                         'usuario cuenta correo
        .password = CLAVE                           'clave cuenta correo
        .POP3Host = Servidor                        'servidor pop
        
        .MaxRecipients = 100                        '
        progresocorreo.Value = 90
        .Send                                       'envia el correo

    End With
    progresocorreo.Value = 100
    Screen.MousePointer = vbDefault
progresocorreo.Value = 1

'Unload Me
Exit Sub
error:
    MsgBox "SU CUENTA DE CORREO NO ESTA CONFIGURADA CORRECTAMENTE; " & vbCr & " SOLICITE AYUDA A COMPUTACION", vbCritical, "ATENCION"
fin:
Exit Sub
End Sub

