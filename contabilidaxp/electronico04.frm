VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form libro_electro04 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envios Obligatorios"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmenvios 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      BackColor       =   16761024
      Caption         =   "Opciones"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin VB.CommandButton Command1 
         Caption         =   "firma envio"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SELECCIONAR TODOS"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   1320
         Width           =   2295
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3015
         Left            =   360
         TabIndex        =   12
         Top             =   2400
         Visible         =   0   'False
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   5318
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.CheckBox chk6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA MAYOR RESUMEN"
         Height          =   495
         Left            =   5160
         TabIndex        =   11
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chk4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA LIBRO DIARIO RESUMEN"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chk5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA MAYOR"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chk2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA BALANCE"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chk3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA LIBRO DIARIO"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chk1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA DICCIONARIO"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ENVIAR XML CORREO ELECTRONICO"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   7080
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "GENERAR XML"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ENVIA PDF CORREO ELECTRONICO"
         Height          =   375
         Left            =   -480
         TabIndex        =   1
         Top             =   6480
         Visible         =   0   'False
         Width           =   3495
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   9000
         Visible         =   0   'False
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   1720
         BackColor       =   16744576
         Caption         =   "Correo Electronico"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   6975
         End
      End
   End
End
Attribute VB_Name = "libro_electro04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public cadena As String
 Private FORMATOGRILLA(10, 20) As String
Private sumas(10) As Double
Private suma(10) As Double
Private sumas2(10) As Double
Private sumas3(10) As Double
Private montos(5) As Double
Private lin As Double
Private ANTED2 As Double
Private ANTEH2 As Double
Private sumade As Double
Private sumaha As Double
Private saldodebeTOTAL As Double
Private saldohaberTOTAL  As Double


 

Private Sub Check1_Click()
    chk1.Value = Check1.Value
    chk2.Value = Check1.Value
    chk3.Value = Check1.Value
    chk4.Value = Check1.Value
    chk5.Value = Check1.Value
    chk6.Value = Check1.Value
    
End Sub

 

Private Sub Command1_Click()
Call ENVIO_LIBRO_DIARIO

End Sub

Private Sub Command5_Click()
     Dim TIMBRA As String
        Call Conectartemporal(Servidor, clientesistema + "conta" + empresaactiva, Usuario, password)
        MES = Format(fechasistema, "mm")
        año = Format(fechasistema, "yyyy")
        TIMBRA = "N"
    If chk1.Value = 1 Then
        Call generadiccionario(rutempresa, Format(fechasistema, "yyyy"))
    End If
    If chk2.Value = 1 Then
        grid1.Rows = 1
        lin = 0
        Call CARGAGRILLABALANCE
        Call CARGABALANCE
        Call generaBalance(rutempresa, Format(fechasistema, "yyyy-mm"))
    End If
    If chk3.Value = 1 Then
        grid1.Rows = 1
        lin = 0
        Call CARGAGRILLA
        Call Consulta_Librodiario
        Call generalibrodiario(rutempresa, Format(fechasistema, "yyyy-mm"))
    End If
    If chk4.Value = 1 Then
        Call generalibrodiarioResumen(rutempresa, Format(fechasistema, "yyyy-mm"))
    End If
    If chk5.Value = 1 Then
        grid1.Rows = 1
        lin = 0
        Call CARGAGRILLAMAYOR
        Call leecuentas
        Call generalibroMayor(rutempresa, Format(fechasistema, "yyyy-mm"))
    End If
    If chk6.Value = 1 Then
        Call generalibroMayorResumen(rutempresa, Format(fechasistema, "yyyy-mm"))
    End If
    
    
    
    MsgBox "TODOS LOS ARCHIVOS SE HAN GENERADO CON EXITO", vbInformation, "ATENCION"
End Sub
Sub ENVIO_LIBRO_DIARIO()
Dim FIRMAENVIO As String
Dim dato_c As String
Dim dato_s As String
Dim dato_o As String
Dim dato_p As String
Dim dato_d As String
Dim dato_b As String
Dim dato_i As String
Dim dato_a As String
Dim dato_f As String
Dim dato_n As String
Dim dato_t As String
Dim dato_r As String

dato_c = "-c c:\patricio.pfx "
dato_s = "-s 123 "
dato_o = "-o u:\fae_admin\Documentos\Envios\lceenviolibrodiario_775753404.xml "
dato_p = "-p u:\fae_admin\Documentos\Bases\LceEnvioLibros.xml "
dato_d = "-d u:\ENVIOS_CONTA_SII\LceDiario_77575340-4.xml "
dato_b = "-b u:\ENVIOS_CONTA_SII\LceBalance_77575340-4.xml "
dato_i = "-i u:\ENVIOS_CONTA_SII\LceDiccionario_77575340-4.xml "
dato_a = "-a u:\fae_admin\Documentos\Bases\LceCal.xml "
dato_f = "-f u:\fae_admin\Documentos\Bases\LceCoCertif.xml "
dato_n = "-n False "
dato_t = "-t 0 "
dato_r = "-r 0 "

FIRMAENVIO = lib_envio_libro_diario + dato_c + dato_s + dato_o + dato_p + dato_d + dato_b + dato_i + dato_a + dato_f + dato_n + dato_t + dato_r
Rem Shell FIRMAENVIO, vbMaximizedFocus
Open "c:\envia.txt" For Output As #20
Print #20, FIRMAENVIO
Close #20

End Sub

 
 Sub leecuentas()
 
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    grid1.AutoRedraw = False

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor  "
        csql2.sql = csql2.sql + "WHERE año='" + año + "' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        If Mid(resultados2(0), 5, 4) <> "0000" Then Call LEERMOVIMIENTOS(resultados2(0), resultados2(1))
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        grid1.Column(8).Locked = True
        grid1.Column(9).Locked = True
        grid1.Column(10).Locked = True
    
  
        grid1.AutoRedraw = True
        grid1.Refresh


End Sub
 
 
Public Sub generadiccionario(rutempresa, periodo)
Dim comi As String
Dim NOMBRE As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = contadb

    comi = Chr(34)
    'NOMBRE = periodo
    cadena = " "
    cadena = cadena + "<?xml version=" + comi + "1.0" + comi + " encoding=" + comi + "ISO-8859-1" + comi + Chr(63) + Chr(62)
    'head
    cadena = cadena + "<LceDiccionario xmlns=" + comi + "http://www.sii.cl/SiiLce" + comi + " "
    cadena = cadena & "xmlns:ds=" & comi & "http://www.w3.org/2000/09/xmldsig#" & comi & " "
    cadena = cadena & "xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiLce LceDic_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
    '/Head
    
    'body
    cadena = cadena & " <DocumentoDiccionario ID=" & comi & "Diccionario_" & rutempresa & "_" & periodo & comi & ">"
    cadena = cadena & " <Identificacion>"
    cadena = cadena & " <RutContribuyente>" & rutempresa & "</RutContribuyente>"
    cadena = cadena & " <PeriodoTributario>" & periodo & "</PeriodoTributario>"
    cadena = cadena & " </identificacion>"
    
    
    
   
        csql.sql = "SELECT codigo,nombre,plansii "
        csql.sql = csql.sql + "FROM cuentasdelmayor where año='" + periodo + "' "
        csql.sql = csql.sql + "order by codigo"
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Call cargarcuerpodiccionario(Mid(resultados(0), 1, 1), resultados(0), resultados(1), resultados(2))
                resultados.MoveNext
            Wend
        End If
    
    cadena = cadena & " <RutFirma>" & rut_representante & "</RutFirma>"
    cadena = cadena & " <TmstFirma>" & Format(fechasistema, "yyyy-mm-dd") & "T" & Time & "</TmstFirma>"
    cadena = cadena & " </DocumentoDiccionario>"
    cadena = cadena & " </LceDiccionario>"
    
    Call xml.LoadXml(cadena)
    raiz = "u:\fae_admin\documentos\bases\"
    nombrearchivo = "LceDiccionario_" & rutempresa & ".xml"
    Call xml.SaveXml(raiz + nombrearchivo)
       
    Rem Shell "notepad " + raiz + nombrearchivo

End Sub
Sub cargarcuerpodiccionario(clasificacion, codigocuenta, glosa, codigosii)
    If Right(codigocuenta, 4) <> "0000" Then
        cadena = cadena & " <Cuenta>"
        If clasificacion = 3 Or clasificacion = "4" Then
            cadena = cadena & " <ClasificacionCuenta>" & "4" & "</ClasificacionCuenta>"
        Else
             cadena = cadena & " <ClasificacionCuenta>" & clasificacion & "</ClasificacionCuenta>"
        End If
        
        cadena = cadena & " <CodigoCuenta>" & codigocuenta & "</CodigoCuenta>"
        cadena = cadena & " <Glosa>" & glosa & "</Glosa>"
        If codigosii = "" Then codigosii = "1"
        cadena = cadena & " <CodigoSII>" & codigosii & "</CodigoSII>"
        cadena = cadena & " </Cuenta>"
    End If
End Sub

Public Sub generalibrodiario(rutempresa, periodo)
Dim comi As String
Dim NOMBRE As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = contadb

    comi = Chr(34)
    'NOMBRE = periodo
    cadena = " "
    cadena = cadena + "<?xml version=" + comi + "1.0" + comi + " encoding=" + comi + "ISO-8859-1" + comi + Chr(63) + Chr(62)
    'head
    cadena = cadena + "<LceDiario xmlns=" + comi + "http://www.sii.cl/SiiLce" + comi + " "
    cadena = cadena & "xmlns:ds=" & comi & "http://www.w3.org/2000/09/xmldsig#" & comi & " "
    cadena = cadena & "xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " "
    cadena = cadena & "xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiLce LceDiario_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
    '/Head
    
    'RESUMEN LIBRO DIARIO
     Call generalibrodiarioRes(rutempresa, periodo)
    '/RESUMEN LIBRO DIARIO
     Call cargarDiario
    cadena = cadena & " </LceDiario>"
    
    Call xml.LoadXml(cadena)
    raiz = "u:\fae_admin\documentos\bases\"
    nombrearchivo = "LceDiario_" & rutempresa & ".xml"
    Call xml.SaveXml(raiz + nombrearchivo)
       
    Shell "notepad " + raiz + nombrearchivo

End Sub


Public Sub generalibrodiarioResumen(rutempresa, periodo)
Dim comi As String
Dim NOMBRE As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = contadb

    comi = Chr(34)
    'NOMBRE = periodo
    cadena = " "
    cadena = cadena + "<?xml version=" + comi + "1.0" + comi + " encoding=" + comi + "ISO-8859-1" + comi + Chr(63) + Chr(62)
    'RESUMEN LIBRO DIARIO
     Call generalibrodiarioRes(rutempresa, periodo)
    '/RESUMEN LIBRO DIARIO
   
    
    Call xml.LoadXml(cadena)
    raiz = "u:\fae_admin\documentos\bases\"
    nombrearchivo = "LceDiarioRes_" & rutempresa & ".xml"
    Call xml.SaveXml(raiz + nombrearchivo)
       
    Shell "notepad " + raiz + nombrearchivo

End Sub

Sub cargarDiario()
    Dim k As Double
    Dim filtro As String
    Dim primero As Double
    Dim NOMBRETIPO As String
    
    primero = 1
    For k = 1 To grid1.Rows - 1
             If grid1.Cell(k, 1).text <> "" Then
                If primero = 1 Then
                    cadena = cadena & " <Comprobante>"
                    cadena = cadena & " <TpoComp>" & grid1.Cell(k, 2).text & "</TpoComp>"
                    cadena = cadena & " <NumComp>" & Val(grid1.Cell(k, 3).text) & "</NumComp>"
                    cadena = cadena & " <FechaContable>" & Format(grid1.Cell(k, 1).text, "yyyy-mm-dd") & "</FechaContable>"
                    cadena = cadena & " <GlosaAnalisis>" & grid1.Cell(k, 6).text & "</GlosaAnalisis>"
                End If
                cadena = cadena & " <Movimientos>"
                cadena = cadena & " <CodigoCuenta>" & Replace(grid1.Cell(k, 5).text, ".", "") & "</CodigoCuenta>"
                
                If grid1.Cell(k, 0).text <> "" Then
                    cadena = cadena & " <Rut>" & Val(Mid(grid1.Cell(k, 0).text, 1, 9)) & "-" & Mid(grid1.Cell(k, 0).text, 10, 1) & "</Rut>"
                    cadena = cadena & " <Nombre>" & LEERNOMBREPROVEEDOR(grid1.Cell(k, 0).text) & "</Nombre>"
                End If
                NOMBRETIPO = leerdatos(conta, "maestrotipodedocumentos", "nombredocumento", "tipos='" & grid1.Cell(k, 7).text & "'")
                If NOMBRETIPO = "" Then NOMBRETIPO = grid1.Cell(k, 7).text
                
                cadena = cadena & " <TpoDocum>" & NOMBRETIPO & "</TpoDocum>"
                
                cadena = cadena & " <Numero>" & Val(grid1.Cell(k, 8).text) & "</Numero>"
                
                If grid1.Cell(k, 11).text <> "" Then
                    cadena = cadena & " <Debe>" & grid1.Cell(k, 11).text & "</Debe>"
                Else
                    cadena = cadena & " <Haber>" & grid1.Cell(k, 12).text & "</Haber>"
                End If
                cadena = cadena & " </Movimientos>"
                 primero = primero + 1
             End If
            If grid1.Cell(k, 11).text <> "" And grid1.Cell(k, 12).text <> "" Then
             If grid1.Cell(k, 1).text = "" And grid1.Cell(k, 11).text <> "" Then
                cadena = cadena & " <ValorComprobante>" & grid1.Cell(k, 12).text & "</ValorComprobante>"
                cadena = cadena & " </Comprobante>"
                primero = 1
             End If
          End If
    Next k
    
    
End Sub
Public Sub generalibrodiarioRes(rutempresa, periodo)
Dim comi As String
Dim NOMBRE As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim totalcomprobantes As Double
Dim totalmovimientos As Double
Dim totalvalor As Double
    
Set csql.ActiveConnection = contadb

    comi = Chr(34)
    'NOMBRE = periodo
    'head
    cadena = cadena + "<LceDiarioRes xmlns=" + comi + "http://www.sii.cl/SiiLce" + comi + " "
    cadena = cadena & "xmlns:ds=" & comi & "http://www.w3.org/2000/09/xmldsig#" & comi & " "
    cadena = cadena & "xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " "
    cadena = cadena & "xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiLce LceDiarioRes_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
    '/Head
    
    'body
    cadena = cadena & " <DocumentoDiarioRes ID=" & comi & "DiarioRes_" & rutempresa & "_" & periodo & comi & ">"
    cadena = cadena & " <Identificacion>"
    cadena = cadena & " <RutContribuyente>" & rutempresa & "</RutContribuyente>"
    cadena = cadena & " <PeriodoTributario>"
    cadena = cadena & " <Inicial>" & periodo & "</Inicial>"
    cadena = cadena & " <Final>" & periodo & "</Final>"
    cadena = cadena & " </PeriodoTributario>"
    cadena = cadena & " </identificacion>"
       
    csql.sql = "SELECT fecha, (SELECT COUNT( DISTINCT numero) FROM movimientoscontables WHERE m2.fecha=fecha GROUP BY fecha) AS cantidadcomprobantes,"
    csql.sql = csql.sql & "COUNT(linea) AS cantidadmovimientos,SUM(IF(dh='D',monto,0)) AS totalcomprobante FROM movimientoscontables AS m2  "
    csql.sql = csql.sql & "WHERE año='" & Format(fechasistema, "yyyy") & "' "
    Rem csql.sql = csql.sql & " and fecha='2013-12-01' "
    csql.sql = csql.sql & "GROUP BY fecha ORDER BY fecha,numero,linea"
    csql.Execute
    totalvalor = 0
    totalcomprobantes = 0
    totalmovimientos = 0
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Call cargarDiarioRes(resultados(0), resultados(1), resultados(2), resultados(3))
            totalcomprobantes = totalcomprobantes + resultados(1)
            totalmovimientos = totalmovimientos + resultados(2)
            totalvalor = totalvalor + resultados(3)
            resultados.MoveNext
        Wend
    End If
    csql.Close
    Set csql = Nothing
        cadena = cadena & " <Cierre>"
        cadena = cadena & " <CantidadComprobantes>" & totalcomprobantes & "</CantidadComprobantes>"
        cadena = cadena & " <CantidadMovimientos>" & totalmovimientos & "</CantidadMovimientos>"
        cadena = cadena & " <SumaValorComprobante>" & totalvalor & "</SumaValorComprobante>"
        cadena = cadena & " <ValorAcumulado>" & totalvalor & "</ValorAcumulado>"
        cadena = cadena & " </Cierre>"
    
    
    cadena = cadena & " <RutFirma>" & rut_representante & "</RutFirma>"
    cadena = cadena & " <TmstFirma>" & Format(fechasistema, "yyyy-mm-dd") & "T" & Time & "</TmstFirma>"
    cadena = cadena & " </DocumentoDiarioRes>"
    cadena = cadena & " </LceDiarioRes>"
    
    
End Sub

Sub cargarDiarioRes(FECHACONTABLE, cantidadcomprobantes, cantidadmovimientos, totalcomprobantes)
        cadena = cadena & " <RegistroDiario>"
        cadena = cadena & " <FechaContable>" & Format(FECHACONTABLE, "yyyy-mm-dd") & "</FechaContable>"
        cadena = cadena & " <CantidadComprobantes>" & cantidadcomprobantes & "</CantidadComprobantes>"
        cadena = cadena & " <CantidadMovimientos>" & cantidadmovimientos & "</CantidadMovimientos>"
        cadena = cadena & " <SumaValorComprobante>" & totalcomprobantes & "</SumaValorComprobante>"
        cadena = cadena & " </RegistroDiario>"
End Sub




Sub Consulta_Librodiario()
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fechacom As String
    Dim lineascom As Double
   
    
    
    Dim totales(31, 3) As Variant
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,tiposii,rutproveedor "
        csql.sql = csql.sql + "FROM movimientoscontables ," + clientesistema + "conta.maestrotipodedocumentos as td WHERE mes='" + MES + "' and año='" + año + "' and tipos=tipo "
        Rem csql.sql = csql.sql & " and fecha='2013-12-01' "
        csql.sql = csql.sql + "order by fecha,tiposii,numero,linea "
        csql.Execute
        
        
        grid1.AutoRedraw = False
'        Barra.Max = csql.RowsAffected + 1
       lineascom = 0
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        lin = 0: PASO = resultados(1) + resultados(2)
        fechacom = resultados(0)
         
         While Not resultados.EOF
          lin = lin + 1
             lineascom = lineascom + 1
             grid1.Rows = grid1.Rows + 1
             If resultados(1) + resultados(2) <> PASO Then
             Call totalcomprobante(lin)
             totales(Format(fechacom, "dd"), 1) = totales(Format(fechacom, "dd"), 1) + 1
             totales(Format(fechacom, "dd"), 2) = totales(Format(fechacom, "dd"), 2) + lineascom
             lineascom = 0
             fechacom = resultados(0)
             End If
             grid1.Cell(lin, 0).text = resultados("rutproveedor")
             For k = 0 To 9
             grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
             
             If resultados(11) = "D" Then
             totales(Format(fechacom, "dd"), 3) = totales(Format(fechacom, "dd"), 3) + resultados(10)
             grid1.Cell(lin, 11).text = resultados(10)
             anted = anted + resultados(10)
             ANTED2 = ANTED2 + resultados(10)
             Else
             grid1.Cell(lin, 12).text = resultados(10)
             anteh = anteh + resultados(10)
             ANTEH2 = ANTEH2 + resultados(10)
             End If
             PASO = resultados(1) + resultados(2)
             resultados.MoveNext

           
         Wend
          grid1.Rows = grid1.Rows + 1
          lin = lin + 1
          Call totalcomprobante(lin)
          totales(Format(fechacom, "dd"), 1) = totales(Format(fechacom, "dd"), 1) + 1
           totales(Format(fechacom, "dd"), 2) = totales(Format(fechacom, "dd"), 2) + lineascom
             
             lineascom = 0
            
             
'          Grid1.Rows = Grid1.Rows + 1
'          lin = lin + 1
'          Call totalcomprobante2(lin)
          
          resultados.Close
            Set resultados = Nothing

        End If
'        If xmllibrodiario = True Then
'        For k = 1 To 31
'          Grid1.Rows = Grid1.Rows + 1
'             lin = Grid1.Rows - 1
'            Grid1.Cell(lin, 7).text = k
'            Grid1.Cell(lin, 8).text = totales(k, 1)
'            Grid1.Cell(lin, 11).text = totales(k, 2)
'            Grid1.Cell(lin, 12).text = totales(k, 3)
'
'          Next k
'
'        End If
grid1.AutoRedraw = True
grid1.Refresh

End Sub
Sub CARGAGRILLA()
    Dim FORMATOGRILLA(40, 40) As String
    
Rem DATOS DE LA COLUMNA
    grid1.DefaultFont.Size = 7.5
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "NL"
    FORMATOGRILLA(1, 5) = "CUENTA"
    FORMATOGRILLA(1, 6) = "GLOSA"
    FORMATOGRILLA(1, 7) = "TP"
    FORMATOGRILLA(1, 8) = "NUMERO"
    FORMATOGRILLA(1, 9) = "EMISION"
    FORMATOGRILLA(1, 10) = "VENCIMIENTO"
    FORMATOGRILLA(1, 11) = "DEBE"
    FORMATOGRILLA(1, 12) = "HABER"
    FORMATOGRILLA(1, 13) = "NOMBRE CUENTA"
    FORMATOGRILLA(1, 14) = "CUENTA CORRIENTE"
    FORMATOGRILLA(1, 15) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "3"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "25"
    FORMATOGRILLA(2, 7) = "3"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "0"
    FORMATOGRILLA(2, 10) = "0"
    FORMATOGRILLA(2, 11) = "12"
    FORMATOGRILLA(2, 12) = "12"
    FORMATOGRILLA(2, 13) = "30"
    FORMATOGRILLA(2, 14) = "30"
    FORMATOGRILLA(2, 15) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "D"
    FORMATOGRILLA(3, 10) = "D"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "S"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 11) = "###,###,###,###"
    FORMATOGRILLA(4, 12) = "###,###,###,###"
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    
    grid1.Cols = 13
    grid1.Rows = 2
    
     'grid1.AllowUserResizing = False
    grid1.DisplayFocusRect = False
    'grid1.ExtendLastCol = True
    grid1.BoldFixedCell = False
    
    grid1.DrawMode = cellOwnerDraw
    
    grid1.Appearance = Flat
    grid1.ScrollBarStyle = Flat
    grid1.FixedRowColStyle = Flat
    
   'grid1.BackColorFixed = RGB(90, 158, 214)
   ' grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' grid1.BackColorBkg = RGB(90, 158, 214)
   ' grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' grid1.BackColor1 = RGB(231, 235, 247)
   ' grid1.BackColor2 = RGB(239, 243, 255)
   ' grid1.GridColor = RGB(148, 190, 231)
    grid1.Column(0).Width = 0
    
    For k = 1 To grid1.Cols - 1
        
        grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * grid1.DefaultFont.Size
        
        
        grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub
Sub CARGAGRILLAMAYOR()
Rem DATOS DE LA COLUMNA
    grid1.DefaultFont.Size = 7
    
    
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "LINEA"
    FORMATOGRILLA(1, 5) = "CUENTA"
    FORMATOGRILLA(1, 6) = "GLOSA"
    FORMATOGRILLA(1, 7) = "TP"
    FORMATOGRILLA(1, 8) = "NUMERO"
    FORMATOGRILLA(1, 9) = "EMISION"
    FORMATOGRILLA(1, 10) = "VENCIMIENTO"
    FORMATOGRILLA(1, 11) = "DEBE"
    FORMATOGRILLA(1, 12) = "HABER"
    FORMATOGRILLA(1, 13) = "SALDO"
    FORMATOGRILLA(1, 14) = "NOMBRE CUENTA"
    FORMATOGRILLA(1, 15) = "CUENTA CORRIENTE"
    FORMATOGRILLA(1, 16) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "4"
    FORMATOGRILLA(2, 5) = "0"
    FORMATOGRILLA(2, 6) = "30"
    FORMATOGRILLA(2, 7) = "3"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "0"
    FORMATOGRILLA(2, 10) = "10"
    FORMATOGRILLA(2, 11) = "11"
    FORMATOGRILLA(2, 12) = "11"
    FORMATOGRILLA(2, 13) = "12"
    FORMATOGRILLA(2, 14) = "30"
    FORMATOGRILLA(2, 15) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "D"
    FORMATOGRILLA(3, 10) = "D"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 11) = "###,###,###,##0"
    FORMATOGRILLA(4, 12) = "###,###,###,##0"
    FORMATOGRILLA(4, 13) = "###,###,###,##0"
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    
    grid1.Cols = 14
    grid1.Rows = 2
    
     'grid1.AllowUserResizing = False
    grid1.DisplayFocusRect = False
    'grid1.ExtendLastCol = True
    grid1.BoldFixedCell = False
    
    grid1.DrawMode = cellOwnerDraw
    
    grid1.Appearance = Flat
    grid1.ScrollBarStyle = Flat
    grid1.FixedRowColStyle = Flat
    grid1.Column(0).Width = 0
    For k = 1 To grid1.Cols - 1
        grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * grid1.DefaultFont.Size
        grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then grid1.Column(k).CellType = cellCalendar
    Next k
End Sub
Sub totalcomprobante(row)
    
    With grid1.Range(row, 11, row, 12)
    .Borders(cellEdgeTop) = cellThin
    
     End With
   With grid1.Range(row, 1, row, 12)
   .FontBold = True
    .FontUnderline = True
    End With
    
    
    
    grid1.Cell(row, 10).CellType = cellTextBox
    
    
    grid1.Cell(row, 10).text = "TOTAL "
    grid1.Cell(row, 11).text = anted
    grid1.Cell(row, 12).text = anteh
    lin = lin + 2
             grid1.Rows = grid1.Rows + 2
        
        anted = 0: anteh = 0
    End Sub
Sub totalcomprobante2(row)
    
    With grid1.Range(row, 11, row, 12)
    
    .Borders(cellEdgeTop) = cellThin
    
    
    
     End With
   With grid1.Range(row, 1, row, 12)
   .FontBold = True
    .FontUnderline = True
    End With
    
    
    
    grid1.Cell(row, 10).CellType = cellTextBox
    
    
    grid1.Cell(row, 10).text = "TOTAL GENERAL"
    grid1.Cell(row, 11).text = ANTED2
    grid1.Cell(row, 12).text = ANTEH2
    lin = lin + 2
             grid1.Rows = grid1.Rows + 2
        
        ANTED2 = 0: ANTEH2 = 0
    End Sub
    
Public Sub generalibroMayor(rutempresa, periodo)
Dim comi As String
Dim NOMBRE As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = contadb

    comi = Chr(34)
    'NOMBRE = periodo
    cadena = " "
    cadena = cadena + "<?xml version=" + comi + "1.0" + comi + " encoding=" + comi + "ISO-8859-1" + comi + Chr(63) + Chr(62)
    'head
    cadena = cadena + "<LceMayor xmlns=" + comi + "http://www.sii.cl/SiiLce" + comi + " "
    cadena = cadena & "xmlns:ds=" & comi & "http://www.w3.org/2000/09/xmldsig#" & comi & " "
    cadena = cadena & "xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " "
    cadena = cadena & "xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiLce LceMayor_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
    '/Head
    
    'RESUMEN LIBRO MAYOR
     Call generalibroMayorRes(rutempresa, periodo)
    '/RESUMEN LIBRO MAYOR
      Call cargarMayor
     
    cadena = cadena & " </LceMayor>"
    
    Call xml.LoadXml(cadena)
    raiz = "u:\fae_admin\documentos\bases\"
    nombrearchivo = "LceMayor_" & rutempresa & ".xml"
    Call xml.SaveXml(raiz + nombrearchivo)
    Shell "notepad " + raiz + nombrearchivo

End Sub

Public Sub generalibroMayorRes(rutempresa, periodo)
Dim comi As String
Dim NOMBRE As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim saldodebe As Double
Dim saldohaber As Double


Set csql.ActiveConnection = contadb

    comi = Chr(34)
    'NOMBRE = periodo
    'head
    cadena = cadena + "<LceMayorRes xmlns=" + comi + "http://www.sii.cl/SiiLce" + comi + " "
    cadena = cadena & "xmlns:ds=" & comi & "http://www.w3.org/2000/09/xmldsig#" & comi & " "
    cadena = cadena & "xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " "
    cadena = cadena & "xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiLce LceMayorRes_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
    '/Head
    
    'body
    cadena = cadena & " <DocumentoMayorRes ID=" & comi & "MayorRes_" & rutempresa & "_" & periodo & comi & ">"
    cadena = cadena & " <Identificacion>"
    cadena = cadena & " <RutContribuyente>" & rutempresa & "</RutContribuyente>"
    cadena = cadena & " <PeriodoTributario>"
    cadena = cadena & " <Inicial>2012-01</Inicial>"
    cadena = cadena & " <Final>2012-12</Final>"
    cadena = cadena & " </PeriodoTributario>"
    cadena = cadena & " </identificacion>"
       
    csql.sql = " SELECT codigocuenta,COUNT(linea) AS cantidadmovimientos,SUM(IF(dh='D',monto,0)) AS totalDebe, "
    csql.sql = csql.sql & "SUM(IF(dh='H',monto,0)) AS totalHaber,SUM(IF(dh='D',monto,0))-SUM(IF(dh='H',monto,0)) AS Saldo "
    csql.sql = csql.sql & "FROM movimientoscontables AS m2  "
    Rem POR SI LO PIDEN DIFERENCIANDO
    csql.sql = csql.sql & "WHERE año='" & Format(fechasistema, "yyyy") & "' "
    csql.sql = csql.sql & "GROUP BY codigocuenta ORDER BY codigocuenta "
    
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            saldodebe = 0
            saldohaber = 0
            If resultados(4) > 0 Then saldodebe = resultados(4)
            If resultados(4) < 0 Then saldohaber = resultados(4) * -1
            
            Call cargarMayorRes(resultados(0), resultados(1), resultados(2), resultados(3), saldodebe, saldohaber, resultados(2), resultados(3), saldodebe, saldohaber)
            resultados.MoveNext
        Wend
    End If
    csql.Close
    Set csql = Nothing
    
    cadena = cadena & " <RutFirma>" & rut_representante & "</RutFirma>"
    cadena = cadena & " <TmstFirma>" & Format(fechasistema, "yyyy-mm-dd") & "T" & Time & "</TmstFirma>"
    cadena = cadena & " </DocumentoMayorRes>"
    cadena = cadena & " </LceMayorRes>"
End Sub

Sub cargarMayorRes(codigocuenta, cantidadmovimientos, Debe1, Haber1, saldoDeudor1, saldoAcreedor1, Debe2, Haber2, saldoDeudor2, saldoAcreedor2)
        Dim saldofinal As Double
        
        cadena = cadena & " <Cuenta>"
        cadena = cadena & " <CodigoCuenta>" & codigocuenta & "</CodigoCuenta>"
        cadena = cadena & " <CantidadMovimientos>" & cantidadmovimientos & "</CantidadMovimientos>"
        cadena = cadena & " <Cierre>"
        cadena = cadena & " <MontosPeriodo>"
        
        If Debe1 >= 0 Then
            cadena = cadena & " <Debe>" & Debe1 & "</Debe>"
        End If
        If Haber1 > 0 Then
            cadena = cadena & " <Haber>" & Haber1 & "</Haber>"
        End If
        If saldoDeudor1 >= 0 Then
            cadena = cadena & " <SaldoDeudor>" & saldoDeudor1 & "</SaldoDeudor>"
        End If
        
        If saldoAcreedor1 < 0 Then
            cadena = cadena & " <SaldoAcreedor>" & saldoAcreedor1 & "</SaldoAcreedor>"
        End If
        
        cadena = cadena & " </MontosPeriodo>"
        
  Rem BUSCA TOTALES
        Call SaldosAnuales(codigocuenta)
        
        cadena = cadena & " <MontosAcumulado>"
        
        If saldodebeTOTAL > 0 Then
            cadena = cadena & " <Debe>" & saldodebeTOTAL & "</Debe>"
        End If
        If saldohaberTOTAL > 0 Then
        cadena = cadena & " <Haber>" & saldohaberTOTAL & "</Haber>"
        End If
        saldofinal = saldodebeTOTAL - saldohaberTOTAL
        If saldofinal >= 0 Then
            cadena = cadena & " <SaldoDeudor>" & saldofinal & "</SaldoDeudor>"
        End If
        If saldofinal < 0 Then
            cadena = cadena & " <SaldoAcreedor>" & saldofinal * -1 & "</SaldoAcreedor>"
        End If
        cadena = cadena & " </MontosAcumulado>"
        cadena = cadena & " </Cierre>"
        cadena = cadena & " </Cuenta>"
        
End Sub

Sub cargarMayor()
    Dim k As Double
    Dim filtro As String
    Dim primero As Double
    primero = 1
    For k = 1 To grid1.Rows - 1
             If grid1.Cell(k, 1).text <> "" Then
                If primero = 1 Then
                    cadena = cadena & " <Cuenta>"
                    cadena = cadena & " <CodigoCuenta>" & Replace(grid1.Cell(k, 0).text, ".", "") & "</CodigoCuenta>"
                    GoTo no:
                 End If
                    cadena = cadena & " <Movimientos>"
                    cadena = cadena & " <TpoComp>" & grid1.Cell(k, 2).text & "</TpoComp>"
                    cadena = cadena & " <NumComp>" & Val(grid1.Cell(k, 3).text) & "</NumComp>"
                    cadena = cadena & " <FechaContable>" & Format(grid1.Cell(k, 1).text, "yyyy-mm-dd") & "</FechaContable>"
                    cadena = cadena & " <GlosaAnalisis>" & grid1.Cell(k, 6).text & "</GlosaAnalisis>"
                    If grid1.Cell(k, 11).text <> "" Then
                        cadena = cadena & " <Debe>" & grid1.Cell(k, 11).text & "</Debe>"
                    Else
                        cadena = cadena & " <Haber>" & grid1.Cell(k, 12).text & "</Haber>"
                    End If
                    cadena = cadena & " </Movimientos>"
no:
                  primero = primero + 1
             End If
             If grid1.Cell(k, 1).text = "" And grid1.Cell(k, 11).text = "" Then
                cadena = cadena & " </Cuenta>"
                primero = 1
             End If
           
    Next k
End Sub

Sub cargarBalance()
    Dim k As Double
    Dim filtro As String
    Dim primero As Double
    primero = 1
    For k = 1 To grid1.Rows - 1
        cadena = cadena & " <Cuenta>"
        cadena = cadena & " <CodigoCuenta>" & Replace(grid1.Cell(k, 1).text, ".", "") & "</CodigoCuenta>"
        cadena = cadena & " <Debe>" & grid1.Cell(k, 3).text & "</Debe>"
        cadena = cadena & " <Haber>" & grid1.Cell(k, 4).text & "</Haber>"
        If Val(grid1.Cell(k, 5).text) > 0 Then
            cadena = cadena & " <SaldoDeudor>" & grid1.Cell(k, 5).text & "</SaldoDeudor>"
        End If
        If Val(grid1.Cell(k, 6).text) > 0 Then
            cadena = cadena & " <SaldoAcreedor>" & grid1.Cell(k, 6).text & "</SaldoAcreedor>"
        End If
        cadena = cadena & " </Cuenta>"
           
    Next k
End Sub

Sub LEERMOVIMIENTOS(cuenta, NOMBRE)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,rutctacte "
        csql.sql = csql.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and mes='" + MES + "' and año='" + año + "' "
        csql.sql = csql.sql + "order by fecha"
        csql.Execute
       
        lin = lin + 1
        grid1.Rows = grid1.Rows + 1
        Call DATOSSALDOS(cuenta)
        For k = 1 To 6
        grid1.Column(k).Locked = False
        Next k
        
        grid1.Range(lin, 1, lin, 12).FontBold = True
        grid1.Range(lin, 1, lin, 12).FontUnderline = True
        
        
        
        
        grid1.Range(lin, 1, lin, 6).Merge
        
        grid1.Cell(lin, 1).CellType = cellTextBox
        
        grid1.Cell(lin, 10).CellType = cellTextBox
        grid1.Cell(lin, 0).text = cuenta
        grid1.Cell(lin, 1).text = NOMBRE
        grid1.Cell(lin, 10).text = "SALDO-->"
        
        grid1.Cell(lin, 13).text = saldo

        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
          lin = lin + 1
             grid1.Rows = grid1.Rows + 1
             For k = 0 To 9
             grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             If resultados(11) = "D" Then grid1.Cell(lin, 11).text = resultados(10): anted = anted + resultados(10): saldo = saldo + resultados(10)
             If resultados(11) = "H" Then grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(10): saldo = saldo - resultados(10)
             grid1.Cell(lin, 13).text = saldo
'            If Check1.Value = 1 Then
'            Grid1.Cell(lin, 14).text = resultados(12)
'            End If
             resultados.MoveNext
           
         Wend
          lin = lin + 1
             grid1.Rows = grid1.Rows + 1
         
         'Call totalcomprobante(lin, infogrilla)
          resultados.Close
            Set resultados = Nothing
        Else
            Call grid1.RemoveItem(lin)
            lin = lin - 1
        End If

End Sub

Sub totalcomprobanteMayor(row)
    'Grid1.Range(Row, 1, Row, 12).FontBold = True
    grid1.Range(row, 1, row, 12).FontUnderline = True
        
    
    grid1.Range(row, 11, row, 12).Borders(cellEdgeTop) = cellThin
    grid1.Cell(row, 10).CellType = cellTextBox
    grid1.Cell(row, 10).text = "TOTAL "
    grid1.Cell(row, 11).text = anted
    grid1.Cell(row, 12).text = anteh
    lin = lin + 2
             grid1.Rows = grid1.Rows + 2
        
        anted = 0: anteh = 0: saldo = 0
    End Sub
    
Sub LEERSALDOS(cuenta)
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
    
    condicion = "codigo=" + "'" + cuenta + "' and año='" + año + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
 '   If sqlconta.status = 4 Then Stop

End Sub
Sub DATOSSALDOS(cuenta)
 

Call LEERSALDOS(cuenta)
sumador = Val(sqlconta.response(2, 3)) - Val(sqlconta.response(3, 3))
For k = 1 To Val(MES) - 1
sumador = sumador + Val(sqlconta.response(k + 3, 3)) - Val(sqlconta.response(k + 15, 3))
 
Next k
saldo = sumador
End Sub
Sub SaldosAnuales(cuenta)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT codigocuenta,SUM(IF(dh='D',monto,0)),SUM(IF(dh='H',monto,0)) "
    csql.sql = csql.sql & "FROM movimientoscontables WHERE fecha "
    csql.sql = csql.sql & "BETWEEN '" & año & "-01-01" & "' AND '" & Format(fechasistema, "yyyy-mm-dd") & "' "
    csql.sql = csql.sql & "AND codigocuenta='" & cuenta & "' GROUP BY codigocuenta "
    csql.Execute
    saldodebeTOTAL = 0
    saldohaberTOTAL = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        saldodebeTOTAL = resultados(1)
        saldohaberTOTAL = resultados(2)
    End If
    csql.Close
    Set csql = Nothing
    
End Sub
Public Sub generalibroMayorResumen(rutempresa, periodo)
Dim comi As String
Dim NOMBRE As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = contadb

    comi = Chr(34)
    'NOMBRE = periodo
    cadena = " "
    cadena = cadena + "<?xml version=" + comi + "1.0" + comi + " encoding=" + comi + "ISO-8859-1" + comi + Chr(63) + Chr(62)
    'RESUMEN LIBRO DIARIO
     Call generalibroMayorRes(rutempresa, periodo)
    '/RESUMEN LIBRO DIARIO
   
    
    Call xml.LoadXml(cadena)
    raiz = "u:\fae_admin\documentos\bases\"
    nombrearchivo = "LceMayorRes_" & rutempresa & ".xml"
    Call xml.SaveXml(raiz + nombrearchivo)
       
    Shell "notepad " + raiz + nombrearchivo
End Sub

Public Sub generaBalance(rutempresa, periodo)
Dim comi As String
Dim NOMBRE As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Set csql.ActiveConnection = contadb

    comi = Chr(34)
    'NOMBRE = periodo
    cadena = " "
    cadena = cadena + "<?xml version=" + comi + "1.0" + comi + " encoding=" + comi + "ISO-8859-1" + comi + Chr(63) + Chr(62)
    'head
    cadena = cadena + "<LceBalance xmlns=" + comi + "http://www.sii.cl/SiiLce" + comi + " "
    cadena = cadena & "xmlns:ds=" & comi & "http://www.w3.org/2000/09/xmldsig#" & comi & " "
    cadena = cadena & "xmlns:xsi=" + comi + "http://www.w3.org/2001/XMLSchema-instance" + comi + " "
    cadena = cadena & "xsi:schemaLocation=" + comi + "http://www.sii.cl/SiiLce LceBalance_v10.xsd" + comi + " version=" + comi + "1.0" + comi + ">"
    '/Head
    
    cadena = cadena & " <DocumentoBalance ID=" & comi & "Balance_" & rutempresa & "_" & periodo & comi & ">"
    cadena = cadena & " <Identificacion>"
    cadena = cadena & " <RutContribuyente>" & rutempresa & "</RutContribuyente>"
    cadena = cadena & " <PeriodoTributario>" & periodo & "</PeriodoTributario>"
    cadena = cadena & " </Identificacion>"
    Call cargarBalance
    
    cadena = cadena & " <RutFirma>" & rut_representante & "</RutFirma>"
    cadena = cadena & " <TmstFirma>" & Format(fechasistema, "yyyy-mm-dd") & "T" & Time & "</TmstFirma>"
    cadena = cadena & " </DocumentoBalance>"
    cadena = cadena & " </LceBalance>"
    
    Call xml.LoadXml(cadena)
    raiz = "u:\fae_admin\documentos\bases\"
    nombrearchivo = "LceBalance_" & rutempresa & ".xml"
    Call xml.SaveXml(raiz + nombrearchivo)
    Shell "notepad " + raiz + nombrearchivo

End Sub

Sub CARGAGRILLABALANCE()
Rem DATOS DE LA COLUMNA
    
    
    FORMATOGRILLA(1, 1) = " CODIGO "
    FORMATOGRILLA(1, 2) = " CUENTA         "
    FORMATOGRILLA(1, 3) = "DEBITOS"
    FORMATOGRILLA(1, 4) = "CREDITOS"
    FORMATOGRILLA(1, 5) = "DEUDOR"
    FORMATOGRILLA(1, 6) = "ACREEDOR"
    FORMATOGRILLA(1, 7) = " ACTIVO"
    FORMATOGRILLA(1, 8) = "PASIVO"
    FORMATOGRILLA(1, 9) = "PERDIDA"
    FORMATOGRILLA(1, 10) = "GANANCIA"
    Rem LARGO DE LOS DATOS
'    If scodigo.Value = True Then
'        FormatoGrilla(2, 1) = "0"
'    Else
        FORMATOGRILLA(2, 1) = "8"
'    End If
    FORMATOGRILLA(2, 2) = "28"
    FORMATOGRILLA(2, 3) = "12"
    FORMATOGRILLA(2, 4) = "12"
    FORMATOGRILLA(2, 5) = "11"
    FORMATOGRILLA(2, 6) = "11"
    FORMATOGRILLA(2, 7) = "11"
    FORMATOGRILLA(2, 8) = "11"
    FORMATOGRILLA(2, 9) = "11"
    FORMATOGRILLA(2, 10) = "11"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = "###,###,###,###"
    FORMATOGRILLA(4, 4) = "###,###,###,###"
    FORMATOGRILLA(4, 5) = "###,###,###,###"
    FORMATOGRILLA(4, 6) = "###,###,###,###"
    FORMATOGRILLA(4, 7) = "###,###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###,###"
    FORMATOGRILLA(4, 10) = "###,###,###,###"
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    
    
    grid1.Cols = 11
    grid1.Rows = 2
    
     'GRID1.AllowUserResizing = False
    grid1.DisplayFocusRect = False
    'GRID1.ExtendLastCol = True
    grid1.BoldFixedCell = False
    
    grid1.DrawMode = cellOwnerDraw
    
    grid1.Appearance = Flat
    grid1.ScrollBarStyle = Flat
    grid1.FixedRowColStyle = Flat
    
   'GRID1.BackColorFixed = RGB(90, 158, 214)
   ' GRID1.BackColorFixedSel = RGB(110, 180, 230)
   ' GRID1.BackColorBkg = RGB(90, 158, 214)
   ' GRID1.BackColorScrollBar = RGB(231, 235, 247)
   ' GRID1.BackColor1 = RGB(231, 235, 247)
   ' GRID1.BackColor2 = RGB(239, 243, 255)
   ' GRID1.GridColor = RGB(148, 190, 231)
    grid1.Column(0).Width = 0
    
    For k = 1 To grid1.Cols - 1
        
        grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * grid1.DefaultFont.Size
        
        
        
        grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Sub CARGABALANCE()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim lin As Double
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo "
        csql.sql = csql.sql + "FROM cuentasdelmayor "
        csql.sql = csql.sql + "WHERE año='" + año + "' "
'        If Option1.Value = True Then
        csql.sql = csql.sql + " and mid(codigo,5,4)<>'0000' "
'        Else
'        csql.sql = csql.sql + " and mid(codigo,5,4)='0000' and mid(codigo,3,2)<>'00' "
'
'        End If
        
        
        csql.sql = csql.sql + "order by codigo,año "
        csql.Execute
        lin = 0
        For k = 1 To 8
        sumas(k) = 0
        sumas2(k) = 0
        sumas3(k) = 0
        sumast(k) = 0
        
        Next k
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
                While Not resultados.EOF
                    
                    Call LEERSALDOSBalance(resultados(0), resultados(2))
                            If suma(1) + suma(2) <> 0 Then
                            lin = lin + 1
                            grid1.Rows = lin + 1
                             
                            grid1.Cell(lin, 1).text = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4)
                            grid1.Cell(lin, 2).text = resultados(1)
                            For k = 1 To 8
                            grid1.Cell(lin, k + 2).text = suma(k)
                            Next k
                            End If
                    
                resultados.MoveNext
                Wend
'            Call totales
            
            resultados.Close
            
            Set resultados = Nothing
        End If
End Sub


Sub LEERSALDOSBalance(LLAVE, tipo)
Dim SUMD As Double
Dim SUMH As Double
Dim anted As Double
Dim anteh As Double
Dim DIFE As Double
Dim fechaproceso As String


    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = ""
    
    condicion = "codigo=" + "'" + LLAVE + "' and año ='" + año + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    
    Call sqlconta.sqlconta(op, condicion)
 '   If sqlconta.status = 4 Then Stop
    anted = sqlconta.response(2, 3)
    anteh = sqlconta.response(3, 3)
    Rem anted = 0
     Rem anteh = 0
    fechaproceso = DateSerial(año, MES + 1, 0)
    
    
    
   Call LEERSALDOSMAYOR(LLAVE, Format(fechaproceso, "yyyy-mm-dd"))
   Rem  sumade = 0
    Rem sumaha = 0
    SUMD = sumade: SUMH = sumaha

For k = 1 To 8
suma(k) = 0
sumas2(k) = 0
sumas3(k) = 0
Next k

suma(1) = anted + SUMD
suma(2) = anteh + SUMH
DIFE = suma(1) - suma(2)

If DIFE > 0 Then suma(3) = DIFE
If DIFE < 0 Then suma(4) = DIFE * -1


If tipo = "1" Or tipo = "2" Then suma(5) = suma(3): suma(6) = suma(4)

If tipo <> "1" And tipo <> "2" Then suma(7) = suma(3): suma(8) = suma(4)
For k = 1 To 8
sumas(k) = sumas(k) + suma(k)
Next k

End Sub
Sub LEERSALDOSMAYOR(codigo, fecha)
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim fecha1 As String
    Dim fecha2 As String
    Dim resultados As rdoResultset
  Dim NIVEL As String
  
    fecha1 = Format(fecha, "yyyy") + "-01-01"
    fecha2 = Format(fecha, "yyyy-mm-dd")
        Set csql2.ActiveConnection = contadb
       NIVEL = "3"
        If Mid(codigo, 5, 5) = "0000" Then NIVEL = "2"
        If Mid(codigo, 3, 6) = "000000" Then NIVEL = "1"
        csql2.sql = "SELECT fecha,sum(monto),dh "
        csql2.sql = csql2.sql + "FROM movimientoscontables WHERE fecha between '" + fecha1 + "' and '" + fecha2 + "' "
        If NIVEL = "1" Then
        csql2.sql = csql2.sql + "and mid(codigocuenta,1,2)='" + Mid(codigo, 1, 2) + "' "
        End If
        If NIVEL = "2" Then
        csql2.sql = csql2.sql + "and mid(codigocuenta,1,4)='" + Mid(codigo, 1, 4) + "' "
        End If
        If NIVEL = "3" Then
        csql2.sql = csql2.sql + "and codigocuenta='" + codigo + "' "
        End If
        
        
        csql2.sql = csql2.sql + " group by dh "
        csql2.Execute
        LINEAS = 0
        sumade = 0: sumaha = 0
        If csql2.RowsAffected > 0 Then
         
        Set resultados = csql2.OpenResultset
        While Not resultados.EOF
        If resultados(2) = "D" Then
        sumade = resultados(1)
        Else
        sumaha = resultados(1)
        End If
        
        
        
        resultados.MoveNext
        Wend
          
          resultados.Close
            Set resultados = Nothing
        End If


  
End Sub

Private Sub Form_Load()
  Call librerias_java_conta
End Sub

