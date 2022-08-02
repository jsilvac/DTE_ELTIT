Attribute VB_Name = "registro"
Public lc_mescontable As String
Public lc_anocontable As String
Public lc_neto As String
Public lc_iva As String
Public lc_total As String
Public estanoaceptado As Boolean
  Public lc_tipodte As String
    Public lc_folio As String
    Public lc_fchemis As String
    Public lc_rut As String
    Public lc_prove As String
    Public lc_dire As String
    Public lc_recep As String
Public lc_tipo(10) As String
Public lc_monto(10) As String

Public Function esta_en_libro_compras(tipo, numero, rut, monto, empresa, MES, año) As Boolean
Dim multi As Integer
Dim empresafae As String

If tipo = "33" Then tipo = "4"
If tipo = "56" Then tipo = "5"
If tipo = "61" Then tipo = "6"
If tipo = "34" Then tipo = "0"
If tipo = "30" Then tipo = "1"
If tipo = "32" Then tipo = "9"
If tipo = "60" Then tipo = "3"
If tipo = "46" Then tipo = "7"
If tipo = "914" Then tipo = "8"


Rem On Error GoTo salida:

rut = Format(Mid(rut, 1, Len(rut) - 2), "000000000") + Right(rut, 1)
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT mescontable,añocontable,neto,iva,total "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta" & empresa & ".facturasdecompras "
        If tipo = "8" Then
        csql2.sql = csql2.sql + "Where tipo ='" + tipo + "' and numero = '" & Format(numero, "0000000000") & "'"
        Else
        
        csql2.sql = csql2.sql + "Where tipo ='" + tipo + "' and numero = '" & Format(numero, "0000000000") & "' and rut = '" + rut + "' "
        
        End If
        
        
        lc_neto = 0
        lc_iva = 0
        lc_total = 0
            
        csql2.Execute
        esta_en_libro_compras = False
        Set resultados2 = csql2.OpenResultset
        If csql2.RowsAffected > 0 Then
            lc_mescontable = resultados2(0)
            lc_anocontable = resultados2(1)
            lc_neto = resultados2(2)
            lc_iva = resultados2(3)
            lc_total = resultados2(4)
            
            esta_en_libro_compras = True
        Else
            esta_en_libro_compras = False
        End If
        
        
Exit Function
salida:
MsgBox "archivo sii para este mes no esta disponible "


End Function

Public Function fecha_pendientes() As String
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT IFNULL(MAX(fecharecepcion),'" + Format(Date, "dd-mm-yyyy") + "') FROM " + clientesistema + "conta" + empresaactiva + ".sii_lp_99"
        csql2.Execute
        fecha_pendientes = ""
            Set resultados2 = csql2.OpenResultset
        If csql2.RowsAffected > 0 Then
            fecha_pendientes = resultados2(0)
        End If


End Function


Public Function fecha_aceptados(MES, año) As String
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        On Error GoTo PASO:
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT fecharecepcion FROM " + clientesistema + "conta" + empresaactiva + ".sii_lc_" + MES + "_" + año + " order by nro desc"
        csql2.Execute
        fecha_aceptados = ""
            Set resultados2 = csql2.OpenResultset
        If csql2.RowsAffected > 0 Then
            fecha_aceptados = resultados2(0)
        End If

PASO:
fecha_aceptados = "2017-01-01"
End Function

Public Function ESTAENSII(tipo, numero, rut, monto, empresa, mesrevisa, añorevisa) As Boolean
Dim multi As Integer
Dim empresafae As String
Dim rut2 As String
rut2 = rut

If tipo = 4 Then
 tipo = "33"
End If
If tipo = 5 Then
 tipo = "56"
End If
If tipo = 6 Then
 tipo = "61"
End If
If tipo = 0 Then
 tipo = "34"
End If
If tipo = 1 Then
 tipo = "30"
End If
If tipo = 9 Then
 tipo = "32"
End If

If tipo = 7 Then
 tipo = "46"
End If
If tipo = 3 Then
 tipo = "60"
End If
If tipo = 8 Then
 tipo = "914"
Rem  rut = "0555555555"
End If

On Error GoTo salida:
'numero = Val(numero)
'empresafae = CONFI_EMPRESAFAE
rut2 = Format(Mid(rut, 2, 8), "########") + "-" + Mid(rut, 10, 1)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT rutproveedor "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta" & empresa & ".sii_lc_" + mesrevisa + "_" + añorevisa + " "
        csql2.sql = csql2.sql + "Where tipodoc ='" + tipo + "' and folio = '" & CDbl(numero) & "' and rutproveedor = '" + rut2 + "' "
        
        
        csql2.Execute
        ESTAENSII = False
        
        
         estanoaceptado = False
        If csql2.RowsAffected > 0 Then
            ESTAENSII = True
        
        End If
        
        If ESTAENSII = False Then
        csql2.sql = "SELECT rutproveedor "
        csql2.sql = csql2.sql + "FROM " & clientesistema & "conta" & empresa & ".sii_lp_99 "
        csql2.sql = csql2.sql + "Where tipodoc ='" + tipo + "' and folio = '" & CDbl(numero) & "' and rutproveedor = '" + rut2 + "' "
        csql2.Execute
        estanoaceptado = False
        
        If csql2.RowsAffected > 0 Then
            ESTAENSII = True
            estanoaceptado = True
          
        End If
        
        
        End If
        
        
Exit Function
salida:
Rem MsgBox "archivo sii para este mes no esta disponible "


End Function


Public Function ESTAENSII_todo(tipo, numero, rut, fecha, empresa) As Boolean
Dim mesdiferencia As Double
Dim mesnuevo As String
Dim añonuevo As String

mesdiferencia = DateDiff("M", fecha, Date)
mesnuevo = Format(fecha, "mm")
añonuevo = Format(fecha, "yyyy")
ESTAENSII_todo = False

If ESTAENSII(tipo, numero, rut, "0", empresaactiva, mesnuevo, añonuevo) = True Then
ESTAENSII_todo = True
Exit Function
End If

For H = 1 To mesdiferencia
fecha = DateAdd("m", H, fecha)
mesnuevo = Format(fecha, "mm")
añonuevo = Format(fecha, "yyyy")

If ESTAENSII(tipo, numero, rut, "0", empresaactiva, mesnuevo, añonuevo) = True Then
ESTAENSII_todo = True
Exit Function
End If


Next H

End Function


Public Function lee_factura_de_compra(tipo, numero, rut) As Boolean
    Dim csql As New rdoQuery
    Dim CUENTA2 As String
   Rem  On Error GoTo no:
  
If tipo = "33" Then tipo = "4"
If tipo = "56" Then tipo = "5"
If tipo = "61" Then tipo = "6"
If tipo = "34" Then tipo = "0"
If tipo = "30" Then tipo = "1"
If tipo = "60" Then tipo = "3"
If tipo = "46" Then tipo = "7"
If tipo = "914" Then tipo = "8"


    
    Set csql.ActiveConnection = contadb
    csql.sql = "select numero from " & clientesistema & "conta" & empresaactiva & ".facturasdecompras "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
    csql.Execute
    lee_factura_de_compra = False
    If csql.RowsAffected > 0 Then
    lee_factura_de_compra = True
    End If
    Exit Function
no:
    desconectado = True
End Function



Sub crearcuentacorriente(rutprove, nombreprove, direccionprove, comunaprove, ciudadprove, giroprove, fonoprove)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        If empresaactiva <> "" Then
            Set csql.ActiveConnection = contadb
            nombreprove = Replace(nombreprove, "'", "")

            csql.sql = "INSERT IGNORE INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono,email_dte) "
            csql.sql = csql.sql & "values ('" & Format(fechasistema, "yyyy") & "','23100026','" & rutprove & "',"
            csql.sql = csql.sql & "'" & Replace(nombreprove, "'", "") & "','" & direccionprove & "','" & comunaprove & "',"
            csql.sql = csql.sql & "'" & ciudadprove & "','" & giroprove & "','" & fonoprove & "','AUTOMATICO')"
            csql.Execute
            
            csql.sql = "INSERT IGNORE INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono,email_dte) "
            csql.sql = csql.sql & "values ('" & Format(fechasistema, "yyyy") & "','11200044','" & rutprove & "',"
            csql.sql = csql.sql & "'" & nombreprove & "','" & direccionprove & "','" & comunaprove & "',"
            csql.sql = csql.sql & "'" & ciudadprove & "','" & giroprove & "','" & fonoprove & "','AUTOMATICO')"
            csql.Execute
            
            
             csql.sql = "INSERT ignore INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "values ('" & Format(fechasistema, "yyyy") & "','23100026','" & rutprove & "')"
            csql.Execute
            
            
            csql.sql = "INSERT ignore INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "values ('" & Format(fechasistema, "yyyy") & "','11200044','" & rutprove & "')"
            csql.Execute
            
            
            
            csql.Close
            Set csql = Nothing
      
        End If
          
            


End Sub



Sub grabafactura(tipo, numero, fecha, fechavencimiento, rut, NETO, iva, EXENTO, retencion, total)
    Dim netos As Double
    Dim DH As String
    Dim DH2 As String
    Dim mesconta As String
    Dim añoconta As String
    Dim diaconta As String
    Dim CUENTA2 As String
    
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    Dim ELECTRONICA As String
    Dim tipodoc As String

    Dim fechacom As String
    Rem On Error GoTo no:
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "retencion"
    campos(9, 0) = "total"
    campos(10, 0) = "añocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = "electronica"
    campos(14, 0) = "activo"
    campos(15, 0) = "fechadigitacion"
    campos(16, 0) = "folio"
    campos(17, 0) = "impuestoespecifico"
    campos(18, 0) = "cuentatipo"
    campos(19, 0) = ""
    If tipo = "34" And EXENTO = 0 Then
    EXENTO = total
    End If
    
    
    If tipo = "33" Then tipo = "4"
    If tipo = "56" Then tipo = "5"
    If tipo = "61" Then tipo = "6"
    If tipo = "34" Then tipo = "0"
    If tipo = "46" Then tipo = "7"
    
    campos(0, 1) = tipo
    campos(1, 1) = Format(numero, "0000000000")
    campos(2, 1) = Format(fecha, "yyyy-mm-dd")
    campos(3, 1) = Format(fechavencimiento, "yyyy-mm-dd")
  Rem   cuenta2 = Right(rut, 1)
    Rem rut = Format(Mid(rut, 1, Len(rut) - 2), "000000000")
    
    campos(4, 1) = rut
    campos(5, 1) = NETO
    campos(6, 1) = iva
    campos(7, 1) = EXENTO
    campos(8, 1) = retencion
    campos(9, 1) = total
    
    MESCONTABLE = MES
    AÑOCONTABLE = año
   
    campos(10, 1) = AÑOCONTABLE
    campos(11, 1) = Format(MESCONTABLE, "00")
    campos(12, 1) = "RECEPCION VIA SII"
        
    campos(13, 1) = "S"
    campos(14, 1) = "N"
    campos(15, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(16, 1) = LEERULTIMOFOLIO(campos(11, 1), campos(10, 1))
 
    campos(17, 1) = "0"
    If tipo = "3" Or tipo = "6" Then
   CUENTA2 = "11200044"
    Else
    CUENTA2 = "23100026"

    End If
    campos(18, 1) = CUENTA2
    condicion = ""
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb

    Call sqlconta.sqlconta(op, condicion)
'    k = sqlconta.Status
    fecha = Format(campos(3, 1), "yyyy-mm-dd")
'
'
    fechacom = año + "-" + MES + "-" + "01"
    If fecha >= fechacom Then
    fechacom = fecha
    End If
'
ivacredito = "11400001"
'
'
    If tipo = "4" Then tipodoc = "FC": DH = "H": DH2 = "D"
    If tipo = "0" Then tipodoc = "EE": DH = "H": DH2 = "D"
    If tipo = "5" Then tipodoc = "DC": DH = "H": DH2 = "D"
    If tipo = "6" Then tipodoc = "NC": DH = "D": DH2 = "H"
    If tipo = "7" Then tipodoc = "FP": DH = "H": DH2 = "D"
     
   Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "001", fechacom, CUENTA2, "", campos(4, 1), "", "CENTRALIZA DOCUMENTO DE COMPRAS " + numero, tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(9, 1), DH, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
   Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "002", fechacom, ivacredito, "", campos(4, 1), "", "CENTRALIZACION I.V.A", tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(6, 1), DH2, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
   If tipo = "7" Then
        Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "003", fechacom, "23200007", "", campos(4, 1), "", "RETENCION I.V.A " & tipodoc, tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), retencion, DH, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
   End If
CRCC = "0101"

    Call grabardetallefactura(campos(0, 1), campos(1, 1), campos(4, 1), campos(2, 1), campos(11, 1), campos(10, 1), tipodoc, CRCC, DH2, fechacom)
    
    Exit Sub
no:
    desconectado = True
End Sub

Sub grabardetallefactura(tipo, numero, rut, fecha, MES, año, tipodoc, CRCC, DH2, fechacom)
    Dim campos(20, 10) As String
    
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim ilas As Double

    Dim cuenta As String
    Dim DH As String
    Dim NOMBRE As String
  Rem   On Error GoTo no:
    año = Format(fecha, "yyyy")
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = "fechacreacion"
    campos(11, 0) = ""
    
    
    
    cuenta = leercuentaproveedores(rut)


Rem CALCULA NETOS
 If tipo = "7" Then
    lin = 4
    Else
    lin = 3
End If
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = cuenta
    campos(5, 1) = leernombrecuenta(cuenta, "n")
    If tipo = "0" Then
    campos(6, 1) = lc_neto + lc_exento
    Else
    campos(6, 1) = lc_neto
    End If
    
    campos(7, 1) = DH2
    campos(8, 1) = CRCC
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
  
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fechacom, campos(4, 1), "", campos(3, 1), "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    
  
    
     empresadte (empresaactiva)
Rem CALCULA ILAS CERVEZAS
ilas = 0
For k = 1 To 7
If lc_tipo(k) = "26" Then
ilas = lc_monto(k)
End If
Next k




    If ilas <> 0 Then
    lin = lin + 1
    
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal2(confi_localempresa, "cuentailacervezas")
    campos(5, 1) = "IMPUESTO ILA CERVEZAS"
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fechacom, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
   
Rem CALCULA HARINA
ilas = 0
For k = 1 To 7
If lc_tipo(k) = "19" Then
ilas = lc_monto(k)
End If
Next k



    
    If ilas <> 0 Then
    lin = lin + 1
    
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal2(confi_localempresa, "cuentaharina")
    campos(5, 1) = "IMPUESTO HARINA"
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fechacom, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
Rem LICORES
ilas = 0
For k = 1 To 7
If lc_tipo(k) = "24" Then
ilas = lc_monto(k)
End If
Next k



    
    If ilas <> 0 Then
    lin = lin + 1
    
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal2(confi_localempresa, "cuentailalicores")
    campos(5, 1) = "IMPUESTO LICORES"
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fechacom, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
Rem vinos
ilas = 0
For k = 1 To 7
If lc_tipo(k) = "25" Then
ilas = lc_monto(k)
End If
Next k



    
    If ilas <> 0 Then
    lin = lin + 1
    
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal2(confi_localempresa, "cuentailavinos")
    campos(5, 1) = "IMPUESTO VINOS"
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fechacom, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem refrescos azucar
ilas = 0
For k = 1 To 7
If lc_tipo(k) = "271" Then
ilas = ilas + lc_monto(k)
End If
Next k



   
    If ilas <> 0 Then
    lin = lin + 1
    
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal2(confi_localempresa, "cuentailarefrescos")
    campos(5, 1) = "IMPUESTO REFRESCOS AZUCAR"
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fechacom, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem refrescos sin azucar

ilas = 0

For k = 1 To 7
If lc_tipo(k) = "27" Then
ilas = ilas + lc_monto(k)
End If
Next k



   
    If ilas <> 0 Then
    lin = lin + 1
    
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    If Format(fecha, "yyyy-mm-dd") > "2014-09-30" Then
    campos(4, 1) = "11400017"
    Else
    campos(4, 1) = leerdatoslocal2(confi_localempresa, "cuentailarefrescos")
    End If
    
    campos(5, 1) = "IMPUESTO REFRESCOS NO AZUCAR"
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fechacom, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If


Rem CARNE
ilas = 0
For k = 1 To 7
If lc_tipo(k) = "18" Then
ilas = lc_monto(k)
End If
Next k



   
    If ilas <> 0 Then
    lin = lin + 1
    
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal2(confi_localempresa, "cuentacarne")
    campos(5, 1) = "IMPUESTO CARNE"
    campos(6, 1) = ilas
    campos(7, 1) = DH2
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fechacom, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
   
    Exit Sub
no:
    desconectado = True
    
End Sub


Public Function LEERULTIMOFOLIO(mesconta, añoconta) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select folio from " & clientesistema & "conta" + empresaactiva + ".facturasdecompras where mescontable = '" & Format(mesconta, "00") & "' AND añocontable = '" & añoconta & "' order by folio desc limit 0,1"
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
        If resultados(0) <> "NULO" Then
        LEERULTIMOFOLIO = resultados(0) + 1
        Else
        LEERULTIMOFOLIO = "0000000001"
        End If
        
    End If
    
End Function

Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    Rem On Error GoTo no:
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = MES
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = clientesistema & "conta" + empresaactiva + ".movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   Exit Sub
no:
   desconectado = True
   
End Sub



Public Function leerdatoslocal2(empresa, dato) As String
    Dim campos(10, 10) As String
empresadte (empresaactiva)
    campos(0, 0) = dato
    campos(1, 0) = ""
   
    campos(0, 2) = clientesistema & "gestion.g_maestroempresas"
    condicion = "codigo=" + "'" + confi_localempresa + "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerdatoslocal2 = sqlconta.response(0, 3)
    Else
        leerdatoslocal2 = ""
    End If
    
End Function
Public Function leercuentaproveedores(rut) As String
    Dim campos(10, 10) As String
    Rem On Error GoTo no:
    campos(0, 0) = "contable"
    campos(1, 0) = ""
   
    campos(0, 2) = clientesistema & "conta.proveedores_cuenta"
    condicion = "rut=" + "'" + rut + "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leercuentaproveedores = sqlconta.response(0, 3)
    Else
        leercuentaproveedores = "11350006"
    End If
    Exit Function
no:
    desconectado = True
End Function

Public Function leernombrecuenta(cuenta22, banco) As String


        Dim op As Integer
        
        
        campos(0, 0) = "nombre"
        campos(1, 0) = ""
        campos(0, 2) = clientesistema & "conta" + empresaactiva + ".cuentasdelmayor"
    
        condicion = "codigo = '" & cuenta22 & "' and año='" + Format(fechasistema, "yyyy") + "' "

        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
        
        leernombrecuenta = sqlconta.response(0, 3)
        Else
        leernombrecuenta = ""
        End If
End Function


