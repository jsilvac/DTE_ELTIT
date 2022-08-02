Attribute VB_Name = "contaelec"
Public lib_java_conta As String
Public lib_envio_libro_diario As String

Public Sub librerias_java_conta()
confi_rutaconta = "u:\fae_admin"
Dim LIBRERIAS(22) As String

lib_base = "java -classpath " & confi_rutaconta + ";"
LIBRERIAS(1) = confi_rutaconta + "\lib\apache-mime4j-0.6.jar;"
LIBRERIAS(2) = confi_rutaconta + "\lib\avalon-framework-4.2.0.jar;"
LIBRERIAS(3) = confi_rutaconta + "\lib\barcode4j-fop-ext-complete.jar;"
LIBRERIAS(4) = confi_rutaconta + "\lib\batik-all-1.7.jar;"
LIBRERIAS(5) = confi_rutaconta + "\lib\commons-io-1.4.jar;"
LIBRERIAS(6) = confi_rutaconta + "\lib\commons-logging-1.1.1.jar;"
LIBRERIAS(7) = confi_rutaconta + "\lib\fop.jar;"
LIBRERIAS(8) = confi_rutaconta + "\lib\httpclient-4.0.jar;"
LIBRERIAS(9) = confi_rutaconta + "\lib\httpcore-4.0.1.jar;"
LIBRERIAS(10) = confi_rutaconta + "\lib\httpmime-4.0.jar;"
LIBRERIAS(11) = confi_rutaconta + "\lib\jargs.jar;"
LIBRERIAS(12) = confi_rutaconta + "\lib\jsr173_1.0_api.jar;"
LIBRERIAS(13) = confi_rutaconta + "\lib\not-yet-commons-ssl-0.3.11.jar;"
LIBRERIAS(14) = confi_rutaconta + "\lib\serializer-2.7.0.jar;"
LIBRERIAS(15) = confi_rutaconta + "\lib\xalan-2.7.0.jar;"
LIBRERIAS(16) = confi_rutaconta + "\lib\xbean.jar;"
LIBRERIAS(17) = confi_rutaconta + "\lib\xml-apis-ext-1.3.04.jar;"
LIBRERIAS(18) = confi_rutaconta + "\lib\xmlgraphics-commons-1.3.jar;"
LIBRERIAS(19) = confi_rutaconta + "\lib\xmlsec-1.4.3.jar;"
LIBRERIAS(20) = confi_rutaconta + "\lib\OpenLibsDte.jar "



lib_java_conta = lib_base
For k = 1 To 20
lib_java_conta = lib_java_conta + LIBRERIAS(k)
Next k

lib_envio_libro_diario = lib_java_conta + "cl.adminerp.lce.envios.GeneraLceEnvioLibrosDiario "
End Sub

