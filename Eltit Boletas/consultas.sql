INSERT INTO eltit_manager.clientes_dte
            (prefijo,
             rut,
             codigo_contable,
             razon_social,
             fecha_activacion,
             ip_servidor,
             fecha_finalizacion,
             activo,
             servidor_destino,
             fecha_resolucion,
             numero_resolucion,
             rut_certificado,
             nombre_certificado,
             critico_33,
             critico_61,
             critico_39,
             cloud,
             mysql_user,
             mysql_pass,
             sube_dte_boletas,
             sube_dte_facturas,
             sube_ventas,
             inicio_sincroniza,
             numero_registros,
             mail_intercambio_smtp,
             mail_intercambio_direccion,
             mail_intercambio_clave)
             
SELECT       "eltit_",
             LPAD(REPLACE(emp.rut,"-",""),10,"0") AS rut,
             emp.codigoempresa ,
             emp.nombre AS  razon_social,
             '2020-09-21' AS fecha_activacion,
             '' AS ip_servidor,
             '000-00-00' AS fecha_finalizacion,
             '1' AS activo,
             '' AS servidor_destino,
             emp.fecharesolucion AS fecha_resolucion,
             emp.numeroresolucion AS  numero_resolucion,
             '7762388-4' AS rut_certificado,
             'juan patricio eltit jadue' AS nombre_certificado,
             '1' AS critico_33,
             '1' AS critico_61,
             '1' AS critico_39,
             '1' AS cloud,
             'root' AS mysql_user,
             '123' AS mysql_pass,
             '0' AS sube_dte_boletas,
             '0' AS sube_dte_facturas,
             '0' AS sube_ventas,
             '2020-09-14' inicio_sincroniza,
             50 AS numero_registros,
             emp.servermail AS mail_intercambio_smtp,
             emp.mailsalida AS mail_intercambio_direccion,
             emp.clavemail AS mail_intercambio_clave FROM eltit_conta.maestroempresas AS emp
	     WHERE emp.profesionrepresentante = '1' ;             
             
             SELECT codigoempresa, nombre, rut, profesionrepresentante FROM maestroempresas ORDER BY codigoempresa ;
             SELECT codigo, nombre,nombrelocal, direccion, emite_39 FROM  clientes_locales;
             
    SELECT dte.tipo AS tipo_doc, dte.numero AS foliosii,dte.fecha AS fecha_emision, dte.fechaenviosii AS fae_fechaenvio_sii, 
 '00:00:00' AS fae_horaenvio_sii, dte.cajadocumento AS caja_doc,  IFNULL(0,0) AS monto_exento, IFNULL(dte.monto,0) AS monto_total, 
 dte.xml AS fae_xml  FROM eltit_fae00.sv_dte00 AS dte  
 LEFT JOIN eltit_ventas00.sv_documento_cabeza_00 AS dc  ON(LPAD(dte.numero,10,'0') = dc.foliosii)  
 AND dc.local = dte.localdocumento AND dc.fecha  = dte.fechadocumento AND dte.cajadocumento = dc.caja  
 AND dte.tipodocumento = dc.tipo  
 WHERE dte.localdocumento = '00' AND dte.fecha = '2020-08-31' AND dte.tipo ='39'
 UNION
  SELECT dte.tipo AS tipo_doc, dte.numero AS foliosii,dte.fecha AS fecha_emision, dte.fechaenviosii AS fae_fechaenvio_sii, 
 '00:00:00' AS fae_horaenvio_sii, dte.cajadocumento AS caja_doc,  IFNULL(0,0) AS monto_exento, IFNULL(dte.monto,0) AS monto_total, 
 dte.xml AS fae_xml  FROM eltit_fae05.sv_dte05 AS dte  
 LEFT JOIN eltit_ventas05.sv_documento_cabeza_05 AS dc  ON(LPAD(dte.numero,10,'0') = dc.foliosii)  
 AND dc.local = dte.localdocumento AND dc.fecha  = dte.fechadocumento AND dte.cajadocumento = dc.caja  
 AND dte.tipodocumento = dc.tipo 
 WHERE dte.localdocumento = '05' AND dte.fecha = '2020-08-31' AND dte.tipo ='39'
  UNION
  SELECT dte.tipo AS tipo_doc, dte.numero AS foliosii,dte.fecha AS fecha_emision, dte.fechaenviosii AS fae_fechaenvio_sii, 
 '00:00:00' AS fae_horaenvio_sii, dte.cajadocumento AS caja_doc,  IFNULL(0,0) AS monto_exento, IFNULL(dte.monto,0) AS monto_total, 
 dte.xml AS fae_xml  FROM eltit_fae25.sv_dte25 AS dte  
 LEFT JOIN eltit_ventas25.sv_documento_cabeza_25 AS dc  ON(LPAD(dte.numero,10,'0') = dc.foliosii)  
 AND dc.local = dte.localdocumento AND dc.fecha  = dte.fechadocumento AND dte.cajadocumento = dc.caja  
 AND dte.tipodocumento = dc.tipo 
 WHERE dte.localdocumento = '25' AND dte.fecha = '2020-08-31' AND dte.tipo ='39'
  UNION
  SELECT dte.tipo AS tipo_doc, dte.numero AS foliosii,dte.fecha AS fecha_emision, dte.fechaenviosii AS fae_fechaenvio_sii, 
 '00:00:00' AS fae_horaenvio_sii, dte.cajadocumento AS caja_doc,  IFNULL(0,0) AS monto_exento, IFNULL(dte.monto,0) AS monto_total, 
 dte.xml AS fae_xml  FROM eltit_fae41.sv_dte41 AS dte  
 LEFT JOIN eltit_ventas41.sv_documento_cabeza_41 AS dc  ON(LPAD(dte.numero,10,'0') = dc.foliosii)  
 AND dc.local = dte.localdocumento AND dc.fecha  = dte.fechadocumento AND dte.cajadocumento = dc.caja  
 AND dte.tipodocumento = dc.tipo 
 WHERE dte.localdocumento = '41' AND dte.fecha = '2020-08-31' AND dte.tipo ='39'
  UNION
  SELECT dte.tipo AS tipo_doc, dte.numero AS foliosii,dte.fecha AS fecha_emision, dte.fechaenviosii AS fae_fechaenvio_sii, 
 '00:00:00' AS fae_horaenvio_sii, dte.cajadocumento AS caja_doc,  IFNULL(0,0) AS monto_exento, IFNULL(dte.monto,0) AS monto_total, 
 dte.xml AS fae_xml  FROM eltit_fae60.sv_dte60 AS dte  
 LEFT JOIN eltit_ventas60.sv_documento_cabeza_60 AS dc  ON(LPAD(dte.numero,10,'0') = dc.foliosii)  
 AND dc.local = dte.localdocumento AND dc.fecha  = dte.fechadocumento AND dte.cajadocumento = dc.caja  
 AND dte.tipodocumento = dc.tipo 
 WHERE dte.localdocumento = '60' AND dte.fecha = '2020-08-31' AND dte.tipo ='39'
 ORDER BY foliosii;          
             
             
             
             
             