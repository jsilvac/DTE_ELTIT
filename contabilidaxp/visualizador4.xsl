<?xml version="1.0" encoding="iso-8859-1"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<xsl:template match="/">
<html>
<head>
<title>Visualizador de Facturas Electronicas</title>
<meta http-equiv="Content-Type" content="text/html" charset="windows-1252"/>
</head>
<body>

<xsl:for-each select="xml-fragment/siid:Documento">

<div align="center">
<table border="0" width="670" bgcolor="#EEEEFF">
  <tr>
    <td width="100%">
      <table border="0" width="100%" cellspacing="0">
        <tr>
         <td width="49%" valign="top">
         
	   <table border="0" width="100%">
              <tr>
                <td width="21%"></td>
                <td width="79%" align="center"><b><font size="5" color="#000080"><xsl:value-of select="siid:Encabezado/siid:Emisor/siid:RznSoc"/></font></b></td>
              </tr>              <tr>
                <td width="21%"></td>
                <td width="79%" align="center"><font size="2" color="#000080">Giro: <xsl:value-of select="siid:Encabezado/siid:Emisor/siid:GiroEmis"/> - Sucursal: <xsl:value-of select="siid:Encabezado/siid:Emisor/siid:Sucursal"/>,<br/> <xsl:value-of select="siid:Encabezado/siid:Emisor/siid:DirOrigen"/>, <xsl:value-of select="siid:Encabezado/siid:Emisor/siid:CmnaOrigen"/></font></td>
              </tr>
              <tr>
                <td width="21%"></td>
                <td width="79%" align="center"><font size="2" color="#000080"><b><xsl:value-of select="siid:Encabezado/siid:Emisor/siid:CiudadOrigen"/></b></font></td>
              </tr>
            </table>
          </td>
            
          <td width="41%" valign="top">
            <table border="0" width="100%" cellspacing="0">
              <tr>
                <td width="100%">
                  <table border="2" width="100%" bordercolor="#FF0000" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="100%">
                        <p align="center"/><b><font size="4">RUT.: <xsl:value-of select="siid:Encabezado/siid:Emisor/siid:RUTEmisor"/> <br/>
  			<xsl:choose>
			      <xsl:when test="siid:Encabezado/siid:IdDoc/siid:TipoDTE[. = 33]">FACTURA ELECTRONICA</xsl:when>
			      <xsl:when test="siid:Encabezado/siid:IdDoc/siid:TipoDTE[. = 34]">FACTURA EXENTA ELECTRONICA</xsl:when>
			      <xsl:when test="siid:Encabezado/siid:IdDoc/siid:TipoDTE[. = 52]">GUIA DE DESPACHO ELECTRONICA</xsl:when>
			      <xsl:when test="siid:Encabezado/siid:IdDoc/siid:TipoDTE[. = 61]">NOTA DE CREDITO ELECTRONICA</xsl:when>
				  <xsl:when test="siid:Encabezado/siid:IdDoc/siid:TipoDTE[. = 56]">NOTA DE DEBITO ELECTRONICA</xsl:when>
				  <xsl:when test="siid:Encabezado/siid:IdDoc/siid:TipoDTE[. = 39]">BOLETA ELECTRONICA</xsl:when>
			      <xsl:otherwise>OTRO DOCUMENTO</xsl:otherwise>
  		        </xsl:choose>
  			<br/>
                        N° <xsl:value-of select="siid:Encabezado/siid:IdDoc/siid:Folio"/></font></b></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td width="100%">
                  <p align="center"/></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%">
      <table border="0" width="100%" cellpadding="0" height="137">
        <tr>
          <td width="14%" height="21"></td>
          <td width="43%" height="21"></td>
          <td width="20%" height="21"><font size="2" color="#000080">Fecha Emisión:</font></td>
          <td width="23%" height="21"><xsl:value-of select="siid:Encabezado/siid:IdDoc/siid:FchEmis"/></td>
        </tr>
        <tr>
          <td width="14%" height="21"><font size="2" color="#000080">Señor(es):</font></td>
          <td width="43%" height="21"><xsl:value-of select="siid:Encabezado/siid:Receptor/siid:RznSocRecep"/></td>
          <td width="20%" height="21"><font size="2" color="#000080">RUT.:</font></td>
          <td width="23%" height="21"><xsl:value-of select="siid:Encabezado/siid:Receptor/siid:RUTRecep"/></td>
        </tr>
        <tr>
          <td width="14%" height="20"><font size="2" color="#000080">Dirección:</font></td>
          <td width="43%" height="20"><xsl:value-of select="siid:Encabezado/siid:Receptor/siid:DirRecep"/></td>
          <td width="20%" height="20"><font size="2" color="#000080">Giro:</font></td>
          <td width="23%" height="20"><xsl:value-of select="siid:Encabezado/siid:Receptor/siid:GiroRecep"/></td>
        </tr>
        <tr>
          <td width="14%" height="21"><font size="2" color="#000080">Comuna:</font></td>
          <td width="43%" height="21"><xsl:value-of select="siid:Encabezado/siid:Receptor/siid:CmnaRecep"/></td>
          <td width="20%" height="21"><font size="2" color="#000080">Fecha Venc:</font></td>
          <td width="23%" height="21"><xsl:value-of select="siid:Encabezado/siid:IdDoc/siid:FchVenc"/></td>
        </tr>
        <tr>
          <td width="14%" height="21"><font size="2" color="#000080">Ciudad:</font></td>
          <td width="43%" height="21"><xsl:value-of select="siid:Encabezado/siid:Receptor/siid:CiudadRecep"/></td>
          <td width="20%" height="21"><font size="2" color="#000080">Forma Pago:</font></td>
          <td width="23%" height="21"><xsl:value-of select="siid:Encabezado/siid:IdDoc/siid:FmaPago"/>
		<xsl:choose>
		      <xsl:when test="Encabezado/IdDoc/FmaPago[. = 1]">Efectivo</xsl:when>
		      <xsl:when test="Encabezado/IdDoc/FmaPago[. = 2]">Crédito</xsl:when>
  		</xsl:choose>          
          </td>
        </tr>
      </table>
    </td>
  </tr>
    <tr>
   <td>
    <hr/>
   </td>
  </tr>
  <tr>
    <td width="100%">

<!--selecciona el tipo de detalle a mostrar con/sin descuentos -->
    
<xsl:choose>        
   <xsl:when test="siid:Detalle/siid:DescuentoMonto">
      <table border="0" width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td width="3%"><font color="#000080" size="2">N°</font></td>
          <td width="12%"><font color="#000080" size="2">Codigo</font></td>
          <td width="8%"><font color="#000080" size="2">Cantidad</font></td>
          <td width="41%"><font color="#000080" size="2">Detalle</font></td>
          <td width="12%" align="right"><font color="#000080" size="2">Valor Unitario</font></td>
          <td width="12%" align="right"><font color="#000080" size="2">Dscto</font></td>
          <td width="12%" align="right"><font color="#000080" size="2">Total</font></td>
        </tr>
                   <tr>
		   <td colspan="7">
		    <hr/>
		   </td>
		  </tr>
     	<xsl:for-each select="siid:Detalle">
     
	        <tr>
	          <td width="3%"><font size="2"><xsl:value-of select="siid:NroLinDet"/></font></td>
	          <td width="12%"><font size="2"><xsl:value-of select="siid:CdgItem/VlrCodigo"/></font></td>
	          <td width="8%"><font size="2"><xsl:value-of select="siid:QtyItem"/></font></td>
	          <td width="41%"><font size="2"><xsl:value-of select="siid:NmbItem"/>
	          	<xsl:choose>        
   				<xsl:when test="siid:DscItem">
   					<br/><xsl:value-of select="siid:DscItem"/>
   				</xsl:when>
   			</xsl:choose></font>
	          
	          </td>
	          <td width="12%" align="right"><font size="2"><xsl:value-of select="siid:PrcItem"/></font></td>
	          <td width="12%" align="right"><font size="2"><xsl:value-of select="siid:DescuentoMonto"/></font></td>
	          <td width="12%" align="right"><font size="2"><xsl:value-of select="siid:MontoItem"/></font></td>
	        </tr>
      
     	</xsl:for-each>
     	 </table>
     </xsl:when>
     <xsl:otherwise> 
            
            <table border="0" width="100%" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="3%"><font color="#000080" size="2">N°</font></td>
	          <td width="12"><font color="#000080" size="2">Codigo</font></td>
	          <td width="8%"><font color="#000080" size="2">Cantidad</font></td>
	          <td width="45%"><font color="#000080" size="2">Detalle</font></td>
	          <td width="16%" align="right"><font color="#000080" size="2">Valor Unitario</font></td>
	          <td width="16%" align="right"><font color="#000080" size="2">Total</font></td>
	        </tr>
                  <tr>
		   <td colspan="6">
		    <hr/>
		   </td>
		  </tr>
     		<xsl:for-each select="siid:Detalle">
     
	        <tr>
	          
	          <td width="3%"><font size="2"><xsl:value-of select="siid:NroLinDet"/></font></td>
	          <td width="12%"><font size="2"><xsl:value-of select="siid:CdgItem/siid:VlrCodigo"/></font></td>
	          <td width="8%"><font size="2"><xsl:value-of select="siid:QtyItem"/></font></td>
	          <td width="45%"><font size="2"><xsl:value-of select="siid:NmbItem"/>
	          	<xsl:choose>        
   				<xsl:when test="siid:DscItem">
   					<br/><xsl:value-of select="siid:DscItem"/>
   				</xsl:when>
   			</xsl:choose>	          
	          
	          </font></td>
	          <td width="16%" align="right"><font size="2"><xsl:value-of select="siid:PrcItem"/></font></td>
	          <td width="16%" align="right"><font size="2"><xsl:value-of select="siid:MontoItem"/></font></td>
	          
	        </tr>
  
     		</xsl:for-each>
        </table>
     </xsl:otherwise>
</xsl:choose>      
    </td>
  </tr>

<!--********descuentos y recargos globales*********-->

  <tr>
   <td width="100%">
       <table border="0" width="100%" cellpadding="0" cellspacing="0">
	<xsl:for-each select="DscRcgGlobal">
	        <tr>
	          <td width="25%"><font size="2"><xsl:value-of select="siid:TpoMov"/></font></td>
	          <td width="50%"><font size="2"><xsl:value-of select="siid:GlosaDR"/></font></td>
	          <td width="10%" align="right"><font size="2"><xsl:value-of select="siid:TpoValor"/></font></td>
	          <td width="15%" align="right"><font size="2"><xsl:value-of select="siid:ValorDR"/></font></td>
	        </tr>
  
  	</xsl:for-each>
       </table>

   </td>
  </tr>  
  
  <tr>
   <td>
    <hr/>
   </td>
  </tr>
  
<!--  REFERENCIAS DEL DTE-->

	<xsl:choose>        
	   <xsl:when test="siid:Referencia/siid:TpoDocRef[. &lt; 104]">
	   <tr>
   	    <td>
	      <table border="0" width="100%" cellpadding="0" cellspacing="0">
	        <tr><td colspan="6" align="center">REFERENCIAS</td></tr>
	        <tr>
	          <td width="3%"><font color="#000080" size="2">Lin</font></td>
	          <td width="10%"><font color="#000080" size="2">Doc</font></td>
	          <td width="7%"><font color="#000080" size="2">Folio</font></td>
	          <td width="10%"><font color="#000080" size="2">Fecha</font></td>
	          <td width="50%"><font color="#000080" size="2">Motivo</font></td>
	          <td width="20%"><font color="#000080" size="2">Corr.Factura</font></td>
	          
	        </tr>
	                   <tr>
			   <td colspan="6">
			    <hr/>
			   </td>
			  </tr>
	     	<xsl:for-each select="siid:Referencia">
	     
		        <tr>
		          <td width="3%"><font size="2"><xsl:value-of select="siid:NroLinRef"/></font></td>
		          <td width="10%"><font size="2"><xsl:value-of select="siid:TpoDocRef"/></font></td>
		          <td width="7%"><font size="2"><xsl:value-of select="siid:FolioRef"/></font></td>
		          <td width="10%"><font size="2"><xsl:value-of select="siid:FchRef"/></font></td>
		          <td width="50%"><font size="2"><xsl:value-of select="siid:RazonRef"/></font></td>
		          <td width="20%"><font size="2"><xsl:value-of select="siid:CorrFact"/></font></td>
		        </tr>
	      
	     	</xsl:for-each>
	     	 </table>
	     </td>
   	    </tr>
   	      <tr>
		 <td>
		   <hr/>
		 </td>
	      </tr>
	     </xsl:when>  
  	</xsl:choose>  
  

<!-- FIN REFERENCIAS DEL DTE-->  
  
  
  
  <tr>
    <td width="100%">
      <table border="0" width="100%" cellspacing="0">
        <tr>
          <td width="33%" align="center"></td>
          <td width="24%"></td>
          <td width="43%" align="right" valign="top">
	          <table>
		          <tr>
		          	<td width="50%"><font size="3" color="#000080"><b>Neto:</b></font></td>
		          	<td width="50%" align="right"><xsl:value-of select="siid:Encabezado/siid:Totales/siid:MntNeto"/></td>
		          </tr>
		          <tr>
		          	<td width="50%"><font size="3" color="#000080"><b>IVA <xsl:value-of select="siid:Encabezado/siid:Totales/siid:TasaIVA"/> %:</b></font></td>
		          	<td width="50%" align="right"><xsl:value-of select="siid:Encabezado/siid:Totales/siid:IVA"/></td>
		          </tr>
		        <xsl:choose> 
 			<xsl:when test="siid:Encabezado/siid:Totales/siid:MntExe">
 				<tr>
			          <td width="50%">Monto Exento</td>
			          <td width="50%" align="right"><xsl:value-of select="siid:Encabezado/siid:Totales/siid:MntExe"/></td>
			        </tr>	
 			</xsl:when>
 			</xsl:choose>

			<!--IMPUESTOS ADICIONALES -->
			    
			<xsl:for-each select="siid:Encabezado/siid:Totales/siid:ImptoReten/siid:TipoImp">
			        <tr>
			          <td width="50%">
		  			<xsl:choose>
					      <xsl:when test=".[. = 18]"><font size="3" color="#000080"><b>IVA Ant.(Carne)</b></font></xsl:when>
					      <xsl:when test=".[. = 19]"><font size="3" color="#000080"><b>IVA Ant.(Harina)</b></font></xsl:when>
					      <xsl:when test=".[. = 23]"><font size="3" color="#000080"><b>Impto.Adic 15%</b></font></xsl:when>
					      <xsl:when test=".[. = 24]"><font size="3" color="#000080"><b>Licores 27%</b></font></xsl:when>
					      <xsl:when test=".[. = 25]"><font size="3" color="#000080"><b>Vinos 15%</b></font></xsl:when>
					      <xsl:when test=".[. = 26]"><font size="3" color="#000080"><b>Cervezas 15%</b></font></xsl:when>
					      <xsl:when test=".[. = 27]"><font size="3" color="#000080"><b>Bebidas 13%</b></font></xsl:when>
					      <xsl:otherwise>IMPTO. xxx</xsl:otherwise>
		  		        </xsl:choose>
			          
			          </td>
			          <td width="50%" align="right"><xsl:value-of select="../siid:MontoImp"/></td>
			        </tr>
			</xsl:for-each>
			        <tr>
			          <td width="50%"><font size="3" color="#000080"><b>Total:</b></font></td>
			          <td width="50%" align="right"><xsl:value-of select="siid:Encabezado/siid:Totales/siid:MntTotal"/></td>
			        </tr>	

<!--Campos Opcionales-->
			<xsl:choose> 
 			<xsl:when test="siid:Encabezado/siid:Totales/siid:SaldoAnterior">
 				<tr>
			          <td width="50%">Saldo Anterior</td>
			          <td width="50%" align="right"><xsl:value-of select="siid:Encabezado/siid:Totales/siid:SaldoAnterior"/></td>
			        </tr>	
       				<tr>
			          <td width="50%">Total a Pagar</td>
			          <td width="50%" align="right"><xsl:value-of select="siid:Encabezado/siid:Totales/siid:VlrPagar"/></td>
			        </tr>	
 			</xsl:when>
 			</xsl:choose>

	          </table>
          </td>
       </tr>
      </table>
    </td>
  </tr>
</table>
<br/>
<br/>
<br/>
<br/>

</div>

</xsl:for-each>

</body>
</html>
</xsl:template>


</xsl:stylesheet>