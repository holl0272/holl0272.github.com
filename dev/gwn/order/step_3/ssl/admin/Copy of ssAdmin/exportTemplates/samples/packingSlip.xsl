<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template match='/'>
    <html>
      <head>
	  </head>
      <body>
	    <center>	
		<xsl:apply-templates select="orders/order" />
	    </center>
      </body>
    </html>

  </xsl:template>

  <xsl:template match="order" >
		<table width="100%" border="1" cellspacing="0" cellpadding="2" STYLE="page-break-before:always">
		<tr>
			<td valign="top"><p>Sandshot Software<br/>
				1958 Devils Reach Road<br/>
				Woodbridge, VA 22192<br/>
				TEL: (571) 723-1540<br/>
				sales@sandshot.net</p>
			</td>
			<td valign="top"><table width="100%" border="1" cellspacing="0" cellpadding="2">
				<tr>
				<td width='50%'>ORDER NUMBER: </td>
				<td><xsl:value-of select="orderID"/></td>
				</tr>
				<tr>
				<td>ORDER DATE:</td>
				<td><xsl:value-of select="orderDate"/></td>
				</tr>
			</table></td>
		</tr>
		<tr><td colspan="2">&amp;nbsp;</td></tr>
		<tr>
			<td valign="top"><table width="100%" border="1" cellspacing="0" cellpadding="0">
				<tr>
				<td>BILL TO:<br/>
					<xsl:apply-templates select="billingAddress" />
				</td>
				</tr>
			</table></td>
			<td valign="top">
			  <table width="100%" border="1" cellspacing="0" cellpadding="0">
				<tr>
				<td>SHIP TO:<br/>
					<xsl:apply-templates select="shippingAddress" />
				</td>
				</tr>
			</table></td>
		</tr>
		<tr align="center">
			<td colspan="2"><br/>
			<table width="75%" border="1" cellspacing="0" cellpadding="2">
				<tr>
				<td>METHOD OF PAYMENT</td>
				<td>P.O. NUMBER</td>
				<td>SHIP VIA</td>
				</tr>
				<tr>
				<td><xsl:value-of select="orderPaymentMethod"/></td>
				<td>&amp;nbsp;<xsl:value-of select="orderPurchaseOrderNumber"/></td>
				<td><xsl:value-of select="orderShipMethod"/></td>
				</tr>
			</table>
			<br/>
			</td>
		</tr>
		<tr>
			<td colspan="2">
			  <table width="100%" border="1" cellspacing="0" cellpadding="2">
				<tr>
				  <td width='60'>N0</td>
				  <td width='160'>ITEM#</td>
				  <td width='70%'>DESCRIPTION</td>
				  <td width='120'>ORDER QTY</td>
				  <td width='160'>PRICE</td>
				</tr>

				<xsl:for-each select="orderDetail">
					<tr>
						<xsl:if test="position() mod 2 = 0">
							<xsl:attribute name="bgcolor">yellow</xsl:attribute>
						</xsl:if>
						<td><xsl:number value="position()" format="1"/></td>
						<td><xsl:value-of select="odrdtProductID" /></td>
						<td><xsl:value-of select="odrdtProductName" /></td>
						<td><xsl:value-of select="odrdtQuantity" /></td>
						<td>$<xsl:value-of select="odrdtSubTotal" /></td>
					</tr>
					<xsl:for-each select="odrdtAttDetailID">
						<tr bgcolor="">
							<td colspan="2">&amp;nbsp;</td>
							<td colspan="3">&amp;nbsp;&amp;nbsp;<span class="orderItemText"><xsl:value-of select="odrattrName" />: <xsl:value-of select="odrattrAttribute" /></span></td>
						</tr>
					</xsl:for-each>
				</xsl:for-each>

				<tr>
				  <td>&amp;nbsp;</td>
				  <td>&amp;nbsp;</td>
				  <td>&amp;nbsp;</td>
				  <td>&amp;nbsp;</td>
				  <td>&amp;nbsp;</td>
				</tr>
				<tr>
				  <td colspan="2">&amp;nbsp;</td>
				  <td colspan="3" align="right">
				  <table border="1" cellspacing="0" cellpadding="2">
					<tr>
						<td>subtotal:</td>
						<td>&amp;nbsp;$<xsl:value-of select="orderAmount" /></td>
					</tr>
					
					<tr>
						<td><xsl:value-of select="orderShipMethod" />:</td>
						<td>&amp;nbsp;$<xsl:value-of select="orderShippingAmount" /></td>
					</tr>
					
					<xsl:if test="orderHandling[.&gt;0]">
					<tr>
					  <td>HANDLING:</td>
					  <td>&amp;nbsp;$<xsl:value-of select="orderHandling" /></td>
					</tr>
					</xsl:if>
					
					<xsl:if test="orderSTax[.&gt;0]">
					<tr>
					  <td><xsl:value-of select="shippingAddress/State" /> State Tax:</td>
					  <td>&amp;nbsp;$<xsl:value-of select="orderSTax" /></td>
					</tr>
					</xsl:if>
					
					<xsl:if test="orderCTax[.&gt;0]">
					<tr>
					  <td><xsl:value-of select="shippingAddress/Country" /> Tax:</td>
					  <td>&amp;nbsp;$<xsl:value-of select="orderCTax" /></td>
					</tr>
					</xsl:if>
					
					<tr>
					  <td>TOTAL:</td>
					  <td>&amp;nbsp;$<xsl:value-of select="orderGrandTotal" /></td>
					</tr>

					<xsl:if test="ssGiftCertificate/RedemptionAmount[.&lt;0]">
					<tr>
					  <td>Certificate (<xsl:value-of select="ssGiftCertificate/CertificateNumber" />)::</td>
					  <td>&amp;nbsp;$<xsl:value-of select="ssGiftCertificate/RedemptionAmount" /></td>
					</tr>
					<tr>
					  <td>Amount Billed::</td>
					  <td>&amp;nbsp;$<xsl:value-of select="ssGiftCertificate/ssGCNewTotalDue" /></td>
					</tr>
					</xsl:if>
					
					</table></td>
				</tr>
			</table></td>
		</tr>
		<tr>
			<td>&amp;nbsp;</td>
			<td>&amp;nbsp;</td>
		</tr>
		<xsl:if test="orderCTax[.&gt;0]">
		<tr>
		  <td colspan="2">
		    <table width="100%" border="1" cellspacing="0" cellpadding="2">
			<tr>
			  <td>
			  COMMENTS:<br/>
			  <xsl:value-of select="orderComments" />
			  </td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td><xsl:value-of select="shippingAddress/Country" /> Tax:</td>
			<td>&amp;nbsp;$<xsl:value-of select="orderCTax" /></td>
		</tr>
		</xsl:if>
		</table>
  </xsl:template>

  <xsl:template match="orderDetail" >
    <tr>
    <th class= "selector">4</th>
    <td><xsl:value-of select="@odrdtProductID"/></td>
    <td><xsl:value-of select="@odrdtProductName" /></td>
    <td><xsl:value-of select="@odrdtQuantity" /></td>
    <td><xsl:value-of select="@odrdtSubTotal" /></td>
	</tr>
  </xsl:template>

  <xsl:template match="shippingAddress" >
	<table>
    <tr>
      <td>
        <xsl:value-of select="FirstName"/>
		<xsl:choose>
			<xsl:when test="MiddleInitial[.&gt;'']">
				<xsl:value-of select="MiddleInitial" />
			</xsl:when>
		</xsl:choose>
		<xsl:value-of select="LastName" />
      </td>
    </tr>
	<xsl:choose>
		<xsl:when test="Company[.&gt;'']">
			<tr><td><xsl:value-of select="Company" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Addr1[.&gt;'']">
			<tr><td><xsl:value-of select="Addr1" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Addr2[.&gt;'']">
			<tr><td><xsl:value-of select="Addr2" /></td></tr>
		</xsl:when>
	</xsl:choose>
    <tr>
      <td>
        <xsl:value-of select="City"/>, 
        <xsl:value-of select="State"/>
        <xsl:value-of select="Zip"/>
      </td>
    </tr>
	<xsl:choose>
		<xsl:when test="Country[.='US']">
			<tr><td><xsl:value-of select="CountryName" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Phone[.&gt;'']">
			<tr><td>Phone: <xsl:value-of select="Phone" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Fax[.&gt;'']">
			<tr><td>Fax: <xsl:value-of select="Fax" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Email[.&gt;'']">
			<tr><td><xsl:value-of select="Email" /></td></tr>
		</xsl:when>
	</xsl:choose>
	</table>
  </xsl:template>

  <xsl:template match="billingAddress" >
	<table>
    <tr>
      <td>
        <xsl:value-of select="FirstName"/>
		<xsl:choose>
			<xsl:when test="MiddleInitial[.&gt;'']">
				<xsl:value-of select="MiddleInitial" />
			</xsl:when>
		</xsl:choose>
		<xsl:value-of select="LastName" />
      </td>
    </tr>
	<xsl:choose>
		<xsl:when test="Company[.&gt;'']">
			<tr><td><xsl:value-of select="Company" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Addr1[.&gt;'']">
			<tr><td><xsl:value-of select="Addr1" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Addr2[.&gt;'']">
			<tr><td><xsl:value-of select="Addr2" /></td></tr>
		</xsl:when>
	</xsl:choose>
    <tr>
      <td>
        <xsl:value-of select="City"/>, 
        <xsl:value-of select="State"/>
        <xsl:value-of select="Zip"/>
      </td>
    </tr>
	<xsl:choose>
		<xsl:when test="Country[.='US']">
			<tr><td><xsl:value-of select="CountryName" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Phone[.&gt;'']">
			<tr><td>Phone: <xsl:value-of select="Phone" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Fax[.&gt;'']">
			<tr><td>Fax: <xsl:value-of select="Fax" /></td></tr>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="Email[.&gt;'']">
			<tr><td><xsl:value-of select="Email" /></td></tr>
		</xsl:when>
	</xsl:choose>
	</table>
  </xsl:template>

</xsl:stylesheet>