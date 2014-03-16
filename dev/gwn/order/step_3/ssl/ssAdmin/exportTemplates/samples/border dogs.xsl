<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template match='/'>
    <html>
      <head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title>Sales Receipt</title>
		<link rel="stylesheet" href="exportTemplates/borderDogs.css" type="text/css" />
	  </head>
      <body>
	    <center>	
		<xsl:apply-templates select="orders/order" />
	    </center>
      </body>
    </html>

  </xsl:template>

  <xsl:template match="order" >
      <xsl:variable name="BackOrderCount" select="ItemOnBackOrder"/>
      
<div class="billingLabel" id="billingLabel">
	<table style="border-collapse: collapse" border="0" width="100%" height="0" cellpadding="0" id="tblBillingLabel">
		<tr>
			<td><xsl:apply-templates select="billingAddress" /></td>
		</tr>
	</table>
</div>

<div class="shippingLabel" id="shippingLabel">
	<table style="border-collapse: collapse" border="0" width="100%" height="0" cellpadding="0" id="tblShippingLabel">
		<tr>
			<td>BorderDogs.com<br />
			PO Box 153162<br />
			Tampa, FL 33684<br />
			USA<hr />
			<xsl:apply-templates select="shippingAddress" />
			</td>
		</tr>
	</table>
</div>

<div class="shippingMethodLabel" id="shippingMethodLabel">
	<table style="border-collapse: collapse" border="0" width="100%" height="0" cellpadding="0" id="tblShippingMethodLabel">
		<tr>
			<td>Ship via: <xsl:value-of select="orderShipMethod"/></td>
		</tr>
	</table>
</div>

<div class="orderItemsLabel" id="orderItemsLabel">
	<table style="border-collapse: collapse" border="0" width="100%" height="0" cellpadding="0" id="tblOrderItemsLabel">
		<tr>
			<td colspan="2">
			  <table width="100%" border="0" cellspacing="0" cellpadding="2">
				<tr>
				  <xsl:attribute name="bgcolor">lightgrey</xsl:attribute>
				  <td width="60">Item</td>
				  <td width="160">Code</td>
				  <td width="70%">Description</td>
				  <td align="center" width="120">Unit&amp;nbsp;Price</td>
				  <td align="center" width="120">Order&amp;nbsp;Qty</td>
				  <xsl:if test="$BackOrderCount[.&gt;0]"><td align="center" width="120">Qty&amp;nbsp;on&amp;nbsp;Backorder</td></xsl:if>
				  <td width='160'>Price</td>
				</tr>

				<xsl:for-each select="orderDetail">
					<xsl:variable name="EvenRow" select="position() mod 2"/>
					<tr>
						<xsl:if test="$EvenRow = 0">
							<xsl:attribute name="bgcolor">lightgrey</xsl:attribute>
						</xsl:if>

						<td><xsl:number value="position()" format="1"/></td>
						<td><xsl:value-of select="odrdtProductID" /></td>
						<td nowrap=""><xsl:value-of select="odrdtProductName" /></td>
						<td align="right">$<xsl:value-of select="format-number(number(odrdtPrice),'###0.00')" />&amp;nbsp;</td>
						<td align="center"><xsl:value-of select="odrdtQuantity" /></td>
						<xsl:if test="$BackOrderCount[.&gt;0]">
						<td align="center"><xsl:value-of select="odrdtBackOrderQTY" /></td>
						</xsl:if>
						<td align="right">$<xsl:value-of select="format-number(number(odrdtSubTotal),'###0.00')" />&amp;nbsp;</td>
					</tr>

					<xsl:for-each select="odrdtAttDetailID">
						<tr>
							<xsl:if test="$EvenRow = 0">
								<xsl:attribute name="bgcolor">lightgrey</xsl:attribute>
							</xsl:if>
							<td colspan="2">&amp;nbsp;</td>
							<td>
							<xsl:if test="$BackOrderCount[.=0]">
								<xsl:attribute name="colspan">4</xsl:attribute>
							</xsl:if>
							<xsl:if test="$BackOrderCount[.&gt;0]">
								<xsl:attribute name="colspan">5</xsl:attribute>
							</xsl:if>
							&amp;nbsp;&amp;nbsp;<span class="orderItemText"><xsl:value-of select="odrattrName" />: <xsl:value-of select="odrattrAttribute" /></span></td>
						</tr>
					</xsl:for-each>

					<xsl:if test="odrdtGiftWrapQTY[.&gt;0]">
					  <tr>
						<xsl:if test="$EvenRow = 0">
							<xsl:attribute name="bgcolor">lightgrey</xsl:attribute>
						</xsl:if>
						<td colspan="2">&amp;nbsp;</td>
						<td>&amp;nbsp;&amp;nbsp;Gift Wrap</td>
						<td align="right">$<xsl:value-of select="format-number(number(odrdtGiftWrapUnitPrice),'###0.00')" />&amp;nbsp;</td>
						<td align="center"><xsl:value-of select="odrdtGiftWrapQTY" /></td>
						<xsl:if test="$BackOrderCount[.&gt;0]">
							<td>&amp;nbsp;</td>
						</xsl:if>
						<td align="right">$<xsl:value-of select="format-number(number(odrdtGiftWrapPrice),'###0.00')" />&amp;nbsp;</td>
					</tr>
					</xsl:if>
				</xsl:for-each>

				<tr>
					<td>
					<xsl:if test="$BackOrderCount[.=0]">
						<xsl:attribute name="colspan">6</xsl:attribute>
					</xsl:if>
					<xsl:if test="$BackOrderCount[.&gt;0]">
						<xsl:attribute name="colspan">7</xsl:attribute>
					</xsl:if>
				  <hr width="100%" /></td>
				</tr>
				
				<tr>
				  <td colspan="2">&amp;nbsp;</td>
				  <td align="right">
					<xsl:if test="$BackOrderCount[.=0]">
						<xsl:attribute name="colspan">4</xsl:attribute>
					</xsl:if>
					<xsl:if test="$BackOrderCount[.&gt;0]">
						<xsl:attribute name="colspan">5</xsl:attribute>
					</xsl:if>
				  <table border="0" cellspacing="0" cellpadding="2">
					<tr>
						<td>Sub Total:</td>
						<td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderSubTotal),'###0.00')" /></td>
					</tr>
					
					<xsl:if test="orderCouponDiscount[.&gt;0]">
					<tr>
					  <td>Coupon (<xsl:value-of select="orderCouponCode" />):</td>
					  <td align="right">&amp;nbsp;($<xsl:value-of select="format-number(number(orderCouponDiscount),'###0.00')" />)</td>
					</tr>
					</xsl:if>
					
					<xsl:if test="orderDiscount[.&gt;0]">
					<tr>
					  <td>Discount:</td>
					  <td align="right">&amp;nbsp;($<xsl:value-of select="format-number(number(orderDiscount),'###0.00')" />)</td>
					</tr>
					</xsl:if>
					
					<tr>
						<td><xsl:value-of select="orderShipMethod" />:</td>
						<td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderShippingAmount),'###0.00')" /></td>
					</tr>
					
					<xsl:if test="orderHandling[.&gt;0]">
					<tr>
					  <td>Handling:</td>
					  <td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderHandling),'###0.00')" /></td>
					</tr>
					</xsl:if>
					
					<xsl:if test="orderSTax[.&gt;0]">
					<tr>
					  <td><xsl:value-of select="shippingAddress/State" /> Tax:</td>
					  <td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderSTax),'###0.00')" /></td>
					</tr>
					</xsl:if>
					
					<xsl:if test="orderCTax[.&gt;0]">
					<tr>
					  <td><xsl:value-of select="shippingAddress/Country" /> Tax:</td>
					  <td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderCTax),'###0.00')" /></td>
					</tr>
					</xsl:if>
					
					<tr>
					  <td colspan="2" align="center"><hr/></td>
					</tr>
					<tr>
					  <td>Grand Total:</td>
					  <td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderGrandTotal),'###0.00')" /></td>
					</tr>
					
					<xsl:for-each select="ssGiftCertificate">
						<xsl:if test="RedemptionAmount[.&gt;0]">
						<tr>
						<td>Gift Certificate (<xsl:value-of select="CertificateNumber" />):</td>
						<td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(RedemptionAmount),'###0.00')" /></td>
						</tr>
						</xsl:if>
					</xsl:for-each>
					
					<xsl:if test="orderBackOrderAmount[.&gt;0]">
					<tr>
					  <td>Billed Amount:</td>
					  <td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderBillAmount),'###0.00')" /></td>
					</tr>
					<tr>
					  <td>Remaining Amount:</td>
					  <td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderBackOrderAmount),'###0.00')" /></td>
					</tr>
					</xsl:if>

					</table></td>
				</tr>
			</table></td>
		</tr>
		<xsl:if test="orderComments[.&gt;'']">
		<tr>
			<td>&amp;nbsp;</td>
			<td>&amp;nbsp;</td>
		</tr>
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
</div>

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
        <xsl:value-of select="FirstName"/>&amp;nbsp;
		<xsl:choose>
			<xsl:when test="MiddleInitial[.&gt;'']">
				<xsl:value-of select="MiddleInitial" />&amp;nbsp;
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
        <xsl:value-of select="City"/>, <xsl:value-of select="State"/><xsl:text> </xsl:text><xsl:value-of select="Zip"/>
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
        <xsl:value-of select="FirstName"/>&amp;nbsp;
		<xsl:choose>
			<xsl:when test="MiddleInitial[.='']">&amp;nbsp;
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
        <xsl:value-of select="City"/>, <xsl:value-of select="State"/><xsl:text> </xsl:text><xsl:value-of select="Zip"/>
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