<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template match='/'>
    <html>
      <head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
		<title>Sales Receipt</title>
		<link rel="stylesheet" href="exportTemplates/ssTemplateStyleSheet.css" type="text/css" />
	  </head>
      <body>
	    <center>	
		<xsl:apply-templates select="orders/order" />
	    </center>
      </body>
    </html>

  </xsl:template>

  <xsl:template match="order" >
	<xsl:variable name="needsPageBreak" select="position() > 1"/>
    <xsl:variable name="BackOrderCount" select="ItemOnBackOrder"/>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<xsl:if test="position() > 1"><xsl:attribute name="style">page-break-before:always</xsl:attribute></xsl:if>
<tr><td align="center">
<table width="99%" border="0" cellpadding="0">
  <tr> 
    <td width="47%" valign="top">
        <img src="images/logo.jpg" alt="Your Logo" /><br />
		<p class="MerchantAddress">
        1958 Brooke Farm Court<br />
        Woodbridge VA 22192<br />
        (703) 507-7330<br />
        <br />
        http://www.sandshot.net<br />
        </p>
        <p></p>
    </td>
    <td width="53%" valign="top" align="right">
        <p class="PageTitle">Sales Receipt</p>
        <table border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td align="right" class="OrderMisc">Order #:&amp;nbsp;&amp;nbsp;</td>
            <td align="right" class="OrderMisc"><xsl:value-of select="orderID"/></td>
          </tr>
          <tr> 
            <td align="right" class="OrderMisc">Date Ordered:&amp;nbsp;&amp;nbsp;</td>
            <td align="right" class="OrderMisc"><xsl:value-of select="shortOrderDate"/></td>
          </tr>
        </table>
        <br />
        <br />
        <p>&amp;nbsp;</p>
      </td>
  </tr>
		<tr><td colspan="2">&amp;nbsp;</td></tr>
		<tr>
			<td width="50%" valign="top">
			<table width="63%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
				<tr> 
				<th align="left" class="AddressHeader">&amp;nbsp;Bill To:</th>
				</tr>
				<tr> 
				<td>
					<table>
						<tr>
							<td class="AddressName">
							<xsl:call-template name="outputName"><xsl:with-param name='currentNode' select='./billingAddress' /></xsl:call-template>
							</td>
						</tr>
						<xsl:if test="string-length(billingAddress/Company)>0">
							<tr><td class="AddressBody"><xsl:value-of select="billingAddress/Company" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(billingAddress/Addr1)>0">
							<tr><td class="AddressBody"><xsl:value-of select="billingAddress/Addr1" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(billingAddress/Addr2)>0">
							<tr><td class="AddressBody"><xsl:value-of select="billingAddress/Addr2" /></td></tr>
						</xsl:if>
						<tr>
							<td class="AddressBody">
							<xsl:value-of select="billingAddress/City"/>, <xsl:value-of select="billingAddress/State"/><xsl:text> </xsl:text><xsl:value-of select="billingAddress/Zip"/>
							</td>
						</tr>
						<xsl:if test="billingAddress/Country[.!='US']">
							<tr><td class="AddressBody"><xsl:value-of select="billingAddress/CountryName" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(billingAddress/Phone)>0">
							<tr><td class="AddressBody"><xsl:value-of select="billingAddress/Phone" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(billingAddress/Fax)>0">
							<tr><td class="AddressBody"><xsl:value-of select="billingAddress/Fax" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(billingAddress/Email)>0">
							<tr><td class="AddressBody"><xsl:value-of select="billingAddress/Email" /></td></tr>
						</xsl:if>
					</table>
				</td>
				</tr> 
			</table>
			</td>
			<td width="50%" valign="top" align="right">
			<table width="63%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
				<tr> 
				<th align="left" class="AddressHeader">&amp;nbsp;Ship To:</th>
				</tr>
				<tr> 
				<td>
					<table>
						<tr>
							<td class="AddressName">
							<xsl:call-template name="outputName"><xsl:with-param name='currentNode' select='./shippingAddress' /></xsl:call-template>
							</td>
						</tr>
						<xsl:if test="string-length(shippingAddress/Company)>0">
							<tr><td class="AddressBody"><xsl:value-of select="shippingAddress/Company" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(shippingAddress/Addr1)>0">
							<tr><td class="AddressBody"><xsl:value-of select="shippingAddress/Addr1" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(shippingAddress/Addr2)>0">
							<tr><td class="AddressBody"><xsl:value-of select="shippingAddress/Addr2" /></td></tr>
						</xsl:if>
						<tr>
							<td class="AddressBody">
							<xsl:value-of select="shippingAddress/City"/>, <xsl:value-of select="shippingAddress/State"/><xsl:text> </xsl:text><xsl:value-of select="shippingAddress/Zip"/>
							</td>
						</tr>
						<xsl:if test="shippingAddress/Country[.!='US']">
							<tr><td class="AddressBody"><xsl:value-of select="shippingAddress/CountryName" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(shippingAddress/Phone)>0">
							<tr><td class="AddressBody"><xsl:value-of select="shippingAddress/Phone" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(shippingAddress/Fax)>0">
							<tr><td class="AddressBody"><xsl:value-of select="shippingAddress/Fax" /></td></tr>
						</xsl:if>
						<xsl:if test="string-length(shippingAddress/EMail)>0">
							<tr><td class="AddressBody"><xsl:value-of select="shippingAddress/EMail" /></td></tr>
						</xsl:if>
					</table>
				</td>
				</tr> 
			</table></td>
		</tr>

		<tr>
		  <td colspan="2" align="center">
		<table border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
			<tr> 
			<th class="AddressHeader">Method of Payment</th>
			<th class="AddressHeader">Ship Via</th>
			</tr>
			<tr>
			<td align="center">&amp;nbsp;&amp;nbsp;<xsl:value-of select="orderPaymentMethod" />&amp;nbsp;&amp;nbsp;</td>
			<td align="center">&amp;nbsp;&amp;nbsp;<xsl:value-of select="orderShipMethod" />&amp;nbsp;&amp;nbsp;</td>
			</tr>
		</table>
			</td>
		</tr>
		<tr>
			<td colspan="2">
<table width="99%" border="1" cellpadding="3" cellspacing="0" bordercolor="#000000">
  <colgroup width="10%" align="left" valign="top" />
  <colgroup width="60%" align="left" valign="top" />
  <colgroup width="10%" align="center" valign="top" />
  <colgroup width="15%" align="right" valign="top" />
  <colgroup width="15%" align="right" valign="top" />
				<tr class="ProductHeader">
				  <td>Code</td>
				  <td>Description</td>
				  <td>Qty</td>
				  <xsl:if test="$BackOrderCount[.&gt;0]"><td>Qty&amp;nbsp;on&amp;nbsp;Backorder</td></xsl:if>
				  <td>Unit&amp;nbsp;Price</td>
				  <td>Price</td>
				</tr>

				<xsl:for-each select="orderDetail">
					<xsl:variable name="EvenRow" select="position() mod 2"/>
					<tr>
						<xsl:if test="$EvenRow = 0">
							<xsl:attribute name="class">ProductEvenRows</xsl:attribute>
						</xsl:if>
						<xsl:if test="$EvenRow != 0">
							<xsl:attribute name="class">ProductOddRows</xsl:attribute>
						</xsl:if>

						<td><xsl:value-of select="odrdtProductID" /></td>
						<td nowrap="">
							<xsl:value-of select="odrdtProductName" />
							<xsl:for-each select="odrdtAttDetailID">
							<br />&amp;nbsp;&amp;nbsp;<span class="orderItemText"><xsl:value-of select="odrattrName" />: <xsl:value-of select="odrattrAttribute" /></span>
							</xsl:for-each>
						</td>
						<td><xsl:value-of select="odrdtQuantity" /></td>
						<xsl:if test="$BackOrderCount[.&gt;0]">
						<td><xsl:value-of select="odrdtBackOrderQTY" /></td>
						</xsl:if>
						<td>$<xsl:value-of select="format-number(number(odrdtPrice),'###0.00')" />&amp;nbsp;</td>
						<td>$<xsl:value-of select="format-number(number(odrdtSubTotal),'###0.00')" />&amp;nbsp;</td>
					</tr>


					<xsl:if test="odrdtGiftWrapQTY[.&gt;0]">
					  <tr>
						<xsl:if test="$EvenRow = 0">
							<xsl:attribute name="class">ProductEvenRows</xsl:attribute>
						</xsl:if>
						<xsl:if test="$EvenRow != 0">
							<xsl:attribute name="class">ProductOddRows</xsl:attribute>
						</xsl:if>

						<td>&amp;nbsp;</td>
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
				  <td align="right">
					<xsl:if test="$BackOrderCount[.=0]">
						<xsl:attribute name="colspan">5</xsl:attribute>
					</xsl:if>
					<xsl:if test="$BackOrderCount[.&gt;0]">
						<xsl:attribute name="colspan">6</xsl:attribute>
					</xsl:if>
				  <table border="0" cellspacing="0" cellpadding="2">
					<tr>
						<td>Sub Total:</td>
						<td align="right">&amp;nbsp;$<xsl:value-of select="format-number(number(orderSubTotal),'###0.00')" /></td>
					</tr>
					
					<xsl:if test="orderCouponDiscount[.&gt;0] and string-length(orderCouponCode)>0">
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
</td></tr>
</table>
<p align="center" class="Footer">Thank you for your business!<br /></p>
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


<xsl:template name="outputName">
	<xsl:param name="currentNode" />
	<xsl:value-of select="$currentNode/FirstName"/>&amp;nbsp;
	<xsl:if test="string-length($currentNode/MiddleInitial)>0"><xsl:value-of select="$currentNode/MiddleInitial" /><xsl:text> </xsl:text></xsl:if>
	<xsl:value-of select="$currentNode/LastName" />
</xsl:template>

</xsl:stylesheet>