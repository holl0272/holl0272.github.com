<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:template match="/">
		<html>
			<head>
	  </head>
			<body>
				<center>
					<xsl:apply-templates select="orders/order"/>
				</center>
			</body>
		</html>
	</xsl:template>
	<xsl:template match="order">
		<table border="0" cellspacing="0" cellpadding="0">
			<xsl:if test="position() = 1"><xsl:attribute name="style">width: 3.5in; border-collapse: collapse</xsl:attribute></xsl:if>
			<xsl:if test="position() > 1"><xsl:attribute name="style">width: 3.5in; border-collapse: collapse; page-break-before:always;</xsl:attribute></xsl:if>
					<colgroup>
						<col align="left" valign="top" width="3.5in"></col>
						<col align="left" valign="top" width="0.5in"></col>
						<col align="right" valign="top" width="3.5in"></col>
					</colgroup>
			<tr>
				<td>
					<table style="width: 3.5in; border-collapse: collapse" border="0" cellpadding="0" id="tblBillingLabel">
						<tr>
							<td><xsl:apply-templates select="billingAddress" /></td>
						</tr>
					</table>
					<br />
					<table style="width: 3.5in; border-collapse: collapse" border="0" cellpadding="0" id="tblShippingMethodLabel">
						<tr>
							<td>Ship via: <xsl:value-of select="orderShipMethod"/></td>
						</tr>
					</table>
				</td>
				<td>&amp;nbsp;</td>
				<td style="width: 3.5in; height:2.0in">
					<table style="width: 3.5in; height:2.0in;border-collapse: collapse" border="0" cellpadding="0" id="tblShippingLabel">
						<tr>
							<td nowrap="">BorderDogs.com<br />
							PO Box 153162<br />
							Tampa, FL 33684<br />
							USA<hr />
							<xsl:apply-templates select="shippingAddress" />
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="3">
				</td>
			</tr>
			<tr>
				<td colspan="3">
					<br />
					<table width="100%" border="0" cellspacing="0" cellpadding="2">
					<colgroup>
						<col align="left" width="25%"></col>
						<col align="left" width="50%"></col>
						<col align="center" width="25%"></col>
					</colgroup>
						<tr>
							<td>Item</td>
							<td>Name</td>
							<td>Quantity</td>
						</tr>
						<xsl:for-each select="orderDetail">
							<tr>
								<xsl:if test="position() mod 2 = 0">
									<xsl:attribute name="bgcolor">yellow</xsl:attribute>
								</xsl:if>
								<td>
									<xsl:value-of select="odrdtProductID"/>
								</td>
								<td>
									<xsl:value-of select="odrdtProductName"/>
								</td>
								<td>
									<xsl:value-of select="odrdtQuantity"/>
								</td>
							</tr>
							<xsl:for-each select="odrdtAttDetailID">
								<tr bgcolor="">
									<td>&amp;nbsp;</td>
									<td colspan="2">&amp;nbsp;&amp;nbsp;<span class="orderItemText">
											<xsl:value-of select="odrattrName"/>: <xsl:value-of select="odrattrAttribute"/>
										</span>
									</td>
								</tr>
							</xsl:for-each>
						</xsl:for-each>
					</table>
				</td>
			</tr>
			<xsl:if test="string-length(orderComments)>0">
				<tr>
					<td colspan="2">
						<table width="100%" border="1" cellspacing="0" cellpadding="2">
							<tr>
								<td>COMMENTS:<br/>
									<xsl:value-of select="orderComments"/>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</xsl:if>
		</table>
	</xsl:template>

	<xsl:template match="shippingAddress">
					<xsl:value-of select="FirstName"/>
					<xsl:choose>
						<xsl:when test="MiddleInitial[.&gt;'']">
							<xsl:text> </xsl:text><xsl:value-of select="MiddleInitial"/>
						</xsl:when>
					</xsl:choose>
					<xsl:text> </xsl:text><xsl:value-of select="LastName"/><br />
					
			<xsl:choose>
				<xsl:when test="string-length(Company)>0">
							<xsl:value-of select="Company"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="string-length(Addr1)>0">
							<xsl:value-of select="Addr1"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="string-length(Addr2)>0">
							<xsl:value-of select="Addr2"/><br />
				</xsl:when>
			</xsl:choose>
					<xsl:value-of select="City"/>, 
        <xsl:value-of select="State"/>
					<xsl:value-of select="Zip"/><br />
			<xsl:choose>
				<xsl:when test="Country[.!='US']">
							<xsl:value-of select="CountryName"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="Phone[.&gt;'']">
						Phone: <xsl:value-of select="Phone"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="Fax[.&gt;'']">
					Fax: <xsl:value-of select="Fax"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="Email[.&gt;'']">
							<xsl:value-of select="Email"/><br />
				</xsl:when>
			</xsl:choose>
	</xsl:template>
	<xsl:template match="billingAddress">
					<xsl:value-of select="FirstName"/>
					<xsl:choose>
						<xsl:when test="MiddleInitial[.&gt;'']">
							<xsl:text> </xsl:text><xsl:value-of select="MiddleInitial"/>
						</xsl:when>
					</xsl:choose>
					<xsl:text> </xsl:text><xsl:value-of select="LastName"/><br />
					
			<xsl:choose>
				<xsl:when test="string-length(Company)>0">
							<xsl:value-of select="Company"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="string-length(Addr1)>0">
							<xsl:value-of select="Addr1"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="string-length(Addr2)>0">
							<xsl:value-of select="Addr2"/><br />
				</xsl:when>
			</xsl:choose>
					<xsl:value-of select="City"/>, 
        <xsl:value-of select="State"/>
					<xsl:value-of select="Zip"/><br />
			<xsl:choose>
				<xsl:when test="Country[.!='US']">
							<xsl:value-of select="CountryName"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="Phone[.&gt;'']">
						Phone: <xsl:value-of select="Phone"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="Fax[.&gt;'']">
					Fax: <xsl:value-of select="Fax"/><br />
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="Email[.&gt;'']">
							<xsl:value-of select="Email"/><br />
				</xsl:when>
			</xsl:choose>
	</xsl:template>
</xsl:stylesheet>
