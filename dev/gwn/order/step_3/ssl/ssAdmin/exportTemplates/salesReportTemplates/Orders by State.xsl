<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
      xmlns:msxsl="urn:schemas-microsoft-com:xslt">
<xsl:import href="../supportingTemplates/supportTemplates.xsl"/>
<xsl:import href="supportingTemplates/salesReport_Common.xsl"/>
<xsl:output method="html"/>

<xsl:key name="keyUniqueProducts" match="orderDetail" use="odrdtProductID"/>
<xsl:key name="keyProducts" match="Products" use="uid"/>
<xsl:key name="keyShipToStates" match="shippingAddress" use="State"/>

  <xsl:template match="/">
	
	<!-- Generate a list of unique products in the order -->
	<xsl:variable name="productKeys">
		<xsl:for-each select="//orderDetail[generate-id(.) = generate-id(key('keyUniqueProducts', odrdtProductID)[1])]">
			<xsl:sort select="odrdtProductID"/>
			<key><xsl:copy-of select="odrdtProductID" /></key>
		</xsl:for-each>
	</xsl:variable>

	<!-- Generate a list of unique products in the order -->
	<xsl:variable name="taxKeys">
		<xsl:for-each select="//shippingAddress[generate-id(.) = generate-id(key('keyShipToStates', State)[1])]">
			<xsl:sort select="State"/>
			<state><xsl:copy-of select="State" /></state>
		</xsl:for-each>
	</xsl:variable>

	<xsl:variable name="rootNode" select="." />

	<xsl:call-template name="StatesshippedToKeysInReportListing">
		<xsl:with-param name="productKeys" select="$productKeys" />
	</xsl:call-template><hr />
	
  </xsl:template>

	<!-- Displays states which have been shipped to -->
	<xsl:template name="StatesshippedToKeysInReportListing">
		<xsl:param name="productKeys" />
		<xsl:variable name="rootNode" select="." />

		<table border="1" cellspacing="0" class="tbl">
			<caption>Product Sales By State</caption>
			<thead>
			<tr class="tblhdr">
				<th>
					<xsl:attribute name="style">cursor:hand</xsl:attribute>
					<xsl:attribute name="onclick">onSort('State', 'descending', 'number', 'imgSort_0');</xsl:attribute>
					<xsl:text>State</xsl:text>
					<img src="images/transparent.gif" id="imgSort_0" border="0" align="bottom" />
				</th>
				<th>
					<xsl:attribute name="style">cursor:hand</xsl:attribute>
					<xsl:attribute name="onclick">onSort('orderID', 'descending', 'number', 'imgSort_1');</xsl:attribute>
					<xsl:text>Order</xsl:text>
					<img src="images/transparent.gif" id="imgSort_1" border="0" align="bottom" />
				</th>
				<th>
					<xsl:attribute name="style">cursor:hand</xsl:attribute>
					<xsl:attribute name="onclick">onSort('orderDetail/odrdtProductID', 'descending', 'text', 'imgSort_2');</xsl:attribute>
					<xsl:text>Items</xsl:text>
					<img src="images/transparent.gif" id="imgSort_2" border="0" align="bottom" />
				</th>
			</tr>
			</thead>
			<tbody>

			<xsl:for-each select="$rootNode/orders/order/shippingAddress[generate-id(.) = generate-id(key('keyShipToStates', State)[1])]">
				<xsl:sort select="State"/>
				<xsl:variable name="State" select="string(State)" />

			<tr>
				<td><xsl:copy-of select="$State" /></td>
				<td>
					<xsl:for-each select="$rootNode/orders/order[shippingAddress/State = $State]">
						<xsl:value-of select="orderID" /><br />
					</xsl:for-each>
				</td>
				<td>
					<xsl:for-each select="$rootNode/orders/order[shippingAddress/State = $State]">
						<xsl:for-each select="orderDetail">
							<xsl:value-of select="odrdtProductID"/>: <xsl:value-of select="odrdtProductName"/> (<xsl:value-of select="odrdtQuantity"/>)<br />
						</xsl:for-each>
					</xsl:for-each>
				</td>
			</tr>
			</xsl:for-each>
			</tbody>
		</table>

	</xsl:template>

</xsl:stylesheet>