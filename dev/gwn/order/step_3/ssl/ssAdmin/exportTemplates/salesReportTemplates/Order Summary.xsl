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

<!-- Inbound sorting parameters -->
<xsl:param name="sortSelect">orderID</xsl:param>
<xsl:param name="sortType">number</xsl:param>
<xsl:param name="sortOrder">descending</xsl:param>

  <xsl:template match="/">
	
	<xsl:call-template name="ordersInReportListing" /><hr />

  </xsl:template>

	<!-- Displays row of order numbers which are in the report -->
	<xsl:template name="ordersInReportListing">
		<table border="1" cellspacing="0" class="tbl">
			<caption>Orders in Report</caption>
			<thead>
			<tr class="tblhdr">
				<th>
					<xsl:attribute name="style">cursor:hand</xsl:attribute>
					<xsl:attribute name="onclick">onSort('orderID', 'descending', 'number', 'imgSort_0');</xsl:attribute>
					<xsl:text>Order Number</xsl:text>
					<img src="images/transparent.gif" id="imgSort_0" border="0" align="bottom" />
				</th>
				<th>
					<xsl:attribute name="style">cursor:hand</xsl:attribute>
					<xsl:attribute name="onclick">onSort('sum(orderDetail/odrdtQuantity)', 'descending', 'number', 'imgSort_1');</xsl:attribute>
					<xsl:text>Item Count</xsl:text>
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
			<xsl:for-each select="orders/order">
				<xsl:sort select="orderID" data-type="number" order="descending" />

			<tr>
				<td><xsl:value-of select="orderID"/></td>
				<td>
					<div>
						<xsl:attribute name="style">cursor:hand</xsl:attribute>
						<xsl:attribute name="onclick">if (document.all("ulOrderDetail<xsl:value-of select="uid"/>").style.display==""){document.all("ulOrderDetail<xsl:value-of select="uid"/>").style.display="none";}else{document.all("ulOrderDetail<xsl:value-of select="uid"/>").style.display="";} return false;</xsl:attribute>
						<xsl:value-of select="sum(orderDetail/odrdtQuantity)"/>
					</div>
				</td>
				<td>
					<xsl:for-each select="orderDetail">
						<xsl:value-of select="odrdtProductID"/>: 
						<span>
						<xsl:value-of select="odrdtProductName"/> (<xsl:value-of select="odrdtQuantity"/>)<br />
						<xsl:for-each select="odrdtAttDetailID"> - <xsl:value-of select="odrattrName" />: <xsl:value-of select="odrattrAttribute" /><br />
						</xsl:for-each>
						</span>
					</xsl:for-each>
				</td>
			</tr>
			</xsl:for-each>
			</tbody>
		</table>
	</xsl:template>

</xsl:stylesheet>