<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
      xmlns:msxsl="urn:schemas-microsoft-com:xslt">
<!--
This file contains supporting Common variables and templates

Common variables:

	
Common templates:


-->

	<!-- Displays row of order numbers which are in the report -->
	<xsl:template name="ordersInReportListing">
		<table border="1" cellspacing="0" class="tbl">
			<caption>Orders in Report</caption>
			<thead>
			<tr class="tblhdr">
				<th>Order Number</th>
				<th>Item Count</th>
			</tr>
			</thead>
			<tbody>
			<xsl:for-each select="orders/order">
				<xsl:sort select="OrderNumber"/>
			<tr>
				<td><xsl:value-of select="OrderNumber"/></td>
				<td>
					<div>
						<xsl:attribute name="style">cursor:hand</xsl:attribute>
						<xsl:attribute name="onclick">if (document.all("ulOrderDetail<xsl:value-of select="uid"/>").style.display==""){document.all("ulOrderDetail<xsl:value-of select="uid"/>").style.display="none";}else{document.all("ulOrderDetail<xsl:value-of select="uid"/>").style.display="";} return false;</xsl:attribute>
						<xsl:value-of select="sum(orderDetail/Quantity)"/>
					</div>
					<ul>
					<xsl:attribute name="style">display:xnone</xsl:attribute>
					<xsl:attribute name="id">ulOrderDetail<xsl:value-of select="uid"/></xsl:attribute>
					<xsl:for-each select="orderDetail">
						<li><xsl:value-of select="ProductCode"/>: <xsl:value-of select="ProductName"/> (<xsl:value-of select="Quantity"/>)</li>
					</xsl:for-each>
					</ul>
				</td>
			</tr>
			</xsl:for-each>
			</tbody>
		</table>
	</xsl:template>

	<!-- Displays product keys which are in the report -->
	<xsl:template name="productKeysInReportListing">
		<xsl:param name="productKeys" />
		<xsl:variable name="rootNode" select="." />
		<table border="1" cellspacing="0" class="tbl">
			<caption>Products in Report</caption>
			<thead>
			<tr class="tblhdr">
				<th>Code</th>
				<th>Name</th>
				<th>Current Sell Price</th>
			</tr>
			</thead>
			<tbody>
			<xsl:for-each select="msxsl:node-set($productKeys)/key">
				<xsl:variable name="productUID" select="string(.)" />
				<tr>
					<td><xsl:value-of select="$productUID" /></td>
					<td>
						<xsl:call-template name="productInfo">
						<xsl:with-param name="productUID" select="$productUID" />
						<xsl:with-param name="returnNode" select="'Name'" />
						<xsl:with-param name="rootNode" select="$rootNode" />
						</xsl:call-template>
					</td>
					<td>
						<xsl:call-template name="productCurrentSellPrice">
						<xsl:with-param name="productUID" select="$productUID" />
						<xsl:with-param name="rootNode" select="$rootNode" />
						</xsl:call-template>
					</td>
				</tr>
			</xsl:for-each>
			</tbody>
		</table>
	</xsl:template>

	<!-- Displays states which have been shipped to -->
	<xsl:template name="StatesshippedToKeysInReportListing">
		<xsl:param name="productKeys" />
		<xsl:variable name="rootNode" select="." />

		<xsl:for-each select="$rootNode/orders/order/shippingAddress[generate-id(.) = generate-id(key('keyShipToStates', State)[1])]">
			<xsl:sort select="State"/>
			<xsl:variable name="State" select="string(State)" />
			<h1><xsl:copy-of select="$State" /></h1>
			<ul>
				<xsl:for-each select="$rootNode/orders/order[shippingAddress/State = $State]">
					<li><xsl:value-of select="number(OrderNumber)" /></li>
				</xsl:for-each>
			</ul>
		</xsl:for-each>


		<table border="1" cellspacing="0" class="tbl">
			<caption>Products in Report</caption>
			<thead>
			<tr class="tblhdr">
				<th>Code</th>
				<th>Name</th>
				<th>Current Sell Price</th>
			</tr>
			</thead>
			<tbody>
			<xsl:for-each select="msxsl:node-set($productKeys)/key">
				<xsl:variable name="productUID" select="string(.)" />
				<tr>
					<td><xsl:value-of select="$productUID" /></td>
					<td>
						<xsl:call-template name="productInfo">
							<xsl:with-param name="productUID" select="$productUID" />
							<xsl:with-param name="returnNode" select="'Name'" />
							<xsl:with-param name="rootNode" select="$rootNode" />
						</xsl:call-template>
					</td>
					<td>
						<xsl:call-template name="productCurrentSellPrice">
							<xsl:with-param name="productUID" select="$productUID" />
							<xsl:with-param name="rootNode" select="$rootNode" />
						</xsl:call-template>
					</td>
				</tr>
			</xsl:for-each>
			</tbody>
		</table>
	</xsl:template>

	<!-- Returns piece of product information given desired uid -->
	<xsl:template name="productInfo">
		<xsl:param name="productUID" />
		<xsl:param name="Default" />
		<xsl:param name="returnNode" />
		<xsl:param name="rootNode" />
		
		<xsl:variable name="returnValue">
			<xsl:for-each select="$rootNode/orders/Products/Product[@uid = $productUID]">
				<xsl:for-each select="child::node()">
					<xsl:if test="name(.) = $returnNode"><xsl:value-of select="." /></xsl:if>
				</xsl:for-each>
			</xsl:for-each>
		</xsl:variable>
		
		<xsl:choose>
			<xsl:when test="string-length($returnValue)>0"><xsl:value-of select="$returnValue" /></xsl:when>
			<xsl:otherwise><xsl:value-of select="$Default" /></xsl:otherwise>
		</xsl:choose>

	</xsl:template>

	<!-- Returns piece of product information given desired uid -->
	<xsl:template name="productCurrentSellPrice">
		<xsl:param name="productUID" />
		<xsl:param name="rootNode" />
		
		<xsl:variable name="productNode" select="$rootNode/orders/Products/Product[@uid = $productUID]" />

		<xsl:choose>
			<xsl:when test="$productNode/IsOnSale = 1"><xsl:value-of select="$productNode/SalePrice" /></xsl:when>
			<xsl:otherwise><xsl:value-of select="$productNode/Price" /></xsl:otherwise>
		</xsl:choose>

	</xsl:template>


	<!-- Tax Collected by State -->
	<!-- 0 - State Only, 1 - Local Only, 2 - State and Local -->
	<xsl:template name="taxCollectedByState">
		<xsl:param name="collectionMode" />
		<xsl:param name="State" />
		<xsl:param name="rootNode" />
		
		<xsl:variable name="salesTaxNodes">
			<xsl:for-each select="$rootNode/orders/order[shippingAddress/State = $State]">
				<State><xsl:value-of select="number(StateTax)" /></State>
				<Local><xsl:value-of select="number(LocalTaxTotal)" /></Local>
				<StateAndLocal><xsl:value-of select="number(StateTax + LocalTaxTotal)" /></StateAndLocal>
				<StateOnly>
					<xsl:choose>
						<xsl:when test="number(LocalTaxTotal) = 0"><xsl:value-of select="number(StateTax)" /></xsl:when>
						<xsl:otherwise>0</xsl:otherwise>
					</xsl:choose>
				</StateOnly>
				<LocalOnly>
					<xsl:choose>
						<xsl:when test="number(StateTax) = 0"><xsl:value-of select="number(LocalTaxTotal)" /></xsl:when>
						<xsl:otherwise>0</xsl:otherwise>
					</xsl:choose>
				</LocalOnly>
				<StateAndLocalOnly>
					<xsl:choose>
						<xsl:when test="number(LocalTaxTotal) > 0"><xsl:value-of select="number(StateTax + LocalTaxTotal)" /></xsl:when>
						<xsl:otherwise>0</xsl:otherwise>
					</xsl:choose>
				</StateAndLocalOnly>
			</xsl:for-each>
		</xsl:variable>
				

		<xsl:choose>
			<xsl:when test="$collectionMode = '0'"><xsl:value-of select="format-number(sum(msxsl:node-set($salesTaxNodes)/State), '#,##0.00')" /></xsl:when>
			<xsl:when test="$collectionMode = '1'"><xsl:value-of select="format-number(sum(msxsl:node-set($salesTaxNodes)/Local), '#,##0.00')" /></xsl:when>
			<xsl:when test="$collectionMode = '2'"><xsl:value-of select="format-number(sum(msxsl:node-set($salesTaxNodes)/StateOnly), '#,##0.00')" /></xsl:when>
			<xsl:when test="$collectionMode = '3'"><xsl:value-of select="format-number(sum(msxsl:node-set($salesTaxNodes)/LocalOnly), '#,##0.00')" /></xsl:when>
			<xsl:when test="$collectionMode = '4'"><xsl:value-of select="format-number(sum(msxsl:node-set($salesTaxNodes)/StateAndLocalOnly), '#,##0.00')" /></xsl:when>
			<xsl:otherwise><xsl:value-of select="format-number(sum(msxsl:node-set($salesTaxNodes)/StateAndLocal), '#,##0.00')" /></xsl:otherwise>
		</xsl:choose>

	</xsl:template>

</xsl:stylesheet>