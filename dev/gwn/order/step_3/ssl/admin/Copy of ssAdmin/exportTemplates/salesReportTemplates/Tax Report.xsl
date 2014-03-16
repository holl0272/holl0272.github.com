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
	<table border="1" cellspacing="0" class="tbl">
		<colgroup>
			<col align="left" />
			<col align="left" />
			<col align="left" />
			<col align="right" />
			<col align="right" />
			<col align="right" />
		</colgroup>
		<caption>Product Sales Report</caption>
		<thead>
		<tr class="tblhdr">
			<th>
				<xsl:attribute name="style">cursor:hand</xsl:attribute>
				<xsl:attribute name="onclick">onSort('orderDetail/odrdtProductID', 'descending', 'text', 'imgSort_0');</xsl:attribute>
				<xsl:text>Product Code</xsl:text>
				<img src="images/transparent.gif" id="imgSort_0" border="0" align="bottom" />
			</th>
			<th>
				<xsl:attribute name="style">cursor:hand</xsl:attribute>
				<xsl:attribute name="onclick">onSort('orderDetail/odrdtProductName', 'descending', 'text', 'imgSort_1');</xsl:attribute>
				<xsl:text>Product Name</xsl:text>
				<img src="images/transparent.gif" id="imgSort_1" border="0" align="bottom" />
			</th>
			<th>
				<xsl:text>Short Description</xsl:text>
			</th>
			<th>
				<xsl:text>Total Qty</xsl:text>
			</th>
			<th>
				<xsl:text>Unit Price</xsl:text>
			</th>
			<th>
				<xsl:text>Total Sales</xsl:text>
			</th>
		</tr>
		</thead>
		<tbody>
			<xsl:for-each select="msxsl:node-set($productKeys)/key">
				<xsl:sort select="odrdtProductID" data-type="text" order="descending" />
				<xsl:variable name="productUID" select="string(.)" />

				<!-- Generate a list of orders for the given product -->
				<xsl:variable name="orderDetailNodes">
					<xsl:for-each select="$rootNode/orders/order/orderDetail[odrdtProductID = $productUID]">
						<xsl:sort select="odrdtProductID"/>
						<xsl:variable name="PriceSoldAt">
							<xsl:choose>
								<xsl:when test="IsOnSale = 1"><xsl:value-of select="number(SalePrice)" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="number(Price)" /></xsl:otherwise>
							</xsl:choose>
						</xsl:variable>
						
						<Price><xsl:value-of select="number($PriceSoldAt)" /></Price>
						<Quantity><xsl:value-of select="number(Quantity)" /></Quantity>
						<extPrice><xsl:value-of select="number($PriceSoldAt * odrdtQuantity)" /></extPrice>
					</xsl:for-each>
				</xsl:variable>

				<xsl:for-each select="$rootNode/orders/order/orderDetail[odrdtProductID = $productUID]">
					<xsl:variable name="PriceSoldAt">
						<xsl:choose>
							<xsl:when test="IsOnSale = 1"><xsl:value-of select="number(SalePrice)" /></xsl:when>
							<xsl:otherwise><xsl:value-of select="number(Price)" /></xsl:otherwise>
						</xsl:choose>
					</xsl:variable>

					<xsl:if test="false">
						<fieldset><legend>Order <xsl:value-of select="../orderID" />: <xsl:value-of select="odrdtProductName" /></legend>
							<xsl:text>Price: </xsl:text><xsl:value-of select="Price" /><br />
							<xsl:text>Sale Price: </xsl:text><xsl:value-of select="SalePrice" /><br />
							<xsl:text>IsOnSale: </xsl:text><xsl:value-of select="IsOnSale" /><br />
							<xsl:text>Sold At: </xsl:text><xsl:value-of select="$PriceSoldAt" /><br />
							<xsl:text>Quantity: </xsl:text><xsl:value-of select="odrdtQuantity" /><br />
							<xsl:text>extPrice: </xsl:text><xsl:value-of select="number($PriceSoldAt * odrdtQuantity)" /><br />
						</fieldset>
					</xsl:if>
				</xsl:for-each>
				
				<tr>
					<td>
						<xsl:call-template name="productInfo">
						<xsl:with-param name="productUID" select="$productUID" />
						<xsl:with-param name="Default" select="$rootNode/orders/order/orderDetail[odrdtProductID = $productUID]/odrdtProductID" />
						<xsl:with-param name="returnNode" select="'Code'" />
						<xsl:with-param name="rootNode" select="$rootNode" />
						</xsl:call-template> 
					</td>
					<td>
						<xsl:call-template name="productInfo">
						<xsl:with-param name="productUID" select="$productUID" />
						<xsl:with-param name="Default" select="$rootNode/orders/order/orderDetail[odrdtProductID = $productUID]/odrdtProductName" />
						<xsl:with-param name="returnNode" select="'Name'" />
						<xsl:with-param name="rootNode" select="$rootNode" />
						</xsl:call-template>
					</td>
					<td>
						<xsl:call-template name="productInfo">
						<xsl:with-param name="productUID" select="$productUID" />
						<xsl:with-param name="Default" select="'- N/A -'" />
						<xsl:with-param name="returnNode" select="'ShortDescription'" />
						<xsl:with-param name="rootNode" select="$rootNode" />
						</xsl:call-template>
					</td>
					<td>
						<xsl:value-of select="format-number(sum(msxsl:node-set($orderDetailNodes)/odrdtQuantity), '#,##0')" />
					</td>
					<td>
						<xsl:value-of select="format-number(number(sum(msxsl:node-set($orderDetailNodes)/extPrice)) div number(sum(msxsl:node-set($orderDetailNodes)/odrdtQuantity)), '#,##0.00')" />
					</td>
					<td>
						<xsl:value-of select="format-number(sum(msxsl:node-set($orderDetailNodes)/extPrice), '#,##0.00')" />
					</td>
				</tr>
			</xsl:for-each>
		</tbody>
		<xsl:variable name="DiscountTotal" select="-1 * sum($rootNode/orders/order/TotalAppliedDiscounts)" />
		<xsl:variable name="SubTotal" select="sum($rootNode/orders/order/SubTotal)" />
		<xsl:variable name="SubTotalBeforeDiscounts" select="$SubTotal - $DiscountTotal" />
		<xsl:variable name="ShippingTotal" select="sum($rootNode/orders/order/ShippingTotal)" />
		<xsl:variable name="HandlingTotal" select="sum($rootNode/orders/order/HandlingTotal)" />
		<xsl:variable name="PreTaxTotal" select="$SubTotalBeforeDiscounts + $ShippingTotal + $HandlingTotal + $DiscountTotal" />
		<tfoot>
			<tr>
				<td colspan="3" align="right">Merchandise Total</td>
				<td colspan="3" align="right"><xsl:value-of select="format-number($SubTotalBeforeDiscounts, '#,##0.00')" /></td>
			</tr>
			<tr>
				<td colspan="3" align="right">Shipping Fees</td>
				<td colspan="3" align="right"><xsl:value-of select="format-number($ShippingTotal, '#,##0.00')" /></td>
			</tr>
			<tr>
				<td colspan="3" align="right">Handling Fees</td>
				<td colspan="3" align="right"><xsl:value-of select="format-number($HandlingTotal, '#,##0.00')" /></td>
			</tr>
			<tr>
				<td colspan="3" align="right">Discounts</td>
				<td colspan="3" align="right"><xsl:value-of select="format-number($DiscountTotal, '#,##0.00')" /></td>
			</tr>
			<tr>
				<td colspan="3" align="right">Total before sales taxes</td>
				<td colspan="3" align="right"><xsl:value-of select="format-number($PreTaxTotal, '#,##0.00')" /></td>
			</tr>
			<tr>
				<td colspan="3" align="right">CA Local and State Tax</td>
				<td colspan="3" align="right">
					<xsl:call-template name="taxCollectedByState">
						<xsl:with-param name="rootNode" select="$rootNode" />
						<xsl:with-param name="State" select="'CA'" />
						<xsl:with-param name="collectionMode" select="'4'" />
					</xsl:call-template>
				</td>
			</tr>
			<tr>
				<td colspan="3" align="right">CA State Tax</td>
				<td colspan="3" align="right">
					<xsl:call-template name="taxCollectedByState">
						<xsl:with-param name="rootNode" select="$rootNode" />
						<xsl:with-param name="State" select="'CA'" />
						<xsl:with-param name="collectionMode" select="'2'" />
					</xsl:call-template>
				</td>
			</tr>
			<tr>
				<td colspan="3" align="right">OH State Tax</td>
				<td colspan="3" align="right">
					<xsl:call-template name="taxCollectedByState">
						<xsl:with-param name="rootNode" select="$rootNode" />
						<xsl:with-param name="State" select="'OH'" />
						<xsl:with-param name="collectionMode" select="'1'" />
					</xsl:call-template>
				</td>
			</tr>
			<tr>
				<td colspan="3" align="right">TX State Tax</td>
				<td colspan="3" align="right">
					<xsl:call-template name="taxCollectedByState">
						<xsl:with-param name="rootNode" select="$rootNode" />
						<xsl:with-param name="State" select="'TX'" />
						<xsl:with-param name="collectionMode" select="'1'" />
					</xsl:call-template>
				</td>
			</tr>
			<tr>
				<td colspan="3" align="right">IN State Tax</td>
				<td colspan="3" align="right">
					<xsl:call-template name="taxCollectedByState">
						<xsl:with-param name="rootNode" select="$rootNode" />
						<xsl:with-param name="State" select="'IN'" />
						<xsl:with-param name="collectionMode" select="'1'" />
					</xsl:call-template>
				</td>
			</tr>
			<tr>
				<td colspan="3" align="right">VA Local and State Tax</td>
				<td colspan="3" align="right">
					<xsl:call-template name="taxCollectedByState">
						<xsl:with-param name="rootNode" select="$rootNode" />
						<xsl:with-param name="State" select="'VA'" />
						<xsl:with-param name="collectionMode" select="'4'" />
					</xsl:call-template>
				</td>
			</tr>
			<tr>
				<td colspan="3" align="right">Grand Total</td>
				<td colspan="3" align="right"><xsl:value-of select="format-number(sum($rootNode/orders/order/GrandTotal) + sum($rootNode/orders/order/GiftCertificateTotal), '#,##0.00')" /></td>
			</tr>
			<tr>
				<td colspan="3" align="right">Gift Certificates Redeemed</td>
				<td colspan="3" align="right"><xsl:value-of select="format-number(sum($rootNode/orders/order/GiftCertificateTotal), '#,##0.00')" /></td>
			</tr>
			<tr>
				<td colspan="3" align="right">Amount Collected</td>
				<td colspan="3" align="right"><xsl:value-of select="format-number(sum($rootNode/orders/order/TotalBilled), '#,##0.00')" /></td>
			</tr>
		</tfoot>
	</table>

  </xsl:template>

</xsl:stylesheet>