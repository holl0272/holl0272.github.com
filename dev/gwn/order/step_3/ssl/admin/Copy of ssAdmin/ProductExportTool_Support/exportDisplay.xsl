<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html"/>
<!--
'********************************************************************************
'*   Product Export Tool - Froogle Module				                        *
'*   Release Version:   1.00.000												*
'*   Release Date:		May 6, 2004												*
'*   Release Date:		May 6, 2004												*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

-->
<xsl:variable name="outputDelimeter"><xsl:text>,</xsl:text></xsl:variable>
<xsl:variable name="outputDelimeter_CR"><xsl:text>&#10;</xsl:text></xsl:variable>

<xsl:template match='/'>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" xrules="none" id="tblSummary">
  <tr>
    <th align="left" class="hdrNonSelected" style="cursor:auto;">Product ID</th>
    <th align="left" class="hdrNonSelected" style="cursor:auto;">Product</th>
    <th class="hdrNonSelected" style="cursor:auto;">Regular Price</th>
    <th class="hdrNonSelected" style="cursor:auto;">Sale Price</th>
    <th class="hdrNonSelected" style="cursor:auto;">On Sale</th>
    <th class="hdrNonSelected" style="cursor:auto;">Category</th>
  </tr>

	<xsl:call-template name="products"></xsl:call-template>
</table>
</xsl:template>

<xsl:template name="products" >
	<xsl:for-each select="products/product">

		<xsl:variable name="SellPrice">
			<xsl:if test="prodSaleIsActive = 1"><xsl:value-of select="number(prodSalePrice)" /></xsl:if>
			<xsl:if test="prodSaleIsActive = 0"><xsl:value-of select="number(prodPrice)" /></xsl:if>
		</xsl:variable>
		<tr>
			<td><xsl:value-of select="prodID" /></td>
			<td><xsl:value-of select="prodName" /></td>
			<td><xsl:value-of select="prodPrice" /></td>
			<td><xsl:value-of select="prodSalePrice" /></td>
			<td>
				<xsl:if test="prodSaleIsActive = 1">On Sale</xsl:if>
				<xsl:if test="prodSaleIsActive = 0">-</xsl:if>
			</td>
			<td>
				<xsl:for-each select="categories/category">
					<xsl:value-of select="categoryName" /><br />
				</xsl:for-each>
			</td>
		</tr>
	</xsl:for-each>

</xsl:template>

<xsl:template name="productByCategory" >
	<xsl:for-each select="products/product">
		<xsl:for-each select="categories/category">

		<xsl:variable name="SellPrice">
			<xsl:if test="../../IsOnSale = 1"><xsl:value-of select="number(../../SalePrice)" /></xsl:if>
			<xsl:if test="../../IsOnSale = 0"><xsl:value-of select="number(../../Price)" /></xsl:if>
		</xsl:variable>
		<tr>
			<td><xsl:value-of select="../../Code" /></td>
			<td><xsl:value-of select="../../Name" /></td>
			<td><xsl:value-of select="../../Cost" /></td>
			<td><xsl:value-of select="../../Price" /></td>
			<td><xsl:value-of select="../../SalePrice" /></td>
			<td>
				<xsl:if test="../../IsOnSale = 1">On Sale</xsl:if>
				<xsl:if test="../../IsOnSale = 0">-</xsl:if>
			</td>
			<td><xsl:value-of select="categoryName" /></td>
		</tr>
		</xsl:for-each>
	</xsl:for-each>

</xsl:template>

</xsl:stylesheet>