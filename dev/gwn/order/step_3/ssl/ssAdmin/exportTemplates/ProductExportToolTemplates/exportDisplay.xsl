<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:import href="../supportingTemplates/supportTemplates.xsl"/>
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
  <colgroup>
    <col width="15%" align="left" />
    <col width="15%" align="left" />
    <col width="15%" align="center" />
    <col width="15%" align="center" />
    <col width="15%" align="center" />
    <col width="15%" align="center" />
    <col width="15%" align="center" />
  </colgroup>
  <thead>
  <tr>
    <th align="left" class="hdrNonSelected" style="cursor:auto;">Product ID</th>
    <th align="left" class="hdrNonSelected" style="cursor:auto;">Product</th>
    <th class="hdrNonSelected" style="cursor:auto;">Cost</th>
    <th class="hdrNonSelected" style="cursor:auto;">Regular Price</th>
    <th class="hdrNonSelected" style="cursor:auto;">Sale Price</th>
    <th class="hdrNonSelected" style="cursor:auto;">On Sale</th>
    <th class="hdrNonSelected" style="cursor:auto;">Active</th>
  </tr>
  </thead>
	<xsl:apply-templates select="products" />
</table>
</xsl:template>

<xsl:template match="products" >
	<xsl:for-each select="product">
		<xsl:variable name="rowStyle">
			<xsl:if test="position() mod 2 = 1"></xsl:if>
			<xsl:if test="position() mod 2 = 0">Inactive</xsl:if>
		</xsl:variable>
		<xsl:for-each select="categories/category">

			<xsl:variable name="SellPrice">
				<xsl:if test="../../IsOnSale = 1"><xsl:value-of select="number(../../SalePrice)" /></xsl:if>
				<xsl:if test="../../IsOnSale = 0"><xsl:value-of select="number(../../Price)" /></xsl:if>
			</xsl:variable>
			
			<tr>
				<xsl:attribute name="class"><xsl:value-of select="$rowStyle" /></xsl:attribute>
				<td><xsl:value-of select="../../Code" /></td>
				<td><xsl:value-of select="../../prodName" /></td>
				<td><xsl:call-template name="customNumber"><xsl:with-param name="currentNode">11s</xsl:with-param></xsl:call-template>-<xsl:value-of select="../../Cost" /></td>
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