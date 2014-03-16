<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:import href="../supportingTemplates/supportTemplates.xsl"/>
<xsl:output method="text"/>
<!--
'********************************************************************************
'*   Product Export Tool - Froogle Module				                        *
'*   Release Version:   1.00.000												*
'*   Release Date:		April 6, 2005											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2005 Sandshot Software.  All rights reserved.                *
'********************************************************************************

-->
<xsl:variable name="outputDelimeter">,</xsl:variable>

<xsl:template match='/'>
	<xsl:if test="false">
	</xsl:if>
		<xsl:text>DetailLink</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>prodName</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>ShortDescription</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Price</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>ImageSmallPath</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>categoryName</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Code</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
	<xsl:apply-templates select="products/product" />
</xsl:template>

<xsl:template match="product" >
	<xsl:for-each select="categories/category">
		<xsl:call-template name="product.CategoryLoop"></xsl:call-template>
	</xsl:for-each>

</xsl:template>

<xsl:template name="product.CategoryLoop" >
	<xsl:text>"</xsl:text><xsl:call-template name="escape-quote"><xsl:with-param name="string" select="../../DetailLink" /></xsl:call-template><xsl:text>"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>"</xsl:text><xsl:call-template name="escape-quote"><xsl:with-param name="string" select="../../Name" /></xsl:call-template><xsl:text>"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>"</xsl:text><xsl:call-template name="escape-quote"><xsl:with-param name="string" select="../../ShortDescription" /></xsl:call-template><xsl:text>"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>"</xsl:text><xsl:call-template name="escape-quote"><xsl:with-param name="string" select="../../Price" /></xsl:call-template><xsl:text>"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>"</xsl:text><xsl:call-template name="escape-quote"><xsl:with-param name="string" select="../../ImageSmallPath" /></xsl:call-template><xsl:text>"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>"</xsl:text><xsl:call-template name="escape-quote"><xsl:with-param name="string" select="categoryName" /></xsl:call-template><xsl:text>"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>"</xsl:text><xsl:call-template name="escape-quote"><xsl:with-param name="string" select="../../Code" /></xsl:call-template><xsl:text>"</xsl:text>
	<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
</xsl:template>

</xsl:stylesheet>