<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html"/>
<!--
'********************************************************************************
'*   Order Manager - Payment Export Module										*
'*   Release Version:   1.00.0001												*
'*   Release Date:		October 13, 2004										*
'*   Revision Date:		October 13, 2004										*
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

<!-- USER DEFINED SETTINGS -->
	<xsl:variable name="displayHeader" select="'N'"/>
	<xsl:variable name="outputDelimeter"><xsl:text>,</xsl:text></xsl:variable>
	<xsl:variable name="outputDelimeter_CR"><xsl:text>&#10;</xsl:text></xsl:variable>
<!-- USER DEFINED SETTINGS -->


<xsl:template match='/'>
	<xsl:if test="$displayHeader = 'Y'">
		<!-- Begin Header -->
		<xsl:text>Email</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		<!-- End Header -->
	</xsl:if>
	<xsl:apply-templates select="items/item" />
</xsl:template>

<xsl:template match="item" >
	<xsl:if test="string-length(custEmail) > 0">
	<xsl:if test="custIsSubscribed='1'">
		<!-- custEmail --><xsl:value-of select="custEmail"/>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
	</xsl:if>
	</xsl:if>

</xsl:template>

</xsl:stylesheet>