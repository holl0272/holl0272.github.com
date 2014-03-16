<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
      xmlns:msxsl="urn:schemas-microsoft-com:xslt">
<xsl:import href="supportingTemplates/supportTemplates.xsl"/>
<xsl:output method="html"/>
<!--
'********************************************************************************
'*   Order Manager - FedEx Module				                                *
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

This module creates a FedEx compatible file:

-->


<!-- ********* USER DEFINED SETTINGS ********* -->

	<xsl:template match="orderShipMethod">
		<xsl:choose>
			<xsl:when test="../orderShipMethod = 'Regular Shipping'">Ground</xsl:when>
			<xsl:when test="../orderShipMethod = 'Premium Shipping'">2nd Day Air</xsl:when>
			<xsl:when test="../orderShipMethod = 'Free Shipping'">Ground</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS Next Day AM ®'">Next Day Early AM</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS Next Day Air ®'">Next Day Air</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS Next Day Air Saver ®'">Next Day Air Saver</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS 2nd Day Air Early AM ®'">2nd Day Air AM</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS 2nd Day Air ®'">2nd Day Air</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS 3 Day Select ®'">3 Day Select</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS Ground'">Ground</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS Canada Standard'">Canada Standard</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS WorldWide Express (sm)'">WorldWide Express</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS WorldWide Express Plus (sm)'">WorldWide Express Plus</xsl:when>
			<xsl:when test="../orderShipMethod = 'UPS WorldWide Expedited (sm)'">WorldWide Expedited</xsl:when>
			<xsl:otherwise>Ground</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

<!-- ********* USER DEFINED SETTINGS ********* -->

<xsl:variable name="outputDelimeter"><xsl:text>,</xsl:text></xsl:variable>

<xsl:template match='/'>
	<xsl:apply-templates select="orders/order" />
</xsl:template>

<xsl:template match="order" >
		<!-- OrderID -->"<xsl:value-of select="../OrderNumber" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- Name -->"<xsl:value-of select="Company" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- Address1 -->"<xsl:value-of select="Address1" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- Address2 -->"<xsl:value-of select="Address2" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- City -->"<xsl:value-of select="City" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- State -->"<xsl:value-of select="State" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- ZIPCode -->"<xsl:value-of select="Zip" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- Telephone -->"<xsl:value-of select="Phone" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- ShipmentNotificationEmail -->"<xsl:value-of select="EMail" /><xsl:text>"</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>



<!-- OrderID --><xsl:value-of select="orderID" /><xsl:value-of select="$outputDelimeter"/>

<!-- Use shipping address by default -->
<xsl:if test="string-length(shippingAddress/Country) > 0">
	<xsl:if test="string-length(shippingAddress/Company) > 0">
		<!-- Name --><xsl:apply-templates select="shippingAddress/Company" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	</xsl:if>
	<xsl:if test="string-length(shippingAddress/Company) = 0">
		<!-- Name --><xsl:apply-templates select="shippingAddress/FirstName" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="shippingAddress/MiddleInitial" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="shippingAddress/LastName" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	</xsl:if>

	<!-- Address1 --><xsl:apply-templates select="shippingAddress/Addr1" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Address2 --><xsl:apply-templates select="shippingAddress/Addr2" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- City --><xsl:apply-templates select="shippingAddress/City" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- State --><xsl:apply-templates select="shippingAddress/State" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- ZIPCode --><xsl:apply-templates select="shippingAddress/Zip" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Telephone --><xsl:apply-templates select="shippingAddress/Phone" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- ShipmentNotificationEmail --><xsl:apply-templates select="shippingAddress/Email" mode="stripComma" /><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
</xsl:if>

<!-- Use shipping address by default -->
<xsl:if test="string-length(billingAddress/Country) = 0">
	<xsl:if test="string-length(billingAddress/Company) > 0">
		<!-- Name --><xsl:apply-templates select="billingAddress/Company" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	</xsl:if>
	<xsl:if test="string-length(billingAddress/Company) = 0">
		<!-- Name --><xsl:apply-templates select="billingAddress/FirstName" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="billingAddress/MiddleInitial" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="billingAddress/LastName" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	</xsl:if>

	<!-- Address1 --><xsl:apply-templates select="billingAddress/Addr1" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Address2 --><xsl:apply-templates select="billingAddress/Addr2" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- City --><xsl:apply-templates select="billingAddress/City" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- State --><xsl:apply-templates select="billingAddress/State" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- ZIPCode --><xsl:apply-templates select="billingAddress/Zip" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Telephone --><xsl:apply-templates select="billingAddress/Phone" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- ShipmentNotificationEmail --><xsl:apply-templates select="billingAddress/Email" mode="stripComma" /><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
</xsl:if>

<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
</xsl:template>

</xsl:stylesheet>