<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
      xmlns:msxsl="urn:schemas-microsoft-com:xslt">
<xsl:import href="supportingTemplates/supportTemplates.xsl"/>
<xsl:output method="html"/>
<!--
'********************************************************************************
'*   Order Manager - UPS Worldship Module		                                *
'*   Release Version:   2.00.002												*
'*   Release Date:		November 15, 2003										*
'*   Revision Date:		December 20, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

This module creates a UPS Worldship compatible file:

-->


<!-- ********* USER DEFINED SETTINGS ********* -->

    <xsl:variable name="ShipmentNotificationOption" select="'Y'"/>
    <xsl:variable name="ShipmentNotificationType" select="'email'"/>
    <xsl:variable name="DeclaredValueOption" select="'Y'"/>
	<xsl:variable name="BillOption"><xsl:text>Prepaid</xsl:text></xsl:variable>
	<xsl:variable name="QvnOption"><xsl:text>Y</xsl:text></xsl:variable>
	<xsl:variable name="packtype"><xsl:text>Package</xsl:text></xsl:variable>
	<xsl:variable name="dept"><xsl:text></xsl:text></xsl:variable>
      
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

	<xsl:text>ShipTo_CustomerID</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipmentInformation_ServiceType</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_Name</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_Attention</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_Address1</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_Address2</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_Address3</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_Country</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_ZIPCode</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_City</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_State</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_Telephone</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipTo_Fax</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipmentInformation_QVN_Option</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipmentInformation_QVN_Ship_Notification_1_Option</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipmentInformation_QVN_Ship_Recipient_1_Type</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ShipmentInformation_QVN_Ship_Recipient_1_Email</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>Package_Weight</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>Package_DeclaredValueOption</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>Package_DeclaredValueAmount</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>Package_BillingOption</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>Package_PackageType</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>Package_Qvnoption</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>dept</xsl:text><xsl:value-of select="$outputDelimeter_CR"/>

	<xsl:apply-templates select="orders/order" />
</xsl:template>

<xsl:template match="order" >
<!-- OrderID --><xsl:value-of select="orderID" /><xsl:value-of select="$outputDelimeter"/>
<!-- ServiceType --><xsl:apply-templates select="orderShipMethod"/><xsl:value-of select="$outputDelimeter"/>

<!-- Use shipping address by default -->
<xsl:if test="string-length(shippingAddress/Country) > 0">
	<xsl:if test="string-length(shippingAddress/Company) > 0">
		<!-- Name --><xsl:apply-templates select="shippingAddress/Company" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Attention --><xsl:apply-templates select="shippingAddress/FirstName" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="shippingAddress/MiddleInitial" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="shippingAddress/LastName" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	</xsl:if>
	<xsl:if test="string-length(shippingAddress/Company) = 0">
		<!-- Name --><xsl:apply-templates select="shippingAddress/FirstName" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="shippingAddress/MiddleInitial" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="shippingAddress/LastName" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Attention --><xsl:apply-templates select="shippingAddress/Company" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	</xsl:if>

	<!-- Address1 --><xsl:apply-templates select="shippingAddress/Addr1" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Address2 --><xsl:apply-templates select="shippingAddress/Addr2" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Address3 --><xsl:value-of select="$outputDelimeter"/>
	<!-- Country --><xsl:apply-templates select="shippingAddress/Country" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- ZIPCode --><xsl:apply-templates select="shippingAddress/Zip" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- City --><xsl:apply-templates select="shippingAddress/City" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- State --><xsl:apply-templates select="shippingAddress/State" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Telephone --><xsl:apply-templates select="shippingAddress/Phone" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Fax --><xsl:apply-templates select="shippingAddress/Fax" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- QVN_Option --><xsl:value-of select="$ShipmentNotificationOption"/><xsl:value-of select="$outputDelimeter"/>
	<!-- QVN_Ship_Notification_1_Option --><xsl:value-of select="$ShipmentNotificationOption"/><xsl:value-of select="$outputDelimeter"/>
	<!-- QVN_Ship_Recipient_1_Type --><xsl:value-of select="$ShipmentNotificationType"/><xsl:value-of select="$outputDelimeter"/>
	<!-- QVN_Ship_Recipient_1_Email --><xsl:apply-templates select="shippingAddress/Email" mode="stripComma" /><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
</xsl:if>

<!-- Use shipping address by default -->
<xsl:if test="string-length(billingAddress/Country) = 0">
	<xsl:if test="string-length(billingAddress/Company) > 0">
		<!-- Name --><xsl:apply-templates select="billingAddress/Company" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Attention --><xsl:apply-templates select="billingAddress/FirstName" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="billingAddress/MiddleInitial" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="billingAddress/LastName" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	</xsl:if>
	<xsl:if test="string-length(billingAddress/Company) = 0">
		<!-- Name --><xsl:apply-templates select="billingAddress/FirstName" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="billingAddress/MiddleInitial" mode="stripComma" /><xsl:text> </xsl:text><xsl:apply-templates select="billingAddress/LastName" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Attention --><xsl:apply-templates select="billingAddress/Company" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	</xsl:if>

	<!-- Address1 --><xsl:apply-templates select="billingAddress/Addr1" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Address2 --><xsl:apply-templates select="billingAddress/Addr2" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Address3 --><xsl:value-of select="$outputDelimeter"/>
	<!-- Country --><xsl:apply-templates select="billingAddress/Country" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- ZIPCode --><xsl:apply-templates select="billingAddress/Zip" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- City --><xsl:apply-templates select="billingAddress/City" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- State --><xsl:apply-templates select="billingAddress/State" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Telephone --><xsl:apply-templates select="billingAddress/Phone" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- Fax --><xsl:apply-templates select="billingAddress/Fax" mode="stripComma" /><xsl:value-of select="$outputDelimeter"/>
	<!-- QVN_Option --><xsl:value-of select="$ShipmentNotificationOption"/><xsl:value-of select="$outputDelimeter"/>
	<!-- QVN_Ship_Notification_1_Option --><xsl:value-of select="$ShipmentNotificationOption"/><xsl:value-of select="$outputDelimeter"/>
	<!-- QVN_Ship_Recipient_1_Type --><xsl:value-of select="$ShipmentNotificationType"/><xsl:value-of select="$outputDelimeter"/>
	<!-- QVN_Ship_Recipient_1_Email --><xsl:apply-templates select="billingAddress/Email" mode="stripComma" /><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
</xsl:if>

<!-- Weight --><xsl:value-of select="OrderWeight" /><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
<!-- DeclaredValueOption --><xsl:value-of select="$DeclaredValueOption"/><xsl:value-of select="$outputDelimeter"/>
<!-- DeclaredValueAmount --><xsl:value-of select="format-number(number(orderGrandTotal),'###0.00')" />
<!-- Bill Option --><xsl:value-of select="$BillOption"/><xsl:value-of select="$outputDelimeter"/>
<!-- Package Type --><xsl:value-of select="$packtype"/><xsl:value-of select="$outputDelimeter"/>
<!-- Qvn Option --><xsl:value-of select="$QvnOption"/><xsl:value-of select="$outputDelimeter"/>
<!-- dept --><xsl:value-of select="$dept"/><xsl:value-of select="$outputDelimeter"/>
<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
</xsl:template>

</xsl:stylesheet>