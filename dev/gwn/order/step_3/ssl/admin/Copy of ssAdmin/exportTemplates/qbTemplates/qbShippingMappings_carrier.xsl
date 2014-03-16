<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<!--
'********************************************************************************
'*   Order Manager QuickBooks Module		                                    *
'*   Release Version:   1.00.0001												*
'*   Release Date:		December 26, 2003										*
'*   Release Date:		December 26, 2003										*
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

<xsl:template match="orderShipMethod">
    <xsl:choose>
        <xsl:when test="../orderShipMethod = 'Regular Shipping'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'Premium Shipping'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'Free Shipping'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Next Day AM ®'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Next Day Air ®'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Next Day Air Saver ®'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 2nd Day Air Early AM ®'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 2nd Day Air ®'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 3 Day Select ®'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Ground'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Canada Standard'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Express (sm)'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Express Plus (sm)'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Expedited (sm)'">UPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'FEDEX Ground'">FEDEX</xsl:when>
        <xsl:when test="../orderShipMethod = 'USPS Parcel'">USPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'USPS Express'">USPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'USPS Priority'">USPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'USPS International'">USPS</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Regular'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Expedited'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Xpresspost'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Priority Courier'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Expedited Evening'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Xpresspost Evening'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Expedited Saturday'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Xpresspost Saturday'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Surface'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Air'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Xpresspost'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Purolator'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Puropak'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - International - Surface'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - International - Air'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - International - Purolator'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - International - Puropak'">Canada Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'DHL'">DHL</xsl:when>
        <xsl:when test="../orderShipMethod = ''">LTL</xsl:when>
        
        <xsl:otherwise>UPS</xsl:otherwise>
    </xsl:choose>
</xsl:template>

</xsl:stylesheet>
