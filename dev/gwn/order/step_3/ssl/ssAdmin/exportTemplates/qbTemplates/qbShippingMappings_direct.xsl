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
        <xsl:when test="../orderShipMethod = 'Regular Shipping'">Regular Shipping</xsl:when>
        <xsl:when test="../orderShipMethod = 'Premium Shipping'">Premium Shipping</xsl:when>
        <xsl:when test="../orderShipMethod = 'Free Shipping'">Free Shipping</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Next Day AM ®'">UPS Next Day AM ®</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Next Day Air ®'">UPS Next Day Air ®</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Next Day Air Saver ®'">UPS Next Day Air Saver ®</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 2nd Day Air Early AM ®'">UPS 2nd Day Air Early AM ®</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 2nd Day Air ®'">UPS 2nd Day Air ®</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 3 Day Select ®'">UPS 3 Day Select ®</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Ground'">UPS Ground</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Canada Standard'">UPS Canada Standard</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Express (sm)'">UPS WorldWide Express (sm)</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Express Plus (sm)'">UPS WorldWide Express Plus (sm)</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Expedited (sm)'">UPS WorldWide Expedited (sm)</xsl:when>
        <xsl:when test="../orderShipMethod = 'FEDEX Ground'">FEDEX Ground</xsl:when>
        <xsl:when test="../orderShipMethod = 'USPS Parcel'">USPS Parcel</xsl:when>
        <xsl:when test="../orderShipMethod = 'USPS Express'">USPS Express</xsl:when>
        <xsl:when test="../orderShipMethod = 'USPS Priority'">USPS Priority</xsl:when>
        <xsl:when test="../orderShipMethod = 'USPS International'">USPS International</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Regular'">Canada Post - Domestic - Regular</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Expedited'">Canada Post - Domestic - Expedited</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Xpresspost'">Canada Post - Domestic - Xpresspost</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Priority Courier'">Canada Post - Domestic - Priority Courier</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Expedited Evening'">Canada Post - Domestic - Expedited Evening</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Xpresspost Evening'">Canada Post - Domestic - Xpresspost Evening</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Expedited Saturday'">Canada Post - Domestic - Expedited Saturday</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - Domestic - Xpresspost Saturday'">Canada Post - Domestic - Xpresspost Saturday</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Surface'">Canada Post - USA - Surface</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Air'">Canada Post - USA - Air</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Xpresspost'">Canada Post - USA - Xpresspost</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Purolator'">Canada Post - USA - Purolator</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - USA - Puropak'">Canada Post - USA - Puropak</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - International - Surface'">Canada Post - International - Surface</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - International - Air'">Canada Post - International - Air</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - International - Purolator'">Canada Post - International - Purolator</xsl:when>
        <xsl:when test="../orderShipMethod = 'Canada Post - International - Puropak'">Canada Post - International - Puropak</xsl:when>
        <xsl:when test="../orderShipMethod = 'DHL'">DHL</xsl:when>
        <xsl:when test="../orderShipMethod = ''">LTL Carriers</xsl:when>
        
		<!-- Begin Postage Rate 2 methods -->
        <xsl:when test="../orderShipMethod = 'Express Mail'">Express Mail</xsl:when>
        <xsl:when test="../orderShipMethod = 'Priority Mail'">Priority Mail</xsl:when>
        <xsl:when test="../orderShipMethod = 'Parcel Post'">Parcel Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'First Class'">First Class</xsl:when>
        <xsl:when test="../orderShipMethod = 'Media Mail'">Media Mail</xsl:when>
        
       <xsl:when test="../orderShipMethod = 'Global Express Guaranteed Document Service'">Global Express Guaranteed Document Service</xsl:when>
        <xsl:when test="../orderShipMethod = 'Global Express Guaranteed Non-Document Service'">Global Express Guaranteed Non-Document Service</xsl:when>
        <xsl:when test="../orderShipMethod = 'Global Express Mail (EMS)'">Global Express Mail (EMS)</xsl:when>
        <xsl:when test="../orderShipMethod = 'Global Priority Mail - Flat-rate Envelope (large)'">Global Priority Mail - Flat-rate Envelope (large)</xsl:when>
        <xsl:when test="../orderShipMethod = 'Global Priority Mail - Flat-rate Envelope (small)'">Global Priority Mail - Flat-rate Envelope (small)</xsl:when>
        <xsl:when test="../orderShipMethod = 'Global Priority Mail - Variable Weight Envelope (single)'">Global Priority Mail - Variable Weight Envelope (single)</xsl:when>
        <xsl:when test="../orderShipMethod = 'Airmail Letter Post'">Airmail Letter Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Airmail Parcel Post'">Airmail Parcel Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Economy (Surface) Letter Post'">Economy (Surface) Letter Post</xsl:when>
        <xsl:when test="../orderShipMethod = 'Economy (Surface) Parcel Post'">Economy (Surface) Parcel Post</xsl:when>

        <xsl:when test="../orderShipMethod = 'UPS Next Day AM'">UPS Next Day AM</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Next Day Air'">UPS Next Day Air</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Next Day Air Saver'">UPS Next Day Air Saver</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 2nd Day Air AM'">UPS 2nd Day Air AM</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 2nd Day Air'">UPS 2nd Day Air</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS 3 Day Select'">UPS 3 Day Select</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Standard Ground'">UPS Standard Ground</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS Canada Standard'">UPS Canada Standard</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Express'">UPS WorldWide Express</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Express Plus'">UPS WorldWide Express Plus</xsl:when>
        <xsl:when test="../orderShipMethod = 'UPS WorldWide Expedited'">UPS WorldWide Expedited</xsl:when>

        <xsl:when test="../orderShipMethod = 'FedEx Priority'">FedEx Priority</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx 2day'">FedEx 2day</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx Standard Overnight'">FedEx Standard Overnight</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx First Overnight'">FedEx First Overnight</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx Express Saver'">FedEx Express Saver</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx Overnight Freight'">FedEx Overnight Freight</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx 2day Freight'">FedEx 2day Freight</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx Express Saver Freight'">FedEx Express Saver Freight</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx International Priority'">FedEx International Priority</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx International Economy'">FedEx International Economy</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx International First'">FedEx International First</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx International Priority Freight'">FedEx International Priority Freight</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx International Economy Freight'">FedEx International Economy Freight</xsl:when>
        <xsl:when test="../orderShipMethod = 'FedEx Home Delivery'">FedEx Home Delivery</xsl:when>
        <xsl:when test="../orderShipMethod = 'U.S. Domestic FedEx Ground Package'">U.S. Domestic FedEx Ground Package</xsl:when>
        <xsl:when test="../orderShipMethod = 'International FedEx Ground Package'">International FedEx Ground Package</xsl:when>

        <xsl:when test="../orderShipMethod = 'Domestic - Regular'">Domestic - Regular</xsl:when>
        <xsl:when test="../orderShipMethod = 'Domestic - Expedited'">Domestic - Expedited</xsl:when>
        <xsl:when test="../orderShipMethod = 'Domestic - Xpresspost'">Domestic - Xpresspost</xsl:when>
        <xsl:when test="../orderShipMethod = 'Domestic - Priority Courier'">Domestic - Priority Courier</xsl:when>
        <xsl:when test="../orderShipMethod = 'Domestic - Expedited Evening'">Domestic - Expedited Evening</xsl:when>
        <xsl:when test="../orderShipMethod = 'Domestic - Xpresspost Evening'">Domestic - Xpresspost Evening</xsl:when>
        <xsl:when test="../orderShipMethod = 'Domestic - Expedited Saturday'">Domestic - Expedited Saturday</xsl:when>
        <xsl:when test="../orderShipMethod = 'Domestic - Expedited Saturday'">Domestic - Expedited Saturday</xsl:when>
        <xsl:when test="../orderShipMethod = 'Domestic - Xpresspost Saturday'">Domestic - Xpresspost Saturday</xsl:when>
        <xsl:when test="../orderShipMethod = 'USA - Surface'">USA - Surface</xsl:when>
        <xsl:when test="../orderShipMethod = 'USA - Air'">USA - Air</xsl:when>
        <xsl:when test="../orderShipMethod = 'USA - Xpresspost'">USA - Xpresspost</xsl:when>
        <xsl:when test="../orderShipMethod = 'USA - Purolator'">USA - Purolator</xsl:when>
        <xsl:when test="../orderShipMethod = 'USA - Puropak'">USA - Puropak</xsl:when>
        <xsl:when test="../orderShipMethod = 'International - Surface'">International - Surface</xsl:when>
        <xsl:when test="../orderShipMethod = 'International - Air'">International - Air</xsl:when>
        <xsl:when test="../orderShipMethod = 'International - Purolator'">International - Purolator</xsl:when>
        <xsl:when test="../orderShipMethod = 'International - Puropak'">International - Puropak</xsl:when>

        <xsl:otherwise><xsl:value-of select="../orderShipMethod" /></xsl:otherwise>
    </xsl:choose>
</xsl:template>

</xsl:stylesheet>
