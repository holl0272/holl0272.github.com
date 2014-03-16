<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:import href="supportingTemplates/supportTemplates.xsl"/>
<xsl:output method="text"/>
<!--
'********************************************************************************
'*   Order Manager - USPS Click-N-Ship Module
'*   Release Version:   1.00.001
'*   Release Date:		July 19, 2006
'*   Revision Date:		July 19, 2006
'*
'*   This module creates a UPS Worldship compatible file:
'*
'*   Release Notes:
'*
'*   1.00.001 (July 19, 2006)
'*   - Initial Release
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.
'********************************************************************************
-->

<!-- ********* USER DEFINED SETTINGS ARE FOUND IN qbAccountSettings. The settings below are specific to the Sales Receipt export. -->
	<xsl:variable name="displayHeader" select="'Y'"/>
	<xsl:variable name="outputDelimeter"><xsl:text>,</xsl:text></xsl:variable>
	<xsl:variable name="orderNumberPrefix" select="''"/>
<!-- ********* USER DEFINED SETTINGS -->

<xsl:template name="carrier">
	<xsl:param name="currentNode" />
    <xsl:choose>
        <xsl:when test="$currentNode/orderShipMethod = 'USPS Parcel'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'USPS Express'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'USPS Priority'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'USPS International'">1</xsl:when>
        
		<!-- Begin Postage Rate 2 methods -->
        <xsl:when test="$currentNode/orderShipMethod = 'Express Mail'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Priority Mail'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Parcel Post'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'First Class'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Media Mail'">1</xsl:when>
        
        <xsl:when test="$currentNode/orderShipMethod = 'Global Express Guaranteed Document Service'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Global Express Guaranteed Non-Document Service'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Global Express Mail (EMS)'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Global Priority Mail - Flat-rate Envelope (large)'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Global Priority Mail - Flat-rate Envelope (small)'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Global Priority Mail - Variable Weight Envelope (single)'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Airmail Letter Post'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Airmail Parcel Post'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Economy (Surface) Letter Post'">1</xsl:when>
        <xsl:when test="$currentNode/orderShipMethod = 'Economy (Surface) Parcel Post'">1</xsl:when>

        <xsl:otherwise>0</xsl:otherwise>
    </xsl:choose>
</xsl:template>

<xsl:template match='/'>
	<xsl:if test="$displayHeader = 'Y'">
		<!-- Begin Header -->
		<xsl:text>Full Name</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Company</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Address 1</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Address 2</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Address 3</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>City</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>State</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Zip Code</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Province</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Country</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Urbanization</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Phone Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Fax Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>E Mail</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Reference Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Short Name</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		<!-- End Header -->
	</xsl:if>
	<xsl:apply-templates select="orders/order" />
</xsl:template>

<xsl:template match="order" >
    <xsl:variable name="isUSPS">
		<xsl:call-template name="carrier">
			<xsl:with-param name='currentNode' select='.' />
		</xsl:call-template>
    </xsl:variable>

	<xsl:if test="$isUSPS=1">
		<!-- Full Name -->"<xsl:value-of select="shippingAddress/FirstName" /><xsl:text> </xsl:text><xsl:if test="string-length(shippingAddress/MiddleInitial)>0"><xsl:value-of select="shippingAddress/MiddleInitial" /><xsl:text> </xsl:text></xsl:if><xsl:value-of select="shippingAddress/LastName" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- Company -->"<xsl:apply-templates select="shippingAddress/Company" mode="escape-CSV" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- Address 1 -->"<xsl:apply-templates select="shippingAddress/Addr1" mode="escape-CSV" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- Address 2 -->"<xsl:apply-templates select="shippingAddress/Addr2" mode="escape-CSV" />"<xsl:value-of select="$outputDelimeter"/>
		<!-- Address 3 --><xsl:value-of select="$outputDelimeter"/>
		<!-- City --><xsl:apply-templates select="shippingAddress/City" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- State --><xsl:apply-templates select="shippingAddress/State" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- ZIP Code --><xsl:value-of select="shippingAddress/Zip" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Province --><xsl:value-of select="$outputDelimeter"/>
		<!-- Country --><xsl:apply-templates select="shippingAddress/Country" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Urbanization --><xsl:value-of select="$outputDelimeter"/>
		<!-- Phone Number --><xsl:value-of select="shippingAddress/Phone" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Fax Number --><xsl:value-of select="shippingAddress/Fax" /><xsl:value-of select="$outputDelimeter"/>
		<!-- E Mail --><xsl:apply-templates select="shippingAddress/Email" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Reference Number --><xsl:value-of select="$orderNumberPrefix"/><xsl:value-of select="orderID" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Full Name -->"<xsl:value-of select="shippingAddress/LastName" /><xsl:value-of select="shippingAddress/Addr1" /><xsl:text>"</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
	</xsl:if>

</xsl:template>

<xsl:template match="node()" mode="escape-CSV" name="escape-CSV">
   <xsl:param name="string" select="." />
   <xsl:choose>
      <xsl:when test="contains($string, '&quot;')">
         <xsl:value-of select="substring-before($string, '&quot;')" />
         <xsl:text>""</xsl:text>
         <xsl:call-template name="escape-CSV">
            <xsl:with-param name="string"
                  select="substring-after($string, '&quot;')" />
         </xsl:call-template>
      </xsl:when>
      <xsl:when test="contains($string, ',')">
         <xsl:value-of select="substring-before($string, ',')" />
         <xsl:text>","</xsl:text>
         <xsl:call-template name="escape-CSV">
            <xsl:with-param name="string"
                  select="substring-after($string, ',')" />
         </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
         <xsl:value-of select="$string" />
      </xsl:otherwise>
   </xsl:choose>
</xsl:template>

<xsl:variable name="str" select="'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'" />
<xsl:template name="stripchars">
  <xsl:param name="x" />
  <xsl:param name="y" />
    
  <xsl:if test="contains($str, $x)">
    <xsl:value-of select="$x" />
  </xsl:if>
    
  <xsl:if test="string-length($y) > 0">
     <xsl:call-template name="stripchars">
       <xsl:with-param name="x" select="substring($y, 1, 1)" /> 
       <xsl:with-param name="y" select="substring($y, 2, string-length($y))" /> 
     </xsl:call-template>
  </xsl:if>
</xsl:template>

<xsl:template name="replace-string">
    <xsl:param name="text"/>
    <xsl:param name="replace"/>
    <xsl:param name="with"/>
    <xsl:choose>
      <xsl:when test="contains($text,$replace)">
        <xsl:value-of select="substring-before($text,$replace)"/>
        <xsl:value-of select="$with"/>
        <xsl:call-template name="replace-string">
          <xsl:with-param name="text" select="substring-after($text,$replace)"/>
          <xsl:with-param name="replace" select="$replace"/>
          <xsl:with-param name="with" select="$with"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$text"/>
      </xsl:otherwise>
    </xsl:choose>
</xsl:template>

</xsl:stylesheet>