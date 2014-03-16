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

First Name - Required.
Last Name - Required.
Street Address - Required.
City - Required.
State - Required. Two-character ISO state code.
Zip Code - Required.
Country - Two-character ISO country code. Defaults to US if left empty.
Phone Number - Enter without spaces or punctuation, for example, 3451231234.
Email Address -
Order Description -
Invoice Number -
Order Amount - Required. Do not include the dollar ($) sign.
Approval Code - Only used for voice authorized transactions. Leave empty if none.
Name on Card - Required.
Card Type - Required. "Visa", "MC", "Disc", or "AmEx"
Card Number - Required.
Expiration Month - Required. Numeric month, with January as 1. Leading zero is optional.
Expiration Year - Required. A four-digit year.
CVV2 Number - See CVV2 Documentation.

-->

<!-- ********* USER DEFINED SETTINGS ARE FOUND IN qbAccountSettings. The settings below are specific to the Sales Receipt export. -->
      <xsl:variable name="displayHeader" select="'N'"/>
      <xsl:variable name="orderNumberPrefix" select="'Order '"/>
<!-- ********* USER DEFINED SETTINGS -->

<xsl:variable name="outputDelimeter"><xsl:text>,</xsl:text></xsl:variable>
<xsl:variable name="outputDelimeter_CR"><xsl:text>&#10;</xsl:text></xsl:variable>

<xsl:template match='/'>
	<xsl:if test="$displayHeader = 'Y'">
		<!-- Begin Header -->
		<xsl:text>First Name</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Last Name</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Street Address</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>City</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>State</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Zip Code</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Country</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Phone Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Email Address</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Order Description</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Invoice Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Order Amount</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Approval Code</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Name on Card</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Card Type</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Card Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Expiration Month</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Expiration Year</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>CVV2 Number</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		<!-- End Header -->
	</xsl:if>
	<xsl:apply-templates select="orders/order" />
</xsl:template>

<xsl:template match="order" >
	<xsl:if test="string-length(payCardType) > 0">
		<!-- FirstName --><xsl:apply-templates select="billingAddress/FirstName" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- LastName --><xsl:apply-templates select="billingAddress/LastName" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Address1 --><xsl:apply-templates select="billingAddress/Addr1" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- City --><xsl:apply-templates select="billingAddress/City" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- State --><xsl:apply-templates select="billingAddress/State" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- ZIP Code -->
			<xsl:variable name="ZIPCode">
				<xsl:call-template name="stripchars">
					<xsl:with-param name="x" select="substring(billingAddress/Zip, 1, 1)" />
					<xsl:with-param name="y" select="substring(billingAddress/Zip, 2, string-length(billingAddress/Phone))" />
				</xsl:call-template>
			</xsl:variable>
		<!-- ZIP Code --><xsl:value-of select="substring($ZIPCode, 1, 10)" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Country --><xsl:apply-templates select="billingAddress/Country" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Telephone -->
			<xsl:variable name="telephoneNumber">
				<xsl:call-template name="stripchars">
					<xsl:with-param name="x" select="substring(billingAddress/Phone, 1, 1)" />
					<xsl:with-param name="y" select="substring(billingAddress/Phone, 2, string-length(billingAddress/Phone))" />
				</xsl:call-template>
			</xsl:variable>
		<!-- Telephone --><xsl:value-of select="substring($telephoneNumber, 1, 10)" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Email Address --><xsl:apply-templates select="billingAddress/Email" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Order Description --><xsl:value-of select="$orderNumberPrefix"/><xsl:value-of select="orderID" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Invoice Number --><xsl:value-of select="orderID" /><xsl:value-of select="$outputDelimeter"/>
		<!-- orderGrandTotal --><xsl:value-of select="orderGrandTotal" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Approval Code --><xsl:value-of select="$outputDelimeter"/>
		<!-- Name on Card --><xsl:value-of select="payCardName" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Card Type --><xsl:apply-templates select="payCardType" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Card Number --><xsl:value-of select="payCardNumber" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Expiration Month --><xsl:value-of select="substring(payCardExpires, 1, 2)" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Expiration Year --><xsl:value-of select="substring(payCardExpires, 4, 4)" /><xsl:value-of select="$outputDelimeter"/>
		<!-- CVV2 Number --><xsl:value-of select="payCardCCV" />
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
	</xsl:if>

</xsl:template>

<xsl:template match="payCardType">
    <xsl:choose>
        <xsl:when test=". = 'American Express'">AmEx</xsl:when>
        <xsl:when test=". = 'Visa'">Visa</xsl:when>
        <xsl:when test=". = 'Discover'">Disc</xsl:when>
        <xsl:when test=". = 'MasterCard'">MC</xsl:when>
        <xsl:when test=". = 'Diners Club'">DC</xsl:when>
        <xsl:when test=". = 'Carte Blanche'">CB</xsl:when>
        <xsl:when test=". = 'Delta'">Delta</xsl:when>
        <xsl:when test=". = 'JCB'">JCB</xsl:when>
        <xsl:when test=". = 'Solo'">Solo</xsl:when>
        <xsl:when test=". = 'Switch'">Switch</xsl:when>
        <xsl:otherwise>Unknown</xsl:otherwise>
    </xsl:choose>
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