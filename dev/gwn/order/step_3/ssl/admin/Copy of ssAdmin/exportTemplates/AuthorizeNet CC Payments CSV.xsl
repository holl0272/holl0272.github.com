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
      <xsl:variable name="displayTestColumns" select="'N'"/>
      <xsl:variable name="orderNumberPrefix" select="'Marine Supplies and Electronics Order '"/>
      <xsl:variable name="transactionType" select="'AUTH_ONLY'"/><!-- AUTH_CAPTURE, AUTH_ONLY, CAPTURE_ONLY, CREDIT, VOID, PRIOR_AUTH_CAPTURE -->
<!-- ********* USER DEFINED SETTINGS -->

<xsl:variable name="outputDelimeter"><xsl:text>,</xsl:text></xsl:variable>
<xsl:variable name="outputDelimeter_CR"><xsl:text>&#10;</xsl:text></xsl:variable>

<xsl:template match='/'>
	<xsl:if test="$displayHeader = 'Y'">
		<!-- Begin Header -->
		<xsl:text>Invoice Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Order Description</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Order Amount</xsl:text><xsl:value-of select="$outputDelimeter"/>
<xsl:if test="$displayTestColumns = 'Y'">
	<xsl:text>Authorized Amount</xsl:text><xsl:value-of select="$outputDelimeter"/>
</xsl:if>
		<xsl:text>Card Type</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Transaction Type</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Authorization Code</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Transaction ID</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Card Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Expiration Date</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Bank Account Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Bank Account Type</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Bank ABA Routing Code</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Bank Name</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Customer ID</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>First Name</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Last Name</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Company</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Street Address</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>City</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>State</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Zip Code</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Country</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Phone Number</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>FAX</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Email Address</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Card Code</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		<!-- End Header -->
	</xsl:if>
	<xsl:apply-templates select="orders/order" />
</xsl:template>

<xsl:template match="order" >
	<xsl:variable name="certificateAmount"><xsl:value-of select="number(ssGiftCertificate/RedemptionAmount)" /></xsl:variable>
	<xsl:variable name="amountDue">
		<xsl:choose>
			<xsl:when test="$certificateAmount='NaN'"><xsl:value-of select="number(orderGrandTotal)-number(orderDiscount)-number(orderCouponDiscount)" /></xsl:when>
			<xsl:otherwise><xsl:value-of select="number(orderGrandTotal) + number($certificateAmount)-number(orderDiscount)-number(orderCouponDiscount)" /></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>

	<xsl:if test="(string-length(payCardType) > 0) or string-length(orderCheckAcctNumber) > 0 or string-length(orderPurchaseOrderNumber) > 0">
		<!-- Invoice Number --><xsl:value-of select="orderID" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Order Description --><xsl:value-of select="$orderNumberPrefix"/><xsl:value-of select="orderID" /><xsl:value-of select="$outputDelimeter"/>
		<!-- orderGrandTotal -->
			<xsl:choose>
				<xsl:when test="orderGrandTotal>=0"><xsl:value-of select="format-number($amountDue,'###0.00')" /></xsl:when>
				<xsl:otherwise><xsl:value-of select="format-number(-1*number(orderGrandTotal),'###0.00')" /></xsl:otherwise>
			</xsl:choose><xsl:value-of select="$outputDelimeter"/>
<xsl:if test="$displayTestColumns = 'Y'">
	<xsl:value-of select="format-number(number(ProcessorResponse/trnsrspAuthorizationAmount),'###0.00')" /><xsl:value-of select="$outputDelimeter"/>
</xsl:if>
		<!-- Card Type --><xsl:apply-templates select="payCardType" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Transaction Type -->
			<xsl:if test="(string-length(orderCheckAcctNumber) = 0)">
				<xsl:choose>
					<xsl:when test="orderGrandTotal>0">
						<xsl:choose>
							<xsl:when test="number(orderGrandTotal)&gt;number(ProcessorResponse/trnsrspAuthorizationAmount)">AUTH_CAPTURE</xsl:when>
							<xsl:otherwise>
							<xsl:if test="(authorizationExpired = 'True')">AUTH_CAPTURE</xsl:if>
							<xsl:if test="(authorizationExpired = 'False')">PRIOR_AUTH_CAPTURE</xsl:if>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="orderGrandTotal&lt;0">Credit</xsl:when>
				</xsl:choose>
			</xsl:if>
			<xsl:value-of select="$outputDelimeter"/>
		<!-- Authorization Code --><xsl:value-of select="ProcessorResponse/trnsrspAuthNo" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Transaction ID --><xsl:value-of select="ProcessorResponse/trnsrspRetrievalCode" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Card Number --><xsl:value-of select="payCardNumber" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Expiration Date --><xsl:value-of select="payCardExpires" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Bank Account Number --><xsl:apply-templates select="orderCheckAcctNumber" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Bank Account Type --><xsl:value-of select="$outputDelimeter"/>
		<!-- Bank ABA Routing Code --><xsl:apply-templates select="orderRoutingNumber" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Bank Name --><xsl:apply-templates select="orderBankName" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Customer ID --><xsl:value-of select="billingAddress/custID" /><xsl:value-of select="$outputDelimeter"/>
		<!-- FirstName --><xsl:apply-templates select="billingAddress/FirstName" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- LastName --><xsl:apply-templates select="billingAddress/LastName" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Company --><xsl:apply-templates select="billingAddress/Company" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
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
			<xsl:value-of select="substring($telephoneNumber, 1, 10)" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Fax -->
			<xsl:variable name="faxNumber">
				<xsl:call-template name="stripchars">
					<xsl:with-param name="x" select="substring(billingAddress/Fax, 1, 1)" />
					<xsl:with-param name="y" select="substring(billingAddress/Fax, 2, string-length(billingAddress/Fax))" />
				</xsl:call-template>
			</xsl:variable>
			<xsl:value-of select="substring($faxNumber, 1, 10)" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Email Address --><xsl:apply-templates select="billingAddress/Email_DONOTUSE" mode="escape-CSV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- Card Code --><xsl:value-of select="payCVV" /><xsl:value-of select="$outputDelimeter"/>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
	</xsl:if>

</xsl:template>

<xsl:template match="payCardType">
    <xsl:choose>
        <xsl:when test=". = 'American Express'">CC</xsl:when>
        <xsl:when test=". = 'Visa'">CC</xsl:when>
        <xsl:when test=". = 'Discover'">CC</xsl:when>
        <xsl:when test=". = 'MasterCard'">CC</xsl:when>
        <xsl:when test=". = 'Diners Club'">CC</xsl:when>
        <xsl:when test=". = 'Carte Blanche'">CC</xsl:when>
        <xsl:when test=". = 'Delta'">CC</xsl:when>
        <xsl:when test=". = 'JCB'">CC</xsl:when>
        <xsl:when test=". = 'Solo'">CC</xsl:when>
        <xsl:when test=". = 'Switch'">CC</xsl:when>
        <xsl:otherwise>echeck</xsl:otherwise>
    </xsl:choose>
</xsl:template>

<xsl:template match="payCardType_old">
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