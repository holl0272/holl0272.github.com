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

<!-- ********* USER DEFINED SETTINGS ********* -->

	<xsl:variable name="accountsReceivableName" select="'Accounts Receivable'"/>
	<xsl:variable name="bankAccountName" select="'Accounts Receivable'"/>
	<xsl:variable name="paymentsAccountName" select="'Undeposited Funds'"/>
	<xsl:variable name="taxAccountName" select="'State Tax'"/>

	<xsl:variable name="salesAccountName" select="'Sales'"/>
	
	<xsl:variable name="salesTaxName" select="'Sales Tax Payable'"/>

	<xsl:variable name="discountName" select="'Order Discounts'"/>
	<xsl:variable name="discountMemo" select="'Order Discount'"/>
	<xsl:variable name="discountInvItem" select="'Discount'"/>

	<xsl:variable name="giftCertificateName" select="'Gift Certificate Redemption'"/>
	<xsl:variable name="giftCertificateMemo" select="'Gift Certificate Redemption'"/>
	<xsl:variable name="giftCertificateInvItem" select="'Gift Certificate'"/>

	<xsl:variable name="shippingName" select="'Freight Income'"/>
	<xsl:variable name="shippingMemo" select="'Shipping &amp; Handling'"/>
	<xsl:variable name="shippingInvItem" select="'shipping'"/>

	<xsl:variable name="handlingName" select="'Handling Income'"/>
	<xsl:variable name="handlingMemo" select="'Handling'"/>
	<xsl:variable name="handlingInvItem" select="'Handling'"/>

	<xsl:variable name="attributeSeparator_productCode" select="' - '"/>
	<xsl:variable name="attributeSeparator_multipleAttributes" select="', '"/>
	<xsl:variable name="attributeSeparator" select="': '"/>

	<xsl:variable name="attributeSeparator_productCode_INV" select="' '"/>
	<xsl:variable name="attributeSeparator_multipleAttributes_INV" select="' '"/>
	<xsl:variable name="attributeSeparator_INV" select="' '"/>
	
	<xsl:variable name="giftCertificateProductCode" select="'GiftCertificate'"/>

	<xsl:variable name="storeMessage" select="'Thank You for shopping at www.storefront.net'"/>
      
	<xsl:variable name="dateToUse" select="'ssDateOrderShipped'"/><!-- shortOrderDate (default), TodaysDate, ssDatePaymentReceived, ssDateOrderShipped  -->

<!-- ********* END USER DEFINED SETTINGS ********* -->

	<xsl:variable name="outputDelimeter_CR"><xsl:text>&#10;</xsl:text></xsl:variable>
	<xsl:variable name="outputDelimeter_TAB"><xsl:text>&#9;</xsl:text></xsl:variable>
	<xsl:variable name="outputDelimeter"><xsl:text>&#9;</xsl:text></xsl:variable>

<xsl:template match="billingAddress">
	<!--
	Use this one for First Name, MI, Last Name
	<xsl:value-of select="../billingAddress/FirstName" /><xsl:text> </xsl:text><xsl:if test="string-length(../billingAddress/MiddleInitial)>0"><xsl:value-of select="../billingAddress/MiddleInitial" /><xsl:text> </xsl:text></xsl:if><xsl:value-of select="../billingAddress/LastName" />

	Use this one for Last Name, First Name MI
	<xsl:value-of select="../billingAddress/LastName" /><xsl:text>, </xsl:text><xsl:value-of select="../billingAddress/FirstName" /><xsl:if test="string-length(../billingAddress/MiddleInitial)>0"><xsl:text> </xsl:text><xsl:value-of select="../billingAddress/MiddleInitial" /></xsl:if>
	-->
	<xsl:value-of select="../billingAddress/LastName" /><xsl:text>, </xsl:text><xsl:value-of select="../billingAddress/FirstName" /><xsl:if test="string-length(../billingAddress/MiddleInitial)>0"><xsl:text> </xsl:text><xsl:value-of select="../billingAddress/MiddleInitial" /></xsl:if>

</xsl:template>

<xsl:template match="node()" mode="product_INVITEM" name="product_INVITEM">
    <xsl:choose>
        <xsl:when test="odrdtProductID = $giftCertificateProductCode">
        	<xsl:value-of select="odrdtProductID" />
        </xsl:when>
        <xsl:otherwise>
			<xsl:value-of select="odrdtProductID" />
			<xsl:for-each select="odrdtAttDetailID">
				<xsl:if test="position() = 1">
					<xsl:value-of select="$attributeSeparator_productCode_INV" />
				</xsl:if>
				<xsl:if test="position() = 2">
					<xsl:value-of select="$attributeSeparator_multipleAttributes_INV" />
				</xsl:if>
				<xsl:value-of select="odrattrName" />
				<xsl:value-of select="$attributeSeparator_INV" />
				<xsl:value-of select="odrattrAttribute" />
			</xsl:for-each>
        </xsl:otherwise>
    </xsl:choose>
</xsl:template>

</xsl:stylesheet>
