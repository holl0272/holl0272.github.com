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

<!-- ********* USER DEFINED SETTINGS -->

	<xsl:variable name="accountsReceivableName" select="'Accounts Receivable'"/>
	<xsl:variable name="bankAccountName" select="'Undeposited Funds'"/>
	<xsl:variable name="paymentsAccountName" select="'Undeposited Funds'"/>

	<xsl:variable name="salesAccountName" select="'Sales'"/>
	
	<xsl:variable name="salesTaxName" select="'Sales Tax Payable'"/>

	<xsl:variable name="discountName" select="'Order Discounts'"/>
	<xsl:variable name="discountMemo" select="'Order Discount'"/>
	<xsl:variable name="discountInvItem" select="'Discount'"/>

	<xsl:variable name="shippingName" select="'Freight Income'"/>
	<xsl:variable name="shippingMemo" select="'Shipping &amp; Handling'"/>
	<xsl:variable name="shippingInvItem" select="'shipping'"/>

	<xsl:variable name="handlingName" select="'Handling Income'"/>
	<xsl:variable name="handlingMemo" select="'Handling'"/>
	<xsl:variable name="handlingInvItem" select="'Handling'"/>

	<xsl:variable name="attributeSeparator" select="' - '"/>
	<xsl:variable name="storeMessage" select="'Thank You for shopping at www.storefront.net'"/>
      
<!-- ********* USER DEFINED SETTINGS -->

</xsl:stylesheet>
