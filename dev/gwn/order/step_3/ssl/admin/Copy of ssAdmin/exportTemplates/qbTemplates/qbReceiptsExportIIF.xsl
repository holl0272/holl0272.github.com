<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
      xmlns:msxsl="urn:schemas-microsoft-com:xslt">
<!--
'********************************************************************************
'*   Order Manager QuickBooks Module		                                    *
'*   Release Version:   1.00.0002												*
'*   Release Date:		December 26, 2003										*
'*   Revision Date:		April 5, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   1.00.002 (April 5, 2004)                                                   *
'*   - Updated export module for Gift Certificate/Discounts                     *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

This module creates a QuickBooks 2002 Pro compatible information file in the format:

!TRNS,TRNSID,TRNSTYPE,DATE,ACCNT,NAME,CLASS,AMOUNT,DOCNUM,MEMO,CLEAR,TOPRINT,TOSEND,ADDR1,ADDR2,ADDR3,ADDR4,ADDR5,DUEDATE,TERMS,PAID,PAYMETH,SHIPVIA,SHIPDATE,REP,FOB,PONUM,INVTITLE,INVMEMO,SADDR1,SADDR2,SADDR3,SADDR4,SADDR5,NAMEISTAXABLE
!SPL,SPLID,TRNSTYPE,DATE,ACCNT,NAME,CLASS,AMOUNT,DOCNUM,MEMO,CLEAR,PRICE,QNTY,INVITEM,PAYMETH,TAXABLE,VALADJ,EXTRA
!ENDTRNS

-->

<xsl:import href="qbAccountSettings.xsl"/>
<xsl:import href="qbShippingMappings.xsl"/>
<xsl:import href="../supportingTemplates/supportTemplates.xsl"/>

<!-- ********* USER DEFINED SETTINGS ARE FOUND IN qbAccountSettings. The settings below are specific to the Sales Receipt export. -->
      <xsl:variable name="outputType" select="'SALE'"/><!-- I've seen this "CASH SALE" AND "SALE" -->
      <xsl:variable name="printInvoice" select="'Y'"/>
<!-- ********* USER DEFINED SETTINGS -->

<xsl:template match='/'>!TRNS,TRNSID,TRNSTYPE,DATE,ACCNT,NAME,CLASS,AMOUNT,DOCNUM,MEMO,CLEAR,TOPRINT,TOSEND,ADDR1,ADDR2,ADDR3,ADDR4,ADDR5,DUEDATE,TERMS,PAID,PAYMETH,SHIPVIA,SHIPDATE,REP,FOB,PONUM,INVTITLE,INVMEMO,SADDR1,SADDR2,SADDR3,SADDR4,SADDR5,NAMEISTAXABLE
!SPL,SPLID,TRNSTYPE,DATE,ACCNT,NAME,CLASS,AMOUNT,DOCNUM,MEMO,CLEAR,PRICE,QNTY,INVITEM,PAYMETH,TAXABLE,VALADJ,EXTRA
!ENDTRNS<xsl:text></xsl:text>
<xsl:apply-templates select="orders/order" />
</xsl:template>

<xsl:template match="order" >

<xsl:variable name="ssGiftCertificateTotal" select="sum(ssGiftCertificate/RedemptionAmount)"/>
<xsl:variable name="orderDiscountTotal" select="number(orderDiscount)"/>
<xsl:variable name="orderTotal" select="number(orderGrandTotal) - number($orderDiscountTotal) + number($ssGiftCertificateTotal)"/>
<!--
SE
<xsl:variable name="orderTotal" select="number(orderGrandTotal) + number($ssGiftCertificateTotal)"/>

AE
<xsl:variable name="orderTotal" select="number(orderGrandTotal) - number($orderDiscountTotal) + number($ssGiftCertificateTotal)"/>
-->

<xsl:variable name="dateToExport">
   <xsl:choose>
      <xsl:when test="$dateToUse='TodaysDate'"><xsl:value-of select="TodaysDate"/></xsl:when>
      <xsl:when test="$dateToUse='ssDatePaymentReceived'"><xsl:value-of select="ssDatePaymentReceived"/></xsl:when>
      <xsl:when test="$dateToUse='ssDateOrderShipped'"><xsl:value-of select="ssDateOrderShipped"/></xsl:when>
      <xsl:otherwise><xsl:value-of select="shortOrderDate"/></xsl:otherwise>
   </xsl:choose>
</xsl:variable>

<xsl:text></xsl:text>
TRNS,<!-- TRNSID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT -->"<xsl:value-of select="$bankAccountName"/>"<xsl:text>,</xsl:text>
<!-- NAME -->"<xsl:apply-templates select="billingAddress"/>"<xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT -->"<xsl:value-of select="format-number(number($orderTotal),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- DOCNUM -->"<xsl:value-of select="InvoiceNumber"/>"<xsl:text>,</xsl:text>
<!-- MEMO --><xsl:text>,</xsl:text>
<!-- CLEAR --><xsl:text>,</xsl:text>
<!-- TOPRINT -->"<xsl:value-of select="$printInvoice"/>"<xsl:text>,</xsl:text>
<!-- TOSEND -->"N"<xsl:text>,</xsl:text>
<!-- BADDR1 -->"<xsl:value-of select="billingAddress/FirstName" /><xsl:text> </xsl:text><xsl:if test="string-length(billingAddress/MiddleInitial)>0"><xsl:value-of select="billingAddress/MiddleInitial" /><xsl:text> </xsl:text></xsl:if><xsl:value-of select="billingAddress/LastName" />"<xsl:text>,</xsl:text>
<!-- BADDR2 --><xsl:if test="string-length(billingAddress/Company)>0">"<xsl:value-of select="billingAddress/Company" />",</xsl:if>
<!-- BADDR3 -->"<xsl:value-of select="billingAddress/Addr1" />"<xsl:text>,</xsl:text>
<!-- BADDR4 --><xsl:if test="string-length(billingAddress/Addr2)>0">"<xsl:value-of select="billingAddress/Addr2" />",</xsl:if>
<!-- BADDR5 -->"<xsl:value-of select="billingAddress/City" /><xsl:text>,</xsl:text><xsl:text> </xsl:text><xsl:value-of select="billingAddress/State" /><xsl:text> </xsl:text><xsl:value-of select="billingAddress/Zip" /><xsl:if test="(string-length(billingAddress/Company)>0 and string-length(billingAddress/Addr2)>0)"><xsl:if test="billingAddress/Country!='US'"><xsl:text> </xsl:text><xsl:value-of select="billingAddress/CountryName" /></xsl:if></xsl:if>"<xsl:text>,</xsl:text>
<!-- BADDR_Alt2 --><xsl:if test="(string-length(billingAddress/Company)=0 or string-length(billingAddress/Addr2)=0)">"<xsl:if test="billingAddress/Country!='US'"><xsl:value-of select="billingAddress/CountryName" /></xsl:if>",</xsl:if>
<!-- BADDR_Alt4 --><xsl:if test="(string-length(billingAddress/Company)=0 and string-length(billingAddress/Addr2)=0)">"",</xsl:if>
<!-- DUEDATE -->"<xsl:value-of select="$dateToExport"/>"<xsl:text>,</xsl:text>
<!-- TERMS --><xsl:if test="string-length(payCardType)>0"><xsl:value-of select="payCardType"/></xsl:if><xsl:if test="string-length(payCardType)=0"><xsl:value-of select="orderPaymentMethod"/></xsl:if><xsl:text>,</xsl:text>
<!-- PAID --><xsl:text>,</xsl:text>
<!-- PAYMETH -->"<xsl:value-of select="payCardType" />"<xsl:text>,</xsl:text>
<!-- SHIPVIA -->"<xsl:apply-templates select="orderShipMethod"/>"<xsl:text>,</xsl:text>
<!-- SHIPDATE -->"<xsl:value-of select="ssDateOrderShipped" />"<xsl:text>,</xsl:text>
<!-- REP --><xsl:text>,</xsl:text>
<!-- FOB --><xsl:text>,</xsl:text>
<!-- PONUM -->"<xsl:value-of select="InvoiceNumber"/>"<xsl:text>,</xsl:text><!-- PO number used for invoices only, Invoice number used for sales receipts -->
<!-- INVTITLE --><xsl:text>,</xsl:text>
<!-- INVMEMO -->"<xsl:value-of select="$storeMessage"/>"<xsl:text>,</xsl:text>
<!-- SADDR1 -->"<xsl:value-of select="shippingAddress/FirstName" /><xsl:text> </xsl:text><xsl:if test="string-length(shippingAddress/MiddleInitial)>0"><xsl:value-of select="shippingAddress/MiddleInitial" /><xsl:text> </xsl:text></xsl:if><xsl:value-of select="shippingAddress/LastName" />"<xsl:text>,</xsl:text>
<!-- SADDR2 --><xsl:if test="string-length(shippingAddress/Company)>0">"<xsl:value-of select="shippingAddress/Company" />",</xsl:if>
<!-- SADDR3 -->"<xsl:value-of select="shippingAddress/Addr1" />"<xsl:text>,</xsl:text>
<!-- SADDR4 --><xsl:if test="string-length(shippingAddress/Addr2)>0">"<xsl:value-of select="shippingAddress/Addr2" />",</xsl:if>
<!-- SADDR5 -->"<xsl:value-of select="shippingAddress/City" /><xsl:text>,</xsl:text><xsl:text> </xsl:text><xsl:value-of select="shippingAddress/State" /><xsl:text> </xsl:text><xsl:value-of select="shippingAddress/Zip" /><xsl:if test="(string-length(shippingAddress/Company)>0 and string-length(shippingAddress/Addr2)>0)"><xsl:if test="shippingAddress/Country!='US'"><xsl:text> </xsl:text><xsl:value-of select="shippingAddress/CountryName" /></xsl:if></xsl:if>"<xsl:text>,</xsl:text>
<!-- SADDR_Alt2 --><xsl:if test="(string-length(shippingAddress/Company)=0 or string-length(shippingAddress/Addr2)=0)">"<xsl:if test="shippingAddress/Country!='US'"><xsl:value-of select="shippingAddress/CountryName" /></xsl:if>",</xsl:if>
<!-- SADDR_Alt4 --><xsl:if test="(string-length(shippingAddress/Company)=0 and string-length(shippingAddress/Addr2)=0)">"",</xsl:if>
<!-- NAMEISTAXABLE -->"<xsl:if test="orderSTax>0">Y</xsl:if><xsl:if test="orderSTax=0">N</xsl:if>"<xsl:text></xsl:text>

<!-- Now for the order details -->
<xsl:for-each select="orderDetail">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT -->"<xsl:value-of select="$salesAccountName"/>"<xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT -->"<xsl:value-of select="format-number(-1*number(odrdtSubTotal),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->
	<xsl:value-of select="$outputDelimeter_Quote"/>
	<xsl:variable name="productName">
		<xsl:value-of select="odrdtProductName" />
		<xsl:for-each select="odrdtAttDetailID">
			<xsl:if test="position() = 1"><xsl:value-of select="$attributeSeparator_productCode" /></xsl:if>
			<xsl:if test="position() = 2"><xsl:value-of select="$attributeSeparator_multipleAttributes" /></xsl:if>
			<xsl:value-of select="odrattrName" /><xsl:value-of select="$attributeSeparator" /><xsl:value-of select="odrattrAttribute" />
		</xsl:for-each>
	</xsl:variable>
	<xsl:apply-templates select="msxsl:node-set($productName)" mode="escape-quote" />
	<xsl:value-of select="$outputDelimeter_Quote"/><xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE -->"<xsl:value-of select="format-number(number(odrdtPrice),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- QNTY -->"-<xsl:value-of select="odrdtQuantity" />"<xsl:text>,</xsl:text>
<!-- INVITEM -->"<xsl:apply-templates select="." mode="product_INVITEM" />"<xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"<xsl:value-of select="prodStateTaxIsActive" />"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->

</xsl:for-each>

<!-- Now for an empty line -->
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO --><xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->

<!-- Now for the discount -->
<xsl:if test="orderCouponDiscount[.&gt;0]">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT -->"<xsl:value-of select="$discountName"/>"<xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT -->"<xsl:value-of select="format-number(number(orderCouponDiscount),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"<xsl:value-of select="$discountMemo"/>"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE -->"-<xsl:value-of select="format-number(number(orderCouponDiscount),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- QNTY -->"-1"<xsl:text>,</xsl:text>
<!-- INVITEM -->"<xsl:value-of select="$discountInvItem"/>"<xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"<xsl:if test="orderSTax>0">Y</xsl:if><xsl:if test="orderSTax=0">N</xsl:if>"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the discount (SE version) -->
<xsl:if test="orderDiscount[.&gt;0]">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT -->"<xsl:value-of select="$discountName"/>"<xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT -->"<xsl:value-of select="format-number(number(orderDiscount),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"<xsl:value-of select="$discountMemo"/>"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE -->"-<xsl:value-of select="format-number(number(orderDiscount),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- QNTY -->"-1"<xsl:text>,</xsl:text>
<!-- INVITEM -->"<xsl:value-of select="$discountInvItem"/>"<xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"<xsl:if test="orderSTax>0">Y</xsl:if><xsl:if test="orderSTax=0">N</xsl:if>"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the shipping -->
<xsl:if test="orderShippingAmount[.&gt;0]">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT -->"<xsl:value-of select="$shippingName"/>"<xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT -->"-<xsl:value-of select="format-number(number(orderShippingAmount),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"<xsl:if test="string-length(orderShipMethod)>0"><xsl:value-of select="orderShipMethod"/></xsl:if><xsl:if test="string-length(orderShipMethod)=0"><xsl:value-of select="$shippingMemo"/></xsl:if>"<xsl:text>,</xsl:text><!-- Alternate usage: select="$shippingMemo" -->
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE -->"<xsl:value-of select="format-number(number(orderShippingAmount),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- QNTY -->"-1"<xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:value-of select="$shippingInvItem"/><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the handling -->
<xsl:if test="orderHandling[.&gt;0]">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT -->"<xsl:value-of select="$handlingName"/>"<xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT -->"-<xsl:value-of select="format-number(number(orderHandling),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"<xsl:value-of select="$handlingMemo"/>"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE -->"<xsl:value-of select="format-number(number(orderHandling),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- QNTY -->"1"<xsl:text>,</xsl:text>
<!-- INVITEM -->"<xsl:value-of select="$handlingInvItem"/>"<xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for an empty line -->
SPL,,<xsl:value-of select="$outputType"/>,<xsl:value-of select="$dateToExport"/>,,,,,,,"N",,,,,"N","N",<xsl:text></xsl:text>

<!-- Now for the gift certificate; NOTE: Requires Gift Certificate Module -->
<xsl:for-each select="ssGiftCertificate">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT -->"<xsl:value-of select="$giftCertificateName"/>"<xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT -->"<xsl:value-of select="format-number(-1*number(RedemptionAmount),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"<xsl:value-of select="$giftCertificateMemo"/>"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE -->"<xsl:value-of select="format-number(number(RedemptionAmount),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- QNTY -->"-1"<xsl:text>,</xsl:text>
<!-- INVITEM -->"<xsl:value-of select="$giftCertificateInvItem"/>"<xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:for-each>

<!-- Now for the purchase order -->
<xsl:if test="orderPurchaseOrderNumber[.!='']">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"Purchase Order: <xsl:value-of select="orderPurchaseOrderName" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"                <xsl:value-of select="orderPurchaseOrderNumber" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the credit card -->
<xsl:if test="CreditCardNumber[.!='']">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"Credit Card # :<xsl:value-of select="CreditCardNumber" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the credit card expiration -->
<xsl:if test="payCardType[.!='']">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"Card Expires:<xsl:value-of select="payCardExpires" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the credit card type -->
<xsl:if test="payCardType[.!='']">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"Card Type:<xsl:value-of select="payCardType" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for an empty line -->
SPL,,<xsl:value-of select="$outputType"/>,<xsl:value-of select="$dateToExport"/>,,,,,,,"N",,,,,"N","N",<xsl:text></xsl:text>

<!-- Now for the email -->
<xsl:if test="billingAddress/Email[.!='']">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"E-Mail:<xsl:value-of select="billingAddress/Email" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the phone -->
<xsl:if test="billingAddress/Phone[.!='']">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"Phone<xsl:text> </xsl:text>:<xsl:value-of select="billingAddress/Phone" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the fax  -->
<xsl:if test="billingAddress/Fax[.!='']">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"Fax <xsl:text> </xsl:text> :<xsl:value-of select="billingAddress/Fax" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for an empty line -->
SPL,,<xsl:value-of select="$outputType"/>,<xsl:value-of select="$dateToExport"/>,,,,,,,"N",,,,,"N","N",<xsl:text></xsl:text>

<!-- Now for the customer comments -->
<xsl:if test="orderComments[.!='']">
SPL<xsl:text>,</xsl:text>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT --><xsl:text>,</xsl:text>
<!-- NAME --><xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT --><xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO -->"Comments:<xsl:value-of select="orderComments" />"<xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE --><xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM --><xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"N"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->
</xsl:if>

<!-- Now for the tax -->
SPL<xsl:text>,</xsl:text><xsl:if test="orderSTax>0"><xsl:value-of select="TaxEntity"/></xsl:if>
<!-- SPLID --><xsl:text>,</xsl:text>
<!-- TRNSTYPE --><xsl:value-of select="$outputType"/><xsl:text>,</xsl:text>
<!-- DATE --><xsl:value-of select="$dateToExport"/><xsl:text>,</xsl:text>
<!-- ACCNT -->"<xsl:value-of select="$salesTaxName"/>"<xsl:text>,</xsl:text>
<!-- NAME -->"<xsl:if test="orderSTax>0"><xsl:value-of select="TaxEntity"/></xsl:if>"<xsl:text>,</xsl:text>
<!-- CLASS --><xsl:text>,</xsl:text>
<!-- AMOUNT -->"-<xsl:value-of select="format-number(number(orderSTax),'###0.00')" />"<xsl:text>,</xsl:text>
<!-- DOCNUM --><xsl:text>,</xsl:text>
<!-- MEMO --><xsl:text>,</xsl:text>
<!-- CLEAR -->"N"<xsl:text>,</xsl:text>
<!-- PRICE -->"<xsl:value-of select="TaxRate" />"<xsl:text>,</xsl:text>
<!-- QNTY --><xsl:text>,</xsl:text>
<!-- INVITEM -->"<xsl:value-of select="TaxEntity"/>"<xsl:text>,</xsl:text>
<!-- PAYMETH --><xsl:text>,</xsl:text>
<!-- TAXABLE -->"<xsl:value-of select="Taxable"/>"<xsl:text>,</xsl:text>
<!-- VALADJ -->"N"<xsl:text>,</xsl:text>
<!-- EXTRA -->"<xsl:value-of select="TaxCalcMeth"/>"<xsl:text></xsl:text>

<!-- Mark the transaction end -->
ENDTRNS<xsl:text></xsl:text>
<!-- End of Line --><xsl:text>&#10;</xsl:text>
</xsl:template>

</xsl:stylesheet>
