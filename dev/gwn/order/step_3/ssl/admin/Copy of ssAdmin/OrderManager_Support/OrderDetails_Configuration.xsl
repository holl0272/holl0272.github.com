<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:variable name="orderSubject" select="'Your Sandshot Software Order'" />
	<xsl:variable name="displayBackOrderByDefault" select="'false'" />
	<xsl:variable name="currencySymbol" select="orders/order/currencySymbol" />
	<xsl:variable name="decimalSeparator" select="orders/order/decimalSeparator" />
	<xsl:variable name="ordersExtra1_Label" select="orders/orderDetailDispalyOptions/ordersExtra1_Label" />
	<xsl:variable name="Insured_Label" select="orders/orderDetailDispalyOptions/Insured_Label" />
	<xsl:variable name="PackageWeight_Label" select="orders/orderDetailDispalyOptions/PackageWeight_Label" />
	<xsl:variable name="orderTrackingExtra1_Label" select="orders/orderDetailDispalyOptions/orderTrackingExtra1_Label" />
</xsl:stylesheet>
