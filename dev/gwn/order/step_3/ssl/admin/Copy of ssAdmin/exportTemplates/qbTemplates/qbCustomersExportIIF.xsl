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

This module creates a QuickBooks 2002 Pro compatible information file in the format:

!CUST,NAME,BADDR1,BADDR2,BADDR3,BADDR4,BADDR5,SADDR1,SADDR2,SADDR3,SADDR4,SADDR5,PHONE1,PHONE2,FAXNUM,CONT1,CONT2,CTYPE,TERMS,TAXABLE,TAXITEM,COMPANYNAME,FIRSTNAME,MIDINIT,LASTNAME,CUSTFLD3,CUSTFLD4,CUSTFLD5,CUSTFLD6,EMAIL

'Notes:
'Shipping lines import as follows:
ADDR1 - Name
ADDR2 - Company, if present
ADDR3 - Address, line 1
ADDR4 - Address, line 2, if present
ADDR5 - City, State, ZIP (Country if not U.S. AND Company AND line 2 is present)
ADDR_Alt2 - Spacer if no company or line 2 And Country outside U.S.
ADDR_Alt4 - Spacer if no company, no line 2 And Country outside U.S.

-->

<xsl:import href="qbAccountSettings.xsl"/>

<xsl:template match='/'>
<xsl:text>!CUST,NAME,BADDR1,BADDR2,BADDR3,BADDR4,BADDR5,SADDR1,SADDR2,SADDR3,SADDR4,SADDR5,PHONE1,PHONE2,FAXNUM,CONT1,CONT2,CTYPE,TERMS,TAXABLE,TAXITEM,COMPANYNAME,FIRSTNAME,MIDINIT,LASTNAME,CUSTFLD1,CUSTFLD2,CUSTFLD3,CUSTFLD4,CUSTFLD5,CUSTFLD6,CUSTFLD7,EMAIL</xsl:text>
<!-- End of Line --><xsl:text>&#10;</xsl:text>
<xsl:apply-templates select="orders/order" />
</xsl:template>

<xsl:template match="order" >
<!-- CUST -->
<!--
Country Name is added if not U.S.
Note: QB can only accept five lines which are all accounted for with company name/address line 2
Country name added to end of ZIP if company and address line 2 present
-->
<xsl:text>CUST,</xsl:text>
<!-- NAME -->"<xsl:apply-templates select="billingAddress"/>"<xsl:text>,</xsl:text>
<!-- BADDR1 -->"<xsl:value-of select="billingAddress/FirstName" /><xsl:text> </xsl:text><xsl:if test="string-length(billingAddress/MiddleInitial)>0"><xsl:value-of select="billingAddress/MiddleInitial" /><xsl:text> </xsl:text></xsl:if><xsl:value-of select="billingAddress/LastName" />"<xsl:text>,</xsl:text>
<!-- BADDR2 --><xsl:if test="string-length(billingAddress/Company)>0">"<xsl:value-of select="billingAddress/Company" />",</xsl:if>
<!-- BADDR3 -->"<xsl:value-of select="billingAddress/Addr1" />"<xsl:text>,</xsl:text>
<!-- BADDR4 --><xsl:if test="string-length(billingAddress/Addr2)>0">"<xsl:value-of select="billingAddress/Addr2" />",</xsl:if>
<!-- BADDR5 -->"<xsl:value-of select="billingAddress/City" /><xsl:text>,</xsl:text><xsl:text> </xsl:text><xsl:value-of select="billingAddress/State" /><xsl:text> </xsl:text><xsl:value-of select="billingAddress/Zip" /><xsl:if test="(string-length(billingAddress/Company)>0 and string-length(billingAddress/Addr2)>0)"><xsl:if test="billingAddress/Country!='US'"><xsl:text> </xsl:text><xsl:value-of select="billingAddress/CountryName" /></xsl:if></xsl:if>"<xsl:text>,</xsl:text>
<!-- BADDR_Alt2 --><xsl:if test="(string-length(billingAddress/Company)=0 or string-length(billingAddress/Addr2)=0)">"<xsl:if test="billingAddress/Country!='US'"><xsl:value-of select="billingAddress/CountryName" /></xsl:if>",</xsl:if>
<!-- BADDR_Alt4 --><xsl:if test="(string-length(billingAddress/Company)=0 and string-length(billingAddress/Addr2)=0)">"",</xsl:if>
<!-- SADDR1 -->"<xsl:value-of select="shippingAddress/FirstName" /><xsl:text> </xsl:text><xsl:if test="string-length(shippingAddress/MiddleInitial)>0"><xsl:value-of select="shippingAddress/MiddleInitial" /><xsl:text> </xsl:text></xsl:if><xsl:value-of select="shippingAddress/LastName" />"<xsl:text>,</xsl:text>
<!-- SADDR2 --><xsl:if test="string-length(shippingAddress/Company)>0">"<xsl:value-of select="shippingAddress/Company" />",</xsl:if>
<!-- SADDR3 -->"<xsl:value-of select="shippingAddress/Addr1" />"<xsl:text>,</xsl:text>
<!-- SADDR4 --><xsl:if test="string-length(shippingAddress/Addr2)>0">"<xsl:value-of select="shippingAddress/Addr2" />",</xsl:if>
<!-- SADDR5 -->"<xsl:value-of select="shippingAddress/City" /><xsl:text>,</xsl:text><xsl:text> </xsl:text><xsl:value-of select="shippingAddress/State" /><xsl:text> </xsl:text><xsl:value-of select="shippingAddress/Zip" /><xsl:if test="(string-length(shippingAddress/Company)>0 and string-length(shippingAddress/Addr2)>0)"><xsl:if test="shippingAddress/Country!='US'"><xsl:text> </xsl:text><xsl:value-of select="shippingAddress/CountryName" /></xsl:if></xsl:if>"<xsl:text>,</xsl:text>
<!-- SADDR_Alt2 --><xsl:if test="(string-length(shippingAddress/Company)=0 or string-length(shippingAddress/Addr2)=0)">"<xsl:if test="shippingAddress/Country!='US'"><xsl:value-of select="shippingAddress/CountryName" /></xsl:if>",</xsl:if>
<!-- SADDR_Alt4 --><xsl:if test="(string-length(shippingAddress/Company)=0 and string-length(shippingAddress/Addr2)=0)">"",</xsl:if>
<!-- PHONE1 -->"<xsl:value-of select="billingAddress/Phone" />"<xsl:text>,</xsl:text>
<!-- PHONE2 -->"<xsl:value-of select="shippingAddress/Phone" />"<xsl:text>,</xsl:text>
<!-- FAXNUM -->"<xsl:value-of select="billingAddress/Fax" />"<xsl:text>,</xsl:text>
<!-- CONT1 --><xsl:text>,</xsl:text>
<!-- CONT2 --><xsl:text>,</xsl:text>
<!-- CTYPE --><xsl:text>,</xsl:text>
<!-- TERMS -->"<xsl:value-of select="orderPaymentMethod" />"<xsl:text>,</xsl:text>
<!-- TAXABLE --><xsl:if test="orderSTax>0">Y</xsl:if><xsl:if test="orderSTax=0">N</xsl:if><xsl:text>,</xsl:text>
<!-- TAXITEM -->"<xsl:value-of select="TaxEntity"/>"<xsl:text>,</xsl:text>
<!-- COMPANYNAME --><xsl:if test="string-length(billingAddress/Company)>0">"<xsl:value-of select="billingAddress/Company" />"</xsl:if><xsl:text>,</xsl:text>
<!-- FIRSTNAME -->"<xsl:value-of select="billingAddress/FirstName" />"<xsl:text>,</xsl:text>
<!-- MIDINIT -->"<xsl:value-of select="billingAddress/MiddleInitial" />"<xsl:text>,</xsl:text>
<!-- LASTNAME -->"<xsl:value-of select="billingAddress/LastName" />"<xsl:text>,</xsl:text>
<!-- CUSTFLD1 --><xsl:text>,</xsl:text>
<!-- CUSTFLD2 --><xsl:text>,</xsl:text>
<!-- CUSTFLD3 --><xsl:text>,</xsl:text>
<!-- CUSTFLD4 --><xsl:text>,</xsl:text>
<!-- CUSTFLD5 --><xsl:text>,</xsl:text>
<!-- CUSTFLD6 --><xsl:text>,</xsl:text>
<!-- CUSTFLD7 --><xsl:text>,</xsl:text>
<!-- EMAIL -->"<xsl:value-of select="billingAddress/Email" />"<xsl:text></xsl:text>
<!-- End of Line --><xsl:text>&#10;</xsl:text>
</xsl:template>
</xsl:stylesheet>