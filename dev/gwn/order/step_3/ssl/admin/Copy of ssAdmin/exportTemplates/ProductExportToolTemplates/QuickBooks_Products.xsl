<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:import href="../supportingTemplates/supportTemplates.xsl"/>
<xsl:output method="text"/>
<!--
'********************************************************************************
'*   Product Export Tool - Froogle Module				                        *
'*   Release Version:   1.00.000												*
'*   Release Date:		April 6, 2005											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2005 Sandshot Software.  All rights reserved.                *
'********************************************************************************

-->
<xsl:variable name="outputDelimeter">,</xsl:variable>

<xsl:template match='/'>

	<xsl:if test="False">
		<xsl:text>!HDR</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>PROD</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>VER</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>REL</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>IIFVER</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>DATE</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>TIME</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>ACCNTNT</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>ACCNTNTSPLITTIME</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>

		<xsl:text>HDR</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>QuickBooks Pro</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Version 13.0D</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Release R8P</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>1</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>2005-10-28</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>1130526712</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>N</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>0</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>

		<xsl:text>!CUSTITEMDICT</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>INDEX</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>LABEL</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>INUSE</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>

		<xsl:text>!ENDCUSTITEMDICT</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>

		<xsl:text>CUSTITEMDICT</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>0</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Name Embroidery</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>Y</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		
		<xsl:text>CUSTITEMDICT</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>1</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>N</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		
		<xsl:text>CUSTITEMDICT</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>2</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>N</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		
		<xsl:text>CUSTITEMDICT</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>3</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>N</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		
		<xsl:text>CUSTITEMDICT</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>4</xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
		<xsl:text>N</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
		
		<xsl:text>ENDCUSTITEMDICT</xsl:text>
		<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
	</xsl:if>

	<xsl:text>!INVITEM</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>NAME</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>REFNUM</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>TIMESTAMP</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>INVITEMTYPE</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>DESC</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>PURCHASEDESC</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ACCNT</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ASSETACCNT</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>COGSACCNT</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>QNTY</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>QNTY</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>PRICE</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>COST</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>TAXABLE</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>SALESTAXCODE</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>PAYMETH</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>TAXVEND</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>TAXDIST</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>PREFVEND</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>REORDERPOINT</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>EXTRA</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>CUSTFLD1</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>CUSTFLD2</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>CUSTFLD3</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>CUSTFLD4</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>CUSTFLD5</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>DEP_TYPE</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ISPASSEDTHRU</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>HIDDEN</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>DELCOUNT</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>USEID</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ISNEW</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>PO_NUM</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>SERIALNUM</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>WARRANTY</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>LOCATION</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>VENDOR</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ASSETDESC</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>SALEDATE</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>SALEEXPENSE</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>NOTES</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ASSETNUM</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>COSTBASIS</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>ACCUMDEPR</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>UNRECBASIS</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<xsl:text>PURCHASEDATE</xsl:text>
	<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
	
	<xsl:for-each select="products/product">
		<xsl:variable name="Code" select="Code" />
		<xsl:variable name="Name" select="Name" />
		<xsl:variable name="Description" select="ShortDescription" />
		<xsl:variable name="VendorId" select="VendorId" />
		<xsl:variable name="prodVend" select="../../products/Vendors/Vendor[@uid = $VendorId]/name" />

		<xsl:call-template name="productItem">
			<xsl:with-param name="prodName" select="$Code" />
			<xsl:with-param name="prodPrice" select="Price" />
			<xsl:with-param name="position" select="position()" />
			<xsl:with-param name="prodDesc" select="$Name" />
			<xsl:with-param name="prodDescLong" select="$Description" />
			<xsl:with-param name="prodVend" select="$prodVend" />

		</xsl:call-template>

		<xsl:for-each select="attributes">
			<xsl:for-each select="attributeCategory">
				<xsl:variable name="AttrCategory">
					<xsl:call-template name="replace-string">
						<xsl:with-param name="text" select="name" />
						<xsl:with-param name="replace" select="':'" />
						<xsl:with-param name="with" select="''" />
					</xsl:call-template>
				</xsl:variable>
				
				<xsl:call-template name="productItem">
					<xsl:with-param name="prodName">
						<xsl:value-of select="$Code"/>
						<xsl:if test="$AttrCategory!=''"><xsl:text>:</xsl:text><xsl:value-of select="$AttrCategory"/></xsl:if>
					</xsl:with-param>
					<xsl:with-param name="prodPrice" select="'0'" />
					<xsl:with-param name="prodDesc">
						<xsl:value-of select="$Name"/><xsl:text>, </xsl:text><xsl:value-of select="$AttrCategory"/>
					</xsl:with-param>
					<xsl:with-param name="prodDescLong" select="$Description" />
					<xsl:with-param name="position" select="position()" />
					<xsl:with-param name="prodVend" select="$prodVend" />
				</xsl:call-template>

				<xsl:for-each select="attribute">
					<xsl:call-template name="productItem">
						<xsl:with-param name="prodName">
							<xsl:value-of select="$Code"/><xsl:text>:</xsl:text><xsl:value-of select="$AttrCategory"/><xsl:text>:</xsl:text><xsl:value-of select="name"/>
						</xsl:with-param>
						<xsl:with-param name="prodPrice" select="price" />
						<xsl:with-param name="prodDesc">
							<xsl:value-of select="$Name"/><xsl:text>, </xsl:text><xsl:value-of select="$AttrCategory"/><xsl:text>:</xsl:text><xsl:value-of select="name"/>
						</xsl:with-param>
						<xsl:with-param name="prodDescLong" select="$Description" />
						<xsl:with-param name="position" select="position()" />
						<xsl:with-param name="prodVend" select="$prodVend" />
					</xsl:call-template>
				</xsl:for-each>
			</xsl:for-each>
		</xsl:for-each>
	</xsl:for-each>


	<!-- !INVITEM --><xsl:text>INVITEM</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- NAME --><xsl:text>Out of State</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- REFNUM --><xsl:text>1</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- TIMESTAMP --><xsl:text>1128623421</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- INVITEMTYPE --><xsl:text>COMPTAX</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- DESC --><xsl:text>"Out-of-state sale, exempt from sales tax"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PURCHASEDESC --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ACCNT --><xsl:text>Sales Tax Payable</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ASSETACCNT --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- COGSACCNT --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- QNTY --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- QNTY --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PRICE --><xsl:text>0.0%</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- COST --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- TAXABLE --><xsl:text>N</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- SALESTAXCODE --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PAYMETH --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- TAXVEND --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- TAXDIST --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PREFVEND --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- REORDERPOINT --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- EXTRA --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD1 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD2 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD3 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD4 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD5 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- DEP_TYPE --><xsl:text>0</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ISPASSEDTHRU --><xsl:text>N</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- HIDDEN --><xsl:text>N</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- DELCOUNT --><xsl:text>0</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- USEID --><xsl:text>N</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ISNEW --><xsl:text>Y</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PO_NUM --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- SERIALNUM --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- WARRANTY --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- LOCATION --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- VENDOR --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ASSETDESC --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- SALEDATE --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- SALEEXPENSE --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- NOTES --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ASSETNUM --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- COSTBASIS --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ACCUMDEPR --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- UNRECBASIS --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PURCHASEDATE --><xsl:text></xsl:text>
	<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>

</xsl:template>

<xsl:template name="productItem">
	<xsl:param name="prodName" />
	<xsl:param name="prodPrice" />
	<xsl:param name="position" />
	<xsl:param name="prodDesc" />
	<xsl:param name="prodVend" />
	<xsl:param name="prodDescLong" />
	
	<!-- !INVITEM --><xsl:text>INVITEM</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- NAME --><xsl:value-of select="$prodName"/><xsl:value-of select="$outputDelimeter"/>
	<!-- REFNUM --><xsl:value-of select="$position"/><xsl:value-of select="$outputDelimeter"/>
	<!-- TIMESTAMP --><xsl:text>1130435347</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- INVITEMTYPE --><xsl:text>INVENTORY</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- DESC --><xsl:text>"</xsl:text><xsl:value-of select="$prodDesc"/><xsl:text>"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PURCHASEDESC --><xsl:text>"</xsl:text><xsl:value-of select="$prodDescLong"/><xsl:text>"</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ACCNT --><xsl:text>Sales:Merchandise</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ASSETACCNT --><xsl:text>Inventory Asset</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- COGSACCNT --><xsl:text>Cost of Goods Sold</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- QNTY --><xsl:text>0</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- QNTY --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PRICE --><xsl:value-of select="$prodPrice"/>
	<!-- COST --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- TAXABLE --><xsl:text>Y</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- SALESTAXCODE --><xsl:text>Tax</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PAYMETH --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- TAXVEND --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- TAXDIST --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PREFVEND --><xsl:value-of select="$prodVend"/><xsl:value-of select="$outputDelimeter"/>
	<!-- REORDERPOINT --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- EXTRA --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD1 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD2 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD3 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD4 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- CUSTFLD5 --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- DEP_TYPE --><xsl:text>0</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ISPASSEDTHRU --><xsl:text>N</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- HIDDEN --><xsl:text>N</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- DELCOUNT --><xsl:text>0</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- USEID --><xsl:text>N</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ISNEW --><xsl:text>Y</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PO_NUM --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- SERIALNUM --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- WARRANTY --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- LOCATION --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- VENDOR --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ASSETDESC --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- SALEDATE --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- SALEEXPENSE --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- NOTES --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ASSETNUM --><xsl:text></xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- COSTBASIS --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- ACCUMDEPR --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- UNRECBASIS --><xsl:text>0.00</xsl:text><xsl:value-of select="$outputDelimeter"/>
	<!-- PURCHASEDATE --><xsl:text></xsl:text>
	<!-- End of Line --><xsl:value-of select="$outputDelimeter_CR"/>
</xsl:template>

</xsl:stylesheet>

