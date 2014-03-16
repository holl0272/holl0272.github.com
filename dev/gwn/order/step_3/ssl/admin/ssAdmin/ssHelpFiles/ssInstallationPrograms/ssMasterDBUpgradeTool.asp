<%Option Explicit 
Server.ScriptTimeout = 600
'********************************************************************************
'*   Sandshot Software StoreFront add-on Master Database Upgrade Tool           *
'*   Release Version: 1.00.013 Beta	                                            *
'*   Release Date: January 30, 2003												*
'*   Revision Date: September 30, 2005											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   1.00.013 (September 30, 2005)												*
'*   - Updated to add custom upgrade						                    *
'*                                                                              *
'*   1.00.012 (May 30, 2005)													*
'*   - Updated to add Tax Rate Manager						                    *
'*                                                                              *
'*   1.00.011 (January 1, 2005)													*
'*   - Updated to function with Master Template				                    *
'*                                                                              *
'*   1.00.010 (November 20, 2004)												*
'*   - Added Order Manager v3 changes						                    *
'*                                                                              *
'*   1.00.009 (October 28, 2004)												*
'*   - Added CCV/StartDate/Issue Number update				                    *
'*                                                                              *
'*   1.00.008 (October 21, 2004)												*
'*   - Added File Download update							                    *
'*                                                                              *
'*   1.00.007 (June 13, 2004)													*
'*   - Added Attribute Extender update						                    *
'*                                                                              *
'*   1.00.006 (February 28, 2004)												*
'*   - Added support for Product Placement					                    *
'*                                                                              *
'*   1.00.005 (November 6, 2003)												*
'*   - Added support for Gift Certificate v1.01.001 changes                     *
'*                                                                              *
'*   1.00.004 (October 30, 2003)												*
'*   - Added support for Dynamic Product Display                                *
'*                                                                              *
'*   1.00.003 (August 18, 2003)													*
'*   - Prior release                                                            *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

sub debugprint(sField1,sField2)
Response.Write "<h4>" & sField1 & ": " & sField2 & "</h4>"
end sub

Const ssTemplateVersion = "1.01.006"

Const enName = 0
Const enInstalled = 1
Const enDBUpgraded = 2
Const enAddonVersion = 3
Const enAddonLink = 4
Const enUpgradeDetails = 5

Dim enAO_AttributeExtender
Dim enAO_GiftCertificate
Dim enAO_PayPalPayments
Dim enAO_PostageRate
Dim enAO_PricingLevel
Dim enAO_PromoMail
Dim enAO_PromoMgr
Dim enAO_PromoMgrII
Dim enAO_OrderMgr
Dim enAO_WebStoreMgr
Dim enAO_ZBS
Dim enAO_DynamicProduct
Dim enAO_ProductPlacement
Dim enAO_FileDownload
Dim enAO_CCV
Dim enAO_MasterTemplate
Dim enAO_ContentManagement
Dim enAO_TaxRateManager
Dim enAO_SEtoAEUpgrade
Dim enAO_VisitorTracking
Dim enAO_SQLSpeedUpgrade
Dim enAO_CustomUpgrade
Dim enAO_BuyersClub
Dim enAO_DatabaseSize

Dim enMaxEN
Dim ArrayCounter: Call ResetArrayIndex

Call SetInstallationOrder

'***********************************************************************************************

Sub SetInstallationOrder
	enAO_AttributeExtender = ArrayIndex
	enAO_BuyersClub = ArrayIndex
	enAO_ContentManagement = ArrayIndex
	enAO_CCV = ArrayIndex
	enAO_DynamicProduct = ArrayIndex
	enAO_FileDownload = ArrayIndex
	enAO_GiftCertificate = ArrayIndex
	enAO_PayPalPayments = ArrayIndex
	enAO_PostageRate = ArrayIndex
	enAO_PricingLevel = ArrayIndex
	enAO_ProductPlacement = ArrayIndex
	enAO_PromoMail = ArrayIndex
	enAO_PromoMgr = ArrayIndex
	enAO_PromoMgrII = ArrayIndex
	enAO_OrderMgr = ArrayIndex
	enAO_TaxRateManager = ArrayIndex
	enAO_VisitorTracking = ArrayIndex
	enAO_WebStoreMgr = ArrayIndex
	enAO_ZBS = ArrayIndex
	enAO_MasterTemplate = ArrayIndex
	enAO_SEtoAEUpgrade = ArrayIndex
	enAO_DatabaseSize = ArrayIndex
	enAO_SQLSpeedUpgrade = ArrayIndex
	enAO_CustomUpgrade = ArrayIndex

	enMaxEN = ArrayCounter
	Call ResetArrayIndex
End Sub

'***********************************************************************************************


Function ArrayIndex()
	ArrayCounter = ArrayCounter + 1
	ArrayIndex = ArrayCounter
End Function
Sub ResetArrayIndex
	ArrayCounter = -1
End Sub

'***********************************************************************************************

Const enUpdate_Action = 0
Const enUpdate_Field = 1
Const enUpdate_Type = 2

Const dbProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
Dim mstrDBPath
Dim mstrDBName

Function adjustSQLServerType(byVal strType)

Dim pstrTempType
				
	If mblnSQLServer Then
		Select Case LCase(strType)
			Case "byte": pstrTempType = "int"
			Case "double": pstrTempType = "decimal (10,2)"
			Case "long": pstrTempType = "int"
			Case "date": pstrTempType = "Datetime"
			Case "yesno": pstrTempType = "bit"
			Case "memo": pstrTempType = "text"
			Case "counter": pstrTempType = "int identity"
			Case Else: pstrTempType = strType
		End Select
	Else
		pstrTempType = strType
	End If

	adjustSQLServerType = pstrTempType
		
End Function	'adjustSQLServerType

'***********************************************************************************************

Sub SetAvailableAddOns(byRef aryAddons)

Dim plngStepCounter

ReDim maryAddons(enMaxEN)
'0 - addon name
'1 - is installed
'2 - is db upgraded
'3 - version

	aryAddons(enAO_AttributeExtender) = Array("Attribute Extender", "N/A", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Attribute Extender</h4></td></tr><tr><td><ol><li>Modify sfAttributes table</li><ul><li>Add field: attrDisplayStyle</li><li>Add field: attrDisplayOrder</li></ul><li>Modify sfOrderAttributes table</li><ul><li>Make odrattrName a memo/text field</li></ul><li>Modify sfTmpOrderAttributes table</li><ul><li>Make odrattrtmpAttrID a memo/text field</li></ul></ol></td></tr></table>")
	aryAddons(enAO_GiftCertificate) = Array("Gift Certificate", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Gift Certificate</h4></td></tr><tr><td><ol><li>adds the ssGiftCertificates table</li><li>adds the ssGiftCertificateRedemptions table</li></ol></td></tr></table>")
	aryAddons(enAO_PayPalPayments) = Array("PayPal Payments", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>PayPal Payments</h4></td></tr><tr><td><OL><LI>Creates a new table PayPalIPNs</LI><LI>Creates a new table PayPalPayments</LI></OL></td></tr></table>")
	aryAddons(enAO_PostageRate) = Array("Postage Rate", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Postage Rate</h4></td></tr><tr><td><OL><LI>adds the ssShippingCarriers table</LI><LI>Adds a new table <i>ssShippingMethods</i></LI></OL></td></tr></table>")
	aryAddons(enAO_PricingLevel) = Array("Pricing Level Manager", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Pricing Level Manager</h4></td></tr><tr><td><OL><LI>Adds a new table <i>PricingLevel</i></LI><LI>upgrades the sfAttributeDetail, sfCustomers, sfMTPrices(AE Only), and sfProducts tables</LI></OL></td></tr></table>")
	aryAddons(enAO_PromoMail) = Array("Promotion Mail Manager", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Promotional Mail Manager</h4></td></tr><tr><td><ol><li>Updates sfCustomers table - adds ssPromoMailSent field</li></ol></td></tr></table>")
	aryAddons(enAO_PromoMgr) = Array("Promotion Manager", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Promotion Manager</h4></td></tr><tr><td><OL><LI>Creates a new table Promotions </LI><LI>Creates a new table orderDiscounts</LI></OL></td></tr></table>")
	aryAddons(enAO_PromoMgrII) = Array("Promotion Manager II", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Promotion Manager</h4></td></tr><tr><td><OL><LI>Creates a new table Promotions </LI><LI>Creates a new table orderDiscounts</LI></OL></td></tr></table>")
	aryAddons(enAO_OrderMgr) = Array("Order Manger", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Order Manager</h4></td></tr><tr><td><ol><li>Adds a new table <i>ssOrderManager</i></li></ol></td></tr></table>")
	aryAddons(enAO_WebStoreMgr) = Array("WebStore Manger - Integrated Security", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>WebStore Manager - Integrated Security</h4></td></tr><tr><td><ol><li>Adds a new table <i>SSUsers</i><ul><li>creates a default username: <i>admin</i><li>creates a default password: <i>pass</i></li></ul></td></tr></table>")
	aryAddons(enAO_ZBS) = Array("Zone Based Shipping", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Zone Based Shipping</h4></td></tr><tr><td><OL><LI>Updates sfShipping table</LI><ul><LI>Increases shipMethod field size to 65</LI><LI>Increases shipCode field size to 60 and requires unique values</li></ul><LI>Creates a new table ssShipZones</LI><LI>Creates a new table ssShippingRates</LI></OL></td></tr></table>")
	aryAddons(enAO_DynamicProduct) = Array("Dynamic Product", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Dynamic Product</h4></td></tr><tr><td><OL><LI>Updates sfProduct Table table</LI><ul><LI>relatedProducts field added</LI></ul></OL></td></tr></table>")
	aryAddons(enAO_ProductPlacement) = Array("Product Placement", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Product Placement</h4></td></tr><tr><td><OL><LI>Updates sfProduct Table table</LI><ul><LI>sortCat field added</LI><LI>sortMfg field added</LI><LI>sortVend field added</LI></ul><li>Updates sfSubCatDetail (AE only)</li><ul><li>sortCatDetail field added</li></ul></OL></td></tr></table>")
	aryAddons(enAO_FileDownload) = Array("File Download", "N/A", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>File Download</h4></td></tr><tr><td>See the read me file</td></tr></table>")
	aryAddons(enAO_CCV) = Array("CCV", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>CCV</h4></td></tr><tr><td><OL><LI>Updates sfCPayments Table table</LI><ul><LI>payCardCCV field added</LI></ul></OL></td></tr></table>")
	aryAddons(enAO_MasterTemplate) = Array("Master Template", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>CCV</h4></td></tr><tr><td><OL><LI>Updates sfCPayments Table table</LI><ul><LI>payCardCCV field added</LI></ul></OL></td></tr></table>")
	aryAddons(enAO_ContentManagement) = Array("Content Management", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>CCV</h4></td></tr><tr><td><OL><LI>Updates sfCPayments Table table</LI><ul><LI>payCardCCV field added</LI></ul></OL></td></tr></table>")
	aryAddons(enAO_TaxRateManager) = Array("Tax Rate Manager", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Tax Rate Manager</h4></td></tr><tr><td><OL><li>Creates ssTaxTable table</li></ul></OL></td></tr></table>")
	aryAddons(enAO_SEtoAEUpgrade) = Array("SE to AE upgrade", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>SE to AE database upgrade</h4></td></tr><tr><td><h3>This can be used to add AE specific tables to an SE database. You must own an AE license</h3><OL><LI>Adds sfCoupons, sfGiftWraps, sfInventory, sfInventoryInfo, sfMTPrices, sfOrderDetailsAE, sfOrdersAE, sfTmpOrderDetailsAE, sfTmpOrdersAE tables</LI></ul></OL></td></tr></table>")
	aryAddons(enAO_CustomUpgrade) = Array("Custom", "Unknown", False, "", "", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Custom Upgrade</h4></td></tr></table>")
	aryAddons(enAO_VisitorTracking) = Array("Visitor Tracking", "Unknown", False, "", "", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Visitor Tracking</h4></td></tr></table>")
	aryAddons(enAO_SQLSpeedUpgrade) = Array("SQL Server Database Speed", "Unknown", False, "", "", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>SQL Server Database Size</h4><p>This upgrade modifies the product and visitor tables to change nvarchar to varchar and ntext to nvarchar(2000) (or the largest existing field size)</p><p style='color:red'>This will disable application support for unicode characters</p></td></tr></table>")
	aryAddons(enAO_BuyersClub) = Array("Buyers Club", "Unknown", False, "", "http://www.addons4storefront.net/", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>CCV</h4></td></tr><tr><td><OL><LI>Updates sfCPayments Table table</LI><ul><LI>payCardCCV field added</LI></ul></OL></td></tr></table>")
	aryAddons(enAO_DatabaseSize) = Array("Database Size", "Unknown", False, "", "", "<table cellpadding=0 cellspacing=0 border=0><tr><td><h4>Database Size</h4><p>This upgrade modifies the product table to reduce the field size for lesser used features</p>" _
																					& "<p style='color:red'>This script will alter the field sizes for lesser used functions. Note: If you want, you may manually edit these settings in the database to retain some specific functionality.<br />The following fields are used for SEO specific functions; if you do not use this you may perform the automated routine which shrinks them to one character.<ul><li>pageName</li><li>metaTitle</li><li>metaDescription</li><li>metaKeywords</li></ul><br />The following fields are used for download specific products<ul><li>prodFileName</li><li>UpgradeVersion</li><li>version</li><li>InstallationHours</li><li>packageCodes</li></ul>Note: Access databases must have a compact and repair accomplished to see the size reduction.")

	If Len(mstrUpgradeItem) > 0 Then
		aryAddons(enAO_AttributeExtender)(enInstalled) = CBool(mstrUpgradeItem = "AttributeExtender")
		aryAddons(enAO_GiftCertificate)(enInstalled) = CBool(mstrUpgradeItem = "GiftCertificate")
		aryAddons(enAO_PayPalPayments)(enInstalled) = CBool(mstrUpgradeItem = "PayPalPayments")
		aryAddons(enAO_PostageRate)(enInstalled) = CBool(mstrUpgradeItem = "PostageRate")
		aryAddons(enAO_PricingLevel)(enInstalled) = CBool(mstrUpgradeItem = "PricingLevel")
		aryAddons(enAO_PromoMail)(enInstalled) = CBool(mstrUpgradeItem = "PromoMail")
		aryAddons(enAO_PromoMgr)(enInstalled) = CBool(mstrUpgradeItem = "PromoMgr")
		aryAddons(enAO_PromoMgrII)(enInstalled) = CBool(mstrUpgradeItem = "PromoMgrII")
		aryAddons(enAO_OrderMgr)(enInstalled) = CBool(mstrUpgradeItem = "OrderMgr")
		aryAddons(enAO_WebStoreMgr)(enInstalled) = CBool(mstrUpgradeItem = "WebStoreMgr")
		aryAddons(enAO_ZBS)(enInstalled) = CBool(mstrUpgradeItem = "ZBS")
		aryAddons(enAO_DynamicProduct)(enInstalled) = CBool(mstrUpgradeItem = "DynamicProduct")
		aryAddons(enAO_ProductPlacement)(enInstalled) = CBool(mstrUpgradeItem = "ProductPlacement")
		aryAddons(enAO_FileDownload)(enInstalled) = CBool(mstrUpgradeItem = "FileDownload")
		aryAddons(enAO_CCV)(enInstalled) = CBool(mstrUpgradeItem = "CCV")
		aryAddons(enAO_MasterTemplate)(enInstalled) = CBool(mstrUpgradeItem = "MasterTemplate")
		aryAddons(enAO_ContentManagement)(enInstalled) = CBool(mstrUpgradeItem = "ContentManagement")
		aryAddons(enAO_TaxRateManager)(enInstalled) = CBool(mstrUpgradeItem = "TaxRate")
		aryAddons(enAO_SEtoAEUpgrade)(enInstalled) = CBool(mstrUpgradeItem = "SEtoAE")
		aryAddons(enAO_VisitorTracking)(enInstalled) = CBool(mstrUpgradeItem = "VisitorTracking")
		aryAddons(enAO_SQLSpeedUpgrade)(enInstalled) = CBool(mstrUpgradeItem = "SQLSpeedUpgrade")
		aryAddons(enAO_DatabaseSize)(enInstalled) = CBool(mstrUpgradeItem = "DatabaseSize")
		aryAddons(enAO_BuyersClub)(enInstalled) = CBool(mstrUpgradeItem = "BuyersClub")
	End If

End Sub	'SetAvailableAddOns

'***********************************************************************************************

Sub DoUpgrades(byRef objCnn, byRef aryAddons)

Dim i
Dim paryUpgradeActions
Dim paryAction

Dim pblnInstall
Dim pbytUpgradeToLatest

		paryUpgradeActions = Split(Request.Form("installAddon"),",")
		
		For i = 0 To UBound(paryUpgradeActions)
			paryAction = Split(paryUpgradeActions(i), ".")
			pblnInstall = CBool(paryAction(1))
			If UBound(paryAction) > 1 Then
				pbytUpgradeToLatest = paryAction(2)
			Else
				pbytUpgradeToLatest = 0
			End If
			
			Select Case CLng(paryAction(0))
				Case enAO_AttributeExtender
					Call Install_AttributeExtenderAddon(objCnn, pblnInstall, pbytUpgradeToLatest)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_GiftCertificate
					Call Install_GiftCertificateManagerAddon(objCnn, pblnInstall, pbytUpgradeToLatest)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_PayPalPayments
					Call Install_PayPalPaymentsAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_PostageRate
					Call Install_PostageRateAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_PricingLevel
					Call Install_PricingLevelAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_PromoMail
					Call Install_PromoMailAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_PromoMgr
					Call Install_PromotionManagerAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_PromoMgrII
					Call Install_PromotionManagerIIAddon(objCnn, pblnInstall, pbytUpgradeToLatest)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_OrderMgr
					Call Install_OrderMgrAddon(objCnn, pblnInstall, pbytUpgradeToLatest)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_WebStoreMgr
					Call Install_WebStoreManagerAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_ZBS
					Call Install_ZBSAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_DynamicProduct
					Call Install_DynamicProductAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_ProductPlacement
					Call Install_ProductPlacementAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_FileDownload
					Call Install_FileDownloadAddon(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_CCV
					Call Install_CCV(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_MasterTemplate
					Call Install_MasterTemplate(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_ContentManagement
					Call Install_ContentManagement(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_SEtoAEUpgrade
					Call Install_SEtoAEUpgrade(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_CustomUpgrade
					Call Install_CustomUpgrade(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_TaxRateManager
					Call Install_TaxRateAddon(objCnn, pblnInstall)
				Case enAO_VisitorTracking
					Call Install_VisitorTracking(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_SQLSpeedUpgrade
					Call Install_SQLSpeedUpgrade(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_DatabaseSize
					Call Install_DatabaseSizeUpgrade(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
				Case enAO_BuyersClub
					Call Install_BuyersClub(objCnn, pblnInstall)
					mstrMessage = "<hr><h3>" & aryAddons(paryAction(0))(enName) & " upgrade results:</h3>" & mstrMessage & "<hr>"
			End Select
'				debugprint "pblnInstall",pblnInstall
'				debugprint "pbytUpgradeToLatest",pbytUpgradeToLatest
		Next 'i

End Sub	'DoUpgrades

'***********************************************************************************************

Function FileExists(byVal strFilePath)

Dim pobjFSO
Dim pblnFileFound

	Set pobjFSO = CreateObject("Scripting.FileSystemObject")
	pblnFileFound = pobjFSO.FileExists(strFilePath)
	Set pobjFSO = Nothing
	
	FileExists = pblnFileFound

End Function	'FileExists
	
'***********************************************************************************************

Sub DetermineInstalledAddOns(aryAddons)

Dim pobjFSO
Dim pstrFilePath
Dim pstrFileToCheck
Dim i
Dim pstrFileFound

	pstrFilePath = Server.MapPath("AdminHeader.asp")
	pstrFilePath = Replace(Lcase(pstrFilePath),"adminheader.asp","")
	pstrFilePath = Replace(Lcase(pstrFilePath),"sshelpfiles\","")
	pstrFilePath = Replace(Lcase(pstrFilePath),"ssinstallationprograms\","")
	'debugprint "pstrFilePath",pstrFilePath

	Set pobjFSO = CreateObject("Scripting.FileSystemObject")
	
	If pobjFSO.FileExists(pstrFilePath & "adminheader.asp") Then
		For i = 0 To UBound(aryAddons)
			Select Case i
				Case enAO_AttributeExtender
					pstrFileToCheck = ""	'no test yet
				Case enAO_DynamicProduct
					pstrFileToCheck = ""	'no test yet
				Case enAO_CCV
					pstrFileToCheck = ""	'no test yet
				Case enAO_GiftCertificate
					pstrFileToCheck = pstrFilePath & "ssGiftCertificateAdmin.asp"
				Case enAO_PayPalPayments
					pstrFileToCheck = pstrFilePath & "ssPayPalPaymentsAdmin.asp"
				Case enAO_PostageRate
					pstrFileToCheck = pstrFilePath & "ssPostageRate_shippingMethodsAdmin.asp"
				Case enAO_PricingLevel
					pstrFileToCheck = pstrFilePath & "ssPricingLevelAdmin.asp"
				Case enAO_PromoMail
					pstrFileToCheck = pstrFilePath & "ssPromoMailAdmin.asp"
				Case enAO_PromoMgr
					pstrFileToCheck = pstrFilePath & "ssPromoAdmin.asp"
				Case enAO_PromoMgrII
					pstrFileToCheck = pstrFilePath & "ssPromotionsAdmin.asp"
				Case enAO_OrderMgr
					pstrFileToCheck = pstrFilePath & "ssOrderAdmin.asp"
				Case enAO_TaxRateManager
					pstrFileToCheck = pstrFilePath & "ssTaxRateAdmin.asp"
				Case enAO_WebStoreMgr
					pstrFileToCheck = pstrFilePath & "sfDesignAdmin.asp"
				Case enAO_ZBS
					pstrFileToCheck = pstrFilePath & "sszbsZoneAdmin.asp"
				Case enAO_ProductPlacement
					pstrFileToCheck = pstrFilePath & "ssProductPlacementAdmin.asp"
				Case enAO_FileDownload
					pstrFileToCheck = ""	'no test yet
				Case enAO_MasterTemplate
					pstrFileToCheck = ""	'no test yet
				Case enAO_SEtoAEUpgrade
					pstrFileToCheck = ""	'no test yet
				Case enAO_ContentManagement
					pstrFileToCheck = pstrFilePath & "ssContentAdmin.asp"
					pstrFileToCheck = ""	'no test yet
				Case enAO_VisitorTracking
					pstrFileToCheck = pstrFilePath & "ssVisitorTrackingAdmin.asp"
					pstrFileToCheck = ""	'no test yet
				Case enAO_SQLSpeedUpgrade
					pstrFileToCheck = pstrFilePath & "ssVisitorTrackingAdmin.asp"
					pstrFileToCheck = ""	'no test yet
				Case enAO_BuyersClub
					pstrFileToCheck = pstrFilePath & "ssBuyersClubAdmin.asp"
					pstrFileToCheck = ""	'no test yet
			End Select
			
			If Len(pstrFileToCheck) > 0 Then
				pstrFileFound = CStr(pobjFSO.FileExists(pstrFileToCheck))
				aryAddons(i)(enInstalled) = CStr(CBool(pstrFileFound))
			Else
				pstrFileFound = "N/A"	'no test yet
				aryAddons(i)(enInstalled) = True
			End If
			aryAddons(i)(enInstalled) = True

			'debugprint maryAddons(i)(enName), maryAddons(i)(enInstalled)
			'debugprint pstrFileToCheck, maryAddons(i)(enInstalled)
		Next 'i
	Else
		For i = 0 To UBound(aryAddons)
			If i <> enAO_AttributeExtender Then maryAddons(i)(enInstalled) = "Unknown"
				aryAddons(i)(enInstalled) = True
		Next 'i
	End If
	
	Set pobjFSO = Nothing

End Sub	'DetermineInstalledAddOns

'***********************************************************************************************

Sub DetermineDatabaseUpgrades(byRef objCnn, byRef aryAddons)
'Determine which database upgrades for the add-ons have been installed and which version

Dim pstrSQL
Dim pobjRS
Dim i, j
Dim paryFieldsToCheck
Dim pblnUpgraded
Dim pstrVersion
	
	On Error Resume Next

	Set	pobjRS = server.CreateObject("adodb.recordset")
	pobjRS.CursorLocation = 2 'adUseClient

	For i = 0 To UBound(aryAddons)
	
		pblnUpgraded = False
		pstrVersion = ""

		Select Case i
			Case enAO_AttributeExtender
				'attrURL added to sfAttributes with v2 
				pstrSQL = "Select odrattrsvdAttrText from sfSavedOrderAttributes Where odrattrsvdID=-1"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
					pstrVersion = ""
				Else
					Err.Clear
					pstrSQL = "Select attrDisplayStyle from sfAttributes Where attrID=-1"
					pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Err.number = 0 Then
						pblnUpgraded = True
						pstrVersion = ".1"
					Else
						pblnUpgraded = False
					End If
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_GiftCertificate
				'ssGCFreeText added with v1.1
				pstrSQL = "Select ssGCFreeText from ssGiftCertificates"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					Err.Clear
					pstrSQL = "Select ssGCID from ssGiftCertificates"
					pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Err.number = 0 Then
						pblnUpgraded = True
						pstrVersion = ".1"
					Else
						pblnUpgraded = False
					End If
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_PayPalPayments
				'PayPalToken added with PayPal Express
				pstrSQL = "Select PayPalToken from sfTmpOrderDetails"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					Err.Clear
					pstrSQL = "Select PayPalIPNID from PayPalIPNs"
					pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Err.number = 0 Then
						pblnUpgraded = True
						pstrVersion = ".1"
					Else
						pblnUpgraded = False
					End If
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_PostageRate
				pstrSQL = "Select ssShippingCarrierID from ssShippingCarriers"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_PricingLevel
				'clubExpDate added with MT 1.01.006 - taking shortcut by adding it here
				pstrSQL = "Select clubExpDate from sfCustomers"
				'pstrSQL = "Select PricingLevelID from sfCustomers"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_PromoMail
				pstrSQL = "Select ssPromoMailSent from sfCustomers"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_PromoMgr
				pstrSQL = "Select PromotionID from Promotions"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_PromoMgrII
				'FreeShippingCode added with v2
				pstrSQL = "Select FreeShippingCode from Promotions"
				pstrSQL = "Select MaxAllowableValue from Promotions"
				'ApplyToBasePrice added for MasterTemplate 11/12/2006
				pstrSQL = "Select NumUsesByCustomer from Promotions"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
					pstrVersion = ""
				Else
					Err.Clear
					pstrSQL = "Select PromotionID from Promotions"
					pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Err.number = 0 Then
						pblnUpgraded = True
						pstrVersion = ".1"
					Else
						pblnUpgraded = False
					End If
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_OrderMgr
				If Err.number <> 0 Then Err.Clear
				paryFieldsToCheck = Array("ssExportedShipping", "ssBackOrderTrackingNumber", "ssOrderFlagged", "ssOrderID")
				'ssVoid added with v2.01.001
				'ssBackOrderTrackingNumber added with v2.00.001
				For j = 0 To UBound(paryFieldsToCheck)
					pstrSQL = "Select " & paryFieldsToCheck(j) & " from ssOrderManager"
					pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Err.number = 0 Then Exit For
					Err.Clear
				Next 'j
				
				If j = 0 Then
					pblnUpgraded = True
				ElseIf j > UBound(paryFieldsToCheck) Then
					pblnUpgraded = False
				Else
					pblnUpgraded = True
					pstrVersion = "." & CStr(j)
				End If 
				If pobjRS.State = 1 Then pobjRS.Close
				If Err.number <> 0 Then Err.Clear
			Case enAO_WebStoreMgr
				If Err.number <> 0 Then Err.Clear
				paryFieldsToCheck = Array("failedLoginAttempts", "userName")
				'ssVoid added with v2.01.001
				'ssBackOrderTrackingNumber added with v2.00.001
				For j = 0 To UBound(paryFieldsToCheck)
					pstrSQL = "Select " & paryFieldsToCheck(j) & " from ssUsers"
					pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Err.number = 0 Then Exit For
					Err.Clear
				Next 'j
				
				If j = 0 Then
					pblnUpgraded = True
				ElseIf j > UBound(paryFieldsToCheck) Then
					pblnUpgraded = False
				Else
					pblnUpgraded = True
					pstrVersion = "." & CStr(j)
				End If 
				If pobjRS.State = 1 Then pobjRS.Close
				If Err.number <> 0 Then Err.Clear
			Case enAO_ZBS
				pstrSQL = "Select ZoneID from ssShipZones"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_DynamicProduct
				pstrSQL = "Select relatedProducts from sfProducts"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_ProductPlacement
				pstrSQL = "Select sortMfg from sfProducts"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_TaxRateManager
				pstrSQL = "Select TaxRateID from ssTaxTable Where TaxRateID=0"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_FileDownload
				pstrSQL = "Select prodMaxDownloads from sfProducts"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_CCV
				pstrSQL = "Select payCardCCV from sfCPayments"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_MasterTemplate
				pstrSQL = "Select ssTemplateVersion from sfAdmin Where ssTemplateVersion='" & ssTemplateVersion & "'"
				'pstrSQL = "Select adminMinOrderMessage from sfAdmin"
				'pstrSQL = "Select prodMinQty from sfProducts"
				'pstrSQL = "Select SumOfssGCRedemptionAmount from qryGiftCertificateRedemptionsByOrder"
				'pstrSQL = "Select prodEnableAlsoBought from sfProducts"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					If pobjRS.EOF Then pstrVersion = ".1"
					pblnUpgraded = True
				Else
					Err.Clear
					pstrSQL = "Select loclctryFraudRating from sfLocalesCountry"
					pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Err.number = 0 Then
						pblnUpgraded = True
						pstrVersion = ".1"
					Else
						pblnUpgraded = False
					End If
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_ContentManagement
				pstrSQL = "Select contentAuthorID from contentAuthors"
				pstrSQL = "Select contentSortOrder from content"
				pstrSQL = "Select contentProductAssignmentID from contentProductAssignments"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_VisitorTracking
				pstrSQL = "Select TypeKeyWordSearch from visitorPageViews"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_SQLSpeedUpgrade
				pstrSQL = "Select prodDescription from sfProducts"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					If pobjRS.Fields("prodDescription").Type = 203 Then
						pblnUpgraded = False
					Else
						pblnUpgraded = True
					End If
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_DatabaseSize
				pstrSQL = "Select prodAdditionalImages from sfProducts"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					If pobjRS.Fields("prodAdditionalImages").Type = 203 Then
						pblnUpgraded = False
					Else
						pblnUpgraded = True
					End If
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_SEtoAEUpgrade
				pstrSQL = "Select cpCouponCode from sfCoupons"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
			Case enAO_BuyersClub
				pstrSQL = "Select ssBuyersClubRedemptionID from ssBuyersClubRedemptions"
				pobjRS.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number = 0 Then
					pblnUpgraded = True
				Else
					pblnUpgraded = False
				End If
				pobjRS.Close
				Err.Clear
		End Select
		'debugprint "pblnUpgraded",pblnUpgraded
		'debugprint "pstrVersion",pstrVersion
		
		maryAddons(i)(enDBUpgraded) = pblnUpgraded
		maryAddons(i)(enAddonVersion) = pstrVersion
		
	Next 'i
	
	Set pobjRS = Nothing

End Sub	'DetermineDatabaseUpgrades

'***********************************************************************************************

Function isValidConnection(byRef objCnn)
'Tests/Establishes the connection to the database

Dim pblnValidConnection
Dim pstrDSN

On Error Resume Next

	pstrDSN = Application("DSN_NAME")
	If Len(pstrDSN) = 0 Then pstrDSN = Session("DSN_NAME")
	If len(pstrDSN) > 0 then
		Set objCnn = Server.CreateObject("ADODB.Connection")
		objCnn.Open pstrDSN
		If (objCnn.State = 1) Then
			pblnValidConnection = True
		Else
			pblnValidConnection = False
			mstrMessage = "<h4><Font Color='Red'>Could not connect to the database. Error: " & Err.number & " - " & Err.Description & "</FONT></H3>" _
						& "DSN: <em>" & pstrDSN & "</em><br />"
			Err.Clear
		End If
	Else
		pblnValidConnection = False
		mstrMessage = "<h4><Font Color='Red'>Could not connect to the database</FONT></H3>"
	End If
	
	isValidConnection = pblnValidConnection

End Function	'isValidConnection

'***********************************************************************************************

	Function OpenDatabase(objCnn, strFilePath, blnDSN)

	On Error Resume Next

		Set objCnn = Server.CreateObject("ADODB.Connection")
		If blnDSN Then
			objCnn.Open strFilePath
		Else
			objCnn.Open dbProvider & "Data Source=" & strFilePath & ";"
		End If
		If Err.number <> 0 Then
			debugprint Err.number,Err.Description
		End If
		OpenDatabase = (objCnn.State = 1)

	End Function	'OpenDatabase

'***********************************************************************************************

Function stepCounter(lngCounter)

	lngCounter = lngCounter + 1
	stepCounter = lngCounter
End Function	'stepCounter

'***********************************************************************************************

Sub WriteFormVariables

Dim pstrFormItem

	Response.Write "<hr><h4>Start Form Contents</h4>" & vbcrlf
	For Each pstrFormItem In Request.Form
		Response.Write "&nbsp;&nbsp;" & pstrFormItem & ": " & Request.Form(pstrFormItem) & "<BR>" & vbcrlf
	Next 'pstrFormItem
	Response.Write "<h4>Start QueryString Contents</h4>" & vbcrlf
	For Each pstrFormItem In Request.QueryString
		Response.Write "&nbsp;&nbsp;" & pstrFormItem & ": " & Request.QueryString(pstrFormItem) & "<BR>" & vbcrlf
	Next 'pstrFormItem
	Response.Write "<hr>" & vbcrlf
	
End Sub	'WriteFormVariables

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Generic DB Routines
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Sub setFieldDefinitionsForRemoval(byRef aryFieldDefinitions)
	'Need to convert ADDs to DROPs
	'Need to change ALTER from new back to original

	Dim plngFieldCounter
	Dim pstrAction

		For plngFieldCounter = 0 To UBound(aryFieldDefinitions)
			pstrAction = aryFieldDefinitions(plngFieldCounter)(0)
			Select Case LCase(aryFieldDefinitions(plngFieldCounter)(0))
				Case "add": aryFieldDefinitions(plngFieldCounter)(0) = "DROP"
				Case "alter": aryFieldDefinitions(plngFieldCounter)(2) = aryFieldDefinitions(plngFieldCounter)(3)
				Case Else
			End Select
		Next
		
	End Sub	'setFieldDefinitionsForRemoval

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Generic DB Upgrade
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function GenericDBUpgrade(byRef objCnn, byVal strTableName, byVal aryUpgradeInstructions, byRef strErrorMessage)
	Dim pstrSQL
	Dim pstrLocalError
	Dim pblnSuccess
	Dim i
	Dim pstrDBAction

	On Error Resume Next

		For i = 0 To UBound(aryUpgradeInstructions)
			pstrDBAction = UCase(Trim(aryUpgradeInstructions(i)(enUpdate_Action)))
			Select Case pstrDBAction
				Case "DROP"
					pstrSQL = "ALTER TABLE " & strTableName & " DROP " & aryUpgradeInstructions(i)(enUpdate_Field)
				Case "SQL"
					pstrSQL = aryUpgradeInstructions(i)(enUpdate_Type)
				Case Else
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & strTableName & " " & aryUpgradeInstructions(i)(enUpdate_Action) & " " & aryUpgradeInstructions(i)(enUpdate_Field) & " " & adjustSQLServerType(aryUpgradeInstructions(i)(enUpdate_Type))
						If LCase(aryUpgradeInstructions(i)(enUpdate_Type)) = "counter" Then pstrSQL = "SET IDENTITY_INSERT " & strTableName & " ON;" & pstrSQL
					Else
						pstrSQL = "ALTER TABLE " & strTableName & " " & aryUpgradeInstructions(i)(enUpdate_Action) & " COLUMN " & aryUpgradeInstructions(i)(enUpdate_Field) & " " & adjustSQLServerType(aryUpgradeInstructions(i)(enUpdate_Type))
					End If
			End Select
			
			objCnn.Execute pstrSQL,, 128
			
			If Err.number <> 0 Then
				If InStr(1, pstrSQL, "SET IDENTITY_INSERT") > 0 And InStr(1, err.Description, "IDENTITY_INSERT is already ON") > 1 Then 
					err.Clear
					pstrSQL = Replace(pstrSQL, "SET IDENTITY_INSERT " & strTableName & " ON;", "")
					objCnn.Execute pstrSQL,, 128

					If Err.number <> 0 Then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
														& "&nbsp;&nbsp;<font color=red>SQL (" & i & "):" & pstrSQL & "</font><BR>"
						Err.Clear
					End If
				Else
					pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
													& "&nbsp;&nbsp;<font color=red>SQL (" & i & "):" & pstrSQL & "</font><BR>"
					Err.Clear
				End If
			End If

		Next 'i
		
		strErrorMessage = pstrLocalError
		pblnSuccess = CBool(Len(pstrLocalError) = 0)

		GenericDBUpgrade = pblnSuccess
		
	End Function	'GenericDBUpgrade

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Generic CreateNewTable
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function CreateNewTable(byRef objCnn, byVal strTableName, byVal strFieldName, byRef strErrorMessage)
	
	Dim pstrSQL
	Dim pstrLocalError
	Dim pblnSuccess

	On Error Resume Next

		If Len(strFieldName) > 0 Then
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & strTableName & " (" & strFieldName & " int Identity, CONSTRAINT " & strTableName & "_pk Primary Key (" & strFieldName & "))"
			Else
				pstrSQL = "CREATE TABLE " & strTableName & " (" & strFieldName & " COUNTER, CONSTRAINT " & strTableName & "_pk Primary Key (" & strFieldName & "))"
			End If
		Else
				pstrSQL = "DROP TABLE " & strTableName
		End If
		objCnn.Execute pstrSQL,, 128
			
		If Err.number <> 0 Then
			pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
							& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
			Err.Clear
		End If

		strErrorMessage = pstrLocalError
		pblnSuccess = CBool(Len(pstrLocalError) = 0)

		CreateNewTable = pblnSuccess
		
	End Function	'CreateNewTable

'******************************************************************************************************************************************************************
'
'	Begin Main Page
'
'******************************************************************************************************************************************************************

Dim mobjCnn
Dim mblnUpgraded
Dim mblnError
Dim mstrMessage
Dim mblnValidConnection

Dim mblnSF5
Dim mblnSF5SE
Dim mblnSF5AE
Dim mblnSQLServer
Dim i
Dim maryAddons
Dim mstrUpgradeItem

	'Call WriteFormVariables

	mstrUpgradeItem = Request.QueryString("UpgradeItem")
	If Len(mstrUpgradeItem) = 0 Then mstrUpgradeItem = Request.Form("UpgradeItem")
	'hard code since all items are present
	mstrUpgradeItem = "MasterTemplate"

	If Len(Request.Form) = 0 Then
		'Determine StoreFront Version
		'StoreFront 5 sets an Application Variable - if it is empty, assume SF2k
		mblnSF5SE = CBool(Application("AppName") = "StoreFront")
		mblnSF5AE = CBool(Application("AppName") = "StoreFrontAE")
		mblnSF5 = (mblnSF5SE Or mblnSF5AE)

		'Determine database Version
		'StoreFront 5 sets an Application Variable - SF2k has to figure it out manually
		If mblnSF5 Then
			mblnSQLServer = CBool(Application("AppDatabase") <> "Access")
		Else
			mblnSQLServer = False
		End If
	Else
		mblnSQLServer = CBool((Len(Request.Form("SQLServer")) > 0))
		mblnSF5SE = CBool(Request.Form("SFVersion") = "SF5SE")
		mblnSF5AE = CBool(Request.Form("SFVersion") = "SF5AE")
		mblnSF5 = (mblnSF5SE Or mblnSF5AE)
	End If

	mblnError = False
	
	Call SetAvailableAddOns(maryAddons)
	Call DetermineInstalledAddOns(maryAddons)
	
	'Test/Establish the database connection
	mblnValidConnection = isValidConnection(mobjCnn)
	
	If mblnValidConnection Then
		Call DoUpgrades(mobjCnn, maryAddons)
		Call DetermineDatabaseUpgrades(mobjCnn, maryAddons)
	Else
	
	End If

%>
<!doctype HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta content="Microsoft FrontPage 4.0" name=GENERATOR>
<title>Sandshot Sofware Database Upgrade Utility</title>

<script language="javascript">
function showInstallationDetails(bytDetail, bytDisplay)
{
var paryDetails = new Array(<%= UBound(maryAddons) %>);
<% 
For i = 0 To UBound(maryAddons) 
	Response.Write "paryDetails[" & i & "] = " & Chr(34) & maryAddons(i)(enUpgradeDetails) & Chr(34) & ";" & vbcrlf
Next 'i
%>

	if (bytDisplay == 1)
	{
		document.all("detailPane").innerHTML = paryDetails[bytDetail];
	}else{
		return false;
		document.all("detailPane").innerHTML = "Select an add-on for details";
	}
	return false;
}


</script>
</HEAD>
<BODY>
<H2>Sandshot Sofware Database Upgrade Utility</H2>

<P>This utility upgrades the StoreFront database to use the various Sandshot Software add-ons listed below.

<h4>General Information</h4>
<ul>
  <li><b>Add-on</b> - this column lists the available add-on name. It is hyperlinked to the respective product information page where you can view the full product details.</li>
  <li><b>Files Installed</b> - this column indicates if the necessary add-on files have been installed.</li>
  <li><b>Database Upgraded</b> - this column indicates if the database has been upgraded to use this add-on. If the necessary add-on files are installed, you will be given the option to install or uninstall the upgrade.</li>
  <li><b>Installation Notes</b> - place your cursor over the product information link to view the details for the database upgrade action performed for this add-on</li>
</ul>

<h4>Instructions for use:</h4>
<OL>
  <LI>This file must be located in your active Storefront web</LI>           
  <LI>You must be running this file from your web browser</LI>
  <li>You must not have the StoreFront databas open (Access only)</li>
  <li>Select the desired options</li>
  <li>Accept the user agreement</li>
  <li>Press the <i>Upgrade Database</i> button</li>   
</OL>
  
<FORM name="frmDatabaseUpgrade" id="frmDatabaseUpgrade" action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="Post">
<input type="hidden" name="upgradeItem" id="upgradeItem" value="<%= mstrUpgradeItem %>">
<table cellpadding="2" cellspacing="0" border="1">
  <% If Len(mstrMessage) > 0 Then %>
  <tr>
    <td colspan="4" align="center">
      <table border="2" cellpadding="2" cellspacing="0">
        <tr><td><%= mstrMessage %></td></tr>
      </table>
    </td>
  </tr>
  <% End If %>
  <tr>
    <th>Add-on</th>
    <th>Files Installed</th>
    <th>Database Upgraded</th>
    <th>Installation Notes</th>
  </tr>
<% For i = 0 To UBound(maryAddons) 
	If (maryAddons(i)(enInstalled) = "True" OR maryAddons(i)(enInstalled) = "N/A") And Len(mstrUpgradeItem) > 0 Then
%>
  <tr>
    <td><a href="<%= maryAddons(i)(enAddonLink) %>" onmouseover="return showInstallationDetails('<%= i %>',1)" onmouseout="return showInstallationDetails('',0);" target="_blank"><%= maryAddons(i)(enName) %></a></td>
    <td><%= maryAddons(i)(enInstalled) %></td>
    <td>
    <%
		If maryAddons(i)(enInstalled) <> "False" Then
			If maryAddons(i)(enDBUpgraded) Then
				%><input type="checkbox" name="installAddon" ID="installAddon<%= i %>" value="<%= i %>.0">&nbsp;<label for="installAddon<%= i %>">Uninstall</label><%
				If Len(maryAddons(i)(enAddonVersion)) > 0 Then
					%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="installAddon" ID="installAddon<%= i %>U" value="<%= i %>.1<%= maryAddons(i)(enAddonVersion) %>">&nbsp;<label for="installAddon<%= i %>U">Upgrade to latest</label><%
				End If
			Else
				%><input type="checkbox" name="installAddon" ID="installAddon<%= i %>" value="<%= i %>.1">&nbsp;<label for="installAddon<%= i %>">Install</label><%
			End If
		Else
			Response.Write "N/A"
		End If
    %>
    </td>
    <% If i = 0 Then %>
    <td rowspan="<%= UBound(maryAddons) %>" id="detailPane">Select an add-on for details</td>
    <% End If %>
  </tr>
<%
	End If	'aryAddons(i)(enInstalled) And Len(mstrUpgradeItem) > 0 Then
   Next 'i %>
  <tr>
    <td colspan="3">
      <table border=1 cellpadding=2 cellspacing=0>
        <tr>
          <th colspan="2">Custom DB upgrade</th>
        </tr>
        <tr>
          <td colspan="2"><ul>
							<li>You can either add a field to an existing table or create a new one</li>
							<li>If you add a new table your first field MUST be the autonumber/increment</li>
							<li>Make sure you check the Install checkbox for the custom upgrade</li>
						  </ul>
		</td>
        </tr>
        <tr>
          <td>Use Existing Table:</td>
          <td><select name="sourceTable" id="sourceTable"><option value="">Select a table</option><%= getAvailableTables %></select></td>
        </tr>
        <tr>
          <td>Create new table</td>
          <td><input type="text" name="sourceTableNew" id="sourceTableNew" value=""></td>
        </tr>
        <tr>
          <td>Field: </td>
          <td><input type="text" name="sourceFieldNew" id="sourceFieldNew" value=""></td>
        </tr>
        <tr>
          <td>Action: </td>
          <td><select name="fieldAction" id="fieldAction">
          <option value="ADD" selected>Add</option>
          <option value="ALTER">Alter</option>
          <option value="DROP">Drop</option>
          
          </td>
        </tr>
        <tr>
          <td>Field Type:</td>
          <td>
			<select name="sourceFieldType" id="sourceFieldType">
							<option value="byte">Byte</option>
							<option value="long">Long Integer</option>
							<option value="date">Date</option>
							<option value="yesno">Yes/No</option>
							<option value="memo">Memo</option>
							<option value="counter">Counter</option>
							<option value="char">Char</option>
							<option value="varchar">VarChar</option>
						</select>
          </td>
        </tr>
        <tr>
          <td>Field Length:</td>
          <td><input type="text" name="sourceFieldLength" id="sourceFieldLength" value=""></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan="3">
	<% If mblnValidConnection Then %>
	<strong>StoreFront Version</strong><br>
	&nbsp;&nbsp;<input type="radio" name="SFVersion" ID="SFVersion5SE" <% If mblnSF5SE Then Response.Write "checked" %> value="SF5SE">&nbsp;<label for="SFVersion5SE">StoreFront 5.0 SE</label><br>
	&nbsp;&nbsp;<input type="radio" name="SFVersion" ID="SFVersion5AE" <% If mblnSF5AE Then Response.Write "checked" %> value="SF5AE">&nbsp;<label for="SFVersion5AE">StoreFront 5.0 AE</label><br>
	&nbsp;&nbsp;<input type="radio" name="SFVersion" ID="SFVersion2k" <% If Not (mblnSF5SE Or mblnSF5AE) Then Response.Write "checked" %> value="SF2k">&nbsp;<label for="SFVersion2k">StoreFront 2000</label><br>
	&nbsp;&nbsp;<input type=checkbox name="SQLSERVER" ID="SQLSERVER" <% If mblnSQLServer Then Response.Write "checked" %>>&nbsp;<label for="SQLSERVER">This is a SQL Server database</label><br>
<% On Error Goto 0 %>
	<% If maryAddons(enAO_TaxRateManager)(enDBUpgraded) Then Call ShowTaxRateImportOption %>
	<% End If %>
	<P>
	<input type="checkbox" name="AcceptAgreement" id="AcceptAgreement" value="Accept" onclick="if (this.checked){this.form.btnAction.disabled=false}else{this.form.btnAction.disabled=true}">&nbsp;
	<label for="AcceptAgreement"><b>Accept Agreement:</b> This utility is provided without warranty. While it has been 
successfully tested using the standard Storefront database, no guarantee regarding fitness for use in your application is 
made. Always make a backup of your database prior to making changes to it.</label></P>
	<INPUT type="submit" name="btnAction" id="btnAction" value="Upgrade Database" disabled><br>
    </td>
  </tr>
</table>
</FORM>
<p><a href="../../admin.asp" title="return to admin page">Return to admin</a></p>

</BODY></HTML>
<%

'***********************************************************************************************
'
'	ADD-ON UPGRADE FUNCTIONS
'
'***********************************************************************************************

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Attribute Extender
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_AttributeExtenderAddon(byRef objCnn, byVal blnInstall, ByVal bytUpgradeToLatest)

	Dim pstrSQL
	Dim pstrTableName
	Dim pstrFieldName
	Dim pstrTempMessage
	Dim pstrLocalError
	Dim pblnError

		On Error Resume Next

		If blnInstall then

			If bytUpgradeToLatest = 0 Then
			
				'------------------------------------------------------------------------------'
				' Update sfAttributes TABLE														   '
				'------------------------------------------------------------------------------'

				'Reset Settings
				pstrTempMessage = ""
				pblnError = False

				pstrTableName = "sfAttributes"
			
				ReDim paryFields(6)			
				paryFields(0) = Array("ADD","attrDisplayStyle","long")	'original
				paryFields(1) = Array("ADD","attrDisplayOrder","long")	'original
				paryFields(2) = Array("ADD","attrDisplay","memo")		'added with v2
				paryFields(3) = Array("ADD","attrImage","char(255)")	'added with v2
				paryFields(4) = Array("ADD","attrSKU","char(255)")		'added with v2
				paryFields(5) = Array("ADD","attrURL","char(255)")		'added with v2
				paryFields(6) = Array("ADD","attrExtra","char(255)")	'added with v2
				
				For i = 0 To UBound(paryFields)
					
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
										& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
			
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if
			
				'------------------------------------------------------------------------------'
				' Update sfAttributeDetail TABLE														   '
				'------------------------------------------------------------------------------'

				'Reset Settings
				pstrTempMessage = ""
				pblnError = False

				pstrTableName = "sfAttributeDetail"
			
				ReDim paryFields(8)			
				paryFields(0) = Array("ADD","attrdtDefault","yesno")
				paryFields(1) = Array("ADD","attrdtDisplay","memo")
				paryFields(2) = Array("ADD","attrdtFileName","char(255)")
				paryFields(3) = Array("ADD","attrdtImage","char(255)")
				paryFields(4) = Array("ADD","attrdtSKU","char(50)")
				paryFields(5) = Array("ADD","attrdtURL","char(255)")
				paryFields(6) = Array("ADD","attrdtWeight","long")
				paryFields(7) = Array("ADD","attrdtExtra","char(255)")
				paryFields(8) = Array("ADD","attrdtExtra1","char(255)")
				
				For i = 0 To UBound(paryFields)
					
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
										& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
			
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if

				'------------------------------------------------------------------------------'
				' Update sfTmpOrderAttributes TABLE														   '
				'------------------------------------------------------------------------------'

				'Reset Settings
				pstrTempMessage = ""
				pblnError = False

				pstrTableName = "sfTmpOrderAttributes"
			
				ReDim paryFields(0)			
				paryFields(0) = Array("ADD","odrattrtmpAttrText","memo")

				For i = 0 To UBound(paryFields)
					
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
										& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
			
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if

				'------------------------------------------------------------------------------'
				' Update sfSavedOrderAttributes TABLE														   '
				'------------------------------------------------------------------------------'

				'Reset Settings
				pstrTempMessage = ""
				pblnError = False

				pstrTableName = "sfSavedOrderAttributes"
			
				ReDim paryFields(0)			
				paryFields(0) = Array("ADD","odrattrsvdAttrText","memo")

				For i = 0 To UBound(paryFields)
					
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
										& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
			
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if

			ElseIf bytUpgradeToLatest = 1 Then
			
				'------------------------------------------------------------------------------'
				' Update sfAttributes TABLE														   '
				'------------------------------------------------------------------------------'

				'Reset Settings
				pstrTempMessage = ""
				pblnError = False

				pstrTableName = "sfAttributes"
			
				ReDim paryFields(4)			
				paryFields(0) = Array("ADD","attrDisplay","memo")		'added with v2
				paryFields(1) = Array("ADD","attrImage","char(255)")	'added with v2
				paryFields(2) = Array("ADD","attrSKU","char(255)")		'added with v2
				paryFields(3) = Array("ADD","attrURL","char(255)")		'added with v2
				paryFields(4) = Array("ADD","attrExtra","char(255)")	'added with v2
				
				For i = 0 To UBound(paryFields)
					
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
										& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
			
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if
			
				'------------------------------------------------------------------------------'
				' Update sfAttributeDetail TABLE														   '
				'------------------------------------------------------------------------------'

				'Reset Settings
				pstrTempMessage = ""
				pblnError = False

				pstrTableName = "sfAttributeDetail"
			
				ReDim paryFields(8)			
				paryFields(0) = Array("ADD","attrdtDefault","yesno")
				paryFields(1) = Array("ADD","attrdtDisplay","memo")
				paryFields(2) = Array("ADD","attrdtFileName","char(255)")
				paryFields(3) = Array("ADD","attrdtImage","char(255)")
				paryFields(4) = Array("ADD","attrdtSKU","char(50)")
				paryFields(5) = Array("ADD","attrdtURL","char(255)")
				paryFields(6) = Array("ADD","attrdtWeight","long")
				paryFields(7) = Array("ADD","attrdtExtra","char(255)")
				paryFields(8) = Array("ADD","attrdtExtra1","char(255)")
				
				For i = 0 To UBound(paryFields)
					
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
										& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
			
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if

				'------------------------------------------------------------------------------'
				' Update sfTmpOrderAttributes TABLE														   '
				'------------------------------------------------------------------------------'

				'Reset Settings
				pstrTempMessage = ""
				pstrLocalError = ""
				pblnError = False

				pstrTableName = "sfTmpOrderAttributes"
				ReDim paryFields(1)			
				paryFields(0) = Array("ADD","odrattrtmpAttrText","memo")
				paryFields(1) = Array("ALTER","odrattrtmpAttrID","long")

				For i = 0 To UBound(paryFields)
					
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
										& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
			
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if

			End If	'upgrade version check
			
			If Not mblnError Then
				mstrMessage = mstrMessage & "<H4>You can edit your products <a href='../../sfProductAdmin.asp'>here</a>.</H4>"			
			End If

			'------------------------------------------------------------------------------'
			' Update sfSavedOrderAttributes TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrTempMessage = ""
			pstrLocalError = ""
			pblnError = False

			pstrTableName = "sfSavedOrderAttributes"
		
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","odrattrsvdAttrText","memo")

			For i = 0 To UBound(paryFields)
				
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				End If
				objCnn.Execute pstrSQL,, 128
				
				if Err.number <> 0 then
					pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
									& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
					Err.Clear
				End If

			Next 'i
		
			if Len(pstrLocalError) = 0 then
				mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
				mblnUpgraded = True
			else
				mblnError = True
				mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
			end if

		Else
		
			On Error Resume Next	'Necessary due to number of versions which may be used

			'------------------------------------------------------------------------------'
			' Remove sfAttributes TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrTempMessage = ""
			pblnError = False

			pstrTableName = "sfTmpOrderAttributes"
		
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","odrattrtmpAttrText","memo")
			
			For i = 0 To UBound(paryFields)
				pstrSQL = "ALTER TABLE " & pstrTableName & " DROP COLUMN " & paryFields(i)(1)
				objCnn.Execute pstrSQL,, 128
				If Err.number = 0 Then
					pstrTempMessage = pstrTempMessage & "<li>" & paryFields(i)(1) & " successfully removed.</li>"
				Else
					pblnError = True
					pstrTempMessage = pstrTempMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & "</FONT></h4>"	
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & paryFields(i)(1) & ": " & Err.description &"</FONT></li>"
					Err.Clear
				end if
			Next 'i

			mblnError = pblnError
			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
			
			'------------------------------------------------------------------------------'
			' Remove sfAttributes TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrTempMessage = ""
			pblnError = False

			pstrTableName = "sfAttributes"
			ReDim paryFields(6)			
			paryFields(0) = Array("ADD","attrDisplayStyle","long")	'original
			paryFields(1) = Array("ADD","attrDisplayOrder","long")	'original
			paryFields(2) = Array("ADD","attrDisplay","memo")		'added with v2
			paryFields(3) = Array("ADD","attrImage","char(255)")	'added with v2
			paryFields(4) = Array("ADD","attrSKU","char(255)")		'added with v2
			paryFields(5) = Array("ADD","attrURL","char(255)")		'added with v2
			paryFields(6) = Array("ADD","attrExtra","char(255)")	'added with v2
			
			For i = 0 To UBound(paryFields)
				pstrSQL = "ALTER TABLE " & pstrTableName & " DROP COLUMN " & paryFields(i)(1)
				objCnn.Execute pstrSQL,, 128
				If Err.number = 0 Then
					pstrTempMessage = pstrTempMessage & "<li>" & paryFields(i)(1) & " successfully removed.</li>"
				Else
					pblnError = True
					pstrTempMessage = pstrTempMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & "</FONT></h4>"	
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & paryFields(i)(1) & ": " & Err.description &"</FONT></li>"
					Err.Clear
				end if
			Next 'i

			mblnError = pblnError
			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
			
			'------------------------------------------------------------------------------'
			' Remove sfAttributeDetail TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrTempMessage = ""
			pblnError = False

			pstrTableName = "sfAttributeDetail"
			ReDim paryFields(8)			
			paryFields(0) = Array("ADD","attrdtDefault","yesno")
			paryFields(1) = Array("ADD","attrdtDisplay","memo")
			paryFields(2) = Array("ADD","attrdtFileName","char(255)")
			paryFields(3) = Array("ADD","attrdtImage","char(255)")
			paryFields(4) = Array("ADD","attrdtSKU","char(50)")
			paryFields(5) = Array("ADD","attrdtURL","char(255)")
			paryFields(6) = Array("ADD","attrdtWeight","long")
			paryFields(7) = Array("ADD","attrdtExtra","char(255)")
			paryFields(8) = Array("ADD","attrdtExtra1","char(255)")
			
			For i = 0 To UBound(paryFields)
				pstrSQL = "ALTER TABLE " & pstrTableName & " DROP COLUMN " & paryFields(i)(1)
				objCnn.Execute pstrSQL,, 128
				If Err.number = 0 Then
					pstrTempMessage = pstrTempMessage & "<li>" & paryFields(i)(1) & " successfully removed.</li>"
				Else
					pblnError = True
					pstrTempMessage = pstrTempMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & "</FONT></h4>"	
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & paryFields(i)(1) & ": " & Err.description &"</FONT></li>"
					Err.Clear
				end if
			Next 'i

			mblnError = pblnError
			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
			
		End If

	End Function	'Install_AttributeExtenderAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Buyer's Club
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_BuyersClub(byRef objCnn, byVal blnInstall)

	Dim pstrTableName
	Dim pstrTempMessage
	Dim pstrLocalErrorMessage
	Dim pblnSuccess
	Dim plngTableCounter
	Dim plngRecordCounter
	Dim paryRecordInsertions

	'Define the upgrades
	
	Dim parydbUpgrades(2)
	'contains array of table name, tableID (for new tables only), array of fieldDefinitions, array of records to insert (optional)
	'fieldDefinitions is array of action, fieldName, ACCESS specific field type, original field type
	
	'ssBuyersClubRedemptions
	ReDim paryFields(4)			
		paryFields(0) = Array("ADD","ssBuyersClubRedemptionCustID","long","")
		paryFields(1) = Array("ADD","ssBuyersClubRedemptionPoints","long","")
		paryFields(2) = Array("ADD","ssBuyersClubRedemptionCertificateID","long","")
		paryFields(3) = Array("ADD","ssBuyersClubRedemptionDate","datetime","")
		paryFields(4) = Array("ADD","ssBuyersClubRedemptionNotes","char(255)","")
	parydbUpgrades(0) = Array("ssBuyersClubRedemptions", "ssBuyersClubRedemptionID", paryFields)

	'sfProducts
	ReDim paryFields(1)			
		paryFields(0) = Array("ADD","buyersClubPointValue","double","")
		paryFields(1) = Array("ADD","buyersClubIsPercentage","long","")
	ReDim paryRecordInsertions(0)
			paryRecordInsertions(0) = "Update sfProducts Set buyersClubPointValue=0.03, buyersClubIsPercentage=1"
	parydbUpgrades(1) = Array("sfProducts", "", paryFields, paryRecordInsertions)

	'sfOrderDetails
	ReDim paryFields(0)			
		paryFields(0) = Array("ADD","buyersClubPointsIssued","double","")

	'Note: this section will back-date point awards - just set the date below (2x)
	If True Then
		parydbUpgrades(2) = Array("sfOrderDetails", "", paryFields)
	Else
		ReDim paryRecordInsertions(0)
			If mblnSQLServer Then
				paryRecordInsertions(0) = "Update sfOrderDetails Set buyersClubPointsIssued = odrdtSubTotal" _
										& " Where odrdtOrderId In (SELECT orderID FROM sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId WHERE (sfOrders.orderDate>'4/1/2005' And orderIsComplete=1))"
			Else
				paryRecordInsertions(0) = "Update sfOrderDetails Set buyersClubPointsIssued = 0.03 * CDbl(odrdtSubTotal)" _
										& " Where odrdtOrderId In (SELECT orderID FROM sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId WHERE (sfOrders.orderDate>#4/1/2005# And orderIsComplete=1))"
			End If

		parydbUpgrades(2) = Array("sfOrderDetails", "", paryFields, paryRecordInsertions)
	End If

	On Error Resume Next

		pblnSuccess = True
		
		If blnInstall then
		
			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				'------------------------------------------------------------------------------'
				' Create table
				'------------------------------------------------------------------------------'
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(1), pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				End If	'new table check
				
				'------------------------------------------------------------------------------'
				' Add/alter field definitions
				'------------------------------------------------------------------------------'
				
				pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

				'Intermediate error checking
				if Len(pstrLocalErrorMessage) = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
				else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
				end if
				
				'Now check for any records to insert
				If UBound(parydbUpgrades(plngTableCounter)) >= 3 Then
					If Err.Number <> 0 Then err.Clear
				
					paryRecordInsertions = parydbUpgrades(plngTableCounter)(3)
					For plngRecordCounter = 0 To UBound(paryRecordInsertions)
						objCnn.Execute paryRecordInsertions(plngRecordCounter),,128
					Next 'plngRecordCounter

					'Intermediate error checking
					If Err.Number = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully populated</B></li><BR>"
					Else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error populating " & pstrTableName & "</FONT></li>" _
														  & "<li>--Error " & Err.Number & ": " & Err.Description & "</li>"
					End If

				End If

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_BuyersClub)(enName) & " database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error adding " & aryAddons(enAO_BuyersClub)(enName) & " database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		Else

			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					'------------------------------------------------------------------------------'
					' Remove table
					'------------------------------------------------------------------------------'
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "", pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully deleted</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error deleting " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				Else
					'------------------------------------------------------------------------------'
					' Undo Add/alter field definitions
					'------------------------------------------------------------------------------'
					Call setFieldDefinitionsForRemoval(parydbUpgrades(plngTableCounter)(2))
					
					pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table changes successfully removed</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error undoing " & pstrTableName & " table changes</FONT></li>" & pstrLocalErrorMessage
					end if
				
				End If	'new table check
				

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_BuyersClub)(enName) & " database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error removing " & aryAddons(enAO_BuyersClub)(enName) & " database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		
		Install_BuyersClub = pblnSuccess

	End Function	'Install_BuyersClub
	
'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	CCV
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_CCV(byRef objCnn, byVal blnInstall)

	Dim pstrTableName
	Dim pstrTempMessage
	Dim pstrLocalErrorMessage
	Dim pblnSuccess

	On Error Resume Next

		pblnSuccess = True
		
		If blnInstall then

			'------------------------------------------------------------------------------'
			' Update sfCPayments TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfCPayments"
			
			ReDim paryFields(2)			
			paryFields(0) = Array("ADD","payCardCCV","char(10)")
			paryFields(1) = Array("ADD","payCardIssueNumber","char(10)")
			paryFields(2) = Array("ADD","payCardStartDate","char(10)")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if
			
			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>CCV database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error adding CCV database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		Else

			'------------------------------------------------------------------------------'
			' REMOVE sfCPayments TABLE UPGRADES														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "sfCPayments"
			
			ReDim paryFields(2)			
			paryFields(0) = Array("DROP","payCardCCV")
			paryFields(1) = Array("DROP","payCardIssueNumber")
			paryFields(2) = Array("DROP","payCardStartDate")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Updates Successfully removed</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing CCV database modifications: <br />" & pstrLocalErrorMessage & "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>CCV database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error removing CCV database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_CCV = pblnSuccess

	End Function	'Install_CCV
	
'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Content Management
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_ContentManagement(byRef objCnn, byVal blnInstall)

	Dim pstrTableName
	Dim pstrTempMessage
	Dim pstrLocalErrorMessage
	Dim pblnSuccess
	Dim plngTableCounter
	Dim plngRecordCounter
	Dim paryRecordInsertions

	'Define the upgrades
	
	Dim parydbUpgrades(6)
	'contains array of table name, tableID (for new tables only), array of fieldDefinitions, array of records to insert (optional)
	'fieldDefinitions is array of action, fieldName, ACCESS specific field type, original field type
	
	'contentAuthors
	ReDim paryFields(2)			
		paryFields(0) = Array("ADD","contentAuthorName","char(255)","")
		paryFields(1) = Array("ADD","contentAuthorEmail","char(255)","")
		paryFields(2) = Array("ADD","contentAuthorURL","char(255)","")
	parydbUpgrades(0) = Array("contentAuthors", "contentAuthorID", paryFields)

	'contentCategoryAssignments
	ReDim paryFields(1)			
		paryFields(0) = Array("ADD","contentCategoryAssignmentCategoryID","long","")
		paryFields(1) = Array("ADD","contentCategoryAssignmentContentID","long","")
	parydbUpgrades(1) = Array("contentCategoryAssignments", "contentCategoryAssignmentID", paryFields)

	'contentFoundUseful
	ReDim paryFields(1)			
		paryFields(0) = Array("ADD","contentFoundUsefulContentID","long","")
		paryFields(1) = Array("ADD","contentFoundUsefulScore","long","")
	parydbUpgrades(2) = Array("contentFoundUseful", "contentFoundUsefulID", paryFields)

	'contentTypes
	ReDim paryFields(5)			
		paryFields(0) = Array("ADD","contentTypeName","char(255)","")
		paryFields(1) = Array("ADD","contentTypeDisplayName","char(255)","")
		paryFields(2) = Array("ADD","contentTypeURL","char(255)","")
		paryFields(3) = Array("ADD","contentTypeDescription","char(100)","")
		paryFields(4) = Array("ADD","contentTypeDisplayInSiteMap","yesno","")
		paryFields(5) = Array("ADD","contentTypeDisplayOrder","long","")

						'If LCase(aryUpgradeInstructions(i)(enUpdate_Type)) = "counter" Then pstrSQL = "SET IDENTITY_INSERT " & strTableName & " ON;" & pstrSQL

		ReDim paryRecordInsertions(7)
		pstrTableName = "contentTypes"
		Call ResetArrayIndex
		paryRecordInsertions(ArrayIndex) = "INSERT INTO contentTypes (contentTypeID, contentTypeName, contentTypeDisplayName, contentTypeURL) Values (1, 'SEO - Product', 'Search by Product', 'viewProducts.asp')"
		paryRecordInsertions(ArrayIndex) = "INSERT INTO contentTypes (contentTypeID, contentTypeName, contentTypeDisplayName, contentTypeURL) Values (2, 'SEO - Category', 'Search by Category', 'viewCategories.asp')"
		paryRecordInsertions(ArrayIndex) = "INSERT INTO contentTypes (contentTypeID, contentTypeName, contentTypeDisplayName, contentTypeURL) Values (3, 'SEO - Manufacturer', 'Search by Manufacturer', 'viewManufacturers.asp')"
		paryRecordInsertions(ArrayIndex) = "INSERT INTO contentTypes (contentTypeID, contentTypeName, contentTypeDisplayName, contentTypeURL) Values (4, 'SEO - Vendor', 'Search by Vendor', 'viewVendors.asp')"
		paryRecordInsertions(ArrayIndex) = "INSERT INTO contentTypes (contentTypeID, contentTypeName, contentTypeDisplayName, contentTypeURL) Values (5, 'Product Review', 'Product Reviews', 'productReviews.asp')"
		paryRecordInsertions(ArrayIndex) = "INSERT INTO contentTypes (contentTypeID, contentTypeName, contentTypeDisplayName, contentTypeURL) Values (6, 'Frequently Asked Questions (F.A.Q.)', 'F.A.Q.s', 'faq.asp')"
		paryRecordInsertions(ArrayIndex) = "INSERT INTO contentTypes (contentTypeID, contentTypeName, contentTypeDisplayName, contentTypeURL) Values (7, 'Content Category', '', 'siteContent.asp')"
		paryRecordInsertions(ArrayIndex) = "INSERT INTO contentTypes (contentTypeID, contentTypeName, contentTypeDisplayName, contentTypeURL) Values (8, 'Content Pages', '', 'siteContent.asp')"
		
		If mblnSQLServer Then
			For i = 0 To UBound(paryRecordInsertions)
				paryRecordInsertions(i) = "SET IDENTITY_INSERT " & pstrTableName & " ON;" & paryRecordInsertions(i)
			Next 'i
		End If
	parydbUpgrades(3) = Array("contentTypes", "contentTypeID", paryFields, paryRecordInsertions)

	'content
	ReDim paryFields(22)			
		paryFields(0) = Array("ADD","contentAuthorID","long","")
		paryFields(1) = Array("ADD","contentContentType","long","")
		paryFields(2) = Array("ADD","contentReferenceID","long","")
		paryFields(3) = Array("ADD","contentApprovedForDisplay","byte","")
		paryFields(4) = Array("ADD","contentTitle","char(100)","")
		paryFields(5) = Array("ADD","contentAbstract","memo","")
		paryFields(6) = Array("ADD","contentContent","memo","")
		paryFields(7) = Array("ADD","contentContentFilePath","char(100)","")
		paryFields(8) = Array("ADD","contentAuthorName","char(100)","")
		paryFields(9) = Array("ADD","contentAuthorEmail","char(100)","")
		paryFields(10) = Array("ADD","contentAuthorShowEmail","yesno","")
		paryFields(11) = Array("ADD","contentAuthorRating","long","")
		paryFields(12) = Array("ADD","contentDateCreated","date","")
		paryFields(13) = Array("ADD","contentDateModified","date","")
		paryFields(14) = Array("ADD","contentTemplatePage","char(50)","")
		paryFields(15) = Array("ADD","contentPageName","char(50)","")
		paryFields(16) = Array("ADD","contentPageTitle","char(100)","")
		paryFields(17) = Array("ADD","contentMetaDescription","char(255)","")
		paryFields(18) = Array("ADD","contentMetaKeywords","char(100)","")
		paryFields(19) = Array("ADD","contentMetaAuthor","char(100)","")
		paryFields(20) = Array("ADD","contentMetaCustom1","char(100)","")
		paryFields(21) = Array("ADD","contentMetaCustom2","char(100)","")
		paryFields(22) = Array("ADD","contentSortOrder","long","")
	parydbUpgrades(4) = Array("content", "contentID", paryFields)

	'sfProducts
	ReDim paryFields(0)			
		paryFields(0) = Array("ADD","sfProductID","counter","")
	parydbUpgrades(5) = Array("sfProducts", "", paryFields)

	'contentProductAssignments
	ReDim paryFields(1)			
		paryFields(0) = Array("ADD","contentProductAssignmentProductID","long","")
		paryFields(1) = Array("ADD","contentProductAssignmentContentID","long","")
	parydbUpgrades(6) = Array("contentProductAssignments", "contentProductAssignmentID", paryFields)

	On Error Resume Next

		pblnSuccess = True
		
		If blnInstall then
		
			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				'------------------------------------------------------------------------------'
				' Create table
				'------------------------------------------------------------------------------'
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(1), pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				End If	'new table check
				
				'------------------------------------------------------------------------------'
				' Add/alter field definitions
				'------------------------------------------------------------------------------'
				
				pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

				'Intermediate error checking
				if Len(pstrLocalErrorMessage) = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
				else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
				end if
				
				'Now check for any records to insert
				If UBound(parydbUpgrades(plngTableCounter)) >= 3 Then
					If Err.Number <> 0 Then err.Clear
				
					paryRecordInsertions = parydbUpgrades(plngTableCounter)(3)
					For plngRecordCounter = 0 To UBound(paryRecordInsertions)
						objCnn.Execute paryRecordInsertions(plngRecordCounter),,128
						'Intermediate error checking
						If Err.Number = 0 then
							pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully populated</B></li><BR>"
						Else
							pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error populating " & pstrTableName & "</FONT></li>" _
															& "--Error " & Err.Number & ": " & Err.Description & "<br/>SQL: " & paryRecordInsertions(plngRecordCounter) & "</li>"
							Err.Clear
						End If
					Next 'plngRecordCounter

				End If

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_ContentManagement)(enName) & " database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error adding " & aryAddons(enAO_ContentManagement)(enName) & " database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		Else

			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					'------------------------------------------------------------------------------'
					' Remove table
					'------------------------------------------------------------------------------'
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "", pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully deleted</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error deleting " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				Else
					'------------------------------------------------------------------------------'
					' Undo Add/alter field definitions
					'------------------------------------------------------------------------------'
					Call setFieldDefinitionsForRemoval(parydbUpgrades(plngTableCounter)(2))
					
					pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table changes successfully removed</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error undoing " & pstrTableName & " table changes</FONT></li>" & pstrLocalErrorMessage
					end if
				
				End If	'new table check
				

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_ContentManagement)(enName) & " database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error removing " & aryAddons(enAO_ContentManagement)(enName) & " database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_ContentManagement = pblnSuccess

	End Function	'Install_ContentManagement

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Dynamic Product
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_DynamicProductAddon(byRef objCnn, byVal blnInstall)

	Dim pstrSQL
	Dim pstrTableName
	Dim pstrFieldName
	Dim pstrTempMessage
	Dim pblnError

		On Error Resume Next

		If blnInstall then

			'------------------------------------------------------------------------------'
			' Update sfProducts TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrTempMessage = ""
			pblnError = False

			pstrTableName = "sfProducts"
			pstrFieldName = "relatedProducts"
			If mblnSQLServer Then
				pstrSQL = "ALTER TABLE " & pstrTableName & " ADD " & pstrFieldName & " text"
			Else
				pstrSQL = "ALTER TABLE " & pstrTableName & " ADD " & pstrFieldName & " memo"
			End If
			objCnn.Execute pstrSQL,, 128
			If Err.number = 0 Then
				pstrTempMessage = pstrTempMessage & "<li>" & pstrFieldName & " successfully added.</li>"
			Else
				pblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & "</FONT></h4>"	
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrFieldName & ": " & Err.description &"</FONT></li>"
				Err.Clear
			end if

			mblnError = pblnError
			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
			
		Else

			'------------------------------------------------------------------------------'
			' Update sfProducts TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrTempMessage = ""
			pblnError = False

			pstrTableName = "sfProducts"
			pstrFieldName = "relatedProducts"
			pstrSQL = "ALTER TABLE " & pstrTableName & " DROP COLUMN " & pstrFieldName
			objCnn.Execute pstrSQL,, 128
			If Err.number = 0 Then
				pstrTempMessage = pstrTempMessage & "<li>" & pstrFieldName & " successfully removed.</li>"
			Else
				pblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & "</FONT></h4>"	
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrFieldName & ": " & Err.description &"</FONT></li>"
				Err.Clear
			end if

			mblnError = pblnError
			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
			
		End If

	End Function	'Install_DynamicProductAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	File Download
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_FileDownloadAddon(byRef objCnn, byVal blnInstall)

	Dim pstrTableName
	Dim pstrTempMessage
	Dim pstrLocalErrorMessage
	Dim pblnSuccess

	On Error Resume Next

		pblnSuccess = True
		
		If blnInstall then

			'------------------------------------------------------------------------------'
			' Create ssFileDownloads TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "ssFileDownloads"
			
			pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "ssFileDownloadID", pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			ReDim paryFields(7)			
			paryFields(0) = Array("ADD","ssFileDownloadInitiated","date")
			paryFields(1) = Array("ADD","ssFileDownloadCompleted","date")
			paryFields(2) = Array("ADD","ssFileDownloadOrderItemID","long")
			paryFields(3) = Array("ADD","ssFileDownloadREMOTE_ADDR","char(255)")
			paryFields(4) = Array("ADD","ssFileDownloadREMOTE_HOST","char(255)")
			paryFields(5) = Array("ADD","ssFileDownloadDNSHost","char(255)")
			paryFields(6) = Array("ADD","ssFileDownloadedFileName","char(255)")
			paryFields(7) = Array("ADD","ssFileDownloadedVersion","char(50)")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfOrderDetails TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfOrderDetails"
			
			ReDim paryFields(2)			
			paryFields(0) = Array("ADD","odrdtDownloadExpiresOn","date")
			paryFields(1) = Array("ADD","odrdtMaxDownloads","long")
			paryFields(2) = Array("ADD","odrdtDownloadAuthorized","yesno")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfProducts TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfProducts"
			
			ReDim paryFields(1)			
			paryFields(0) = Array("ADD","prodMaxDownloads","long")
			paryFields(1) = Array("ADD","prodDownloadValidFor","long")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfAttributeDetail TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfAttributeDetail"
			
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","attrdtFileName","char(255)")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if
			
			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>File Download database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error adding File Download database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		Else

			'------------------------------------------------------------------------------'
			' Create ssFileDownloads TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "ssFileDownloads"
			
			pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "", pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully removed</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing File Download database modifications: <br />" & pstrLocalErrorMessage & "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE sfOrderDetails TABLE UPGRADES														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "sfOrderDetails"
			
			ReDim paryFields(2)			
			paryFields(0) = Array("DROP","odrdtDownloadExpiresOn")
			paryFields(1) = Array("DROP","odrdtMaxDownloads")
			paryFields(2) = Array("DROP","odrdtDownloadAuthorized")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Updates Successfully removed</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing File Download database modifications: <br />" & pstrLocalErrorMessage & "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE sfProducts TABLE UPGRADES														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "sfProducts"
			
			ReDim paryFields(1)			
			paryFields(0) = Array("DROP","prodMaxDownloads")
			paryFields(1) = Array("DROP","prodDownloadValidFor")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Updates Successfully removed</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing File Download database modifications: <br />" & pstrLocalErrorMessage & "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE sfAttributeDetail TABLE UPGRADES														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "sfAttributeDetail"
			
			ReDim paryFields(0)			
			paryFields(0) = Array("DROP","attrdtFileName")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Updates Successfully removed</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing File Download database modifications: <br />" & pstrLocalErrorMessage & "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>File Download database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error removing File Download database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_FileDownloadAddon = pblnSuccess

	End Function	'Install_FileDownloadAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Gift Certificate Manager
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_GiftCertificateManagerAddon(byRef objCnn, byVal blnInstall, byVal bytUpgradeToLatest)

	Dim pstrSQL
	Dim pstrTableName

	On Error Resume Next
	
		If blnInstall then

			If bytUpgradeToLatest = 0 Then
				'------------------------------------------------------------------------------'
				' ADD ssGiftCertificates TABLE														   '
				'------------------------------------------------------------------------------'

				pstrTableName = "ssGiftCertificates"
				If mblnSQLServer Then
					pstrSQL = "CREATE TABLE " & pstrTableName & " " _
							& "(ssGCID int Identity," _
							& " ssGCCode char (20) PRIMARY KEY," _
							& " ssGCExpiresOn Datetime," _
							& " ssGCSingleUse tinyint," _
							& " ssGCElectronic tinyint," _
							& " ssGCCustomerID int," _
							& " ssGCIssuedToEmail char (255)," _
							& " ssGCCreatedOn Datetime," _
							& " ssGCModifiedOn Datetime" _
							& " )"
				Else
					pstrSQL = "CREATE TABLE " & pstrTableName & " " _
							& "(ssGCID COUNTER," _
							& " ssGCCode char (20) PRIMARY KEY," _
							& " ssGCExpiresOn Date," _
							& " ssGCSingleUse YESNO," _
							& " ssGCElectronic YESNO," _
							& " ssGCCustomerID long," _
							& " ssGCIssuedToEmail char (255)," _
							& " ssGCCreatedOn Date," _
							& " ssGCModifiedOn Date" _
							& " )"
				End If

				objCnn.Execute pstrSQL,, 128

				if Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully updated.</h4>"
				else
					mblnError = True
					mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
					Err.Clear
				end if

				'------------------------------------------------------------------------------'
				' ADD ssGiftCertificateRedemptions TABLE														   '
				'------------------------------------------------------------------------------'

				pstrTableName = "ssGiftCertificateRedemptions"
				If mblnSQLServer Then
					pstrSQL = "CREATE TABLE " & pstrTableName & " " _
							& "(ssGCRedemptionID int Identity PRIMARY KEY," _
							& " ssGCRedemptionCGCode char (20)," _
							& " ssGCRedemptionAmount Decimal (10,2)," _
							& " ssGCRedemptionOrderID int," _
							& " ssGCRedemptionInternalNotes text," _
							& " ssGCRedemptionExternalNotes text," _
							& " ssGCRedemptionType int," _
							& " ssGCRedemptionCreatedOn Datetime," _
							& " ssGCRedemptionActive tinyint" _
							& " )"
				Else
					pstrSQL = "CREATE TABLE " & pstrTableName & " " _
							& "(ssGCRedemptionID COUNTER PRIMARY KEY," _
							& " ssGCRedemptionCGCode char (20)," _
							& " ssGCRedemptionAmount double," _
							& " ssGCRedemptionOrderID long," _
							& " ssGCRedemptionInternalNotes memo," _
							& " ssGCRedemptionExternalNotes memo," _
							& " ssGCRedemptionType long," _
							& " ssGCRedemptionCreatedOn Date," _
							& " ssGCRedemptionActive YESNO" _
							& " )"
				End If

				objCnn.Execute pstrSQL,, 128

				if Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully updated.</h4>"
				else
					mblnError = True
					mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
					Err.Clear
				end if

			ElseIf bytUpgradeToLatest = 1 Then
			
				'------------------------------------------------------------------------------'
				' Update ssGiftCertificates TABLE for additional fields						   '
				'------------------------------------------------------------------------------'

				pstrTableName = "ssGiftCertificates"
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD ssGCFreeText text"
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD COLUMN ssGCFreeText memo"
				End If
				objCnn.Execute pstrSQL,, 128
				If Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & ", Field ssGCFreeText successfully added.</B></LI>"
				Else
					mblnError = True
					mstrMessage = mstrMessage & "<LI><Font Color='Red'><B>Table " & pstrTableName & ", Field ssGCFreeText not added.</B></FONT><br>" & Err.Description & "</LI>"
					Err.Clear
				End If

				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD ssGCToName text"
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD COLUMN ssGCToName memo"
				End If
				objCnn.Execute pstrSQL,, 128
				If Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & ", Field ssGCToName successfully added.</B></LI>"
				Else
					mblnError = True
					mstrMessage = mstrMessage & "<LI><Font Color='Red'><B>Table " & pstrTableName & ", Field ssGCToName not added.</B></FONT><br>" & Err.Description & "</LI>"
					Err.Clear
				End If
				
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD ssGCFromName text"
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD COLUMN ssGCFromName memo"
				End If
				objCnn.Execute pstrSQL,, 128
				If Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & ", Field ssGCFromName successfully added.</B></LI>"
				Else
					mblnError = True
					mstrMessage = mstrMessage & "<LI><Font Color='Red'><B>Table " & pstrTableName & ", Field ssGCFromName not added.</B></FONT><br>" & Err.Description & "</LI>"
					Err.Clear
				End If
				
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD ssGCFromEmail text"
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD COLUMN ssGCFromEmail memo"
				End If
				objCnn.Execute pstrSQL,, 128
				If Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & ", Field ssGCFromEmail successfully added.</B></LI>"
				Else
					mblnError = True
					mstrMessage = mstrMessage & "<LI><Font Color='Red'><B>Table " & pstrTableName & ", Field ssGCFromEmail not added.</B></FONT><br>" & Err.Description & "</LI>"
					Err.Clear
				End If
				
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD ssGCMessage text"
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " ADD COLUMN ssGCMessage memo"
				End If
				objCnn.Execute pstrSQL,, 128
				If Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & ", Field ssGCMessage successfully added.</B></LI>"
				Else
					mblnError = True
					mstrMessage = mstrMessage & "<LI><Font Color='Red'><B>Table " & pstrTableName & ", Field ssGCMessage not added.</B></FONT><br>" & Err.Description & "</LI>"
					Err.Clear
				End If
			Else
				Response.Write "<h4>No Upgrade Instructions given.</h4>"
			End If	'bytUpgradeToLatest

		Else

			'------------------------------------------------------------------------------'
			' REMOVE ssGiftCertificates TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssGiftCertificates"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE ssGiftCertificateRedemptions TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssGiftCertificateRedemptions"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

		End If

	End Function	'Install_GiftCertificateManagerAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Master Template
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_MasterTemplate(byRef objCnn, byVal blnInstall)

	Dim pstrTableName
	Dim pstrTempMessage
	Dim pstrLocalErrorMessage
	Dim pblnSuccess

	On Error Resume Next

		pblnSuccess = True
		
		If blnInstall then

			'------------------------------------------------------------------------------'
			' Create ssConfigurationSettings TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "ssConfigurationSettings"
			
			ReDim paryFields(63)			
			Call ResetArrayIndex
			If mblnSQLServer Then
				paryFields(ArrayIndex) = Array("SQL","Create ssConfigurationSettings TABLE","CREATE TABLE " & pstrTableName & " (configID int Identity, storeID int DEFAULT 1, configName varchar(50), CONSTRAINT " & pstrTableName & "_pk Primary Key (configName,storeID))")
			Else
				paryFields(ArrayIndex) = Array("SQL","Create ssConfigurationSettings TABLE","CREATE TABLE " & pstrTableName & " (configID COUNTER, storeID long DEFAULT 1, configName varchar(50), CONSTRAINT " & pstrTableName & "_pk Primary Key (configName,storeID))")
			End If

			paryFields(ArrayIndex) = Array("ADD","configCategory","varchar(50)")
			paryFields(ArrayIndex) = Array("ADD","configTitle","varchar(50)")
			paryFields(ArrayIndex) = Array("ADD","configValue","varchar(255)")
			paryFields(ArrayIndex) = Array("ADD","configDescription","varchar(255)")
			paryFields(ArrayIndex) = Array("ADD","configDataType","long")
			paryFields(ArrayIndex) = Array("ADD","configDisplayType","varchar(25)")
			paryFields(ArrayIndex) = Array("ADD","configUserDefined","yesno")
			paryFields(ArrayIndex) = Array("ADD","configOptions","memo")
			
			'paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,0)")
			'Buyer's Club
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Buyer''s Club','BuyersClubEarningsMultiple','Earnings Multiple','1','points earned = multiple x product sell price x product multiple',1,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Buyer''s Club','BuyersClubRedemptionMultiple','Redemption Multiple','1','certificate value = multiple x points redeemed',1,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Buyer''s Club','BuyersClubMinimumRedemption','Minimum Certificate Value','5','Minimum number of points which can be redeemed for the certificate',1,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Buyer''s Club','BuyersClubCertificateMultiple','Certificate Multiples','5','Multiples available for certificates (in cents)',1,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Buyer''s Club','BuyersClubEnabled','Enabled','5','Enable Buyer''s Club display in myAccount',4,'checkbox',0)")
			
			'New Products Page
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'New Products Page','NewProductsDaysSinceAdded','Show New Products Added','30','Number of days to use for new products',1,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'New Products Page','NewProductField','New Products Field','prodDateModified','Field to use for date comparison - prodDateAdded and prodDateModified are the usual',0,'textbox',0)")

			'Site Settings
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Site Settings','TimeoutRedirectPage','Timeout Redirect Page','order.asp','Page customers are sent to if their login expires during checkout',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Settings','TrackPageViews','Track Page Views','0','Enable detailed tracking of page views. <em>Note: This may result in a large database.</em>',4,'checkbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Settings','PrimaryEmailToSendErrorTo','Email To Send Errors To','support@sandshot.net','Email address to send error messages to.',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Settings','SubWebPath','SubWeb Path','/mastertemplate','If you are running the site in a subweb you must set this value. Note: This must be lowercase',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Settings','NotifyMeEnabled','Enable Notify Me','0','Enable customers to request being notified by email if inventory is out of stock',4,'checkbox',0)")

			'Google Settings
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Google Settings','GoogleAnalytics','Google Analytics ID','','Google Analytics ID',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Google Settings','GoogleAdwords','Google AdWords ID','','Google AdWords ID',0,'textbox',0)")

			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Overture Settings','OvertureID','Overture ID','','Overture ID',0,'textbox',0)")
			'Checkout
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','CartDisplay_ProductID','Product ID Prefix','','Text to display in front of product code. Leave empty to not display product code.',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','CartDisplay_MfgName','Manufacturer Prefix','','Text to display in front of manufacturer name. Leave empty to not display manufacturer.',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','POTerm','PO Term','PO','Term to use for POs',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','ECheckTerm','eCheck Term','eCheck','Term to use for eChecks',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','CODTerm','COD Term','COD','Term to use for CODs',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','PhoneFaxTerm','PhoneFax Term','PhoneFax','Term to use for PhoneFaxs',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined,configOptions) Values (1,'Checkout Settings','OrderViewImageSrc','Display Image In Cart','8','Display image in cart',0,'select',0,'8,Small;,None')")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','IncludeEmailVerification','Require Email Verification','1','Require customer to type in verification email address.',4,'checkbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','DisableLogin','DisableLogin','1','Check to hide returning customer login.',4,'checkbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','CCV_SaveToDB','CCV - Save To DB','1','Check to save in database, otherwise info sent in merchant email. Note: Please make sure you comply with your merchant account agreement. Email processing only.',4,'checkbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','CCV_Optional','CCV - Optional','1','Make CCV optional.',4,'checkbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Checkout Settings','CCVFieldName','CCVFieldName','payCardCCV','Field name in sfCPayments table to store CCV information. Note: Leave empty to not collect. You must have a value for this even if you do not save the data.',0,'textbox',0)")

			'Display Settings
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','BodyStyle','Body Style','','Style or other attributes which is added by default to the body tag.',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','PageWidth','Page Width','100%','Width of page.',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','PageBackground','Page Background','','Background color',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','FormFieldStyle','Form Field Style','','Style to apply to form fields',0,'textbox',0)")

			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','TopNavStyle','Top Nav Style','','id and style to apply to top navigation section',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','LeftNavStyle','Left Nav Style','','id and style to apply to left navigation section',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','RightNavStyle','Right Nav Style','','id and style to apply to right navigation section',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','BottomNavStyle','Bottom Nav Style','','id and style to apply to bottom navigation section',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','ContentStyle','Content Style','','id and style to apply to content section',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Site Design','ExpandCategoriesByDefault','Expand Categories','1','Check to display all categories by default.',4,'checkbox',0)")

			'Search Result Settings
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined,configOptions) Values (1," _
								   & "'Search Result Settings','SearchResultsDisplayType','Display Type','1','Display format for search results',1,'select',0,'0,Standard;1,Search Grid;2,Tabular')")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined,configOptions) Values (1," _
								   & "'Search Result Settings','SearchResultsDisplayTypeFixed','Display Type - User Customizable','','Set to float to allow the user to toggle between the standard and grid views',0,'select',0,',Float;0,Standard;1,Search Grid;2,Tabular')")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','DefaultPageSize','Default Page Size','12','Records Per Page',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','DefaultMaxRecords','Default Max Records','0','Maximum amount of records returned, 0 is no maximum',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','HighlightSearchTermClass','Highlight Search Term Class','highlightSearch','Used to highlight exact matches for search terms. Set to style sheet class if desired. Default is highlightSearch',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','MaxLengthDescription','Max Description Length','255','Maximum number of characters to display in search results. This only applies to tabular search results when displaying description as well.<br />-1 - do not show; 0 - show all; xxx - max length',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','ShowSearchCustomizationOption','Show Search Customization Bar','1','Display Search Customization Bar. Includes paging and sorting options',4,'checkbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','ShowOptionToTurnOffImages','Show Option to Turn Off Images','1','Enable user to toggle images on/off',4,'checkbox',0)")

			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','DisplayCategoryDescriptions','Display Category Descriptions','1','',4,'checkbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','SearchAttributes','Search Attributes','1','Check this option to include attribute text in search',4,'checkbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Search Result Settings','ImageDetailInstructions','Image Detail Instructions','<br />Click image for details','Text to appear in standard view immediately after the image.',0,'textbox',0)")
'53 so far
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Search Result Settings','URLTemplate_Manufacturer','Manufacturer URL Template','search_results_manufacturer.asp?txtsearchParamMan={ID}','URL template to use for automatic manufacturer links.',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Search Result Settings','URLTemplate_Vendor','Vendor URL Template','search_results_manufacturer.asp?txtsearchParamMan={ID}','URL template to use for automatic vendor links.',0,'textbox',0)")

			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Detail Page Settings','LargeWindowImageField','Large Window Image Field','','Field in sfProducts table which is used for extra large image. Leave blank not to use.',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Detail Page Settings','TextLinkToLargerWindow','Text Link To Larger Window','<br>Click to see larger image','Text to appear in immediately after the image.',0,'textbox',0)")

			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Quick Product Entry Page Settings','NumEntries','Number of Order Lines','10','Set to the number of order lines you desire on the page.',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Quick Product Entry Page Settings','DisplayProductIDs','Display Product IDs','1','Check to display a dropdown of product IDs instead of a textbox. Note: this can substantially increase page load time if you have a large catalog.',4,'checkbox',0)")

			'Downloadable Product Settings
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1,'Downloadable Product Settings','DownloadRootLocation_Default','Download Folder Location','{root}Downloads\','Location of download files. Use {root} to use application root',0,'textbox',0)")
			paryFields(ArrayIndex) = Array("SQL","Configuration","Insert Into " & pstrTableName & " (storeID,configCategory,configName,configTitle,configValue,configDescription,configDataType,configDisplayType,configUserDefined) Values (1," _
								   & "'Downloadable Product Settings','Download_CheckForAuthorization','Check For Authorization','0','Check to require individual authorization of downloads.',4,'checkbox',0)")

			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Create notifyMe TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "notifyMe"
			
			pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "notifyMeID", pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			ReDim paryFields(10)			
			paryFields(0) = Array("ADD","notifyCustID","long")
			paryFields(1) = Array("ADD","notifyProdID","varchar(50)")
			paryFields(2) = Array("ADD","notifyStoreID","long")
			paryFields(3) = Array("ADD","notifyLastName","varchar(50)")
			paryFields(4) = Array("ADD","notifyFirstName","varchar(50)")
			paryFields(5) = Array("ADD","notifyEmail","varchar(100)")
			paryFields(6) = Array("ADD","notifyType","varchar(50)")
			paryFields(7) = Array("ADD","notifyDateCreated","date")
			paryFields(8) = Array("ADD","notifyDateNotified","date")
			paryFields(9) = Array("ADD","notifyNotifyCount","long")
			paryFields(10) = Array("ADD","notifyInventoryID","long")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Create notifyMeAttributes TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "notifyMeAttributes"
			
			pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "notifyMeAttributeID", pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			ReDim paryFields(1)			
			paryFields(0) = Array("ADD","notifyMeID","long")
			paryFields(1) = Array("ADD","AttrdtID","long")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Create visitors TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "visitors"
			
			pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "visitorID", pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			ReDim paryFields(23)			
			paryFields(0) = Array("ADD","visitorSessionID","long")
			paryFields(1) = Array("ADD","visitorCustomerID","long")
			paryFields(2) = Array("ADD","visitorDateCreated","date")
			paryFields(3) = Array("ADD","visitorLastVisited","date")
			paryFields(4) = Array("ADD","vistorDiscountCodes","varchar(255)")
			paryFields(5) = Array("ADD","visitorySelectedFreeProducts","varchar(255)")
			paryFields(6) = Array("ADD","visitorCertificateCodes","varchar(255)")
			paryFields(7) = Array("ADD","visitorLastSearch","varchar(255)")
			paryFields(8) = Array("ADD","visitorRecentlyViewedProducts","varchar(255)")
			paryFields(9) = Array("ADD","visitorCity","varchar(50)")
			paryFields(10) = Array("ADD","visitorState","varchar(3)")
			paryFields(11) = Array("ADD","visitorZIP","varchar(10)")
			paryFields(12) = Array("ADD","visitorCountry","varchar(3)")
			paryFields(13) = Array("ADD","visitor_REFERER","long")
			paryFields(14) = Array("ADD","vistor_HTTP_REFERER","varchar(255)")
			paryFields(15) = Array("ADD","visitor_REMOTE_ADDR","varchar(255)")
			paryFields(16) = Array("ADD","visitorPreferredCurrency","varchar(3)")
			paryFields(17) = Array("ADD","visitorPreferredShippingCode","varchar(65)")
			paryFields(18) = Array("ADD","visitorLoggedInCustomerID","long")
			paryFields(19) = Array("ADD","visitorShipAddressID","long")
			paryFields(20) = Array("ADD","visitorOrderID","long")
			paryFields(21) = Array("ADD","visitorInstructions","varchar(255)")
			paryFields(22) = Array("ADD","visitorPaymentmethod","varchar(20)")
			paryFields(23) = Array("ADD","visitorEstimatedShipping","varchar(10)")
			
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update ssShippingMethods TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "ssShippingMethods"
			
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","ssShippingMethodIsSpecial","YESNO")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if
			
			'------------------------------------------------------------------------------'
			' Update sfManufacturers TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfManufacturers"
			
			ReDim paryFields(5)			
			paryFields(0) = Array("ADD","mfgMetaTitle","char(100)")
			paryFields(1) = Array("ADD","mfgMetaDescription","char(100)")
			paryFields(2) = Array("ADD","mfgMetaKeywords","char(100)")
			paryFields(3) = Array("ADD","mfgDescription","MEMO")
			paryFields(4) = Array("ADD","mfgIsActive","YESNO")
			paryFields(5) = Array("SQL","mfgIsActive","Update sfManufacturers Set mfgIsActive=1")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if
			
			'------------------------------------------------------------------------------'
			' Update sfVendors TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfVendors"
			
			ReDim paryFields(5)			
			paryFields(0) = Array("ADD","vendMetaTitle","char(100)")
			paryFields(1) = Array("ADD","vendMetaDescription","char(100)")
			paryFields(2) = Array("ADD","vendMetaKeywords","char(100)")
			paryFields(3) = Array("ADD","vendDescription","MEMO")
			paryFields(4) = Array("ADD","vendIsActive","YESNO")
			paryFields(5) = Array("SQL","vendIsActive","Update sfVendors Set vendIsActive=1")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if
			
			'------------------------------------------------------------------------------'
			' Update sfAdmin TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfAdmin"
			
			ReDim paryFields(14)			
			paryFields(0) = Array("ADD","adminTechnicalEmail","char(255)")
			paryFields(1) = Array("ADD","adminTermsAndConditions","MEMO")
			paryFields(2) = Array("ADD","adminTermsAndConditionsIsactive","YESNO")
			paryFields(3) = Array("ADD","adminMinOrderAmount","char(20)")
			paryFields(4) = Array("ADD","adminGlobalConfirmationMessage","memo")
			paryFields(5) = Array("ADD","adminGlobalConfirmationMessageIsactive","YESNO")
			paryFields(6) = Array("ADD","adminMinOrderMessage","memo")
			paryFields(7) = Array("ADD","ssTemplateVersion","char(10)")

			paryFields(8) = Array("ADD","hoursBetweenCleanings","long")
			paryFields(9) = Array("ADD","daysToSaveIncompleteOrders","long")
			paryFields(10) = Array("ADD","daysToSaveTempOrders","long")
			paryFields(11) = Array("ADD","daysToSaveSavedOrders","long")
			paryFields(12) = Array("ADD","daysToKeepVisitors","long")
			paryFields(13) = Array("ADD","CCReplace","char(20)")
			paryFields(14) = Array("ADD","adminOriginState","char(3)")

			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if
			
			objCnn.Execute "Update sfAdmin Set ssTemplateVersion='" & ssTemplateVersion & "'",,128

			'------------------------------------------------------------------------------'
			' Update sfOrders TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfOrders"
			
			ReDim paryFields(1)			
			paryFields(0) = Array("ADD","orderStoreID","long")
			paryFields(1) = Array("ALTER","orderGrandTotal","varchar(50)")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfTransactionResponse TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfTransactionResponse"
			
			ReDim paryFields(1)			
			paryFields(0) = Array("ADD","trnsrspAuthorizationAmount","char(50)")
			paryFields(1) = Array("ADD","trnsrspCCV2","char(50)")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfProducts TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfProducts"
			
			ReDim paryFields(33)			
			paryFields(0) = Array("ADD","prodAdditionalImages","MEMO")
			paryFields(1) = Array("ADD","prodEnableReviews","YESNO")
			paryFields(2) = Array("ADD","prodEnableAlsoBought","YESNO")
			paryFields(3) = Array("ADD","prodHandlingFee","varchar(20)")
			paryFields(4) = Array("ADD","prodSetupFee","varchar(20)")
			
			paryFields(5) = Array("SQL","qryGiftCertificateRedemptionsByOrder","")
			paryFields(5)(2) = "CREATE VIEW qryGiftCertificateRedemptionsByOrder AS " _
							& "SELECT Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount) AS SumOfssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionOrderID" _
							& " FROM ssGiftCertificateRedemptions" _
							& " GROUP BY ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
							& " HAVING ((ssGiftCertificateRedemptions.ssGCRedemptionType=1) AND (ssGiftCertificateRedemptions.ssGCRedemptionActive<>0))"
			paryFields(6) = Array("SQL","prodHandlingFee","Update sfProducts Set prodHandlingFee=0 Where prodHandlingFee is Null")
			paryFields(7) = Array("SQL","prodSetupFee","Update sfProducts Set prodSetupFee=0 Where prodSetupFee is Null")

			paryFields(8) = Array("SQL","CDOSYS Mail Method","Update sfSelectValues Set slctvalMailMethod='CDOSYS' Where slctvalID=9")
			paryFields(9) = Array("SQL","writeHTML Mail Method","Update sfSelectValues Set slctvalMailMethod='writeHTML' Where slctvalID=10")
			paryFields(10) = Array("ADD","prodMinQty","double")
			paryFields(11) = Array("ADD","prodIncrement","double")
			paryFields(12) = Array("ADD","prodLimitQtyToMTP","YESNO")
			paryFields(13) = Array("SQL","prodLimitQtyToMTP","Update sfProducts Set prodLimitQtyToMTP=0 Where prodLimitQtyToMTP Is Not Null")

			paryFields(14) = Array("ADD","prodFixedShippingCharge","double")
			paryFields(15) = Array("SQL","prodFixedShippingCharge","Update sfProducts Set prodFixedShippingCharge=0 Where prodFixedShippingCharge Is Not Null")
			paryFields(16) = Array("ADD","prodSpecialShippingMethods","varchar(20)")
			paryFields(17) = Array("ADD","prodDisplayAdditionalImagesInWindow","byte")
			paryFields(18) = Array("ADD","pageName","varchar(100)")
			paryFields(19) = Array("ADD","metaTitle","varchar(100)")
			paryFields(20) = Array("ADD","metaDescription","varchar(255)")
			paryFields(21) = Array("ADD","metaKeywords","varchar(100)")
			paryFields(22) = Array("ADD","version","varchar(100)")
			paryFields(23) = Array("ADD","releaseDate","DATETIME")
			paryFields(24) = Array("ADD","InstallationHours","varchar(1)")
			paryFields(25) = Array("ADD","MyProduct","YESNO")
			paryFields(26) = Array("ADD","InstallationRequired","YESNO")
			paryFields(27) = Array("ADD","IncludeInSearch","YESNO")
			paryFields(28) = Array("ADD","IncludeInRandomProduct","YESNO")
			paryFields(29) = Array("ADD","UpgradeVersion","varchar(255)")
			paryFields(30) = Array("ADD","packageCodes","varchar(255)")
			paryFields(31) = Array("SQL","IncludeInSearch","Update sfProducts Set IncludeInSearch=1 Where IncludeInSearch<>0")
			paryFields(32) = Array("SQL","IncludeInSearch","Update sfProducts Set IncludeInRandomProduct=1 Where IncludeInRandomProduct<>0")
			paryFields(33) = Array("ADD","prodSetupFeeOneTime","varchar(20)")

			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfSub_Categories TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfSub_Categories"
			
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","subcatHttpAdd","char(255)")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update ssShippingMethods TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "ssShippingMethods"
			
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","ssShippingMethodCountryRule","long")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfCategories TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfCategories"
			
			ReDim paryFields(1)			
			paryFields(0) = Array("ALTER","catHttpAdd","char(255)")
			paryFields(1) = Array("ALTER","catDescription","char(255)")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfLocalesCountry TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfLocalesCountry"
			
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","loclctryFraudRating","long")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Create fraudEmail TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "fraudEmails"
			
			pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "fraudEmailID", pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			ReDim paryFields(1)			
			paryFields(0) = Array("ADD","fraudEmail","char(255)")
			paryFields(1) = Array("ADD","fraudScore","long")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Create fraudIPAddresses TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "fraudIPAddresses"
			
			pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "fraudEmailID", pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			ReDim paryFields(1)			
			paryFields(0) = Array("ADD","fraudIPAddress","char(255)")
			paryFields(1) = Array("ADD","fraudScore","long")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if
			
			'------------------------------------------------------------------------------'
			' Update tables to support fractional ordering													   '
			'------------------------------------------------------------------------------'
			pstrTableName = "sfOrderDetails"
			ReDim paryFields(0): paryFields(0) = Array("ALTER","odrdtQuantity","double")
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			pstrTableName = "sfSavedOrderDetails"
			ReDim paryFields(0): paryFields(0) = Array("ALTER","odrdtsvdQuantity","double")
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			pstrTableName = "sfTmpOrderDetails"
			ReDim paryFields(0): paryFields(0) = Array("ALTER","odrdttmpQuantity","double")
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>File Download database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error adding File Download database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		Else

			'------------------------------------------------------------------------------'
			' Create ssFileDownloads TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "ssFileDownloads"
			
			pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "", pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully removed</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing File Download database modifications: <br />" & pstrLocalErrorMessage & "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE sfOrders TABLE UPGRADES														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "sfOrders"
			
			ReDim paryFields(0)			
			paryFields(0) = Array("DROP","orderStoreID")
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)
			
			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Updates Successfully removed</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing File Download database modifications: <br />" & pstrLocalErrorMessage & "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>File Download database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error removing File Download database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_MasterTemplate = pblnSuccess

	End Function	'Install_MasterTemplate

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Order Manager
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_OrderMgrAddon(byRef objCnn, byVal blnInstall, bytUpgradeToLatest)

	Dim pblnError
	Dim pstrSQL
	Dim pstrTableName
	Dim pstrLocalError

	On Error Resume Next
	
		If blnInstall then

			'------------------------------------------------------------------------------'
			' ADD ssOrderManager TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssOrderManager"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " (ssorderID int PRIMARY KEY)"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " (ssorderID long PRIMARY KEY)"
			End If
			objCnn.Execute pstrSQL,, 128

			if Err.number <> 0 then
				pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
								& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
				Err.Clear
			End If

			If Len(pstrLocalError) = 0 Then
				mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully added</B></H3><BR>"
				mblnUpgraded = True
			Else
				pblnError = True
				mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
			End If

			'------------------------------------------------------------------------------'
			' Update ssOrderManager TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrLocalError = ""
			pblnError = False

			pstrTableName = "ssOrderManager"
		
			ReDim paryFields(18)			
			paryFields(0) = Array("ADD","ssExternalNotes","memo")				'original
			paryFields(1) = Array("ADD","ssInternalNotes","memo")				'original
			paryFields(2) = Array("ADD","ssDatePaymentReceived","Date")			'original
			paryFields(3) = Array("ADD","ssPaidVia","char (50)")				'original
			paryFields(4) = Array("ADD","ssDateOrderShipped","Date")			'original
			paryFields(5) = Array("ADD","ssDateEmailSent","Date")				'original
			paryFields(6) = Array("ADD","ssShippedVia","long")					'original
			paryFields(7) = Array("ADD","ssTrackingNumber","char (50)")			'original
			paryFields(8) = Array("ADD","ssOrderStatus","long")					'original
			paryFields(9) = Array("ADD","ssExported","long")					'added with v1
			paryFields(10) = Array("ADD","ssOrderFlagged","long")				'added with v1
			paryFields(11) = Array("ADD","ssBackOrderDateNotified","Date")		'added with v1
			paryFields(12) = Array("ADD","ssBackOrderDateExpected","Date")		'added with v1
			paryFields(13) = Array("ADD","ssBackOrderMessage","memo")			'added with v1
			paryFields(14) = Array("ADD","ssBackOrderInternalMessage","memo")	'added with v1
			paryFields(15) = Array("ADD","ssBackOrderTrackingNumber","memo")	'added with v2
			paryFields(16) = Array("ADD","ssInternalOrderStatus","long")		'added with v3
			paryFields(17) = Array("ADD","ssExportedPayment","long")			'added with v3
			paryFields(18) = Array("ADD","ssExportedShipping","long")			'added with v3

			For i = 0 To UBound(paryFields)
				
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				End If
				objCnn.Execute pstrSQL,, 128
				
				if Err.number <> 0 then
					pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
									& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
					Err.Clear
				End If

			Next 'i
		
			if Len(pstrLocalError) = 0 then
				mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
				mblnUpgraded = True
			else
				pblnError = True
				mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
			end if

					
			'------------------------------------------------------------------------------'
			' Update sfOrders TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrLocalError = ""
			pblnError = False

			pstrTableName = "sfOrders"
		
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","orderVoided","YESNO")		'added with v2
			
			For i = 0 To UBound(paryFields)
				
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				End If
				objCnn.Execute pstrSQL,, 128
				
				if Err.number <> 0 then
					pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
									& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
					Err.Clear
				End If

			Next 'i
		
			if Len(pstrLocalError) = 0 then
				mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
				mblnUpgraded = True
			else
				pblnError = True
				mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
			end if

		Else

			'------------------------------------------------------------------------------'
			' REMOVE ssOrderManager TABLE												   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssOrderManager"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute pstrSQL,, 128
			if Err.number = 0 then
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

		End If

		mblnError = mblnError & blnError

	End Function	'Install_OrderMgrAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	PayPal Payments
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_PayPalPaymentsAddon(byRef objCnn, byVal blnInstall)

	Dim pstrSQL
	Dim pstrTableName

	On Error Resume Next
	
		If blnInstall then

			'------------------------------------------------------------------------------'
			' ADD ExpressCheckout TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssExpressCheckout"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(ExpressCheckoutID int Identity PRIMARY KEY," _
						& " SessionID int,"_
						& " ExpressCheckoutType int,"_
						& " token char (50)"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(ExpressCheckoutID COUNTER PRIMARY KEY," _
						& " SessionID long,"_
						& " ExpressCheckoutType long,"_
						& " token char (50))"
			End If

			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				mblnUpgraded = True
				mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully added.</h4>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

			'------------------------------------------------------------------------------'
			' Update sfTmpOrderDetails TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrLocalError = ""
			pblnError = False

			pstrTableName = "sfTmpOrderDetails"
		
			ReDim paryFields(0)			
			paryFields(0) = Array("ADD","PayPalToken","char (50)")		'added with v2
			
			For i = 0 To UBound(paryFields)
				
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				End If
				objCnn.Execute pstrSQL,, 128
				
				if Err.number <> 0 then
					pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
									& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
					Err.Clear
				End If

			Next 'i
		
			if Len(pstrLocalError) = 0 then
				mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
				mblnUpgraded = True
			else
				pblnError = True
				mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
			end if


			'------------------------------------------------------------------------------'
			' ADD PayPalIPNs TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "PayPalIPNs"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(PayPalIPNID int Identity PRIMARY KEY," _
						& " txn_id char (50),"_
						& " payment_status char (50),"_
						& " pending_reason char (50),"_
						& " DateIPNReceived Datetime)"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(PayPalIPNID COUNTER PRIMARY KEY," _
						& " txn_id char (50),"_
						& " payment_status char (50),"_
						& " pending_reason char (50),"_
						& " DateIPNReceived Date)"
			End If

			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				mblnUpgraded = True
				mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully added.</h4>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

			'------------------------------------------------------------------------------'
			' ADD PayPalPayments TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "PayPalPayments"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(txn_id char (50) UNIQUE,"_
						& " receiver_email text,"_
						& " item_name char (100),"_
						& " item_number char (50),"_
						& " quantity Decimal (10,2),"_
						& " invoice char (50),"_
						& " custom text,"_
						& " payment_status char (50),"_
						& " pending_reason char (50),"_
						& " payment_date Datetime,"_
						& " payment_gross Decimal (10,2),"_
						& " payment_fee Decimal (10,2),"_
						& " txn_type char (50),"_
						& " first_name char (50),"_
						& " last_name char (50),"_
						& " address_street char (50),"_
						& " address_city char (50),"_
						& " address_state char (50),"_
						& " address_zip char (50),"_
						& " address_country char (50),"_
						& " address_status char (50),"_
						& " payer_email text,"_
						& " payer_status char (50),"_
						& " payment_type char (50),"_
						& " notify_version char (5),"_
						& " verify_sign text,"_
						& " Category int,"_
						& " ActionCompleted_Completed Datetime,"_
						& " ActionCompleted_Pending Datetime,"_
						& " ActionCompleted_Failed Datetime,"_
						& " ActionCompleted_Denied Datetime)"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(txn_id char (50) UNIQUE,"_
						& " receiver_email memo,"_
						& " item_name char (100),"_
						& " item_number char (50),"_
						& " quantity double,"_
						& " invoice char (50),"_
						& " custom memo,"_
						& " payment_status char (50),"_
						& " pending_reason char (50),"_
						& " payment_date Date,"_
						& " payment_gross double,"_
						& " payment_fee double,"_
						& " txn_type char (50),"_
						& " first_name char (50),"_
						& " last_name char (50),"_
						& " address_street char (50),"_
						& " address_city char (50),"_
						& " address_state char (50),"_
						& " address_zip char (50),"_
						& " address_country char (50),"_
						& " address_status char (50),"_
						& " payer_email memo,"_
						& " payer_status char (50),"_
						& " payment_type char (50),"_
						& " notify_version char (5),"_
						& " verify_sign memo,"_
						& " Category int,"_
						& " ActionCompleted_Completed Date,"_
						& " ActionCompleted_Pending Date,"_
						& " ActionCompleted_Failed Date,"_
						& " ActionCompleted_Denied Date)"
			End If

			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				mblnUpgraded = True
				mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully added.</h4>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

			'------------------------------------------------------------------------------'
			' ADD PayPal Transaction Type												   '
			'------------------------------------------------------------------------------'

			Dim rsUpgrade

			pstrSQL = "Select * from sfTransactionTypes where transType='PayPal'"
			Set rsUpgrade = Server.CreateObject("ADODB.RECORDSET")
			rsUpgrade.Open pstrSQL,objCnn
			
			If rsUpgrade.EOF Then
				pstrSQL = "Insert Into sfTransactionTypes (transType, transName,transIsActive) Values ('PayPal','PayPal','1')"
				objCnn.Execute pstrSQL,, 128

				if Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<h4><B>PayPal successfully added as an available transaction method</B></h4>"
				else
					mblnError = True
					mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error adding PayPal Transaction Method: " & Err.description	& "</FONT></h4>"	
					Err.Clear
				end if

			Else
				mblnError = True
				mstrMessage = mstrMessage & "<H4><B>PayPal already exists as an available transaction method</B></H4><BR>"
			End If
			rsUpgrade.Close
			Set rsUpgrade = Nothing

			'------------------------------------------------------------------------------'
			' ADD PayPal Transaction Type												   '
			'------------------------------------------------------------------------------'

			pstrSQL = "Select * from sfTransactionMethods where trnsmthdName='PayPal WebPayments Pro'"
			Set rsUpgrade = Server.CreateObject("ADODB.RECORDSET")
			rsUpgrade.Open pstrSQL,objCnn
			
			If rsUpgrade.EOF Then
				pstrSQL = "Insert Into sfTransactionMethods (trnsmthdName) Values ('PayPal WebPayments Pro')"
				objCnn.Execute pstrSQL,, 128

				if Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<h4><B>PayPal WebPayments Pro successfully added as an available transaction method</B></h4>"
				else
					mblnError = True
					mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error adding PayPal WebPayments Pro Transaction Method: " & Err.description	& "</FONT></h4>"	
					Err.Clear
				end if

			Else
				mblnError = True
				mstrMessage = mstrMessage & "<H4><B>PayPal WebPayments Pro already exists as an available transaction method</B></H4><BR>"
			End If
			rsUpgrade.Close
			Set rsUpgrade = Nothing

		Else

			'------------------------------------------------------------------------------'
			' REMOVE PayPalIPNs TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "PayPalIPNs"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE PayPalPayments TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "PayPalPayments"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

		End If

	End Function	'Install_PayPalPaymentsAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Pricing Level
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_PricingLevelAddon(byRef objCnn, byVal blnInstall)

	Dim pstrSQL
	Dim pstrTableName
	Dim pstrFieldName
	Dim pstrTempMessage
	Dim pblnError

	On Error Resume Next

		If blnInstall then

			'------------------------------------------------------------------------------'
			' ADD PricingLevels TABLE													   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrTempMessage = ""
			pblnError = False

			pstrTableName = "PricingLevels"
			pstrFieldName = "attrDisplayStyle"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(PricingLevelID int Identity PRIMARY KEY," _
						& " PricingLevelName char (20),"_
						& " PricingLevelNotes char (255))"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(PricingLevelID Counter PRIMARY KEY," _
						& " PricingLevelName char (20),"_
						& " PricingLevelNotes char (255))"
			End If
			objCnn.Execute pstrSQL,, 128
			If Err.number = 0 Then
				pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & " successfully added.</li>"
			Else
				pblnError = True
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrFieldName & " Error " & Err.number & ": " & Err.description &"</FONT></li>"
				Err.Clear
			end if

			'------------------------------------------------------------------------------'
			' Update sfAttributes TABLE														   '
			'------------------------------------------------------------------------------'

			'Reset Settings
			pstrTempMessage = ""
			pblnError = False

			If mblnSQLServer Then
				objCnn.Execute "Alter TABLE sfCustomers ADD PricingLevelID int",, 128
				objCnn.Execute "Alter TABLE sfCustomers ADD clubCode char (50)",, 128
				objCnn.Execute "Alter TABLE sfCustomers ADD clubExpDate Datetime",, 128
			Else
				objCnn.Execute "Alter TABLE sfCustomers ADD PricingLevelID long",, 128
				objCnn.Execute "Alter TABLE sfCustomers ADD clubCode char (50)",, 128
				objCnn.Execute "Alter TABLE sfCustomers ADD clubExpDate Date",, 128
			End If

			If mblnSQLServer Then
				objCnn.Execute "Alter TABLE sfProducts ADD prodPLPrice text",, 128
				objCnn.Execute "Alter TABLE sfProducts ADD prodPLSalePrice text",, 128
			Else
				objCnn.Execute "Alter TABLE sfProducts ADD prodPLPrice memo",, 128
				objCnn.Execute "Alter TABLE sfProducts ADD prodPLSalePrice memo",, 128
			End If

			If mblnSQLServer Then
				objCnn.Execute "Alter TABLE sfAttributeDetail ADD attrdtPLPrice text",, 128
			Else
				objCnn.Execute "Alter TABLE sfAttributeDetail ADD attrdtPLPrice memo",, 128
			End If

			If mblnSF5AE Then
				If mblnSQLServer Then
					objCnn.Execute "Alter TABLE sfMTPrices ADD mtPLValue text",, 128
				Else
					objCnn.Execute "Alter TABLE sfMTPrices ADD mtPLValue memo",, 128
				End If
			End If

			If Err.number = 0 Then
				pstrTempMessage = pstrTempMessage & "<li>Pricing Level database modification installed successfully.</li>"
			Else
				pblnError = True
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error installing Pricing Level modifications: " & Err.description &"</FONT></li>"
				Err.Clear
			end if

			mblnError = pblnError
			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
			
		Else

			'------------------------------------------------------------------------------'
			' REMOVE PricingLevels TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "PricingLevels"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE TABLE UPGRADES														   '
			'------------------------------------------------------------------------------'

			objCnn.Execute "Alter TABLE sfCustomers DROP PricingLevelID",, 128
			objCnn.Execute "Alter TABLE sfProducts DROP prodPLPrice",, 128
			objCnn.Execute "Alter TABLE sfProducts DROP prodPLSalePrice",, 128
			objCnn.Execute "Alter TABLE sfAttributeDetail DROP attrdtPLPrice",, 128

			If mblnSF5AE Then objCnn.Execute "Alter TABLE sfMTPrices DROP mtPLValue",, 128

			if Err.number = 0 then
				mblnUpgraded = False
				pstrTempMessage = pstrTempMessage & "<LI><B>Pricing Level database modifications successfully removed.</B></LI>"
			else
				mblnError = True
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing Pricing Level database modifications: " & Err.description	& "</FONT></LI>"	
			end if

			mblnError = pblnError
			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
			
		End If

	End Function	'Install_PricingLevelAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Product Placement
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_ProductPlacementAddon(byRef objCnn, byVal blnInstall)

	Dim pstrSQL
	Dim pstrTableName
	Dim pstrFieldName
	Dim pstrTempMessage
	Dim pblnError
	Dim i
	Dim pstrLocalError

	On Error Resume Next

		If blnInstall then

			'------------------------------------------------------------------------------'
			' Update sfProducts TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfProducts"
			
			ReDim paryFields(2)			
			paryFields(0) = Array("ADD","sortCat","long")
			paryFields(1) = Array("ADD","sortMfg","long")
			paryFields(2) = Array("ADD","sortVend","long")
			
			For i = 0 To UBound(paryFields)
				
				If mblnSQLServer Then
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				Else
					pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(enUpdate_Action) & " COLUMN " & paryFields(i)(enUpdate_Field) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
				End If
				objCnn.Execute pstrSQL,, 128
				
				if Err.number <> 0 then
					pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
									& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
					Err.Clear
				End If

			Next 'i
			
			if Len(pstrLocalError) = 0 then
				mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
				mblnUpgraded = True
			else
				mblnError = True
				mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
			end if

			'------------------------------------------------------------------------------'
			' Update sfSubCatDetail TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfSubCatDetail"
			
			If mblnSF5AE Then
				ReDim paryFields(0)
				paryFields(0) = Array("ADD","sortCatDetail","long")

				For i = 0 To UBound(paryFields)
					
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(0) & " " & paryFields(i)(1) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(0) & " COLUMN " & paryFields(i)(1) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
										& "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
				
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if

			End If
						
			If Not mblnError Then
				mstrMessage = mstrMessage & "<H4>You can order your products <a href='../../ssProductPlacementAdmin.asp'>here</a>.</H4>"			
			End If
			
		Else

			'------------------------------------------------------------------------------'
			' REMOVE TABLE UPGRADES														   '
			'------------------------------------------------------------------------------'

			objCnn.Execute "Alter TABLE sfProducts DROP sortCat",, 128
			objCnn.Execute "Alter TABLE sfProducts DROP sortMfg",, 128
			objCnn.Execute "Alter TABLE sfProducts DROP sortVend",, 128
			If mblnSF5AE Then
				objCnn.Execute "Alter TABLE sfSubCatDetail DROP sortCatDetail",, 128
			End If

			if Err.number = 0 then
				mblnUpgraded = False
				pstrTempMessage = pstrTempMessage & "<LI><B>Product Placement database modifications successfully removed.</B></LI>"
			else
				mblnError = True
				pstrTempMessage = pstrTempMessage & "<LI><Font Color='Red'>Error removing Product Placement database modifications: " & Err.description	& "</FONT></LI>"	
			end if

			mblnError = pblnError
			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
			
		End If

	End Function	'Install_ProductPlacementAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Promo Mail
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_PromoMailAddon(byRef objCnn, byVal blnInstall)

	Dim pstrSQL
	Dim pstrTableName

	On Error Resume Next
	
		If blnInstall then

			'------------------------------------------------------------------------------'
			' ADD ssPromoMailsToSend TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "sfCustomers"
			pstrSQL = "ALTER TABLE sfCustomers ADD ssPromoMailSent Char (1)"

			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				mblnUpgraded = True
				mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully updated.</h4>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

		Else

			'------------------------------------------------------------------------------'
			' REMOVE ssPromoMailsToSend TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "sfCustomers"
			pstrSQL = "ALTER TABLE sfCustomers DROP COLUMN ssPromoMailSent"
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully updated.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

		End If

	End Function	'Install_PromoMailAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Promotion Manager
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_PromotionManagerAddon(byRef objCnn, byVal blnInstall)

	Dim pstrSQL
	Dim pstrTableName

	On Error Resume Next
	
		If blnInstall then

			'------------------------------------------------------------------------------'
			' ADD ordersDiscounts TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ordersDiscounts"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(ordersDiscountID int Identity PRIMARY KEY," _
						& " OrderID int,"_
						& " PromotionID int,"_
						& " DiscountAmount Decimal (10,2))"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(ordersDiscountID COUNTER PRIMARY KEY," _
						& " OrderID long,"_
						& " PromotionID long,"_
						& " DiscountAmount double)"
			End If

			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				mblnUpgraded = True
				mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully updated.</h4>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

			'------------------------------------------------------------------------------'
			' ADD Promotions TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "Promotions"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(PromotionID int Identity PRIMARY KEY," _
						& " PromoCode char (50) UNIQUE,"_
						& " PromoTitle char (50),"_
						& " PromoRules text,"_
						& " StartDate Datetime,"_
						& " EndDate Datetime,"_
						& " Duration int,"_
						& " MaxUses int,"_
						& " NumUses int,"_
						& " Discount Decimal (10,2),"_
						& " Percentage bit,"_
						& " MinsubTotal Decimal (10,2),"_
						& " Combineable bit,"_
						& " ModifiedOn Datetime,"_
						& " Inactive bit,"_
						& " ApplyAutomatically bit,"_
						& " ExcludeSaleItems bit,"_
						& " ProductID text)"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(PromotionID COUNTER PRIMARY KEY," _
						& " PromoCode char (50) UNIQUE,"_
						& " PromoTitle char (50),"_
						& " PromoRules memo,"_
						& " StartDate Date,"_
						& " EndDate Date,"_
						& " Duration long,"_
						& " MaxUses long,"_
						& " NumUses long,"_
						& " Discount double,"_
						& " Percentage YESNO,"_
						& " MinsubTotal double,"_
						& " Combineable YESNO,"_
						& " ModifiedOn Date,"_
						& " Inactive YESNO,"_
						& " ApplyAutomatically YESNO,"_
						& " ExcludeSaleItems YESNO,"_
						& " ProductID memo)"
			End If

			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				mblnUpgraded = True
				mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully updated.</h4>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

		Else

			'------------------------------------------------------------------------------'
			' REMOVE ordersDiscounts TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ordersDiscounts"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE Promotions TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "Promotions"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

		End If

	End Function	'Install_PromotionManagerAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Promotion Manager II
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_PromotionManagerIIAddon(byRef objCnn, byVal blnInstall, ByVal bytUpgradeToLatest)

	Dim pstrSQL
	Dim pstrTableName

	On Error Resume Next
	
		If blnInstall then

			If bytUpgradeToLatest = 0 Then
			
				'------------------------------------------------------------------------------'
				' ADD ordersDiscounts TABLE														   '
				'------------------------------------------------------------------------------'

				pstrTableName = "ordersDiscounts"
				If mblnSQLServer Then
					pstrSQL = "CREATE TABLE " & pstrTableName & " " _
							& "(ordersDiscountID int Identity PRIMARY KEY," _
							& " OrderID int,"_
							& " PromotionID int,"_
							& " DiscountAmount Decimal (10,2))"
				Else
					pstrSQL = "CREATE TABLE " & pstrTableName & " " _
							& "(ordersDiscountID COUNTER PRIMARY KEY," _
							& " OrderID long,"_
							& " PromotionID long,"_
							& " DiscountAmount double)"
				End If

				objCnn.Execute pstrSQL,, 128

				if Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully updated.</h4>"
				else
					mblnError = True
					mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
					Err.Clear
				end if

				'------------------------------------------------------------------------------'
				' ADD Promotions TABLE														   '
				'------------------------------------------------------------------------------'

				pstrTableName = "Promotions"
				If mblnSQLServer Then
					pstrSQL = "CREATE TABLE " & pstrTableName & " " _
							& "(PromotionID int Identity PRIMARY KEY," _
							& " PromoCode char (50) UNIQUE,"_
							& " PromoTitle char (50),"_
							& " PromoRules text,"_
							& " StartDate Datetime,"_
							& " EndDate Datetime,"_
							& " Duration int,"_
							& " MaxUses int,"_
							& " NumUses int,"_
							& " Discount Decimal (10,2),"_
							& " Percentage bit,"_
							& " MinsubTotal Decimal (10,2),"_
							& " Combineable bit,"_
							& " ModifiedOn Datetime,"_
							& " Inactive bit,"_
							& " ApplyAutomatically bit,"_
							& " ExcludeSaleItems bit,"_
							& " ProductID text," _
							& " productCountLimit int,"_
							& " buyX int,"_
							& " getY int,"_
							& " FreeShippingLimit int,"_
							& " likeItem bit,"_
							& " ProductIDExclusion text,"_
							& " Category text,"_
							& " CategoryExclusion text,"_
							& " Manufacturer text,"_
							& " ManufacturerExclusion text,"_
							& " Vendor text,"_
							& " VendorExclusion text,"_
							& " FreeProductID text,"_
							& " offerFreeGiftAutomatically bit,"_
							& " FreeShippingCode char(50))"
				Else
					pstrSQL = "CREATE TABLE " & pstrTableName & " " _
							& "(PromotionID COUNTER PRIMARY KEY," _
							& " PromoCode char (50) UNIQUE,"_
							& " PromoTitle char (50),"_
							& " PromoRules memo,"_
							& " StartDate Date,"_
							& " EndDate Date,"_
							& " Duration long,"_
							& " MaxUses long,"_
							& " NumUses long,"_
							& " Discount double,"_
							& " Percentage YESNO,"_
							& " MinsubTotal double,"_
							& " Combineable YESNO,"_
							& " ModifiedOn Date,"_
							& " Inactive YESNO,"_
							& " ApplyAutomatically YESNO,"_
							& " ExcludeSaleItems YESNO,"_
							& " ProductID memo," _
							& " productCountLimit long,"_
							& " buyX long,"_
							& " getY long,"_
							& " FreeShippingLimit long,"_
							& " likeItem YESNO,"_
							& " ProductIDExclusion Memo,"_
							& " Category Memo,"_
							& " CategoryExclusion Memo,"_
							& " Manufacturer Memo,"_
							& " ManufacturerExclusion Memo,"_
							& " Vendor Memo,"_
							& " VendorExclusion Memo,"_
							& " FreeProductID Memo,"_
							& " offerFreeGiftAutomatically YESNO,"_
							& " FreeShippingCode char(50))"
				End If

				objCnn.Execute pstrSQL,, 128

				if Err.number = 0 then
					mblnUpgraded = True
					mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully updated.</h4>"
				else
					mblnError = True
					mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error updating " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
					Err.Clear
				end if
			
			ElseIf bytUpgradeToLatest = 1 Then
			
				'------------------------------------------------------------------------------'
				' ADD Promotions TABLE														   '
				'------------------------------------------------------------------------------'


				pstrTableName = "Promotions"
				Dim paryFields(18)
				
				paryFields(0) = Array("ADD","productCountLimit","long")
				paryFields(1) = Array("ADD","buyX","long")
				paryFields(2) = Array("ADD","getY","long")
				paryFields(3) = Array("ADD","FreeShippingLimit","long")
				
				paryFields(4) = Array("ADD","likeItem","YESNO")
				paryFields(5) = Array("ADD","ProductIDExclusion","Memo")
				paryFields(6) = Array("ADD","Category","Memo")
				paryFields(7) = Array("ADD","CategoryExclusion","Memo")
				paryFields(8) = Array("ADD","Manufacturer","Memo")
				paryFields(9) = Array("ADD","ManufacturerExclusion","Memo")
				paryFields(10) = Array("ADD","Vendor","Memo")
				paryFields(11) = Array("ADD","VendorExclusion","Memo")
				paryFields(12) = Array("ADD","FreeProductID","Memo")
				paryFields(13) = Array("ADD","offerFreeGiftAutomatically","YESNO")
				paryFields(14) = Array("ADD","FreeShippingCode","varchar(50)")
				paryFields(15) = Array("ADD","MaxAllowableValue","varchar(50)")
				paryFields(16) = Array("ADD","MaxAllowableValuePerItem","varchar(50)")
				paryFields(17) = Array("ADD","ApplyToBasePrice","YESNO")
				paryFields(18) = Array("ADD","NumUsesByCustomer","long")
				
				Dim pstrLocalError
				For i = 0 To UBound(paryFields)
					If mblnSQLServer Then
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(0) & " " & paryFields(i)(1) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					Else
						pstrSQL = "ALTER TABLE " & pstrTableName & " " & paryFields(i)(0) & " COLUMN " & paryFields(i)(1) & " " & adjustSQLServerType(paryFields(i)(enUpdate_Type))
					End If
					objCnn.Execute pstrSQL,, 128
					
					if Err.number <> 0 then
						pstrLocalError = pstrLocalError & "&nbsp;&nbsp;<font color=red>" & Err.Number & ": " & Err.Description & "</font><br>" & vbcrlf _
									   & "&nbsp;&nbsp;<font color=red>" & pstrSQL & "</font><BR>"
						Err.Clear
					End If

				Next 'i
				
				if Len(pstrLocalError) = 0 then
					mstrMessage = mstrMessage & "<H3><B>" & pstrTableName & " Table Successfully Upgraded</B></H3><BR>"
					mblnUpgraded = True
				else
					mblnError = True
					mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></H3>" & pstrLocalError
				end if

			End If
			
			If Not mblnError Then
				mstrMessage = mstrMessage & "<H4>You can create your promotions <a href='../../ssPromotionsAdmin.asp'>here</a>.</H4>"			
			End If

		Else

			'------------------------------------------------------------------------------'
			' REMOVE ordersDiscounts TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ordersDiscounts"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE Promotions TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "Promotions"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

		End If

	End Function	'Install_PromotionManagerIIAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Postage Rate
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_PostageRateAddon(byRef objCnn, byVal blnInstall)

	Dim paryShippingCarriers(8)
	Dim paryShippingMethods(66)
	Dim pstrSQL
	Dim pstrTableName
	Dim pstrTempMessage
	Dim pbytEnabled
	Dim plngStepCounter	'exists to save reordering
	
		'Set the Carriers
		paryShippingCarriers(1) = Array ("Unknown", "RateURL", "Username", "Password", "TrackingURL", "Imagepath")
		paryShippingCarriers(2) = Array ("U.S.P.S.","http://Production.ShippingApis.com/ShippingAPI.dll?API=","Username","Password","http://www.framed.usps.com/cgi-bin/cttgate/ontrack.cgi?tracknbr=[TrackingNumber]","Imagepath")
		paryShippingCarriers(3) = Array ("UPS","http://www.ups.com/using/services/rave/qcost_dss.cgi", "Username", "Password", "http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=1&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=[TrackingNumber]","Imagepath")
		paryShippingCarriers(4) = Array ("FedEx","http://grd.fedex.com/cgi-bin/rrr2010.exe","Username","Password","http://www.fedex.com/cgi-bin/tracking?action=track&language=english&cntry_code=us&tracknumbers=[TrackingNumber]","Imagepath")
		paryShippingCarriers(5) = Array ("Canada Post","http://206.191.4.228:30000","Username","Password","http://204.104.133.7/scripts/tracktrace.dll?MfcIsApiCommand=TraceE&referrer=CPCNewPop&i_num=[TrackingNumber]","Imagepath")
		paryShippingCarriers(6) = Array ("Airborne","RateURL","Username","Password","TrackingURL","Imagepath")
		paryShippingCarriers(7) = Array ("DHL","RateURL","Username","Password","TrackingURL","Imagepath")
		paryShippingCarriers(8) = Array ("Freight Quote","RateURL","Username","Password","TrackingURL","Imagepath")

		'Set the Methods
		'paryShippingMethods(0) = Array ("ssShippingCarrierID", "ssShippingMethodName", "ssShippingMethodCode", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

		'U.S.P.S. Methods
		'domestic
		plngStepCounter = 0
		pbytEnabled = 0
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Express Mail", "Express", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Priority Mail", "Priority", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Parcel Post", "Parcel", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "First Class", "First Class", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Media", "Media", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

		'international
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Global Express Guaranteed Document Service", "Global Express Guaranteed Document Service", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Global Express Guaranteed Non-Document Service", "Global Express Guaranteed Non-Document Service", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Global Express Mail (EMS)", "Global Express Mail (EMS)", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Global Priority Mail - Flat-rate Envelope (large)", "Global Priority Mail - Flat-rate Envelope (large)", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Global Priority Mail - Flat-rate Envelope (small)", "Global Priority Mail - Flat-rate Envelope (small)", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Global Priority Mail - Variable Weight Envelope (single)", "Global Priority Mail - Variable Weight Envelope (single)", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Airmail Letter Post", "Airmail Letter Post", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Airmail Parcel Post", "Airmail Parcel Post", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Economy (Surface) Letter Post", "Economy (Surface) Letter Post", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (2, "Economy (Surface) Parcel Post", "Economy (Surface) Parcel Post", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

		'UPS Methods
		pbytEnabled = 1
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS Next Day AM", "1DM", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS Next Day Air", "1DA", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS Next Day Air Saver", "1DP", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS 2nd Day Air AM", "2DM", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS 2nd Day Air", "2DA", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS 3 Day Select", "3DS", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS Standard Ground", "GND", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS Canada Standard", "STD", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS WorldWide Express", "XPR", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS WorldWide Express Plus", "XDM", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (3, "UPS WorldWide Expedited", "XPD", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

		'FedEx Methods
		pbytEnabled = 0
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx Priority", "01", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx 2day", "03", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx Standard Overnight", "05", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx First Overnight", "06", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx Express Saver", "20", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx Overnight Freight", "70", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx 2day Freight", "80", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx Express Saver Freight", "83", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx International Priority", "01i", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx International Economy", "03i", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx International First", "06i", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx International Priority Freight", "70i", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx International Economy Freight", "86i", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "FedEx Home Delivery", "90", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "U.S. Domestic FedEx Ground Package", "92", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (4, "International FedEx Ground Package", "92i", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

		'Canada Post Methods
		pbytEnabled = 0
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Domestic - Regular", "1010", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Domestic - Expedited", "1020", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Domestic - Xpresspost", "1030", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Domestic - Priority Courier", "1040", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Domestic - Expedited Evening", "1120", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Domestic - Xpresspost Evening", "1130", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Domestic - Expedited Saturday", "1220", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Domestic - Xpresspost Saturday", "1230", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "USA - Small Packet Surface", "2010", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "USA - Small Packet Air", "2015", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "USA - Expedited Business", "2020", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "Expedited Commercial", "2025", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "USA - Xpresspost", "2030", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "USA - Purolator", "2040", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "USA - Puropak", "2050", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "International - Small Packet Surface", "3005", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "International - Surface", "3010", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "International - Air", "3020", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "International - Xpresspost", "3025", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "International - Purolator", "3040", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "International - Puropak", "3050", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "International - Purolator Air", "5010", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (5, "International - Purolator Surface", "5020", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

		'Freight Quote
		pbytEnabled = 0
		paryShippingMethods(stepCounter(plngStepCounter)) = Array (8, "Freight Quote", "FreightQuote", pbytEnabled, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

		'debugprint "plngStepCounter",plngStepCounter

		'For i = 1 To plngStepCounter
		'	Response.Write i & ": " & paryShippingMethods(i)(1) & "<BR>"
		'Next 'i

	On Error Resume Next

		If blnInstall then

			If mblnSF5 Then 
				'------------------------------------------------------------------------------'
				'UPGRADE sfOrders TABLE														   '
				'------------------------------------------------------------------------------'

				pstrSQL = "ALTER TABLE sfOrders ALTER COLUMN orderShipMethod Char (65)"
				objCnn.Execute(pstrSQL)

				if Err.number = 0 then
					pstrTempMessage = pstrTempMessage & "<li>sfOrders table successfully upgraded</li>"
				else
					mblnError = True
					mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error upgrading sfOrders Table: " & Err.description	& "</FONT></h4>"	
					Err.Clear
				end if
			Else
				'------------------------------------------------------------------------------'
				'UPGRADE customer TABLE														   '
				'------------------------------------------------------------------------------'

				pstrSQL = "ALTER TABLE customer ALTER COLUMN SHIPPING_METHOD Char (65)"
				objCnn.Execute(pstrSQL)

				if Err.number = 0 then
					pstrTempMessage = pstrTempMessage & "<li>sfOrders table successfully upgraded</li>"
				else
					mblnError = True
					mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error upgrading customer Table: " & Err.description	& "</FONT></h4>"	
					Err.Clear
				end if
			End If

			'------------------------------------------------------------------------------'
			' ADD ssShippingCarriers TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssShippingCarriers"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(ssShippingCarrierID int Identity PRIMARY KEY," _
						& " ssShippingCarrierName char (65)," _
						& " ssShippingCarrierUserName char (65)," _
						& " ssShippingCarrierPassword char (65)," _
						& " ssShippingCarrierRateURL char (255)," _
						& " ssShippingCarrierTrackingURL char (255)," _
						& " ssShippingCarrierImagePath char (255)" _
						& " )"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(ssShippingCarrierID COUNTER PRIMARY KEY," _
						& " ssShippingCarrierName char (65)," _
						& " ssShippingCarrierUserName char (65)," _
						& " ssShippingCarrierPassword char (65)," _
						& " ssShippingCarrierRateURL char (255)," _
						& " ssShippingCarrierTrackingURL char (255)," _
						& " ssShippingCarrierImagePath char (255)" _
						& " )"
			End If

			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & " table successfully added</li>"
				For i = 1 To UBound(paryShippingCarriers)
					pstrSQL = "Insert Into " & pstrTableName & " (ssShippingCarrierID, ssShippingCarrierName, ssShippingCarrierUserName, ssShippingCarrierPassword, ssShippingCarrierRateURL, ssShippingCarrierTrackingURL, ssShippingCarrierImagePath)" _
						& " Values (" & i & ", '" & paryShippingCarriers(i)(0) & "', '" & paryShippingCarriers(i)(2) & "', '" & paryShippingCarriers(i)(3) & "', '" & paryShippingCarriers(i)(1) & "', '" & paryShippingCarriers(i)(4) & "', '" & paryShippingCarriers(i)(5) & "')"
					If mblnSQLServer Then pstrSQL = "SET IDENTITY_INSERT " & pstrTableName & " ON;" & pstrSQL
					objCnn.Execute pstrSQL,, 128
				Next	'i
				If mblnSQLServer Then objCnn.Execute "SET IDENTITY_INSERT " & pstrTableName & " OFF;",, 128
			
				mblnUpgraded = True
			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error adding " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

			'------------------------------------------------------------------------------'
			' ADD ssShippingMethods TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssShippingMethods"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(ssShippingMethodID int Identity PRIMARY KEY," _
						& " ssShippingCarrierID int," _
						& " ssShippingMethodCode char (65)," _
						& " ssShippingMethodName char (65)," _
						& " ssShippingMethodEnabled bit," _
						& " ssShippingMethodIsSpecial bit," _
						& " ssShippingMethodLocked bit," _
						& " ssShippingMethodMinCharge Decimal (10,2)," _
						& " ssShippingMethodMultiple Decimal (10,2)," _
						& " ssShippingMethodPerPackageFee Decimal (10,2)," _
						& " ssShippingMethodPerShipmentFee Decimal (10,2)," _
						& " ssShippingMethodOfferFreeShippingAbove Decimal (10,2)," _
						& " ssShippingMethodLimitFreeShippingByWeight Decimal (10,2)," _
						& " ssShippingMethodClass Decimal (10,2)," _
						& " ssShippingMethodOrderBy Decimal (10,2)," _
						& " ssShippingMethodDefault bit," _
						& " ssShippingMethodMinWeight Decimal (10,2)," _
						& " ssShippingMethodPrefWeight Decimal (10,2)," _
						& " ssShippingMethodMaxLength Decimal (10,2)," _
						& " ssShippingMethodMaxWidth Decimal (10,2)," _
						& " ssShippingMethodMaxHeight Decimal (10,2)," _
						& " ssShippingMethodMaxWeight Decimal (10,2)," _
						& " ssShippingMethodMaxGirth Decimal (10,2)" _
						& " )"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(ssShippingMethodID COUNTER PRIMARY KEY," _
						& " ssShippingCarrierID long," _
						& " ssShippingMethodCode char (65)," _
						& " ssShippingMethodName char (65)," _
						& " ssShippingMethodEnabled YESNO," _
						& " ssShippingMethodIsSpecial YESNO," _
						& " ssShippingMethodLocked YESNO," _
						& " ssShippingMethodMinCharge double," _
						& " ssShippingMethodMultiple double," _
						& " ssShippingMethodPerPackageFee double," _
						& " ssShippingMethodPerShipmentFee double," _
						& " ssShippingMethodOfferFreeShippingAbove double," _
						& " ssShippingMethodLimitFreeShippingByWeight double," _
						& " ssShippingMethodClass double," _
						& " ssShippingMethodOrderBy double," _
						& " ssShippingMethodDefault YESNO," _
						& " ssShippingMethodMinWeight double," _
						& " ssShippingMethodPrefWeight double," _
						& " ssShippingMethodMaxLength double," _
						& " ssShippingMethodMaxWidth double," _
						& " ssShippingMethodMaxHeight double," _
						& " ssShippingMethodMaxWeight double," _
						& " ssShippingMethodMaxGirth double" _
						& " )"
			End If

			objCnn.Execute (pstrSQL)

			if Err.number = 0 then
				mblnUpgraded = True
				pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & " table successfully added</li>"
				For i = 1 To plngStepCounter
					pstrSQL = "Insert Into " & pstrTableName & " (ssShippingMethodID, ssShippingCarrierID, ssShippingMethodCode, ssShippingMethodName, ssShippingMethodEnabled, ssShippingMethodIsSpecial, ssShippingMethodLocked, ssShippingMethodMultiple, ssShippingMethodPerPackageFee, ssShippingMethodPerShipmentFee, ssShippingMethodOfferFreeShippingAbove, ssShippingMethodLimitFreeShippingByWeight, ssShippingMethodClass, ssShippingMethodOrderBy, ssShippingMethodDefault, ssShippingMethodMinCharge, ssShippingMethodMinWeight, ssShippingMethodPrefWeight, ssShippingMethodMaxLength, ssShippingMethodMaxWidth, ssShippingMethodMaxHeight, ssShippingMethodMaxWeight, ssShippingMethodMaxGirth)" _
						& " Values (" & i & ", '" & paryShippingMethods(i)(0) & "', '" & paryShippingMethods(i)(2) & "', '" & paryShippingMethods(i)(1) & "', " & paryShippingMethods(i)(3) & ", 0," & paryShippingMethods(i)(4) & ", " & paryShippingMethods(i)(5) & ", " & paryShippingMethods(i)(6) & ", " & paryShippingMethods(i)(7) & ", " & paryShippingMethods(i)(8) & ", " & paryShippingMethods(i)(9) & ", " & paryShippingMethods(i)(10) & ", " & paryShippingMethods(i)(11) & ", " & paryShippingMethods(i)(12) & ",0,0,50,0,0,0,170,0)"
					If mblnSQLServer Then pstrSQL = "SET IDENTITY_INSERT " & pstrTableName & " ON;" & pstrSQL
					objCnn.Execute pstrSQL,, 128
					If Err.number <> 0 Then
						Response.Write "SQL: " & pstrSQL & "<BR>"
						Err.Clear
					End If
				Next	'i
				If mblnSQLServer Then objCnn.Execute "SET IDENTITY_INSERT " & pstrTableName & " OFF;",, 128

			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error adding " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"

		Else

			'------------------------------------------------------------------------------'
			' REMOVE ssShippingCarriers TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssShippingCarriers"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

			'------------------------------------------------------------------------------'
			' REMOVE ssShippingMethods TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssShippingMethods"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute (pstrSQL)
			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

		End If

	End Function	'Install_PostageRateAddon
	
'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Tax Rate
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_TaxRateAddon(byRef objCnn, byVal blnInstall)

	Dim pstrSQL
	Dim pstrTableName
	Dim pstrTempMessage
	Dim pbytEnabled
	
	'On Error Resume Next

		If blnInstall then

		'------------------------------------------------------------------------------'
		' ADD ssTaxTable TABLE														   '
		'------------------------------------------------------------------------------'

			pstrTableName = "ssTaxTable"
			If mblnSQLServer Then
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(TaxRateID Int Identity," _
						& " City char (50),"_
						& " County char (50),"_
						& " PostalCode char (10),"_
						& " LocaleAbbr char (3),"_
						& " TaxRate Decimal Not Null)"
			Else
				pstrSQL = "CREATE TABLE " & pstrTableName & " " _
						& "(TaxRateID Counter PRIMARY KEY," _
						& " City char (50),"_
						& " County char (50),"_
						& " LocaleAbbr char (3),"_
						& " PostalCode char (10),"_
						& " TaxRate Single Not Null)"
			End If

			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & " table successfully added</li>"
				mblnUpgraded = True
			else
				mblnError = True
				mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error adding " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
				Err.Clear
			end if

			mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"

		Else

			'------------------------------------------------------------------------------'
			' REMOVE ssTaxTable TABLE														   '
			'------------------------------------------------------------------------------'

			pstrTableName = "ssTaxTable"
			pstrSQL = "DROP TABLE " & pstrTableName
			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				mblnUpgraded = False
				mstrMessage = mstrMessage & "<LI><B>Table " & pstrTableName & " successfully removed.</B></LI>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></LI>"	
			end if

		End If

	End Function	'Install_TaxRateAddon
	
'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	WebStore Manager
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Function Install_WebStoreManagerAddon(byRef objCnn, byVal blnInstall)

Dim pstrTableName
Dim pstrTempMessage
Dim pstrLocalErrorMessage
Dim pblnSuccess
Dim plngTableCounter
Dim plngRecordCounter
Dim paryRecordInsertions

	Dim parydbUpgrades(1)
	'contains array of table name, tableID (for new tables only), array of fieldDefinitions, array of records to insert (optional)
	'fieldDefinitions is array of action, fieldName, ACCESS specific field type, original field type
	
	'ssUsers
	ReDim paryFields(13)			
		paryFields(0) = Array("ADD","userName","varchar(20)")
		paryFields(1) = Array("ADD","userPass","varchar(10)")
		paryFields(2) = Array("ADD","userInitials","varchar(10)")
		paryFields(3) = Array("ADD","isActive","yesno")
		paryFields(4) = Array("ADD","isAdmin","yesno")
		paryFields(5) = Array("ADD","failedLoginAttempts","long")
		paryFields(6) = Array("ADD","LastLogin","datetime")
		paryFields(7) = Array("ADD","Orders_View","yesno")
		paryFields(8) = Array("ADD","Orders_ViewCC","yesno")
		paryFields(9) = Array("ADD","Orders_Edit","yesno")
		paryFields(10) = Array("ADD","Products_Edit","yesno")
		paryFields(11) = Array("ADD","dateEdited","date")
		paryFields(12) = Array("ADD","Orders_Delete","yesno")
		paryFields(13) = Array("ADD","Orders_Report","yesno")
		
	ReDim paryRecordInsertions(0)
		paryRecordInsertions(0) = "INSERT INTO ssUsers (userName, userPass, isAdmin, Orders_View, Orders_ViewCC, Orders_Edit, Products_Edit, Orders_Delete, Orders_Report) Values ('admin', 'pass', 1, 1, 1, 1, 1, 1, 1)"
		
	parydbUpgrades(0) = Array("ssUsers", "ssUserID", paryFields, paryRecordInsertions)

	'ssUserLog
	ReDim paryFields(3)			
		paryFields(0) = Array("ADD","ssUserLogUserID","varchar(20)")
		paryFields(1) = Array("ADD","ssUserLogIP","varchar(20)")
		paryFields(2) = Array("ADD","ssUserLogInDate","datetime")
		paryFields(3) = Array("ADD","ssUserLogInResult","varchar(40)")
		
	parydbUpgrades(1) = Array("ssUserLog", "ssUserLogID", paryFields, "")

	On Error Resume Next

	pblnSuccess = True
	
	If blnInstall then
	
		For plngTableCounter = 0 To UBound(parydbUpgrades)
			pstrTableName = parydbUpgrades(plngTableCounter)(0)
			
			'------------------------------------------------------------------------------'
			' Create table
			'------------------------------------------------------------------------------'
			If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
				pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(1), pstrLocalErrorMessage)
				
				'Intermediate error checking
				if Len(pstrLocalErrorMessage) = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
				else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
				end if
			End If	'new table check
			
			'------------------------------------------------------------------------------'
			' Add/alter field definitions
			'------------------------------------------------------------------------------'
			
			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if
			
			'Now check for any records to insert
			If UBound(parydbUpgrades(plngTableCounter)) >= 3 Then
				If Err.Number <> 0 Then err.Clear
			
				paryRecordInsertions = parydbUpgrades(plngTableCounter)(3)
				For plngRecordCounter = 0 To UBound(paryRecordInsertions)
					objCnn.Execute paryRecordInsertions(plngRecordCounter),,128
				Next 'plngRecordCounter

				'Intermediate error checking
				If Err.Number = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully populated</B></li><BR>"
				Else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error populating " & pstrTableName & "</FONT></li>" _
														& "<li>--Error " & Err.Number & ": " & Err.Description & "</li>"
				End If

			End If

		Next 'plngTableCounter

		'------------------------------------------------------------------------------'
		' Final Error Checking
		'------------------------------------------------------------------------------'
		if pblnSuccess then
			pstrTempMessage = "<LI><B>" & maryAddons(enAO_WebStoreMgr)(enName) & " database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
		else
			mblnError = True
			pstrTempMessage = "<LI><Font Color='Red'>Error adding " & maryAddons(enAO_WebStoreMgr)(enName) & " database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
		end if

	Else

		For plngTableCounter = 0 To UBound(parydbUpgrades)
			pstrTableName = parydbUpgrades(plngTableCounter)(0)
			
			If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
				'------------------------------------------------------------------------------'
				' Remove table
				'------------------------------------------------------------------------------'
				pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "", pstrLocalErrorMessage)
				
				'Intermediate error checking
				if Len(pstrLocalErrorMessage) = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully deleted</B></li><BR>"
				else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error deleting " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
				end if
			Else
				'------------------------------------------------------------------------------'
				' Undo Add/alter field definitions
				'------------------------------------------------------------------------------'
				Call setFieldDefinitionsForRemoval(parydbUpgrades(plngTableCounter)(2))
				
				pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

				'Intermediate error checking
				if Len(pstrLocalErrorMessage) = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table changes successfully removed</B></li><BR>"
				else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error undoing " & pstrTableName & " table changes</FONT></li>" & pstrLocalErrorMessage
				end if
			
			End If	'new table check
			

		Next 'plngTableCounter

		'------------------------------------------------------------------------------'
		' Final Error Checking
		'------------------------------------------------------------------------------'
		if pblnSuccess then
			pstrTempMessage = "<LI><B>" & maryAddons(enAO_WebStoreMgr)(enName) & " database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
		else
			mblnError = True
			pstrTempMessage = "<LI><Font Color='Red'>Error removing " & maryAddons(enAO_WebStoreMgr)(enName) & " database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
		end if

	End If	'blnInstall

	'------------------------------------------------------------------------------'
	' Record success or failure
	'------------------------------------------------------------------------------'
	mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
	mblnUpgraded = pblnSuccess
	
	Install_WebStoreManagerAddon = mblnUpgraded

End Function	'Install_WebStoreManagerAddon

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Zone Based Shipping
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Function Install_ZBSAddon(byRef objCnn, byVal blnInstall)

Dim pstrSQL
Dim pstrTableName
Dim pstrFieldName
Dim pstrTempMessage
Dim pblnError

	On Error Resume Next

	If blnInstall then

		'------------------------------------------------------------------------------'
		'UPGRADE sfShipping TABLE														   '
		'------------------------------------------------------------------------------'

		pstrTableName = "sfShipping"
		pstrSQL = "ALTER TABLE sfShipping ALTER COLUMN shipMethod Char (65)"
		objCnn.Execute(pstrSQL)

		If mblnSQLServer Then
			pstrSQL = "ALTER TABLE sfShipping ALTER COLUMN shipCode Char (60) "
			objCnn.Execute(pstrSQL)

			pstrSQL = "CREATE UNIQUE INDEX shipCode_ind ON sfShipping (shipCode)"
			objCnn.Execute(pstrSQL)
		Else
			pstrSQL = "ALTER TABLE sfShipping ALTER COLUMN shipCode Char (60) Unique"
			objCnn.Execute(pstrSQL)
		End If

		if Err.number = 0 then
			mstrMessage = mstrMessage & "<li>" & pstrTableName & " table successfully upgraded</li>"
		else
			mblnError = True
			mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error upgrading " & pstrTableName & " Table: " & Err.description	& "</FONT></h4>"	
		end if

		'------------------------------------------------------------------------------'
		' ADD ssShipZones TABLE														   '
		'------------------------------------------------------------------------------'

		pstrTableName = "ssShipZones"
		If mblnSQLServer Then
			pstrSQL = "CREATE TABLE " & pstrTableName & " " _
					& "(ZoneID int Identity," _
					& " ZoneName char (50) PRIMARY KEY,"_
					& " ZoneCountries text,"_
					& " ZoneStates text,"_
					& " ZoneZips text)"
		Else
			pstrSQL = "CREATE TABLE " & pstrTableName & " " _
					& "(ZoneID Counter," _
					& " ZoneName char (50) PRIMARY KEY,"_
					& " ZoneCountries memo,"_
					& " ZoneStates memo,"_
					& " ZoneZips memo)"
		End If

		objCnn.Execute pstrSQL,, 128

		if Err.number = 0 then
			mstrMessage = mstrMessage & "<li>" & pstrTableName & " table successfully added</li>"
			mblnUpgraded = True
		else
			mblnError = True
			mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error adding " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
			Err.Clear
		end if

		'------------------------------------------------------------------------------'
		' ADD ssShippingRates TABLE														   '
		'------------------------------------------------------------------------------'

		pstrTableName = "ssShippingRates"
		If mblnSQLServer Then
			pstrSQL = "CREATE TABLE " & pstrTableName & " " _
					& "(" _
					& " ShipRateID int Identity," _
					& " ShipMethod int," _
					& " ShipZone int," _
					& " ShipWeight decimal (10,3)," _
					& " ShipRate decimal (10,2)," _
					& " ShipRatePercentage tinyint," _
					& " CONSTRAINT ssShippingRates_pk Primary Key (ShipMethod,ShipZone,ShipWeight)" _
					& ")"
		Else
			pstrSQL = "CREATE TABLE " & pstrTableName & " " _
					& "(ShipRateID COUNTER," _
					& " ShipMethod long," _
					& " ShipZone long," _
					& " ShipWeight Double," _
					& " ShipRate Double, " _
					& " ShipRatePercentage YESNO," _
					& " CONSTRAINT MyTableConstraint Primary Key (ShipMethod,ShipZone,ShipWeight))"
		End If

		objCnn.Execute pstrSQL,, 128

		if Err.number = 0 then
			mstrMessage = mstrMessage & "<li>" & pstrTableName & " table successfully added</li>"
			mblnUpgraded = True
		else
			mblnError = True
			mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error adding " & pstrTableName & ": " & Err.description	& "</FONT></h4>"	
			Err.Clear
		end if

		mblnError = pblnError
		
	Else

		'------------------------------------------------------------------------------'
		' REPAIR sfShipping TABLE													   '
		'------------------------------------------------------------------------------'

		'Reset Settings
		pstrTempMessage = ""
		pblnError = False

		pstrTableName = "sfShipping"
		If mblnSQLServer Then
			pstrSQL = "DROP INDEX sfShipping.shipCode_ind"
			objCnn.Execute pstrSQL,, 128

			if Err.number = 0 then
				mstrMessage = mstrMessage & "<H3><B>sfShipping table successfully restored.</B></H3><BR>"
			else
				mblnError = True
				mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error removing " & pstrTableName & ": " & Err.description	& "</FONT></H3><BR>"	
				Err.Clear
			end if
		End If
			
		pstrSQL = "ALTER TABLE sfShipping ALTER COLUMN shipMethod Char (50)"
		objCnn.Execute pstrSQL,, 128

		pstrSQL = "ALTER TABLE sfShipping ALTER COLUMN shipCode Char (25)"
		objCnn.Execute pstrSQL,, 128

		If Err.number = 0 Then
			mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully removed.</h4>"
		Else
			pblnError = True
			mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error removing " & pstrTableName & "</FONT></h4>"	
			mstrMessage = mstrMessage & "<ul><li><Font Color='Red'>Error " & Err.number & ": " & Err.description &"</FONT></li></ul>"
			Err.Clear
		end if
		
		'------------------------------------------------------------------------------'
		' REMOVE ssShipZones TABLE														   '
		'------------------------------------------------------------------------------'

		pstrTableName = "ssShipZones"
		pstrSQL = "DROP TABLE " & pstrTableName
		objCnn.Execute pstrSQL,, 128
		If Err.number = 0 Then
			mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully removed.</h4>"
		Else
			pblnError = True
			mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error removing " & pstrTableName & "</FONT></h4>"	
			mstrMessage = mstrMessage & "<ul><li><Font Color='Red'>Error " & Err.number & ": " & Err.description &"</FONT></li></ul>"
			Err.Clear
		end if
		
		'------------------------------------------------------------------------------'
		' REMOVE ssShippingRates TABLE														   '
		'------------------------------------------------------------------------------'

		pstrTableName = "ssShippingRates"
		pstrSQL = "DROP TABLE " & pstrTableName
		objCnn.Execute pstrSQL,, 128
		If Err.number = 0 Then
			mstrMessage = mstrMessage & "<h4>" & pstrTableName & " table successfully removed.</h4>"
		Else
			pblnError = True
			mstrMessage = mstrMessage & "<h4><Font Color='Red'>Error removing " & pstrTableName & "</FONT></h4>"	
			mstrMessage = mstrMessage & "<ul><li><Font Color='Red'>Error " & Err.number & ": " & Err.description &"</FONT></li></ul>"
			Err.Clear
		end if
		
	End If

End Function	'Install_ZBSAddon

Function ShowTaxRateImportOption

Dim pblnDatabaseExists
Dim paryCustom(6)

	If Request.Form("ImportTaxTables") = "" Then
		paryCustom(0) = mstrDBPath & "NYSTaxRates.mdb"
		paryCustom(1) = "ssTaxTable"
		paryCustom(2) = "City"
		paryCustom(3) = "County"
		paryCustom(4) = "PostalCode"
		paryCustom(5) = "TaxRate"
		paryCustom(6) = "LocaleAbbr"
		
	Else
		paryCustom(0) = Trim(Request.Form("ImportPath"))
		paryCustom(1) = Trim(Request.Form("Table"))
		paryCustom(2) = Trim(Request.Form("City"))
		paryCustom(3) = Trim(Request.Form("County"))
		paryCustom(4) = Trim(Request.Form("PostalCode"))
		paryCustom(5) = Trim(Request.Form("TaxRate"))
		paryCustom(6) = Trim(Request.Form("State"))
		
	End If
	
	pblnDatabaseExists = FileExists(paryCustom(0))
	If pblnDatabaseExists And Request.Form("ImportTaxTables") <> "" Then Call ImportTaxTables
%>
<table border=1 cellpadding=2 cellspacing=0 width=95% ID="Table1">
	<tr>
		<td colspan=3><h4>Import Tax Information from a database</h4>
		<i>To import tax rate information from an external database fill in the information below. The form
		is prefilled with the default database table structure. The imported database can be of different 
		structure, but you need to fill in the appropriate table and field names. If the imported database 
		does not contain all of the fields, just leave the corresponding field name blank. You should locate the database in the fpdb folder and only enter the actual name of the datatabase to be imported.</i> 
		</td>
	</tr>
	<tr>
		<td colspan=3><div id="divAction">&nbsp;</div></td>
	</tr>
	<tr>
		<td colspan=3><div id="divPosition">&nbsp;</div><div id="divRecordCount"></div>&nbsp;</td>
	</tr>

<tr>
<td colspan=3><i>Database structure</i></td>
</tr>
<tr>
<td><b>&nbsp;</b></td><td>Imported Database</td><td>StoreFront Database</td>
</tr>
<tr>
	<td align=right>Database:</td><td><input name="ImportPath" ID="ImportPath" value="<%= paryCustom(0) %>"></td>
	<td>
	<%
		If pblnDatabaseExists Then
			Response.Write "Database confirmed"
		Else
			Response.Write "<font color=red>Invalid database path</font>"
		End If
	%>
	</td>
</tr>
<tr>
<td align=right>Table:</td><td><input name="Table" ID="Table" value="<%= paryCustom(1) %>"></td><td><i>ssTaxTable</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="City" ID="City" value="<%= paryCustom(2) %>"></td><td><i>City</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="County" ID="County" value="<%= paryCustom(3) %>"></td><td><i>County</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="PostalCode" ID="PostalCode" value="<%= paryCustom(4) %>"></td><td><i>PostalCode</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="TaxRate" ID="TaxRate" value="<%= paryCustom(5) %>"></td><td><i>TaxRate</i></td>
</tr>
<tr>
<td align=right>Field:</td><td><input name="State" ID="State" value="<%= paryCustom(6) %>"></td><td><i>State</i></td>
</tr>
<tr>
<td>&nbsp;</td>
<td colspan=2>
	<input type=radio name="overwrite" value=0 checked ID="overwrite0">&nbsp;Overwrite existing data.<br>
	<input type=radio name="overwrite" value=1 ID="overwrite1">&nbsp;Append to existing data.
</td>
</tr>
<tr>
<td>&nbsp;</td>
<td colspan=2><input type=submit value="Import Tax Tables" id="ImportTaxTables" name="ImportTaxTables"></td>
</tr>
</table>
<%
End Function	'ShowTaxRateImportOption

'***********************************************************************************************

	Sub ImportTaxTables
	
	On Error Goto 0
	
	Dim pobjcnnImport
	Dim prsTarget
	Dim prsSource
		
	Dim pstrFilePath
	Dim pstrTableName
	Dim pstrFieldName_City
	Dim pstrFieldName_County
	Dim pstrFieldName_PostalCode
	Dim pstrFieldName_TaxRate
	Dim pstrFieldName_State
	Dim pblnAppend
	Dim pblnSuccess
	
		pblnSuccess = True
		
		pstrFilePath = Trim(Request.Form("ImportPath"))
		pstrTableName = Trim(Request.Form("Table"))
		pstrFieldName_City = Trim(Request.Form("City"))
		pstrFieldName_County = Trim(Request.Form("County"))
		pstrFieldName_PostalCode = Trim(Request.Form("PostalCode"))
		pstrFieldName_State = Trim(Request.Form("State"))
		pstrFieldName_TaxRate = Trim(Request.Form("TaxRate"))
		pblnAppend = (Request.Form("overwrite") = 1)

		'debugprint "pstrFilePath",server.MapPath("/")
		'debugprint "pstrFilePath",pstrFilePath
		'debugprint "pstrTableName",pstrTableName
		'debugprint "pstrFieldName_City",pstrFieldName_City
		'debugprint "pstrFieldName_County",pstrFieldName_County
		'debugprint "pstrFieldName_PostalCode",pstrFieldName_PostalCode
		'debugprint "pstrFieldName_State",pstrFieldName_State
		'debugprint "pstrFieldName_TaxRate",pstrFieldName_TaxRate
		'debugprint "pblnAppend",pblnAppend

		If OpenDatabase(pobjcnnImport, pstrFilePath, False) Then
	
			Response.Write "Importing tax tables<ul>"
			If Not pblnAppend Then 
				Response.Write "<li>Deleting existing tax rates</li>"
				mobjCnn.Execute "Delete from ssTaxTable",,128
			End If

			'Open the source table
			Set prsSource = server.CreateObject("ADODB.RECORDSET")
			prsSource.CursorLocation =	adUseClient
			prsSource.Open "[" & pstrTableName & "]", pobjcnnImport, adOpenStatic, adLockOptimistic, adCmdTable
			If cBool(prsSource.State) Then
				Response.Write "<li>Source table " & pstrTableName & " successfully opened</li>"
			Else
				pblnSuccess = False
				If Err.number = 0 Then
					Response.Write "<font color=red><b>Error opening " & pstrTableName & "</b></font>"
				Else
					Response.Write "<font color=red><b>Error opening " & pstrTableName & " " & Err.number & ":" & Err.Description & "</b></font>"
				End If
			End If
			'debugprint "prsSource.RecordCount",prsSource.RecordCount
			
			'Open the target table
			Set prsTarget = server.CreateObject("ADODB.RECORDSET")
			prsTarget.CursorLocation =	adUseClient
			prsTarget.CacheSize = 100
			prsTarget.Open "ssTaxTable", mobjCnn, adUseServer, adLockOptimistic, adCmdTable
			If cBool(prsSource.State) Then
				Response.Write "<li>Target table " & pstrTableName & " successfully opened</li>"
			Else
				pblnSuccess = False
				If Err.number = 0 Then
					Response.Write "<font color=red><b>Error opening ssTaxTable</b></font>"
				Else
					Response.Write "<font color=red><b>Error opening ssTaxTable" & Err.number & ":" & Err.Description & "</b></font>"
				End If
			End If
			
			'debugprint "prsTarget.RecordCount",prsTarget.RecordCount
'On Error Resume Next

			If pblnSuccess Then

				Response.Write "<li>Inserting " & prsSource.RecordCount & " new tax rate(s)<li>"
				For i=1 to prsSource.RecordCount
					prsTarget.AddNew

					If len(pstrFieldName_City) > 0 Then prsTarget.Fields("City").Value = prsSource.Fields(pstrFieldName_City).Value
					If len(pstrFieldName_County) > 0 Then prsTarget.Fields("County").Value = prsSource.Fields(pstrFieldName_County).Value
					If len(pstrFieldName_PostalCode) > 0 Then prsTarget.Fields("PostalCode").Value = prsSource.Fields(pstrFieldName_PostalCode).Value
					If len(pstrFieldName_State) > 0 Then prsTarget.Fields("LocaleAbbr").Value = prsSource.Fields(pstrFieldName_State).Value
					If len(pstrFieldName_TaxRate) > 0 Then 
						prsTarget.Fields("TaxRate").Value = prsSource.Fields(pstrFieldName_TaxRate).Value
					Else
						prsTarget.Fields("TaxRate").Value = 0
					End If
					
					If Err.number <> 0 Then
						pblnSuccess = False
						Response.Write "<script>UpdatePosition('<font color=red><b>Error " & Err.number & ":" & Err.Description & "</b></font>');</script>"
						Exit For
					End If				
					prsSource.MoveNext
				Next
On Error Goto 0						
				If pblnSuccess Then prsTarget.UpdateBatch
			
On Error Resume Next
			End If
			
			prsTarget.Close
			Set prsTarget = Nothing

			prsSource.Close
			Set prsSource = Nothing
			
			pobjcnnImport.Close
			Set pobjcnnImport = Nothing
			
			If pblnSuccess Then Response.Write "<li>Database imported</li>"
			Response.Write "</ul>"
		End If
		
	End Sub	'ImportTaxTables

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	SE to AE Upgrade
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_SEtoAEUpgrade(byRef objCnn, byVal blnInstall)

	Dim pstrTableName
	Dim pstrTempMessage
	Dim pstrLocalErrorMessage
	Dim pblnSuccess
	Dim plngTableCounter
	Dim plngRecordCounter
	Dim paryRecordInsertions

	'Define the upgrades
	
	Dim parydbUpgrades(11)
	'contains array of table name, tableID (for new tables only), array of fieldDefinitions, array of records to insert (optional)
	'fieldDefinitions is array of action, fieldName, ACCESS specific field type, original field type
	
	'sfAdmin
	ReDim paryFields(2)			
		paryFields(0) = Array("ADD","adminBackOrderBilling","long","")
		paryFields(1) = Array("ADD","ltlUN","varchar(50)","")
		paryFields(2) = Array("ADD","ltlEMail","varchar(50)","")
	parydbUpgrades(0) = Array("sfAdmin", "", paryFields)
	'Note sfCouponID isn't part of the actual AE upgrade but is used because of limitation in this upgrade script

	'sfCoupons
	Call ResetArrayIndex
	ReDim paryFields(7)			
		paryFields(ArrayIndex) = Array("SQL","Create sfCoupons TABLE","CREATE TABLE sfCoupons (cpCouponCode varchar(100), CONSTRAINT sfCoupons_pk Primary Key (cpCouponCode))")
		paryFields(ArrayIndex) = Array("ADD","cpActivate","long","")
		paryFields(ArrayIndex) = Array("ADD","cpNeverExpire","long","")
		paryFields(ArrayIndex) = Array("ADD","cpDescription","char(100)","")
		paryFields(ArrayIndex) = Array("ADD","cpValue","long","")
		paryFields(ArrayIndex) = Array("ADD","cpType","char(10)","")
		paryFields(ArrayIndex) = Array("ADD","cpExpirationDate","date","")
		paryFields(ArrayIndex) = Array("ADD","cpMin","long","")
	parydbUpgrades(1) = Array("sfCoupons", "", paryFields)

	'sfGiftWraps
	Call ResetArrayIndex
	ReDim paryFields(2)			
		paryFields(ArrayIndex) = Array("SQL","Create sfGiftWraps TABLE","CREATE TABLE sfGiftWraps (gwProdID varchar(50), CONSTRAINT sfGiftWraps_pk Primary Key (gwProdID))")
		paryFields(ArrayIndex) = Array("ADD","gwActivate","long","")
		paryFields(ArrayIndex) = Array("ADD","gwPrice","varchar(20)","")
	parydbUpgrades(2) = Array("sfGiftWraps", "", paryFields)

	'sfInventory
	Call ResetArrayIndex
	ReDim paryFields(4)			
		paryFields(ArrayIndex) = Array("ADD","invenProdId","varchar(50)","")
		paryFields(ArrayIndex) = Array("ADD","invenAttDetailID","varchar(255)","")
		paryFields(ArrayIndex) = Array("ADD","invenAttName","memo","")
		paryFields(ArrayIndex) = Array("ADD","invenInStock","long","")
		paryFields(ArrayIndex) = Array("ADD","invenLowFlag","long","")
	parydbUpgrades(3) = Array("sfInventory", "invenId", paryFields)

	'sfInventoryInfo
	Call ResetArrayIndex
	ReDim paryFields(6)			
		paryFields(ArrayIndex) = Array("SQL","Create sfInventoryInfo TABLE","CREATE TABLE sfInventoryInfo (invenProdId varchar(50), CONSTRAINT sfInventoryInfo_pk Primary Key (invenProdId))")
		paryFields(ArrayIndex) = Array("ADD","invenbBackOrder","long","")
		paryFields(ArrayIndex) = Array("ADD","invenbTracked","long","")
		paryFields(ArrayIndex) = Array("ADD","invenbStatus","long","")
		paryFields(ArrayIndex) = Array("ADD","invenbNotify","long","")
		paryFields(ArrayIndex) = Array("ADD","invenInStockDEF","long","")
		paryFields(ArrayIndex) = Array("ADD","invenLowFlagDEF","long","")
	parydbUpgrades(4) = Array("sfInventoryInfo", "", paryFields)

	'sfMTPrices
	Call ResetArrayIndex
	ReDim paryFields(4)			
		paryFields(ArrayIndex) = Array("SQL","Create sfMTPrices TABLE","CREATE TABLE sfMTPrices (mtProdID varchar(50), mtQuantity long, CONSTRAINT sfMTPrices_pk Primary Key (mtProdID, mtQuantity))")
		paryFields(ArrayIndex) = Array("ADD","mtIndex","long","")
		paryFields(ArrayIndex) = Array("ADD","mtValue","long","")
		paryFields(ArrayIndex) = Array("ADD","mtType","varchar(10)","")
		paryFields(ArrayIndex) = Array("ADD","mtPLValue","memo","")
	parydbUpgrades(5) = Array("sfMTPrices", "", paryFields)

	'sfOrderDetailsAE
	Call ResetArrayIndex
	ReDim paryFields(4)			
		paryFields(ArrayIndex) = Array("SQL","Create sfOrderDetailsAE TABLE","CREATE TABLE sfOrderDetailsAE (odrdtAEID long, CONSTRAINT sfOrderDetailsAE_pk Primary Key (odrdtAEID))")
		paryFields(ArrayIndex) = Array("ADD","odrdtGiftWrapPrice","varchar(50)","")
		paryFields(ArrayIndex) = Array("ADD","odrdtGiftWrapQTY","long","")
		paryFields(ArrayIndex) = Array("ADD","odrdtBackOrderQTY","long","")
		paryFields(ArrayIndex) = Array("ADD","odrdtAttDetailID","varchar(255)","")
	parydbUpgrades(6) = Array("sfOrderDetailsAE", "", paryFields)

	'sfOrdersAE
	Call ResetArrayIndex
	ReDim paryFields(4)			
		paryFields(ArrayIndex) = Array("SQL","Create sfOrdersAE TABLE","CREATE TABLE sfOrdersAE (orderAEID long, CONSTRAINT sfOrdersAE_pk Primary Key (orderAEID))")
		paryFields(ArrayIndex) = Array("ADD","orderCouponCode","varchar(100)","")
		paryFields(ArrayIndex) = Array("ADD","orderBillAmount","varchar(50)","")
		paryFields(ArrayIndex) = Array("ADD","orderBackOrderAmount","varchar(50)","")
		paryFields(ArrayIndex) = Array("ADD","orderCouponDiscount","varchar(50)","")
	parydbUpgrades(7) = Array("sfOrdersAE", "", paryFields)

	'sfSub_Categories
	Call ResetArrayIndex
	ReDim paryFields(9)			
		paryFields(ArrayIndex) = Array("ADD","subcatCategoryId","long","")
		paryFields(ArrayIndex) = Array("ADD","subcatName","varchar(50)","")
		paryFields(ArrayIndex) = Array("ADD","subcatDescription","varchar(255)","")
		paryFields(ArrayIndex) = Array("ADD","subcatImage","varchar(50)","")
		paryFields(ArrayIndex) = Array("ADD","subcatIsActive","long","")
		paryFields(ArrayIndex) = Array("ADD","CatHierarchy","varchar(50)","")
		paryFields(ArrayIndex) = Array("ADD","Depth","long","")
		paryFields(ArrayIndex) = Array("ADD","HasProds","byte","")
		paryFields(ArrayIndex) = Array("ADD","bottom","byte","")
		paryFields(ArrayIndex) = Array("ADD","subcatHttpAdd","varchar(255)","")
	parydbUpgrades(8) = Array("sfSub_Categories", "subcatID", paryFields)

	'sfSubCatDetail
	Call ResetArrayIndex
	ReDim paryFields(3)			
		paryFields(ArrayIndex) = Array("ADD","subcatCategoryId","long","")
		paryFields(ArrayIndex) = Array("ADD","ProdID","varchar(50)","")
		paryFields(ArrayIndex) = Array("ADD","ProdName","varchar(200)","")
		paryFields(ArrayIndex) = Array("ADD","sortCatDetail","long","")
	parydbUpgrades(9) = Array("sfSubCatDetail", "subcatDetailID", paryFields)

	'sfTmpOrderDetailsAE
	Call ResetArrayIndex
	ReDim paryFields(2)			
		paryFields(ArrayIndex) = Array("SQL","Create sfTmpOrderDetailsAE TABLE","CREATE TABLE sfTmpOrderDetailsAE (odrdttmpAEID long, CONSTRAINT sfTmpOrderDetailsAE_pk Primary Key (odrdttmpAEID))")
		paryFields(ArrayIndex) = Array("ADD","odrdttmpGiftWrapQTY","long","")
		paryFields(ArrayIndex) = Array("ADD","odrdttmpBackOrderQTY","long","")
	parydbUpgrades(10) = Array("sfTmpOrderDetailsAE", "", paryFields)

	'sfTmpOrdersAE
	Call ResetArrayIndex
	ReDim paryFields(1)			
		paryFields(ArrayIndex) = Array("SQL","Create sfTmpOrdersAE TABLE","CREATE TABLE sfTmpOrdersAE (odrtmpSessionID long, CONSTRAINT sfTmpOrdersAE_pk Primary Key (odrtmpSessionID))")
		paryFields(ArrayIndex) = Array("ADD","odrtmpCouponCode","varchar(100)","")
	parydbUpgrades(11) = Array("sfTmpOrdersAE", "", paryFields)


	On Error Resume Next

		pblnSuccess = True
		
		If blnInstall then
		
			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				'------------------------------------------------------------------------------'
				' Create table
				'------------------------------------------------------------------------------'
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(1), pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				End If	'new table check
				
				'------------------------------------------------------------------------------'
				' Add/alter field definitions
				'------------------------------------------------------------------------------'
				
				pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

				'Intermediate error checking
				if Len(pstrLocalErrorMessage) = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
				else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
				end if
				
				'Now check for any records to insert
				If UBound(parydbUpgrades(plngTableCounter)) >= 3 Then
					If Err.Number <> 0 Then err.Clear
				
					paryRecordInsertions = parydbUpgrades(plngTableCounter)(3)
					For plngRecordCounter = 0 To UBound(paryRecordInsertions)
						objCnn.Execute paryRecordInsertions(plngRecordCounter),,128
					Next 'plngRecordCounter

					'Intermediate error checking
					If Err.Number = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully populated</B></li><BR>"
					Else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error populating " & pstrTableName & "</FONT></li>" _
														  & "<li>--Error " & Err.Number & ": " & Err.Description & "</li>"
					End If

				End If

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_ContentManagement)(enName) & " database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error adding " & aryAddons(enAO_ContentManagement)(enName) & " database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		Else

			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					'------------------------------------------------------------------------------'
					' Remove table
					'------------------------------------------------------------------------------'
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "", pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully deleted</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error deleting " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				Else
					'------------------------------------------------------------------------------'
					' Undo Add/alter field definitions
					'------------------------------------------------------------------------------'
					Call setFieldDefinitionsForRemoval(parydbUpgrades(plngTableCounter)(2))
					
					pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table changes successfully removed</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error undoing " & pstrTableName & " table changes</FONT></li>" & pstrLocalErrorMessage
					end if
				
				End If	'new table check
				

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_ContentManagement)(enName) & " database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error removing " & aryAddons(enAO_ContentManagement)(enName) & " database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_SEtoAEUpgrade = pblnSuccess

	End Function	'Install_SEtoAEUpgrade

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Visitor Tracking
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_VisitorTracking(byRef objCnn, byVal blnInstall)

	Dim pstrTableName
	Dim pstrTempMessage
	Dim pstrLocalErrorMessage
	Dim pblnSuccess
	Dim plngTableCounter
	Dim plngRecordCounter
	Dim paryRecordInsertions

	'Define the upgrades
	
	Dim parydbUpgrades(1)
	'contains array of table name, tableID (for new tables only), array of fieldDefinitions, array of records to insert (optional)
	'fieldDefinitions is array of action, fieldName, ACCESS specific field type, original field type
	
	'visitors
	ReDim paryFields(5)			
		paryFields(0) = Array("ADD","visitorSessionID","long")
		paryFields(1) = Array("ADD","visitorCustomerID","long")
		paryFields(2) = Array("ADD","visitorDateCreated","date")
		paryFields(3) = Array("ADD","vistor_HTTP_REFERER","char(255)")
		paryFields(4) = Array("ADD","visitor_REMOTE_ADDR","char(255)")
		paryFields(5) = Array("ADD","visitorLoggedInCustomerID","long")
	parydbUpgrades(0) = Array("visitors", "visitorID", paryFields)

	'visitorPageViews
	ReDim paryFields(7)			
		paryFields(0) = Array("ADD","visitorID","long","")
		paryFields(1) = Array("ADD","PageName","char(255)")
		paryFields(2) = Array("ADD","PageQueryString","char(255)")
		paryFields(3) = Array("ADD","TimeViewed","date")
		paryFields(4) = Array("ADD","PageReferrer","char(255)")
		paryFields(5) = Array("ADD","SearchKeyWords","char(255)")
		paryFields(6) = Array("ADD","SearchResultCount","long","")
		paryFields(7) = Array("ADD","TypeKeyWordSearch","char(5)")
	parydbUpgrades(1) = Array("visitorPageViews", "visitorPageViewID", paryFields)

	On Error Resume Next

		pblnSuccess = True
		
		If blnInstall then
		
			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				'------------------------------------------------------------------------------'
				' Create table
				'------------------------------------------------------------------------------'
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(1), pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				End If	'new table check
				
				'------------------------------------------------------------------------------'
				' Add/alter field definitions
				'------------------------------------------------------------------------------'
				
				pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

				'Intermediate error checking
				if Len(pstrLocalErrorMessage) = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
				else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
				end if
				
				'Now check for any records to insert
				If UBound(parydbUpgrades(plngTableCounter)) >= 3 Then
					If Err.Number <> 0 Then err.Clear
				
					paryRecordInsertions = parydbUpgrades(plngTableCounter)(3)
					For plngRecordCounter = 0 To UBound(paryRecordInsertions)
						objCnn.Execute paryRecordInsertions(plngRecordCounter),,128
					Next 'plngRecordCounter

					'Intermediate error checking
					If Err.Number = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully populated</B></li><BR>"
					Else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error populating " & pstrTableName & "</FONT></li>" _
														  & "<li>--Error " & Err.Number & ": " & Err.Description & "</li>"
					End If

				End If

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_VisitorTracking)(enName) & " database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error adding " & aryAddons(enAO_VisitorTracking)(enName) & " database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		Else

			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					'------------------------------------------------------------------------------'
					' Remove table
					'------------------------------------------------------------------------------'
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "", pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully deleted</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error deleting " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				Else
					'------------------------------------------------------------------------------'
					' Undo Add/alter field definitions
					'------------------------------------------------------------------------------'
					Call setFieldDefinitionsForRemoval(parydbUpgrades(plngTableCounter)(2))
					
					pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table changes successfully removed</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error undoing " & pstrTableName & " table changes</FONT></li>" & pstrLocalErrorMessage
					end if
				
				End If	'new table check
				

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_VisitorTracking)(enName) & " database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error removing " & aryAddons(enAO_VisitorTracking)(enName) & " database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_VisitorTracking = pblnSuccess

	End Function	'Install_VisitorTracking

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	SQLSpeedUpgrade
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_SQLSpeedUpgrade(byRef objCnn, byVal blnInstall)

	Dim i, j
	Dim paryData
	Dim paryFields
	Dim pblnSuccess
	Dim plngMaxLength
	Dim pstrLocalErrorMessage
	Dim plngSplit
	Dim pstrTableName
	Dim pstrTempCode
	Dim pstrTempMessage
	Dim pobjRS
	Dim pstrSQL
	
		pblnSuccess = True
		
		If blnInstall then
			'------------------------------------------------------------------------------'
			' Update sfProducts TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfProducts"
			
			ReDim paryFields(14)			
			paryFields(0) = Array("ALTER","prodDescription",2000)
			paryFields(1) = Array("ALTER","prodMessage",500)
			paryFields(2) = Array("ALTER","relatedProducts",100)
			paryFields(3) = Array("ALTER","prodPLPrice",100)
			paryFields(4) = Array("ALTER","prodPLSalePrice",100)
			paryFields(5) = Array("ALTER","prodAdditionalImages",2000)
			plngSplit = 5
			paryFields(6) = Array("ALTER","prodFileName","varchar(255)")
			paryFields(7) = Array("ALTER","prodName","varchar(255)")
			paryFields(8) = Array("ALTER","prodNamePlural","varchar(255)")
			paryFields(9) = Array("ALTER","prodShortDescription","varchar(255)")
			paryFields(10) = Array("ALTER","prodImageSmallPath","varchar(255)")
			paryFields(11) = Array("ALTER","prodImageLargePath","varchar(255)")
			paryFields(12) = Array("ALTER","prodShip","varchar(15)")
			paryFields(13) = Array("ALTER","prodHandlingFee","varchar(15)")
			paryFields(14) = Array("ALTER","prodSetupFee","varchar(15)")
			
			Set	pobjRS = server.CreateObject("adodb.recordset")
			With pobjRS
				.CursorLocation = 2 'adUseClient
				
				For i = 0 To plngSplit
					pstrSQL = "Select prodID, " & paryFields(i)(1) & " from " & pstrTableName
					.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Not .EOF Then
						paryData = .GetRows()
						plngMaxLength = 0
						For j = 0 To UBound(paryData,2)
							If plngMaxLength < Len(paryData(1,j)) Then
								plngMaxLength = Len(paryData(1,j))
								pstrTempCode = paryData(0,j)
							End If
						Next 'j
						
						If CBool(plngMaxLength < paryFields(i)(2)) Then
							paryFields(i)(2) = "varchar(" & paryFields(i)(2) & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
						Else
							paryFields(i)(2) = "varchar(" & plngMaxLength & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default actual length of " & plngMaxLength & ". Product <em>" & pstrTempCode & "</em> set this length.</li>"
						End If
					End If
					.Close
				Next 'i
				
				For i = plngSplit + 1 To UBound(paryFields)
					pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
				Next 'i
			End With
			Set pobjRS = Nothing
			pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & "changes:<ul>" & pstrLocalErrorMessage & "</li></ul>"

			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfAttributes TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfAttributes"
			
			ReDim paryFields(5)			
			paryFields(0) = Array("ALTER","attrDisplay",300)
			plngSplit = 0
			paryFields(1) = Array("ALTER","attrName","varchar(255)")
			paryFields(2) = Array("ALTER","attrImage","varchar(255)")
			paryFields(3) = Array("ALTER","attrSKU","varchar(255)")
			paryFields(4) = Array("ALTER","attrURL","varchar(255)")
			paryFields(5) = Array("ALTER","attrExtra","varchar(255)")
			
			Set	pobjRS = server.CreateObject("adodb.recordset")
			With pobjRS
				.CursorLocation = 2 'adUseClient
				
				For i = 0 To plngSplit
					pstrSQL = "Select attrProdId, " & paryFields(i)(1) & " from " & pstrTableName
					.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Not .EOF Then
						paryData = .GetRows()
						plngMaxLength = 0
						For j = 0 To UBound(paryData,2)
							If plngMaxLength < Len(paryData(1,j)) Then
								plngMaxLength = Len(paryData(1,j))
								pstrTempCode = paryData(0,j)
							End If
						Next 'j
						
						If CBool(plngMaxLength < paryFields(i)(2)) Then
							paryFields(i)(2) = "varchar(" & paryFields(i)(2) & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
						Else
							paryFields(i)(2) = "varchar(" & plngMaxLength & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default actual length of " & plngMaxLength & ". Product <em>" & pstrTempCode & "</em> set this length.</li>"
						End If
					End If
					.Close
				Next 'i
				
				For i = plngSplit + 1 To UBound(paryFields)
					pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
				Next 'i
			End With
			Set pobjRS = Nothing
			pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & "changes:<ul>" & pstrLocalErrorMessage & "</li></ul>"

			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update sfAttributeDetail TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfAttributeDetail"
			
			ReDim paryFields(9)			
			paryFields(0) = Array("ALTER","attrdtPLPrice",100)
			paryFields(1) = Array("ALTER","attrdtDisplay",500)
			plngSplit = 1
			paryFields(2) = Array("ALTER","attrdtName","varchar(255)")
			paryFields(3) = Array("ALTER","attrdtPrice","varchar(50)")
			paryFields(4) = Array("ALTER","attrdtImage","varchar(255)")
			paryFields(5) = Array("ALTER","attrdtFileName","varchar(255)")
			paryFields(6) = Array("ALTER","attrdtSKU","varchar(50)")
			paryFields(7) = Array("ALTER","attrdtURL","varchar(255)")
			paryFields(8) = Array("ALTER","attrdtExtra","varchar(255)")
			paryFields(9) = Array("ALTER","attrdtExtra1","varchar(255)")
			
			Set	pobjRS = server.CreateObject("adodb.recordset")
			With pobjRS
				.CursorLocation = 2 'adUseClient
				
				For i = 0 To plngSplit
					pstrSQL = "Select attrdtName, " & paryFields(i)(1) & " from " & pstrTableName
					.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Not .EOF Then
						paryData = .GetRows()
						plngMaxLength = 0
						For j = 0 To UBound(paryData,2)
							If plngMaxLength < Len(paryData(1,j)) Then
								plngMaxLength = Len(paryData(1,j))
								pstrTempCode = paryData(0,j)
							End If
						Next 'j
						
						If CBool(plngMaxLength < paryFields(i)(2)) Then
							paryFields(i)(2) = "varchar(" & paryFields(i)(2) & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
						Else
							paryFields(i)(2) = "varchar(" & plngMaxLength & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default actual length of " & plngMaxLength & ". Product <em>" & pstrTempCode & "</em> set this length.</li>"
						End If
					End If
					.Close
				Next 'i
				
				For i = plngSplit + 1 To UBound(paryFields)
					pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
				Next 'i
			End With
			Set pobjRS = Nothing
			pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & "changes:<ul>" & pstrLocalErrorMessage & "</li></ul>"

			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

			'------------------------------------------------------------------------------'
			' Update visitors TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "visitors"
			
			ReDim paryFields(14)			
			paryFields(0) = Array("ALTER","visitorLastSearch",1000)
			plngSplit = 0
			paryFields(1) = Array("ALTER","vistorDiscountCodes","varchar(255)")
			paryFields(2) = Array("ALTER","visitorySelectedFreeProducts","varchar(255)")
			paryFields(3) = Array("ALTER","visitorCertificateCodes","varchar(255)")
			paryFields(4) = Array("ALTER","visitorRecentlyViewedProducts","varchar(255)")
			paryFields(5) = Array("ALTER","visitorCity","varchar(50)")
			paryFields(6) = Array("ALTER","visitorState","varchar(3)")
			paryFields(7) = Array("ALTER","visitorZIP","varchar(10)")
			paryFields(8) = Array("ALTER","visitorCountry","varchar(3)")
			paryFields(9) = Array("ALTER","vistor_HTTP_REFERER","varchar(255)")
			paryFields(10) = Array("ALTER","visitor_REMOTE_ADDR","varchar(255)")
			paryFields(11) = Array("ALTER","visitorPreferredCurrency","varchar(3)")
			paryFields(12) = Array("ALTER","visitorPreferredShippingCode","varchar(65)")
			paryFields(13) = Array("ALTER","visitorInstructions","varchar(255)")
			paryFields(14) = Array("ALTER","visitorPaymentmethod","varchar(20)")
			
			Set	pobjRS = server.CreateObject("adodb.recordset")
			With pobjRS
				.CursorLocation = 2 'adUseClient
				
				For i = 0 To plngSplit
					pstrSQL = "Select visitorID, " & paryFields(i)(1) & " from " & pstrTableName
					.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Not .EOF Then
						paryData = .GetRows()
						plngMaxLength = 0
						For j = 0 To UBound(paryData,2)
							If plngMaxLength < Len(paryData(1,j)) Then
								plngMaxLength = Len(paryData(1,j))
								pstrTempCode = paryData(0,j)
							End If
						Next 'j
						
						If CBool(plngMaxLength < paryFields(i)(2)) Then
							paryFields(i)(2) = "varchar(" & paryFields(i)(2) & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
						Else
							paryFields(i)(2) = "varchar(" & plngMaxLength & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default actual length of " & plngMaxLength & ". Product <em>" & pstrTempCode & "</em> set this length.</li>"
						End If
					End If
					.Close
				Next 'i
				
				For i = plngSplit + 1 To UBound(paryFields)
					pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
				Next 'i
			End With
			Set pobjRS = Nothing
			pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & "changes:<ul>" & pstrLocalErrorMessage & "</li></ul>"

			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_SQLSpeedUpgrade = pblnSuccess

	End Function	'Install_SQLSpeedUpgrade
	
'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	DatabaseSize
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_DatabaseSizeUpgrade(byRef objCnn, byVal blnInstall)

	Dim i, j
	Dim paryData
	Dim paryFields
	Dim pblnSuccess
	Dim plngMaxLength
	Dim pstrLocalErrorMessage
	Dim plngSplit
	Dim pstrTableName
	Dim pstrTempCode
	Dim pstrTempMessage
	Dim pobjRS
	Dim pstrSQL
	
		pblnSuccess = True
		
		If blnInstall then
			'------------------------------------------------------------------------------'
			' Update sfProducts TABLE													   '
			'------------------------------------------------------------------------------'
				
			pstrTableName = "sfProducts"
			
			ReDim paryFields(8)			
			paryFields(0) = Array("ALTER","pageName",1)
			paryFields(1) = Array("ALTER","metaTitle",1)
			paryFields(2) = Array("ALTER","metaDescription",1)
			paryFields(3) = Array("ALTER","metaKeywords",1)
			paryFields(4) = Array("ALTER","prodFileName",1)
			paryFields(5) = Array("ALTER","UpgradeVersion",1)
			paryFields(6) = Array("ALTER","version",1)
			paryFields(7) = Array("ALTER","InstallationHours",1)
			paryFields(8) = Array("ALTER","packageCodes",1)
			plngSplit = 8
			
			Set	pobjRS = server.CreateObject("adodb.recordset")
			With pobjRS
				.CursorLocation = 2 'adUseClient
				
				For i = 0 To plngSplit
					pstrSQL = "Select prodID, " & paryFields(i)(1) & " from " & pstrTableName
					.Open pstrSQL, objCnn, 3, 1	'adOpenStatic, adLockReadOnly
					If Not .EOF Then
						paryData = .GetRows()
						plngMaxLength = 0
						For j = 0 To UBound(paryData,2)
							If plngMaxLength < Len(paryData(1,j)) Then
								plngMaxLength = Len(paryData(1,j))
								pstrTempCode = paryData(0,j)
							End If
						Next 'j
						
						If CBool(plngMaxLength < paryFields(i)(2)) Then
							paryFields(i)(2) = "varchar(" & paryFields(i)(2) & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
						Else
							paryFields(i)(2) = "varchar(" & plngMaxLength & ")"
							pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default actual length of " & plngMaxLength & ". Product <em>" & pstrTempCode & "</em> set this length.</li>"
						End If
					End If
					.Close
				Next 'i
				
				For i = plngSplit + 1 To UBound(paryFields)
					pstrLocalErrorMessage = pstrLocalErrorMessage & "<li>" & paryFields(i)(1) & " set to default maximum length of " & paryFields(i)(2) & "</li>"
				Next 'i
			End With
			Set pobjRS = Nothing
			pstrTempMessage = pstrTempMessage & "<li>" & pstrTableName & "changes:<ul>" & pstrLocalErrorMessage & "</li></ul>"

			pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, paryFields, pstrLocalErrorMessage)

			'Intermediate error checking
			if Len(pstrLocalErrorMessage) = 0 then
				pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
			else
				pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_DatabaseSizeUpgrade = pblnSuccess

	End Function	'Install_DatabaseSizeUpgrade

'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Custom Upgrade
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

	Function Install_CustomUpgrade(byRef objCnn, byVal blnInstall)

	Dim pstrTableName
	Dim pstrTempMessage
	Dim pstrLocalErrorMessage
	Dim pblnSuccess
	Dim plngTableCounter
	Dim plngRecordCounter
	Dim paryRecordInsertions

	'Define the upgrades
	
	Dim parydbUpgrades(0)
	'contains array of table name, tableID (for new tables only), array of fieldDefinitions, array of records to insert (optional)
	'fieldDefinitions is array of action, fieldName, ACCESS specific field type, original field type
	
	Dim sourceTable
	Dim sourceTableNew
	Dim sourceFieldNew
	Dim sourceFieldType
	Dim sourceFieldLength
	Dim fieldAction

		sourceTable = Trim(Request.Form("sourceTable"))
		If Len(sourceTable) = 0 Then sourceTableNew = Trim(Request.Form("sourceTableNew"))
		sourceFieldNew = Trim(Request.Form("sourceFieldNew"))
		sourceFieldType = Trim(Request.Form("sourceFieldType"))
		sourceFieldLength = Trim(Request.Form("sourceFieldLength"))
		fieldAction = Trim(Request.Form("fieldAction"))
		If sourceFieldType = "char" And Len(sourceFieldLength) > 0 Then sourceFieldType = sourceFieldType & "(" & sourceFieldLength & ")"

	If Len(sourceTable) = 0 And Len(sourceTableNew) = 0 Then
		mstrMessage = mstrMessage & "<ul><li><Font Color='Red'>Error in custom upgrade. No table specified.</FONT></li></ul>"
		Install_CustomUpgrade = False
		Exit Function
	ElseIf Len(sourceFieldNew) = 0 Then
		mstrMessage = mstrMessage & "<ul><li><Font Color='Red'>Error in custom upgrade. No field specified.</FONT></li></ul>"
		Install_CustomUpgrade = False
		Exit Function
	End If

	ReDim paryFields(0)
	paryFields(0) = Array(fieldAction, sourceFieldNew, sourceFieldType, "")
	
	If Len(sourceTableNew) > 0 Then
		parydbUpgrades(0) = Array(sourceTableNew, sourceFieldNew, paryFields)
	Else
		parydbUpgrades(0) = Array(sourceTable, "", paryFields)
	End If

	On Error Resume Next

		pblnSuccess = True
		
		If blnInstall then
		
			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				'------------------------------------------------------------------------------'
				' Create table
				'------------------------------------------------------------------------------'
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(1), pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Added</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error adding " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				End If	'new table check
				
				'------------------------------------------------------------------------------'
				' Add/alter field definitions
				'------------------------------------------------------------------------------'
				
				pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

				'Intermediate error checking
				if Len(pstrLocalErrorMessage) = 0 then
					pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " Table Successfully Upgraded</B></li><BR>"
				else
					pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error upgrading " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
				end if
				
				'Now check for any records to insert
				If UBound(parydbUpgrades(plngTableCounter)) >= 3 Then
					If Err.Number <> 0 Then err.Clear
				
					paryRecordInsertions = parydbUpgrades(plngTableCounter)(3)
					For plngRecordCounter = 0 To UBound(paryRecordInsertions)
						objCnn.Execute paryRecordInsertions(plngRecordCounter),,128
					Next 'plngRecordCounter

					'Intermediate error checking
					If Err.Number = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully populated</B></li><BR>"
					Else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error populating " & pstrTableName & "</FONT></li>" _
														  & "<li>--Error " & Err.Number & ": " & Err.Description & "</li>"
					End If

				End If

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_ContentManagement)(enName) & " database modifications successfully added.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error adding " & aryAddons(enAO_ContentManagement)(enName) & " database modifications: </FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		Else

			For plngTableCounter = 0 To UBound(parydbUpgrades)
				pstrTableName = parydbUpgrades(plngTableCounter)(0)
				
				If Len(parydbUpgrades(plngTableCounter)(1)) > 0 Then
					'------------------------------------------------------------------------------'
					' Remove table
					'------------------------------------------------------------------------------'
					pblnSuccess = pblnSuccess And CreateNewTable(objCnn, pstrTableName, "", pstrLocalErrorMessage)
					
					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table successfully deleted</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error deleting " & pstrTableName & "</FONT></li>" & pstrLocalErrorMessage
					end if
				Else
					'------------------------------------------------------------------------------'
					' Undo Add/alter field definitions
					'------------------------------------------------------------------------------'
					Call setFieldDefinitionsForRemoval(parydbUpgrades(plngTableCounter)(2))
					
					pblnSuccess = pblnSuccess And GenericDBUpgrade(objCnn, pstrTableName, parydbUpgrades(plngTableCounter)(2), pstrLocalErrorMessage)

					'Intermediate error checking
					if Len(pstrLocalErrorMessage) = 0 then
						pstrTempMessage = pstrTempMessage & "<li><B>" & pstrTableName & " table changes successfully removed</B></li><BR>"
					else
						pstrTempMessage = pstrTempMessage & "<li><Font Color='Red'>Error undoing " & pstrTableName & " table changes</FONT></li>" & pstrLocalErrorMessage
					end if
				
				End If	'new table check
				

			Next 'plngTableCounter

			'------------------------------------------------------------------------------'
			' Final Error Checking
			'------------------------------------------------------------------------------'
			if pblnSuccess then
				pstrTempMessage = "<LI><B>" & aryAddons(enAO_ContentManagement)(enName) & " database modifications successfully removed.</B><ul>" & pstrTempMessage & "</ul></LI>"
			else
				mblnError = True
				pstrTempMessage = "<LI><Font Color='Red'>Error removing " & aryAddons(enAO_ContentManagement)(enName) & " database modifications:</FONT><ul>" & pstrTempMessage & "</ul></LI>"	
			end if

		End If	'blnInstall

		'------------------------------------------------------------------------------'
		' Record success or failure
		'------------------------------------------------------------------------------'
		mstrMessage = mstrMessage & "<ul>" & pstrTempMessage & "</ul>"
		mblnUpgraded = pblnSuccess
		Install_CustomUpgrade = pblnSuccess

	End Function	'Install_CustomUpgrade


'///////////////////////////////////////////////////////////////////////////////////////////////
'/
'|	Page clean-up
'\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



	On Error Resume Next
	mobjCnn.Close 
	Set mobjCnn = Nothing

'**************************************************************************************************************************************************

Function getAvailableTables

Dim paryTables
Dim pstrOut

	paryTables = getAvailableTables_Array
	For i = 0 To UBound(paryTables)
		pstrOut = pstrOut & "<option value=""" & paryTables(i)(0) & """>" & paryTables(i)(1) & "</option>"
	Next 'i

	getAvailableTables = pstrOut

End Function	'getAvailableTables

'**************************************************************************************************************************************************

Function getAvailableTables_Array

Dim objSchema
Dim pblnInserted
Dim plngPos
Dim pstrTableValue
Dim pstrTableName
Dim pstrType
Dim pstrprevTableName
Dim paryPrimary
Dim paryTemp
Dim plngCounter
Dim plngPointer

	Set objSchema = mobjCnn.OpenSchema(20) 
	'adSchemaTables = 20
	'adSchemaPrimaryKeys = 28

	plngCounter = -1
	Do Until objSchema.EOF
		'debugprint "TABLE_TYPE", objSchema("TABLE_TYPE")
		pstrType = objSchema("TABLE_TYPE")
		Select Case UCase(objSchema("TABLE_TYPE"))
			Case "TABLE", "VIEW"
				pstrTableName = objSchema("TABLE_NAME")
				plngPos = InStr(1, pstrTableName, "$")
				If plngPos > 1 Then pstrTableName = Left(pstrTableName, plngPos-1)
				If pstrprevTableName <> pstrTableName Then
					plngCounter = plngCounter + 1
					pstrTableValue = UCase(objSchema("TABLE_TYPE")) & " - " & pstrTableName
					pstrTableValue = pstrTableName & " (" & objSchema("TABLE_TYPE") & ")" 
					If plngCounter = 0 Then
						ReDim paryPrimary(plngCounter)
						paryPrimary(plngCounter) = Array(pstrTableName, pstrTableValue)
					Else
						ReDim paryTemp(plngCounter)
						plngPointer = 0
						pblnInserted  = False

						For i = 0 To UBound(paryPrimary)
							'If pstrTableValue > paryPrimary(i)(1) And Not pblnInserted Then
							If LCase(pstrTableName) < LCase(paryPrimary(i)(0)) And Not pblnInserted Then
								paryTemp(plngPointer) = Array(pstrTableName, pstrTableValue)
								plngPointer = plngPointer + 1
								pblnInserted = True
							End If
							paryTemp(plngPointer) = paryPrimary(i)
							plngPointer = plngPointer + 1
						Next
						
						'check if item was at bottom of list
						If Not isArray(paryTemp(plngCounter)) Then 
							paryTemp(plngCounter) = Array(pstrTableName, pstrTableValue)
						End If
						
						paryPrimary = paryTemp
					End If					

				End If
				pstrprevTableName = pstrTableName
			
			Case Else
				'do nothing
		End Select
		
		objSchema.MoveNext
	Loop
	
	objSchema.Close 
	Set mobjCnn = objSchema
	
	getAvailableTables_Array = paryPrimary

End Function	'getAvailableTables_Array

%>