<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Release Version:   2.00.0001												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		November 15, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<!--#include file="ssLibrary/ssmodSF5Addons.asp"-->
<% Sub WriteHeader(strBodyOnload, blnDisplayTopMenu) %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= mstrPageTitle %></title>
<link rel="stylesheet" href="ssLibrary/ssStyleSheet.css" type="text/css">
<script language="javascript" src="ssLibrary/ssFormValidation.js" type="text/javascript"></script>
<script language="javascript" src="ssLibrary/calendar.js" type="text/javascript"></script>
<script language="javascript" src="ssLibrary/sorter.js" type="text/javascript"></script>
<script language="javascript" src="ssLibrary/tipMessage.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">

function showDataEntryTip(theLabel)
{
	stm(tipMessage[theLabel.htmlFor],tipStyle['dataEntry']);
}

	var FiltersEnabled = 0 // if your not going to use transitions or filters in any of the tips set this to 0
	tipStyle['dataEntry']=["white","black","steelblue","whitesmoke","","","","","","","","","","",200,"",2,2,10,10,"","","","simple","gray"]

	//original examples
	// tipStyle[...]=[TitleColor,TextColor,TitleBgColor,TextBgColor,TitleBgImag,TextBgImag,TitleTextAlign,TextTextAlign, TitleFontFace, TextFontFace, TipPosition, StickyStyle, TitleFontSize, TextFontSize, Width, Height, BorderSize, PadTextArea, CoordinateX , CoordinateY, TransitionNumber, TransitionDuration, TransparencyLevel ,ShadowType, ShadowColor]
	tipStyle[0]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,51,1,0,"",""]
	tipStyle[1]=["white","black","#000099","#E8E8FF","","","","","","","center","","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[2]=["white","black","#000099","#E8E8FF","","","","","","","left","","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[3]=["white","black","#000099","#E8E8FF","","","","","","","float","","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[4]=["white","black","#000099","#E8E8FF","","","","","","","fixed","","","",200,"",2,2,1,1,"","","","",""]
	tipStyle[5]=["white","black","#000099","#E8E8FF","","","","","","","","sticky","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[6]=["white","black","#000099","#E8E8FF","","","","","","","","keep","","",200,"",2,2,10,10,"","","","",""]
	tipStyle[7]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,40,10,"","","","",""]
	tipStyle[8]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,50,"","","","",""]
	tipStyle[9]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,51,0.5,75,"simple","gray"]
	tipStyle[10]=["white","black","black","white","","","right","","Impact","cursive","center","",3,5,200,150,5,20,10,0,50,1,80,"complex","gray"]
	tipStyle[11]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,51,0.5,45,"simple","gray"]
	tipStyle[12]=["white","black","#000099","#E8E8FF","","","","","","","","","","",200,"",2,2,10,10,"","","","",""]

</script>
</head>
<% 
If Len(strBodyOnload) > 0 Then
	Response.Write "<body onload=" & Chr(34) & strBodyOnload & Chr(34) & ">"
Else
	Response.Write "<body>"
End If
%>
<div id="TipLayer" style="visibility:hidden;position:absolute;z-index:1000;top:-100"></div>
<%
If blnDisplayTopMenu Then
%>
<div id="topMenu">
	<a HREF="../../../"><IMG alt="Sandshot Sofware" border=0 src="Images/logo.jpg" width="303" height="88" ></a>
	<% If isAllowedToViewReports Then %><a class="topMenu" href="default.asp">Dashboard</a><% End If	'isAllowedToViewReports %>
	<a class="topMenu" href="admin.asp">Main&nbsp;Menu</a>
	<% If isAllowedToViewOrder Then %><a class="topMenu" href="ssOrderAdmin.asp">Orders</a><% End If	'isAllowedToViewOrder %>
	<% If isAllowedToEditProducts Then %><a class="topMenu" href="sfProductAdmin.asp">Catalog</a><% End If	'isAllowedToEditProducts %>
	<a class="topMenu" href="ssHelpFiles/help.htm">Help</a>
	<% If isLoggedIn Then Response.Write " <a class=""topMenu"" href=""Admin.asp?Action=LogOff"">Log Off " & userName & "</a>" %>
</div>
<% End If	'blnDisplayTopMenu %>
<!-- End Sandshot Header -->
<% End Sub	'WriteHeader 

'Determine which add-ons are installed
'Known possibilities include:

Dim cblnAddon_CustomerMgr		'Customer Manager
Dim cblnAddon_GCMgr				'Gift Certificate Manager
'Dim cblnAddon_ImportProducts	'Import Products
Dim cblnAddon_OrderMgr			'Order Manager
Dim cblnAddon_PayPalPayments	'PayPal Payments
Dim cblnAddon_PostageRate		'Postage Rate Component
Dim cblnAddon_PricingLevelMgr	'Pricing Level Manager
Dim cblnAddon_ProductMgr		'Product Manager
'Dim cblnAddon_ProductPricing	'Product Pricing
Dim cblnAddon_PromoMailMgr		'Promotional Mail Manager
Dim cblnAddon_PromoMgr			'Promotion Manager
Dim cblnAddon_PromotionMgrII	'Promotion Manager
'Dim cblnAddon_SalesCentral		'Sales Central
Dim cblnAddon_SiteMonitor		'Site Monitor
Dim cblnAddon_TaxRateMgr		'Tax Rate Manager
Dim cblnAddon_WebStoreMgr		'WebStore Manager
Dim cblnAddon_ZBS				'Zone Based Shipping
Dim cblnAddon_ProductPlacement	'Product Placement
Dim cblnAddon_ProductReview		'Product Review

Call DetermineAddOns

Sub DetermineAddOns

Dim pobjFSO
Dim pstrFilePath

	'On Error Resume Next

	pstrFilePath = ssAdminPath

	Set pobjFSO = CreateObject("Scripting.FileSystemObject")
	
	cblnAddon_CustomerMgr = pobjFSO.FileExists(pstrFilePath & "sfCustomerAdmin.asp")
	cblnAddon_GCMgr = pobjFSO.FileExists(pstrFilePath & "ssGiftCertificateAdmin.asp")
	'cblnAddon_ImportProducts = pobjFSO.FileExists(pstrFilePath & "ssImportProducts.asp")
	cblnAddon_OrderMgr = pobjFSO.FileExists(pstrFilePath & "ssOrderAdmin.asp")
	cblnAddon_PayPalPayments = pobjFSO.FileExists(pstrFilePath & "ssPayPalPaymentsAdmin.asp")
	cblnAddon_PostageRate = pobjFSO.FileExists(pstrFilePath & "ssPostageRate_shippingMethodsAdmin.asp")
	cblnAddon_PricingLevelMgr = pobjFSO.FileExists(pstrFilePath & "ssPricingLevelAdmin.asp")
	'cblnAddon_ProductExport = pobjFSO.FileExists(pstrFilePath & "ssProductExportTool.asp")
	cblnAddon_ProductMgr = pobjFSO.FileExists(pstrFilePath & "sfProductAdmin.asp")
	'cblnAddon_ProductPricing = pobjFSO.FileExists(pstrFilePath & "ssProductPricingTool.asp")
	cblnAddon_PromoMailMgr = pobjFSO.FileExists(pstrFilePath & "ssPromoMailAdmin.asp")
	cblnAddon_PromoMgr = pobjFSO.FileExists(pstrFilePath & "ssPromoAdmin.asp")
	cblnAddon_PromotionMgrII = pobjFSO.FileExists(pstrFilePath & "ssPromotionsAdmin.asp")
	If cblnAddon_PromotionMgrII Then cblnAddon_PromoMgr = False
	'cblnAddon_SalesCentral = pobjFSO.FileExists(pstrFilePath & "ssReports.asp")
	cblnAddon_SiteMonitor = pobjFSO.FileExists(pstrFilePath & "ssSiteMonitor.htm")
	cblnAddon_TaxRateMgr = pobjFSO.FileExists(pstrFilePath & "ssTaxRateAdmin.asp")
	cblnAddon_WebStoreMgr = pobjFSO.FileExists(pstrFilePath & "sfDesignAdmin.asp")
	cblnAddon_ZBS = pobjFSO.FileExists(pstrFilePath & "sszbsZoneAdmin.asp")
	cblnAddon_ProductPlacement = pobjFSO.FileExists(pstrFilePath & "ssProductPlacementAdmin.asp")
	cblnAddon_ProductReview = pobjFSO.FileExists(pstrFilePath & "ssProductReviewsAdmin.asp")
	'debugprint "cblnAddon_ProductPlacement",cblnAddon_ProductPlacement

	Set pobjFSO = Nothing

End Sub	'DetermineAddOns
'
%>