<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Revision Date: September 21, 2002			                                *
'*   Version 2.0                                                                *
'*                                                                              *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<% Sub WriteHeader(strBodyOnload, blnDisplayTopMenu) %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= mstrPageTitle %></title>
<style>
BODY
{
    BACKGROUND-ATTACHMENT: fixed;
    BACKGROUND-COLOR: white;
    COLOR: black;
    FONT-FAMILY: Arial, 'Century Gothic'
}
HR
{
    BORDER-BOTTOM: 90ex;
    BORDER-LEFT: 90ex;
    BORDER-RIGHT: 90ex;
    BORDER-TOP: 90ex;
    COLOR: AE7EFE
}
.Selected
{
    background-color: #FFFF00;
}
.Active
{
    background-color: #66CCFF;
}
.Inactive
{
    background-color: #CCCCCC;
}
TD
{
}
TD.label
{
  text-align: right;
}
TD.SectionHeading
{
  text-align: center;
  font-style: italic;
  font-weight: bold;
}
.tdHighlight
{
    background-color: lightsteelblue;
}
.tdNormal
{
    background-color: whitesmoke;
}
	.selector 	{font-family: webdings; padding: 0pt; width: 1em; text-align: center; cursor: default; color:whitesmoke;}
	.pagetitle 	{font-family:Tahoma; font-size :20px; font-weight :bold; color :steelblue; width:500;
	 filter:Shadow(color='lightsteelblue', Direction='135');}
	.tblhdr 	{font-family:Tahoma; font-size :12px; font-weight :normal; color :azure;  background-color:steelblue;}
	.tbl 	{font-family:Tahoma; font-size :11px; font-weight :normal; color :black; border-color:steelblue;}
	.butn 	{font-family:Tahoma; font-size :11px; font-weight :bold; background:steelblue; color :azure; cursor:hand;}
	.editfld 	{font-family:Tahoma; font-size :11px; font-weight :normal; background:whitesmoke; color :black;}
	.errline 	{font-family:Tahoma; font-size :11px; font-weight :bold; color :red;}
	.keyfld {font-family:webdings; font-size :16px; font-weight :normal; color :red;}
	.mandtory {font-family:wingdings; font-size :16px; font-weight :normal; color :red;}
.hdrNonSelected
{
	cursor:hand;
	color:white;
	font-family:Tahoma;
	font-size :12px;
	font-weight :bold;
	background-color:steelblue;
}
.hdrSelected
{
	cursor:hand;
	color:yellow;
	font-family:Tahoma;
	font-size :12px;
	font-weight :bold;
	background-color:steelblue;
}
	.FatalError {font-family:Tahoma; font-size :26px; font-weight :bold; color :red;}
img.Selector
{
	cursor: hand;
	width: 100%;
	height: 100%;
}
.MenuItem
{
    BACKGROUND-COLOR: steelblue;
    COLOR: white;
    FLOAT: left;
    FONT-FAMILY: Arial;
    FONT-SIZE: xx-small;
    FONT-WEIGHT: lighter;
    LEFT: 0pt;
    POSITION: relative;
    TOP: 0pt
}
.ToolBar
{
    FLOAT: left;
    MARGIN: 0px;
    PADDING-BOTTOM: 0px;
    PADDING-LEFT: 0px;
    PADDING-RIGHT: 0px;
    PADDING-TOP: 0px;
    TEXT-ALIGN: right
}
.MenuHeader
{
    BACKGROUND-COLOR: steelblue;
    COLOR: white;
    FONT-FAMILY: Arial;
    FONT-SIZE: xx-small
}
</style>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE="VBSCRIPT">

dim menuHTML
dim currentMenu
dim x,y,x2,y2
dim defaultWidth
dim objCurrentHighLight

function resetColorMenuItem(objSubMenu)
		objSubMenu.style.color="white"
end function

function highLightMenuItem(objSubMenu)
	objSubMenu.style.color="lightBlue"
end function

function hideSubMenu()

  xClick = window.event.clientX
  yClick = window.event.clientY

  if xClick>x and xClick<x2 and yClick>y and yClick<y2 then
 
  else
	  currentMenu.style.display="none"
	  objCurrentHighLight.style.color="white"
  end if
  
  
end function

function showSubMenu(objMenuHeader,objSubMenu)
  if objCurrentHighLight <> "" then
	objCurrentHighLight.style.color="white"
  end if
  objMenuHeader.style.cursor="hand"
  set objCurrentHighLight = objMenuHeader.children("headerText")
  objCurrentHighLight.style.color="lightBlue"
  
  if currentMenu <> "" then
	currentMenu.style.zIndex = 100
	currentMenu.style.display = "none"
  end if
  
  if (objMenuHeader.offsetLeft+objSubMenu.style.pixelWidth)>document.body.clientWidth then
	objSubMenu.style.left=document.body.clientWidth-objSubMenu.style.pixelWidth
  else
    objSubMenu.style.left=objMenuHeader.offsetLeft
  end if
  objSubMenu.style.top=objMenuHeader.offsetTop+14
  objSubMenu.style.display=""
  set currentMenu=objSubMenu
  y = objSubMenu.style.pixelTop-10
  x = objSubMenu.style.pixelLeft 
  'alert(objSubMenu.offsetHeight)
  y2 = eval(objSubMenu.style.pixelTop + objSubMenu.offsetHeight)
  x2 = eval(x + objSubMenu.style.pixelWidth) 

end function

function clearMenu()
	menuHTML=""
	defaultWidth = 100
	currentMenu = ""
	Menu.innerHTML = ""
	objCurrentHighLight = ""
end function

function showMenu()
  Menu.innerHTML="<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=100%><TR bgColor=steelblue><TD><DIV></DIV></TD>" & menuHTML & "</TR></TABLE>"
end function

function createHeader(sId,sHeaderName,sUrl,sTarget)
  dim strHTML

  If len(sUrl) = 0 Then  
	strHTML = "<TD onMouseOver=""showSubMenu(" & sId & ",subMenuItem" & sId & ")"" onMouseOut=""hideSubMenu()""ID=""" & sID & """ CLASS=""MenuItem"" >|&nbsp;&nbsp;<SPAN ID=""headerText"">" & sHeaderName & "</SPAN>&nbsp;&nbsp;</TD>"
  Else
	strHTML = "<TD CLASS=""MenuItem"" >|&nbsp;&nbsp;<SPAN ID=""headerText"">" _
			& "<A ID='AS_" & sId & "'" & _
				"   STYLE='width:" & defaultWidth & ";text-decoration:none;cursor:hand;font-family:Verdana;font-size:xx-small;font-size:10;color:white'" & _
				"   HREF='" & sUrl & "' TARGET='" & sTarget & "' onMouseOver='highLightMenuItem(this)' onMouseOut='resetColorMenuItem(this)'>" & sHeaderName & _
				"   </A></SPAN>&nbsp;&nbsp;</TD>"
  End If
  
  Menu.style.display=""
  menuHTML = menuHTML & strHTML 
end function

function createSubMenu(sId,sHeaderName,sUrl,sTarget)

  dim htmlStr
  dim iPos
  dim strHTML

  htmlStr = subMenu.innerHTML
  iPos = instr(htmlStr,"<!-- submenu" & sId & " -->")
  if sHeaderName = "-" then 
    
    strHTML = "<HR WIDTH=100%><!-- submenu" & sId & " -->" 
    SubMenu.innerHTML=replace(htmlstr,"<!-- submenu" & sID & " -->",strHTML)
  else

   if iPos<=0 then
     strHTML = "<SPAN CLASS='MenuItem' ID='subMenuItem" & sId & "' onMouseOut='hideSubMenu()' STYLE='display:none;width:" & defaultWidth & ";position:absolute;left:0;top:60;padding-top:0;padding-left:0;padding-bottom:20;z-index:118;'>" & _
               "<DIV STYLE='width:" & defaultWidth & ";position:relative;left:0;top:8;z-index:118;' >" & _
				"<A ID='AS_" & sId & "'" & _
				"   STYLE='width:" & defaultWidth & ";text-decoration:none;cursor:hand;font-family:Verdana;font-size:xx-small;font-size:10;color:white'" & _
				"   HREF='" & sUrl & "' TARGET='" & sTarget & "' onMouseOver='highLightMenuItem(this)' onMouseOut='resetColorMenuItem(this)'>" & sHeaderName & _
				"   </A>" & _
				"<!-- submenu" & sId & " --></DIV></SPAN>"
   else
     strHTML =	"<A ID='subMenuRef" & MenuIDStr & "'" & _
	          	"   STYLE='width:" & defaultWidth & ";position:relative;left:0;text-decoration:none;font-family:Verdana;font-size:xx-small;font-size:10;color:white'" & _
		        "   HREF='" & sUrl & "' TARGET='" & sTarget & "' onMouseOver='highLightMenuItem(this)' onMouseOut='resetColorMenuItem(this)'>" & _
		        sHeaderName & "</A><!-- submenu" & sId & " -->"
   end if

   if iPos<=0 then
    SubMenu.innerHTML=SubMenu.innerHTML & strHTML
   else
    SubMenu.innerHTML=replace(htmlstr,"<!-- submenu" & sID & " -->",strHTML)
   end if
  end if
end function

</SCRIPT>
</head>
<% 
If Len(strBodyOnload) > 0 Then
	Response.Write "<body onload=" & Chr(34) & strBodyOnload & Chr(34) & ">"
Else
	Response.Write "<body>"
End If

If blnDisplayTopMenu Then
%>
<table border="1" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td><A HREF="http://www.sandshot.net/"><IMG alt="Sandshot Sofware" border=0 src="Images/logo_blue.gif" width="303" height="88" ></A></td>
  </tr>
</table>
	<span ID="Menu" CLASS="ToolBar" STYLE="display:none"></span>
<span ID="SubMenu">
<script LANGUAGE="VBSCRIPT">
clearMenu
	createHeader "Home","Home","../../../default.asp",""
	createHeader "Administrative","Administration","",""
		createSubMenu "Administrative","Main","Admin.asp",""
		<% If cblnAddon_WebStoreMgr Or cblnAddon_ProductMgr Then %>createSubMenu "Administrative","Product&nbsp;Administration","sfProductAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_ProductPricing Then %>createSubMenu "Administrative","Product&nbsp;Pricing","ssProductPricingTool.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr And cblnSF5AE Then %>createSubMenu "Administrative","Coupon&nbsp;Administration","sfCouponAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Administrative","Clean&nbsp;Up&nbsp;Database","ssDBcleanup.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Administrative","Change&nbsp;Username/Password","Admin.asp?Action=ChangePwd",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_CustomerMgr Then %>createSubMenu "Administrative","Customer&nbsp;Manager","sfCustomerAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_GCMgr Then %>createSubMenu "Administrative","Gift&nbsp;Certificate&nbsp;Manager","ssGiftCertificateAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_PayPalPayments Then %>createSubMenu "Administrative","PayPal&nbsp;Payments","ssPayPalPaymentsAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_PricingLevelMgr Then %>createSubMenu "Administrative","Pricing&nbsp;Level&nbsp;Manager","ssPricingLevelAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_PromoMailMgr Then %>createSubMenu "Administrative","Promotional&nbsp;Mail&nbsp;Manager","ssPromoMailAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_PromoMgr Then %>createSubMenu "Administrative","Promotion&nbsp;Manager","ssPromoAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_OrderMgr Then %>createSubMenu "Administrative","Order&nbsp;Manager","ssOrderAdmin.asp",""<% End If %><%= vbcrlf %>
	createHeader "Reports","Reports","",""
		createSubMenu "Reports","SF&nbsp;Reports","../sfreports.asp",""
	<% If cblnAddon_WebStoreMgr Then %>
	createHeader "Configuration","Website&nbsp;Configuration","",""
		createSubMenu "Configuration","Application&nbsp;Settings","sfAdmin.asp?Show=Application",""
		createSubMenu "Configuration","Design&nbsp;Settings","sfDesignAdmin.asp",""
		createSubMenu "Configuration","Geographical&nbsp;Settings","sfAdmin.asp?Show=Geographical",""
		createSubMenu "Configuration","Mail&nbsp;Settings","sfAdmin.asp?Show=Mail",""
		createSubMenu "Configuration","Search&nbsp;Engine&nbsp;Design&nbsp;Settings","sfTextAdmin.asp",""
		createSubMenu "Configuration","Shipping&nbsp;Settings","sfAdmin.asp?Show=Shipping",""
		createSubMenu "Configuration","Transaction&nbsp;Settings","sfAdmin.asp?Show=Transaction",""
	<% End If %><%= vbcrlf %>

	<% If cblnAddon_WebStoreMgr Or cblnAddon_PostageRate Or cblnAddon_TaxRateMgr Then %>
	createHeader "Support","Settings","",""
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Company&nbsp;Contact&nbsp;Information","sfCompanyAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Configure&nbsp;Affiliates","sfAffiliatesAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr And cblnSF5AE Then %>createSubMenu "Support","Configure&nbsp;Categories","sfCategoryAdminAE.asp",""<% Else %>createSubMenu "Support","Configure&nbsp;Categories","sfCategoryAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Configure&nbsp;Colors","sfColorAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Configure&nbsp;Fonts","sfFontAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Configure&nbsp;Manufacturers","sfManufacturersAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Configure&nbsp;Payment&nbsp;Methods","sfTransactionTypesAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Configure&nbsp;Supported&nbsp;Shipping&nbsp;Carriers","sfShippingAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Configure&nbsp;Value&nbsp;Based&nbsp;Rates","sfValueShippingAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Configure&nbsp;Vendors","sfVendorsAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_PostageRate Then %>createSubMenu "Support","Postage&nbsp;Rate Component","ssShippingMethodAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Supported&nbsp;Countries&nbsp;and&nbsp;Tax&nbsp;Rate","sfLocalesCountryAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_WebStoreMgr Then %>createSubMenu "Support","Supported&nbsp;States/Provinces&nbsp;and&nbsp;Tax&nbsp;Rate","sfLocalesStateAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_TaxRateMgr Then %>createSubMenu "Support","Tax&nbsp;Rate&nbsp;Manager","ssTaxRateAdmin.asp",""<% End If %><%= vbcrlf %>
		<% If cblnAddon_ZBS Then %>createSubMenu "Support","Zone&nbsp;Based&nbsp;Shipping","sszbsZoneAdmin.asp",""<% End If %><%= vbcrlf %>
	<% End If %><%= vbcrlf %>

	<% If cblnUseIntegratedSecurity Then %>
		<% If len(session("login"))> 0 then %>
			createHeader "LogOff","LogOff","Admin.asp?Action=LogOff",""
		<% End If %>
	<% End If %><%= vbcrlf %>
Call showMenu
</script>
</span>
<P><hr>
<% End If	'blnDisplayTopMenu %>
<!-- End Sandshot Header -->
<% End Sub	'WriteHeader 

'Determine which add-ons are installed
'Known possibilities include:

Dim cblnAddon_CustomerMgr		'Customer Manager
Dim cblnAddon_GCMgr				'Gift Certificate Manager
Dim cblnAddon_OrderMgr			'Order Manager
Dim cblnAddon_PayPalPayments	'PayPal Payments
Dim cblnAddon_PostageRate		'Postage Rate Component
Dim cblnAddon_PricingLevelMgr	'Pricing Level Manager
Dim cblnAddon_ProductMgr		'Product Manager
Dim cblnAddon_ProductPricing	'Product Pricing
Dim cblnAddon_PromoMailMgr		'Promotional Mail Manager
Dim cblnAddon_PromoMgr			'Promotion Manager
Dim cblnAddon_TaxRateMgr		'Tax Rate Manager
Dim cblnAddon_WebStoreMgr		'WebStore Manager
Dim cblnAddon_ZBS				'Zone Based Shipping

Call DetermineAddOns

Sub DetermineAddOns

Dim pobjFSO
Dim pstrFilePath

	'On Error Resume Next

	pstrFilePath = Server.MapPath("AdminHeader.asp")
	pstrFilePath = Replace(Lcase(pstrFilePath),"adminheader.asp","")
	'debugprint "pstrFilePath",pstrFilePath

	Set pobjFSO = CreateObject("Scripting.FileSystemObject")
	
	cblnAddon_CustomerMgr = pobjFSO.FileExists(pstrFilePath & "sfCustomerAdmin.asp")					'Customer Manager
	cblnAddon_GCMgr = pobjFSO.FileExists(pstrFilePath & "ssGiftCertificateAdmin.asp")					'Gift Certificate Manager
	cblnAddon_PayPalPayments = pobjFSO.FileExists(pstrFilePath & "ssPayPalPaymentsAdmin.asp")			'PayPal Payments
	cblnAddon_PostageRate = pobjFSO.FileExists(pstrFilePath & "ssPostageRate_shippingMethodsAdmin.asp")	'Postage Rate Component
	cblnAddon_PricingLevelMgr = pobjFSO.FileExists(pstrFilePath & "ssPricingLevelAdmin.asp")			'Pricing Level Manager
	cblnAddon_ProductMgr = pobjFSO.FileExists(pstrFilePath & "sfProductAdmin.asp")						'Product Manager
	cblnAddon_ProductPricing = pobjFSO.FileExists(pstrFilePath & "ssProductPricingTool.asp")					'Product Pricing
	cblnAddon_PromoMailMgr = pobjFSO.FileExists(pstrFilePath & "ssPromoMailAdmin.asp")					'Promotional Mail Manager
	cblnAddon_PromoMgr = pobjFSO.FileExists(pstrFilePath & "ssPromoAdmin.asp")							'Promotion Manager
	cblnAddon_OrderMgr = pobjFSO.FileExists(pstrFilePath & "ssOrderAdmin.asp")							'Order Manager
	cblnAddon_TaxRateMgr = pobjFSO.FileExists(pstrFilePath & "ssTaxRateAdmin.asp")						'Tax Rate Manager
	cblnAddon_WebStoreMgr = pobjFSO.FileExists(pstrFilePath & "sfDesignAdmin.asp")						'WebStore Manager
	cblnAddon_ZBS = pobjFSO.FileExists(pstrFilePath & "sszbsZoneAdmin.asp")								'Zone Based Shipping
	'debugprint "cblnAddon_ZBS",cblnAddon_ZBS

	Set pobjFSO = Nothing

End Sub	'DetermineAddOns
'
%>
