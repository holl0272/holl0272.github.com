<% Option Explicit 
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

Response.Buffer = true
%>
<!--#include file="SSLibrary/modDatabase.asp"-->
<!--#include file="SSLibrary/modLogin.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
dim mstrLoginMessage
dim mstrAction
dim mstrPrevPage

mstrPageTitle = "Sandshot Software WebStore Manager"
mstrAction = Request.QueryString("Action")
if len(mstrAction)=0 then mstrAction = Request.Form("Action")
if mstrAction = "LogOff" then session("login") = ""

If cblnUseIntegratedSecurity Then
	if len(session("login"))=0 then

		mstrPrevPage = Request.Form("PrevPage")
		if len(mstrPrevPage) = 0 then mstrPrevPage = Request.QueryString("PrevPage")

		mstrLoginMessage = ValidUserName
		if  mstrLoginMessage <> "True" then	
				Call WriteHeader("", True)
				Call ShowLoginForm(mstrLoginMessage)
		else
			if len(mstrPrevPage) = 0 then
				Call WriteHeader("", True)
				Call ShowAdminPage
			else
				Response.Clear
				Response.Redirect mstrPrevPage
			end if	
		end if
	else
		if mstrAction = "ChangePwd" then
			if len(Request.Form("Action")) <> 0 then mstrLoginMessage = ChangePassword
			Call WriteHeader("",True)
			Call ShowChangePasswordForm(mstrLoginMessage)
		else
			Call WriteHeader("",True)
			Call ShowAdminPage
		end if
	end if
else
	Call WriteHeader("",True)
	Call ShowAdminPage
end if

Response.Flush

'*******************************************************************************************************************************************

Sub ShowAdminPage
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>WebStore Manager Main Admin Page</title>
</head>
<script LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></script>

<body>
<center>
  <table border="0" cellpadding=3 cellspacing=8 ID="Table1">
    <tr>
      <td width="100%" colspan="2">
        <p align="center"><h2>Welcome to the DemoStore Adminstration Menu</h2>
      <p></p>
      </td>
    </tr>
    <tr>
      <td width="50%"></td>
      <td width="50%"></td>
    </tr>
    <tr>
      <td width="50%" valign="top">
        <p align="left">Website Configuration
        <ul>
        <% If cblnAddon_WebStoreMgr Then %>
          <li><a href="sfDesignAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Design Settings</a>
          <li><a href="sfTextAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Search Engine Design Settings</a>
          <li><a href="sfAdmin.asp?Show=Application" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Application Settings</a>
          <li><a href="sfAdmin.asp?Show=Mail" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Mail Settings</a>
          <li><a href="sfAdmin.asp?Show=Transaction" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Transaction Settings</a>
          <li><a href="sfAdmin.asp?Show=Geographical" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Geographical Settings</a>
          <li><a href="sfAdmin.asp?Show=Shipping" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Shipping Settings</a></li>
          <ul>
          <li><a href="sfValueShippingAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Value Based</a></li>
          <li><a href="sfShippingAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Carrier Based</a></li>
          </ul>
        <% End If %>
        </ul>
      </td>
      <td width="50%" align="left" valign="top" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Administrative Functions
        <ul>
<% If cblnAddon_OrderMgr Then %><li><a href="ssOrderAdmin.asp">Order&nbsp;Manager</a></li><% End If %>
<% If cblnAddon_PayPalPayments Then %><li><a href="ssPayPalPaymentsAdmin.asp">PayPal&nbsp;Payments</a></li><% End If %>
<% If cblnAddon_PromoMgr Then %><li><a href="ssPromoAdmin.asp">Promotion&nbsp;Manager</a></li><% End If %>
<% If cblnAddon_GCMgr Then %><li><a href="ssGiftCertificateAdmin.asp">Gift&nbsp;Certificate&nbsp;Manager</a></li><% End If %>
<% If cblnAddon_CustomerMgr Then %><li><a href="sfCustomerAdmin.asp">Customer&nbsp;Manager</a></li><% End If %>
<% If cblnAddon_PricingLevelMgr Then %><li><a href="ssPricingLevelAdmin.asp">Pricing&nbsp;Level&nbsp;Manager</a></li><% End If %>
<% If cblnAddon_WebStoreMgr OR cblnAddon_ProductMgr Then %><li><a href="sfProductAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Product Administration</a></li><% End If %>
<% If cblnAddon_ProductPricing Then %><li><a href="ssProductPricingTool.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Product Pricing Tool</a></li><% End If %>
<% If cblnAddon_WebStoreMgr Then %><li><a href="ssDBcleanup.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Clean Up Database</a></li><% End If %>
        </ul>
        <p>&nbsp;</p>
<% If cblnAddon_WebStoreMgr Then %>
        <ul>
          <li><a href="Admin.asp?Action=ChangePwd" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Change Username/Password</a>
        </li>
        </ul>
<% End If %>
      </td>
    </tr>
    <tr>
      <td width="50%">Supporting Settings
        <ul>
        <% If cblnAddon_WebStoreMgr Then %>
          <li><a href="sfAffiliatesAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Affiliates</a></li>
        <% If cblnSF5AE Then %>
          <li><a href="sfCategoryAdminAE.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Categories</a></li>
        <% Else %>
          <li><a href="sfCategoryAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Categories</a></li>
        <% End If %>
          <li><a href="sfColorAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Colors</a></li>
          <li><a href="sfCompanyAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Company Contact Information</a></li>
          <li><a href="sfFontAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Fonts</a></li>
          <li><a href="sfManufacturersAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Manufacturers</a></li>
          <li><a href="sfTransactionTypesAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Payment Methods</a></li>
          <li><a href="sfShippingAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Supported Shipping Carriers</a></li>
          <li><a href="sfValueShippingAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Value Based Rates</a></li>
          <li><a href="sfVendorsAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Vendors</a></li>
          <li><a href="sfLocalesCountryAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Supported Countries and Tax Rate</a></li>
          <li><a href="sfLocalesStateAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Supported States/Provinces and Tax Rate</a></li>
        <% End If %>
          
<% If cblnAddon_PostageRate Then %>
<li>Postage&nbsp;Rate&nbsp;Component</li>
<ul>
<li><a href="ssPostageRate_shippingMethodsAdmin.asp">Configure Shipping Methods</a></li>
<li><a href="ssPostageRate_ShippingCarriersAdmin.asp">Configure Shipping Carriers</a></li>
<li><a href="ssPostageRate_ShippingMethodConfigurationAdmin.asp">Set Store Shipping Calculation Method</a></li>
</ul>
<% End If %>
<% If cblnAddon_TaxRateMgr Then %><li><a href="ssTaxRateAdmin.asp">Tax&nbsp;Rate&nbsp;Manager</a></li><% End If %>
<% If cblnAddon_ZBS Then %><li><a href="sszbsZoneAdmin.asp">Zone&nbsp;Based&nbsp;Shipping</a></li><% End If %>
        </ul>
        <p></p>
      </td>
      <td width="50%" valign="top" align="left">Reports
        <ul>
          <li><a href="../sfreports.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Standard SF Reporting Tools</a> 
        </li>
<% If cblnAddon_PromoMailMgr Then %><li><a href="ssPromoMailAdmin.asp">Promotional&nbsp;Mail&nbsp;Manager</a></li><% End If %>
        </ul>
      </td>
    </tr>
    <tr>
      <td width="100%" colspan="2"></td>
    </tr>
  </table>
</center>
</body>

</html>
<% End Sub %>