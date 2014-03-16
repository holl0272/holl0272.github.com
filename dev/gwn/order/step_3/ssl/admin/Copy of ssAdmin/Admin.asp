<% Option Explicit 
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

Response.Buffer = true
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/ssmodAdminReports.asp"-->
<!--#include file="ssReports.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Call CheckLoginStatus_AdminPage(False)
%><!--#include file="adminFooter.asp"--><%
	If Response.Buffer Then Response.Flush
    Call ReleaseObject(cnn)

'*******************************************************************************************************************************************

Dim maryAddOns

Function decrementCounter(byRef lngCounter)
	decrementCounter = lngCounter
	lngCounter = lngCounter - 1
End Function

Sub InitializeAddonReferences

Dim i

	If isArray(maryAddOns) Then Exit Sub

	'array decoder (Display Name, URL)
	i = 17
	ReDim maryAddOns(i)
	maryAddOns(decrementCounter(i)) = Array("Zone Based Shipping Calculator", "")
	maryAddOns(decrementCounter(i)) = Array("WebStore Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Tax Rate Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Sales Central", "")
	maryAddOns(decrementCounter(i)) = Array("Promotional Mail Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Promotion Manager II", "")
	maryAddOns(decrementCounter(i)) = Array("Promotion Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Product Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Product Pricing Tool", "")
	maryAddOns(decrementCounter(i)) = Array("Product Import Tool", "")
	maryAddOns(decrementCounter(i)) = Array("Pricing Level Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Postage Rate Calculator", "")
	maryAddOns(decrementCounter(i)) = Array("PayPal Payment w/ IPN", "")
	maryAddOns(decrementCounter(i)) = Array("Order Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Gift Certificate Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Customer Manager", "")
	maryAddOns(decrementCounter(i)) = Array("Product Placement", "")
	If i <> 0 Then
		Response.Write "<h3><font color=red>Error initializing addon references. Remaining items: " & i & "</font></h3>"
		Response.Flush
	End If
End Sub	'InitializeAddonReferences

Function WriteAddonReference(strAddonName, blnDisplay)

Dim i

	If Not blnDisplay Then Exit Function
	
	Call InitializeAddonReferences

	For i = 1 To UBound(maryAddOns)
		If maryAddOns(i)(0) = strAddonName Then
			WriteAddonReference = "&nbsp;<sup><font size='-1'><a href='#" & maryAddOns(i)(0) & "' title='Included with " & maryAddOns(i)(0) & "'>" & i & "</a></font></sup>"
			Exit For
		End If
	Next 'i

End Function	'WriteAddonReference

'*******************************************************************************************************************************************

Sub ShowAdminPage

Dim pblnShowAddonReferences
Dim i

pblnShowAddonReferences = True
pblnShowAddonReferences = False
%>
  <table border="0" cellpadding="8" cellspacing="0" ID="Table1">
    <tr>
      <th width="100%" colspan="3" align=center>
      <div colspan="2" class="clsCurrentLocation">Administration Menu</div>
      </th>
    </tr>
    <tr>
      <td valign="top">
  <table border="1" cellpadding="8" cellspacing="0">
    <tr>
      <td valign="top">

		<% If isAdmin And False Then %>
		<div class="Section"><span class="SectionHeader"><span class="SectionHeaderTitle">Sandshot Settings</span> <img id="imgSandshotSpecific" src="images/UI_OM_expand.gif" onclick="hidePageSection(this, 'SandshotSpecific');" class="SectionHeaderExpander" /></span>
		  <div id="divSandshotSpecificContent" style="display:none">
        <ul>
          <li><a href="sskbArticlesAdmin.asp" onmouseover="return DisplayTitle(this);" onmouseout="ClearTitle();" title="">Knowledge Base</a></li>
          <li><a href="ssVersionChecks.asp" onmouseover="return DisplayTitle(this);" onmouseout="ClearTitle();" title="">Version Checks</a></li>
          <li><a href="ssInvoiceViews.asp" onmouseover="return DisplayTitle(this);" onmouseout="ClearTitle();" title="">Invoice Views</a></li>
          <li><a href="sskbCategoriesAdmin.asp">KB Categories</a></li>
          <li><a href="sskbsubCategoriesAdmin.asp">KB sub-Categories</a></li>
          <li><a href="sskbArticleTypesAdmin.asp">KB Types</a></li>
          <li><a href="ssProblemReportsAdmin.asp">Problem Reports</a></li>
          <li><a href="ssProductDowloadFileChecker.asp">Verify Files</a></li>
        </ul>
        <% Call showOutstandingProblemReports %>
		  </div>
        </div>
		<% End If	'isAdmin %>

		<% If isAllowedToViewOrder Then %>
		<div class="Section"><span class="SectionHeader"><span class="SectionHeaderTitle">Orders</span> <img id="imgOrderMenu" src="images/UI_OM_expand.gif" onclick="hidePageSection(this, 'OrderMenu');" class="SectionHeaderExpander" /></span>
		  <div id="divOrderMenuContent" style="display:none">
		  <ul>
			<li><a href="ssOrderAdmin.asp">View&nbsp;Orders</a><%= WriteAddonReference("Order Manager", pblnShowAddonReferences) %></li>
			<li><a href="ssOrderAdmin_Process.asp">Process&nbsp;Orders</a><%= WriteAddonReference("Order Manager", pblnShowAddonReferences) %></li>
			<li><a href="ssOrderAdmin.asp?Action=ViewOrder&optDisplay=1&optDate_Filter=0&optPayment_Filter=2&optShipment_Filter=0">Mass&nbsp;Process&nbsp;Payments</a></li>
			<li><a href="ssOrderAdmin.asp?Action=ViewOrder&optDisplay=2&optDate_Filter=0&optPayment_Filter=0&optShipment_Filter=1">Mass&nbsp;Process&nbsp;Shipments</a></li>
			<li><a href="ssOrderAdmin.asp?Action=Filter&optDisplay=3">Import&nbsp;Tracking&nbsp;Numbers</a></li>
			<li><a href="ssOrderAdmin_UpdatePastOrders.asp">Update&nbsp;Old&nbsp;Orders</a></li>
			<li><a href="sfCustomerAdmin.asp">Customers</a><%= WriteAddonReference("Customer Manager", pblnShowAddonReferences) %></li>
		  </ul>
<span style="margin-left: 20;">
<fieldset style="display: inline">
  <legend><a href="ssOrderAdmin.asp">Find Order by Order Number</a></legend>
  <form name="frmQuickOrderManager" id="frmQuickOrderManager" action="ssOrderAdmin.asp" style="display: inline">
	<input type="hidden" name="Action" id="Action" value="Filter">
	<input type="hidden" name="Flag_Voided" id="Flag_Voided" value="0">
	<input type="text" name="Text_Filter" id="Text_Filter" value="">&nbsp;<input type="image" src="../../Images/Buttons/go3.gif" ID="Image1" NAME="Image1"><br />
	<input type="radio" name="optText_Filter" id="optText_Filter0" value="1" checked><label for="optText_Filter0">Order Number</label>
	<input type="radio" name="optText_Filter" id="optText_Filter1" value="4"><label for="optText_Filter1">Email</label>
  </form>
</fieldset>
</span>

		  </div>
		</div>
		<% End If	'isAllowedToViewOrder %>

		<% If isAllowedToEditProducts Then %>
		<div class="Section"><span class="SectionHeader"><span class="SectionHeaderTitle">Products</span> <img id="imgProductMenu" src="images/UI_OM_expand.gif" onclick="hidePageSection(this, 'ProductMenu');" class="SectionHeaderExpander" /></span>
		  <div id="divProductMenuContent" style="display:none">
		  <ul>
			<li><a href="sfProductAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Product Administration</a><%= WriteAddonReference("Product Manager", pblnShowAddonReferences) %></li>
			<li><a href="ssProductPlacementAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Product Placement</a><%= WriteAddonReference("Product Placement", pblnShowAddonReferences) %></li>
			<li><a href="ssProductExportTool.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Product Export</a><%= WriteAddonReference("Product Manager", pblnShowAddonReferences) %></li>
			<li><a href="ssProductPricingTool.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Product Pricing Tool</a><%= WriteAddonReference("Product Pricing Tool", pblnShowAddonReferences) %></li>
			<li><a href="ssImportProducts.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Import Products</a></li>
			<% If cblnSF5AE Then %><li><a href="ssInventoryImportFile.asp">Load Inventory File</a></li><% End If %>
			<li><a href="ssProductImageCheck.asp">Product&nbsp;Image&nbsp;Check</a><%= WriteAddonReference("Product Manager", pblnShowAddonReferences) %></li>
			<li><a href="ssProductReviewsAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Product Reviews</a><%= WriteAddonReference("Product Placement", pblnShowAddonReferences) %></li>
		  </ul>
<fieldset style="display: inline">
  <legend><a href="sfProductAdmin.asp">Find Product (Product Mgr)</a></legend>
  <form name="frmQuickProductManager" id="frmQuickProductManager" action="sfProductAdmin.asp" style="display: inline">
	<input type="hidden" name="Action" id="Hidden1" value="Filter">
	<input type="text" name="TextSearch" id="TextSearch" value="">&nbsp;<input type="image" src="../../Images/Buttons/go3.gif"><br />
	<input type="radio" name="radTextSearch" id="radTextSearch0" value="0" checked><label for="radTextSearch0">Code</label>
	<input type="radio" name="radTextSearch" id="radTextSearch1" value="3"><label for="radTextSearch1">Name</label>
  </form>
</fieldset>

<fieldset style="display: inline">
  <legend><a href="frmQuickProductPricing.asp">Find Product (Pricing Tool)</a></legend>
  <form name="frmQuickProductPricing" id="frmQuickProductPricing" action="ssProductPricingTool.asp" style="display: inline">
	<input type="hidden" name="Action" id="Hidden2" value="Filter">
	<input type="text" name="TextSearch" id="Text1" value="">&nbsp;<input type="image" src="../../Images/Buttons/go3.gif"><br />
	<input type="radio" name="radTextSearch" id="radTextSearch2" value="0" checked><label for="radTextSearch2">Code</label>
	<input type="radio" name="radTextSearch" id="radTextSearch3" value="3"><label for="radTextSearch3">Name</label>
  </form>
</fieldset>

		  </div>
		</div>
		<% End If	'isAllowedToEditProducts %>
		
		<% If isAllowedToViewReports Then %>
		<div class="Section"><span class="SectionHeader"><span class="SectionHeaderTitle">Reports</span> <img id="imgReportMenu" src="images/UI_OM_expand.gif" onclick="hidePageSection(this, 'ReportMenu');" class="SectionHeaderExpander" /></span>
		  <div id="divReportMenuContent" style="display:none">
		  <ul>
			<li><a href="default.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Sales&nbsp;Central</a><%= WriteAddonReference("Sales Central", pblnShowAddonReferences) %></li>
			<li><a href="ssSalesReports.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Sales&nbsp;Report</a><%= WriteAddonReference("Sales Central", pblnShowAddonReferences) %></li>
			<li><a href="ssSalesReports.asp?chkShowReport0=1" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Top&nbsp;Products</a></li>
			<li><a href="ssSalesReports.asp?chkShowReport8=1" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Top&nbsp;Viewed&nbsp;Products</a></li>
			<li><a href="ssSalesReports.asp?chkShowReport7=1" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Top&nbsp;Customers</a></li>
			<li><a href="ssSalesReportAdvanced.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Sales&nbsp;Report, Advanced</a><%= WriteAddonReference("Sales Central", pblnShowAddonReferences) %></li>
			<li><a href="ssAnnualSalesReports.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Annual&nbsp;Reports</a><%= WriteAddonReference("Sales Central", pblnShowAddonReferences) %></li>
			<li><a href="../sfreports.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Standard SF Reporting Tools</a></li>
			<li><a href="ssPromoMailAdmin.asp">Promotional&nbsp;Mail&nbsp;Manager</a><%= WriteAddonReference("Promotional Mail Manager", pblnShowAddonReferences) %></li>
			<li><a href="ssSiteMonitor.htm">Monitor&nbsp;Site</a><%= WriteAddonReference("Promotional Mail Manager", pblnShowAddonReferences) %></li>
			<li><a href="ssPayPalPaymentsAdmin.asp">PayPal&nbsp;Payments</a><%= WriteAddonReference("PayPal Payment w/ IPN", pblnShowAddonReferences) %></li>
			<li><a href="ssActiveShoppingCarts.asp">Active Carts</a></li>
			<li><a href="ssActiveWishLists.asp">Saved Carts</a></li>
			<li><a href="ssActiveNotifications.asp">Notification Requests</a></li>
			<li><a href="ssPromotionsReport.asp">Discounts Usage</a></li>
			<% If cblnSF5AE Then %><li><a href="ssInventoryList.asp">List Inventory</a></li><% End If %>
		  </ul>
		  </div>
		</div>
		<% End If	'isAllowedToViewReports %>

		<% If isAdmin Then %>
		<div class="Section"><span class="SectionHeader"><span class="SectionHeaderTitle">Site Level Settings</span> <img id="imgSiteSettingsMenu" src="images/UI_OM_expand.gif" onclick="hidePageSection(this, 'SiteSettingsMenu');" class="SectionHeaderExpander" /></span>
		  <div id="divSiteSettingsMenuContent" style="display:none">
		  <ul>
			<li><a href="sfAdmin.asp?Show=Application" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Store Configuration Settings</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfColorAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Colors</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfFontAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Fonts</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfDesignAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Design Settings</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfAdmin.asp?Show=Geographical" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Geographical Settings</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfAdmin.asp?Show=Mail" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Mail Settings</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfTextAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Search Result Text Settings</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfAdmin.asp?Show=Shipping" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Shipping Settings</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfLocalesCountryAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Supported Countries and Tax Rate</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfLocalesStateAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Supported States/Provinces and Tax Rate</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
			<li><a href="ssTaxRateAdmin.asp">Local&nbsp;Tax&nbsp;Settings</a><%= WriteAddonReference("Tax Rate Manager", pblnShowAddonReferences) %></li>
			<li><a href="sfAdmin.asp?Show=Transaction" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Transaction Settings</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
		  </ul>
		  </div>
		</div>
		<% End If	'isAdmin %>

		<% If isAdmin Then %>
		<div class="Section"><span class="SectionHeader"><span class="SectionHeaderTitle">Content Administration</span> <img id="imgContentMenu" src="images/UI_OM_expand.gif" onclick="hidePageSection(this, 'ContentMenu');" class="SectionHeaderExpander" /></span>
		  <div id="divContentMenuContent" style="display:none">
		  <ul>
            <li><a href="ssCMS_PageFragmentAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Page Content</a></li>
            <li><a href="ssCMS_ManufacturerAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Manufacturers</a></li>
            <li><a href="ssCMS_ContentTypeAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Content Types</a></li>
		  </ul>
		  </div>
		</div>
		<% End If	'isAdmin %>

		<% If isAdmin Then %>
		<div class="Section"><span class="SectionHeader"><span class="SectionHeaderTitle">Supporting Settings</span> <img id="imgSupportingSettingsMenu" src="images/UI_OM_expand.gif" onclick="hidePageSection(this, 'SupportingSettingsMenu');" class="SectionHeaderExpander" /></span>
		  <div id="divSupportingSettingsMenuContent" style="display:none">
		  <ul>
            <% If isAdmin Then %>
            <li><a href="ssUserAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Configure Site Users</a></li>
            <% End If %>
            <li><a href="sfAffiliatesAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Affiliates</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
            <% If cblnSF5AE Then %>
            <li><a href="sfCategoryAdminAE.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Categories</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
            <% Else %>
            <li><a href="sfCategoryAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Categories</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
            <% End If %>
            <li><a href="sfManufacturersAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Manufacturers</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
            <li><a href="ssPricingLevelAdmin.asp">Pricing&nbsp;Levels</a><%= WriteAddonReference("Pricing Level Manager", pblnShowAddonReferences) %></li>
            <li><a href="sfVendorsAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Vendors</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
		  </ul>
		  </div>
		</div>
		<% End If	'isAdmin %>

		<div class="Section">
		  <span class="SectionHeader">
		  <span class="SectionHeaderTitle">Administrative</span> <img id="imgAdministrativeMenu" src="images/UI_OM_expand.gif" onclick="hidePageSection(this, 'AdministrativeMenu');" class="SectionHeaderExpander" /></span>
		  <div id="divAdministrativeMenuContent" style="display:none">
		  <ul>
          <li><a href="ssPromotionsAdmin.asp">Discounts and Promotions</a></li>
          <li><a href="ssBuyersClubRedemptionAdmin.asp">Buyer's&nbsp;Club&nbsp;Remptions</a></li>
          <li><a href="ssGiftCertificateAdmin.asp">Gift&nbsp;Certificates</a><%= WriteAddonReference("Gift Certificate Manager", pblnShowAddonReferences) %></li>
          <% If isAdmin Then %>
          <% If cblnUseIntegratedSecurity And isAdmin Then %>
          <li><a href="Admin.asp?Action=ChangePwd" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Change Username/Password</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
          <% End If %>
          <li><a href="ssUserAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Users</a> / <a href="ssUserLoginAttemptsAdmin.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Review Login Attempts</a></li>
          <li><a href="ssDBcleanup.asp" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="">Clean Up/Back-up Database</a><%= WriteAddonReference("WebStore Manager", pblnShowAddonReferences) %></li>
          <% End If	'isAdmin %>
		  </ul>
		  </div>
		</div>


      </td>
    </tr>
  </table>
      </td>
      <td valign="top">
      <%
		If isAdmin  Then
			Call showTodaysOrders
			Response.Write "<br />"
			Call showCurrentVisitors
		End If
      %>
      
      <% Call CheckIfCleanDBNeeded(1) %>
      </td>
    </tr>
  </table>

<script language="javascript">
function hidePageSection(theImage, strSectionName)
{
	if (theImage == null) return false;
	if (theImage.src.indexOf("images/UI_OM_expand.gif") > 0)
	{
		theImage.src = "images/UI_OM_collapse.gif";
		setCookie("Display" + strSectionName, 1);
	}else{
		theImage.src = "images/UI_OM_expand.gif"
		deleteCookie("Display" + strSectionName);
	}
	showHideElement(document.getElementById("div" + strSectionName + "Content"))
}

function setPageDisplaySettings()
{
	var strSectionHeader;
	var arySectionHeaders = new Array("OrderMenu","ProductMenu","ReportMenu","SiteSettingsMenu","ContentMenu","SupportingSettingsMenu","AdministrativeMenu");
	
	for (var i = 0;  i < arySectionHeaders.length;  i++)
	{
		strSectionHeader = arySectionHeaders[i];
		if (getCookie("Display" + strSectionHeader) == 1)
		{
			hidePageSection(document.getElementById("img" + strSectionHeader), strSectionHeader);
		}
	}
	
}

setPageDisplaySettings();

</script>

</body>

</html>
<% End Sub %>