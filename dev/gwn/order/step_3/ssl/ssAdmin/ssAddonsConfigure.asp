<%@ Language=VBScript %>
<%Option Explicit
'********************************************************************************
'*   Common Support File For StoreFront 6.0 add-ons
'*   Purpose: Configure add-on settings
'*   Release Version:	2.00.001		
'*   Release Date:		January 1, 2006
'*   Revision Date:		January 1, 2006
'*
'*   Release Notes:
'*
'*   2.00.001 (January 1, 2006)
'*	 - Initial Release
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************

Response.Buffer = True

'********************************************************************************

Sub ShowVersionCheckStatus(byVal strAddon)

Dim cblnShowAlreadyRegisteredVersionForm
Dim pstrMessage
Dim pstrPurchaseKey
Dim pstrExpirationDate
Dim pblnValidRegistration

	cblnShowAlreadyRegisteredVersionForm = False
	pblnValidRegistration = False
	If mclsSSAddonConfig.isAddonInstalled(strAddon) Then
%>
		<table class="tbl" cellpadding="1" cellspacing="0" border="1" rules="none">
		  <colgroup align="right" valign="top">
		  <colgroup align="left" valign="top">
		  <tr>
			<td colspan="2" align="center">
			<%
				pstrMessage = mclsSSAddonConfig.GetValue(strAddon, "versionCheckMessage")
				pstrPurchaseKey = mclsSSAddonConfig.GetValue(strAddon, "purchaseKey")
				
				If Len(pstrPurchaseKey) = 0 Then
					pstrExpirationDate = getAddonConfigurationSetting(strAddon, "expirationDate")
					If isDate(pstrExpirationDate) Then
						If CDate(pstrExpirationDate) < Date() Then
							Response.Write "<h4><font color=""red"">Unregistered Version - Expired on " & FormatDateTime(pstrExpirationDate) & "</font></h4>"
						Else
							Response.Write "<font color=""red"">Unregistered Version - Expires on " & FormatDateTime(pstrExpirationDate) & "</font><br />"
						End If
					Else
						Response.Write "<font color=""red"">Unregistered Version</font><br />"
					End If
				Else
					Response.Write "<h4>Registered</h4>"
					pblnValidRegistration = True
				End If
				If Len(pstrMessage) > 0 Then Response.Write pstrMessage
				
				If Not pblnValidRegistration Then cblnShowAlreadyRegisteredVersionForm = True
			%>
			</td>
		  </tr>
		  <% If cblnShowAlreadyRegisteredVersionForm Then %>
		  <tr>
			<td><label for="<%= strAddon %>PurchaseEmail" title="The billing email address used for this order">Purchase Email</label>:&nbsp;</td>
			<td><input name="<%= strAddon %>PurchaseEmail" id="<%= strAddon %>PurchaseEmail" value="<%= mclsSSAddonConfig.GetValue(strAddon, "purchaseEmail") %>" ></td>
		  </tr>
		  <tr>
			<td><label for="<%= strAddon %>PurchaseOrderNumber" title="The order number">Order Number</label>:&nbsp;</td>
			<td><input name="<%= strAddon %>PurchaseOrderNumber" id="<%= strAddon %>PurchaseOrderNumber" value="<%= mclsSSAddonConfig.GetValue(strAddon, "purchaseOrderNumber") %>" ></td>
		  </tr>
		  <tr>
			<td>&nbsp;</td>
			<td>
			  <% If Len(mclsSSAddonConfig.GetValue(strAddon, "purchaseKey")) = 0 Then %>
			  <input type="submit" class="butn" name="btnRegisterAddon" id="btnRegisterAddon<%= strAddon %>" value="Register" onclick="return registerAddon('<%= strAddon %>');">
			  <% Else %>
			  <input type="submit" class="butn" name="btnRegisterAddon" id="btnRegisterAddon<%= strAddon %>" value="Re-Register" onclick="return registerAddon('<%= strAddon %>');">
			  <% End If %>
			</td>
		  </tr>
		  <% End If	'cblnShowRegisteredVersionForm %>
		</table>
<%
	Else
		Response.Write "Add-on not installed"
	End If	'
	
End Sub	'ShowVersionCheckStatus

'********************************************************************************

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mblnSaveError
Dim mstrAction
Dim mstrAddonToRegister
Dim mstrSelectedCategory

	mstrPageTitle = "Sandshot Software Add-ons Configuration Editor"

	mblnSaveError = False
	
    mstrAction = LoadRequestValue("Action")
    mstrAddonToRegister = LoadRequestValue("AddonRegistrationCode")

    mstrSelectedCategory = LoadRequestValue("Category")
    If Len(mstrSelectedCategory) = 0 Then mstrSelectedCategory = "Addons"

    Set mclsSSAddonConfig = New clsSSAddonConfig
    
    Select Case mstrAction
        Case "Update"
            mblnSaveError = Not mclsSSAddonConfig.Update
			mclsSSAddonConfig.Load
        Case "Register"
            mblnSaveError = Not mclsSSAddonConfig.Update
			mclsSSAddonConfig.Load
			Application.Contents.Remove(getVersionCheckKey(getAddonConfigurationSetting(mstrAddonToRegister, "addonCode")))
			Call setAddonConfigurationSetting(mstrAddonToRegister, "purchaseEmail", Request.Form(mstrAddonToRegister & "purchaseEmail"))
			Call setAddonConfigurationSetting(mstrAddonToRegister, "purchaseOrderNumber", Request.Form(mstrAddonToRegister & "purchaseOrderNumber"))
			Call CheckForUpdatedVersion(getAddonConfigurationSetting(mstrAddonToRegister, "addonCode"), "NewRegistration")
        Case "CheckForUpdates"
            mblnSaveError = Not mclsSSAddonConfig.Update
			mclsSSAddonConfig.Load
			If Len(getAddonConfigurationSetting(mstrAddonToRegister, "addonCode")) > 0 Then
				Application.Contents.Remove(getVersionCheckKey(getAddonConfigurationSetting(mstrAddonToRegister, "addonCode")))
				Call CheckForUpdatedVersion(getAddonConfigurationSetting(mstrAddonToRegister, "addonCode"), getAddonConfigurationSetting(mstrAddonToRegister, "versionInstalled"))
			End If
        Case Else
			mclsSSAddonConfig.Load
    End Select
    
    If Len(Request.QueryString) > 0 Then
		Call WriteHeader("body_onload();",False)
		Response.Write "<p>&nbsp;</p>"	'need spacer
    Else
		Call WriteHeader("body_onload();",True)
    End If

    With mclsSSAddonConfig
%>

<SCRIPT LANGUAGE=javascript>
<!--
var aryAddons = new Array("Addons", "ImageChecker", "ProductExport", "ProductImport", "ProductManager", "ProductPricing", "OrderManager", "OrderExport", "SalesCentral", "SalesReport", "ProductPlacement", "PromotionalMail");

function body_onload()
{
	DisplaySection('<%= mstrSelectedCategory %>');
}

function ValidInput(theForm)
{

    return(true);
}

function DisplaySection(strSection)
{

	for (var i=0; i < aryAddons.length;i++)
	{
		if (aryAddons[i] == strSection)
		{
			if (document.all("tbl" + aryAddons[i]) != null)
			{
				document.all("tbl" + aryAddons[i]).style.display = "";
				document.all("td" + aryAddons[i]).className = "hdrSelected";
			}
		}else{
			if (document.all("tbl" + aryAddons[i]) != null)
			{
				document.all("tbl" + aryAddons[i]).style.display = "none";
				document.all("td" + aryAddons[i]).className = "hdrNonSelected";
			}
		}
	}
		
	document.forms(0).Category.value = strSection
	return(false);
}

function registerAddon(strAddon)
{
	var purchaseEmail = document.all(strAddon + "PurchaseEmail")
	var purchaseOrderNumber = document.all(strAddon + "PurchaseOrderNumber")
	
	if (purchaseEmail.value == "")
	{
		alert("Please enter an email address.");
		purchaseEmail.focus();
		return false;
	}

	if (purchaseOrderNumber.value == "")
	{
		alert("Please enter an order number.");
		purchaseOrderNumber.focus();
		return false;
	}

	frmData.Action.value = "Register";
	frmData.AddonRegistrationCode.value = strAddon;
	frmData.submit();
}

function checkForUpdates(strAddon)
{
	frmData.Action.value = "CheckForUpdates";
	frmData.AddonRegistrationCode.value = strAddon;
	frmData.submit();
	return false;
}

//-->
</SCRIPT>
<script language="vbscript">
dim maryCells()

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function saveFile(byRef strFileName, byRef strContents)
	'Initialize and script ActiveX controls not marked as safe needs to be set to Enable or Promopt

	Dim fso
	Dim MyFile
	Dim pstrErrorMessage
	
		On Error Resume Next
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'trim trailing returns
		'If Right(strContents, 1) = Chr(10) Then strContents = Left(strContents, Len(strContents)-1)
		'If Right(strContents, 1) = Chr(13) Then strContents = Left(strContents, Len(strContents)-1)
		'If Right(strContents, 1) = vbcrlf Then strContents = Left(strContents, Len(strContents)-1)

		If Err.number = 429 Then
			pstrErrorMessage = "You do not have the security settings set properly for this item. " & vbcrlf _
							& "To enable this functionality do the following: "  & vbcrlf _
							& "  - In the Internet Explorer toolbar select Tools --> Internet Options "  & vbcrlf _
							& "  - Select the security tab "  & vbcrlf _
							& "  - Select Custom Level "  & vbcrlf _
							& "  - Find the option 'Initialize and Script ActiveX Components not marked as safe' "  & vbcrlf _
							& "    Change this setting to Prompt "  & vbcrlf _
							& "  - Select OK and OK "
			msgbox(pstrErrorMessage)
		ElseIf Err.number > 0 Then
			pstrErrorMessage = "There was an error opening the file " & pstrFilePath & ". " & vbcrlf _
							& "Error " & Err.number & ": " & Err.Description & vbcrlf
			msgbox(pstrErrorMessage)
		Else
			Set MyFile = fso.CreateTextFile(strFileName, True)
			MyFile.Write strContents
			MyFile.close
			Set MyFile = Nothing
		End If

		Set fso = Nothing
		
		saveFile = (Err.number = 0)

	End Function	'saveFile

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function saveExportFile(strExportSection)

	Dim pobjFileChooser
	Dim pobjExportSection
	Dim pstrFile
	Dim pstrTextToSave

		Set pobjExportSection = document.all(strExportSection)
		Set pobjFileChooser = document.all("tempFile")
		pobjFileChooser.click()

		pstrFile = pobjFileChooser.value
		If Len(pstrFile) > 0 Then
			pstrTextToSave = pobjExportSection.value
			msgbox pstrTextToSave
			Call saveFile(pstrFile, pstrTextToSave)
		End If

		saveExportFile = False

		Set pobjFileChooser = Nothing
		Set pobjExportSection = Nothing
		
		saveExportFile = (Err.number <> 0)

	End Function	'saveExportFile

	'-----------------------------------------------------------------------------------------------------------------------------------------------

</script>

<%= .writeErrorMessages %>
<% If mblnSaveError Then %>
<h4><font color=red>There was an error saving the configuration file. This is likely due to the security settings on your server. You should manually save the file by clicking this link. Note you will need to upload it to the server as you can only save it to your local workstation.</font></h4>
<a href="" onclick="saveExportFile('addonsConfigXML'); return false;">Manually save ssAdmin\exportTemplates\qbTemplates\qbAccountConfiguration.xsl</a><br />
<% End If %>
<FORM action="ssAddonsConfigure.asp" id="frmData" name="frmData" onsubmit="return ValidInput();" method="post">
<input type="hidden" name="Category" id="Category" value="<%= mstrSelectedCategory %>">
<input type="hidden" name="AddonRegistrationCode" id="AddonRegistrationCode" value="">
<input type="hidden" name="Action" id="Action" value="Update">

<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th id="tdAddons" class="hdrSelected" nowrap onclick="return DisplaySection('Addons');">&nbsp;Add-ons&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdImageChecker" class="hdrNonSelected" nowrap onclick="return DisplaySection('ImageChecker');">&nbsp;Image&nbsp;Checker&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdOrderExport" class="hdrSelected" nowrap onclick="return DisplaySection('OrderExport');">&nbsp;Order&nbsp;Export&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdOrderManager" class="hdrSelected" nowrap onclick="return DisplaySection('OrderManager');">&nbsp;Order&nbsp;Manager&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdProductExport" class="hdrNonSelected" nowrap onclick="return DisplaySection('ProductExport');">&nbsp;Product&nbsp;Export&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdProductImport" class="hdrNonSelected" nowrap onclick="return DisplaySection('ProductImport');">&nbsp;Product&nbsp;Import&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdProductManager" class="hdrNonSelected" nowrap onclick="return DisplaySection('ProductManager');">&nbsp;Product&nbsp;Manager&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdProductPlacement" class="hdrNonSelected" nowrap onclick="return DisplaySection('ProductPlacement');">&nbsp;Product&nbsp;Placement&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdProductPricing" class="hdrNonSelected" nowrap onclick="return DisplaySection('ProductPricing');">&nbsp;Product&nbsp;Pricing&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdPromotionalMail" class="hdrNonSelected" nowrap onclick="return DisplaySection('PromotionalMail');">&nbsp;Promotional&nbsp;Mail&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdSalesCentral" class="hdrNonSelected" nowrap onclick="return DisplaySection('SalesCentral');">&nbsp;Sales&nbsp;Central&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdSalesReport" class="hdrNonSelected" nowrap onclick="return DisplaySection('SalesReport');">&nbsp;Sales&nbsp;Report&nbsp;</th>
	<th nowrap width="2pt"></th>

	<th width="90%" align=right>&nbsp;</th>
  </tr>
  <tr>
	<td colspan="24" class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td colspan="24" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
	
<table class="tbl" cellpadding="1" cellspacing="0" border="1" id="tblAddons" rules="none" width="100%">
  <colgroup align="left" valign="top">
  <colgroup align="center" valign="top">
  <colgroup align="center" valign="top">
  <colgroup align="center" valign="top">
  <tr class="tblhdr">
	<th colspan=4 align="left"><a href="admin.asp"><font color="white">Sandshot Software Add-ons</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/help.htm')" name="btnHelp"></th>
  </tr>
  <% If Len(Session("versionCheck_ImmediateMessage")) > 0 Then %>
  <tr>
	<td colspan=4 align="left"><%= Session("versionCheck_ImmediateMessage") %></td>
  </tr>
  <% End If %>
  <tr class="tblhdr">
	<th>Add-on&nbsp;</th>
	<th>Registration&nbsp;Status&nbsp;</th>
	<th>Installed&nbsp;Version&nbsp;</th>
	<th>Updated Last</th>
  </tr>
  <%
  Dim paryAddons
  Dim pstrVersionCheck
  paryAddons = .AddonsArray
  For i = 0 To UBound(.AddonsArray)
	If .isAddonInstalled(i) Then
  %>
  <tr>
    <td><%= paryAddons(i)(enAddon_Name) %></td>
    <td>
    <%
		If Len(.GetValue(AddonConfigurationCode(paryAddons(i)(enAddon_ProductID)), "purchaseKey")) = 0 Then
			Response.Write "<a href=""ssAddonsConfigure.asp?Category=" & paryAddons(i)(enAddon_ConfigCode) & """>Register</a>"
		Else
			Response.Write "Registered"
		End If
	%>
    </td>
    <td><%= .GetValue(AddonConfigurationCode(paryAddons(i)(enAddon_ProductID)), "versionInstalled") %></td>
    <td>
    <% If Len(getAddonConfigurationSetting(paryAddons(i)(enAddon_ConfigCode), "addonCode")) > 0 Then %>
		<a href="" onclick="return checkForUpdates('<%= paryAddons(i)(enAddon_ConfigCode) %>');" title="Check for updates now">
		<%
			pstrVersionCheck = Trim(Application(getVersionCheckKey(paryAddons(i)(enAddon_ProductID))))
			If Len(pstrVersionCheck) = 0 Then
				Response.Write "-"
			ElseIf isDate(pstrVersionCheck) Then
				Response.Write pstrVersionCheck
			Else
				Response.Write pstrVersionCheck
			End If
		%></a>
    <% Else
			If Len(pstrVersionCheck) = 0 Then
				Response.Write "-"
			ElseIf isDate(pstrVersionCheck) Then
				Response.Write pstrVersionCheck
			Else
				Response.Write pstrVersionCheck
			End If
	   End If
	%>
    </td>
  </tr>
  <%
	Else
  %>
  <tr>
    <td><%= paryAddons(i)(enAddon_Name) %></td>
    <td><strong>Not Installed</strong></td>
    <td colspan="2">View the <a href="" onclick="OpenHelp('ssHelpFiles/help.htm'); return false;">Master Help File</a> for more information</td>
  </tr>
  <%
	End If	'.isAddonInstalled(i)
  Next 'i
  %>

  <tr><td colspan="4">&nbsp;</td></tr>
  <tr class="tblhdr">
	<th colspan=4 align="left">StoreFront 5.0 Version Information</th>
  </tr>
  <tr>
    <td colspan="4" align="left"> 
    <%
		If cblnSF5AE Then
			Response.Write "StoreFront 5.0 AE<br />"
		Else
			Response.Write "StoreFront 5.0 SE<br />"
		End If
		If cblnSQLDatabase Then
			Response.Write "Database Type: SQL Server<br />"
		Else
			Response.Write "Database Type: Access<br />"
		End If
    %>
    </td>
  </tr>
</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblImageChecker">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left">
	<a href="ssProductImageCheck.asp"><font color="white">Image Checker Settings</font></a>
	&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ImageChecker/help_ImageChecker.htm')" name="btnHelp"></th>
  </tr>

  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("ImageChecker") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>

  <tr>
	<td>Products to check:&nbsp;</td>
	<td>
		<fieldset>
		<input type="radio" name="cbytImageCheckerCheckActiveProductsOnly" id="cbytImageCheckerCheckActiveProductsOnly2" value="-1" <%= isChecked(.GetValue("ImageChecker", "cbytImageCheckerCheckActiveProductsOnly") = "-1") %>>&nbsp;<label for="cbytImageCheckerCheckActiveProductsOnly2" title="">Active Products Only</label><br />
		<input type="radio" name="cbytImageCheckerCheckActiveProductsOnly" id="cbytImageCheckerCheckActiveProductsOnly0" value="0" <%= isChecked(.GetValue("ImageChecker", "cbytImageCheckerCheckActiveProductsOnly") = "0") %>>&nbsp;<label for="cbytImageCheckerCheckActiveProductsOnly0" title="">Inactive Products Only</label><br />
		<input type="radio" name="cbytImageCheckerCheckActiveProductsOnly" id="cbytImageCheckerCheckActiveProductsOnly1" value="1" <%= isChecked(.GetValue("ImageChecker", "cbytImageCheckerCheckActiveProductsOnly") = "1") %>>&nbsp;<label for="cbytImageCheckerCheckActiveProductsOnly1" title="">Show All</label><br />
		</fieldset>
	</td>
  </tr>

  <tr>
	<td>Products to show:&nbsp;</td>
	<td>
		<fieldset>
		<input type="radio" name="cbytImageCheckerDefaultOnlyShowErrors" id="cbytImageCheckerDefaultOnlyShowErrors2" value="-1" <%= isChecked(.GetValue("ImageChecker", "cbytImageCheckerDefaultOnlyShowErrors") = "-1") %>>&nbsp;<label for="cbytImageCheckerDefaultOnlyShowErrors2" title="">Invalid paths only</label><br />
		<input type="radio" name="cbytImageCheckerDefaultOnlyShowErrors" id="cbytImageCheckerDefaultOnlyShowErrors0" value="0" <%= isChecked(.GetValue("ImageChecker", "cbytImageCheckerDefaultOnlyShowErrors") = "0") %>>&nbsp;<label for="cbytImageCheckerDefaultOnlyShowErrors0" title="">Invalid or undefined paths</label><br />
		<input type="radio" name="cbytImageCheckerDefaultOnlyShowErrors" id="cbytImageCheckerDefaultOnlyShowErrors1" value="1" <%= isChecked(.GetValue("ImageChecker", "cbytImageCheckerDefaultOnlyShowErrors") = "1") %>>&nbsp;<label for="cbytImageCheckerDefaultOnlyShowErrors1" title="">Show All</label><br />
		</fieldset>
	</td>
  </tr>

  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnImageCheckerUpdateImageEvenIfNonExistent" id="cblnImageCheckerUpdateImageEvenIfNonExistent" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ImageChecker", "cblnImageCheckerUpdateImageEvenIfNonExistent"), False)) %>>&nbsp;<label for="cblnImageCheckerUpdateImageEvenIfNonExistent" title="Ordinarily an image will be validated prior to being updated. This setting provides an override so that the image path will be set even if the image itself does not exist on the server.">Update image path even if not present</label></td>
  </tr>

  <tr>
	<td><label for="cstrImageCheckerDefaultSmallImagePattern" title="The pattern to use to auto generate the image. Any field in the products table can be used as long as the field name is wrapped as follows: Ex. {Code}">Small Image Pattern</label>:&nbsp;</td>
	<td><input name="cstrImageCheckerDefaultSmallImagePattern" id="cstrImageCheckerDefaultSmallImagePattern" value="<%= .GetValue("ImageChecker", "cstrImageCheckerDefaultSmallImagePattern") %>" > (Ex. images/{Code}_sm.jpg)</td>
  </tr>

  <tr>
	<td><label for="cstrImageCheckerDefaultLargeImagePattern" title="The pattern to use to auto generate the image. Any field in the products table can be used as long as the field name is wrapped as follows: Ex. {Code}">Large Image Pattern</label>:&nbsp;</td>
	<td><input name="cstrImageCheckerDefaultLargeImagePattern" id="cstrImageCheckerDefaultLargeImagePattern" value="<%= .GetValue("ImageChecker", "cstrImageCheckerDefaultLargeImagePattern") %>" > (Ex. images/{Code}_lg.jpg)</td>
  </tr>

  <tr>
	<td><label for="cstrImageCheckerLargeImageField" title="The field to be used for the large image update">Large image field</label>:&nbsp;</td>
	<td><input name="cstrImageCheckerLargeImageField" id="cstrImageCheckerLargeImageField" value="<%= .GetValue("ImageChecker", "cstrImageCheckerLargeImageField") %>" > (Ex. ImageLargePath)</td>
  </tr>

  <tr>
	<td><label for="cstrImageCheckerSmallImageField" title="The field to be used for the small image update">Small image field</label>:&nbsp;</td>
	<td><input name="cstrImageCheckerSmallImageField" id="cstrImageCheckerSmallImageField" value="<%= .GetValue("ImageChecker", "cstrImageCheckerSmallImageField") %>" > (Ex. ImageSmallPath)</td>
  </tr>

  <tr>
	<td><label for="cstrImageCheckerCustomFields" title="">Custom fields from Products table</label>:&nbsp;</td>
	<td><input name="cstrImageCheckerCustomFields" id="cstrImageCheckerCustomFields" value="<%= .GetValue("ImageChecker", "cstrImageCheckerCustomFields") %>" size="20">&nbsp;(Separated with a leading comma: ex. ", Field1, Field 2")</td>
  </tr>

</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblOrderExport">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="ssOrderAdmin_Export.asp"><font color="white">Order Export Tool Settings</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/OrderManagerExportModule/help_OMExportModule.htm')" name="btnHelp0"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("OrderExport") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>
  <tr>
	<td><label for="cstrDefaultExportFilename" title="The name of the file the download link generates"></label>Default export file name:&nbsp;</td>
	<td><input name="cstrDefaultExportFilename" id="cstrDefaultExportFilename" value="<%= .GetValue("OrderExport", "cstrDefaultExportFilename") %>" ></td>
  </tr>
  <tr>
	<td><label for="cstrInvoiceOrderPrefix" title="Prefix which will appear in front of the custom invoice number">Invoice Prefix</label>:&nbsp;</td>
	<td><input name="cstrInvoiceOrderPrefix" id="cstrInvoiceOrderPrefix" value="<%= .GetValue("OrderExport", "cstrInvoiceOrderPrefix") %>" ></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnUseCustomInvoiceNumber" id="cblnUseCustomInvoiceNumber" value="1" <%= isChecked(ConvertToBoolean(.GetValue("OrderExport", "cblnUseCustomInvoiceNumber"), False)) %>>&nbsp;<label for="cblnUseCustomInvoiceNumber" title="Check to use custom invoice numbers (increment for each export). If unchecked order numbers are used as invoice numbers">Use custom invoice number</label></td>
  </tr>
  <tr>
	<td><label for="clngMaxShortDescriptionLength" title="The maximum length of the short description to make available in the export. Set to 0 for unlimited."></label>Max. short description length:&nbsp;</td>
	<td><input name="clngMaxShortDescriptionLength" id="clngMaxShortDescriptionLength" value="<%= .GetValue("OrderExport", "clngMaxShortDescriptionLength") %>" ></td>
  </tr>
  <tr>
	<td><label for="cbytTaxFraction" title="Used to round tax percentages to a fixed fraction. Value is used to limit available results. Ex. a value of 8 means all tax percentages are fractions of 1/8 of 1 percent"></label>Tax fraction:&nbsp;</td>
	<td><input name="cbytTaxFraction" id="cbytTaxFraction" value="<%= .GetValue("OrderExport", "cbytTaxFraction") %>" ></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnAutoExport" id="cblnAutoExport" value="1" <%= isChecked(ConvertToBoolean(.GetValue("OrderExport", "cblnAutoExport"), False)) %>>&nbsp;<label for="cblnAutoExport" title="">Automatically mark orders as exported</label></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><a href="ssOrderAdmin_ExportQuickBooksConfigure.asp">QuickBooks Configuration Settings</a></td>
  </tr>
</table>
  
<table class="tbl" cellpadding="1" cellspacing="0" border="1" id="tblOrderManager" rules="none" width="100%">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="ssOrderAdmin.asp"><font color="white">Order Manager Settings</font></a>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/OrderManager/help_OrderManager.htm')" name="btnHelp"></th>
  </tr>
  <tr>
    <td colspan="2" align="center">
      <a href="#orderManager_General">General Settings</a> | 
      <a href="#orderManager_Detail">Order Summary Settings</a> | 
      <a href="#orderManager_Summary">Order Detail Settings</a> | 
      <a href="#orderManager_OrderTracking">Order Tracking Settings</a>
    </td>
  </tr>

  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("OrderManager") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>

  <tr class="tblhdr"><td colspan="2" align="center"><a name="orderManager_General"><strong>General Settings</strong></a></td></tr>
  <tr>
	<td>On initially loading Order Manager:&nbsp;</td>
	<td>
		<fieldset>
		<input type="radio" name="cblnShowFilterInitially" id="cblnShowFilterInitially1" value="1" <%= isChecked(ConvertToBoolean(.GetValue("OrderManager", "cblnShowFilterInitially"), False)) %>>&nbsp;<label for="cblnShowFilterInitially1" title="This setting controls if the order summary tab is displayed automatically on the initial page load. If checked the order summary tab will be automatically displayed as soon as the page loads. If it is not checked, the Order Filter tab will automatically be initially selected. Note if you display the Order Summary tab initially you must automatically load the order summaries for any summaries to appear.">Display Order Filter tab</label><br />
		<input type="radio" name="cblnShowFilterInitially" id="cblnShowFilterInitially0" value="0" <%= isChecked(Not ConvertToBoolean(.GetValue("OrderManager", "cblnShowFilterInitially"), False)) %>>&nbsp;<label for="cblnShowFilterInitially0" title="This setting controls if the order summary tab is displayed automatically on the initial page load. If checked the order summary tab will be automatically displayed as soon as the page loads. If it is not checked, the Order Filter tab will automatically be initially selected. Note if you display the Order Summary tab initially you must automatically load the order summaries for any summaries to appear.">Display Order Summary tab</label>
		</fieldset>
	</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnEnableOrderDeletion" id="cblnEnableOrderDeletion" value="1" <%= isChecked(ConvertToBoolean(.GetValue("OrderManager", "cblnEnableOrderDeletion"), False)) %>>&nbsp;<label for="cblnEnableOrderDeletion" title="Check to enable order deletion. If unchecked the mass delete option in the Order Summaries tab and the individual delete button on the Order Detail tab will be unavailable"> Enable order deletion</label></td>
  </tr>
  <tr>
	<td><label for="cstrExportTemplateFolder" title="Directory under the ssAdmin directory where the templates are located. The default location is 'exportTemplates\'">Export template directory</label>:&nbsp;</td>
	<td><input name="cstrExportTemplateFolder" id="cstrExportTemplateFolder" value="<%= .GetValue("OrderManager", "cstrExportTemplateFolder") %>" size="50"></td>
  </tr>

  <tr>
	<td><label for="cstrAdminGeneralLocation" title="Location of the interface file required to send email and decrypt credit card numbers. This file should be named something like ssOrderManager/ssAdmin_General_93729476t9322.aspx where a random number is placed at the end of the file name.">Order Manager interface file: </label></td>
	<td><input name="cstrAdminGeneralLocation" id="cstrAdminGeneralLocation" value="<%= .GetValue("OrderManager", "cstrAdminGeneralLocation") %>" size="50"></td>
  </tr>
  <tr>
	<td>Date Format to Use:&nbsp;</td>
	<td>
		<fieldset>
		<input type="radio" name="cblnEurpoeanFormat" id="cblnEurpoeanFormat1" value="1" <%= isChecked(ConvertToBoolean(.GetValue("OrderManager", "cblnEurpoeanFormat"), False)) %>>&nbsp;<label for="cblnEurpoeanFormat1" title="Force date formatting to use European format. This is primarily used for non-U.S. companies hosting in the U.S.">Use European Format (MM-DD-YY hh:mm)</label><br />
		<input type="radio" name="cblnEurpoeanFormat" id="cblnEurpoeanFormat0" value="0" <%= isChecked(Not ConvertToBoolean(.GetValue("OrderManager", "cblnEurpoeanFormat"), False)) %>>&nbsp;<label for="cblnEurpoeanFormat0" title="Format the date using the default server configuration. This is the usual setting.">Use Default Server Format</label>
		</fieldset>
	</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><a href="ssOrderAdmin_ReportsConfigure.asp">Sales Receipt/Packing Slip Configuration Settings</a></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><a href="ssOrderAdmin_EmailConfigure.asp">Add/Edit/Delete Email Templates</a></td>
  </tr>

  <tr class="tblhdr"><td colspan="2" align="center"><a name="orderManager_Summary"><strong>Order Summary Settings</strong></a></td></tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnAutoShowTable" id="cblnAutoShowTable" value="1" <%= isChecked(.GetValue("OrderManager", "cblnAutoShowTable")) %>>&nbsp;<label for="cblnAutoShowTable" title="This setting controls if the order summary is automatically loaded on the initial page load. If checked the order summary will be automatically loaded on the initial page load. If it not checked, then you must select a filter criteria prior to any orders being loaded for display in the summary section.">Automatically load order summary on initial page load</label></td>
  </tr>
  <tr>
	<td><label for="clngDefaultRecords" title="Order Manager contains a paging utility. This sets the default number of records and the corresponding page size to show. A value of 0 will show all records:">Default records to show: </label></td>
	<td><input name="clngDefaultRecords" id="clngDefaultRecords" value="<%= .GetValue("OrderManager", "clngDefaultRecords") %>" size="5"></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnShowPageNumbers" id="cblnShowPageNumbers" value="1" <%= isChecked(.GetValue("OrderManager", "cblnShowPageNumbers")) %>>&nbsp;<label for="cblnShowPageNumbers" title="Check to show order numbers in the paging section. Leave unchecked to show page numbers.">Show Page numbers instead of order numbers</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnUseOrderFlags" id="cblnUseOrderFlags" value="1" <%= isChecked(.GetValue("OrderManager", "cblnUseOrderFlags")) %>>&nbsp;<label for="cblnUseOrderFlags" title="Uncheck this option to remove the order flag from the order summary table only. It will still be available in for filtering and editing in the detail.">Use order flags</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnUseBackOrder" id="cblnUseBackOrder" value="1" <%= isChecked(.GetValue("OrderManager", "cblnUseBackOrder")) %>>&nbsp;<label for="cblnUseBackOrder" title="Check to enable the ability to manage backorder notifications to your customers. Uncheck if you do not need this feature to reclaim some of the space it utilizes.  Unchecking this will remove this column from the filter, summary table, and detail window.">Use Back Order flags</label></td>
  </tr>
  <tr>
	<td><label for="cstrDisplayMemoField" title="Database field to display as a line of information below your order item in the Order Summary tab. The field ssInternalNotes which was added as part of Order Manager is the default value though it can be any field from the Orders table. The only limitation is if you're using SQL Server it cannot be a text or ntext field.">Additional info to display on Order Summary</label>:&nbsp;</td>
	<td><input name="cstrDisplayMemoField" id="cstrDisplayMemoField" value="<%= .GetValue("OrderManager", "cstrDisplayMemoField") %>" >&nbsp;<em>(Must be a field name in the Orders table</em></td>
  </tr>

  <tr>
	<td><label for="cstrPaymentReceived_True" title="Text to use in the Order Summary section for orders with received payments">Payment Received Text</label>:&nbsp;</td>
	<td><input name="cstrPaymentReceived_True" id="cstrPaymentReceived_True" value="<%= .GetValue("OrderManager", "cstrPaymentReceived_True") %>" ></td>
  </tr>
  <tr>
	<td><label for="cstrPaymentReceived_False" title="Text to use in the Order Summary section for orders which are awaiting payment">Awaiting Payment Text</label>:&nbsp;</td>
	<td><input name="cstrPaymentReceived_False" id="cstrPaymentReceived_False" value="<%= .GetValue("OrderManager", "cstrPaymentReceived_False") %>" ></td>
  </tr>
  <tr>
	<td><label for="cstrOrderShipped_True" title="Text to use in the Order Summary section for orders which have been shipped">Order Shipped Text</label>:&nbsp;</td>
	<td><input name="cstrOrderShipped_True" id="cstrOrderShipped_True" value="<%= .GetValue("OrderManager", "cstrOrderShipped_True") %>" ></td>
  </tr>
  <tr>
	<td><label for="cstrOrderShipped_False" title="Text to use in the Order Summary section for orders which are awaiting shipment">Awaiting Shipment Text</label>:&nbsp;</td>
	<td><input name="cstrOrderShipped_False" id="cstrOrderShipped_False" value="<%= .GetValue("OrderManager", "cstrOrderShipped_False") %>" ></td>
  </tr>

  <tr>
	<td><label for="cstrShippingExportFile" title="XSL template located in the exportTemplates directory">Shipping Template</label>:&nbsp;</td>
	<td><input name="cstrShippingExportFile" id="cstrShippingExportFile" value="<%= .GetValue("OrderManager", "cstrShippingExportFile") %>" size="50"></td>
  </tr>
  <tr>
	<td><label for="cstrPaymentExportFile" title="XSL template located in the exportTemplates directory">Payment Template</label>:&nbsp;</td>
	<td><input name="cstrPaymentExportFile" id="cstrPaymentExportFile" value="<%= .GetValue("OrderManager", "cstrPaymentExportFile") %>" size="50"></td>
  </tr>

  <tr>
	<td><label for="cstrOrdersExtra1_Label" title="Extra field available in the Orders table. Appears in the Order Detail view immediately under external notes. Leave empty to not display.">Custom Orders Label</label>:&nbsp;</td>
	<td><input name="cstrOrdersExtra1_Label" id="cstrOrdersExtra1_Label" value="<%= .GetValue("OrderManager", "cstrOrdersExtra1_Label") %>" size="50"></td>
  </tr>
  <tr>
	<td><label for="cstrInsured_Label" title="Used to indicate package insurance. Leave empty to not display.">Insurance</label>:&nbsp;</td>
	<td><input name="cstrInsured_Label" id="cstrInsured_Label" value="<%= .GetValue("OrderManager", "cstrInsured_Label") %>" size="50"></td>
  </tr>
  <tr>
	<td><label for="cstrPackageWeight_Label" title="Used to manually set a package weight. Leave empty to not display.">Package Weight</label>:&nbsp;</td>
	<td><input name="cstrPackageWeight_Label" id="cstrPackageWeight_Label" value="<%= .GetValue("OrderManager", "cstrPackageWeight_Label") %>" size="50"></td>
  </tr>
  <tr>
	<td><label for="cstrOrderTrackingExtra1_Label" title="Extra field available in the Order Tracking table. Appears in the order tracking section. Leave empty to not display.">Custom Order Tracking Label</label>:&nbsp;</td>
	<td><input name="cstrOrderTrackingExtra1_Label" id="cstrOrderTrackingExtra1_Label" value="<%= .GetValue("OrderManager", "cstrOrderTrackingExtra1_Label") %>" ></td>
  </tr>
  <tr class="tblhdr"><td>&nbsp;</td><td><strong><a name="orderManager_Detail">Order Detail Settings</strong></a></td></tr>
  <tr>
	<td><label for="cstrSalesReceiptTemplate" title="Template you wish for single-click printing of the Sales Receipt. Do not include the template directory in this. It must be a .xsl file.">Sales Receipt Template</label>:&nbsp;</td>
	<td><input name="cstrSalesReceiptTemplate" id="cstrSalesReceiptTemplate" value="<%= .GetValue("OrderManager", "cstrSalesReceiptTemplate") %>" size="50"></td>
  </tr>
  <tr>
	<td><label for="cstrPackingSlipTemplate" title="Template you wish for single-click printing of the Packing Slip. Do not include the template directory in this. It must be a .xsl file.">Packing Slip Template</label>:&nbsp;</td>
	<td><input name="cstrPackingSlipTemplate" id="cstrPackingSlipTemplate" value="<%= .GetValue("OrderManager", "cstrPackingSlipTemplate") %>" size="50"></td>
  </tr>
  <tr>
	<td><label for="cstrChecklistTemplate" title="Template you wish for single-click printing of the order processing. Do not include the template directory in this. It must be a .xsl file.">Checklist Template</label>:&nbsp;</td>
	<td><input name="cstrChecklistTemplate" id="cstrChecklistTemplate" value="<%= .GetValue("OrderManager", "cstrCheckListTemplate") %>" size="50"></td>
  </tr>

  <tr>
	<td><label for="cstrEmailTemplateFolder" title="Directory under the ssAdmin directory where the email templates are located. The default location is 'emailTemplates\'">Email template directory</label>:&nbsp;</td>
	<td><input name="cstrEmailTemplateFolder" id="cstrEmailTemplateFolder" value="<%= .GetValue("OrderManager", "cstrEmailTemplateFolder") %>" size="50"></td>
  </tr>
  <tr>
	<td><label for="cstrDefaultOrderShippedEmailTemplate" title="Email template which will be selected by default. Do not include the template directory in this. It must be a .txt file.">Default Email Template</label>:&nbsp;</td>
	<td><input name="cstrDefaultOrderShippedEmailTemplate" id="cstrDefaultOrderShippedEmailTemplate" value="<%= .GetValue("OrderManager", "cstrDefaultOrderShippedEmailTemplate") %>" size="50"></td>
  </tr>
  <tr>
	<td><label for="cstrDefaultEmailFromAddress" title="From address to use. The Primary Email will be used if this is empty.">Default Email From Address</label>:&nbsp;</td>
	<td><input name="cstrDefaultEmailFromAddress" id="cstrDefaultEmailFromAddress" value="<%= .GetValue("OrderManager", "cstrDefaultEmailFromAddress") %>" size="50"></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="mblnShowFullCountryName" id="mblnShowFullCountryName" value="1" <%= isChecked(ConvertToBoolean(.GetValue("OrderManager", "mblnShowFullCountryName"), False)) %>>&nbsp;<label for="mblnShowFullCountryName" title="Check to to display the full country name instead of the abbreviation">Show Full Country Name</label></td>
  </tr>
  <tr>
	<td>On saving the order detail:&nbsp;</td>
	<td>
		<fieldset>
		<input type="radio" name="onSaveAction" id="onSaveAction0" value="0" <%= isChecked(Not ConvertToBoolean(.GetValue("OrderManager", "cblnAutoShowSummaryOnSave"), False) And Not ConvertToBoolean(.GetValue("OrderManager", "cblnAutoShowFilterOnSave"), False)) %>>&nbsp;<label for="onSaveAction0">Return to Order Detail tab</label><br />
		<input type="radio" name="onSaveAction" id="onSaveAction1" value="1" <%= isChecked(.GetValue("OrderManager", "cblnAutoShowSummaryOnSave")) %>>&nbsp;<label for="onSaveAction1">Display Order Summary tab</label><br />
		<input type="radio" name="onSaveAction" id="onSaveAction2" value="2" <%= isChecked(.GetValue("OrderManager", "cblnAutoShowFilterOnSave")) %>>&nbsp;<label for="onSaveAction2">Display Filter tab</label>
		</fieldset>
  </tr>

<script>
function AddOrderStatus()
{
var pNewRow;
var pNewCell;
var pstrCell1 = "<input type='text' id='orderStatusOption' name='orderStatusOption' value='Order Status Text'>"
var pstrCell2 = "<input type='text' id='orderStatusClass' name='orderStatusClass' value=''>"
var pstrCell3 = "<input type='text' id='orderStatusImage' name='orderStatusImage' value=''><img src='images/delete.gif' onclick='DeleteOrderStatus();'>"      

	pNewRow = document.all("tblOrderStatusInput").insertRow();
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell1;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell2;
	
	pNewCell = pNewRow.insertCell();
	pNewCell.innerHTML = pstrCell3;
	
}

function DeleteOrderStatus(theCell)
{
var ptheRow = window.event.srcElement.parentElement.parentElement;
ptheRow.parentElement.deleteRow(ptheRow.rowIndex);
}

</script>
<tr>
  <td>&nbsp;</td>
  <td>
	<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" id="tblOrderStatusInput">
		<tr class="tblhdr">
			<th colspan="3">Order Status Options</th>
		</tr>
		<tr>
			<td title="Status text to display">Status</td>
			<td title="Class name in stylesheet to use for displaying summary line item">Class</td>
			<td title="Image path; displays next to checkbox in order summaries section">Image</td>
		</tr>
	<%
	Dim i
	Dim maryOrderStatuses
	
	maryOrderStatuses = .GetArray("OrderManager", "orderStatusOptions/orderStatusOption")

		For i = 0 To UBound(maryOrderStatuses)
	%>
		<tr>
			<td><input type="hidden" name="orderStatusID" id="orderStatusID<%= maryOrderStatuses(i)(0) %> %>" size="5" value="<%= maryOrderStatuses(i)(0) %>">
			    <input type="text" name="orderStatusOption" id="orderStatusOption" value="<%= maryOrderStatuses(i)(1) %>"></td>
			<td><input type="text" name="orderStatusClass" id="orderStatusClass<%= maryOrderStatuses(i)(0) %> %>" value="<%= maryOrderStatuses(i)(2) %>"></td>
			<td><input type="text" name="orderStatusImage" id="orderStatusImage<%= maryOrderStatuses(i)(0) %> %>" value="<%= maryOrderStatuses(i)(3) %>"></td>
			<!--<img src="images/delete.gif" onclick="DeleteOrderStatus(this);">-->
		</tr>
	<%	
		Next 'i
	%>
	<tr>
		<td colspan="3" align="center"><input class='butn' id=btnNewOrderStatus name=btnNewOrderStatus type='button' value='New Order Status' onclick="AddOrderStatus();"></td></tr>
	</table>
  </td>
</tr>


  <tr class="tblhdr"><td>&nbsp;</td><td><a name="orderManager_OrderTracking"><strong>Order Tracking Settings</strong></a></td></tr>
  <tr>
	<td>Default Tracking File Format:&nbsp;</td>
	<td>
		<fieldset>
		<input type="radio" name="cstrDefault_ImportProfile" id="cstrDefault_ImportProfile0" value="Default" <%= isChecked(.GetValue("OrderManager", "cstrDefault_ImportProfile") = "Default") %>>&nbsp;<label for="cstrDefault_ImportProfile0" title="">Default</label><br />
		<input type="radio" name="cstrDefault_ImportProfile" id="cstrDefault_ImportProfile1" value="Endicia" <%= isChecked(.GetValue("OrderManager", "cstrDefault_ImportProfile") = "Endicia") %>>&nbsp;<label for="cstrDefault_ImportProfile1" title="">Endicia</label><br />
		<input type="radio" name="cstrDefault_ImportProfile" id="cstrDefault_ImportProfile2" value="Endicia XML" <%= isChecked(.GetValue("OrderManager", "cstrDefault_ImportProfile") = "Endicia XML") %>>&nbsp;<label for="cstrDefault_ImportProfile2" title="">Endicia XML</label><br />
		<input type="radio" name="cstrDefault_ImportProfile" id="cstrDefault_ImportProfile3" value="FedEx" <%= isChecked(.GetValue("OrderManager", "cstrDefault_ImportProfile") = "FedEx") %>>&nbsp;<label for="cstrDefault_ImportProfile3" title="">FedEx</label><br />
		<input type="radio" name="cstrDefault_ImportProfile" id="cstrDefault_ImportProfile4" value="UPS Worldship" <%= isChecked(.GetValue("OrderManager", "cstrDefault_ImportProfile") = "UPS Worldship") %>>&nbsp;<label for="cstrDefault_ImportProfile4" title="">UPS Worldship</label><br />
		</fieldset>
	</td>
  </tr>

  <tr>
	<td>Default Carrier To Use:&nbsp;</td>
	<td>
		<fieldset>
		<input type="radio" name="cstrDefault_CarrierID" id="cstrDefault_CarrierID1" value="1" <%= isChecked(.GetValue("OrderManager", "cstrDefault_CarrierID") = "1") %>>&nbsp;<label for="cstrDefault_CarrierID1" title="">UPS</label><br />
		<input type="radio" name="cstrDefault_CarrierID" id="cstrDefault_CarrierID2" value="2" <%= isChecked(.GetValue("OrderManager", "cstrDefault_CarrierID") = "2") %>>&nbsp;<label for="cstrDefault_CarrierID2" title="">U.S.P.S.</label><br />
		<input type="radio" name="cstrDefault_CarrierID" id="cstrDefault_CarrierID3" value="3" <%= isChecked(.GetValue("OrderManager", "cstrDefault_CarrierID") = "3") %>>&nbsp;<label for="cstrDefault_CarrierID3" title="">FedEx</label><br />
		<input type="radio" name="cstrDefault_CarrierID" id="cstrDefault_CarrierID4" value="4" <%= isChecked(.GetValue("OrderManager", "cstrDefault_CarrierID") = "4") %>>&nbsp;<label for="cstrDefault_CarrierID4" title="">Canada Post</label><br />
		</fieldset>
	</td>
  </tr>

  <tr>
	<td><label for="cTrackingOrderNumberPrefix" title="Prefix which may appear in an order number and should be removed prior to importing the file."> Tracking Order Number Prefix</label>:&nbsp;</td>
	<td><input name="cTrackingOrderNumberPrefix" id="cTrackingOrderNumberPrefix" value="<%= .GetValue("OrderManager", "cTrackingOrderNumberPrefix") %>" size="10"></td>
  </tr>

  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnDefault_SendEmail" id="cblnDefault_SendEmail" value="1" <%= isChecked(ConvertToBoolean(.GetValue("OrderManager", "cblnDefault_SendEmail"), False)) %>>&nbsp;<label for="cblnDefault_SendEmail" title="Check 'Automatically send shipment email' on Import Tracking Numbers tab by default">Check <em>Automatically send shipment email</em> by default</label></td>
  </tr>

</table>
  
<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblProductExport">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="ssProductExportTool.asp"><font color="white">Product Export Settings</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ProductManagerExportModule/help_PMExportModule.htm')" name="btnHelp"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("ProductExport") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>
  <tr>
	<td><label for="cstrDefaultProductExportFilename" title="The name of the file the download link generates"></label>Default export file name:&nbsp;</td>
	<td><input name="cstrDefaultProductExportFilename" id="cstrDefaultProductExportFilename" value="<%= .GetValue("ProductExport", "cstrDefaultProductExportFilename") %>" ></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnDefaultProductExportUseAttributes" id="cblnDefaultProductExportUseAttributes" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductExport", "cblnDefaultProductExportUseAttributes"), False)) %>>&nbsp;<label for="cblnDefaultProductExportUseAttributes" title="Check to include attributes in product export results">Include attributes in results</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnDefaultProductExportUseAllAttributes" id="cblnDefaultProductExportUseAllAttributes" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductExport", "cblnDefaultProductExportUseAllAttributes"), False)) %>>&nbsp;<label for="cblnDefaultProductExportUseAllAttributes" title="Check to include all attribute combinations in product export results. This option overrides the Prevent Duplicates setting">Include all attributes in result set </label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnDefaultProductExportPreventDuplicates" id="cblnDefaultProductExportPreventDuplicates" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductExport", "cblnDefaultProductExportPreventDuplicates"), False)) %>>&nbsp;<label for="cblnDefaultProductExportPreventDuplicates" title="Check this option so that only one product instance will be used. Duplicates arise if attributes are included and when products are assigned to multiple categories. Use all attributes overrides this setting.">Prevent duplicates</label></td>
  </tr>
</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblProductImport">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="ssImportProducts.asp"><font color="white">Product Import Tool Settings</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ProductImport/help_ProductImport.htm')" name="btnHelp"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("ProductImport") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>

  <tr>
	<td><label for="ProductImportScriptTimeout" title="This is the maximum time the page can run. Longer times result in faster imports. Notes: - IIS may be configured such there is an additional limit in place in place which overrides this value. Consult your host for specific values. - SQL Server databases may terminate a connection prior to the page completing its actions. This setting is controlled by your database administrator.">Single Page Load Time</label>:&nbsp;</td>
	<td><input name="ProductImportScriptTimeout" id="ProductImportScriptTimeout" value="<%= .GetValue("ProductImport", "ProductImportScriptTimeout") %>" size="5"></td>
  </tr>
  <tr>
	<td><label for="ProductImportBufferTimeout" title="This is the maximum estimated time it will take for a single product to import. Note: if your products have a lot of attributes this should be higher or if your running SQL Server. This acts as a safety factor to the page and subtracts from the Single Page Load Time.">Single Product Load Time</label>:&nbsp;</td>
	<td><input name="ProductImportBufferTimeout" id="ProductImportBufferTimeout" value="<%= .GetValue("ProductImport", "ProductImportBufferTimeout") %>" size="5"></td>
  </tr>
  <tr>
	<td><label for="ProductImportPageRefreshDelay" title="The time in seconds the page will wait before reposting on large imports; longer delays slow the import but have been known to be necessary on some servers to let them 'catch up'">Page Refresh Delay</label>:&nbsp;</td>
	<td><input name="ProductImportPageRefreshDelay" id="ProductImportPageRefreshDelay" value="<%= .GetValue("ProductImport", "ProductImportPageRefreshDelay") %>" size="5"></td>
  </tr>
  <tr>
	<td><label for="ProductImportWhatIsTrue" title="The import tool can map many import values to a True/False condition. Notes: - All text entries must be in lower case and the import test is case insensitive - Entries must be in lower case and separated with a semi-colon. No spaces should be entered after the semi-colon. - If an entry is in your data which does not map to one of the values below it will map to False - The above discussion refers to True/False values and/or fields which map to 0/1 only">What is true?</label>:&nbsp;</td>
	<td><input name="ProductImportWhatIsTrue" id="ProductImportWhatIsTrue" value="<%= .GetValue("ProductImport", "ProductImportWhatIsTrue") %>" size="25"> (Ex. 1;true;yes;y;on;-1;active)</td>
  </tr>

  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductImportAssignProductsToAllCategoryLevels" id="ProductImportAssignProductsToAllCategoryLevels" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductImport", "ProductImportAssignProductsToAllCategoryLevels"), False)) %>>&nbsp;<label for="ProductImportAssignProductsToAllCategoryLevels" title="When importing a product to a sub-category you have the option to assign it to all levels of the category tree. In earlier versions of StoreFront this was necessary for the product to appear in the search results at the top-level category. It is not necessary to do this in later versions.">Assign Products At All Category Levels</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductImportNumericCategoriesAreUID" id="ProductImportNumericCategoriesAreUID" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductImport", "ProductImportNumericCategoriesAreUID"), False)) %>>&nbsp;<label for="ProductImportNumericCategoriesAreUID" title="The import tool automatically treats any numeric entry as a uid and will assign the product to the corresponding entry. If this is unchecked than any numeric entry in the data source will be treated as a name.">Treat Numeric Category Entries as IDs</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductImportNumericManufacturersAreUID" id="ProductImportNumericManufacturersAreUID" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductImport", "ProductImportNumericManufacturersAreUID"), False)) %>>&nbsp;<label for="ProductImportNumericManufacturersAreUID" title="The import tool automatically treats any numeric entry as a uid and will assign the product to the corresponding entry. If this is unchecked than any numeric entry in the data source will be treated as a name.">Treat Numeric Manufacturer Entries as IDs</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductImportNumericVendorsAreUID" id="ProductImportNumericVendorsAreUID" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductImport", "ProductImportNumericVendorsAreUID"), False)) %>>&nbsp;<label for="ProductImportNumericVendorsAreUID" title="The import tool automatically treats any numeric entry as a uid and will assign the product to the corresponding entry. If this is unchecked than any numeric entry in the data source will be treated as a name.">Treat Numeric Vendor Entries as IDs</label></td>
  </tr>
</table>




<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblProductManager">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="sfProductAdmin.asp"><font color="white">Product Manager Settings</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ProductManager/help_ProductManager.htm')" name="btnHelp"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("ProductManager") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>

  <tr>
	<td><label for="ProductManagerDefaultDetailLinkPath" title="Double-clicking the Link label will automatically set the link path to use this pre-defined path.">Double-Click Detail Link Path</label>:&nbsp;</td>
	<td><input name="ProductManagerDefaultDetailLinkPath" id="ProductManagerDefaultDetailLinkPath" value="<%= Server.HTMLEncode(.GetValue("ProductManager", "ProductManagerDefaultDetailLinkPath")) %>" size="20"></td>
  </tr>
  <tr>
	<td><label for="ProductManagerDefaultMaxRecords" title="Sets the initial number of records to show in summary section is. This value is only used for the initial page load.">Default Summary Page Size</label>:&nbsp;</td>
	<td><input name="ProductManagerDefaultMaxRecords" id="ProductManagerDefaultMaxRecords" value="<%= Server.HTMLEncode(.GetValue("ProductManager", "ProductManagerDefaultMaxRecords")) %>" size="5"></td>
  </tr>

  <tr>
	<td><label for="ProductManagerSummaryTableHeight" title="Controls the size, in pixels, of how large the summary section is. Set to 0 to disable internal scrolling and use browser vertical scroll bar.">Summary Height</label>:&nbsp;</td>
	<td><input name="ProductManagerSummaryTableHeight" id="ProductManagerSummaryTableHeight" value="<%= Server.HTMLEncode(.GetValue("ProductManager", "ProductManagerSummaryTableHeight")) %>" size="5"></td>
  </tr>
  <tr>
	<td><label for="ProductManagerShortDescriptionLength" title="By default this field is limited to 255 characters in the StoreFront application. If you desire, you can set it higher; this will required a change to the Products table.">Short Description Length</label>:&nbsp;</td>
	<td><input name="ProductManagerShortDescriptionLength" id="ProductManagerShortDescriptionLength" value="<%= Server.HTMLEncode(.GetValue("ProductManager", "ProductManagerShortDescriptionLength")) %>" size="5"></td>
  </tr>

  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductManagerAutoShowTable" id="ProductManagerAutoShowTable" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductManager", "ProductManagerAutoShowTable"), False)) %>>&nbsp;<label for="ProductManagerAutoShowTable" title="Automatically loads the product summary table on the initial page load. This is useful for very large databases. If unchecked then you must apply a filter before the summary will load.">Automatically Load Summary Table</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductManagerShowTabs" id="ProductManagerShowTabs" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductManager", "ProductManagerShowTabs"), False)) %>>&nbsp;<label for="ProductManagerShowTabs" title="Used to show the product detail information in a series of tabs. If unchecked all information will be shown at one time.">Use Detail Tabs</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductManagerAutoShowDetailInWindow" id="ProductManagerAutoShowDetailInWindow" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductManager", "ProductManagerAutoShowDetailInWindow"), False)) %>>&nbsp;<label for="ProductManagerAutoShowDetailInWindow" title="Used to set the default value for whether you want to automatically display the product detail in a new window. This value is only used for the initial page load.">Automatically Show Detail In Existing Window</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductManagerDeleteExistingProductAttributesOnCopy" id="ProductManagerDeleteExistingProductAttributesOnCopy" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductManager", "ProductManagerDeleteExistingProductAttributesOnCopy"), False)) %>>&nbsp;<label for="ProductManagerDeleteExistingProductAttributesOnCopy" title="Normally when copying attributes from one product to another the application will add to the existing product's attributes. This setting, when checked, deletes the existing product's attributes prior to copying them so that an exact set of attributes remains.">Delete Existing Product Attributes On Copy</label></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="ProductManagerShowCustomTab" id="ProductManagerShowCustomTab" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductManager", "ProductManagerShowCustomTab"), False)) %>>&nbsp;<label for="ProductManagerShowCustomTab" title="Used to override display of custom data entry tab. The custom tab will automatically display if custom fields are used.">Show Custom Tab</label></td>
  </tr>
</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblProductPricing">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="ssProductPricingTool.asp"><font color="white">Product Pricing Tool Settings</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ProductPricing/help_ProductPricing.htm')" name="btnHelp"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("ProductPricing") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>
  <tr>
	<td><label for="cstrProductPricingAttributeStyle" title="Style to use for attributes. Contrasting styles should be used for visibility. An empty value will default to the even/odd row style">Attribute Style</label>:&nbsp;</td>
	<td><input name="cstrProductPricingAttributeStyle" id="cstrProductPricingAttributeStyle" value="<%= Server.HTMLEncode(.GetValue("ProductPricing", "cstrProductPricingAttributeStyle")) %>" size="20">&nbsp;(ex. style="someStyle" or bgcolor="lightgrey")</td>
  </tr>
  <tr>
	<td><label for="cstrProductPricingEvenStyle" title="Style to use for even rows of products. Contrasting styles should be used for visibility.">Even Row Style</label>:&nbsp;</td>
	<td><input name="cstrProductPricingEvenStyle" id="cstrProductPricingEvenStyle" value="<%= Server.HTMLEncode(.GetValue("ProductPricing", "cstrProductPricingEvenStyle")) %>" size="20">&nbsp;(ex. style="someStyle" or bgcolor="lightgrey")</td>
  </tr>
  <tr>
	<td><label for="cstrProductPricingOddStyle" title="Style to use for odd rows of products. Contrasting styles should be used for visibility.">Odd Row Style</label>:&nbsp;</td>
	<td><input name="cstrProductPricingOddStyle" id="cstrProductPricingOddStyle" value="<%= Server.HTMLEncode(.GetValue("ProductPricing", "cstrProductPricingOddStyle")) %>" size="20">&nbsp;(ex. style="someStyle" or bgcolor="white")</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="cblnProductPricingUseAttributes" id="cblnProductPricingUseAttributes" value="1" <%= isChecked(ConvertToBoolean(.GetValue("ProductPricing", "cblnProductPricingUseAttributes"), False)) %>>&nbsp;<label for="cblnProductPricingUseAttributes" title="Check this option to include attributes in the output. This can significantly increase load time.">Display Attribute Options</label></td>
  </tr>
</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblSalesCentral">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="default.asp"><font color="white">Sales Central</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/SalesCentral/help_SalesCentral.htm')" name="btnHelp" ID="btnHelpSalesCentral"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("SalesCentral") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>
</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblProductPlacement">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="ssProductPlacement.asp"><font color="white">Product Placement Settings</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ProductPlacement/help_ProductPlacement.htm')" name="btnHelp" ID="btnHelpProductPlacement"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("ProductPlacement") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>
</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblPromotionalMail">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="ssPromotionMailAdmin.asp"><font color="white">Promotional Mail Settings</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/PromotionalMail/help_PromotionalMail.htm')" name="btnHelp" ID="btnHelpPromotionalMail"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("PromotionalMail") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>
</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="1" rules="none" id="tblSalesReport">
  <colgroup align="right" valign="top">
  <colgroup align="left" valign="top">
  <tr class="tblhdr">
	<th colspan=2 align="left"><a href="ssSalesReport.asp"><font color="white">Sales Report Settings</font>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/SalesReports/help_SalesReports.htm')" name="btnHelp"></th>
  </tr>
  <tr>
	<td>Registration Status:&nbsp;</td>
	<td><% Call ShowVersionCheckStatus("SalesReport") %></td>
  </tr>
  <tr><td>&nbsp;</td><td><hr /></td></tr>

  <tr>
	<td><label for="SalesReportNumTopSellersToShow" title="The number of top sellers for each period to display. A value of 0 will display all. Note: If you have a large catalog the page load time may be substantial.">Number of top sellers to show</label>:&nbsp;</td>
	<td><input name="SalesReportNumTopSellersToShow" id="SalesReportNumTopSellersToShow" value="<%= Server.HTMLEncode(.GetValue("SalesReport", "SalesReportNumTopSellersToShow")) %>" size="20"></td>
  </tr>
  <tr>
	<td><label for="SalesReportOpenTag" title="The top sellers list is output as a list with each item wrapped in a <li></li> tag. You can customize the output by setting the opening and closing tag.">List tag - opening tag</label>:&nbsp;</td>
	<td><input name="SalesReportOpenTag" id="SalesReportOpenTag" value="<%= Server.HTMLEncode(.GetValue("SalesReport", "SalesReportOpenTag")) %>" size="20">&nbsp;(ex. &lt;ol style="MARGIN-LEFT: 18pt; MARGIN-RIGHT: 0pt; MARGIN-BOTTOM: 0;"&gt;)</td>
  </tr>
  <tr>
	<td><label for="SalesReportCloseTag" title="The top sellers list is output as a list with each item wrapped in a <li></li> tag. You can customize the output by setting the opening and closing tag.">List tag - closing tag</label>:&nbsp;</td>
	<td><input name="SalesReportCloseTag" id="SalesReportCloseTag" value="<%= Server.HTMLEncode(.GetValue("SalesReport", "SalesReportCloseTag")) %>" size="20">&nbsp;(ex. &lt;/ol&gt;)</td>
  </tr>

  <tr>
	<td>&nbsp;</td>
	<td><input type="checkbox" name="SalesReportAutoSelectAllOrders" id="SalesReportAutoSelectAllOrders" value="1" <%= isChecked(ConvertToBoolean(.GetValue("SalesReport", "SalesReportAutoSelectAllOrders"), False)) %>>&nbsp;<label for="SalesReportAutoSelectAllOrders" title="Check this to automatically select and include all orders in the sales report. If unchecked, only the summary will appear and the sales report will only include the first order.">Automatically select all orders</label></td>
  </tr>
  <tr>
	<td><label for="SalesReportDefaultSalesReport" title="The file name of the template in the Sales Report template directory to use.">Default Sales Report to view</label>:&nbsp;</td>
	<td><input name="SalesReportDefaultSalesReport" id="SalesReportDefaultSalesReport" value="<%= Server.HTMLEncode(.GetValue("SalesReport", "SalesReportDefaultSalesReport")) %>" size="20">&nbsp;(ex. Order Summary.xsl)</td>
  </tr>
</table>

<hr>
<table class="tbl" cellpadding="3" cellspacing="0" border="0" width="100%" style="border-collapse: collapse" id="tblData">
  <colgroup align=right valign=top>
  <colgroup align=center valign=top>
  <TR>
    <TD>&nbsp;</TD>
    <TD align="left">&nbsp;&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset>&nbsp;&nbsp;
        <INPUT class='butn' name="btnUpdate" id="btnUpdate" type=submit value='Update'>
    </TD>
  </TR>
</TABLE>
</FORM>

<% If mblnSaveError Then %>
<span id=spantempFile style="display:none"><input type=file id=tempFile name=tempFile size="20"></span>
<textarea name="addonsConfigXML" id="addonsConfigXML" rows="57" cols="120"><%= .xml %></textarea>
<% End If %><%
    End With
    Set mclsSSAddonConfig = Nothing
%><!--#include file="adminFooter.asp"--></BODY></HTML><!--
pstraccountsReceivableName
pstrbankAccountName
pstrpaymentsAccountName
pstrsalesAccountName
pstrsalesTaxName
pstrTaxEntity_NoTax
pstrdiscountName
pstrdiscountMemo
pstrdiscountInvItem
pstrgiftCertificateName
pstrgiftCertificateMemo
pstrgiftCertificateInvItem
pstrshippingName
pstrshippingMemo
pstrshippingInvItem
pstrhandlingName
pstrhandlingMemo
pstrhandlingInvItem
pstrattributeSeparator_productCode
pstrattributeSeparator_multipleAttributes
pstrattributeSeparator
pstrattributeSeparator_productCode_INV
pstrattributeSeparator_multipleAttributes_INV
pstrattributeSeparator_INV
pstrstoreMessage
pstrdefaultCarrierCode
pstrprintInvoice
pstrREP_To_Use
pstrdisplayLastNameFirst
pstrdateToUse
-->