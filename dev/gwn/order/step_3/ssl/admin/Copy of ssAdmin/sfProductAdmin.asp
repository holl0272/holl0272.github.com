 <%Option Explicit
'********************************************************************************
'*   Product Manager Version SF 5.0 					                        *
'*   Release Version:	2.00.003		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		April 18, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   2.00.003 (April 18, 2004)													*
'*   - Miscellaneous enhancements									            *
'*                                                                              *
'*   2.00.002 (September 11, 2003)                                              *
'*   - Added support for sites which must use https - ignore warning            *
'*   - Added support for Attribute Extender custom types				        *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True
Server.ScriptTimeout = 600

Call DetermineAddOns
Call LoadUserSettings

mstrssAddonVersion = "2.00.003"

'Call writeTimer("Start Page")
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="sfProductAdmin_custom.asp"-->
<!--#include file="Common/ssProduct_CommonFilter.asp"-->
<!--#include file="Common/ssProducts_Common.asp"-->
<!--#include file="ProductManager_Support/sfProductAdmin_class.asp"-->
<!--#include file="ProductManager_Support/sfProductAdmin_detail.asp"-->
<%

Function SetImagePath(strImage)

	If len(trim(strImage & "")) > 0 Then
		SetImagePath = mstrBaseHRef & strImage
	Else
		SetImagePath = "images/NoImage.gif"
	End If

End Function	'SetImagePath

'******************************************************************************************************************************************************************

Function ShortsImageManager()

dim pobjFSO
dim pstrTestPath

	pstrTestPath = Replace(Server.MapPath("sfProductAdmin.asp"),"sfProductAdmin.asp","ImageManager\imageupload.asp")
	set pobjFSO = server.CreateObject("Scripting.FileSystemObject")
	ShortsImageManager = pobjFSO.FileExists(pstrTestPath)
	set pobjFSO = nothing

End Function	'ShortsImageManager

'******************************************************************************************************************************************************************

Function DeleteSelected(byRef objclsProduct)

Dim pstrchkProduct
Dim parychkProduct
Dim i

	pstrchkProduct = Request.Form("chkProductID")
	parychkProduct = Split(pstrchkProduct, ",")
	For i = 0 To UBound(parychkProduct)
		objclsProduct.DeleteProduct Trim(parychkProduct(i))
	Next 'i
	
End Function	'DeleteSelected

'******************************************************************************************************************************************************************

Function ActivateSelected(byRef objclsProduct, byVal blnActivate)

Dim pstrchkProduct
Dim parychkProduct
Dim i

	pstrchkProduct = Request.Form("chkProductID")
	parychkProduct = Split(pstrchkProduct, ",")
	For i = 0 To UBound(parychkProduct)
		If len(Trim(parychkProduct(i))) > 0 Then objclsProduct.Activate Trim(parychkProduct(i)), blnActivate
	Next 'i
	
End Function	'DeleteSelected

'******************************************************************************************************************************************************************

Function getSelectedUIDs(byVal blnIncludeCodes, byVal blnIncludeSummary)

Dim i
Dim parychkProduct
Dim paryProductCodes
Dim plngTempID
Dim plngNumBaseProducts
Dim plngNumSummaryProducts
Dim pstrchkProduct
Dim pstrProductCodes

	If blnIncludeCodes Then pstrProductCodes = Request.Form("CopyProduct")
	If Len(pstrProductCodes) > 0 Then
		If InStr(1, pstrProductCodes, ";") > 0 Then
			paryProductCodes = Split(pstrProductCodes, ";")
		ElseIf InStr(1, pstrProductCodes, vbTab) > 0 Then
			paryProductCodes = Split(pstrProductCodes, vbTab)
		Else
			paryProductCodes = Split(pstrProductCodes, ",")
		End If
		
		For i = 0 To UBound(paryProductCodes)
			plngTempID = getProductUIDByCode(paryProductCodes(i))
			If plngTempID = -1 Then
				Call addMessageItem("No product with a code of <em>" & paryProductCodes(i) & "</em> exists", False)
				paryProductCodes(i) = ""
			Else
				paryProductCodes(i) = plngTempID
			End If
		Next 'i
	End If	'Len(pstrProductCodes) > 0
	
	If blnIncludeSummary Then
		pstrchkProduct = Request.Form("chkProductID")
		parychkProduct = Split(pstrchkProduct, ",")
		
		If isArray(paryProductCodes) Then
			plngNumBaseProducts = UBound(paryProductCodes)
			plngNumSummaryProducts = UBound(parychkProduct)
			ReDim Preserve paryProductCodes(plngNumBaseProducts + plngNumSummaryProducts + 1)
			For i = 0 To UBound(parychkProduct)
				paryProductCodes(plngNumBaseProducts + i + 1) = parychkProduct(i)
			Next 'i
		Else
			paryProductCodes = parychkProduct
		End If
		
	End If
	
	For i = 0 To UBound(paryProductCodes)
		paryProductCodes(i) = Trim(paryProductCodes(i))
		'Call getProductInfo(paryProductCodes(i), pstrProductCodes)
		'debugprint i, paryProductCodes(i) & "(" & pstrProductCodes & ")"
	Next 'i
	
	getSelectedUIDs = paryProductCodes
	
End Function	'getSelectedUIDs

'******************************************************************************************************************************************************************

Sub UpdateSelected()

Dim i
Dim paryTargetProducts
Dim pobjCmd
Dim pstrProductName
Dim pstrTargetField
Dim pstrTargetValue
Dim pstrTargetValue_Temp
Dim pstrSQL

	paryTargetProducts = getSelectedUIDs(False, True)
	'Check to see if any products selected
	If UBound(paryTargetProducts) = -1 Then
		Call addMessageItem("<font color=red>No products where selected to update.</font>", False)
		Exit Sub
	End If

	pstrTargetField = Request.Form("updateSelectedField")
	pstrTargetValue = Trim(Request.Form("updateSelectedValue"))

	Select Case pstrTargetField
		Case "prodCategoryID":
				If CBool(Len(pstrTargetValue) = 0) Then
					Call addMessageItem("<font color=red>No value given for the category.</font>", False)
				ElseIf Not isNumeric(pstrTargetValue) Then
					pstrTargetValue_Temp = pstrTargetValue
					pstrTargetValue = getNameFromID("sfCategories", "catName", "catID", True, pstrTargetValue_Temp)

					If Len(pstrTargetValue) = 0 Then
						Call addMessageItem("Category <em>" & pstrTargetValue_Temp & "</em> does not exist.", False)
						Call resetCombo_Saved("manufacturer")
					End If
				End If
				If Len(pstrTargetValue) > 0 Then pstrSQL = "Update sfProducts Set prodCategoryId=" & pstrTargetValue & " Where prodID=?"
		Case "prodManufacturerId":
				If CBool(Len(pstrTargetValue) = 0) Then
					Call addMessageItem("<font color=red>No value given for the manufacturer.</font>", False)
				ElseIf Not isNumeric(pstrTargetValue) Then
					pstrTargetValue_Temp = pstrTargetValue
					pstrTargetValue = getManufacturerByName(pstrTargetValue, True)
					If pstrTargetValue = -1 Then
						pstrTargetValue = ""	'Reset to nothing if an error resulted
					Else
						Call addMessageItem("Manufacturer <em>" & pstrTargetValue_Temp & "</em> created.", False)
						Call resetCombo_Saved("manufacturer")
					End If
				End If
				If Len(pstrTargetValue) > 0 Then pstrSQL = "Update sfProducts Set prodManufacturerId=" & pstrTargetValue & " Where prodID=?"
		Case "prodVendorId":
				If CBool(Len(pstrTargetValue) = 0) Then
					Call addMessageItem("<font color=red>No value given for the vendor.</font>", False)
				ElseIf Not isNumeric(pstrTargetValue) Then
					pstrTargetValue_Temp = pstrTargetValue
					pstrTargetValue = getVendorByName(pstrTargetValue, True)
					If pstrTargetValue = -1 Then
						pstrTargetValue = ""	'Reset to nothing if an error resulted
					Else
						Call addMessageItem("Vendor <em>" & pstrTargetValue_Temp & "</em> created.", False)
						Call resetCombo_Saved("vendor")
					End If
				End If
				If Len(pstrTargetValue) > 0 Then pstrSQL = "Update sfProducts Set prodVendorId=" & pstrTargetValue & " Where prodID=?"
		Case "prodEnabledIsActive", _
			 "prodSaleIsActive", _
			 "prodShipIsActive", _
			 "prodCountryTaxIsActive", _
			 "prodStateTaxIsActive":
				If CBool(Len(pstrTargetValue) = 0) Then
					Call addMessageItem("<font color=red>No value given to update the products to.</font>", False)
				ElseIf Not isNumeric(pstrTargetValue) Then
					If ConvertToBoolean(pstrTargetValue, False) Then
						pstrTargetValue = 1
					Else
						pstrTargetValue = 0
					End If
				End If
				If Len(pstrTargetValue) > 0 Then pstrSQL = "Update sfProducts Set " & makeSQLUpdate(pstrTargetField, pstrTargetValue, True, enDatatype_number) & " Where prodID=?"
		Case "prodWeight", _
			 "prodLength", _
			 "prodWidth", _
			 "prodHeight":
				If CBool(Len(pstrTargetValue) = 0) Or Not isNumeric(pstrTargetValue) Then
					Call addMessageItem("<font color=red>Invalid data for field <em>" & pstrTargetField & "</em>. {" & pstrTargetValue & "} is not a numeric value</font>", False)
					pstrTargetValue = ""
				End If
				If Len(pstrTargetValue) > 0 Then pstrSQL = "Update sfProducts Set " & makeSQLUpdate(pstrTargetField, pstrTargetValue, True, enDatatype_number) & " Where prodID=?"
		Case Else:
				pstrSQL = "Update sfProducts Set " & makeSQLUpdate(pstrTargetField, pstrTargetValue, True, enDatatype_string) & " Where prodID=?"
	End Select	'pstrTargetField

	If Len(pstrSQL) > 0 Then
		Call addMessageItem("<strong>Updating [" & pstrTargetField & "] to a value of [" & pstrTargetValue & "]</strong><br />", False)
		Set pobjCmd = Server.CreateObject("ADODB.COMMAND")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 50, paryTargetProducts(0))

			For i = 0 To UBound(paryTargetProducts)
				Call getProductInfo(paryTargetProducts(i), pstrProductName)
				
				.Parameters("prodID").Value = paryTargetProducts(i)
				.Execute adExecuteNoRecords
				
				If Err.number = 0 Then
					Call addMessageItem("&nbsp;&nbsp;- <em>" & paryTargetProducts(i) & ": " & pstrProductName & "</em> updated.", False)
				Else
					Call addMessageItem("&nbsp;&nbsp;- Error updating <em>" & paryTargetProducts(i) & ": " & pstrProductName & "</em>. Error " & err.number & ": " & err.Description & ".", True)
					Err.Clear
				End If
			Next 'i
		End With
		Set pobjCmd = Nothing
	End If	'Len(pstrSQL) > 0
	
End Sub	'UpdateSelected

'******************************************************************************************************************************************************************
'
'	Begin Main Page
'
'******************************************************************************************************************************************************************

mstrPageTitle = "Product Administration"

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsProduct
Dim mstrprodID, mlngattrID, mlngattrdtID
Dim mblnShortsImageManager
Dim i

Dim mblnShowFilter, mblnShowSummary, mstrShow
Dim cblnAddon_DynamicProductDisplay
Dim maryDisplayField

Dim mblnShowHeader
Dim mblnShowDetail
Dim mlngPageCount
Dim mlngAbsolutePage
Dim mlngMaxRecords
Dim mlngShortDescriptionLength
Dim mblnAttrPrice
Dim mblnMTPrice
Dim mblnShowTabs
Dim mblnAutoShowTable
Dim mblnExpandAttributesAutomatically
Dim cstrDefaultSmallImagePath
Dim cstrDefaultLargeImagePath
Dim cstrDefaultDetailLinkPath

	mlngMaxRecords = LoadRequestValue("PageSize")
	If len(mlngMaxRecords) = 0 Then mlngMaxRecords = clngDefaultMaxRecords

	mblnShowHeader = True
	mblnShowDetail = True

	mstrShow = Request.Form("Show")
	If Len(mstrShow) = 0 Then mstrShow = "General"
	
	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	
	If mblnDetailInNewWindow Then
		mblnShowHeader = False
		mblnShowDetail = True
	End If
	
	mstrprodID = LoadRequestValue("prodID")
	mlngattrID = LoadRequestValue("attrID")
	mlngattrdtID = LoadRequestValue("attrdtID")
	mAction = LoadRequestValue("Action")
	If Len(mAction) = 0 Then mAction = "Filter"
	mlngAbsolutePage = LoadRequestValue("AbsolutePage")
	mblnShortsImageManager = ShortsImageManager
	
    Call LoadFilter
    Call LoadSort
    
'debugprint "mstrprodID",mstrprodID
'debugprint "pstrSQL",pstrSQL
'debugprint "mAction",mAction
    Set mclsProduct = New clsProduct
    With mclsProduct
    Select Case mAction
        Case "Activate", "Deactivate"
			If len(mstrprodID) > 0 Then	.Activate mstrprodID, mAction = "Activate"
        Case "ActivateMarked", "DeactivateMarked"
			Call ActivateSelected(mclsProduct, CBool(mAction = "ActivateMarked"))
		Case "CopyProd"
			.DuplicateProduct mstrprodID, Request.Form("CopyProduct")
			If Request.Form("CopyProductCategories") = "1" Then .CopyProductCategories mstrprodID, Request.Form("CopyProduct")
			mstrprodID = Request.Form("CopyProduct")
			mlngattrID = ""
			mlngattrdtID = ""
		Case "CopyCategories"
			.CopyProductCategories mstrprodID, Request.Form("CopyProduct")
			mstrprodID = Request.Form("CopyProduct")
			mlngattrID = ""
			mlngattrdtID = ""
		Case "CopyAttributesToProd" 'copies all product attributes to existing product
			.CopyProductAttributesToExistingProduct mstrprodID, Request.Form("CopyProduct"), ""
			mstrprodID = Request.Form("CopyProduct")
			mlngattrID = ""
			mlngattrdtID = ""
			.prodID = mstrProdID
			.UpdateInventoryFields
		Case "CopyAttr" 'copies single selected attribute to existing product
			.CopyProductAttributesToExistingProduct mstrprodID, Request.Form("CopyProduct"), mlngattrID
			mstrprodID = Request.Form("CopyProduct")
			mlngattrID = ""
			mlngattrdtID = ""
			.prodID = mstrProdID
			.UpdateInventoryFields
			.UpdateInventoryFields
        Case "DeleteProduct"
			.DeleteProduct mstrprodID
			mstrprodID = ""
			mblnShowSummary = True
        Case "DeleteMarked"
			Call DeleteSelected(mclsProduct)
			mstrprodID = ""
        Case "DeleteAttribute"
			.DeleteAttribute mlngattrID
			.prodID = mstrProdID
			.UpdateInventoryFields
        Case "DeleteAllAttributes"
			.DeleteAllAttributes mstrProdID
			.prodID = mstrProdID
			.UpdateInventoryFields
        Case "DeleteAttrDetail"
			.DeleteAttributeDetail mlngattrdtID
			.prodID = mstrProdID
			.UpdateInventoryFields
		Case "DuplicateAttr" 'creates a duplicate of the selected product attribute within the current product
			.CopyAttribute mlngattrID,"",Request.Form("CopyProduct")
			.prodID = mstrProdID
		Case "Filter"
			mstrprodID = ""
			If mblnDetailInNewWindow Then 
				mblnShowHeader = True
				mblnShowDetail = False
			End If
			mblnShowSummary = True
        Case "New", "Update"
            .Update
		Case "UpdateSelected"
			Call UpdateSelected
		Case "ViewProduct"
			mstrprodID = LoadRequestValue("ViewID")
			mlngattrID = ""
			mlngattrdtID = ""
		Case "ViewAttribute"
			mstrprodID = ""
			mlngattrID = Request.Form("ViewID")
			mlngattrdtID = ""
		Case "ViewAttrDetail"
			mstrprodID = ""
			mlngattrID = ""
			mlngattrdtID = Request.Form("ViewID")
		Case Else
			mblnShowFilter = True
    End Select
    
	Call .LoadSummary(mstrprodID)

	If len(mstrprodID) > 0 Then
		Call .LoadIndividualProduct(mstrprodID)
	Elseif len(mlngattrID) > 0 Then
		.FindAttribute mlngattrID
	Elseif len(mlngattrdtID) > 0 Then
		.FindAttrDetail mlngattrdtID
	ElseIf Not .rsProducts.EOF Then
		mstrprodID = .rsProducts.Fields("prodID").Value
		Call .LoadIndividualProduct(mstrprodID)
	End If
	If cblnSF5AE Then .LoadAE
	
	dim mrsCategory, mrsVendor, mrsManufacturer

	Set mrsCategory = GetRS("Select catID,catName from sfCategories Order By catName")
	Set mrsVendor = GetRS("Select vendID,vendName from sfVendors Order By vendName")
	Set mrsManufacturer = GetRS("Select mfgID,mfgName from sfManufacturers Order By mfgName")

	If .PricingLevel Then
		Dim mstrHeaderRow

		Dim maryPLPrices
		Dim mstrPLPrice
		Dim mobjrsPricingLevels
		Dim mlngNumPricingLevels
		Set mobjrsPricingLevels = GetRS("Select PricingLevelID,PricingLevelName from PricingLevels Order By PricingLevelID")
		If mobjrsPricingLevels.EOF Then
			mlngNumPricingLevels = 0
		Else
			mlngNumPricingLevels = mobjrsPricingLevels.RecordCount
			ReDim maryPricingLevels(mlngNumPricingLevels-1)
		End If

		mstrHeaderRow = "<tr>"
		For i = 1 To mlngNumPricingLevels
			maryPricingLevels(i-1) = Trim(mobjrsPricingLevels.Fields("PricingLevelName").Value)
			mstrHeaderRow = mstrHeaderRow & "<td><font size=-1><i><b>" & maryPricingLevels(i-1) & "</b></i></font></td>"
			mobjrsPricingLevels.MoveNext
		Next 'i
		mstrHeaderRow = mstrHeaderRow & "</tr>"
		If mlngNumPricingLevels > 0 Then mobjrsPricingLevels.MoveFirst
	End If
End With


If mblnShowHeader Then
	Call WriteHeader("body_onload();",True)
	Call WriteSupportingScripts
	Call WriteFormOpener
%>
<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th id="tdFilter" class="hdrSelected" nowrap onclick="return DisplayMainSection('Filter');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Set your order filter criteria here.">&nbsp;Product&nbsp;Filter&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdSummary" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('Summary');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Orders which meet the filter criteria">&nbsp;Product&nbsp;Summaries&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tditemDetail" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('itemDetail');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View results">&nbsp;Product&nbsp;Detail&nbsp;</th>
	<th width="90%" align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/OrderManager/help_OrderManager.htm')" id="btnHelp" name="btnHelp" title="Release Version <%= mstrssAddonVersion %>"></th>
  </tr>
  <tr>
	<td colspan="6" class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td colspan="6" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
<%
		Call WriteProductFilter(True)
		Response.Write outputMessage
		Call mclsProduct.OutputMessage
		If (len(Request.Form) > 0 or mblnAutoShowTable) Then
			Response.Write mclsProduct.OutputSummary
		Else
			Response.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke' id='tblSummary'>" _
						 & "<tr><th><h2>Product Summary Disabled. Please select a filter criteria</h2></th></tr>" _
						 & "</table>"
		End If
		Call WriteProductDetail(mstrProdID)
%>
	</td>
  </tr>
</table>
<%
	Else
		Call WriteHeader("body_onload();",False)
		Call WriteSupportingScripts
		Call WriteFormOpener
		Call mclsProduct.OutputMessage
		Call WriteProductDetail(mstrProdID)
	End If
	
	Call ReleaseObject(mrsCategory)
	Call ReleaseObject(mrsVendor)
	Call ReleaseObject(mrsManufacturer)
	If mclsProduct.PricingLevel Then Call ReleaseObject(mobjrsPricingLevels)
	
	Call ReleaseObject(cnn)
	'Call writeTimer("Page Complete")

    Response.Flush
%>
</FORM>
</center>
<!--#include file="adminFooter.asp"-->
</BODY>
</HTML>
<%

'************************************************************************************************************************************
'
'	SUPPORTING ROUTINES
'
'************************************************************************************************************************************

Sub WriteSupportingScripts

With mclsProduct
%>
<% On Error Goto 0 %>
<% If mblnShortsImageManager Then %><script language="javascript" src="ImageManager/SOSLibrary/incPureUpload.js"></script><% End If %>
<script LANGUAGE=javascript>
<!--
var theDataForm;
var theKeyField;
var strDetailTitle = "<% If len(.prodID) > 0 Then Response.Write .prodID & ": " & EncodeString(.prodName,False) %>";
var blnIsDirty;
var maryimgSizes;
var mlngImageWidth;
var mlngImageHeight;
var disableAttributeFocus = true;
var mlngPricingLevels = 0;

<% If .PricingLevel Then Response.Write "mlngPricingLevels = " & mlngNumPricingLevels & ";" %>

//tipMessage[...]=[title,text]
tipMessage['prodShortDescription']=["Data Entry Help","This is the product description which appears on the search result page.<br />Limited to 255 characters."]
tipMessage['prodDescription']=["Data Entry Help","This is the product description which appears on the detail page. You may leave this field empty if it is the same as the short description.<br />Unlimited length."]
tipMessage['prodMessage']=["Data Entry Help","This is the message that appears in the 'Thank You' window.<br />Unlimited length."]

tipMessage['buyersClubPointValue']=["Data Entry Help","This sets the point value each product receives during checkout. If you want it to be x times the purchase price check the Is Percentage checkbox."]
tipMessage['prodHandlingFee']=["Data Entry Help","This is a product specific handling fee. It is applied for <b>each</b> quantity of the item in the purchase."]
tipMessage['prodSetupFee']=["Data Entry Help","This is a product specific handling fee. It is applied <b>independent</b> of the quantity of the unique product ordered. Note: Products with different attributes are <b>NOT</b> considered unique and each instance will incur a separate setup charge."]
tipMessage['prodSetupFeeOneTime']=["Data Entry Help","This is a product specific handling fee. It is applied <b>independent</b> of the quantity of the unique product ordered. Note: Only products with different product codes are considered unique and will incur a separate setup charge."]

tipMessage['prodMinQty']=["Data Entry Help","This is the minimum qty which can be ordered. Empty, or 0 means no minimum qty."]
tipMessage['prodIncrement']=["Data Entry Help","This is the fraction of a product which can be ordered. Empty, 0, or 1 means only whole quantites can be ordered. Positive integers permit fractional ordering. Ex. 4 equates to fraction ordering by 1/4."]

tipMessage['prodSpecialShippingMethods']=["Data Entry Help","This is the list of special shipping methods which can be used for this product. IDs are from Shipping Methods Administration. Multiple IDs can be separated with commas.<br />"]
tipMessage['prodFixedShippingCharge']=["Data Entry Help","This is the shipping charge to be used if carrier based shipping is used."]
tipMessage['prodShip']=["Data Entry Help","This is the shipping charge to be used for product based shipping."]
tipMessage['prodFileName']=["Data Entry Help","This is the file path to the download.<br />Unlimited length."]
tipMessage['prodMaxDownloads']=["Data Entry Help","This is the maximum number of downloads before the download is locked out. Empty or 0 means unlimited.<br />Numeric."]
tipMessage['prodDownloadValidFor']=["Data Entry Help","This is the maximum number of days after the order before the download is locked out. Empty or 0 means unlimited.<br />Numeric."]

tipMessage['pageName']=["Data Entry Help",""]
tipMessage['metaTitle']=["Data Entry Help",""]
tipMessage['metaDescription']=["Data Entry Help",""]
tipMessage['metaKeywords']=["Data Entry Help",""]

function showFullImage(theImage, bytSwitch)
{

var pstrName = theImage.name;
	if (bytSwitch == 0)
	{
		mlngImageHeight = theImage.height;
		mlngImageWidth = theImage.width;
		
		theImage.height = maryimgSizes[pstrName][0];
		theImage.width = maryimgSizes[pstrName][1];
	}else{
		theImage.height = mlngImageHeight;
		theImage.width = mlngImageWidth;
	}
}

function MakeDirty(theItem)
{
var theForm = theItem.form;

	theForm.btnReset.disabled = false;
	blnIsDirty = true;
}
<% Response.Flush %>
function body_onload()
{
	theDataForm = document.frmData;
	theKeyField = theDataForm.prodID;
	blnIsDirty = false;
	
	<% If mblnShowDetail Then %>
	maryimgSizes = new Array();
	<% For i = 0 To UBound(maryImageFields) %>
	maryimgSizes["img<%= maryImageFields(i)(1) %>"] = new Array(document.all("img<%= maryImageFields(i)(1) %>").height, document.all("img<%= maryImageFields(i)(1) %>").width);
	document.all("img<%= maryImageFields(i)(1) %>").height = 50;
	<% Next 'i %>
	<% End If %>
	
<%
If mblnShowSummary Then
	Response.Write "DisplayMainSection('Summary');" & vbcrlf
ElseIf mblnShowFilter Then
	Response.Write "DisplayMainSection('Filter');" & vbcrlf
Else
	If mblnShowHeader Then Response.Write "DisplayMainSection('itemDetail');" & vbcrlf
	Response.Write "ScrollToElem('selectedSummaryItem');" & vbcrlf
	
	If len(.attrdtID) > 0 Then 
		Response.Write "SelectAttrDetail(" & .attrID & "," & .attrdtID & ");"
	ElseIf len(.attrID) > 0 Then 
		Response.Write "SelectAttr(" & .attrID & ");"
	End If
	Response.Write "DisplaySection(" & chr(34) & mstrShow & chr(34) & ");"
End If
%>
<% 
'If Not mblnShowDetail then Response.Write "return false;" & vbcrlf

If cblnAddon_DynamicProductDisplay Then
%>
	FillItem("relatedProducts");
<% End If %>
	document.all("spanprodName").innerHTML = strDetailTitle;

<% If cblnSF5AE Then %>
var arySections = new Array("General","Detail","Attributes","Shipping","MTP","Inventory","Category");
InitializeCategory();
FillCategory();
<% End If %>
}

var gobjImage;
var gblnSwitch;

function SelectImage(theImage)
{
	gblnSwitch = true;
	gobjImage = theImage;
	document.frmData.tempFile.click();
	return false;
}

function ProcessPath(theFile)
{
var pstrFilePath = theFile.value;
var pstrBaseHRef = document.frmData.strBaseHRef.value;
var pstrBasePath = document.frmData.strBasePath.value;
var pstrHREF;
var pstrItem;
var xyz = "\\";

	if (gblnSwitch)
	{
	gobjImage.src = pstrFilePath;
	pstrItem = gobjImage.name.replace("img","");
	pstrHREF = pstrFilePath.replace(pstrBasePath,"");
	eval("document.frmData." + pstrItem).value = pstrHREF.replace(xyz,"/");
	MakeDirty(eval("document.frmData." + pstrItem));
	document.frmData.btnReset.disabled = false;
	gblnSwitch = false;
	theFile.value = "";
	}
}

function btnNewProduct_onclick(theButton)
{
var theForm = theButton.form;

	theForm.OrigprodID.value = "";
	DisplaySection("Attributes");
	theForm.attrID[0].selected = true;
	theForm.attrID.length = 1;
	ChangeAttr(theForm.attrID);
//	theForm.attrdtID[0].selected = true;
//	theForm.attrName.value = "";

    theForm.btnUpdate.value = "Add Product";
    theForm.btnDeleteProduct.disabled = true;
    theForm.btnCopyProduct.disabled = true;
	theForm.btnReset.disabled = false;
	
	SetDefaults(theForm);
	
	// AE Specific
	if (<%= LCase(CStr(cblnSF5AE)) %>)
	{
		var MTPTable = document.all("tblMTPInput");
		for (var i=MTPTable.rows.length-1; i>0; i--)
		{
		MTPTable.deleteRow(i);
		}

		var InventoryTable = document.all("tblInventoryLevels");
		for (i=InventoryTable.rows.length-1; i>0; i--)
		{
		InventoryTable.deleteRow(i);
		}

		var theSelect = theDataForm.Categories;
		var theKey;
		for (var i=theSelect.length-1; i >=0 ;i--)
		{
			theKey = theSelect.options[i].value;
			mdicCategory.Remove(theKey);
			theSelect.options.remove(i);
		}
	//	mdicCategory.Add (3,"")
		FillCategory();
		
		theForm.ChangeInventory.value = true;
		theForm.ChangeMTP.value = true;
		theForm.ChangeCategory.value = true;
		theForm.invenInStockDEF.value = 0;
		theForm.invenLowFlagDEF.value = 0;

		theForm.invenbTracked.checked = false;
		theForm.invenbNotify.checked = false;
		theForm.invenbBackOrder.checked = false;
		theForm.invenbStatus.checked = false;
		
		theForm.gwPrice.value = 0;
		theForm.gwActivate.checked = false;
	}
	
	DisplaySection("General");
    theForm.prodID.focus();
    document.all("spanprodName").innerHTML = theDataForm.btnUpdate.value;

}

function btnDeleteProduct_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete the product " + theForm.prodID.value + ": " + theForm.prodName.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "DeleteProduct";
    theForm.submit();
    return(true);
    }
    else
    {
    return(false);
    }
}

function btnReset_onclick(theButton)
{
var theForm = theButton.form;

	blnIsDirty = false;
    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    theForm.btnDeleteProduct.disabled = false;
    
		if (aryAttributeNames.length != 0)
		{
			theForm.attrID.size = aryAttributeNames.length/2 + 1;
			for (var i=0; i < aryAttributeNames.length/2;i++)
			{
				theForm.attrID.options[i+1] = new Option(aryAttributeNames[i*2+1], aryAttributeNames[i*2]);
			}
		}
<%
If len(.attrdtID) > 0 Then 
	Response.Write "SelectAttrDetail(" & .attrID & "," & .attrdtID & ");"
ElseIf len(.attrID) > 0 Then 
	Response.Write "SelectAttr(" & .attrID & ");"
End If
%>
}

function SetDefaults(theForm)
{
    theForm.prodID.value = "";
    theForm.prodName.value = "";
    theForm.prodNamePlural.value = "";
    theForm.prodShortDescription.value = "";
    theForm.prodDescription.value = "";
    theForm.prodPrice.value = "0";
    theForm.prodEnabledIsActive.checked = false;
    theForm.prodSaleIsActive.checked = false;
    theForm.prodSalePrice.value = "0";
    theForm.prodDateAdded.value = "";

	<% For i = 0 To UBound(maryImageFields) %>
    theForm.<%= maryImageFields(i)(1) %>.value = "";
	<% Next 'i %>
    theForm.prodFileName.value = "";
    theForm.prodLink.value = "";
    theForm.prodMessage.value = "";
	SetSelect(theForm.prodCategoryId,0);
	SetSelect(theForm.prodManufacturerId,0);
	SetSelect(theForm.prodVendorId,0);

    theForm.prodWeight.value = "0";
    theForm.prodHeight.value = "0";
    theForm.prodWidth.value = "0";
    theForm.prodLength.value = "0";
    theForm.prodShip.value = "0";
    theForm.prodCountryTaxIsActive.checked = false;
    theForm.prodShipIsActive.checked = false;
    theForm.prodStateTaxIsActive.checked = false;
    
<%  
Dim i
Dim paryCustomValues

paryCustomValues = mclsProduct.CustomValues
If isArray(paryCustomValues) Then 
	For i = 0 To UBound(paryCustomValues)
		Response.Write "theForm." & paryCustomValues(i)(1) & ".value = " & Chr(34) & Chr(34) & ";" & vbcrlf
	Next 'i
End If
%>
    
    
return(true);
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.Action.value = "Filter";
	theDataForm.OrderBy.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.submit();
	return false;
}

function ViewPage(theValue)
{
	theDataForm.AbsolutePage.value = theValue;
	theDataForm.Action.value = "Filter";
	theDataForm.submit();
	return false;
}

function ViewProduct(theValue)
{
	var pblnDetailInNewWindow = (getRadio(document.frmData.chkDetailInNewWindow) == 1);

	if (pblnDetailInNewWindow)
	{
		var strURL = "sfProductAdmin.asp?Action=ViewProduct&chkDetailInNewWindow=1&ViewID=" + theValue;
		var DetailWindow = window.open(strURL,"ProductDetail","height=600,width=800,copyhistory=0,scrollbars=1,resizable=1");
		DetailWindow.focus();
		return false;
	}else{
		theDataForm.ViewID.value = theValue;
		theDataForm.Action.value = "ViewProduct";
		theDataForm.submit();
		return false;
	}
}

function ViewAttribute(theValue)
{
	theDataForm.ViewID.value = theValue;
	theDataForm.Action.value = "ViewAttribute";
	theDataForm.submit();
	return false;
}
function ViewAttrDetail(theValue)
{
	theDataForm.ViewID.value = theValue;
	theDataForm.Action.value = "ViewAttrDetail";
	theDataForm.submit();
	return false;
}

function SelectAttr(lngID)
{
	var pblnDetailInNewWindow = (getRadio(document.frmData.chkDetailInNewWindow) == 1);

	if (!pblnDetailInNewWindow)
	{
		theDataForm.attrID.value = lngID;
		DisplaySection("Attributes");
		ChangeAttr(theDataForm.attrID);
	}

}

function SelectAttrDetail(lngAttrID,lngID)
{

	var pblnDetailInNewWindow = (getRadio(document.frmData.chkDetailInNewWindow) == 1);

	if (!pblnDetailInNewWindow)
	{
		DisplaySection("Attributes");
		SelectAttr(lngAttrID);
		theDataForm.attrdtID.value = lngID;
		ChangeAttrDetail(theDataForm.attrdtID);
	}
}

<% 
.OutputAttrValues 
.OutputAttrDetailValues
%>

function ChangeAttr(theSelect)
{
var theForm = theSelect.form;
var intIndex = theSelect.selectedIndex;
var plngSelectedValue = theSelect.value;

	if (intIndex == 0)
	{
		setElementValue(theForm.attrName, "");
		setElementValue(theForm.attrDisplay, "");
		setElementValue(theForm.attrSKU, "");
		setElementValue(theForm.attrDisplayStyle, "");
		setElementValue(theForm.attrImage, "");
		setElementValue(theForm.attrURL, "");
		setElementValue(theForm.attrExtra, "");

		setElementValue(theForm.attrDisplayStyle,0);

		theForm.btnDeleteAttr.disabled = true;
		theForm.btnDuplicateAttr.disabled = true;
		theForm.btnCopyAttr.disabled = true;
		theForm.attrdtID.length = 1;
		theForm.attrdtID.size = 3;
		theForm.attrdtID[0].selected = true;
		ChangeAttrDetail(theForm.attrdtID);
		theForm.btnReset.disabled = false;
		theForm.attrName.focus();
		document.all("divAttrOptions").innerHTML = "&nbsp;";
		document.all("spanAttrOptions").innerHTML = "&nbsp;";
		document.all("imgUp").disabled = true;
		document.all("imgDown").disabled = true;		
	}
	else
	{
		theForm.btnDeleteAttr.disabled = false;
		theForm.btnDuplicateAttr.disabled = false;
		theForm.btnCopyAttr.disabled = false;
		theForm.attrdtID.length = 1;
		if (aryAttributes instanceof Array)
		{
		if (aryAttributes[intIndex] instanceof Array)
		{
			theForm.attrdtID.size = aryAttributes[intIndex].length/2 + 1;
			for (var i=0; i < aryAttributes[intIndex].length/2;i++)
			{
				theForm.attrdtID.options[i+1] = new Option(aryAttributes[intIndex][i*2], aryAttributes[intIndex][i*2+1]);
			}
		}
		document.all("imgUp").disabled = false;
		document.all("imgDown").disabled = false;		

		//added for text based attributes
<% If .TextBasedAttribute Then Response.Write "		setElementValue(theForm.attrDisplayStyle, aryAttributeDisplay[plngSelectedValue][0]);" %>
		}
	theForm.attrdtID[0].selected = true;
	ChangeAttrDetail(theForm.attrdtID);
	theForm.attrName.value = theSelect.item(intIndex).text;
<% If Len(.attrDisplay_Field) > 0 Then Response.Write "		setElementValue(theForm.attrDisplay, aryAttributeDisplay[plngSelectedValue][1]);" & vbcrlf %>
<% If Len(.attrExtra_Field) > 0 Then Response.Write "		setElementValue(theForm.attrExtra, aryAttributeDisplay[plngSelectedValue][2]);" & vbcrlf %>
<% If Len(.attrImage_Field) > 0 Then Response.Write "		setElementValue(theForm.attrImage, aryAttributeDisplay[plngSelectedValue][3]);" & vbcrlf %>
<% If Len(.attrSKU_Field) > 0 Then Response.Write "		setElementValue(theForm.attrSKU, aryAttributeDisplay[plngSelectedValue][4]);" & vbcrlf %>
<% If Len(.attrURL_Field) > 0 Then Response.Write "		setElementValue(theForm.attrURL, aryAttributeDisplay[plngSelectedValue][5]);" & vbcrlf %>
	document.all("divAttrOptions").innerHTML = theSelect.item(intIndex).text;
	document.all("spanAttrOptions").innerHTML = theSelect.item(intIndex).text + " Attributes";
	}
}

function ChangeAttrDetail(theSelect)
{
var theForm = theSelect.form;
var intValue = theSelect.value;

	if (theSelect.selectedIndex == 0)
	{
		theForm.btnDeleteAttrDetail.disabled = true;
		
		setElementValue(theForm.attrdtID, "");
		setElementValue(theForm.attrdtName, "");
		setElementValue(theForm.attrdtDisplay, "");
		setElementValue(theForm.attrdtSKU, "");
		setElementValue(theForm.attrdtPrice, "");
		setElementValue(theForm.attrdtType, 0);
		setElementValue(theForm.attrdtWeight, "");
		setElementValue(theForm.attrdtImage, "");
		setElementValue(theForm.attrdtURL, "");
		setElementValue(theForm.attrdtFileName, "");
		setElementValue(theForm.attrdtExtra, "");
		setElementValue(theForm.attrdtExtra1, "");
		setElementValue(theForm.attrdtDefault, false);
		
		if (mlngPricingLevels > 0)
		{
			if (mlngPricingLevels == 1)
			{
				setElementValue(theForm.attrdtPLPrice, "");
			}else{
				for (var i=0; i < mlngPricingLevels;i++){setElementValue(theForm.attrdtPLPrice[i], "");}
			}
		}

		if (!disableAttributeFocus) {theForm.attrdtName.focus();}

	}
	else
	{
		theForm.btnDeleteAttrDetail.disabled = false;
		if (aryAttributeDetails[intValue] instanceof Array)
		{

			setElementValue(theForm.attrdtName, aryAttributeDetails[intValue][0]);
			setElementValue(theForm.attrdtPrice, aryAttributeDetails[intValue][1]);
			setElementValue(theForm.attrdtType, aryAttributeDetails[intValue][2]);
			setElementValue(theForm.attrdtWeight, aryAttributeDetails[intValue][3]);
			setElementValue(theForm.attrdtDisplay, aryAttributeDetails[intValue][5]);
			setElementValue(theForm.attrdtFileName, aryAttributeDetails[intValue][6]);
			setElementValue(theForm.attrdtSKU, aryAttributeDetails[intValue][7]);
			setElementValue(theForm.attrdtURL, aryAttributeDetails[intValue][8]);
			setElementValue(theForm.attrdtDefault, aryAttributeDetails[intValue][9]);
			setElementValue(theForm.attrdtExtra, aryAttributeDetails[intValue][10]);
			setElementValue(theForm.attrdtExtra1, aryAttributeDetails[intValue][11]);

			if (theForm.attrdtImage != undefined)
			{
				if (aryAttributeDetails[intValue][4] == "")
				{
					theForm.attrdtImage.value = aryAttributeDetails[intValue][4];
					document.all("imgattrdtImage").src = "<%= SetImagePath("") %>";
				}else{
					theForm.attrdtImage.value = aryAttributeDetails[intValue][4];
					document.all("imgattrdtImage").src = "<%= SetImagePath("") %>" + aryAttributeDetails[intValue][4];					
				}
			}
			
		}
		if (mlngPricingLevels > 0)
		{
			if (mlngPricingLevels == 1)
			{
				setElementValue(theForm.attrdtPLPrice, "");
			}else{
				for (var i=0; i < mlngPricingLevels;i++){setElementValue(theForm.attrdtPLPrice[i], "");}
			}
			
			// added for Pricing Levels
			var paryPrices;
			var pstrPrices = new String(aryAttributePLDetails[intValue]);
			paryPrices = pstrPrices.split(";");
				
			for (var i=0; i < mlngPricingLevels;i++)
			{
				if (i < paryPrices.length)
				{
					if (mlngPricingLevels == 1)
					{
						setElementValue(theForm.attrdtPLPrice, paryPrices[i]);
					}else{
						setElementValue(theForm.attrdtPLPrice[i], paryPrices[i]);
					}
				}else{
					if (mlngPricingLevels == 1)
					{
						setElementValue(theForm.attrdtPLPrice, "");
					}else{
						setElementValue(theForm.attrdtPLPrice[i], "");
					}
				}
			}
		}
	}
}

function ChangeAttr_MTP(theSelect)
{
var theForm = theSelect.form;
var intIndex = theSelect.selectedIndex + 1;

	theForm.attrdtID_MTP.length = 1;
	if (aryAttributes instanceof Array)
	{
	if (aryAttributes[intIndex] instanceof Array)
	{
		theForm.attrdtID_MTP.size = aryAttributes[intIndex].length/2;
		for (var i=0; i < aryAttributes[intIndex].length/2;i++)
		{
			theForm.attrdtID_MTP.options[i] = new Option(aryAttributes[intIndex][i*2], aryAttributes[intIndex][i*2+1]);
		}
	}

	theForm.attrdtID_MTP[0].selected = true;
	ChangeAttrDetail_MTP(theForm.attrdtID_MTP);
	}
}

function ChangeAttrDetail_MTP(theSelect)
{
var theForm = theSelect.form;
var intValue = theSelect.value;

	if (theSelect.selectedIndex == 0)
	{
//		theForm.attrdtName.value = "";
//		theForm.attrdtPrice.value = "";
//		setElementValue(theForm.attrdtType,0);
//		theForm.attrdtID.value = "";
//		theForm.attrdtName.focus();
	}
	else
	{
		if (aryAttributeDetails[intValue] instanceof Array)
		{
//			theForm.attrdtName.value = aryAttributeDetails[intValue][0];
//			theForm.attrdtPrice.value = aryAttributeDetails[intValue][1];
//			setElementValue(theForm.attrdtType,aryAttributeDetails[intValue][2]);
		}
	}
}

function UpItem(strTarget)
{

var theSelect;
if (strTarget == "attribute")
{
	theSelect = theDataForm.attrID;
}else{
	theSelect = theDataForm.attrdtID;
}
var intSelected = theSelect.selectedIndex;

	if (intSelected > 1)
	{
		var optText = theSelect.options[intSelected].text;
		var optValue = theSelect.options[intSelected].value;
		
		theSelect.options[intSelected].value = theSelect.options[intSelected-1].value;
		theSelect.options[intSelected].text = theSelect.options[intSelected-1].text;
		theSelect.options[intSelected-1].value = optValue;
		theSelect.options[intSelected-1].text = optText;
		theSelect.selectedIndex = intSelected - 1;
		MakeDirty(theSelect);
	}
}

function DownItem(strTarget)
{

var theSelect;
if (strTarget == "attribute")
{
	theSelect = theDataForm.attrID;
}else{
	theSelect = theDataForm.attrdtID;
}

var intSelected = theSelect.selectedIndex;

	if (intSelected > 0)
	{
	if (intSelected < (theSelect.length - 1))
	{
		var optText = theSelect.options[intSelected].text;
		var optValue = theSelect.options[intSelected].value;
		
		theSelect.options[intSelected].value = theSelect.options[intSelected+1].value;
		theSelect.options[intSelected].text = theSelect.options[intSelected+1].text;
		theSelect.options[intSelected+1].value = optValue;
		theSelect.options[intSelected+1].text = optText;
		theSelect.selectedIndex = intSelected + 1;
		MakeDirty(theSelect);
	}
	}
}

function GetSortOrder(theForm)
{
var strOrder = "";

	if (theForm.attrID.length > 1)
	{
	for (var i=1; i < theForm.attrID.length;i++)
	{
	strOrder += theForm.attrID.options[i].value + ","
	theForm.attrDisplayOrder.value = strOrder;
	}
	}

	if (theForm.attrdtID.length > 1)
	{
	for (var i=1; i < theForm.attrdtID.length;i++)
	{
	strOrder += theForm.attrdtID.options[i].value + ","
	theForm.attrdtOrder.value = strOrder;
	}
	}
}

function CheckAll(blnCheck)
{
var plngCount = document.frmData.chkProductID.length;
var i;

if (document.frmData.chkProductID.checked==undefined)
{
	for (i=0; i < plngCount;i++)
	{
	document.frmData.chkProductID[i].checked = blnCheck;
	}
}else{
document.frmData.chkProductID.checked = blnCheck;
}
}

function ValidInput(theForm)
{
var  strSection = frmData.Show.value;

//  if (!blnIsDirty)
//  {
//  alert("No changes");
//  return(false);
//  }

  DisplaySection("General");
  if (theDataForm.prodID.value == "")
  {
    alert("Please enter a Product ID.")
    theDataForm.prodID.focus();
    return(false);
  }
    
  if (theDataForm.prodName.value == "")
  {
    alert("Please enter a Product Name.")
    theDataForm.prodName.focus();
    return(false);
  }
    
  if (!isNumeric(theForm.prodPrice,false,"Please enter a number for the product price.")) {return(false);}
  if (!isNumeric(theForm.prodSalePrice,true,"Please enter a number for the product sale price.")) {return(false);}

	DisplaySection("Shipping");
  if (!isNumeric(theForm.prodWeight,true,"Please enter a number for the product weight.")) {return(false);}
  if (!isNumeric(theForm.prodHeight,true,"Please enter a number for the product height.")) {return(false);}
  if (!isNumeric(theForm.prodWidth,true,"Please enter a number for the product width.")) {return(false);}
  if (!isNumeric(theForm.prodLength,true,"Please enter a number for the product length.")) {return(false);}
  if (!isNumeric(theForm.prodShip,true,"Please enter a number for the product ship price.")) {return(false);}
	
	DisplaySection("Attributes");
  if (!isNumeric(theForm.attrdtPrice,true,"Please enter a number for the price variance.")) {return(false);}

  if (theForm.attrdtPrice.value != 0)
  {
	  if (theForm.attrdtType[2].checked)
	  {
	    alert("You have entered a price variance AND selected No Change.\n\n Please choose either an increase or decrease in the price.");
	    theForm.attrdtPrice.focus();
	    theForm.attrdtPrice.select();
	    return(false);
	  }
  }

	GetSortOrder(theForm);
	frmData.Show.value = strSection;
	
<% If cblnSF5AE Then %>
	//set categories values
	var theSelect = theDataForm.Categories;
	for (var i=0; i < theSelect.length;i++){theSelect.options[i].selected = true;}
<% End If %>
	
<% If cblnAddon_DynamicProductDisplay Then %>
	var theSelect = theForm.relatedProducts;
	for (var i=0; i < theSelect.length;i++){theSelect.options[i].selected = true;}
<% End If %>

	theDataForm.submit();
    return(true);
}

function ExpandProduct(theLink,lngID)
{

	if (theLink.innerHTML == "+")
	{
		theLink.innerHTML = "-";
		theLink.title = "Hide attributes";
		eval("tbl" + lngID).style.display = "";
	}
	else
	{
		theLink.innerHTML = "+";
		theLink.title = "Show attributes";
		eval("tbl" + lngID).style.display = "none";
	}
	return false;

}

function ExpandAttr(theLink,lngID)
{
	if (theLink.innerHTML == "+")
	{
		theLink.innerHTML = "-";
		theLink.title = "Hide attributes";
		eval("tblAttDetail" + lngID).style.display = "";
	}
	else
	{
		theLink.innerHTML = "+";
		theLink.title = "Show attributes";
		eval("tblAttDetail" + lngID).style.display = "none";
	}
	return false;
}

var strSubSection = "Status";
function DisplaySection(strSection)
{
<% If Not mblnShowTabs Then Response.Write "return false;" %>

<% 
Dim pstrTempHeaderRow

pstrTempHeaderRow = "'General','Detail','Attributes','Shipping'"
If cblnSF5AE Then pstrTempHeaderRow = pstrTempHeaderRow & ",'MTP','Inventory','Category'"
If mclsProduct.CustomMTP Then pstrTempHeaderRow = pstrTempHeaderRow & ",'MTP'"
If cblnAddon_DynamicProductDisplay Then pstrTempHeaderRow = pstrTempHeaderRow & ",'RelatedProducts'"
If cblnAddon_ProductReview Then pstrTempHeaderRow = pstrTempHeaderRow & ",'ProductReview'"
If isArray(mclsProduct.CustomValues) Then pstrTempHeaderRow = pstrTempHeaderRow & ",'Custom'"
pstrTempHeaderRow = pstrTempHeaderRow & ",'ProductSales','SEO'"
%>
var arySections = new Array(<%= pstrTempHeaderRow %>);

	frmData.Show.value = strSection;

	for (var i=0; i < arySections.length;i++)
	{
		if (arySections[i] == strSection)
		{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "";
				document.all("td" + arySections[i]).className = "hdrSelected";
			}
		}else{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "none";
				document.all("td" + arySections[i]).className = "hdrNonSelected";
			}
		}
	}
}

function DisplayMainSection(strSection)
{

	var arySections = new Array('Filter', 'Summary', 'itemDetail');

	for (var i=0; i < arySections.length;i++)
	{
		if (arySections[i] == strSection)
		{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "";
				document.all("td" + arySections[i]).className = "hdrSelected";
			}
		}else{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "none";
				document.all("td" + arySections[i]).className = "hdrNonSelected";
			}
		}
	}
	
	if (document.all("tblSummaryFunctions") != null)
	{
 		if (strSection == "Summary")
		{
			document.all("tblSummaryFunctions").style.display = "";
		}else{
			document.all("tblSummaryFunctions").style.display = "none";
		}
	}

	return(false);
}

var mdicCategory = new ActiveXObject("Scripting.Dictionary");
var mblnFoundNoCatKey = false;
var mstrNoCatKey = "3";

<%= mclsProduct.ProdDictionaryList %>

function InitializeCategory()
{
	var theSelect = theDataForm.CatSource;
	
	for (var i=0; i < theSelect.length;i++)
	{
		if (mdicCategory.Exists(theSelect.options[i].value))
		{
		mdicCategory(theSelect.options[i].value) = theSelect.options[i].text;
		}
	}

}

function CleanCategory()
{

	if (!mblnFoundNoCatKey)
	{
		var theSelect = theDataForm.CatSource;	
		for (var i=0; i < theSelect.length;i++)
		{
			if (theSelect.options[i].text == "No Category")
			{
				mstrNoCatKey = theSelect.options[i].value;
				mblnFoundNoCatKey = true;
				break;
			}
		}
//		alert(mstrNoCatKey);
	}
	
	if (mdicCategory.Count > 1)
	{
		if (mdicCategory.Exists(mstrNoCatKey)){mdicCategory.Remove(mstrNoCatKey)}
	}else{
		if (mdicCategory.Count == 0){mdicCategory.Add (mstrNoCatKey,"No Category")}
	}
}

function FillCategory()
{
	CleanCategory();
	var theSelect = theDataForm.Categories;
	var pary = (new VBArray(mdicCategory.Keys())).toArray();
	var plngKey;
	var theOption;
	
	theSelect.length = 0;
	
	try
	{
	for (var i=0; i < pary.length;i++)
	{
		plngKey = pary[i];
		theOption = new Option(mdicCategory(plngKey), plngKey);
		theSelect.options.add(theOption);
	}
	}
	catch(e)
	{
	return false;
	}
}

function AddCategory()
{
	var theSelect = theDataForm.CatSource;
	var mblnAdded = false;
	if (theSelect.length > 0)
	{
		for (var i=0; i < theSelect.length;i++)
		{
			if (theSelect.options[i].selected)
			{
				if (!mdicCategory.Exists(theSelect.options[i].value))
				{
				mblnAdded = true;
//alert(theSelect.options[i].value + ": " + theSelect.options[i].text);
				mdicCategory.Add (theSelect.options[i].value,theSelect.options[i].text)
				}
			}
		}
	}
	if (mblnAdded){FillCategory();}
}

function DeleteCategory()
{
	var theSelect = theDataForm.Categories;
	var mblnAdded = false;
	
	for (var i=theSelect.length-1; i >=0 ;i--)
	{
		if (theSelect.options[i].selected)
		{
			mdicCategory.Remove(theSelect.options[i].value);
			mblnAdded = true;
		}
	}
	if (mblnAdded){FillCategory();}

}

function ChangeAE(strSection)
{
document.all("Change" + strSection).value = "True";
}

function setDblClickDefault(theField)
{

var strFieldName = theField.name;

	switch (strFieldName)
	{
	<% For i = 0 To UBound(maryImageFields) %>
		case "<%= maryImageFields(i)(1) %>":
			theField.value = replaceProductID('<%= maryImageFields(i)(3) %>');
			break;
	<% Next 'i %>
		case "prodLink":
			document.frmData.prodLink.value = replaceProductID('<%= cstrDefaultDetailLinkPath %>');
			break;
		default:
	}

}

function replaceProductID(strText)
{
var strProductID = document.frmData.prodID.value;
var strToFind = "<prodID>";
var strTemp;

strTemp = strText.replace(strToFind, strProductID);

return strTemp;
}

function ConfirmDuplicateProduct(theField)
{
	var theForm = theField.form;
	var pstrNewProdID = theForm.CopyProduct1.value;
	if (pstrNewProdID.length > 0)
	{
		theForm.CopyProduct.value = pstrNewProdID;
		theForm.Action.value = "CopyProd";
		theForm.submit();
	}
}

function DuplicateProduct(theField)
{

if (true)
{
var pbytH = window.screen.availHeight;
var pbytW = window.screen.availWidth;
var theTable = document.all("tblDuplicateProduct");

theTable.style.top = (pbytH - 100)/2*2;
theTable.style.left = (pbytW - 200)/2*2;
theTable.style.top = (pbytH - 100)/2*2;
theTable.style.left = 50;
theTable.style.display = "";
document.frmData.CopyProduct1.focus();

return false;
}else{
	var theForm = theField.form;
	var pstrNewprodName = prompt("Enter New Product ID","New Product ID");
	if (pstrNewprodName != null){
		theForm.CopyProduct.value = pstrNewprodName;
		theForm.Action.value = "CopyProd";
		theForm.submit();
	}
}
}

function keepFocusOnDuplicateProduct(theElement)
{
return true;

var theTable = document.all("tblDuplicateProduct");
var strTableStyle = theTable.style.display;
if (strTableStyle.length == 0)
{
theElement.focus();
return false;
}

}
-->
</SCRIPT>
<center>
<%
	End With	'mclsProduct
	
End Sub 'WriteSupportingScripts

'************************************************************************************************************************************

Sub WriteFormOpener
%>
<form action="sfProductAdmin.asp" id=frmData name=frmData onsubmit="return ValidInput(this);" method=post>
<input type=hidden id=OrigprodID name=OrigprodID value="<%= mclsProduct.prodID %>">
<input type=hidden id=CopyProduct name=CopyProduct>
<input type=hidden id="attrDisplayOrder" name="attrDisplayOrder" value="">
<input type=hidden id="attrdtOrder" name="attrdtOrder" value="">
<input type=hidden id=ViewID name=ViewID>
<input type=hidden id=Action name=Action value="Update">
<input type=hidden id=blnShowSummary name=blnShowSummary value="">
<input type=hidden id=blnShowFilter name=blnShowFilter value="">
<input type=hidden id=Show name=Show value="<%= mstrShow %>">
<input type=hidden id=MainSection name=MainSection value="<%= mstrShow %>">
<input type=hidden id=AbsolutePage name=AbsolutePage value="<%= mlngAbsolutePage %>">
<input type=hidden id=OrderBy name=OrderBy value="<%= mstrOrderBy %>">
<input type=hidden id=SortOrder name=SortOrder value="<%= mstrSortOrder %>">
<input type=hidden id="chkDetailInNewWindow2" name="chkDetailInNewWindow" value="<%= mblnDetailInNewWindow %>">

<input type=hidden id=strBaseHRef name=strBaseHRef Value="<%= mstrBaseHRef %>">
<input type=hidden id=strBasePath name=strBasePath Value="<%= mstrBasePath %>">

<table style="display: none; border-style: outset; border-width: 3" cellpadding="3" cellspacing="0" border="1"  bgcolor="#FFFFFF" id="tblDuplicateProduct">
  <tr><td style="border-style: outset; border-color: steelblue; border-width: 3" bgcolor="steelblue">
	<table cellpadding="3" cellspacing="0" border="1"  bgcolor="#FFFFFF" id="tblDuplicateProduct_inner">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
      <tr>
        <th colspan="2" align="center">Duplicate Product</th>
      </tr>
      <tr>
        <td align="center" colspan="2">Current Product to be duplicated: <%= mclsProduct.prodName & " - " & mclsProduct.prodID %></td>
      </tr>
      <tr>
        <td class="Label" nowrap><span title="Enter the product ID to copy this product to">New Product ID:</SPAN></td>
        <td><input name="CopyProduct1" id="CopyProduct1" onchange='' onblur="return keepFocusOnDuplicateProduct(this);" value="" size=50></td>
      </tr>
      <tr>
        <td class="Label">&nbsp;</td>
        <td><input type="checkbox" name="CopyProductCategories" id="CopyProductCategories" value="1" checked>&nbsp;<label for="CopyProductCategories">Copy category assignments</label></td>
      </tr>
      <tr>
        <td class="Label">ProductID:</td>
        <td>
          <input class='butn' title='Create a new product based on this product' id="btnDuplicateProduct2" name=btnDuplicateProduct2 type=button value='Duplicate Product' onclick='ConfirmDuplicateProduct(this);'>
          <input class='butn' title='Cancel this operation' id="btnCancelDuplicateProduct" name="btnCancelDuplicateProduct" type="button" value="Cancel" onclick="document.all('tblDuplicateProduct').style.display = 'none';">
        </td>
      </tr>
	</table>
  </td></tr>
</table>
<% End Sub	'WriteFormOpener %>