<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
'********************************************************************************
'*
'*   search_results.asp -
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins are search_results.asp APPVERSION: 50.4014.0.3
'*   from LaGarde, Incorporated. It has been heavily modified by Sandshot Software
'*   with permission from LaGarde, Incorporated who has NOT released any of the
'*   original copyright protections. As such, this is a derived work covered by
'*   the copyright provisions of both companies.
'*
'*   LaGarde, Incorporated Copyright Statement                                                                           *
'*   The contents of this file is protected under the United States copyright
'*   laws and is confidential and proprietary to LaGarde, Incorporated.  Its
'*   use ordisclosure in whole or in part without the expressed written
'*   permission of LaGarde, Incorporated is expressly prohibited.
'*   (c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'*
'*   Sandshot Software Copyright Statement
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file="SFLib/ssSearchGrid.asp"-->
<%

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim iPageSize, iMaxRecords
Dim txtsearchParamTxt, txtsearchParamType, txtsearchParamCat, txtFromSearch,  txtsearchParamMan
Dim txtCatName, txtsearchParamVen, txtImagePath, txtOutput, txtDateAddedStart
Dim txtDateAddedEnd, txtPriceStart, txtPriceEnd, txtSale, SQL, sAmount, rsCatImage
Dim iAttCounter, irsSearchAttRecordCount, iAttDetailCounter, irsSearchAttDetailRecordCount
Dim iPage, iRec, iNumOfPages, iDesignCounter, iVarPageSize, iSearchRecordCount, icounter, iDesign
Dim rsCat, rsSearch, rsSearchAtt, rsSearchAttDetail, arrAttDetail, arrProduct, arrAtt, rsManufacturer, rsVendor
Dim sSubCat,sALLSUB,X,sMainCat ,iLevel
Dim iSubCat
Dim pstrSQL
Dim pstrSQL_OrderBy

'**********************************************************
'*	Functions
'**********************************************************

'**********************************************************
'*	Begin Page Code
'**********************************************************

	On Error Resume Next
	Response.Buffer = CBool(CStr(vDebug) <> "1")
	If Err.number <> 0 Then Err.Clear
	'On Error Goto 0

	'Added to require login prior to seeing pricing
	'If Not isLoggedIn Then Response.Redirect "myAccount.asp?PrevPage=search_results.asp" & Server.URLEncode("?" & Request.QueryString)

	iDesign	= C_DesignType		'Layout Selection
	iDesign	= 1		'0: None; 1: Left; 2: Right; 3: Alternating
	iDesignCounter = 2

	Call setVisitorLastSearch(Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString)

	Call LoadSearchParameters

	'search_results, salespage, and newproduct are all based off this page
	'comment out the specific section below for each of the pages

	'for salespage use the following section
	'txtSale = 1

	'for newproduct use the following section
	'Dim mIDaysBack
	'If Len(visitorLastVisited) > 0 Then mIDaysBack = DateDiff("d", Now(), visitorLastVisited)
	'If mIDaysBack < clngNewProductsDaysSinceAdded Then mIDaysBack = clngNewProductsDaysSinceAdded
	'txtDateAddedStart = MakeUSDate(Date() - clngNewProductsDaysSinceAdded)

	'set search parameters for the category navigation menu
	mstrCurrentCategory = txtsearchParamCat
	mstrCurrentSubCategory = sSubCat
	mstrCurrentSubSubCategory = iLevel

	Call LoadSort(pstrSQL_OrderBy, iPageSize)
	iVarPageSize = iPageSize

	' Determine the page user is requesting
	iPage = Trim(Request.QueryString("PAGE"))
	If isNumeric(iPage) And Len(iPage) > 0 Then
		iPage = CLng(iPage)
		If iPage < 1 Then iPage = 1
	Else
		iPage = 1
	End If

	Call BuildCustomSearchFilter ( _
				SQL, _
				txtsearchParamType, _
				txtsearchParamTxt, _
				txtsearchParamCat, _
				txtsearchParamMan, _
				txtsearchParamVen, _
				txtDateAddedStart, _
				txtDateAddedEnd, _
				txtPriceStart, _
				txtPriceEnd, _
				txtSale, _
				sSubCat, _
				iLevel)	'added for Sandshot Software Customer Search Filter Display

	If Len(SQL) > 0 Then
		SQL = SQL & pstrSQL_OrderBy
	Else
		'This case results when no products meet search criteria
		SQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription, sfProducts.prodPLPrice, sfProducts.prodPLSalePrice, sfVendors.vendName, sfManufacturers.mfgName" _
			& " FROM sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN sfProducts ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId" _
			& " WHERE (sfProducts.prodEnabledIsActive=1) AND (sfProducts.prodEnabledIsActive=0)"
	End If
	If vDebug = 1 Then Response.Write SQL & "<br /><br />"

	'On Error Resume Next
	'On Error Goto 0
	Set rsSearch = CreateObject("ADODB.RecordSet")
	With rsSearch
		.CursorLocation = adUseClient
		.Open SQL, cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
		If Err.number <> 0 Then	' And vDebug = 1
			Response.Write "<fieldset><legend>Error getting search results</legend>"
			Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />"
			Response.Write "SQL: " & SQL & "<br />"
			Response.Write "</fieldset>"
			Response.Flush
			Err.Clear
		End If

		If .EOF then
			iSearchRecordCount = 0
		Else
			Call AdjustRecordPricingLevel(rsSearch, "SearchProduct")
			arrProduct = .GetRows()
			iSearchRecordCount = mlngNumProductsFound
			iNumOfPages = Int(iSearchRecordCount / iPageSize)

			SQL = "SELECT attrID, attrName, attrProdID, attrDisplayStyle FROM sfAttributes WHERE attrProdID In (" & mstrProductIDList & ")" _
				& " ORDER BY attrDisplayOrder"
			If Len(SQL) > 0 Then
				If vDebug = 1 Then Response.Write  "SearchAtt SQL: " & SQL & "<br /><br />"
				Set rsSearchAtt = CreateObject("ADODB.RecordSet")
				rsSearchAtt.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
				If Not rsSearchAtt.EOF Then
					irsSearchAttRecordCount = rsSearchAtt.RecordCount - 1

					arrAtt = rsSearchAtt.GetRows
					SQL = arrAtt(0,0)
					For iRec = 1 To irsSearchAttRecordCount
						SQL = SQL & "," & arrAtt(0,iRec)
					Next 'iRec
					SQL = "SELECT attrdtID, attrdtAttributeId, attrdtName, attrdtPrice, attrdtType, attrdtOrder, attrdtPLPrice" _
						& " FROM sfAttributeDetail " _
						& " WHERE attrdtAttributeId In (" & SQL & ")" _
						& " ORDER BY attrdtOrder"
					If vDebug = 1 Then Response.Write "SearchAttDetail SQL: " & SQL & "<br /><br />"

					If Len(SQL) > 0 Then
						Set rsSearchAttDetail = CreateObject("ADODB.RecordSet")
						rsSearchAttDetail.CursorLocation = adUseClient
						rsSearchAttDetail.Open SQL, cnn, adOpenStatic,adLockBatchOptimistic, adCmdText

						If Not rsSearchAttDetail.EOF Then
							Call AdjustRecordPricingLevel(rsSearchAttDetail, "SearchAttDetail")
							arrAttDetail = rsSearchAttDetail.GetRows
							irsSearchAttDetailRecordCount = rsSearchAttDetail.RecordCount - 1
						End If
					End If	'Len(SQL) > 0

				End If	'Not rsSearchAtt.EOF

			End If	'Len(SQL) > 0

		End If	'.EOF

	End With	'rsSearch
	Call closeObj(rsSearchAttDetail)
	Call closeObj(rsSearchAtt)
	Call closeObj(rsSearch)
	'On Error Goto 0
	Call recordSearchResults(txtsearchParamType, txtsearchParamTxt, iSearchRecordCount)

	If CInt(iNumOfPages + 1) = CInt(iPage) Then iVarPageSize = iSearchRecordCount - (iNumOfPages * iPageSize)
	If iSearchRecordCount mod iPageSize <> 0 Then iNumOfPages = iNumOfPages + 1

	If mblnCategorySearch Then

		If cblnSF5AE Then
			Set mclsCategory = New ssCategoryAE
			With mclsCategory
				If Len(iLevel) > 0 And isNumeric(iLevel) Then
					If iLevel > 1 Then
						.SubcategoryFilter = sSubCat
					Else
						.CategoryFilter = txtsearchParamCat
					End If
				End If

				If .LoadCategories Then
					mstrCustomDescription = .CurrentCategoryDescription
					mstrCustomName = .CurrentCategoryName
					mstrCustomImage = .CurrentCategoryImage
				End If
			End With
			Set mclsCategory = Nothing
		Else
			Set mclsCategory = New ssCategorySE
			With mclsCategory
				If .LoadCategory(txtsearchParamCat, txtsearchParamMan, txtsearchParamVen) Then
					mstrCustomDescription = .CurrentCategoryDescription
					mstrCustomName = .CurrentCategoryName
					mstrCustomImage = .CurrentCategoryImage
				End If
			End With
			Set mclsCategory = Nothing
		End If

	ElseIf mblnManufacturerSearch Then
		If isNumeric(txtsearchParamMan) Then
			mstrCustomName = getMfgVendItem(txtsearchParamMan, "Name", True)
		End If
	ElseIf mblnVendorSearch Then
		If isNumeric(txtsearchParamVen) Then
			mstrCustomName = getMfgVendItem(txtsearchParamVen, "Name", False)
		End If
	Else
		mstrCustomName = txtCatName
	End If
	mstrCustomNameWrapped = mstrCustomName

	txtCatName = mstrCustomName
	'debugprint "txtCatName", txtCatName

	If vDebug = 1 Then Response.Write SQL & "<br /><br />"
	If cblnSF5AE Then
		'txtCatName = "All " & C_CategoryNameP
		'mstrCustomName = txtCatName
	Else
		'mstrCustomName = Response.write  arrProduct(7, iRec)
	End If
	If txtsearchParamTxt = "" Then txtsearchParamTxt = "*"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= Server.HTMLEncode(mstrCustomName) %></title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% Call preventPageCache %>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="<%= Server.HTMLEncode(txtCatName) %>">
<meta name="description" content="<%= Server.HTMLEncode(mstrCustomDescription) %>">
<meta name="Robot" content="Index,ALL">
<meta name="revisit-after" content="15 Days">
<meta name="Rating" content="General">
<meta name="Language" content="en">
<meta name="distribution" content="Global">
<meta name="Classification" content="classification">
<link runat="server" rel="shortcut icon" type="image/png" href="favicon.ico">
<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">
<link rel="stylesheet" href="css/main.css">
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/incae.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfEmailFriend.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/ssAttributeExtender.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/jquery-1.11.0.min.js" type="text/javascript"></script>
<script language=javascript type="text/javascript">

	$(document).ready(function() {

	var media = navigator.userAgent.toLowerCase();
	var isMobile = media.indexOf("mobile") > -1;
	if(isMobile) {
		$('#horizontal-nav li').css('padding-right', '10px');
	};

		WebFontConfig = {
		  google: { families: [ 'Lato:100,400,900:latin', 'Josefin+Sans:100,400,700,400italic,700italic:latin' ] }
		  };
		  (function() {
		    var wf = document.createElement('script');
		    wf.src = ('https:' == document.location.protocol ? 'https' : 'http') +
		      '://ajax.googleapis.com/ajax/libs/webfont/1/webfont.js';
		    wf.type = 'text/javascript';
		    wf.async = 'true';
		    var s = document.getElementsByTagName('script')[0];
		    s.parentNode.insertBefore(wf, s);
		})();

		$('table.tdTopBanner').next().css('margin', '0 auto 10%');
		$('#frmPromo table').css('margin', '0 auto');

		$(".not_selected").hover(
		  function() {
		    $('#current_page a').css('color','#cccdce');
		  }, function() {
		    $('#current_page a').css('color','#e8d606');
		  }
		);
	});

<!--
function ValidateForm(theForm)
{
var i;
var pos;
var ProductID;
var elemName;
var blnValidForm = false;

	for (i = 0; i < theForm.length; i++)
	{
		if (theForm[i].type == "text")
		{
			elemName = theForm[i].name;
			if (elemName.indexOf("QUANTITY.") != -1)
			{
				if (theForm[i].value != "")
				{
					pos = elemName.indexOf(".");
					ProductID = elemName.substring(pos + 1, elemName.length);

					if (validateCustomProducts(theForm, ProductID))
					{
						blnValidForm = true;
					}else{
						return false;
					}
				}
			}
		}
	}

	if (blnValidForm)
	{
		return true;
	}else{
		alert("You must enter a quantity for at least one item.");
		return false;
	}

}

function validateCustomProducts(theForm, strProductID)
{
var i;
var elemName;

	for (i = 0; i < theForm.length; i++)
	{
		elemName = theForm[i].name;
		if (elemName.indexOf(strProductID) != -1)
		{
			if (theForm[i].optional == false)
			{
				if (theForm[i].value == "")
				{
					alert("Please enter a value for "+ theForm[i].title);
					theForm[i].focus();
					return false;
				}
			}
		}
	}
	return true;
}

//-->
</SCRIPT>
<script language="javascript" type="text/javascript">
	function setCategoryDetails(theForm)
	{
		var theSelect=theForm.CatSource;
		var theValue=theSelect.options[theSelect.options.selectedIndex].value;
		var pstrTemp=theValue.split(".");

		theForm.txtsearchParamCat.value=pstrTemp[0];
		theForm.subcat.value=pstrTemp[1];
		theForm.iLevel.value=pstrTemp[2];
		theForm.txtCatName.value=pstrTemp[2];
	}

</script>
<% writeCurrencyConverterOpeningScript %>

<style>
body {
	background-image: url('images/splash_bg.jpg');
	text-align: center;
}
#tdContent {
	padding:	10px;
}
#tblCategoryMenu, .tdTopBanner, .tdLeftNav {
	display: none;
}
</style>

</head>

<body <%= mstrBodyStyle %>>

	<div id="header">
    <div id="gwn_logo">
      <a href="index.html" title="Home"><image src="images/gwn_logo.png" alt="GameWearNow Logo" style="margin-left: -25px;"></a>
    </div>
    <div id="heading">
      <span class="title_txt" id="title">CUSTOM JERSEYS FOR<br>YOUR SPORTS TEAM</span>
        <br>
      <span class="title_txt" id="sub_title">SEARCH RESULTS</span>
    </div>
  </div>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
	<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
	  <tr>
		<td>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
		<tr>
			<td align="left" class="tdContent"><!--&nbsp;&nbsp;<div class="clsCategoryTrail" id="divCategoryTrail"><% Call writeCategoryTrail(txtsearchParamCat, iLevel - 1, sSubcat, "", True) %></div>--></td>
			<td align="right" class="tdContent">
			<form id='frmSearchSearch' name='frmSearchSearch' action='search_results.asp' method='Get' onsubmit="setCategoryDetails(this);">
			<input type="hidden" id="txtFromSearch" name="txtFromSearch" value="fromSearch">
			<input type="hidden" id="txtsearchParamType" name="txtsearchParamType" value="ALL">
			<input type="hidden" id="txtsearchParamMan" name="txtsearchParamMan" value="ALL">
			<input type="hidden" id="txtsearchParamVen" name="txtsearchParamVen" value="ALL">
			<input type="hidden" id="iLevel" name="iLevel" value="1">
			<input type="hidden" id="subcat" name="subcat" value="">
			<input type="hidden" id="txtCatName" name="txtCatName" value="">
			<input name="txtsearchParamTxt" ID="txtsearchParamTxt" value="<% If txtsearchParamTxt <> "*" Then Response.Write txtsearchParamTxt %>" size="8" />&nbsp;in&nbsp;
			<% Call writeCategorySelect(sSubCat) %>
			<a href="<%= C_HomePath %>advancedSearch.asp" onclick="setCategoryDetails(document.frmSearchSearch); document.frmSearchSearch.submit(); return false;">Search</a>
			</form>
			</td>
		</tr>
		</table>
		</td>
	  </tr>
	  <tr>
		<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td class="tdContent2" align="center"><br /><% Call displaySearchCustomizationBar(0) %></td>
		  </tr>
		  <!--#include file="include_files/search_results_output.asp"-->
		  <tr>
			<td class="tdContent2" align="center"><% Call displaySearchCustomizationBar(1) %><br /></td>
		  </tr>
		</table>
		</td>
	  </tr>
	</table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->

	<div id="footer">
    <ul id="horizontal-nav">
      <li class="not_selected"><a href="order.asp" title="Shopping Cart"><span><image src="../../images/shopping_cart.png" alt="Shopping Cart" id="shopping_cart">MY SHOPPING CART</span></a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="myAccount.asp" title="My Account">MY ACCOUNT</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/faqs/faqs.html" title="FAQ's">FAQ'S</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/privacy_policy/privacy_policy.html" title="Contact Us">PRIVACY POLICY</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/contact_us/contact_us.html" title="Contact Us">CONTACT US <font>(877) 796-6639</font></a></li>
    </ul>
  </div>

</body>
</html>
<%
' Object Cleanup
On Error Resume Next

If Not isEmpty(mclsCategory) Then Set mclsCategory = Nothing
Call cleanup_dbconnopen
%>