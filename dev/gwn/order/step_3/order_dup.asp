<%@ Language=VBScript %>
<% Option Explicit %>
<%
'********************************************************************************
'*
'*   order.asp
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins is order.asp APPVERSION: 50.4014.0.6
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

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim mblnOrderAltered: mblnOrderAltered = False	'added to store cart to session variable
Dim iCounter
Dim iNewQuantity
Dim sBtnAction, sSaveCart, sDelete, iSvdCartID, sRecalculate, iSaveFind, iDeleteFind
Dim iTmpOrderID, iOldQuantity
Dim sProdID
Dim iQuantity
Dim aProdAttr
Dim iProdAttrNum
Dim mstrPageMessage

Dim mlngNumProducts
Dim plngCurrentQty
Dim bCustIdExists

'**********************************************************
'*	Functions
'**********************************************************

	Sub cleanupPageObjects

	On Error Resume Next

		Set mclsCartTotal = Nothing
		Call cleanup_dbconnopen

	End Sub

'**********************************************************
'*	Begin Page Code
'**********************************************************

%>
<!--#include file="incCoreFiles.asp"-->
<%

Response.Buffer= CBool(vDebug = 0)

If SessionID = "" Then
	Call cleanup_dbconnopen	'This line needs to be included to close database connection
	Response.Redirect("search.asp")
End If

mlngNumProducts = Request.Form("iProductCounter")
If Len(mlngNumProducts) = 0 Then mlngNumProducts = -1

' Determine action and OrderID
For iCounter = 0 to mlngNumProducts
	sSaveCart = Request.Form("SaveToCart" & iCounter & ".x")
	If sSaveCart <> "" Then
		iSaveFind = iCounter
		sBtnAction = "SaveToCart"
		Exit For
	End If

	sDelete	= Request.Form("DeleteFromOrder" & iCounter & ".x")
	If sDelete <> "" Then
		iDeleteFind = iCounter
		sBtnAction = "DeleteFromCart"
		Exit For
	End If
Next

' Check to see if custID exists in customer table
bCustIdExists = validCustIDCookie

' Determine if it is recalculate action
sRecalculate  = Request.Form("recalc")
If Len(Request.Form("MoveAll.x")) > 0 Then
	sBtnAction = "MoveAll"
ElseIf sRecalculate = "1" AND lcase(sBtnAction) <> "savetocart" Then
	sBtnAction = "Recalculate"
End If

Select Case sBtnAction
	Case "Recalculate"

		For iCounter = 0 To mlngNumProducts
			iNewQuantity = Request.Form("FormQuantity" & iCounter)
			iOldQuantity = Request.Form("iQuantity" & iCounter)
			iTmpOrderID = Request.Form("iOrderID" & iCounter)
			if Not isnumeric(iNewQuantity) or trim(iNewQuantity) = "" then
				iNewQuantity = iOldQuantity
			end if
			If cblnSF5AE Then Call Order_Update_GiftWrapsBackOrder(iTmpOrderID)

			If iNewQuantity <> "" Then
				If iNewQuantity = 0 Then
					' Delete if 0
					Call setDeleteOrder("odrdttmp", iTmpOrderID)

				ElseIf iNewQuantity <> iOldQuantity Then
					' Update Quantity For Product
					Call setReplaceQuantity("odrdttmp",iNewQuantity,iTmpOrderID)
				End If
			Else
				Call setDeleteOrder("odrdttmp", iTmpOrderID)
			End If
		Next
		If cblnSF5AE Then Order_AdjustCart 'SFAE b2
		mblnOrderAltered = True

		mstrPageMessage = "<h4>The quantites have been updated.</h4><br />"

	Case "SaveToCart"
		sProdID = Request.Form("sProdID" & iSaveFind)
		iTmpOrderID = Request.Form("iOrderID" & iSaveFind)
		iQuantity = Request.Form("iQuantity" & iSaveFind)
		iProdAttrNum = Request.Form("iProdAttrNum" & iSaveFind)
		aProdAttr = getProdAttr("odrattrtmp", iTmpOrderID, iProdAttrNum)
		iCustID = custID_cookie
	  	iNewQuantity = Request.Form("FormQuantity" & iSaveFind)

	  	' In the case that one types in a new quantity and hits save
	  	If iNewQuantity <> iQuantity And iNewQuantity <> "" And iNewQuantity <> 0 Then
	  		iQuantity = iNewQuantity
	  	End If

		' Check if cookies are set
		If Len(custID_cookie) = 0 OR CBool(CStr(getCookie_SessionID) <> CStr(SessionID)) Then
			' Write to cookie identifying place
			Call getSavedTable(aProdAttr, sProdID, iNewQuantity, 0, vistor_HTTP_REFERER)
			Response.Cookies("sfThanks")("PreviousAction") = "FromShopCart"
			Response.Cookies("sfThanks")("DeleteTmpOrderID") = iTmpOrderID
			Response.Cookies("sfThanks").Expires = Date() + 1
			Call cleanup_dbconnopen	'This line needs to be included to close database connection
			Response.Redirect("login.asp")
		End If

		iSvdCartID = getOrderID("odrdtsvd", "odrattrsvd", sProdID, aProdAttr, cInt(iProdAttrNum), plngCurrentQty)

		If iSvdCartID <> "" Then
			If iSvdCartID < 0 Then		' New Row in SavedCartDetails
		  		Call getSavedTable(aProdAttr,sProdID,iQuantity,iCustID,vistor_HTTP_REFERER)	'Write as new row
			Else		' Existing cart
				Call setUpdateQuantity("odrdtsvd",iQuantity,iSvdCartID)	'Update Quantity
			End If		' End iSvdCartID exists If
		Else
			'sErrorDescription =  "Number of attributes not equal to the product specs or database writing error."
			'Response.Redirect("error.asp?strPageName=order.asp&strErrorDescription="&sErrorDescription)
		End If	' End iSvdCartID Null If

		Call setDeleteOrder("odrdttmp", iTmpOrderID)	'delete from sfTmpOrderDetails
		mblnOrderAltered = True

		mstrPageMessage = "<h4>" & getProductInfo(sProdID, enProduct_Name) & " has been placed in your saved cart.</h4><p align=""center""><a href=""savecart.asp"" alt=""View your saved cart""><img class=""inputImage"" src=""" & C_BTN08 & """ alt=""View your saved cart""></a></p><br />"

	Case "MoveAll"
		iCustID = custID_cookie

		For iSaveFind = 0 To mlngNumProducts
			sProdID = Request.Form("sProdID" & iSaveFind)
			iTmpOrderID = Request.Form("iOrderID" & iSaveFind)
			iQuantity = Request.Form("iQuantity" & iSaveFind)
			iProdAttrNum = Request.Form("iProdAttrNum" & iSaveFind)
	  		iNewQuantity = Request.Form("FormQuantity" & iSaveFind)
	  		If iNewQuantity <> iQuantity And iNewQuantity <> "" And iNewQuantity <> 0 Then iQuantity = iNewQuantity

			aProdAttr = getProdAttr("odrattrtmp", iTmpOrderID, iProdAttrNum)
			If True Then
				Response.Write "<fieldset><legend>Move All Items (" & iSaveFind & ")</legend>"
				Response.Write "sProdID: " & sProdID & "<br />"
				Response.Write "iSvdOrderID: " & iSvdOrderID & "<br />"
				Response.Write "iQuantity: " & iQuantity & "<br />"
				Response.Write "iNewQuantity: " & iNewQuantity & "<br />"
				Response.Write "iProdAttrNum: " & iProdAttrNum & "<br />"
				Response.Write "</fieldset>"
			End If

			' Check if cookies are set
			If Len(custID_cookie) = 0 OR CBool(CStr(getCookie_SessionID) <> CStr(SessionID)) Then
				Call getSavedTable(aProdAttr, sProdID, iNewQuantity, 0, vistor_HTTP_REFERER)
			Else
				iSvdCartID = getOrderID("odrdtsvd", "odrattrsvd", sProdID, aProdAttr, cInt(iProdAttrNum), plngCurrentQty)

				If iSvdCartID <> "" Then
					If iSvdCartID < 0 Then		' New Row in SavedCartDetails
		  				Call getSavedTable(aProdAttr,sProdID,iQuantity,iCustID,vistor_HTTP_REFERER)	'Write as new row
					Else		' Existing cart
						Call setUpdateQuantity("odrdtsvd",iQuantity,iSvdCartID)	'Update Quantity
					End If		' End iSvdCartID exists If
				Else
					'sErrorDescription =  "Number of attributes not equal to the product specs or database writing error."
					'Response.Redirect("error.asp?strPageName=order.asp&strErrorDescription="&sErrorDescription)
				End If	' End iSvdCartID Null If
			End If
			Call setDeleteOrder("odrdttmp", iTmpOrderID)	'delete from sfTmpOrderDetails
			mblnOrderAltered = True

			If Len(mstrPageMessage) = 0 Then
				mstrPageMessage = "<h4>" & getProductInfo(sProdID, enProduct_Name) & " has been placed in your saved cart.</h4>"
			Else
				mstrPageMessage = mstrPageMessage & "<h4>" & getProductInfo(sProdID, enProduct_Name) & " has been placed in your saved cart.</h4>"
			End If

		Next
		mstrPageMessage = mstrPageMessage & "<p align=""center""><a href=""savecart.asp"" alt=""View your saved cart""><img class=""inputImage"" src=""" & C_BTN08 & """ alt=""View your saved cart""></a></p><br />"

		' Check if cookies are set
		If Len(custID_cookie) = 0 OR CBool(CStr(getCookie_SessionID) <> CStr(SessionID)) Then
			' Write to cookie identifying place
			Response.Cookies("sfThanks")("PreviousAction") = "FromShopCart"
			Response.Cookies("sfThanks")("DeleteTmpOrderID") = iTmpOrderID
			Response.Cookies("sfThanks").Expires = Date() + 1
			Call cleanup_dbconnopen	'This line needs to be included to close database connection
			Response.Redirect("login.asp")
		Else
			Call closeObj(cnn)
			Response.Redirect "savecart.asp"
		End If

	Case "DeleteFromCart"
		iTmpOrderID = Request.Form("iOrderID" & iDeleteFind)
		Call setDeleteOrder("odrdttmp", iTmpOrderID)
		mblnOrderAltered = True

		mstrPageMessage = "<h4>The selected item has been removed from your cart.</h4><br />"

End Select	'sBtnAction

If mblnOrderAltered Then
	Call InitializeCart
	mclsCartTotal.removeCartFromSession
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Begin Checkout Page</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Pragma" content="no-cache">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="keywords">
<meta name="description" content="description">
<meta name="Robot" content="Index,ALL">
<meta name="revisit-after" content="15 Days">
<meta name="Rating" content="General">
<meta name="Language" content="en">
<meta name="distribution" content="Global">
<meta name="Classification" content="classification">
  <link runat="server" rel="shortcut icon" type="image/png" href="favicon.ico">
  <link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Lato:100,400,900|Josefin+Sans:100,400,700,400italic,700italic">
  <link rel="stylesheet" href="css/main.css">
  <link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">
<script language="javascript" src="SFLib/jquery-1.10.2.min.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/incae.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">

var noState_State = "N/A";

function visitorCountry_onchange(theSelect)
{
	var strSelectedCountry = theSelect.options[theSelect.selectedIndex].value;
	if ((strSelectedCountry == 'US') || (strSelectedCountry == 'CA'))
	{
		// do nothing
	}else{
		letSelectValue(theSelect.form.visitorState, noState_State);
	}
}

function validShippingPreference(theForm)
{

	var strSelectedCountry = theForm.visitorCountry.options[theForm.visitorCountry.selectedIndex].value;
	if ((strSelectedCountry == 'US') || (strSelectedCountry == 'CA'))
	{
		if (theForm.visitorState.options[theForm.visitorState.selectedIndex].value == noState_State)
		{
			alert("Please select the State or Province.");
			theForm.visitorState.focus();
			return (false);
		}

		if (theForm.visitorZIP.value == "")
		{
			alert("Please enter the Zip Code.");
			theForm.visitorZIP.focus();
			return (false);
		}

	}

	return (true);
}

function letSelectValue(theSelect,theValue)
{

	for (var i = 0;  i < theSelect.options.length;  i++)
	{
		if (theSelect.options[i].value == theValue)
		{
			theSelect.selectedIndex = i;
			return true;
		}
	}

	return false;
}

</script>

<script>
$(document).ready(function() {
	$('#tblMainContent td').css('padding', '5px');
})
</script>

<% If cblnSF5AE Then Call Order_ShowInventoryMessage 'SFAE %>
<% writeCurrencyConverterOpeningScript %>
<style>
body {
	text-align: center;
}
</style>
</head>
<body <%= mstrBodyStyle %>>

	<div id="header" style="margin-bottom: 2%;">
    <div id="gwn_logo">
      <a href="index.html" title="Home"><image src="images/gwn_logo.png" alt="GameWearNow Logo"></a>
    </div>
    <div id="heading">
      <span class="title_txt" id="title">CUSTOM JERSEYS FOR<br>YOUR SPORTS TEAM</span>
        <br>
      <span class="title_txt" id="sub_title">ORDER SUMMARY</span>
    </div>
  </div>

<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent" style="margin: 0 auto 5%;">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td class="tdBottomTopBanner"><span class="Error"><strong><br>
			</strong></span>Please review your order as shown below. To modify the quantity of any item ordered,
            input the desired quantity and select the <b>Update Quantity Changes</b>
            button below. To delete an item click on <b>DELETE</b>.<br><br>
            <%if IsSaveCartActive = 1 then %>
            <!--webbot bot="PurpleText" preview="Begin Save Cart Message" -->
            To save an item to return and purchase at a later time click on <b>
          Save for Later</b>. If you want to add new items return to the correct page and select additional items to be added to your order. When you have
			completed your order, select the Check Out button below to connect to our secure directory and complete the order process.
			<!--webbot bot="PurpleText" preview="End Saved Cart Message" -->
			<% end if %>
        </tr>
		<!--webbot bot="PurpleText" PREVIEW="Begin Optional Confirmation Message Display" -->
		<%
			Call WriteThankYouMessage
			If Len(mstrPageMessage) > 0 Then Response.Write "<tr><td class=""tdContent2""><br>" & mstrPageMessage & "</td></tr>"
		%>
		<!--webbot bot="PurpleText" PREVIEW="End Optional Confirmation Message Display" -->
        <tr>
          <td class="tdContent2">
          <% 'Call DisplayShippingTimeMessage %>
          <%
			Call InitializeCart
			With mclsCartTotal

				.City = visitorCity
				.State = visitorState
				.ZIP = visitorZIP
				.Country = visitorCountry
				.isCODOrder = False

				.ShipMethodCode = visitorPreferredShippingCode
				.LoadAllShippingMethods = True

				.checkInventoryLevels
				'.writeDebugCart

				.MinimumOrderMessage = adminMinOrderMessage
				.MinimumOrderAmount = adminMinOrderAmount

				.displayOrder
			End With	'mclsCartTotal

          %>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--webbot bot="PurpleText" preview="End Content Section" -->

  <div id="footer">
    <ul id="horizontal-nav">
      <li id="current_page"><a title="Shopping Cart"><span><image src="images/shopping_cart.png" alt="Shopping Cart" id="shopping_cart">MY SHOPPING CART</span></a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="myAccount.asp" title="My Account">MY ACCOUNT</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/faqs/faqs.html" title="FAQ's">FAQ'S</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/privacy_policy/privacy_policy.html" title="Privacy Policy">PRIVACY POLICY</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/contact_us/contact_us.html" title="Contact Us">CONTACT US <font>(877) 796-6639</font></a></li>
    </ul>
  </div>

<script>
	$(document).ready(function(){
		$('hr').hide();
		$('.tdAltFont1 a > b').unwrap();
		$('.tdAltFont2 a > b').unwrap();
		$('#frmCheckout').attr('action', 'http://dev.gamewearnow.com/sll/process_order.asp');
		$("img[src='images/buttons/continueshop.gif']").css('margin-bottom', '10px').parent().attr('href','index.html');
	});
</script>
</body>
</html>
<% Call cleanupPageObjects %>