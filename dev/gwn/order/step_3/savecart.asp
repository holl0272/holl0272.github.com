<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="incCoreFiles.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.2

'@FILENAME: savecart.asp

'@DESCRIPTION: Displayes all products on sale

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'Modified 10/25/01
'Storefront Ref#'s: 177 'JF

Dim sCondition
Dim sEmail
Dim sPassword
Dim iAuthenticate
Dim bCustIdExists
Dim fontclass
Dim pstrAttributeText
Dim mstrMessage
Dim mstrMessageItem
Dim sSql, rsAllSvdOrders, sProdID, aProdAttr, sProdName, sProdPrice, iProdAttrNum, iCounter
Dim sAttrUnitPrice, sUnitPrice, iQuantity, iNewQuantity, sProductSubtotal, dProductSubtotal, iSvdOrderID, aProdAttrID, sTotalPrice, sProductPrice
Dim iProductCounter, sBgColor, sFontFace, sFontColor, iFontSize
Dim sPaymentList, bHasProducts, sBtnAction, sAddCart, sDelete, iTmpCartID, iAddFind, iDeleteFind, sReferer
Dim plngCurrentQty
Dim iTmpOrderID, iOldQuantity
Dim pstrProdLink
Dim pstrImageSRC

bCustIdExists = validCustIDCookie

'-------------------------------------------------------
' If login button is depressed
'-------------------------------------------------------

If Trim(Request.Form("btnLogin.x")) <> "" Then
	sEmail			= Trim(Request.Form("Email"))
	sPassword		= Trim(Request.Form("Password"))

	' Authenticate
	iCustID = customerAuth(sEmail, sPassword, "loose")
	If iCustID = -1  Then
		If customerAuth(sEmail, sPassword, "loosest") > 0 Then
			sCondition = "EmailMatch"
			Call expireCookie_sfCustomer
		Else
			sCondition = "WrongCombination"
			Call expireCookie_sfCustomer
		End If
	End If
ElseIf Trim(Request.Form("SignUp.x")) <> "" Then
	sEmail			= Trim(Request.Form("Email"))
	sPassword		= Trim(Request.Form("Password"))

	' Authenticate
	iCustID	= customerAuth(sEmail, sPassword, "loose")
	If iCustID = -1 Then
		iCustID = getCustomerID(sEmail, sPassword)
		Call SetSessionLoginParameters(iCustID, sEmail)
	End If
End If

' Determine action and OrderID

' Determine if it is recalculate action
If Len(Request.Form("Recalculate.x")) > 0 Then
	sBtnAction = "Recalculate"
ElseIf Len(Request.Form("MoveAll.x")) > 0 Then
	sBtnAction = "MoveAll"
Else
	For iCounter = 1 to Request.Form("iProductCounter")
		sAddCart = Request.Form("AddToCart" & iCounter & ".x")
		If sAddCart <> "" Then
			iAddFind = iCounter
			sBtnAction = "AddToCart"
			Exit For
		End If

		sDelete	= Request.Form("DeleteFromOrder" & iCounter & ".x")
		If sDelete <> "" Then
			iDeleteFind = iCounter
			sBtnAction = "DeleteFromCart"
			Exit For
		End If
	Next
End If

Select Case sBtnAction
	Case "MoveAll"
		iCustID = custID_cookie
		sReferer = vistor_HTTP_REFERER

		For iCounter = 1 To Request.Form("iProductCounter")
			iAddFind = iCounter
			iSvdOrderID = Request.Form("iSvdOrderID" & iAddFind)
			sProdID = Request.Form("sProdID" & iAddFind)
			iQuantity = Request.Form("iQuantity" & iAddFind)
			iProdAttrNum = Request.Form("iProdAttrNum" & iAddFind)
			iNewQuantity = Request.Form("FormQuantity" & iAddFind)
			If Not isNumeric(iNewQuantity) or trim(iNewQuantity) = "" then iNewQuantity = iOldQuantity

			If True Then
				Response.Write "<fieldset><legend>Move All Items (" & iAddFind & ")</legend>"
				Response.Write "sProdID: " & sProdID & "<br />"
				Response.Write "iSvdOrderID: " & iSvdOrderID & "<br />"
				Response.Write "iQuantity: " & iQuantity & "<br />"
				Response.Write "iNewQuantity: " & iNewQuantity & "<br />"
				Response.Write "iProdAttrNum: " & iProdAttrNum & "<br />"
				Response.Write "</fieldset>"
			End If

	  		' In the case that one types in a new quantity and hits add
	  		If iNewQuantity <> iQuantity And iNewQuantity <> "" Then iQuantity = iNewQuantity

			aProdAttr = getProdAttr("odrattrsvd", iSvdOrderID, iProdAttrNum)

			iTmpCartID = getOrderID("odrdttmp", "odrattrtmp", sProdID, aProdAttr, cInt(iProdAttrNum), plngCurrentQty)

			If Len(iTmpCartID) > 0 Then
				If iTmpCartID < 0 Then	'New Row in SavedCartDetails,
			  		iTmpCartID = getTmpTable(aProdAttr, sProdID, iQuantity, sReferer, getProductInfo(sProdID, enProduct_Ship))
			  		mstrMessageItem = "<strong>" & getProductInfo(sProdID, enProduct_Name) & "</strong> has been moved to your shopping cart."
				Else	'Existing cart, Update Quantity
					Call setUpdateQuantity("odrdttmp",iQuantity,iTmpCartID)
			  		mstrMessageItem = "The quantity of <strong>" & getProductInfo(sProdID, enProduct_Name) & "</strong> in your shopping cart has been updated."
				End If
			Else
				Response.Write "<p>Number of attributes not equal to the product specs or database writing error"
			  	mstrMessageItem = "There was an error adding <strong>" & getProductInfo(sProdID, enProduct_Name) & "</strong> to your shopping cart."
			End If

			If Len(mstrMessage) = 0 Then
				mstrMessage = mstrMessageItem
			Else
				mstrMessage = mstrMessage & "<br />" & mstrMessageItem
			End If

			If cblnSF5AE Then Call SaveCart_WriteSvdtmpAERecord(iSvdOrderID, iTmpCartID)

			' delete from sfSavedOrderDetails
			Call setDeleteOrder("odrdtsvd", iSvdOrderID)

		Next
		Call cleanup_dbconnopen	'This line needs to be included to close database connection
		Response.Redirect "order.asp"

	Case "Recalculate"

		For iCounter = 1 To Request.Form("iProductCounter")
			iNewQuantity = Request.Form("FormQuantity" & iCounter)
			iOldQuantity = Request.Form("iQuantity" & iCounter)
			iSvdOrderID = Request.Form("iSvdOrderID" & iCounter)
			if not isnumeric(iNewQuantity) or trim(iNewQuantity) = "" then iNewQuantity = iOldQuantity
			If iNewQuantity <> "" Then
				If iNewQuantity = 0 Then
					' Delete if 0
					Call setDeleteOrder("odrdtsvd",iSvdOrderID)
				ElseIf iNewQuantity <> iOldQuantity Then
					' Update Quantity For Product
					Call setReplaceQuantity("odrdtsvd",iNewQuantity,iSvdOrderID)
				End If
			Else
				' Delete if Null Value
				Call setDeleteOrder("odrdtsvd",iSvdOrderID)
			End If
		Next

	Case "AddToCart"

		sProdID = Request.Form("sProdID" & iAddFind)
		iSvdOrderID = Request.Form("iSvdOrderID" & iAddFind)
		iQuantity = Request.Form("iQuantity" & iAddFind)
		iProdAttrNum = Request.Form("iProdAttrNum" & iAddFind)
		iCustID = custID_cookie
		sReferer = vistor_HTTP_REFERER
		iNewQuantity = Request.Form("FormQuantity" & iAddFind)

	  	' In the case that one types in a new quantity and hits add
	  	If iNewQuantity <> iQuantity And iNewQuantity <> "" Then iQuantity = iNewQuantity

		aProdAttr = getProdAttr("odrattrsvd", iSvdOrderID, iProdAttrNum)

		iTmpCartID = getOrderID("odrdttmp", "odrattrtmp", sProdID, aProdAttr, cInt(iProdAttrNum), plngCurrentQty)

		If Len(iTmpCartID) > 0 Then
			If iTmpCartID < 0 Then	'New Row in SavedCartDetails,
			  	iTmpCartID = getTmpTable(aProdAttr, sProdID, iQuantity, sReferer, getProductInfo(sProdID, enProduct_Ship))
			  	mstrMessageItem = "<strong>" & getProductInfo(sProdID, enProduct_Name) & "</strong> has been moved to your shopping cart."
			Else	'Existing cart, Update Quantity
				Call setUpdateQuantity("odrdttmp",iQuantity,iTmpCartID)
			  	mstrMessageItem = "The quantity of <strong>" & getProductInfo(sProdID, enProduct_Name) & "</strong> in your shopping cart has been updated."
			End If
		Else
			Response.Write "<p>Number of attributes not equal to the product specs or database writing error"
			mstrMessageItem = "There was an error adding <strong>" & getProductInfo(sProdID, enProduct_Name) & "</strong> to your shopping cart."
		End If

		If Len(mstrMessage) = 0 Then
			mstrMessage = mstrMessageItem
		Else
			mstrMessage = mstrMessage & "<br />" & mstrMessageItem
		End If

		If cblnSF5AE Then Call SaveCart_WriteSvdtmpAERecord(iSvdOrderID, iTmpCartID)

		' delete from sfSavedOrderDetails
		Call setDeleteOrder("odrdtsvd", iSvdOrderID)

	Case "DeleteFromCart"
		iSvdOrderID = Request.Form("iSvdOrderID" & iDeleteFind)
		Call setDeleteOrder("odrdtsvd",iSvdOrderID)
End Select	'sBtnAction

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Save Cart Page</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% Call preventPageCache %>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
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
<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">
<link rel="stylesheet" href="css/main.css">
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/incae.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfEmailFriend.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/jquery-1.11.0.min.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">

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
    $("img[alt='View Cart']").css('margin', '10px');
    $("img[alt='Return to Shop']").css('margin', '10px').parent().attr('href', 'index.html');

    $(window).load(function() {
      $('#tdContent').css('opacity', 1);
      $('#tblMainContent').css('opacity', 1);
    });
  });
</script>

<style>
body {
  background-image: url('images/splash_bg.jpg');
  text-align: center;
}
#tdContent {
  margin: 0 auto 5%;
  opacity: 0;
}
#tblCategoryMenu, .tdTopBanner, .tdLeftNav {
  display: none;
}
.tdAltFont1, .tdAltFont2 {
  background-color: white;
}
.inputImage {
  padding: 10px;
  margin-top: 0;
}
input[type='text'] {
  margin: 10px;
}
input[type='password'] {
  margin: 10px;
}
</style>

<% writeCurrencyConverterOpeningScript %>
</head>
<body <%= mstrBodyStyle %>>

  <div id="header">
    <div id="gwn_logo">
      <a href="index.html" title="Home"><image src="images/gwn_logo.png" alt="GameWearNow Logo" style="margin-left: -25px;"></a>
    </div>
    <div id="heading">
      <span class="title_txt" id="title">CUSTOM JERSEYS FOR<br>YOUR SPORTS TEAM</span>
    </div>
  </div>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="15" cellpadding="0" id="tblMainContent" style="opacity: 0;">
  <tr>
    <td>
      <% If Len(mstrMessage) > 0 Then %>
      <div class="tdContent" style="border: 1 solid black;text-align:center;padding-top:10pt;padding-bottom:10pt"><%= mstrMessage %></div>
      <% End If %>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="center" class="tdMiddleTopBanner">Your <%=Application("CartName")%></td>
        </tr>
        <tr>
          <td class="tdBottomTopBanner">Please
            review your <%=Application("CartName")%> items as shown below. To
            modify the quantity of any item, simply input the desired quantity
            and select the <b>Recalculate <%=Application("CartName")%></b> button.
            To delete an item click on <b>DELETE</b>. To ADD an item to your order
            for purchase, click on <b>ADD</b>. If you want to add new items, you
            can do so by pressing RETURN TO SHOP and click on <%=Application("CartSaveButton")%>
            for the appropriate product. You can access your <%=Application("CartName")%>
            at any time.
        </tr>
        <tr>
          <td class="tdContent2">
            <table border="0" width="100%" cellspacing="0" cellpadding="4">
              <tr>
                <td width="1%" align="left" class="tdContentBar"></td>
                <td width="49%" class="tdContentBar">Item</td>
                <td width="15%" align="center" class="tdContentBar">Unit Price</td>
                <td width="10%" align="center" class="tdContentBar">Qty</td>
                <td width="10%" align="center" class="tdContentBar">Price</td>
                <td width="15%" align="center" class="tdContentBar">action</td>
              </tr>
              <% If Not isLoggedIn Then	%>
              <tr>
                <td colspan="6" align="center"> <br />
                  <form action="savecart.asp" method="post" onSubmit="this.form=true;return sfCheck(this);">
                    <table border="0" width="50%" cellpadding="0" cellspacing="1">
                      <tr>
                        <td width="50%" align="center" class="tdContent" valign="middle">
                          <table border="0" width="100%" cellpadding="3" cellspacing="1">
                            <tr>
                              <td width="100%" align="center" class="tdBottomTopBanner2"><b><%=Application("CartName")%> Login</b></td>
                            </tr>
                            <% If sCondition = "EmailMatch" Then %>
                            <tr>
                              <td width="100%" align="center" class="tdContent">
                                <font class="Error"> <b> Email Match </b> </font>
                                <br />
                                Please choose another email account</td>
                            </tr>
                            <% ElseIf sCondition = "WrongCombination" Then %>
                            <tr>
                              <td width="100%" align="center" class="tdContent"><font class="Error"><b>Wrong
                                Combination of email/password</b></font> <br />
                                Please try again</td>
                            </tr>
                            <% End If %>
                            <tr>
                              <td width="100%" align="center" valign="middle" class="tdContent">
                                <table border="0" width="100%" cellpadding="2">
                                  <tr>
                                    <td width="15%" align="right"><b>
                                      E-Mail:</b></td>
                                    <td width="85%">
                                      <input type="text" size="40" name="Email"  title="E-Mail Address" class="formDesign">
                                    </td>
                                  </tr>
                                  <tr>
                                    <td width="15%" align="right"><b>Password:</b></td>
                                    <td width="85%">
                                      <input type="password" size="40" name="Password" title="Password" class="formDesign">
                                    </td>
                                  </tr>
                                  <tr>
                                    <td width="100%" align="center" colspan="2">
                                      <input type="image" class="inputImage" src="<%= C_BTN16 %>" name="btnLogin">
                                      <input type="image" class="inputImage" src="<%= C_BTN12 %>" name="SignUp">
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                            <tr>
                              <td width="100%" align="center" class="tdContent"><a href="myAccount.asp">Forgot your password?</a> <br />
                                </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              <%
	Else
		iCustID = visitorLoggedInCustomerID	'custID_cookie
		Call setCombineProducts(iCustID)
		sSql = "SELECT * FROM sfSavedOrderDetails WHERE odrdtsvdCustID=" & makeInputSafe(iCustID)
		If vDebug = 1 Then Response.Write "<br /> " & sSql

		Set rsAllSvdOrders = cnn.execute(sSql)

		' Check for no orders
		If rsAllSvdOrders.EOF Then
			bHasProducts = False
			%>
              <tr>
                <td colspan="6" class="tdAltFont1"><font class='Middle_Top_Banner_Small'>
                  <p style="margin-top:25pt">
                    <center>
                      <b><font size="+1">No Items in <%=Application("CartName")%></font></b>
                      <br />
                      Please press return to shop button to begin searching for items. <br />
                    </center>
                  </p>
                  </font> </td>
              </tr>
              <%
		Else
			bHasProducts = True

			Do While NOT rsAllSvdOrders.EOF

				iSvdOrderID = rsAllSvdOrders.Fields("odrdtsvdID").Value
				sProdID = rsAllSvdOrders.Fields("odrdtsvdProductID").Value
				iQuantity = rsAllSvdOrders.Fields("odrdtsvdQuantity").Value

			If Not getProductInfo(sProdID, enProduct_Exists) Then
				Call setDeleteOrder("odrdtsvd", iSvdOrderID)
				Response.Write "<br />Product Does Not Exist"
			Else
	  			sProdName = getProductInfo(sProdID, enProduct_Name)
	  			sProdPrice = getProductInfo(sProdID, enProduct_SellPrice)
	  			iProdAttrNum = getProductInfo(sProdID, enProduct_AttrNum)
				If NOT IsNumeric(iProdAttrNum)Then iProdAttrNum = 0

				' Get Associated Attribute IDs in an array
				If iProdAttrNum <> "" Then
					ReDim aProdAttrID(iProdAttrNum)
					aProdAttrID = getProdAttr("odrattrsvd", iSvdOrderID, iProdAttrNum)
				End If

				' Response Write all Output
				If vDebug = 1 Then
					Response.Write "<p>Product = " & sProdID & "<br />ProdName = " & sProdName & "<br />ProdPrice = " & sProdPrice & "<br />ProdAttrNum = " & iProdAttrNum
					'Call ShowRow("odrdtsvd","odrattrsvd",iSvdOrderID,sProdID)
					If IsArray(aProdAttrID) Then
						For iCounter = 0 To iProdAttrNum -1
							Response.Write "<br />Attribute :" & aProdAttrID(iCounter)
						Next
					End If

				End If

				iProductCounter = iProductCounter + 1
				' Do alternating colors and fonts
				If (iProductCounter mod 2) = 1 Then
					fontclass="tdAltFont1"
				Else
					fontclass="tdAltFont2"
				End If

				If Len(CStr(cbytOrderViewImageSrc)) > 0 Then
					pstrImageSRC = getProductInfo(sProdID, enProduct_ImageSmallPath)
				Else
					pstrImageSRC = ""
				End If

				pstrProdLink = getProductInfo(sProdID, enProduct_Link)
				If Len(pstrProdLink) = 0 Then pstrProdLink = "detail.asp?product_id=" & sProdID

				%>
              <form method="POST" action="savecart.asp" id="form2" name="form2">
                <tr>
                  <td class="<%= fontclass %>" valign="top" align="left"><% If Len(pstrImageSRC) > 0 Then %><a href="<%= pstrProdLink %>"><img src="<%= pstrImageSRC %>" alt="<%= stripHTML(sProdName) %>" /></a><% End If 'Len(pstrImageSRC) > 0 %></td>
                  <td nowrap align="left" valign="top" class='<%= fontclass %>'><a href="<%= pstrProdLink %>"><b><%= sProdName %></b></a><br />
                    <%
						' Begin with 0
						sAttrUnitPrice = 0

						' Iterate Through Attributes
						If iProdAttrNum > 0 And IsArray(aProdAttrID) Then
							Dim sAttrSubtotal, aAttrDetails, sAttrName, sAttrPrice, iAttrType
							For iCounter = 0 To iProdAttrNum - 1
								aAttrDetails = getAttrDetails(aProdAttrID(iCounter))
								sAttrName = aAttrDetails(0)
								sAttrPrice = aAttrDetails(1)
								iAttrType = aAttrDetails(2)

								' Calculate Subtotal
								sAttrUnitPrice =  getAttrUnitPrice(sAttrUnitPrice,sAttrPrice,iAttrType)

					If aProdAttrID(iCounter) = GetAttributeID(aProdAttrID(iCounter)) Then
						pstrAttributeText = Trim(getSavedAttributeText(iSvdOrderID, aProdAttrID(iCounter)))
					Else
						pstrAttributeText = GetAttributeValue(aProdAttrID(iCounter))
					End If

					If Right(sAttrName, 1) = ":" Then sAttrName = Left(sAttrName, Len(sAttrName) - 1)
					If Len(pstrAttributeText) > 0 Then sAttrName = sAttrName & ": " & pstrAttributeText
					%>
                    &nbsp;&nbsp;<%= sAttrName %> <br />
                    <%
							' ProdAttr Loop
							Next
						' Else the attributes don't exist in database.  Best to delete it?
						Elseif iProdAttrNum > 0 And NOT IsArray(aProdAttrID) Then
							Response.Write "<br />Error: No Attributes found for " & iSvdOrderID
							Response.Write "<br /> Deleting from" & Application("CartName") & ". Sorry for the inconvenience."

							Call setDeleteOrder("odrdtsvd",iSvdOrderID)
							If vDebug = 1 Then Response.Write "<p><font color=""red"">" & iSvdOrderID & "</font>"
						' End Product Attribute If
						End If

						' Set Unit Price for Product
						If getProductInfo(sProdID, enProduct_IsActive) = 1 Then 'djp
							If iConverion = 1 Then
								sUnitPrice = "<script>document.write(""" & FormatCurrency(cDbl(sAttrUnitPrice) + cDbl(sProdPrice)) & " = ("" + OANDAconvert(" & cDbl(sAttrUnitPrice) + cDbl(sProdPrice) & "," & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"
							Else
								sUnitPrice = FormatCurrency(cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
							End If
							dProductSubtotal = iQuantity * (cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
							If iConverion = 1 Then
								sProductSubtotal = "<script>document.write(""" & FormatCurrency(dProductSubtotal) & " = ("" + OANDAconvert(" & dProductSubtotal & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"
							Else
								sProductSubtotal = FormatCurrency(dProductSubtotal)
							End If
						End if

						If getProductInfo(sProdID, enProduct_IsActive) = 0 Then Response.Write "<h4>This product is no longer available</h4>"
					%>
                    </td>
                  <td align="center" class='<%=fontClass%>' valign="top" nowrap><%= sUnitPrice %></td>
                  <td align="center" class='<%=fontClass%>' valign="top" nowrap>
                    <% If getProductInfo(sProdID, enProduct_IsActive) = 1 Then %>
                    <input type="text" class="formDesign" name="FormQuantity<%= iProductCounter%>" size="2" value="<%= iQuantity %>">
                    <% Else %>
                    -<input type="hidden" name="FormQuantity<%= iProductCounter%>" size="2" value="0">
                    <% End If %>
                  </td>
                  <td align="center" class='<%=fontClass%>' valign="top" nowrap><%= sProductSubtotal %></td>
                  <td align="center" class='<%=fontClass%>' valign="top" nowrap>
                    <input type="hidden" name="sProdID<%= iProductCounter%>" value="<%=sProdID%>">
                    <input type="hidden" name="iSvdOrderID<%= iProductCounter%>" value="<%=iSvdOrderID%>">
                    <input type="hidden" name="iQuantity<%= iProductCounter%>" value="<%=iQuantity%>">
                    <input type="hidden" name="iProdAttrNum<%= iProductCounter%>" value="<%=iProdAttrNum%>">
                    <input type="image" class="inputImage" src="<%= C_BTN06 %>" name="DeleteFromOrder<%= iProductCounter%>">
                    <br />
                    <%If getProductInfo(sProdID, enProduct_IsActive) = 1 Then%>
                    <input type="image" class="inputImage" src="<%= C_BTN22 %>" name="AddToCart<%= iProductCounter%>">
                    <%End If %>
                  </td>
                </tr>
                <%
				End If

			rsAllSvdOrders.MoveNext
		Loop

	'@ENDCODE

	'-----------------------------------------------------------
	' END PRODUCT DETAIL OUTPUT --------------------------------
	'-----------------------------------------------------------
	%>
                <tr>
                  <td colspan="6" align="center">
                    <hr noshade size="1" width="90%">
                  </td>
                </tr>
                <tr>
                  <td colspan="6" align="right" valign="top">
                    <input type="hidden" name="iProductCounter" value="<%= iProductCounter%>">
                    <input type="image" class="inputImage" src="<%= C_BTN14 %>" name="Recalculate"><br />
                    <input type="image" class="inputImage" src="images/buttons/addallitems.gif" name="MoveAll" id="MoveAll">
                  </td>
                </tr>
                <tr>
                  <td colspan="6" align="center">
                    <hr noshade size="1" width="90%">
                  </td>
                </tr>
              </form>
            </table>
            <%
	'-----------------------------------------------------------
	' SUBTOTAL OUTPUT  taken out 'SFUPDATE
	'-----------------------------------------------------------
	%>
            <table border="0" width="100%" cellspacing="0" cellpadding="2">
              <%
	 ' End rsAllSvdOrders If
	End If

' End Cookie If
End If

	%>
              <tr>
                <td width="100%" colspan="6" align="center"> <a href="order.asp"><img src="<%= C_BTN10 %>" alt="View Cart" border="0"></a>
                  <a href="<%= getLastSearch %>"><img src="<%= C_BTN09 %>" border="0" alt="Return to Shop"></a>
                  <% If isLoggedIn Then %>
                  <form action="login.asp" method="post" id=form1 name=form1>
                    <input type="image" class="inputImage" src="<%= C_BTN11 %>" name="ChangeCart">
                    <% If bHasProducts And iEmailActive = 1 Then %> <br /><a href="javascript:emailwishlist()"><img border="0" src="<%= C_BTN24 %>" alt="Email your Wish List to Friend(s)"></a><% End If %>
                  </form>
                  <%End If %>
                  <%If bHasProducts Then %>
                  <div class="Content_Small">Please Note: None of these items will
                  be in checkout unless you explicitly add them to your order.</div>
                  <%End If %>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
                    </tr>
                  </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->

  <div id="footer">
    <ul id="horizontal-nav">
      <li id="current_page"><a href="order.asp" title="Shopping Cart"><span><image src="../../images/shopping_cart.png" alt="Shopping Cart" id="shopping_cart">MY SHOPPING CART</span></a></li>
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

closeObj(rsAllSvdOrders)
Call cleanup_dbconnopen

%>