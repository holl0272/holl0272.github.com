<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
'********************************************************************************
'*                                                                              *
'*   1.00.001 (June 15, 2006)													*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file="SFLib/myAccountSupportingFunctions.asp"-->
<%
'/////////////////////////////////////////////////
'/
'/  User Parameters
'/


'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Global variables
'**********************************************************

Dim fontclass
Dim mblnDownloadAvailable
Dim mlngOrderID
Dim mobjRSPriorOrders
Dim mstrProductLink
Dim mstrPrevOrderDetailID
Dim mstrOrderDetailID
Dim mblnOddRow

'**********************************************************
'*	Functions
'**********************************************************

	Function getPriorOrderSummary(byVal lngCustID, byRef objRS)

	Dim pstrSQL

		pstrSQL = "SELECT sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrders.orderID, sfOrderAttributes.odrattrName, sfOrderAttributes.odrattrAttribute, sfProducts.prodLink, sfOrders.orderDate, sfOrderDetails.odrdtID, sfOrderDetails.odrdtDownloadExpiresOn, sfOrderDetails.odrdtMaxDownloads, sfOrderDetails.odrdtDownloadAuthorized, sfProducts.prodFileName" _
				& " FROM ((sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
				& " WHERE ((sfOrders.orderIsComplete=1) AND (sfOrders.orderCustId=" & lngCustID & "))" _
				& " ORDER BY sfOrderDetails.odrdtProductName, sfOrders.orderID DESC"

		Set	objRS = CreateObject("adodb.recordset")
		with objRS
			.CursorLocation = 2 'adUseClient

			On Error Resume Next
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			If Err.number <> 0 Then
				Response.Write "<font color=red>Error in getPriorOrderSummary: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
				Response.Write "<font color=red>SQL: " & pstrSQL & "</font><br />" & vbcrlf
				Response.Flush
				Err.Clear
				getPriorOrderSummary = False
			Else
				getPriorOrderSummary = True
			End If
		End With

	End Function	'getPriorOrderSummary

'**********************************************************
'*	Begin Page Code
'**********************************************************

'Only let in valid logins
If Not isLoggedIn Then
	Call cleanup_dbconnopen	'This line needs to be included to close database connection
	Response.Redirect "myAccount.asp?PrevPage=" & Request.ServerVariables("SCRIPT_NAME") & Server.URLEncode("?" & Request.QueryString)
End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Order History by Products</title>
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
  <link runat="server" rel="shortcut icon" type="../image/png" href="favicon.ico">
  <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Lato:100,400,900|Josefin+Sans:100,400,700,400italic,700italic">
  <link rel="stylesheet" href="../css/main.css">
<script language="javascript" src="../SFLib/jquery-1.10.2.min.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/incae.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfEmailFriend.js" type="text/javascript"></script>

<script>
$(document).ready(function() {
		$('.tdAltFont1 a > b').unwrap();
		$('.tdAltFont2 a > b').unwrap();
	$('#tblMainContent td').css('padding', '5px');
	$('#divCategoryTrail').hide();
})
</script>

<style>
body {
	text-align: center;
}
</style>
</head>
<body <%= mstrBodyStyle %>>

		<div id="header" style="margin-bottom: 2%;">
    <div id="gwn_logo">
      <a href="../index.html" title="Home"><image src="../images/gwn_logo.png" alt="GameWearNow Logo"></a>
    </div>
    <div id="heading">
      <span class="title_txt" id="title">CUSTOM JERSEYS FOR<br>YOUR SPORTS TEAM</span>
        <br>
      <span class="title_txt" id="sub_title">ORDER HISTORY</span>
    </div>
  </div>


<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent" style="background-color:#DEDEDE; margin: 0 auto 5%;">
	<tr>
		<td>
	<%
	Call ShowMyAccountBreadCrumbsTrail("", True)
	If getPriorOrderSummary(visitorLoggedInCustomerID, mobjRSPriorOrders) Then
		With mobjRSPriorOrders
			If .EOF Then
			%>
			<table>
			  <tr>
                <th><font color="red" class="Error">We could not locate any orders for you.</font></th>
			  </tr>
			</table>
			<%
			Else
			%>
			<table width="100%" border="0" cellpadding="2" cellspacing="0" align="center">
			  <colgroup>
			    <col align="left" width="" />
			    <col width="" />
			    <col width="" />
			  </colgroup>
              <tr>
                <td class="tdContentBar">Product Title</td>
                <td align="center" class="tdContentBar">Order Number</td>
                <td align="center" class="tdContentBar">Order Date</td>
              </tr>
              <%
			    mblnOddRow = True
			    Do While Not .EOF
					' Do alternating colors and fonts
					If mblnOddRow Then
						fontclass = "tdAltFont1"
					Else
						fontclass = "tdAltFont2"
					End If

					mlngOrderID = Trim(.Fields("orderID").Value & "")
					mstrOrderDetailID = Trim(.Fields("odrdtID").Value & "")
					If mstrPrevOrderDetailID <> mstrOrderDetailID Then
						mstrProductLink = Trim(.Fields("prodlink").Value & "")
						If Len(mstrProductLink) = 0 Then mstrProductLink = "detail.asp?product_id=" & Trim(.Fields("odrdtProductID").Value & "")
						%>
							<tr>
							<td class='<%= fontClass %>'><a href="<%= mstrProductLink %>"><%=  Trim(.Fields("odrdtProductName").Value & "") %></a></td>
							<td class='<%= fontClass %>'><a href="OrderHistory.asp?OrderID=<%= mlngOrderID %>"><%= mlngOrderID %></a></td>
							<td class='<%= fontClass %>'><a href="OrderHistory.asp?OrderID=<%= mlngOrderID %>"><%= FormatDateTime(Trim(.Fields("orderDate").Value & ""), 2) %></a></td>
							</tr>
						<%
					End If 'mstrPrevOrderDetailID <> mstrOrderDetailID

					Dim pstrAttributeCategory
					Dim pstrAttributeDetail

					If Len(Trim(.Fields("odrattrName").Value & "")) > 0 Then
						pstrAttributeCategory = Trim(.Fields("odrattrName").Value & "")
						pstrAttributeDetail = Trim(.Fields("odrattrAttribute").Value & "")
						'Now adjust for attribute extender which MAY save the category name: in the detail
						If inStr(1, pstrAttributeDetail, pstrAttributeCategory & ": ") = 1 Then pstrAttributeDetail = Replace(pstrAttributeDetail,  pstrAttributeCategory & ": ", "")
						%>
						<tr><td class='<%= fontClass %>' colspan="3">&nbsp;&nbsp;<%= pstrAttributeCategory %>: <%= pstrAttributeDetail %></td></tr>
						<%
						End If	'Len(Trim(.Fields("odrattrName").Value & "")) > 0

						mstrPrevOrderDetailID = mstrOrderDetailID
					.MoveNext
					mblnOddRow = Not mblnOddRow
			    Loop
              %>
			</table>
			<%
			End If	'.EOF
			.Close
		End With
		Set mobjRSPriorOrders = Nothing
	End If	'getPriorOrderSummary(visitorLoggedInCustomerID, mobjRSPriorOrders)
	%>
        </td>
      </tr>
    </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->


  <div id="footer">
    <ul id="horizontal-nav">
      <li class="not_selected"><a title="Shopping Cart"><span><image src="../images/shopping_cart.png" alt="Shopping Cart" id="shopping_cart">MY SHOPPING CART</span></a></li>
      <li class="pipe">|</li>
      <li id="current_page"><a href="myAccount.asp" title="My Account">MY ACCOUNT</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/faqs/faqs.html" title="FAQ's">FAQ'S</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/privacy_policy/privacy_policy.html" title="Privacy Policy">PRIVACY POLICY</a></li>
      <li class="pipe">|</li>
      <li class="not_selected"><a href="footer/contact_us/contact_us.html" title="Contact Us">CONTACT US <font>(877) 796-6639</font></a></li>
    </ul>
  </div>
</body>
</html>
<%
	closeObj(mobjRSPriorOrders)
	Call cleanup_dbconnopen
%>