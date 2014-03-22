<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
'********************************************************************************
'*   Sandshot Software Product Download Page                                    *
'*   Release Version   1.0	                                                    *
'*   Release Date      November 16, 2002										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file="ssl/SFLib/ssmodDownload.asp"-->
<!--#include file="SFLib/myAccountSupportingFunctions.asp"-->
<%
'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim mstrSQL
Dim mobjRSPriorOrders
Dim mlngOrderID
Dim mstrProductLink
Dim pstrPrevOrderDetailID
Dim pstrOrderDetailID
Dim pblnOddRow
Dim mblnDownloadAvailable
Dim fontclass

'**********************************************************
'*	Functions
'**********************************************************

	Sub CheckForUpgrade(byVal strProdID)
	'Purpose: Checks to see if upgraded version of product is avaialable

	Dim pstrProdLink
	Dim pstrProdPrice

	Dim pstrSQL
	Dim pobjUpgradeProducts

		If Len(strProdID) = 0 Then
			Response.Write "-"
		Else
			Set	pobjUpgradeProducts = CreateObject("adodb.recordset")
			with pobjUpgradeProducts
				.CursorLocation = 2 'adUseClient

				On Error Resume Next
				pstrSQL = "Select prodName, prodPrice, prodLink, version  From sfProducts Where prodID='" & strProdID & "'"
				.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Err.number <> 0 Then
					Response.Write "<font color=red>Error in GetRS: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
					Response.Write "<font color=red>Error in GetRS: sql = " & pstrSQL & "</font><br />" & vbcrlf
					Response.Flush
					Err.Clear
				Else
					If Not .EOF Then
						pstrProdLink = Trim(.Fields("prodLink").Value & "")
						If Len(pstrProdLink) = 0 Then pstrProdLink = "detail.asp?product_id=" & strProdID
						pstrProdLink = "<a href=" & pstrProdLink & ">Details</a>"
						pstrProdPrice = Trim(.Fields("prodPrice").Value & "")
					Else
						pstrProdLink = "Could not find product <em>" & strProdID & "</em>"
						pstrProdLink = pstrProdLink & "<br /><font color=red>Error in GetRS: sql = " & pstrSQL & "</font><br />"
					End If
				End If
			End With
			Set pobjUpgradeProducts = Nothing

			Response.Write pstrProdLink
		End If

	End Sub	'GetUpgradeProducts

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
<title><%= C_STORENAME %> - Purchase History</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<% Call preventPageCache %>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
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
<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">
<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>

<script>
$(document).ready(function() {
		$('.tdAltFont1 a > b').unwrap();
		$('.tdAltFont2 a > b').unwrap();
	$('#tblMainContent td').css('padding', '5px');
	$('#tblMainContent').css('background-image', 'none');
	$('#divCategoryTrail').hide();
	$('#divCategoryTrail').closest('td').parent().next().find('ul').css('list-style', 'none')
})
</script>

<style>
body {
	text-align: center;
}
#footer {
	background: #11013b;
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
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent" style="background-color:#DEDEDE !important; margin: 0 auto 5%;">
	<tr>
		<td>
	<%
		Call ShowMyAccountBreadCrumbsTrail("", True)
		If hasDownloadableItems Then
		Set	mobjRSPriorOrders = CreateObject("adodb.recordset")
		with mobjRSPriorOrders
			.CursorLocation = 2 'adUseClient

			On Error Resume Next

			mstrSQL = "SELECT sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrders.orderID, sfOrderAttributes.odrattrName, sfOrderAttributes.odrattrAttribute, sfProducts.prodlink, sfProducts.version, sfProducts.releaseDate, sfProducts.UpgradeVersion, sfOrders.orderDate, sfOrderDetails.odrdtID, sfOrderDetails.odrdtDownloadExpiresOn, sfOrderDetails.odrdtMaxDownloads, sfOrderDetails.odrdtDownloadAuthorized" _
					& " FROM ((sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
					& " WHERE ((sfOrders.orderIsComplete=1) AND (sfOrders.orderCustId=" & VisitorLoggedInCustomerID & "))" _
					& " ORDER BY sfOrderDetails.odrdtProductName, sfOrders.orderID DESC"

			.Open mstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			If Err.number <> 0 Then
				Response.Write "<font color=red>Error in GetRS: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
				Response.Write "<font color=red>Error in GetRS: sql = " & mstrSQL & "</font><br />" & vbcrlf
				Response.Flush
				Err.Clear
			Else

				On Error Goto 0
				If .EOF Then
				%>
				<table width="100%" cellpadding="0" cellspacing="0" border="0" rules="none"
				<tr>
					<tr><th><font color="red">We could not locate any orders for you.</font></th></tr>
					<tr><td>Please note this may be a result of you having multiple logins to our system. Please contact us with your order number(s) at <a href="mailto:support@sandshot.net&subject=Missing%20Order">support@sandshot.net</a> </td></tr>
				</table>
				<%
				Else

				%>

				<span class="clsCurrentLocation">Product Order History</span>
				<table width="75%" cellpadding="2" cellspacing="0" border="1" style="border-collapse: collapse; background-color:#DEDEDE; margin: 0 auto 5%;">
					<colgroup>
					  <col valign="top" />
					  <col valign="top" />
					  <col valign="top" />
					  <col valign="top" />
					  <col valign="top" />
					  <col valign="top" />
					  <col valign="top" />
					</colgroup>
					<tr>
					<td class="tdContentBar"><strong>Product</strong></td>
					<td align="center" class="tdContentBar"><strong>Order</strong></td>
					<td align="center" class="tdContentBar"><strong>Order Date</strong></td>
					<td align="center" class="tdContentBar"><strong>Download</strong></td>
					<td align="center" class="tdContentBar"><strong>Current Version</strong></td>
					<td align="center" class="tdContentBar"><strong>Release Date</strong></td>
					<td align="center" class="tdContentBar"><strong>Upgrade Available</strong></td>
					</tr>
					<%
					pblnOddRow = True
					Do While Not .EOF
						' Do alternating colors and fonts
						If pblnOddRow Then
							fontclass = "tdAltFont1"
						Else
							fontclass = "tdAltFont2"
						End If

						mlngOrderID = Trim(.Fields("orderID").Value & "")
						pstrOrderDetailID = Trim(.Fields("odrdtID").Value & "")
						If pstrPrevOrderDetailID <> pstrOrderDetailID Then
							mstrProductLink = Trim(.Fields("prodlink").Value & "")
							If Len(mstrProductLink) = 0 Then mstrProductLink = "detail.asp?product_id=" & Trim(.Fields("odrdtProductID").Value & "")
							mblnDownloadAvailable = HasDownloadAvailable_orderDetail(VisitorLoggedInCustomerID, pstrOrderDetailID)
							If Download_RequestStatus <> enDownloadRequest_NoDownloadAvailable Then
							%>
							<tr>
							<td class="<%= fontClass %>"><a href="<%= mstrProductLink %>"><%=  Trim(.Fields("odrdtProductName").Value & "") %></a></td>
							<td align="center" class="<%= fontClass %>">
							<a href="OrderHistory.asp?OrderID=<%= mlngOrderID %>"><%= mlngOrderID %></a>
							<div style="margin-top: 6pt;padding: 5pt 5pt 5pt 5pt;border: dashed 0pt white"><a href="requestSupport.asp?OrderItem=<%= pstrOrderDetailID %>">Request Support</a></div>
							</td>
							<td align="center" class="<%= fontClass %>"><a href="OrderHistory.asp?OrderID=<%= mlngOrderID %>"><%= FormatDateTime(Trim(.Fields("orderDate").Value & ""), 2) %></a></td>
							<td align="center" class="<%= fontClass %>">
							<%
								Select Case Download_RequestStatus
									Case enDownloadRequest_Valid:
										%>
										<a href="download.asp?OrderDetailID=<%= pstrOrderDetailID %>&amp;FileName=<%= Server.URLEncode(Download_FileName) %>" target="_blank">Download</a>
										<br /><%= FormatNumber(Download_FileSize/1000, 1,,,True) %> kb
										<%
										If Download_MaxDownloads <> 0 Then Response.Write "<br />" & Download_CurrentDownloadCount & " of " & Download_MaxDownloads & " downloads"
									Case enDownloadRequest_DownloadCountReached:	Response.Write "Download limit (<em>" & Download_MaxDownloads & "</em>) reached"
									Case enDownloadRequest_InvalidFilePath:			Response.Write "Contact customer service: Code(" & enDownloadRequest_InvalidFilePath & ")"

									Case Else: Response.Write "Unavailable - Code: <em>" & Download_RequestStatus & "</em>"
								End Select
							%>
							</td>
							<td align="center" class="<%= fontClass %>">&nbsp;<%= Trim(.Fields("version").Value & "") %></td>
							<td align="center" class="<%= fontClass %>">&nbsp;<%= Trim(.Fields("releaseDate").Value & "") %></td>
							<td align="center" class="<%= fontClass %>"><% Call CheckForUpgrade(Trim(.Fields("UpgradeVersion").Value & "")) %></td>
							</tr>
							<% End If %>
					<%
						End If	'pstrPrevOrderDetailID <> pstrOrderDetailID
						If False Then
						'If Len(Trim(.Fields("odrattrName").Value & "")) > 0 Then
							%>
							<tr><td colspan="4" class="<%= fontClass %>">&nbsp;&nbsp;<%= Trim(.Fields("odrattrName").Value & "") %>: <%= Trim(.Fields("odrattrAttribute").Value & "") %></a></td></tr>
							<%
						End If
						pstrPrevOrderDetailID = pstrOrderDetailID
						.MoveNext
						pblnOddRow = Not pblnOddRow
					Loop
					%>
				</table>
			</td>
			</tr>
			</table>
				<%
				End If
			End If
		end with
		mobjRSPriorOrders.Close
		Set mobjRSPriorOrders = Nothing
		Else
		%>
		<table width="95%" cellpadding="0" cellspacing="0" border="0" rules="none">
		<tr>
			<tr><th><font color="red">We could not locate any orders with downloadable items for you.</font></th></tr>
			<tr><td>Please note this may be a result of you having multiple logins to our system. Please contact us with your order number(s) at <a href="mailto:support@sandshot.net&subject=Missing%20Order">support@sandshot.net</a> </td></tr>
		</table>
		<%
		End If	'hasDownloadableItems
		%>
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
Call cleanup_dbconnopen
%>