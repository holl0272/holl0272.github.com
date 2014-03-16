<% 
'********************************************************************************
'*   myAccount Version SF 5.0		                                            *
'*   Release Version:	1.00.003                                                *
'*   Release Date:		September 29, 2002										*
'*   Revision Date:		September 30, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Dim pblnShowMyAccount_ReturnLink:				pblnShowMyAccount_ReturnLink = True
Dim pblnShowMyAccount_ReOrderLink:				pblnShowMyAccount_ReturnLink = True
Dim pblnShowMyAccount_OrderHistoryLink:			pblnShowMyAccount_OrderHistoryLink = True
Dim pblnShowMyAccount_ProductOrderHistoryLink:	pblnShowMyAccount_ProductOrderHistoryLink = True
Dim pblnShowMyAccount_DownloadLink:				pblnShowMyAccount_DownloadLink = True
Dim pblnShowMyAccount_PendingNotifications:		pblnShowMyAccount_PendingNotifications = False
Dim pstrShowMyAccount_DefaultCurrentText

'*************************************************************************************************************************

Sub checkMyAccountLinks()

Dim pstrCurrentPage

	'check broad categories
	If hasPriorOrders Then
		pstrCurrentPage = LCase(currentPage)
		If inStr(1, pstrCurrentPage, "orderhistorybyproducts.asp") > 0 Then
			pblnShowMyAccount_ProductOrderHistoryLink = False
			pstrShowMyAccount_DefaultCurrentText = "Product Order History"
		ElseIf inStr(1, pstrCurrentPage, "orderhistory.asp") > 0 Then
			pblnShowMyAccount_OrderHistoryLink = False
			pstrShowMyAccount_DefaultCurrentText = "Order History"
		ElseIf inStr(1, pstrCurrentPage, "priororders.asp") > 0 Then
			pblnShowMyAccount_DownloadLink = False
			pstrShowMyAccount_DefaultCurrentText = "Product Downloads"
		ElseIf inStr(1, pstrCurrentPage, "neworder.asp") > 0 Then
			pblnShowMyAccount_ReOrderLink = False
			pstrShowMyAccount_DefaultCurrentText = "Re-Order"
		ElseIf inStr(1, pstrCurrentPage, "myaccount_pendingnotifications.asp") > 0 Then
			pblnShowMyAccount_PendingNotifications = False
			pstrShowMyAccount_DefaultCurrentText = "Pending Notifications"
		ElseIf inStr(1, pstrCurrentPage, "myaccount.asp") > 0 Then
			pblnShowMyAccount_ReturnLink = False
			pstrShowMyAccount_DefaultCurrentText = ""
		Else
			pstrShowMyAccount_DefaultCurrentText = ""
		End If

		pblnShowMyAccount_DownloadLink = pblnShowMyAccount_DownloadLink And hasDownloadableItems
	Else
		pblnShowMyAccount_ReOrderLink = False
		pblnShowMyAccount_OrderHistoryLink = False
		pblnShowMyAccount_ProductOrderHistoryLink = False
		pblnShowMyAccount_DownloadLink = False
	End If	'hasPriorOrders

End Sub	'checkMyAccountLinks

'*************************************************************************************************************************

Sub ShowMyAccountBreadCrumbsTrail(byVal strBreadcrumbsTrailFinalText, byVal blnShowQuickMenu)

Dim paryBreadcrumbs
Dim pstrBreadcrumbsText

	Call checkMyAccountLinks
	If strBreadcrumbsTrailFinalText = "-" Then
		pstrBreadcrumbsText = ""
	ElseIf Len(strBreadcrumbsTrailFinalText) > 0 Then
		pstrBreadcrumbsText = strBreadcrumbsTrailFinalText
	Else
		pstrBreadcrumbsText = pstrShowMyAccount_DefaultCurrentText
	End If
	
	If InStr(1, pstrBreadcrumbsText, "{custGreeting}") > 0 Then
		If Session("custGreeting") <> "" Then pstrBreadcrumbsText = Replace(pstrBreadcrumbsText, "{custGreeting}", Session("custGreeting"))
	End If
	
	If Len(pstrBreadcrumbsText) > 0 Then
		paryBreadcrumbs = Array("Home", "default.asp", "Return to home page", "My Account", "myAccount.asp", "Return to my account", pstrBreadcrumbsText)
	Else
		paryBreadcrumbs = Array("Home", "default.asp", "Return to home page", "My Account")
	End If
%>
		<br />
		<table border="0" cellspacing="0" width="95%">
			<tr>
			<td><%= BreadCrumbsTrail(paryBreadcrumbs) %></td>
			</tr>
			<% If blnShowQuickMenu Then %>
			<tr>
			<td align="right">
				<br />
				<div align="right">
				<ul>
				<% If pblnShowMyAccount_ReturnLink Then %><li><A href="myAccount.asp">Return to myAccount</A></li><% End If %>
				<% If pblnShowMyAccount_ReOrderLink Then %><li><A href="neworder.asp">Re-Order from a past order</A></li><% End If %>
				<% If pblnShowMyAccount_OrderHistoryLink Then %><li><A href="orderHistory.asp">Order History by Order</A></li><% End If %>
				<% If pblnShowMyAccount_ProductOrderHistoryLink Then %><li><A href="orderHistoryByProducts.asp">Order History by Product</A></li><% End If %>
				<% If pblnShowMyAccount_DownloadLink Then %><li><A href="priorOrders.asp">Download Products</A></li><% End If %>
				<% If pblnShowMyAccount_PendingNotifications Then %><li><A href="myAccount_PendingNotifications.asp">Pending Product Arrival Notifications</A></li><% End If %>
				</ul>
				</div>
				</td>
			</tr>
			<% End If	'blnShowQuickMenu %>
		</table>
<% 
End Sub 'ShowMyAccountBreadCrumbsTrail

'*************************************************************************************************************************

Sub ShowMenu(blnShowName)

%>
	<br />
	<table class="myAccount" border="0" cellpadding="2" cellspacing="0">
		<tr>
		<td>
  		<% If blnShowName Then %>
		<span class="clsCurrentLocation"><font size="5">Welcome Back <%= Session("custGreeting") %></font></span>
		<% End If %>
		</td>
		</tr>
		<tr>
		<td align="left">
		<ul>
		    <li><A href="MyAccount.asp?Action=View">My Profile</A></li>
		    <li><A href="MyAccount.asp?Action=ChangePwd">Change Password</A></li>
		    <% If hasPriorOrders Then %>
		    <li><A href="orderHistory.asp">Order History by Order</A></li>
		    <li><A href="orderHistoryByProducts.asp">Order History by Products</A></li>
		    <li><A href="neworder.asp">Re-Order from a past order</A></li>
		    <% If hasDownloadableItems Then %><li><A href="priorOrders.asp">Download Products</A></li><% End If	'hasDownloadableItems %>
		    <% End If	'hasPriorOrders %>
		    <% If pblnShowMyAccount_PendingNotifications Then %><li><A href="myAccount_PendingNotifications.asp">Pending Product Arrival Notifications</A></li><% End If %>
		    <li><A href="MyAccount.asp?Action=LogOff">Logout</A></li>
		</ul>
		</td>
		</tr>
	</table>
	<%
	Call ShowGiftCertificates
	If cBuyersClubEnabled Then Call ShowBuyersClubSummaryStatus(visitorLoggedInCustomerID)
	If Len(mstrProblemReportID) = 0 Then
		Call showProblemReportByCustomerID(VisitorLoggedInCustomerID)
	Else
		Call ShowProblemReportStatus(mstrProblemReportID)
	End If
	Call ShowCustomersReviews(visitorLoggedInCustomerID)

End Sub 'ShowMenu

'**********************************************************************************************************

Sub ShowCustomersReviews(byVal lngCustID)

Dim i
Dim paryProductReviews
	'0-contentReferenceID
	'1-contentContent
	'2-contentAuthorName
	'3-contentAuthorEmail
	'4-contentAuthorShowEmail
	'5-contentAuthorRating
	'6-contentDateCreated
Dim pstrAuthor
Dim pstrURL

	'Response.Write "mlngNumReviews: " & mlngNumReviews & "<br />"
	If Len(lngCustID) = 0 Or Not isNumeric(lngCustID) Then Exit Sub
	If loadProductReviewsByCustomer(lngCustID, paryProductReviews) Then
	%>
	<table class="myAccount" border="1" cellpadding="2" cellspacing="0">
	<colgroup>
	  <col valign="top" />
	  <col valign="top" />
	</colgroup>
	<tr class="myAccount"><th colspan="2">Product Review History</th></tr>
	<tr class="myAccount">
	  <th>Product</th>
	  <th>Review</th>
	</tr>
	<%
	For i = 0 To UBound(paryProductReviews, 2)
		Response.Write "<tr>"
		Response.Write "<td>"

		'Now display product
		Set mclsDynamicProducts = New clsDynamicProducts
		With mclsDynamicProducts
			.Connection = cnn
			.DisplayType = 6
			.CurrentProductID = paryProductReviews(9, i)
			.TemplateName = "recentlyViewed_RightColumn.htm"
			.NumColumns = 1
			.NumRows = 1
			.ImageNotPresentURL = "images/NoImage.gif"

			If .LoadDynamicProducts Then
				.DisplayDynamicProducts
			Else
				Response.Write "<em>" & paryProductReviews(10, i) & "</em> is no longer available."
			End If
		End With
		
		Response.Write "</td>"				
		Response.Write "<td align=""left"">"				
				If Abs(paryProductReviews(4, i)) = 1 Then
					pstrAuthor = "<a href='mailTo:" & Trim(paryProductReviews(3, i)) & "'>" & Trim(paryProductReviews(2, i)) & "</a>"
				Else
					pstrAuthor = Trim(paryProductReviews(2, i))
				End If
				
				Call DisplayProductReviewEditLink(paryProductReviews, i)
				Response.Write pstrAuthor & " " & ratingDisplayImage(paryProductReviews(5, i)) & "<br />"
				If Len(Trim(paryProductReviews(1, i))) > 0 Then
					Response.Write "<strong>Comments:</strong> " & Trim(paryProductReviews(1, i)) & "<br />"
				End If
				Response.Write "<strong>Date Reviewed:</strong> " & FormatDateTime(paryProductReviews(6, i)) & "<br />"
				
				If paryProductReviews(7, i) > 0 Or paryProductReviews(8, i) > 0 Then
					Response.Write paryProductReviews(7, i) & " of " & paryProductReviews(7, i) + paryProductReviews(8, i) & " found this review useful.<br />"
				End If
		Response.Write "</td>"				
		Response.Write "</tr>"
	Next 'i
	%>
	</table>
	<%
	End If
	
End Sub	'ShowCustomersReviews

'*************************************************************************************************************************

Sub ShowGiftCertificates()

Dim paryCertificates	'Requires Gift Certificate Module
Dim i
Dim certificateRegistrationLink
Dim plngCertificates
Dim certificateViewLink
Dim pblnDisplay

	If loadCertificatesByCustID(VisitorLoggedInCustomerID, paryCertificates) Then
		If isArray(paryCertificates) Then
		%>
		<table class="myAccount" border="1" cellpadding="2" cellspacing="0">
		  <tr class="myAccount"><th colspan="5">Available Certificates</th></tr>
		  <tr class="myAccount">
			<th>Certificate Number</th>
			<th>Amount Remaining</th>
			<th>Expiration Date</th>
			<th>Redeem</th>
			<th>View</th>
		  </tr>
		<%
		plngCertificates = 0
		For i = 0 To UBound(paryCertificates)
			If Len(paryCertificates(i)(2) & "") = 0 Then
				pblnDisplay = True
			Else
				pblnDisplay = CBool(paryCertificates(i)(2) > Now())
			End If
			pblnDisplay = pblnDisplay And paryCertificates(i)(3) And (CDbl(paryCertificates(i)(1)) > 0)
			
			If pblnDisplay Then
			'If paryCertificates(i)(3) And (paryCertificates(i)(1) > 0) And (paryCertificates(i)(2) > Now() Or isNull(paryCertificates(i)(2))) Then
				plngCertificates = plngCertificates + 1
				
				If isGCRegisteredForUse(paryCertificates(i)(0)) Then
					certificateRegistrationLink = "<a href='myAccount.asp?Action=deleteGCRegistration&Certificate=" & paryCertificates(i)(0) & "' title='Do not use this certificate'>Remove</a>"
				Else
					certificateRegistrationLink = "<a href='myAccount.asp?Certificate=" & paryCertificates(i)(0) & "' title='Use this certificate'>Redeem</a>"
				End If
				certificateViewLink = "viewCertificate.asp?Certificate=" & paryCertificates(i)(0)
				
		%>
		<tr>
		  <td align="center"><%= paryCertificates(i)(0) %></td>
		  <td align="right"><%= FormatCurrency(paryCertificates(i)(1),2) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		  <td align="center"><%
				If isDate(paryCertificates(i)(2)) Then 
					Response.Write FormatDateTime(paryCertificates(i)(2))
				Else
					Response.Write "-"
				End If
			  %></td>
		  <td align="center"><%= certificateRegistrationLink %></td>
		  <td align="center"><a href="<%= certificateViewLink %>">View</a></td>
		</tr>
		<%
			End If
		Next 'i 
		
		If Len(mstrGiftCertificateRegistrationMessage) > 0 Then
		%>
		<tr><td colspan="5"><%= mstrGiftCertificateRegistrationMessage %></td></tr>
		<%
		End If
		
		If plngCertificates = 0 Then
		%>
		<tr><td colspan="5" align="center">No Certificates Available for Redemption</td></tr>
		<%
		End If	'plngCertificates = 0
		%>
		</table>
		<%
		End If
	Else
		'Response.Write "No Gift Certificates are Available for Redemption"
	End If

End Sub	'ShowGiftCertificates

'**********************************************************************************************************

Sub showProblemReportByCustomerID(byVal lngCustID)

Dim pobjRS
Dim pstrSQL

	If Len(lngCustID) = 0 Or Not isNumeric(lngCustID) Then Exit Sub
	
	pstrSQL = "SELECT problemReports.problemReportID, problemReports.dateOpened, problemReports.dateClosed, sfProducts.prodName, problemReports.problemDescription, problemReports.problemResolution" _
			& " FROM problemReports LEFT JOIN sfProducts ON problemReports.ProductID = sfProducts.prodID" _
			& " WHERE problemReports.custId=" & lngCustID _
			& " ORDER BY problemReports.dateOpened DESC"
	'Set	pobjRS = CreateObject("adodb.recordset")
	Set	pobjRS = GetRS(pstrSQL)
	With pobjRS
		If .State = 1 Then
			If Not .EOF Then
				%>
				<table class="myAccount" border="1" cellpadding="2" cellspacing="0">
				<tr><th>View</th><th>Opened On</th><th>Product</th><th>Closed On</th></tr>
				<%
				Do While Not .EOF
					Response.Write "<tr>"
					Response.Write "<td><a href=""myAccount.asp?Action=viewProblemReport&amp;problemReportID=" & .Fields("problemReportID").Value & """>" & .Fields("problemReportID").Value & "</a></td>"				
					Response.Write "<td>" & .Fields("dateOpened").Value & "</td>"				
					Response.Write "<td>" & .Fields("prodName").Value & "</td>"				
					Response.Write "<td>" & .Fields("dateClosed").Value & "&nbsp;</td>"				
					Response.Write "</tr>"
					Response.Write "<tr>"
					Response.Write "<td>&nbsp;</td>"				
					Response.Write "<td colspan=""3"" align=""left"">" & .Fields("problemDescription").Value & "</td>"				
					Response.Write "</tr>"
					.MoveNext
				Loop
				%>
				</table>
				<%
			End If
			.Close
		End If
	End With
	Set pobjRS = Nothing
	
End Sub	'showProblemReportByCustomerID

'**********************************************************************************************************

Sub ShowProblemReportStatus(byVal strProblemReportID)


Dim pobjRS
Dim pstrSQL

	If Len(strProblemReportID) = 0 Or Not isNumeric(strProblemReportID) Then Exit Sub
	
	pstrSQL = "SELECT problemReports.problemReportID, problemReports.dateOpened, problemReports.dateClosed, problemReports.dateModified, problemReports.kbArticleID, sfProducts.prodName, problemReports.problemDescription, problemReports.problemResolution" _
			& " FROM problemReports LEFT JOIN sfProducts ON problemReports.ProductID = sfProducts.prodID" _
			& " WHERE (problemReports.custId=" & VisitorLoggedInCustomerID & " Or problemReports.custId Is Null) And problemReportID = '" & strProblemReportID & "'" _
			& " ORDER BY problemReports.dateOpened DESC"
	Set	pobjRS = GetRS(pstrSQL)
	With pobjRS
		If .EOF Then
			If isLoggedIn Then
				Response.Write "<div style=""border:solid 1pt black;"">We were unable to locate problem report <em>" & strProblemReportID & "</em></div>"
			Else
				Response.Write "<div style=""border:solid 1pt black;"">We were unable to locate problem report <em>" & strProblemReportID & "</em><br />Please log in and and select from your availble options.</div>"
			End If
		Else
			%>
			<table class="myAccount" border="1" cellpadding="2" cellspacing="0">
			<tr>
				<th colspan="2">Problem Report <%= .Fields("problemReportID").Value %></th>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td align="right" valign="top">
				<p align="right">Product:&nbsp;</td>
				<td valign="top"><%= .Fields("prodName").Value %>&nbsp;</td>
			</tr>
			<tr>
				<td align="right" valign="top">Description:&nbsp;</td>
				<td valign="top"><%= .Fields("problemDescription").Value %>&nbsp;</td>
			</tr>
			<tr>
				<td align="right" valign="top">Date Opened:&nbsp;</td>
				<td valign="top"><%= .Fields("dateOpened").Value %>&nbsp;</td>
			</tr>
			<tr>
				<td align="right" valign="top">Resolution:&nbsp;</td>
				<td valign="top"><%= .Fields("problemResolution").Value %>&nbsp;</td>
			</tr>
			<tr>
				<td align="right" valign="top">Date Closed:&nbsp;</td>
				<td valign="top"><%= .Fields("dateClosed").Value %>&nbsp;</td>
			</tr>
			<% If Len(.Fields("kbArticleID").Value & "") > 0 Then %>
			<tr>
				<td>&nbsp;</td>
				<td>This report is covered by Knowledge Base Article <A href="kb/kb.asp?Action=ViewArticle&kbID=<%= .Fields("kbArticleID").Value %>"><%= .Fields("kbArticleID").Value %></A></td>
			</tr>
			<% End If %>
			<tr>
				<td>&nbsp;</td>
				<td>This report was last updated on <%= FormatDateTime(.Fields("dateModified").Value ,1) & " " & FormatDateTime(.Fields("dateModified").Value ,3) %></td>
			</tr>
			</table>
			<%
		End If
		.Close
	End With
	Set pobjRS = Nothing
	
End Sub	'ShowProblemReportStatus %>

