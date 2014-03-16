<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Revision Date: November 26, 2004											*
'*   Version 1.01.001                                                           *
'*                                                                              *
'*   1.00.001 (November 26, 2004)                                               *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Dim mblnShowEditLinks:	mblnShowEditLinks = False

'******************************************************************************************************************************

Sub DisplayAdminMenu()

Dim pblnShowMenu
Dim pstrEditLinks

	pblnShowMenu = isAdminLoggedIn
	
	If Not pblnShowMenu Then pblnShowMenu = isAdminAutoLoginCookieSet
	If Not pblnShowMenu Then pblnShowMenu = isIPAnAdminIP
	
	If pblnShowMenu Then
		
		Response.Write "<div id=""adminMenu""><strong>Admin Menu: </strong>"
		Response.Write "<a href=""MyAccount.asp?Action=ResetVisitor"">Reset Cart</a>"
		Response.Write " | <a href=""ssl/ssAdmin/admin.asp"">Admin Screen</a>"
		Response.Write " | <a href=""ssDebuggingConsole.asp"">Debugging Console</a>"
		
		pstrEditLinks = Request.QueryString("ShowEditLinks")
		If Len(pstrEditLinks) > 0 Then
			mblnShowEditLinks = CBool(pstrEditLinks = "True")
			Session("ShowEditLinks") = mblnShowEditLinks
		Else
			mblnShowEditLinks = Session("ShowEditLinks")
		End If
		
		If mblnShowEditLinks Then
			Response.Write " | <a href=""" & CurrentPage & "?" & Request.QueryString & "&amp;ShowEditLinks=False"">Turn Off Edit Links</a>"
		Else
			Response.Write " | <a href=""" & CurrentPage & "?" & Request.QueryString & "&amp;ShowEditLinks=True"">Turn On Edit Links</a>"
		End If
		Response.Write "</div>"
	End If

End Sub	'DisplayAdminMenu

'******************************************************************************************************************************

Sub DisplayProductEditLink(byVal strProductCode)

Dim paryRawData
Dim paryResults
Dim plngDataCounter
Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	If isAdminLoggedIn And mblnShowEditLinks Then
		pstrSQL = "SELECT sfOrders.orderDate, sfOrderDetails.odrdtSubTotal, sfOrderDetails.odrdtQuantity" _
				& " FROM sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
				& " Where sfOrderDetails.odrdtProductID=?" _
				& " Order By orderDate Desc"
				
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn

			addParameter pobjCmd, "odrdtProductID", adVarChar, strProductCode, 50, 2

			Set pobjRS = .Execute
			If Not pobjRS.EOF Then
				paryRawData = pobjRS.GetRows
				paryResults = PartitionData(paryRawData)
			End If
			Set pobjRS = Nothing
		End With	'pobjCmd
		Set pobjCmd = Nothing

		'images/PENCIL.ICO
		Response.Write "<div id=""adminMenu"" style=""text-align:left"">" & "<a class=""NoLine"" href=""ssl/ssAdmin/sfProductAdmin.asp?Action=ViewProduct&amp;ViewID=" & strProductCode & """ target=""_blank""><img src=""images/PENCIL.ICO"" alt=""Edit"" onmouseover=""showHideElement(document.getElementById('tblProductEdit'));"" onmouseout=""showHideElement(document.getElementById('tblProductEdit'));"" border=0 /></a>"
		Response.Write "<table border=1 cellspacing=0 cellpadding=2 id=tblProductEdit style=""display:none"">"
		Response.Write "<caption>Product Sales History for " & getProductInfo(strProductCode, enProduct_Name) & "</caption>"
		Response.Write "<thead><tr><th>Period</th><th>Units</th><th>Sales</th></tr></thead>"
		If isArray(paryResults) Then
			Response.Write "<tbody id=""productSalesHistory"" style=""display:none"">"
			For plngDataCounter = 1 To UBound(paryResults)
				Response.Write "<tr><td>" & paryResults(plngDataCounter)(0) & "</td><td>" & paryResults(plngDataCounter)(2) & "</td><td>" & FormatCurrency(paryResults(plngDataCounter)(1), 2) & "</td></tr>"
			Next 'plngDataCounter
			plngDataCounter = 0
			Response.Write "</tbody>"
			Response.Write "<tr><th><a href=""Show Sales Data"" onclick=""showHideElement(document.getElementById('productSalesHistory')); return false;"">Sales History</a>: Last Sale on " & paryResults(plngDataCounter)(4) & "</th><th>" & paryResults(plngDataCounter)(2) & "</th><th>" & FormatCurrency(paryResults(plngDataCounter)(1), 2) & "</th></tr>"
		Else
			Response.Write "<tr><th align=""center"" colspan=""3"">No Sales</th></tr>"
		End If
		Response.Write "</table>"
		Response.Write "</div>"
	End If

End Sub	'DisplayProductEditLink

'******************************************************************************************************************************

Sub DisplayProductReviewEditLink(byVal aryReviews, byVal lngIndex)

	If isAdminLoggedIn And mblnShowEditLinks Then
		Response.Write "<div class=""adminCMS""><a class=""adminCMS"" href=""ssl/ssAdmin/ssProductReviewsAdmin.asp?Action=viewItem&amp;ViewID=" & aryReviews(0, lngIndex) & """><img src=""images/PENCIL.ICO"" alt=""Edit"" border=0 /></a></div>"
	End If

End Sub	'DisplayProductEditLink

'******************************************************************************************************************************

Sub DisplayCMSEditLink(byVal lngID, byVal strContent, byVal blnApproved)

	If isAdminLoggedIn Then
		If mblnShowEditLinks Or Not blnApproved Then
			Response.Write "<div class=""adminCMS""><a class=""adminCMS"" href=""ssl/ssAdmin/ssCMS_PageFragmentAdmin.asp?Action=viewItem&amp;ViewID=" & lngID & """><img src=""images/PENCIL.ICO"" alt=""Edit"" border=0 /></a>"
  			If Not blnApproved Then Response.Write "<br /><font color=red><strong>Not Approved for Display!</strong></font>" & strContent
			Response.Write "</div>"
		End If
	End If

End Sub	'DisplayCMSEditLink

'******************************************************************************************************************************

Function PartitionData(byRef aryRawData)

Dim paryResults(7)
Dim pdtNow
Dim plngNumRows
Dim plngRowCounter
Dim plngDataCounter

	pdtNow = Date()
	plngNumRows = UBound(aryRawData, 2)
	paryResults(0) = Array("Summary", 0, 0, "", aryRawData(0,0))
	paryResults(1) = Array("Today", 0, 0, DateAdd("d", -1, pdtNow), pdtNow)
	paryResults(2) = Array("This Week", 0, 0, DateAdd("d", -7, pdtNow), pdtNow)
	paryResults(3) = Array("Last Week", 0, 0, DateAdd("d", -14, pdtNow), DateAdd("d", -7, pdtNow))
	paryResults(4) = Array("This Month", 0, 0, DateAdd("d", -30, pdtNow), pdtNow)
	paryResults(5) = Array("Last Month", 0, 0, DateAdd("d", -60, pdtNow), DateAdd("d", -30, pdtNow))
	paryResults(6) = Array("This Year", 0, 0, DateAdd("d", -365, pdtNow), pdtNow)
	paryResults(7) = Array("Last Year", 0, 0, DateAdd("d", -730, pdtNow), DateAdd("d", -365, pdtNow))
	For plngRowCounter = 0 To plngNumRows
		For plngDataCounter = 0 To UBound(paryResults)
			If Len(paryResults(plngDataCounter)(3)) = 0 Then
				paryResults(plngDataCounter)(1) = paryResults(plngDataCounter)(1) + aryRawData(1, plngRowCounter)
				paryResults(plngDataCounter)(2) = paryResults(plngDataCounter)(2) + aryRawData(2, plngRowCounter)
			ElseIf aryRawData(0, plngRowCounter) >= paryResults(plngDataCounter)(3) And aryRawData(0, plngRowCounter) <= paryResults(plngDataCounter)(4) Then
				paryResults(plngDataCounter)(1) = paryResults(plngDataCounter)(1) + aryRawData(1, plngRowCounter)
				paryResults(plngDataCounter)(2) = paryResults(plngDataCounter)(2) + aryRawData(2, plngRowCounter)
			End If
		Next 'plngDataCounter
	Next 'plngRowCounter
	
	PartitionData = paryResults

End Function	'PartitionData
%>