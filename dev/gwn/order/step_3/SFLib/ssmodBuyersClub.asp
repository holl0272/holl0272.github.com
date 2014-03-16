<%
'********************************************************************************
'*   Buyers' Club Module														*
'*   Release Version:	2.00.001 												*
'*   Release Date:		October 1, 2006											*
'*   Revision Date:		October 1, 2006											*
'*																				*
'*   Release 2.00.001 (October 1, 2006)											*
'*	   - Initial Release														*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.				*
'********************************************************************************

Dim mdblBuyersClubPointsAvailable
Dim mstrMessage_BuyersClub

Dim cBuyersClubEarningsMultiple:	cBuyersClubEarningsMultiple = CDbl(getConfigurationSettingFromCache("BuyersClubEarningsMultiple", 1))
Dim cBuyersClubRedemptionMultiple:	cBuyersClubRedemptionMultiple = CDbl(getConfigurationSettingFromCache("BuyersClubRedemptionMultiple", 1))
Dim cBuyersClubMinimumRedemption:	cBuyersClubMinimumRedemption = CDbl(getConfigurationSettingFromCache("BuyersClubMinimumRedemption", 25))
Dim cBuyersClubCertificateMultiple:	cBuyersClubCertificateMultiple = CDbl(getConfigurationSettingFromCache("BuyersClubCertificateMultiple", 0.5))
Dim cBuyersClubEnabled:				cBuyersClubEnabled = ConvertToBoolean(getConfigurationSettingFromCache("BuyersClubEnabled", False), False)

'***********************************************************************************************

Function BuyersClubPointBalance(byVal lngCustID)

Dim i
Dim paryResults
Dim pdblBalance
Dim pdblPointSubTotal

	pdblBalance = 0

	mdblBuyersClubPointsAvailable = Session("BuyersClubPointsAvailable")
	If Len(mdblBuyersClubPointsAvailable) > 0 And isNumeric(mdblBuyersClubPointsAvailable) Then
		pdblBalance = mdblBuyersClubPointsAvailable
	Else
		Call LoadBuyersClubHistory(paryResults, lngCustID)
		If isArray(paryResults) Then
			For i = 0 To UBound(paryResults)
				pdblPointSubTotal = Trim(paryResults(i)(2))
				If Len(pdblPointSubTotal) = 0 Or Not isNumeric(pdblPointSubTotal) Then pdblPointSubTotal = 0
				pdblBalance = pdblBalance + pdblPointSubTotal
			Next 'i
		End If	'isArray(paryResults)
	End If
	
	BuyersClubPointBalance = pdblBalance

End Function	'BuyersClubPointBalance

'***********************************************************************************************

Sub SaveBuyersClubOrder(byRef objcnn, byVal lngOrderID)

Dim i
Dim paryResults
Dim pdblProductSubTotal
Dim pdblPointSubTotal
Dim pdblPointsToIssue
Dim pobjCMD
Dim pobjRS
Dim pstrSQL

	pstrSQL = "SELECT sfOrderDetails.odrdtID, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, sfProducts.buyersClubPointValue, sfProducts.buyersClubIsPercentage" _
			& " FROM sfOrderDetails INNER JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
			& " WHERE sfOrderDetails.odrdtOrderId=?"

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("orderID", adInteger, adParamInput, 4, lngOrderID)
		Set pobjRS = .Execute
		If Not pobjRS.EOF Then
			paryResults = pobjRS.getRows
		End If
		pobjRS.Close
		Set pobjRS = Nothing
		
		.Parameters.Delete "orderID"
		
		If isArray(paryResults) Then
			pstrSQL = "Update sfOrderDetails Set buyersClubPointsIssued=? Where odrdtID=?"
			.Commandtext = pstrSQL
			.Parameters.Append .CreateParameter("buyersClubPointsIssued", adDouble, adParamInput, 8, 0)
			.Parameters.Append .CreateParameter("odrdtID", adInteger, adParamInput, 4, lngOrderID)
			For i = 0 To UBound(paryResults, 2)
				pdblPointSubTotal = Trim(paryResults(3, i))
				If Len(pdblPointSubTotal) = 0 Or Not isNumeric(pdblPointSubTotal) Then pdblPointSubTotal = 0

				If paryResults(4, i) = 1 Then
					pdblProductSubTotal = Trim(paryResults(2, i))
					If Len(pdblProductSubTotal) = 0 Or Not isNumeric(pdblProductSubTotal) Then pdblProductSubTotal = 0
					
					pdblPointsToIssue = pdblPointSubTotal * pdblProductSubTotal
				Else
					pdblProductSubTotal = Trim(paryResults(1, i))
					If Len(pdblProductSubTotal) = 0 Or Not isNumeric(pdblProductSubTotal) Then pdblProductSubTotal = 0
					
					pdblPointsToIssue = pdblPointSubTotal * paryResults(1, i)
				End If
				.Parameters("buyersClubPointsIssued").Value = pdblPointsToIssue * cBuyersClubEarningsMultiple
				.Parameters("odrdtID").Value = paryResults(0, i)
				.Execute , , 128
			Next 'i
		End If
		
	End With
	Set pobjCMD = Nothing
	
	Session.Contents.Remove("BuyersClubPointsAvailable")

End Sub	'SaveBuyersClubOrder

'***********************************************************************************************

Sub SaveBuyersClubRedemption(byVal lngCustID, byVal dblPointsToRedeem, byVal lngCertificateID)

Dim pobjCmd
Dim pobjRS
Dim pstrSQL
'SELECT @@IDENTITY

	pstrSQL = "Insert Into ssBuyersClubRedemptions (ssBuyersClubRedemptionCustID, ssBuyersClubRedemptionPoints, ssBuyersClubRedemptionCertificateID, ssBuyersClubRedemptionDate)" _
			& " Values (?, ?, ?, ?)"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("ssBuyersClubRedemptionCustID", adInteger, adParamInput, 4, lngCustID)
		.Parameters.Append .CreateParameter("ssBuyersClubRedemptionPoints", adInteger, adParamInput, 4, -1 * dblPointsToRedeem)
		.Parameters.Append .CreateParameter("ssBuyersClubRedemptionCertificateID", adInteger, adParamInput, 4, dblPointsToRedeem)
		.Parameters.Append .CreateParameter("ssBuyersClubRedemptionDate", adDBTimeStamp, adParamInput, 16, Now())
		.Execute , , 128
		
	End With
	Set pobjCMD = Nothing
	
End Sub	'SaveBuyersClubRedemption

'***********************************************************************************************

Sub ShowBuyersClubHistory(byVal lngCustID, byVal blnComplete)

Dim i
Dim paryResults
Dim pdblBalance
Dim pdblPointSubTotal

	pdblBalance = 0

	Call LoadBuyersClubHistory(paryResults, lngCustID)
	If isArray(paryResults) Then
		%>
		<table class="myAccount" border="1" cellpadding="2" cellspacing="0">
		<tr class="myAccount"><th colspan="4">Buyer's Club Point Earning History</th></tr>
		<tr class="myAccount">
			<th>Order Number</th>
			<th>Order Date</th>
			<th>Points Earned</th>
			<th>Balance</th>
		</tr>
		<%
		For i = 0 To UBound(paryResults)
			pdblPointSubTotal = Trim(paryResults(i)(2))
			If Len(pdblPointSubTotal) = 0 Or Not isNumeric(pdblPointSubTotal) Then pdblPointSubTotal = 0
			pdblBalance = pdblBalance + pdblPointSubTotal
		%>
		<tr>
		<td align="center">
		<% If paryResults(i)(3) = 1  Then %>
		<a href="OrderHistory.asp?OrderID=<%= paryResults(i)(0) %>" title="View order details">Order <%= paryResults(i)(0) %></a></td>
		<% Else %>
		<a href="myAccount.asp?Action=ViewBuyersClubRempdtion&amp;RedemptionID=<%= paryResults(i)(0) %>" title="View redemption details"></a>Redemption</td>
		<% End If %>
		<td align="right"><%= FormatDateTime(paryResults(i)(1), 2) %></td>
		<td align="right"><%= pdblPointSubTotal %></td>
		<td align="right"><%= pdblBalance %></td>
		</tr>
		<%
		Next 'i
		%>
		</table>
		<%
	End If	'isArray(paryResults)

End Sub	'ShowBuyersClubHistory

'***********************************************************************************************

Sub ShowBuyersClubSummaryStatus(byVal lngCustID)

Dim i
Dim pdblBalance

	pdblBalance = BuyersClubPointBalance(lngCustID)
	
	'Response.Write "Buyer's Club Points Available: <strong><a href='myAccount.asp?Action=ShowBuyersClubDetail' title='View point history'>" & BuyersClubPointBalance(lngCustID) & "</a>"
	Response.Write "<table class=""myAccount"" border=""1"" cellpadding=""2"" cellspacing=""0"">"
	Response.Write "<tr class=""myAccount""><th colspan=""2"">Buyer's Club Point Summary</th></tr>"
	If BuyersClubPointBalance(lngCustID) = 0 Then
		Response.Write "<tr><td>Buyer's Club Points Available:</td><td><strong><a href='myAccount.asp?Action=ShowBuyersClubDetail'>" & FormatCurrency(BuyersClubPointBalance(lngCustID)) & "</a></td></tr>"
	Else
		Response.Write "<tr><td>Buyer's Club Points Available:</td><td><strong><a href='myAccount.asp?Action=ShowBuyersClubDetail' title='View point history'>" & FormatCurrency(BuyersClubPointBalance(lngCustID)) & "</a></td></tr>"
	End If
	If pdblBalance = 0 Then
		'do nothing
	ElseIf pdblBalance < cBuyersClubMinimumRedemption Then
		'Response.Write " Minimum redemption amount is " & cBuyersClubMinimumRedemption & " points."
		Response.Write "<tr><td>Minimum redemption amount:</td><td>" & FormatCurrency(cBuyersClubMinimumRedemption) & "</td></tr>"
	Else
		Response.Write "<tr><td>&nbsp;</td><td><a href=myAccount.asp?Action=ShowBuyersClubRedemptionOptions title='Redeem your points'>Redeem</a></td></tr>"
	End If 
	Response.Write "</table>"

End Sub	'ShowBuyersClubSummaryStatus

'***********************************************************************************************

Sub ShowBuyersClubRedemptionOptions(byVal lngCustID)

Dim i
Dim pdblBalance
Dim plngNumOptions

	If Len(lngCustID) = 0 Or Not isNumeric(lngCustID) Then Exit Sub
	
	pdblBalance = BuyersClubPointBalance(lngCustID)
	If pdblBalance = 0 Then
		Response.Write "You have no points to redeem."
	ElseIf pdblBalance < cBuyersClubMinimumRedemption Then
		Response.Write "Minimum redemption amount is " & cBuyersClubMinimumRedemption & " points."
	Else
		If Len(mstrMessage_BuyersClub) > 0 Then
			Response.Write mstrMessage_BuyersClub
		End If
		
		Response.Write getPageFragmentByKey("BuyersClubRedemptionInstructions")
		Response.Write "<form action='myAccount.asp' method=post onsubmit='return true;'>"
		Response.Write "<input type=hidden name=Action id=Action value=BuyersClubCreateRedemption>"
		Response.Write "<table class=""myAccount"" border=""1"" cellpadding=""2"" cellspacing=""0"">"
		If cBuyersClubCertificateMultiple = 1 Then
			Response.Write "Redeem <input type=text name=pointsToRedeem id=pointsToRedeem size=6 value=" & pdblBalance & "> points X " & FormatCurrency(cBuyersClubRedemptionMultiple) & " per point = " & FormatCurrency(cBuyersClubRedemptionMultiple * pdblBalance) & "."
		Else
			plngNumOptions = Int((pdblBalance - cBuyersClubMinimumRedemption) / cBuyersClubCertificateMultiple)
			'Response.Write "<tr><th>Option</th><th>Points to Redeem</th><th>Rate</th><th>Certificate Value</th></tr>"
			Response.Write "<tr class=""myAccount""><th>Certificate Value</th></tr>"
			For i = 0 To plngNumOptions
				'Response.Write "<tr><td align=center><input type=radio name=pointsToRedeem id=pointsToRedeem" & i & " value=" & i * cBuyersClubCertificateMultiple & "></td><td align=center>" & i * cBuyersClubCertificateMultiple & "</td><td align=center>" & FormatCurrency(cBuyersClubRedemptionMultiple) & " per point</td><td align=center>" & FormatCurrency(i * cBuyersClubCertificateMultiple * cBuyersClubRedemptionMultiple) & ".</td></tr>"
				Response.Write "<tr><td><input type=radio name=pointsToRedeem id=pointsToRedeem" & i & " value=" & (cBuyersClubMinimumRedemption + i * cBuyersClubCertificateMultiple) & ">&nbsp;<label for=pointsToRedeem" & i & ">" & FormatCurrency((cBuyersClubMinimumRedemption + i * cBuyersClubCertificateMultiple) * cBuyersClubRedemptionMultiple) & "</label></td></tr>"
			Next 'i
			'Response.Write "<tr><td>&nbsp;</td><td colspan=3><input type=image name=btnSubmit id=btnSubmit src=" & Chr(34) & C_BTN18 & Chr(34) & " value=Submit></td></tr>"
			Response.Write "<tr><td colspan=1><input type=image name=btnSubmit id=btnSubmit src=" & Chr(34) & C_BTN18 & Chr(34) & " value=Submit></td></tr>"
		End If
		Response.Write "</table>"
		Response.Write "</form>"
	End If 

End Sub	'ShowBuyersClubRedemptionOptions

'***********************************************************************************************

Function LoadBuyersClubEarnings(byRef aryEarnings, byVal lngCustID)

Dim pobjCMD
Dim pobjRS
Dim pstrSQL

	If isNumeric(lngCustID) And Len(lngCustID) > 0 Then
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, Sum(sfOrderDetails.buyersClubPointsIssued) AS SumOfbuyersClubPointsIssued" _
				& " FROM ssOrderManager INNER JOIN (sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON ssOrderManager.ssorderID = sfOrders.orderID" _
				& " GROUP BY sfOrders.orderID, sfOrders.orderDate, sfOrders.orderCustId, ssOrderManager.ssDatePaymentReceived" _
				& " HAVING ((Sum(sfOrderDetails.buyersClubPointsIssued)>0) AND (ssOrderManager.ssDatePaymentReceived Is Not Null) AND (sfOrders.orderCustId=?))"

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("orderID", adInteger, adParamInput, 4, lngCustID)
			Set pobjRS = .Execute
			If Not pobjRS.EOF Then
				aryEarnings = pobjRS.getRows
			End If
			pobjRS.Close
			Set pobjRS = Nothing
			
			LoadBuyersClubEarnings = isArray(aryEarnings)
			
		End With
		Set pobjCMD = Nothing
	End If

End Function	'LoadBuyersClubEarnings

'***********************************************************************************************

Function LoadBuyersClubHistory(byRef aryHistory, byVal lngCustID)

Dim paryEarnings
Dim paryRedemptions
Dim paryTemp
Dim pblnUseEarning
Dim pblnUseRedemption
Dim i
Dim plngCount
Dim plngEarningsCounter
Dim plngRedemptionsCounter
Dim plngPos
Dim pdt

	plngCount = -1
	If LoadBuyersClubEarnings(paryEarnings, lngCustID) Then plngCount = plngCount + UBound(paryEarnings, 2) + 1
	If LoadBuyersClubRedemptions(paryRedemptions, lngCustID) Then plngCount = plngCount + UBound(paryRedemptions, 2) + 1

	If plngCount <> -1 Then
		ReDim aryHistory(plngCount)
		plngEarningsCounter = 0
		plngRedemptionsCounter = 0
		
		If isArray(paryEarnings) And isArray(paryRedemptions) Then
			For i = 0 To plngCount
				pblnUseEarning = False
				pblnUseRedemption = False
				
				If plngRedemptionsCounter > UBound(paryRedemptions, 2) Then
					pblnUseEarning = True
				ElseIf plngEarningsCounter > UBound(paryEarnings, 2) Then
					pblnUseRedemption = True
				ElseIf paryEarnings(1, plngEarningsCounter) > paryRedemptions(1, plngRedemptionsCounter) Then
					pblnUseRedemption = True
				Else
					pblnUseEarning = True
				End If
				
				If pblnUseEarning Then
					aryHistory(i) = Array(paryEarnings(0,plngEarningsCounter), paryEarnings(1,plngEarningsCounter), paryEarnings(2,plngEarningsCounter), 1)
					plngEarningsCounter = plngEarningsCounter + 1
				ElseIf pblnUseRedemption Then
					aryHistory(i) = Array(paryRedemptions(0,plngRedemptionsCounter), paryRedemptions(1,plngRedemptionsCounter), paryRedemptions(2,plngRedemptionsCounter), 0)
					plngRedemptionsCounter = plngRedemptionsCounter + 1
				End If
			
			Next 'i
		ElseIf Not isArray(paryRedemptions) Then
			For i = 0 To UBound(paryEarnings, 2)
				aryHistory(i) = Array(paryEarnings(0,i), paryEarnings(1,i), paryEarnings(2,i), 1)
			Next 'i
		ElseIf Not isArray(paryEarnings) Then
			For i = 0 To UBound(paryRedemptions, 2)
				aryHistory(i) = Array(paryRedemptions(0,i), paryRedemptions(1,i), paryRedemptions(2,i), 0)
			Next 'i
		End If	'Not isArray(paryRedemptions)
		
		LoadBuyersClubHistory = True
	Else
		LoadBuyersClubHistory = False
	End If	'plngCount <> -1

End Function	'LoadBuyersClubHistory

'***********************************************************************************************

Function LoadBuyersClubRedemptions(byRef aryRedemptions, byVal lngCustID)

Dim pobjCMD
Dim pobjRS
Dim pstrSQL

	If isNumeric(lngCustID) And Len(lngCustID) > 0 Then
		pstrSQL = "SELECT ssBuyersClubRedemptions.ssBuyersClubRedemptionID, ssBuyersClubRedemptions.ssBuyersClubRedemptionDate, ssBuyersClubRedemptions.ssBuyersClubRedemptionPoints" _
				& " FROM ssBuyersClubRedemptions" _
				& " WHERE ssBuyersClubRedemptions.ssBuyersClubRedemptionCustID=?"

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("orderID", adInteger, adParamInput, 4, lngCustID)
			Set pobjRS = .Execute
			If Not pobjRS.EOF Then
				aryRedemptions = pobjRS.getRows
			End If
			pobjRS.Close
			Set pobjRS = Nothing
			
			LoadBuyersClubRedemptions = isArray(aryRedemptions)
			
		End With
		Set pobjCMD = Nothing
	End If

End Function	'LoadBuyersClubRedemptions

'***********************************************************************************************

Function LoadBuyersClubRedemptionByID(byRef aryRedemptions, byVal lngCustID, byVal lngRedemptionID)

Dim pobjCMD
Dim pobjRS
Dim pstrSQL

	If isNumeric(lngCustID) And Len(lngCustID) > 0 And isNumeric(lngRedemptionID) And Len(lngRedemptionID) > 0 Then
		pstrSQL = "SELECT ssBuyersClubRedemptions.ssBuyersClubRedemptionID, ssBuyersClubRedemptions.ssBuyersClubRedemptionDate, ssBuyersClubRedemptions.ssBuyersClubRedemptionPoints" _
				& " FROM ssBuyersClubRedemptions" _
				& " WHERE ssBuyersClubRedemptions.ssBuyersClubRedemptionCustID=? And ssBuyersClubRedemptions.ssBuyersClubRedemptionID=?"

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("CustID", adInteger, adParamInput, 4, lngCustID)
			.Parameters.Append .CreateParameter("RedemptionID", adInteger, adParamInput, 4, lngRedemptionID)
			Set pobjRS = .Execute
			If Not pobjRS.EOF Then
				aryRedemptions = pobjRS.getRows
			End If
			pobjRS.Close
			Set pobjRS = Nothing
			
			LoadBuyersClubRedemptionByID = isArray(aryRedemptions)
			
		End With
		Set pobjCMD = Nothing
	End If

End Function	'LoadBuyersClubRedemptionByID

'************************************************************************************************

Function ssBuyersClub_CreateRedemption(byVal lngCustID)

Dim pclsssGiftCertificate
Dim pdblPointsToRedeem
Dim pdblRedemptionAmount
Dim pstrCertificateCode
Dim plngCertificateID

	pdblPointsToRedeem = Request.Form("pointsToRedeem")
	'pdblPointsToRedeem = 6

	'validate the request
	If Len(pdblPointsToRedeem) = 0 Or Not isNumeric(pdblPointsToRedeem) Then
		mstrMessage_BuyersClub = "<em>" & pdblPointsToRedeem & "</em> is not a valid redemption request."
		ssBuyersClub_CreateRedemption =  False
		Exit Function
	ElseIf CDbl(pdblPointsToRedeem) < cBuyersClubMinimumRedemption Then
		mstrMessage_BuyersClub = "<em>" & pdblPointsToRedeem & "</em> is not a valid redemption request. The minimum redemption amount is <strong>" & cBuyersClubMinimumRedemption & "</strong>."
		ssBuyersClub_CreateRedemption =  False
		Exit Function
	ElseIf CDbl(pdblPointsToRedeem) > BuyersClubPointBalance(lngCustID) Then
		mstrMessage_BuyersClub = "Your redemption request of <em>" & pdblPointsToRedeem & "</em> exceeds your available balance of <strong>" & BuyersClubPointBalance(lngCustID) & "</strong>."
		ssBuyersClub_CreateRedemption =  False
		Exit Function
	End If
	
	pdblRedemptionAmount = pdblPointsToRedeem * cBuyersClubRedemptionMultiple
	
	'Create the certificate
	Set pclsssGiftCertificate = New clsssGiftCertificate
	With pclsssGiftCertificate
		.Connection = cnn
		.ssGCCustomerID = lngCustID
		.ssGCExpiresOn = DateAdd("m", 6, Date())
		.ssGCSingleUse = False
		
		'The next two aren't currently implemented
		.ssGCElectronic = False
		.ssGCFreeText = ""

		'If it is a self-issue "purchase card" then set the customer ID
		
		.createCertificate_New ""
		
		pstrCertificateCode = .ssGCCode
		plngCertificateID = .ssGCID

		Call pclsssGiftCertificate.CreateRedemption(True, "", pstrCertificateCode, enStoreCredit, pdblRedemptionAmount, True, "", "", "")
		Call SendCertificateEmail(pclsssGiftCertificate, enStoreCredit)
		
	End With
	Set pclsssGiftCertificate = Nothing
	
	Call SaveBuyersClubRedemption(lngCustID, pdblPointsToRedeem, plngCertificateID)
	
	mstrMessage_BuyersClub = "Certificate " & pstrCertificateCode & " for " & FormatCurrency(pdblRedemptionAmount) & " is available for use."
	ssBuyersClub_CreateRedemption = True

End Function	'ssBuyersClub_CreateRedemption

'***********************************************************************************************

%>
