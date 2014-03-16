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
'Items to look for

'AVS Match
'CCV/CVV
'Different billing/shipping
'Address Line 2 used
'IP Ranges
'Specific countries
'Email domains

'**********************************************************
'*	Page Level variables
'**********************************************************

Const clngFraudThreshhold_ImmediateDownload = 5
Dim mlngFraudScore:	mlngFraudScore = 0

'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'*	Begin Page Code
'**********************************************************

Sub evaluateFraudPotentialScore(byRef lngOrderStatus, byRef lngInternalOrderStatus)

	If ssDebug_FraudScore Then Response.Write "<fieldset><legend>Fraud Scoring</legend>"
	If ssDebug_FraudScore Then Response.Write "Score (AVS/CCV): " & mlngFraudScore & "<br />"

	'AVS/CVV is already scored from calls in processor.asp
	Call scoreAddress(mclsCustomer, mclsCustomerShipAddress)
	If ssDebug_FraudScore Then Response.Write "Score (Address): " & mlngFraudScore & "<br />"
	Call scoreEmail(mclsCustomer.custEmail)
	If ssDebug_FraudScore Then Response.Write "Score (Email): " & mlngFraudScore & "<br />"
	mlngFraudScore = mlngFraudScore + getCountryFraudScore(mclsCustomer.custCountry)
	If ssDebug_FraudScore Then Response.Write "Score (Country): " & mlngFraudScore & "<br />"
	Call scorePriorCustomer
	If ssDebug_FraudScore Then Response.Write "Score (Prior Customer): " & mlngFraudScore & "<br />"
	
	'lngOrderStatus	- corresponds to maryOrderStatuses item defined in ssl/ssAdmin/ssOrderAdmin_common
	'lngInternalOrderStatus	- corresponds to maryInternalOrderStatuses item defined in ssl/ssAdmin/ssOrderAdmin_common
	
	Select Case mlngFraudScore
		Case mlngFraudScore = 0
			lngOrderStatus = 5
			lngInternalOrderStatus = 2
		Case mlngFraudScore <= 1
			lngOrderStatus = 1
			lngInternalOrderStatus = 3
		Case mlngFraudScore <= 5
			lngOrderStatus = 1
			lngInternalOrderStatus = 3
		Case Else
			lngOrderStatus = 1
			lngInternalOrderStatus = 3
	End Select
	
	If ssDebug_FraudScore Then
		Response.Write "<hr />Final Fraud Score: " & mlngFraudScore & "<br />"
		Response.Write "External Order Status: " & lngOrderStatus & "<br />"
		Response.Write "Internal Order Status: " & lngInternalOrderStatus & "<br />"
		Response.Write "</fieldset>"
	End If

End Sub	'evaluateFraudPotentialScore

'**********************************************************

Sub scorePriorCustomer()

Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	'look for repeat orders over 48 hours which have shipped
	pstrSQL = "SELECT sfOrders.orderDate" _
			& " FROM ssOrderManager INNER JOIN sfOrders ON ssOrderManager.ssorderID = sfOrders.orderID" _
			& " WHERE ((ssOrderManager.ssDateOrderShipped Is Not Null) AND (sfOrders.orderCustId=?) AND (sfOrders.orderDate<?))"

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, visitorLoggedInCustomerID)
		.Parameters.Append .CreateParameter("orderDate", adDBTimeStamp, adParamInput, 16, DateAdd("h", -48, Now()))
		Set pobjRS = .Execute
		If pobjRS.EOF Then
			'New customer
			'Don't do anything
		Else
			'Returning customer so reset fraud potential back to zero
			mlngFraudScore = 0
		End If
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Sub	'scorePriorCustomer

'**********************************************************

Sub scoreEmail(byVal strEmail)

Dim plngIncrementalValue
Dim plngPos
Dim pobjCmd
Dim pobjRS
Dim pstrEmailToScore
Dim pstrEmailDomain

	plngIncrementalValue = 0
	plngPos = InStrRev(strEmail, "@")
	If plngPos > 0 Then
		pstrEmailDomain = Right(strEmail, Len(strEmail) - plngPos + 1)
	End If
	
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "SELECT fraudEmail, fraudScore FROM fraudEmails WHERE fraudEmail Like ? Order by fraudEmail"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("fraudEmail", adVarChar, adParamInput, 255, pstrEmailDomain)
		Set pobjRS = .Execute
		Do While Not pobjRS.EOF
			pstrEmailToScore = LCase(Trim(pobjRS.Fields("fraudEmail").Value & ""))
			If pstrEmailToScore = strEmail Then
				'Exact Match
				plngIncrementalValue = pobjRS.Fields("fraudScore").Value
				Exit Do
			ElseIf Left(pstrEmailToScore, 1) = "@" Then
				'Looking for a domain Match
				plngIncrementalValue = pobjRS.Fields("fraudScore").Value
				'Do not exit do since an exact match may appear later
			End If
			pobjRS.MoveNext
		Loop
		mlngFraudScore = mlngFraudScore + plngIncrementalValue
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Sub	'scoreEmail

'**********************************************************

Sub scoreAddress(byRef objclsCustomer, byRef objclsCustomerShipAddress)

Dim pstrBillingAddress
Dim pstrShippingAddress

	With objclsCustomer
		pstrBillingAddress = .custCompany _
						   & .custFirstName _
						   & .custMiddleInitial _
						   & .custLastName _
						   & .custAddr1 _
						   & .custAddr2 _
						   & .custCity _
						   & .custState _
						   & .custZip _
						   & .custCountry
	End With
	
	With objclsCustomerShipAddress
		pstrShippingAddress= .Company _
						   & .FirstName _
						   & .MiddleInitial _
						   & .LastName _
						   & .Addr1 _
						   & .Addr2 _
						   & .City _
						   & .State _
						   & .Zip _
						   & .Country
						   
		'Check if line 2 is used
		If Len(.Addr2) > 0 Then mlngFraudScore = mlngFraudScore + 1
	End With
	
	If CBool(LCase(pstrBillingAddress) = LCase(pstrBillingAddress)) Then
		mlngFraudScore = mlngFraudScore + 0
	Else
		mlngFraudScore = mlngFraudScore + 1
	End If

End Sub	'scoreAddress

'**********************************************************

Sub scoreAVS(byVal strAVS)
	Select Case strAVS
		Case "A"	'"Address (Street)matches, ZIP does not. (Code A)"
			mlngFraudScore = mlngFraudScore + 1
		Case "D"	'"Street address and Postal Code match (International Issuer). (Code D)"
			mlngFraudScore = mlngFraudScore + 1
		Case "E"	'"Ineligible transaction. (Code E)"
			mlngFraudScore = mlngFraudScore + 1
		Case "G"	'"Service not supported by issuer (International). (Code G)"
			mlngFraudScore = mlngFraudScore + 1
		Case "N"	'"Neither address nor ZIP matches. (Code N)"
			mlngFraudScore = mlngFraudScore + 1
		Case "R"	'"Retry (system unavailable or timed out). (Code R)"
			mlngFraudScore = mlngFraudScore + 1
		Case "S"	'"Card type not supported. (Code S)"
			mlngFraudScore = mlngFraudScore + 1
		Case "U"	'"Address information unavailable. (Code U)"
			mlngFraudScore = mlngFraudScore + 1
		Case "W"	'"9 digit zip match, address does not. (Code W)"
			mlngFraudScore = mlngFraudScore + 1
		Case "X"	'"Exact match (9 digit zip and address). (Code X)"
			mlngFraudScore = mlngFraudScore + 0
		Case "Y"	'"Address and 5 digit zip match. (Code Y)"
			mlngFraudScore = mlngFraudScore + 0
		Case "Z"	'"5 digit zip matches, address does not. (Code Z)"
			mlngFraudScore = mlngFraudScore + 1
		Case Else	'"Unknown AVS Code."
			mlngFraudScore = mlngFraudScore + 1
	End Select	
End Sub	'scoreAVS

'**********************************************************

Sub scoreCVV(byVal strCVV)
	Select Case strAVS
		Case "CVV2 MATCH"			'SecurePay Specific?
			mlngFraudScore = mlngFraudScore + 0
		Case "CVV2 NOT AVAILABLE"	'SecurePay Specific?
			mlngFraudScore = mlngFraudScore + 1
		Case "CVV2 NOMATCH"			'SecurePay Specific?
			mlngFraudScore = mlngFraudScore + 1
		Case Else	'"Unknown"
			mlngFraudScore = mlngFraudScore + 1
	End Select	
End Sub	'scoreAVS

%>








