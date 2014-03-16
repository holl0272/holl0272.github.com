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

'Note: recordPageView is called in db.conn.open.asp immediately after db connection opened
'		cnn.open
'		Call recordPageView
'Note: recordSearchResults is called on search_results
'		Call closeObj(rsSearch)
'		Call recordSearchResults(txtsearchParamType, txtsearchParamTxt, iSearchRecordCount)

'**********************************************************
'*	Page Level variables
'**********************************************************

Const visitorLastSearchMaxLength = 255

Dim maryVisitorPreferences
Dim mlngVisitorID
Dim visitorPageViewID

'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'*	Begin Page Code
'**********************************************************

If Len(CStr(visitorPageViewID)) = 0 Then visitorPageViewID = 0	'check the length since the order of the include files can affect this

'**********************************************************
'**********************************************************

Sub displayVisitorPreferences()

	Response.Write "<fieldset><legend>Visitor Preferences for visitor <em>" & getvisitorPreference("visitorID") & "</em></legend>"
	Response.Write "visitorSessionID: " & getvisitorPreference("visitorSessionID") & "<br />"
	Response.Write "visitorCustomerID: " & getvisitorPreference("visitorCustomerID") & "<br />"
	Response.Write "visitorLastVisited: " & getvisitorPreference("visitorLastVisited") & "<br />"
	Response.Write "vistorDiscountCodes: " & getvisitorPreference("vistorDiscountCodes") & "<br />"
	Response.Write "visitorSelectedFreeProducts: " & getvisitorPreference("visitorSelectedFreeProducts") & "<br />"
	Response.Write "visitorCertificateCodes: " & getvisitorPreference("visitorCertificateCodes") & "<br />"
	Response.Write "visitorLastSearch: " & getvisitorPreference("visitorLastSearch") & "<br />"
	Response.Write "visitorRecentlyViewedProducts: " & getvisitorPreference("visitorRecentlyViewedProducts") & "<br />"
	Response.Write "visitorCity: " & getvisitorPreference("visitorCity") & "<br />"
	Response.Write "visitorState: " & getvisitorPreference("visitorState") & "<br />"
	Response.Write "visitorZIP: " & getvisitorPreference("visitorZIP") & "<br />"
	Response.Write "visitorCountry: " & getvisitorPreference("visitorCountry") & "<br />"
	Response.Write "visitorPreferredCurrency: " & getvisitorPreference("visitorPreferredCurrency") & "<br />"
	Response.Write "visitorPreferredShippingCode: " & getvisitorPreference("visitorPreferredShippingCode") & "<br />"
	Response.Write "visitorEstimatedShipping: " & getvisitorPreference("visitorEstimatedShipping") & "<br />"
	Response.Write "visitor_REFERER: " & getvisitorPreference("visitor_REFERER") & "<br />"
	Response.Write "vistor_HTTP_REFERER: " & getvisitorPreference("vistor_HTTP_REFERER") & "<br />"
	Response.Write "visitor_REMOTE_ADDR: " & getvisitorPreference("visitor_REMOTE_ADDR") & "<br />"
	Response.Write "visitorLoggedInCustomerID: " & getvisitorPreference("visitorLoggedInCustomerID") & "<br />"
	Response.Write "visitorShipAddressID: " & getvisitorPreference("visitorShipAddressID") & "<br />"
	Response.Write "visitorOrderID: " & getvisitorPreference("visitorOrderID") & "<br />"
	Response.Write "visitorInstructions: " & getvisitorPreference("visitorInstructions") & "<br />"
	Response.Write "visitorPaymentmethod: " & getvisitorPreference("visitorPaymentmethod") & "<br />"
	Response.Write "</fieldset>"
	
End Sub	'displayVisitorPreferences

'**********************************************************

Function getvisitorPreference(byVal strPreference)

	If Not isArray(maryVisitorPreferences) Then Call loadVisitorPreferences("")
	
	If isArray(maryVisitorPreferences) Then
		Select Case strPreference
			Case "visitorSessionID":				getvisitorPreference = maryVisitorPreferences(0)
			Case "visitorCustomerID":				getvisitorPreference = maryVisitorPreferences(1)
			Case "visitorLastVisited":				getvisitorPreference = maryVisitorPreferences(2)
			Case "vistorDiscountCodes":				getvisitorPreference = maryVisitorPreferences(3)
			Case "visitorSelectedFreeProducts":		getvisitorPreference = maryVisitorPreferences(4)
			Case "visitorCertificateCodes":			getvisitorPreference = maryVisitorPreferences(5)
			Case "visitorLastSearch":				getvisitorPreference = maryVisitorPreferences(6)
			Case "visitorRecentlyViewedProducts":	getvisitorPreference = maryVisitorPreferences(7)
			Case "visitorCity":						getvisitorPreference = maryVisitorPreferences(8)
			Case "visitorState":					getvisitorPreference = maryVisitorPreferences(9)
			Case "visitorZIP":						getvisitorPreference = maryVisitorPreferences(10)
			Case "visitorCountry":					getvisitorPreference = maryVisitorPreferences(11)
			Case "visitorPreferredCurrency":		getvisitorPreference = maryVisitorPreferences(12)
			Case "visitorPreferredShippingCode":	getvisitorPreference = maryVisitorPreferences(13)
			Case "visitor_REFERER":					getvisitorPreference = maryVisitorPreferences(14)
			Case "vistor_HTTP_REFERER":				getvisitorPreference = maryVisitorPreferences(15)
			Case "visitor_REMOTE_ADDR":				getvisitorPreference = maryVisitorPreferences(16)
			Case "visitorLoggedInCustomerID":		getvisitorPreference = maryVisitorPreferences(17)
			Case "visitorShipAddressID":			getvisitorPreference = maryVisitorPreferences(18)
			Case "visitorOrderID":					getvisitorPreference = maryVisitorPreferences(19)
			Case "visitorID":						getvisitorPreference = maryVisitorPreferences(20)
			Case "visitorInstructions":				getvisitorPreference = maryVisitorPreferences(21)
			Case "visitorPaymentmethod":			getvisitorPreference = maryVisitorPreferences(22)
			Case "visitorEstimatedShipping":		getvisitorPreference = maryVisitorPreferences(23)
		End Select
	End If	'isArray(maryVisitorPreferences)
	
End Function	'getvisitorPreference

'**********************************************************

Function loadVisitorPreferences(byVal vntVisitorID)

Dim pblnSuccess
Dim pobjCmd
Dim pobjRS

	pblnSuccess = False
	
	'use SessionID by default, option to specify visitorID added
	'Note this check MUST BE BEFORE the isArray check since getCookie_SessionID may reference this function
	If Len(vntVisitorID) = 0 Then vntVisitorID = getCookie_SessionID
	
	If Len(vntVisitorID) = 0 Or Not isNumeric(vntVisitorID) Then
		Call addSessionDebugMessage("Invalid or No visitor id: " &  vntVisitorID)
		vntVisitorID = createVisitor(Session.SessionID, Request.QueryString("REFERER"), Request.ServerVariables("HTTP_REFERER"), Request.ServerVariables("REMOTE_ADDR"))
		Call addSessionDebugMessage("createVisitor: " &  vntVisitorID)

		Call setSessionID(vntVisitorID)
		Call setCookie_SessionID(vntVisitorID, Date() + 365)
		Call setCookie_visitorID(vntVisitorID, Date() + 365)

		'debug.PrintCookies
		'Call write_KnownCookies
	End If

	If isArray(maryVisitorPreferences) Then
		loadVisitorPreferences = True
		Exit Function
	End If
	
	Call DebugRecordSplitTime("Loading visitor preferences . . .")
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		If cblnSQLDatabase Then	'needed to split this out due to record size limitation in SQL Server
			.Commandtext = "Select visitorSessionID, visitorCustomerID, visitorLastVisited, vistorDiscountCodes, visitorySelectedFreeProducts, visitorCertificateCodes, visitorLastSearch, visitorRecentlyViewedProducts, visitorCity, visitorState, visitorZIP, visitorCountry, visitorPreferredCurrency, visitorPreferredShippingCode, visitorEstimatedShipping" _
						 & " From visitors" _
						 & " Where visitorID=?"
		Else
			.Commandtext = "Select visitorSessionID, visitorCustomerID, visitorLastVisited, vistorDiscountCodes, visitorySelectedFreeProducts, visitorCertificateCodes, visitorLastSearch, visitorRecentlyViewedProducts, visitorCity, visitorState, visitorZIP, visitorCountry, visitorPreferredCurrency, visitorPreferredShippingCode, visitorEstimatedShipping" _
						 & " ,visitor_REFERER, vistor_HTTP_REFERER, visitor_REMOTE_ADDR, visitorLoggedInCustomerID, visitorShipAddressID, visitorOrderID, visitorInstructions, visitorPaymentmethod" _
						 & " From visitors" _
						 & " Where visitorID=?"
		End If
		Set .ActiveConnection = cnn
		
		On Error Resume Next
		.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, vntVisitorID)
		Set pobjRS = .Execute
		If Err.number <> 0 Then
			loadVisitorPreferences = False
			Err.Clear
			Exit Function
		End If

		If Not pobjRS.EOF Then
			ReDim maryVisitorPreferences(23)
			maryVisitorPreferences(0) = Trim(pobjRS.Fields("visitorSessionID").Value & "")
			maryVisitorPreferences(1) = Trim(pobjRS.Fields("visitorCustomerID").Value & "")
			maryVisitorPreferences(2) = Trim(pobjRS.Fields("visitorLastVisited").Value & "")
			maryVisitorPreferences(3) = Trim(pobjRS.Fields("vistorDiscountCodes").Value & "")
			maryVisitorPreferences(4) = Trim(pobjRS.Fields("visitorySelectedFreeProducts").Value & "")
			maryVisitorPreferences(5) = Trim(pobjRS.Fields("visitorCertificateCodes").Value & "")
			maryVisitorPreferences(6) = Trim(pobjRS.Fields("visitorLastSearch").Value & "")
			maryVisitorPreferences(7) = Trim(pobjRS.Fields("visitorRecentlyViewedProducts").Value & "")
			maryVisitorPreferences(8) = Trim(pobjRS.Fields("visitorCity").Value & "")
			maryVisitorPreferences(9) = Trim(pobjRS.Fields("visitorState").Value & "")
			maryVisitorPreferences(10) = Trim(pobjRS.Fields("visitorZIP").Value & "")
			maryVisitorPreferences(11) = Trim(pobjRS.Fields("visitorCountry").Value & "")
			maryVisitorPreferences(12) = Trim(pobjRS.Fields("visitorPreferredCurrency").Value & "")
			maryVisitorPreferences(13) = Trim(pobjRS.Fields("visitorPreferredShippingCode").Value & "")
			maryVisitorPreferences(23) = Trim(pobjRS.Fields("visitorEstimatedShipping").Value & "")
			
			If cblnSQLDatabase Then
				'DebugPrintRecordset "Visitor " & vntVisitorID & " (part 1)", pobjRS
				pobjRS.Close
				.Commandtext = "Select visitor_REFERER, vistor_HTTP_REFERER, visitor_REMOTE_ADDR, visitorLoggedInCustomerID, visitorShipAddressID, visitorOrderID, visitorInstructions, visitorPaymentmethod From visitors Where visitorID=?"	'needed to split this out due to record size limitation
				Set pobjRS = .Execute
			End If
			'DebugPrintRecordset "Visitor " & vntVisitorID, pobjRS
			
			maryVisitorPreferences(14) = Trim(pobjRS.Fields("visitor_REFERER").Value & "")
			maryVisitorPreferences(15) = Trim(pobjRS.Fields("vistor_HTTP_REFERER").Value & "")
			maryVisitorPreferences(16) = Trim(pobjRS.Fields("visitor_REMOTE_ADDR").Value & "")
			maryVisitorPreferences(17) = Trim(pobjRS.Fields("visitorLoggedInCustomerID").Value & "")
			maryVisitorPreferences(18) = Trim(pobjRS.Fields("visitorShipAddressID").Value & "")
			maryVisitorPreferences(19) = Trim(pobjRS.Fields("visitorOrderID").Value & "")
			maryVisitorPreferences(20) = vntVisitorID
			maryVisitorPreferences(21) = Trim(pobjRS.Fields("visitorInstructions").Value & "")
			maryVisitorPreferences(22) = Trim(pobjRS.Fields("visitorPaymentmethod").Value & "")
			pblnSuccess = True
			
			'Now enforce some defaults, just in case
			If Len(CStr(maryVisitorPreferences(1))) = 0 Then maryVisitorPreferences(1) = 0
			If Len(CStr(maryVisitorPreferences(23))) = 0 Then maryVisitorPreferences(23) = 0	'visitorEstimatedShipping
			
			mlngVisitorID = vntVisitorID
			'Call displayVisitorPreferences
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing
	Call DebugRecordSplitTime("Visitor preferences loaded")
	
	loadVisitorPreferences = pblnSuccess
	
End Function	'loadVisitorPreferences

'**********************************************************

Sub resetVisitorPreferences()

	'need to clear the cookies first
	Call setCookie_visitorID("", Now())
	Call setCookie_SessionID("", Now())
	Call setCookie_custID("", Now())
	Session.Abandon
	
	Set maryVisitorPreferences = Nothing
	Call loadVisitorPreferences("")

	Set mclsCartTotal = New clsCartTotal
	Call removeCartFromSession
	
End Sub	'resetVisitorPreferences

'**********************************************************

Sub setvisitorPreference(byVal strPreference, byVal vntValue, byVal blnUpdateDatabase)

	If Not isArray(maryVisitorPreferences) Then Call loadVisitorPreferences("")
	
	If isArray(maryVisitorPreferences) Then
		Select Case strPreference
			Case "visitorSessionID":				maryVisitorPreferences(0) = vntValue
			Case "visitorCustomerID":				maryVisitorPreferences(1) = vntValue
			Case "visitorLastVisited":				maryVisitorPreferences(2) = vntValue
			Case "vistorDiscountCodes":				maryVisitorPreferences(3) = vntValue
			Case "visitorSelectedFreeProducts":		maryVisitorPreferences(4) = vntValue
			Case "visitorCertificateCodes":			maryVisitorPreferences(5) = vntValue
			Case "visitorLastSearch":				maryVisitorPreferences(6) = vntValue
			Case "visitorRecentlyViewedProducts":	maryVisitorPreferences(7) = vntValue
			Case "visitorCity":						maryVisitorPreferences(8) = vntValue
			Case "visitorState":					maryVisitorPreferences(9) = vntValue
			Case "visitorZIP":						maryVisitorPreferences(10) = vntValue
			Case "visitorCountry":
													maryVisitorPreferences(11) = vntValue
													'Call setCountrySpecificPricingLevel(vntValue)
			Case "visitorPreferredCurrency":		maryVisitorPreferences(12) = vntValue
			Case "visitorPreferredShippingCode":	maryVisitorPreferences(13) = vntValue
			Case "visitor_REFERER":					maryVisitorPreferences(14) = vntValue
			Case "vistor_HTTP_REFERER":				maryVisitorPreferences(15) = vntValue
			Case "visitor_REMOTE_ADDR":				maryVisitorPreferences(16) = vntValue
			Case "visitorLoggedInCustomerID":		maryVisitorPreferences(17) = vntValue
			Case "visitorShipAddressID":			maryVisitorPreferences(18) = vntValue
			Case "visitorOrderID":					maryVisitorPreferences(19) = vntValue
			Case "visitorID":						maryVisitorPreferences(20) = vntValue
			Case "visitorInstructions":				maryVisitorPreferences(21) = vntValue
			Case "visitorPaymentmethod":			maryVisitorPreferences(22) = vntValue
			Case "visitorEstimatedShipping":		maryVisitorPreferences(23) = vntValue
		End Select
	End If	'isArray(maryVisitorPreferences)
	If blnUpdateDatabase Then Call saveVisitorPreference(strPreference, vntValue)
	
End Sub	'setvisitorPreference

'**********************************************************

Function visitorCustomerID()
	visitorCustomerID = getvisitorPreference("visitorCustomerID")
End Function

Function visitorLoggedInCustomerID()
	visitorLoggedInCustomerID = getvisitorPreference("visitorLoggedInCustomerID")
End Function

Function visitorLastVisited()
	visitorLastVisited = getvisitorPreference("visitorLastVisited")
End Function

Function vistorDiscountCodes()
	vistorDiscountCodes = getvisitorPreference("vistorDiscountCodes")
End Function

Function visitorSelectedFreeProducts()
	visitorSelectedFreeProducts = getvisitorPreference("visitorSelectedFreeProducts")
End Function

Function visitorCertificateCodes()
	visitorCertificateCodes = getvisitorPreference("visitorCertificateCodes")
End Function

Function visitorLastSearch()

Dim pstrSearchPath

	pstrSearchPath = getvisitorPreference("visitorLastSearch")
	If Len(pstrSearchPath) = 0 Then pstrSearchPath = Request.Cookies("sfSearch")("SearchPath")
	If Len(pstrSearchPath) > 0 Then
		If InStr(LCase(pstrSearchPath), "login.asp") <> 0 Then pstrSearchPath = "search.asp"
	Else
		pstrSearchPath = "search.asp"
	End If

	visitorLastSearch = pstrSearchPath
	
End Function

Function visitorRecentlyViewedProducts()
	visitorRecentlyViewedProducts = getvisitorPreference("visitorRecentlyViewedProducts")
End Function

Function visitorCity()
	visitorCity = getvisitorPreference("visitorCity")
End Function

Function visitorState()
	Call setDefaultVisitorLocation
	visitorState = getvisitorPreference("visitorState")
End Function

Function visitorZIP()
	Call setDefaultVisitorLocation
	visitorZIP = getvisitorPreference("visitorZIP")
End Function

Sub setDefaultVisitorLocation

Dim pblnChanged
Dim p_State
Dim p_Zip
Dim p_Country

	pblnChanged = False
	p_Country = getvisitorPreference("visitorCountry")
	Select Case LCase(p_Country)
		Case "us":
			p_State = getvisitorPreference("visitorState")
			If Len(p_State) = 0 Then
				p_State = adminOriginState
				pblnChanged = True
			End If

			p_Zip = getvisitorPreference("adminOriginZip")
			If Len(p_Zip) = 0 Then
				p_Zip = adminOriginZip
				pblnChanged = True
			End If
			
		Case "":
			p_State = adminOriginState
			p_Zip = adminOriginState
			p_Country = adminOriginCountry
			pblnChanged = True
		Case Else:
	End Select
	'If pblnChanged Then Call updateVisitorShippingPreferences(p_State, p_Zip, p_Country, getvisitorPreference("visitorPreferredShippingCode"))
End Sub

Function visitorCountry()
	Call setDefaultVisitorLocation
	visitorCountry = getvisitorPreference("visitorCountry")
End Function

Function visitorPreferredCurrency()
	visitorPreferredCurrency = getvisitorPreference("visitorPreferredCurrency")
End Function

Function visitorPreferredShippingCode()
	visitorPreferredShippingCode = getvisitorPreference("visitorPreferredShippingCode")
End Function

Function visitorEstimatedShipping()
	visitorEstimatedShipping = getvisitorPreference("visitorEstimatedShipping")
End Function

Function visitor_REFERER()
	visitor_REFERER = getvisitorPreference("visitor_REFERER")
End Function

Function vistor_HTTP_REFERER()
	vistor_HTTP_REFERER = getvisitorPreference("vistor_HTTP_REFERER")
End Function

Function visitor_REMOTE_ADDR()
	visitor_REMOTE_ADDR = getvisitorPreference("visitor_REMOTE_ADDR")
End Function

Function visitorLoggedInCustomerID()
	visitorLoggedInCustomerID = getvisitorPreference("visitorLoggedInCustomerID")
End Function

Function visitorShipAddressID()
	visitorShipAddressID = getvisitorPreference("visitorShipAddressID")
End Function

Function visitorOrderID()
	visitorOrderID = getvisitorPreference("visitorOrderID")
End Function

Function visitorInstructions()
	visitorInstructions = getvisitorPreference("visitorInstructions")
End Function

Function visitorPaymentmethod()
	visitorPaymentmethod = getvisitorPreference("visitorPaymentmethod")
End Function

Function isCustomerLoggedIn()
	isCustomerLoggedIn = CBool(Len(CStr(visitorLoggedInCustomerID)) > 0)
End Function

'**********************************************************

Function visitorHasPriorOrders()

Dim pblnResult
Dim pobjCmd
Dim pobjRS

	If isCustomerLoggedIn Then
		pblnResult = Session("hasPriorOrders")
		If Len(pblnResult) = 0 Then
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = "SELECT Top 1 orderID FROM sfOrders Where orderCustId=?"
				Set .ActiveConnection = cnn
				.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, visitorLoggedInCustomerID)
				Set pobjRS = .Execute
				If pobjRS.EOF Then
					pblnResult = False
				Else
					pblnResult = True
				End If	'pobjRS.EOF
				pobjRS.Close
				Set pobjRS = Nothing
				Session("hasPriorOrders") = CStr(pblnResult)
			End With	'pobjCmd
			Set pobjCmd = Nothing
		Else
			pblnResult = CBool(pblnResult)
		End If
	Else
		pblnResult = False
	End If

	visitorHasPriorOrders = pblnResult
	
End Function	'visitorHasPriorOrders

'**********************************************************

Function setRecentlyViewedProducts_complete(byVal strProductID, byVal strCurrentPage)
'this function was replaced (but not deleted) since the page should always already have
'the current recently viewed products list
Dim i
Dim pobjCmd
Dim pobjRS
Dim paryRecentlyViewedProducts
Dim pstrRecentlyViewedProducts

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select visitorRecentlyViewedProducts From visitors Where visitorID=?"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, getCookie_SessionID)
		Set pobjRS = .Execute
		
		If pobjRS.EOF Then
			pstrRecentlyViewedProducts = ""
		Else
			pstrRecentlyViewedProducts = Trim(pobjRS.Fields("visitorRecentlyViewedProducts").Value & "")
			paryRecentlyViewedProducts = Split(pstrRecentlyViewedProducts, "|")
			
			pstrRecentlyViewedProducts = strProductID
			For i = 0 To UBound(paryRecentlyViewedProducts)
				If paryRecentlyViewedProducts(i) <> strProductID Then
					pstrRecentlyViewedProducts = pstrRecentlyViewedProducts & "|" & paryRecentlyViewedProducts(i)
				End If
			Next 'i
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
		
		If Len(pstrRecentlyViewedProducts) > 0 Then
			.Parameters.Delete "visitorID"
			.Commandtext = "Update visitors Set visitorRecentlyViewedProducts=?, visitorLastSearch=? Where visitorID=?"
			addParameter pobjCmd, "visitorRecentlyViewedProducts", adVarChar, pstrRecentlyViewedProducts, 255, 2
			addParameter pobjCmd, "visitorLastSearch", adLongVarWChar, strCurrentPage, visitorLastSearchMaxLength, 2
			.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, getCookie_SessionID)
			.Execute , , adExecuteNoRecords
		End If	'Len(pstrRecentlyViewedProducts) > 0

	End With	'pobjCmd
	Set pobjCmd = Nothing
	
End Function	'setRecentlyViewedProducts_complete

'**********************************************************

Function setRecentlyViewedProducts(byVal strProductID, byVal strCurrentPage)

Dim i
Dim pobjCmd
Dim paryRecentlyViewedProducts
Dim pstrRecentlyViewedProducts

	On Error Resume Next

	paryRecentlyViewedProducts = Split(visitorRecentlyViewedProducts, "|")
	
	pstrRecentlyViewedProducts = strProductID
	For i = 0 To UBound(paryRecentlyViewedProducts)
		If paryRecentlyViewedProducts(i) <> strProductID Then
			If Len(pstrRecentlyViewedProducts & "|" & paryRecentlyViewedProducts(i)) > visitorLastSearchMaxLength Then Exit For
			pstrRecentlyViewedProducts = pstrRecentlyViewedProducts & "|" & paryRecentlyViewedProducts(i)
		End If
	Next 'i
	
	If Len(pstrRecentlyViewedProducts) > 0 Then
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "Update visitors Set visitorRecentlyViewedProducts=?, visitorLastSearch=? Where visitorID=?"
			Set .ActiveConnection = cnn

			addParameter pobjCmd, "visitorRecentlyViewedProducts", adVarChar, pstrRecentlyViewedProducts, 255, 2
			addParameter pobjCmd, "visitorLastSearch", adVarChar, strCurrentPage, visitorLastSearchMaxLength, 2

			.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, getCookie_SessionID)
			.Execute , , adExecuteNoRecords
		End With	'pobjCmd
		Set pobjCmd = Nothing
	End If	'Len(pstrRecentlyViewedProducts) > 0

End Function	'setVisitorViewedProducts

'**********************************************************

Sub setVisitorLastSearch(byVal strCurrentPage)
	If Len(strCurrentPage) > 0 Then Call setVisitorPreference("visitorLastSearch", strCurrentPage, True)
End Sub	'setVisitorLastSearch

'**********************************************************

Sub saveVisitorPreference(byVal strPreference, byVal vntValue)

Dim pobjCmd
Dim pblnAbortUpdate

	pblnAbortUpdate = False
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		Select Case strPreference
			Case "visitorSessionID", "visitorCustomerID", "visitorLoggedInCustomerID", "visitor_REFERER", "visitorShipAddressID", "visitorOrderID"
				If isNumeric(vntValue) Then
					If Len(vntValue) = 0 Then vntValue = Null
					.Parameters.Append .CreateParameter("preference", adInteger, adParamInput, 4, vntValue)
				Else
					pblnAbortUpdate = True
				End If
			Case "visitorLastVisited"
				If isDate(vntValue) Then
					If Len(vntValue) = 0 Then vntValue = Null
					.Parameters.Append .CreateParameter("preference", adDBTimeStamp, adParamInput, 16, vntValue)
				Else
					pblnAbortUpdate = True
				End If
			Case "visitorLastSearch"
				addParameter pobjCmd, "preference", adVarChar, vntValue, visitorLastSearchMaxLength, 2
			Case "vistorDiscountCodes","visitorSelectedFreeProducts" , "visitorCertificateCodes", "vistor_HTTP_REFERER", "visitor_REMOTE_ADDR", "visitorRecentlyViewedProducts", "visitorInstructions"
				addParameter pobjCmd, "preference", adVarChar, vntValue, 255, 2
				If strPreference = "visitorSelectedFreeProducts" Then strPreference = "visitorySelectedFreeProducts"
			Case "visitorPreferredShippingCode":	
				addParameter pobjCmd, "preference", adVarChar, vntValue, 65, 2
			Case "visitorCity"
				addParameter pobjCmd, "preference", adVarChar, vntValue, 50, 2
			Case "visitorPaymentmethod":						
				addParameter pobjCmd, "preference", adVarChar, vntValue, 20, 2
			Case "visitorZIP", "visitorEstimatedShipping":						
				addParameter pobjCmd, "preference", adVarChar, vntValue, 10, 2
			Case "visitorState", "visitorCountry", "visitorPreferredCurrency"
				addParameter pobjCmd, "preference", adVarChar, vntValue, 3, 2
			Case Else
				pblnAbortUpdate = True
		End Select

		.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, getCookie_SessionID)
		'Call WriteCommandParameters(pobjCmd, "Abort update? " & pblnAbortUpdate, "saveVisitorPreference")
		If Not pblnAbortUpdate Then
			.Commandtype = adCmdText
			.Commandtext = "Update visitors Set " & strPreference & "=? Where visitorID=?"
			Set .ActiveConnection = cnn
			
			On Error Resume Next
			.Execute , , adExecuteNoRecords
			If Err.number <> 0 And isAdminLoggedIn Then
				Response.Write "<fieldset><legend>Error in saveVisitorPreference</legend>"
				Response.Write "Error " & err.number & ": " & err.Description & "<br />"
				Response.Write "Commandtext: " & .Commandtext & "<br />"
				Response.Write "Preference: " & strPreference & "<br />"
				Response.Write "Value: " & vntValue & "<br />"
				Response.Write "visitorID: " & getCookie_SessionID & "<br />"
				Response.Write "</fieldset>"
				Err.Clear
			End If
		End If
		
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Sub	'saveVisitorPreference

'**********************************************************

Sub updateVisitorSetOrderComplete()

Dim pobjCmd

	If Len(getCookie_SessionID) = 0 Or Not isNumeric(getCookie_SessionID) Then Exit Sub
	
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Update visitors Set visitorySelectedFreeProducts=?, vistorDiscountCodes=?, visitorCertificateCodes=? Where visitorID=?"
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("visitorySelectedFreeProducts", adVarChar, adParamInput, 3, Null)
		.Parameters.Append .CreateParameter("vistorDiscountCodes", adVarChar, adParamInput, 3, Null)
		.Parameters.Append .CreateParameter("visitorCertificateCodes", adVarChar, adParamInput, 3, Null)
		.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, getCookie_SessionID)
		.Execute , , adExecuteNoRecords
		
	End With	'pobjCmd
	Set pobjCmd = Nothing

	'Call setVisitorPreference("visitorSelectedFreeProducts", "", False)	'this is commented out otherwise it will not show up on the order confirmation
	Call setVisitorPreference("vistorDiscountCodes", "", False)
	Call setVisitorPreference("visitorCertificateCodes", "", False)
	Session.Contents.Remove("hasPriorOrders")

End Sub	'updateVisitorSetOrderComplete

'**********************************************************

Sub updateVisitorShippingPreferences(byVal strvisitorState, byVal strvisitorZIP, byVal strvisitorCountry, byVal strvisitorPreferredShippingCode)

Dim pobjCmd

	If Len(getCookie_SessionID) = 0 Or Not isNumeric(getCookie_SessionID) Then Exit Sub
	
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Update visitors Set visitorState=?, visitorZIP=?, visitorCountry=?, visitorPreferredShippingCode=? Where visitorID=?"
		Set .ActiveConnection = cnn

		addParameter pobjCmd, "visitorState", adVarChar, strvisitorState, 3, 2
		addParameter pobjCmd, "visitorZIP", adVarChar, strvisitorZIP, 10, 2
		addParameter pobjCmd, "visitorCountry", adVarChar, strvisitorCountry, 3, 2
		addParameter pobjCmd, "visitorPreferredShippingCode", adVarChar, strvisitorPreferredShippingCode, 65, 2

		.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, getCookie_SessionID)
		.Execute , , adExecuteNoRecords
		
	End With	'pobjCmd
	Set pobjCmd = Nothing
	
	Call setVisitorPreference("visitorState", strvisitorState, False)
	Call setVisitorPreference("visitorZIP", strvisitorZIP, False)
	Call setVisitorPreference("visitorCountry", strvisitorCountry, False)
	Call setVisitorPreference("visitorPreferredShippingCode", strvisitorPreferredShippingCode, False)

End Sub	'updateVisitorShippingPreferences

'**********************************************************

Sub updateVisitorOrderVerification(byVal lngvisitorLoggedInCustomerID, byVal lngvisitorShipAddressID, byVal strVisitorPaymentmethod, byVal strVisitorPreferredShippingCode, byVal strVisitorInstructions)

Dim pobjCmd

	If Len(getCookie_SessionID) = 0 Or Not isNumeric(getCookie_SessionID) Then Exit Sub
	
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Update visitors Set visitorLoggedInCustomerID=?, visitorShipAddressID=?, visitorPaymentmethod=?, visitorPreferredShippingCode=?, visitorInstructions=? Where visitorID=?"
		Set .ActiveConnection = cnn

		If isNumeric(lngvisitorLoggedInCustomerID) Then
			If Len(lngvisitorLoggedInCustomerID) = 0 Then lngvisitorLoggedInCustomerID = Null
			.Parameters.Append .CreateParameter("visitorLoggedInCustomerID", adInteger, adParamInput, 4, lngvisitorLoggedInCustomerID)
		End If

		If isNumeric(lngvisitorShipAddressID) Then
			If Len(lngvisitorShipAddressID) = 0 Then lngvisitorShipAddressID = Null
			.Parameters.Append .CreateParameter("visitorShipAddressID", adInteger, adParamInput, 4, lngvisitorShipAddressID)
		End If

		addParameter pobjCmd, "visitorPaymentmethod", adVarChar, strVisitorPaymentmethod, 20, 2
		addParameter pobjCmd, "visitorPreferredShippingCode", adVarChar, strvisitorPreferredShippingCode, 65, 2
		addParameter pobjCmd, "preference", adVarChar, strVisitorInstructions, 255, 2

		.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, getCookie_SessionID)
		.Execute , , adExecuteNoRecords
		
	End With	'pobjCmd
	Set pobjCmd = Nothing

	Call setVisitorPreference("visitorLoggedInCustomerID", lngvisitorLoggedInCustomerID, False)
	Call setVisitorPreference("visitorShipAddressID", lngvisitorShipAddressID, False)
	Call setVisitorPreference("visitorPaymentmethod", strVisitorPaymentmethod, False)
	Call setVisitorPreference("visitorPreferredShippingCode", strvisitorPreferredShippingCode, False)
	Call setVisitorPreference("visitorInstructions", strVisitorInstructions, False)

End Sub	'updateVisitorOrderVerification

'******************************************************************************************************************************************************

Function createVisitor(byVal visitorSessionID, byVal visitor_REFERER, byVal vistor_HTTP_REFERER, byVal visitor_REMOTE_ADDR)
'This code must is duplicated in global.asa

Dim pobjCmd
Dim pobjRS
Dim pstrSQL
'SELECT @@IDENTITY

	pstrSQL = "Insert Into visitors (visitorSessionID, visitorDateCreated, visitorLastVisited, visitorCountry, visitor_REFERER, vistor_HTTP_REFERER, visitor_REMOTE_ADDR, visitorPreferredCurrency) Values (?, ?, ?, ?, ?, ?, ?, ?)"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		'.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInputOutput, 4, NULL)
		.Parameters.Append .CreateParameter("visitorSessionID", adInteger, adParamInput, 4, visitorSessionID)
		.Parameters.Append .CreateParameter("visitorDateCreated", adDBTimeStamp, adParamInput, 16, Now())
		.Parameters.Append .CreateParameter("visitorLastVisited", adDBTimeStamp, adParamInput, 16, Now())
		.Parameters.Append .CreateParameter("visitorCountry", adVarChar, adParamInput, 3, "US")
		
		If Len(visitor_REFERER) > 0 And isNumeric(visitor_REFERER) Then
			.Parameters.Append .CreateParameter("visitor_REFERER", adInteger, adParamInput, 4, visitor_REFERER)
		Else
			.Parameters.Append .CreateParameter("visitor_REFERER", adInteger, adParamInput, 4, 0)
		End If
		
		addParameter pobjCmd, "vistor_HTTP_REFERER", adVarChar, vistor_HTTP_REFERER, 255, 2
		addParameter pobjCmd, "visitor_REMOTE_ADDR", adVarChar, visitor_REMOTE_ADDR, 255, 2
		addParameter pobjCmd, "visitorPreferredCurrency", adVarChar, "US", 3, 2

		.Execute , , adExecuteNoRecords
		
		.Parameters.Delete "visitorLastVisited"
		.Parameters.Delete "visitorCountry"
		.Parameters.Delete "visitor_REFERER"
		.Parameters.Delete "vistor_HTTP_REFERER"
		.Parameters.Delete "visitor_REMOTE_ADDR"
		.Parameters.Delete "visitorPreferredCurrency"
		pstrSQL = "Select visitorID From visitors Where visitorSessionID=? And visitorDateCreated=?"
		.Commandtext = pstrSQL

		Set pobjRS = .Execute
		If pobjRS.EOF Then
			createVisitor = -1
		Else
			createVisitor = pobjRS.Fields("visitorID").Value
		End If
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Function	'createVisitor

'******************************************************************************************************************************************************

Function visitorExists(byVal visitorID)
'This code must is duplicated in global.asa

Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	pstrSQL = "Select visitorSessionID From visitors Where visitorID=?"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select visitorSessionID From visitors Where visitorID=?"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, visitorID)
		Set pobjRS = .Execute
		
		If pobjRS.EOF Then
			visitorExists = False
		Else
			.Parameters.Delete "visitorID"
			.Commandtext = "Update visitors Set visitorLastVisited=? Where visitorID=?"
			.Parameters.Append .CreateParameter("visitorLastVisited", adDBTimeStamp, adParamInput, 16, Now())
			.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, visitorID)
			.Execute , , adExecuteNoRecords
			visitorExists = True
		End If

		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Function	'visitorExists

'**********************************************************

Sub setVisitorLoggedInCustomerID(byVal vntValue)
	Call setVisitorPreference("visitorLoggedInCustomerID", vntValue, True)
End Sub

Sub setVisitorCustomerID(byVal vntValue)
	Call setVisitorPreference("visitorCustomerID", vntValue, True)
End Sub

Sub setVisitorShipAddressID(byVal vntValue)
	Call setVisitorPreference("visitorShipAddressID", vntValue, True)
End Sub

Sub setVisitorOrderID(byVal vntValue)
	Call setVisitorPreference("visitorOrderID", vntValue, True)
End Sub

'**********************************************************

Sub setVisitorShippingLocationByCustomerID(byVal lngCustID)

Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	pstrSQL = "SELECT custState, custZip, custCountry, shipCode" _
			& " FROM (sfCustomers LEFT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN sfShipping ON sfOrders.orderShipMethod = sfShipping.shipMethod" _
			& " WHERE orderCustId=?" _
			& " ORDER BY sfOrders.orderDate DESC"

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, lngCustID)
		Set pobjRS = .Execute
		
		If Not pobjRS.EOF Then
			'You could check to make sure you're not overwriting a pre-exisiting value
			Call updateVisitorShippingPreferences(pobjRS.Fields("custState").Value, pobjRS.Fields("custZip").Value, pobjRS.Fields("custCountry").Value, pobjRS.Fields("shipCode").Value)
		End If	'pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Sub	'setVisitorShippingLocationByCustomerID

'**********************************************************

Function isValidRecordID(byVal vntID)
'a valid id would be a value > 0

	If Len(CStr(vntID)) = 0 Or Not isNumeric(vntID) Then
		isValidRecordID = False
	ElseIf vntID < 1 Then
		isValidRecordID = False
	Else
		isValidRecordID = True
	End If

End Function	'isValidRecordID

'**********************************************************

Sub setCountrySpecificPricingLevel(byVal strCountryAbbr)

Dim plngPricingLevelID

	Select Case UCase(strCountryAbbr)
		Case "US":	plngPricingLevelID = ""	'Use base prices
		Case "CA":	plngPricingLevelID = 1
		Case Else:	plngPricingLevelID = 1
	End Select

	Session("custPricingLevel") = plngPricingLevelID
	
	Call setPricingLevel

End Sub	'setCountrySpecificPricingLevel

'**********************************************************

Sub recordPageView()

Dim PageName
Dim PageQueryString
Dim TimeViewed
Dim PageReferrer
Dim SearchKeyWords
Dim SearchResultCount
Dim pobjCmd
Dim pobjRS

	If Not cblnTrackPageViews Then Exit Sub
	'Collect values
	PageName = Request.ServerVariables("SCRIPT_NAME")
	PageQueryString = Request.QueryString
	TimeViewed = Now()
	PageReferrer = Request.ServerVariables("HTTP_REFERER")
	
	If Len(PageName) > 0 And visitorPageViewID = 0 Then
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "Insert Into visitorPageViews (visitorID, PageName, PageQueryString, PageReferrer, TimeViewed) Values (?,?,?,?,?)"
			Set .ActiveConnection = cnn

			.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, getCookie_SessionID)
			
			addParameter pobjCmd, "PageName", adVarChar, PageName, 255, 3
			addParameter pobjCmd, "PageQueryString", adVarChar, PageQueryString, 255, 2
			addParameter pobjCmd, "PageReferrer", adVarChar, PageReferrer, 255, 3

			.Parameters.Append .CreateParameter("TimeViewed", adDBTimeStamp, adParamInput, 16, TimeViewed)

			.Execute , , adExecuteNoRecords
			
			.Commandtext = "Select visitorPageViewID From visitorPageViews Where visitorID=? And PageName=? And PageQueryString=? And PageReferrer=? And TimeViewed=?"
			Set pobjRS = .Execute
			If Not pobjRS.EOF Then	visitorPageViewID = pobjRS.Fields("visitorPageViewID").Value
			pobjRS.Close
			Set pobjRS = Nothing
			
		End With	'pobjCmd
		Set pobjCmd = Nothing
	End If	'Len(PageName) > 0 And visitorPageViewID = 0

End Sub	'recordPageView

'**********************************************************

Sub recordSearchResults(byVal TypeKeyWordSearch, byVal SearchKeyWords, byVal SearchResultCount)

Dim pobjCmd

	If Not cblnTrackPageViews Then Exit Sub
	If Len(SearchResultCount) = 0 Or Not isNumeric(SearchResultCount) Then Exit Sub
	
	If Len(SearchKeyWords) > 0 And visitorPageViewID <> 0 Then
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "Update visitorPageViews Set TypeKeyWordSearch=?, SearchKeyWords=?, SearchResultCount=? Where visitorPageViewID=?"
			Set .ActiveConnection = cnn

			addParameter pobjCmd, "TypeKeyWordSearch", adVarChar, TypeKeyWordSearch, 5, 2
			addParameter pobjCmd, "SearchKeyWords", adVarChar, SearchKeyWords, 255, 2

			.Parameters.Append .CreateParameter("SearchResultCount", adInteger, adParamInput, 4, SearchResultCount)
			.Parameters.Append .CreateParameter("visitorPageViewID", adInteger, adParamInput, 4, visitorPageViewID)

			.Execute , , adExecuteNoRecords
			
		End With	'pobjCmd
		Set pobjCmd = Nothing
	End If	'Len(PageName) > 0 And visitorPageViewID = 0

	If Not cblnTrackPageViews Then Exit Sub
	If Len(SearchResultCount) = 0 Or Not isNumeric(SearchResultCount) Then Exit Sub
	
	If visitorPageViewID <> 0 Then
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "Update visitorPageViews Set SearchResultCount=? Where visitorPageViewID=?"
			Set .ActiveConnection = cnn

			.Parameters.Append .CreateParameter("SearchResultCount", adInteger, adParamInput, 4, SearchResultCount)
			.Parameters.Append .CreateParameter("visitorPageViewID", adInteger, adParamInput, 4, visitorPageViewID)

			.Execute , , adExecuteNoRecords
			
		End With	'pobjCmd
		Set pobjCmd = Nothing
	End If	'Len(PageName) > 0 And visitorPageViewID = 0

End Sub	'recordSearchResults
%>
