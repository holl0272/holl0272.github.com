<%
'********************************************************************************
'*   Customer Class module for StoreFront 5.0                                   *
'*   Release Version:   1.00.001                                                *
'*   Release Date:      August 1, 2002											*
'*   Revision Date:     April 30, 2004											*
'*                                                                              *
'*   Revision History															*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Class clsCustomer

Private pobjCnn
Private pblnConnectionOpenedInternally

Private plngcustID
Private pstrcustFirstName
Private pstrcustMiddleInitial
Private pstrcustLastName
Private pstrcustCompany
Private pstrcustAddr1
Private pstrcustAddr2
Private pstrcustCity
Private pstrcustState
Private pstrcustStateName
Private pblnStatePreDefined
Private pstrcustZip
Private pstrcustCountry
Private pstrcustCountryName
Private pstrcustPhone
Private pstrcustFax
Private pstrcustPasswd
Private pstrcustEmail
Private plngcustTimesAccessed
Private pdtcustLastAccess
Private pbytcustIsSubscribed
Private plngPricingLevelID
Private pstrClubExpDate
Private pstrClubCode
Private plngNumAccounts
Private plngPriorOrderCount
Private pblnHasDownloadableItems

Private pstrConfirmPassword
Private cstrFieldGreeting
Private pstrGreeting

Private pstrMessage
Private pblnError

Private cstrSavedCartCustomerName

	'****************************************************************************************************************

	Private Sub Class_Initialize
	
		cstrFieldGreeting = "custFirstName" 
		cstrSavedCartCustomerName = "Saved Cart Customer"
		pblnConnectionOpenedInternally = False
		pblnError = False

		'Now for the defaults
		plngcustID = 0
		plngcustTimesAccessed = 1
		pdtcustLastAccess = Now()
		pbytcustIsSubscribed = 1
		plngPricingLevelID = 0
		pblnStatePreDefined = False
		plngNumAccounts = 0
		
		'Intentionally not initialized
		plngPriorOrderCount = ""
		pblnHasDownloadableItems = ""

	End Sub	'Class_Initialize

	'****************************************************************************************************************

	Private Sub Class_Terminate
		If pblnConnectionOpenedInternally Then
			On Error Resume Next
			pobjcnn.Close
			Set pobjcnn = Nothing
		End If
	End Sub

	'****************************************************************************************************************

	Private Function InitializeConnection
	'returns a connection to the database 

		On Error Resume Next
		If Err.number <> 0 Then Err.Clear
		
		If isObject(cnn) Then
			Set pobjcnn = cnn
			InitializeConnection = True
		Else
			Set pobjcnn = CreateObject("ADODB.Connection")

			If len(Application("DSN_NAME")) > 0 Then
				pobjcnn.Open Application("DSN_NAME")
			End If

			pblnConnectionOpenedInternally = CBool(pobjcnn.State = 1)
			If Not pblnConnectionOpenedInternally Then Response.Write "<font color=red>Error opening database: " & pobjcnn.ConnectionString & "</font>"
			InitializeConnection = pblnConnectionOpenedInternally
		End If
				
	End Function		' InitializeConnection
	
	'****************************************************************************************************************

	Public Property Get Message
		Message = pstrMessage
	End Property
	
	Public Property Get Error
		Error = pblnError
	End Property
	
	Public Property Set Connection(byRef objCnn)
		If isObject(objCnn) Then Set pobjCnn = objCnn
	End Property

	Public Property Let custID(byVal value)
		plngcustID = value
	End Property
	Public Property Get custID
		custID = plngcustID
	End Property
	
	Public Property Let custFirstName(byVal value)
		pstrcustFirstName = value
	End Property
	Public Property Get custFirstName
		custFirstName = pstrcustFirstName
	End Property
	
	Public Property Let custMiddleInitial(byVal value)
		pstrcustMiddleInitial = value
	End Property
	Public Property Get custMiddleInitial
		custMiddleInitial = pstrcustMiddleInitial
	End Property
	
	Public Property Let custLastName(byVal value)
		pstrcustLastName = value
	End Property
	Public Property Get custLastName
		custLastName = pstrcustLastName
	End Property
	
	Public Property Let custCompany(byVal value)
		pstrcustCompany = value
	End Property
	Public Property Get custCompany
		custCompany = pstrcustCompany
	End Property
	
	Public Property Let custAddr1(byVal value)
		pstrcustAddr1 = value
	End Property
	Public Property Get custAddr1
		custAddr1 = pstrcustAddr1
	End Property
	
	Public Property Let custAddr2(byVal value)
		pstrcustAddr2 = value
	End Property
	Public Property Get custAddr2
		custAddr2 = pstrcustAddr2
	End Property
	
	Public Property Let custCity(byVal value)
		pstrcustCity = value
	End Property
	Public Property Get custCity
		custCity = pstrcustCity
	End Property

	Public Property Let custState(byVal value)
		pstrcustState = value
	End Property
	Public Property Get custState
		custState = pstrcustState
	End Property
	
	Public Property Let custZip(byVal value)
		pstrcustZip = value
	End Property
	Public Property Get custZip
		custZip = pstrcustZip
	End Property
	
	Public Property Let custCountry(byVal value)
		pstrcustCountry = value
	End Property
	Public Property Get custCountry
		custCountry = pstrcustCountry
	End Property
	
	Public Property Let custPhone(byVal value)
		pstrcustPhone = value
	End Property
	Public Property Get custPhone
		custPhone = pstrcustPhone
	End Property
	
	Public Property Let custFax(byVal value)
		pstrcustFax = value
	End Property
	Public Property Get custFax
		custFax = pstrcustFax
	End Property
	
	Public Property Let ConfirmPassword(byVal value)
		pstrConfirmPassword = value
	End Property
	Public Property Let custPasswd(byVal value)
		pstrcustPasswd = value
	End Property
	Public Property Get custPasswd
		custPasswd = pstrcustPasswd
	End Property

	Public Property Let custEmail(byVal value)
		pstrcustEmail = value
	End Property
	Public Property Get custEmail
		custEmail = pstrcustEmail
	End Property
	
	Public Property Let custTimesAccessed(byVal value)
		plngcustTimesAccessed = value
	End Property
	Public Property Get custTimesAccessed
		custTimesAccessed = plngcustTimesAccessed
	End Property
	
	Public Property Let custLastAccess(byVal value)
		pdtcustLastAccess = value
	End Property
	Public Property Get custLastAccess
		custLastAccess = pdtcustLastAccess
	End Property
	
	Public Property Let custIsSubscribed(byVal value)
		pbytcustIsSubscribed = value
	End Property
	Public Property Get custIsSubscribed
		custIsSubscribed = pbytcustIsSubscribed
	End Property
	
	Public Property Let PricingLevelID(byVal value)
		plngPricingLevelID = value
	End Property
	Public Property Get PricingLevelID
		PricingLevelID = plngPricingLevelID
	End Property

	Public Property Get Greeting
		Greeting = pstrGreeting
	End Property
	
	Public Property Get IsSubscribed
		IsSubscribed = CBool(pbytcustIsSubscribed = 1)
	End Property
	
	Public Property Get IsSavedCartCustomer
		IsSavedCartCustomer = CBool(pstrcustFirstName = cstrSavedCartCustomerName)
	End Property

	Public Property Let ClubExpDate(byVal value)
		pstrClubExpDate = value
	End Property
	Public Property Get ClubExpDate
		ClubExpDate = pstrClubExpDate
	End Property
	Public Property Let ClubCode(byVal value)
		pstrClubCode = value
	End Property
	Public Property Get ClubCode
		ClubCode = pstrClubCode
	End Property
	
	Public Property Get NumAccounts
		NumAccounts = plngNumAccounts
	End Property

	'***********************************************************************************************

	Public Property Get DisplayName
		If IsSavedCartCustomer Then
			DisplayName = ""
		ElseIf Len(pstrcustMiddleInitial) > 0 Then
			DisplayName = pstrcustFirstName & " " & pstrcustMiddleInitial & " " & pstrcustLastName
		Else
			DisplayName = pstrcustFirstName & " " & pstrcustLastName
		End If
	End Property

	'***********************************************************************************************

	Public Property Get countryName
		If Len(pstrcustCountryName) = 0 And Len(pstrcustCountry) > 0 Then pstrcustCountryName = Trim(getNameWithID("sfLocalesCountry", pstrcustCountry, "loclctryAbbreviation", "loclctryName", 1))
		countryName = pstrcustCountryName
	End Property

	'***********************************************************************************************

	Public Property Get stateName
		If Len(pstrcustStateName) = 0 And Len(pstrcustState) > 0 Then
			pstrcustStateName = Trim(getNameWithID("sfLocalesState", pstrcustState, "loclstAbbreviation", "loclstName", 1))
			If Len(pstrcustStateName) = 0 Then
				pstrcustStateName = pstrcustState
				pblnStatePreDefined = False
			Else
				pblnStatePreDefined = True
			End If
		End If
		stateName = pstrcustStateName
	End Property

	'***********************************************************************************************

	Public Property Get StatePreDefined
		StatePreDefined = pblnStatePreDefined
	End Property

	'***********************************************************************************************

	Private Function ValidateValues
	
	Dim pblnResult

		Call checkLength(pstrcustFirstName, 50, True, True)
		Call checkLength(pstrcustMiddleInitial, 1, True, True)
		Call checkLength(pstrcustLastName, 50, True, True)
		Call checkLength(pstrcustCompany, 60, True, True)
		Call checkLength(pstrcustAddr1, 50, True, True)
		Call checkLength(pstrcustAddr2, 50, True, True)
		Call checkLength(pstrcustCity, 50, True, True)
		Call checkLength(pstrcustState, 25, True, True)
		Call checkLength(pstrcustZip, 12, True, True)
		Call checkLength(pstrcustCountry, 50, True, True)
		Call checkLength(pstrcustPhone, 25, True, True)
		Call checkLength(pstrcustFax, 25, True, True)
		
		If Len(plngcustID) > 0 And Not isNumeric(plngcustID) Then plngcustID = ""
		If Len(plngcustTimesAccessed) = 0 Or Not isNumeric(plngcustTimesAccessed) Then plngcustTimesAccessed = 1
		If Not isDate(pdtcustLastAccess) Then pdtcustLastAccess = Now()
		If Len(pbytcustIsSubscribed) = 0 Or Not isNumeric(pbytcustIsSubscribed) Then pbytcustIsSubscribed = 1
		If Len(plngPricingLevelID) = 0 Or Not isNumeric(plngPricingLevelID) Then plngPricingLevelID = 0
		
		'now for the required fields
		If fncEmailValid(pstrcustEmail) Then
			pblnResult = checkLength(pstrcustEmail, 100, False, True)
		Else
			pstrMessage = "Invalid email"
			pblnResult = False
		End If

		pblnResult = pblnResult And checkLength(pstrcustPasswd, 10, False, True)
		
		ValidateValues = pblnResult

	End Function	'ValidateValues

	'***********************************************************************************************

	Private Sub LoadValues(byRef objRS)
	
		If Not isObject(objRS) Then Exit Sub
		
		With objRS
			If Not .EOF Then
				plngcustID = trim(.Fields("custID").Value & "")
				pstrcustFirstName = trim(.Fields("custFirstName").Value & "")
				pstrcustMiddleInitial = trim(.Fields("custMiddleInitial").Value & "")
				pstrcustLastName = trim(.Fields("custLastName").Value & "")
				pstrcustCompany = trim(.Fields("custCompany").Value & "")
				pstrcustAddr1 = trim(.Fields("custAddr1").Value & "")
				pstrcustAddr2 = trim(.Fields("custAddr2").Value & "")
				pstrcustCity = trim(.Fields("custCity").Value & "")
				pstrcustState = trim(.Fields("custState").Value & "")
				pstrcustZip = trim(.Fields("custZip").Value & "")
				pstrcustCountry = trim(.Fields("custCountry").Value & "")
				pstrcustPhone = trim(.Fields("custPhone").Value & "")
				pstrcustFax = trim(.Fields("custFax").Value & "")
				pstrcustPasswd = trim(.Fields("custPasswd").Value & "")
				pstrcustEmail = trim(.Fields("custEmail").Value & "")
				plngcustTimesAccessed = trim(.Fields("custTimesAccessed").Value & "")
				If Len(plngcustTimesAccessed) = 0 Then plngcustTimesAccessed = 0
				
				pdtcustLastAccess = trim(.Fields("custLastAccess").Value & "")
				pbytcustIsSubscribed = trim(.Fields("custIsSubscribed").Value & "")
				plngPricingLevelID = trim(.Fields("PricingLevelID").Value & "")
				pstrClubExpDate = trim(.Fields("clubExpDate").Value & "")
				pstrClubCode = trim(.Fields("clubCode").Value & "")
				
				If Len(pbytcustIsSubscribed) = 0 Or Not isNumeric(pbytcustIsSubscribed) Then pbytcustIsSubscribed = 1

				Call SetGreeting
			End If
		End With
		
	End Sub	'LoadValues

	'***********************************************************************************************

	Private Sub SetGreeting
	
		If LCase(pstrcustFirstName) = "saved cart customer" Then
			pstrGreeting = "Guest"
		Else
			pstrGreeting = pstrcustFirstName
		End If
		
	End Sub	'SetGreeting

	'***********************************************************************************************

	Private Function CreateEmptyCustomer

	Dim pblnResult
	Dim pobjCmd
	Dim pobjRS
	Dim pstrSQL
	Dim pstrTempID
	
	'On Error Resume Next

		pstrTempID = SessionID & CStr(Now())

		pstrSQL = "Insert Into sfCustomers (custEmail) Values (?)"

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtext = pstrSQL
			.Commandtype = adCmdText
			'.Commandtype = adCmdStoredProc
			Set .ActiveConnection = pobjcnn
			addParameter pobjCmd, "custEmail", adWChar, pstrTempID, 100, 0
			.Execute , , adExecuteNoRecords
		
			'plngcustID = .Parameters("custID")
			
			pstrSQL = "Select custID From sfCustomers Where custEmail=?"
			.Commandtext = pstrSQL
			Set pobjRS = .Execute			
			If pobjRS.EOF Then
				plngcustID = -1
				Response.Write "<h4><font color=red>Error Retrieving New Customer</font></h4>"
				pblnResult = False
			Else
				plngcustID = pobjRS.Fields("custID").Value
				pblnResult = True
			End If
			Call closeObj(pobjRS)

		End With	'pobjCmd
		Call closeObj(pobjCmd)

		CreateEmptyCustomer = pblnResult

	End Function	'CreateEmptyCustomer

	'***********************************************************************************************

	Public Function AddCustomer

	Dim pblnAlreadyExists
	Dim pblnResult
	Dim pobjCmd
	Dim pstrSQL
	
	'On Error Resume Next

		If Not ValidateValues Then
			AddCustomer = False
			Exit Function
		End If
		
		'See if customer exists
		If len(plngcustID) > 0 Then
			pblnAlreadyExists = LoadCustomer(plngcustID)
		ElseIf Len(pstrcustEmail) > 0 And Len(pstrcustPasswd) > 0 Then
			If fncEmailValid(pstrcustEmail) Then
				pblnAlreadyExists = LoadCustomerByEmailPassword(pstrcustEmail, pstrcustPasswd)
				If pblnAlreadyExists Then
					pstrMessage = "This user already exists."
					pblnError = True
				End If
			Else
				pstrMessage = "Invalid email address."
				pblnError = True
			End If
		Else
			If Len(pstrcustEmail) = 0 Then
				pstrMessage = "Invalid email address."
				pblnError = True
			Else
				pstrMessage = "Invalid password."
				pblnError = True
			End If
		End If
		
		If pblnError Then
			pblnResult = False
		Else
			If Not pblnAlreadyExists Then pblnAlreadyExists = CreateEmptyCustomer 
			If pblnAlreadyExists Then
				'debugprint "pstrSQL", pstrSQL

				pstrSQL = "Update sfCustomers Set " _
						& "custFirstName=?," _
						& "custMiddleInitial=?," _
						& "custLastName=?," _
						& "custCompany=?," _
						& "custAddr1=?," _
						& "custAddr2=?," _
						& "custCity=?," _
						& "custState=?," _
						& "custZip=?," _
						& "custCountry=?," _
						& "custPhone=?," _
						& "custFax=?," _
						& "custPasswd=?," _
						& "custEmail=?," _
						& "custTimesAccessed=?," _
						& "custLastAccess=?," _
						& "custIsSubscribed=?" _
						& " Where custID=?"

						'& "PricingLevelID=?," _
						'& "clubCode=?," _
						'& "clubExpDate=?" _
				Set pobjCmd  = CreateObject("ADODB.Command")
				With pobjCmd
					.Commandtext = pstrSQL
					.Commandtype = adCmdText
					'.Commandtype = adCmdStoredProc
					Set .ActiveConnection = pobjcnn

					addParameter pobjCmd, "custFirstName", adWChar, pstrcustFirstName, 50, 2
					addParameter pobjCmd, "custMiddleInitial", adWChar, pstrcustMiddleInitial, 1, 2
					addParameter pobjCmd, "custLastName", adWChar, pstrcustLastName, 50, 2
					addParameter pobjCmd, "custCompany", adWChar, pstrcustCompany, 60, 2
					addParameter pobjCmd, "custAddr1", adWChar, pstrcustAddr1, 50, 2
					addParameter pobjCmd, "custAddr2", adWChar, pstrcustAddr2, 50, 2
					addParameter pobjCmd, "custCity", adWChar, pstrcustCity, 50, 2
					addParameter pobjCmd, "custState", adWChar, pstrcustState, 25, 2
					addParameter pobjCmd, "custZip", adWChar, pstrcustZip, 12, 2
					addParameter pobjCmd, "custCountry", adWChar, pstrcustCountry, 50, 2
					addParameter pobjCmd, "custPhone", adWChar, pstrcustPhone, 25, 2
					addParameter pobjCmd, "custFax", adWChar, pstrcustFax, 25, 2
					addParameter pobjCmd, "custPasswd", adWChar, pstrcustPasswd, 10, 2
					addParameter pobjCmd, "custEmail", adWChar, pstrcustEmail, 100, 2

					.Parameters.Append .CreateParameter("custTimesAccessed", adInteger, adParamInput, 4, plngcustTimesAccessed)
					.Parameters.Append .CreateParameter("custLastAccess", adDBTimeStamp, adParamInput, 16, pdtcustLastAccess)
					.Parameters.Append .CreateParameter("custIsSubscribed", adSmallInt, adParamInput, 2, pbytcustIsSubscribed)
					'.Parameters.Append .CreateParameter("PricingLevelID", adInteger, adParamInput, 4, plngPricingLevelID)
					'.Parameters.Append .CreateParameter("clubCode", adWChar, adParamInput, 50, checkFieldLength(pstrClubCode, 50, 2))
					'.Parameters.Append .CreateParameter("clubExpDate", adDBTimeStamp, adParamInput, 16, pstrClubExpDate)
					.Parameters.Append .CreateParameter("custID", adInteger, adParamInput, 4, plngcustID)

					On Error Resume Next
					.Execute , , adExecuteNoRecords
					If Err.number = 0 Then
						pblnResult = True
					Else
						Response.Write "<h4>Error " & Err.number & ": " & Err.Description & "</h4>"
						err.Clear
						pblnResult = False
					End If
				
				End With	'pobjCmd
				Call closeObj(pobjCmd)

			Else
				pblnResult = False
			End If
		End If

		AddCustomer = pblnResult

	End Function	'AddCustomer

	'***********************************************************************************************

	Public Function Update

	Dim pblnResult
	Dim pobjCmd
	Dim pstrSQL
	
	'On Error Resume Next

		If ValidateValues Then
			pstrSQL = "Update sfCustomers Set " _
					& "custFirstName=?," _
					& "custMiddleInitial=?," _
					& "custLastName=?," _
					& "custCompany=?," _
					& "custAddr1=?," _
					& "custAddr2=?," _
					& "custCity=?," _
					& "custState=?," _
					& "custZip=?," _
					& "custCountry=?," _
					& "custPhone=?," _
					& "custFax=?," _
					& "custPasswd=?," _
					& "custEmail=?," _
					& "custTimesAccessed=?," _
					& "custLastAccess=?," _
					& "custIsSubscribed=?," _
					& "PricingLevelID=?" _
					& " Where custID=?"

			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtext = pstrSQL
				.Commandtype = adCmdText
				'.Commandtype = adCmdStoredProc
				Set .ActiveConnection = pobjcnn

				addParameter pobjCmd, "custFirstName", adWChar, pstrcustFirstName, 50, 2
				addParameter pobjCmd, "custMiddleInitial", adWChar, pstrcustMiddleInitial, 1, 2
				addParameter pobjCmd, "custLastName", adWChar, pstrcustLastName, 50, 2
				addParameter pobjCmd, "custCompany", adWChar, pstrcustCompany, 60, 2
				addParameter pobjCmd, "custAddr1", adWChar, pstrcustAddr1, 50, 2
				addParameter pobjCmd, "custAddr2", adWChar, pstrcustAddr2, 50, 2
				addParameter pobjCmd, "custCity", adWChar, pstrcustCity, 50, 2
				addParameter pobjCmd, "custState", adWChar, pstrcustState, 25, 2
				addParameter pobjCmd, "custZip", adWChar, pstrcustZip, 12, 2
				addParameter pobjCmd, "custCountry", adWChar, pstrcustCountry, 50, 2
				addParameter pobjCmd, "custPhone", adWChar, pstrcustPhone, 25, 2
				addParameter pobjCmd, "custFax", adWChar, pstrcustFax, 25, 2
				addParameter pobjCmd, "custPasswd", adWChar, pstrcustPasswd, 10, 2
				addParameter pobjCmd, "custEmail", adWChar, pstrcustEmail, 100, 2

				.Parameters.Append .CreateParameter("custTimesAccessed", adInteger, adParamInput, 4, plngcustTimesAccessed)
				.Parameters.Append .CreateParameter("custLastAccess", adDBTimeStamp, adParamInput, 16, pdtcustLastAccess)
				.Parameters.Append .CreateParameter("custIsSubscribed", adSmallInt, adParamInput, 2, pbytcustIsSubscribed)
				.Parameters.Append .CreateParameter("PricingLevelID", adInteger, adParamInput, 4, plngPricingLevelID)
				'.Parameters.Append .CreateParameter("clubCode", adWChar, adParamInput, 50, checkFieldLength(pstrClubCode, 50, 2))
				'.Parameters.Append .CreateParameter("clubExpDate", adDBTimeStamp, adParamInput, 16, pstrClubExpDate)
				.Parameters.Append .CreateParameter("custID", adInteger, adParamInput, 4, plngcustID)

				.Execute , , adExecuteNoRecords
			
			End With	'pobjCmd
			Call closeObj(pobjCmd)

			If Err.number = 0 Then
				pblnResult = True
			Else
				pblnResult = False
				Err.Clear
			End If
		Else
			pblnResult = False
		End If

		Update = pblnResult

	End Function	'Update

	'***********************************************************************************************

	Public Function LoadCustomer(byVal lngCustID)

	Dim pblnResult
	Dim pobjCmd
	Dim pobjRS
	
	'On Error Resume Next

		If validNumber(lngCustID) Then
			If Not isObject(pobjcnn) Then
				If Not InitializeConnection Then Exit Function
			End If
			
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtext = "Select * From sfCustomers Where custID=?"
				.Commandtype = adCmdText
				'.Commandtype = adCmdStoredProc
				Set .ActiveConnection = pobjcnn

				.Parameters.Append .CreateParameter("custID", adInteger, adParamInput, 4, lngCustID)
				Set pobjRS = .Execute			
				If pobjRS.EOF Then
					pblnResult = False
				Else
					Call LoadValues(pobjRS)
					pblnResult = True
				End If
				Call closeObj(pobjRS)

			End With	'pobjCmd
			Call closeObj(pobjCmd)
		Else
			pstrMessage = "Invalid Customer ID"
			pblnError = True
			pblnResult = False
		End If	'validNumber(lngCustID)

		LoadCustomer = pblnResult

	End Function	'LoadCustomer

	'***********************************************************************************************

	Public Function LoadCustomerByEmail(byVal strEmail)

	Dim pblnResult
	Dim pobjCmd
	Dim pobjRS
	
	'On Error Resume Next

		If len(strEmail) = 0 then
			pstrMessage = "Please enter an Email address"
			pblnError = True
			pblnResult = False
		ElseIf Not fncEmailValid(strEmail) Then
			pstrMessage = "Invalid email address."
			pblnError = True
			pblnResult = False
		Else
			If Not isObject(pobjcnn) Then
				If Not InitializeConnection Then
					pstrMessage = "Unable to connect to database"
					pblnError = True
					LoadCustomerByEmail = False
					Exit Function
				End If
			End If
			
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtext = "Select * From sfCustomers where custEmail=? Order By custID Desc"
				.Commandtype = adCmdText
				'.Commandtype = adCmdStoredProc
				Set .ActiveConnection = pobjcnn

				addParameter pobjCmd, "custEmail", adWChar, strEmail, 100, 2
				Set pobjRS = .Execute			
				If pobjRS.EOF Then
					pblnResult = False
				Else
					Call LoadValues(pobjRS)
					Do While Not pobjRS.EOF
						plngNumAccounts = plngNumAccounts + 1
						pobjRS.MoveNext
					Loop
					pblnResult = True
				End If
				Call closeObj(pobjRS)
			End With	'pobjCmd
			Call closeObj(pobjCmd)
		End If
		
		LoadCustomerByEmail = pblnResult

	End Function	'LoadCustomerByEmail

	'***********************************************************************************************

	Public Function LoadCustomerByEmailPassword(byVal strEmail, byVal strPassword)

	Dim pobjCmd
	Dim pobjRS
	Dim pblnResult
	
	'On Error Resume Next

		If len(strEmail) = 0 then
			pstrMessage = "Please enter an Email address"
			pblnError = True
			pblnResult = False
		ElseIf Not fncEmailValid(strEmail) Then
			pstrMessage = "Invalid email address."
			pblnError = True
			pblnResult = False
		Else
			If Not isObject(pobjcnn) Then
				If Not InitializeConnection Then
					pstrMessage = "Unable to connect to database"
					pblnError = True
					LoadCustomerByEmail = False
					Exit Function
				End If
			End If
			
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtext = "Select * From sfCustomers where custEmail=? And custPasswd=?"
				.Commandtype = adCmdText
				'.Commandtype = adCmdStoredProc
				Set .ActiveConnection = pobjcnn

				addParameter pobjCmd, "custEmail", adWChar, strEmail, 100, 2
				addParameter pobjCmd, "custPasswd", adWChar, strPassword, 10, 2
				Set pobjRS = .Execute			
				If pobjRS.EOF Then
					pblnResult = False
				Else
					Call LoadValues(pobjRS)
					pblnResult = True
				End If
				Call closeObj(pobjRS)
			End With	'pobjCmd
			Call closeObj(pobjCmd)
		End If
		
		LoadCustomerByEmailPassword = pblnResult

	End Function	'LoadCustomerByEmailPassword

	'********************************************************************************
	
	Function SetSubscribed(byVal blnSubscribe)

	Dim pstrSQL

	'On Error Resume Next

		If Err.number <> 0 Then Err.Clear

		If Not isObject(pobjcnn) Then
			If Not InitializeConnection Then Exit Function
		End If
		
		If blnSubscribe Then
			pstrSQL = "Update sfCustomers Set custIsSubscribed=1 Where custID=" & wrapSQLValue(plngcustID, False, enDatatype_number)
		Else
			pstrSQL = "Update sfCustomers Set custIsSubscribed=0 Where custID=" & wrapSQLValue(plngcustID, False, enDatatype_number)
		End If
		cnn.Execute pstrSQL,,128
			
		SetSubscribed = CBool(Err.number = 0)

	End Function 'SetSubscribed

	'********************************************************************************
	
	Function ChangePassword(strUsername, strPassword, strNewPassword1, strNewPassword2)

	dim pstrNewPassword1, pstrNewPassword2
	dim pstrSQL, pobjRS

	'On Error Resume Next

		pstrUsername = strUsername
		pstrPassword = strPassword
		pstrNewPassword1 = strNewPassword1
		pstrNewPassword2 = strNewPassword2

		If len(pstrUsername) = 0 then
			ChangePassword = "<h3><b>Please enter a Email</b></h3>"
			Exit Function
		Elseif len(pstrPassword) = 0 then
			ChangePassword = "<h3><b>Please enter your current password</b></h3>"
			Exit Function
		Elseif len(pstrNewPassword1) = 0 then
			ChangePassword = "<h3><b>Please enter the new password</b></h3>"
			Exit Function
		Elseif len(pstrNewPassword2) = 0 then
			ChangePassword = "<h3><b>Please re-type your password</b></h3>"
			Exit Function
		Elseif (pstrNewPassword2 <> pstrNewPassword1) then
			ChangePassword = "<h3><b>The passwords you entered do not match.</b></h3>"
			Exit Function
		ElseIf Not checkInput(pstrUsername, pstrNewPassword1) Then
			ChangePassword = "<h3><b>There was a problem with your login. Please contact the system administrator.</b></h3>"
			Exit Function
		ElseIf Not checkInput(pstrUsername, pstrNewPassword2) Then
			ChangePassword = "<h3><b>There was a problem with your login. Please contact the system administrator.</b></h3>"
			Exit Function
		End If

		If Not isObject(pobjcnn) Then
			If Not InitializeConnection Then Exit Function
		End If
		
		pstrSQL = "Select custID, custEmail, custPasswd, [" & cstrFieldGreeting & "] as ssFieldGreeting from sfCustomers where custEmail=?"
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtext = pstrSQL
			.Commandtype = adCmdText
			'.Commandtype = adCmdStoredProc
			Set .ActiveConnection = pobjcnn

			addParameter pobjCmd, "userName", adWChar, strUsername, 100, 2
			Set pobjRS = .Execute			
			If Err.number = 0 Then
				if pobjRS.eof or pobjRS.bof then
					ChangePassword = "<h3><b>You entered an invalid Email. Please try again.</b></h3>"
				ElseIf trim(pobjRS.Fields("custPasswd").Value) = pstrPassword then
					'pstrSQL = "Update sfCustomers set custPasswd = '" & pstrNewPassword1 & "' where custEmail = '" & pstrUsername & "'"
					'pobjcnn.Execute pstrSQL,,128
					
					pstrSQL = "Update sfCustomers set custPasswd=? where custEmail=?"
					.Commandtext = pstrSQL
					.Parameters.Delete 0
					addParameter pobjCmd, "userPass", adWChar, strNewPassword1, 10, 2
					addParameter pobjCmd, "userName", adWChar, strUsername, 100, 2
					.Execute ,,adExecuteNoRecords

					pstrUserID = Trim(pobjRS.Fields("custID").Value & "")
					pstrPassword = Trim(pobjRS.Fields("custPasswd").Value & "")
					pstrEmail = Trim(pobjRS.Fields("custEmail").Value & "")
					pstrGreeting = Trim(pobjRS.Fields("ssFieldGreeting").Value & "")

					Call SetLoginParameters

					ChangePassword = "Username/Password Successfully Changed"
		'			ChangePassword = "<b>This functionality has been disabled.</b> The Username/Password would have been changed"
				else
					ChangePassword = "<h3><b>You entered an invalid password. Please try again.</b></h3>"
				end if
			Else
				If instr(1,Err.Description,"cannot find the input table or query 'sfCustomers'") <> 0 Then
					ChangePassword = "<div class='FatalError'>You need to upgrade your database to use integrated security</div>"
				Else
					ChangePassword = "<div class='FatalError'>Error: " & Err.number & " - " & Err.Description & "</div>"
				End If
			End If	'Err.number = 0
			
			Call closeObj(pobjRS)
		End With	'pobjCmd
		Call closeObj(pobjCmd)


	End Function 'ChangePassword

	'********************************************************************************
	
	Public Function HasDownloadableItems(byVal lngCustID)
	
	Dim pobjCmd
	Dim pobjRS
	Dim pstrSQL
	
		If Len(CStr(pblnHasDownloadableItems)) > 0 Then
			HasDownloadableItems = pblnHasDownloadableItems
			Exit Function
		End If

		pstrSQL = "SELECT sfOrders.orderID FROM (sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID WHERE sfOrders.orderCustId=? AND sfOrders.orderIsComplete=1 AND sfOrders.orderVoided=0 AND sfProducts.prodFileName<>''"

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, lngCustID)
			Set pobjRS = .Execute
			pblnHasDownloadableItems = Not pobjRS.EOF
			pobjRS.Close
			Set pobjRS = Nothing
		End With	'pobjCmd
		Set pobjCmd = Nothing
	
		HasDownloadableItems = pblnHasDownloadableItems

	End Function	'HasDownloadableItems

	'********************************************************************************
	
	Public Function PriorOrderCount(byVal lngCustID)
	
	Dim pobjCmd
	Dim pobjRS
	Dim pstrSQL
	
		If Len(CStr(plngPriorOrderCount)) > 0 Then
			PriorOrderCount = plngPriorOrderCount
			Exit Function
		End If

		pstrSQL = "SELECT Count(sfOrders.orderID) AS CountOforderID FROM sfOrders WHERE sfOrders.orderCustId=? AND sfOrders.orderIsComplete=1 AND sfOrders.orderVoided=0 GROUP BY sfOrders.orderCustId"

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, lngCustID)
			Set pobjRS = .Execute
			If pobjRS.EOF Then
				plngPriorOrderCount = 0
			Else
				'Returning customer so reset fraud potential back to zero
				plngPriorOrderCount = pobjRS.Fields("CountOforderID").Value
			End If
			pobjRS.Close
			Set pobjRS = Nothing
		End With	'pobjCmd
		Set pobjCmd = Nothing

		PriorOrderCount = plngPriorOrderCount

	End Function	'PriorOrderCount

	'********************************************************************************
	
	Private Function checkInput(byRef strEmail, byRef strPassword)
	'Purpose: protect username (email address) and password from SQL Injection attacks
	
		checkInput = fncEmailValid(strEmail) And validatepassword(strPassword)
	
	End Function	'checkInput

	'********************************************************************************
	
End Class	'clsCustomer

'********************************************************************************

Function getCustomerID(byVal sEmail, byVal sPassword)

Dim plngCustID

	plngCustID = - 1
	Set mclsCustomer = New clsCustomer
	With mclsCustomer
		Set .Connection = cnn
		.custFirstName		= "Saved Cart Customer"	
		.custPasswd			= trim(sPassword)
		.custEmail			= trim(sEmail)
		.custTimesAccessed	= 1
		.custLastAccess		= Date()
		If .AddCustomer Then
			plngCustID = .custID
		End If
	End With
	Set mclsCustomer = Nothing

	getCustomerID = plngCustID
	
End Function	'getCustomerID

'********************************************************************************

Function hasDownloadableItems()

Dim pclsCustomer

	If isLoggedIn Then
		Set pclsCustomer = New clsCustomer
		With pclsCustomer
			Set .Connection = cnn
			If .HasDownloadableItems(VisitorLoggedInCustomerID) = 0 Then
				hasDownloadableItems = False
			Else
				hasDownloadableItems = True
			End If
		End With
		Set pclsCustomer = Nothing
	Else
		hasDownloadableItems = False
	End If

End Function	'hasDownloadableItems

'********************************************************************************

Function hasPriorOrders()

Dim pclsCustomer

	If isLoggedIn Then
		Set pclsCustomer = New clsCustomer
		With pclsCustomer
			Set .Connection = cnn
			If .PriorOrderCount(VisitorLoggedInCustomerID) = 0 Then
				hasPriorOrders = False
			Else
				hasPriorOrders = True
			End If
		End With
		Set pclsCustomer = Nothing
	Else
		hasPriorOrders = False
	End If

End Function	'hasPriorOrders

'********************************************************************************

Dim mclsCustomer
'Set mclsCustomer = New clsCustomer

%>