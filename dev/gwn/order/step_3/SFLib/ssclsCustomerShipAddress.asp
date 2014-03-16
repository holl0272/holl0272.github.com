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

Class clsCustomerShipAddress

Private pobjCnn
Private pblnConnectionOpenedInternally

Private plngcshpaddrID
Private plngcshpaddrCustID
Private pstrcshpaddrFirstName
Private pstrcshpaddrMiddleInitial
Private pstrcshpaddrLastName
Private pstrcshpaddrCompany
Private pstrcshpaddrAddr1
Private pstrcshpaddrAddr2
Private pstrcshpaddrCity
Private pstrcshpaddrState
Private pstrcshpaddrStateName
Private pblnStatePreDefined
Private pstrcshpaddrZip
Private pstrcshpaddrCountry
Private pstrcshpaddrCountryName
Private pstrcshpaddrPhone
Private pstrcshpaddrFax
Private pstrcshpaddrEmail
Private pbytcshpaddrIsActive
Private plngNumAddresses

Private pstrMessage
Private pblnError

	'****************************************************************************************************************

	Private Sub Class_Initialize
	
		pblnConnectionOpenedInternally = False
		pblnError = False

		'Now for the defaults
		pblnStatePreDefined = False
		plngNumAddresses = 0

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
	
	Public Property Let Connection(byRef objCnn)
		If isObject(objCnn) Then Set pobjCnn = objCnn
	End Property

	Public Property Let addressID(byVal value)
		plngcshpaddrID = value
	End Property
	Public Property Get addressID
		addressID = plngcshpaddrID
	End Property
	
	Public Property Let CustID(byVal value)
		plngcshpaddrCustID = value
	End Property
	Public Property Get CustID
		CustID = plngcshpaddrCustID
	End Property
	
	Public Property Let FirstName(byVal value)
		pstrcshpaddrFirstName = value
	End Property
	Public Property Get FirstName
		FirstName = pstrcshpaddrFirstName
	End Property
	
	Public Property Let MiddleInitial(byVal value)
		pstrcshpaddrMiddleInitial = value
	End Property
	Public Property Get MiddleInitial
		MiddleInitial = pstrcshpaddrMiddleInitial
	End Property
	
	Public Property Let LastName(byVal value)
		pstrcshpaddrLastName = value
	End Property
	Public Property Get LastName
		LastName = pstrcshpaddrLastName
	End Property
	
	Public Property Let Company(byVal value)
		pstrcshpaddrCompany = value
	End Property
	Public Property Get Company
		Company = pstrcshpaddrCompany
	End Property
	
	Public Property Let Addr1(byVal value)
		pstrcshpaddrAddr1 = value
	End Property
	Public Property Get Addr1
		Addr1 = pstrcshpaddrAddr1
	End Property
	
	Public Property Let Addr2(byVal value)
		pstrcshpaddrAddr2 = value
	End Property
	Public Property Get Addr2
		Addr2 = pstrcshpaddrAddr2
	End Property
	
	Public Property Let City(byVal value)
		pstrcshpaddrCity = value
	End Property
	Public Property Get City
		City = pstrcshpaddrCity
	End Property

	Public Property Let State(byVal value)
		pstrcshpaddrState = value
	End Property
	Public Property Get State
		State = pstrcshpaddrState
	End Property
	
	Public Property Let Zip(byVal value)
		pstrcshpaddrZip = value
	End Property
	Public Property Get Zip
		Zip = pstrcshpaddrZip
	End Property
	
	Public Property Let Country(byVal value)
		pstrcshpaddrCountry = value
	End Property
	Public Property Get Country
		Country = pstrcshpaddrCountry
	End Property
	
	Public Property Let Phone(byVal value)
		pstrcshpaddrPhone = value
	End Property
	Public Property Get Phone
		Phone = pstrcshpaddrPhone
	End Property
	
	Public Property Let Fax(byVal value)
		pstrcshpaddrFax = value
	End Property
	Public Property Get Fax
		Fax = pstrcshpaddrFax
	End Property
	
	Public Property Let Email(byVal value)
		pstrcshpaddrEmail = value
	End Property
	Public Property Get Email
		Email = pstrcshpaddrEmail
	End Property
	
	Public Property Let IsActive(byVal value)
		pbytcshpaddrIsActive = value
	End Property
	Public Property Get IsActive
		IsActive = pbytcshpaddrIsActive
	End Property

	Public Property Get NumAddresses
		NumAddresses = plngNumAddresses
	End Property
	
	'***********************************************************************************************

	Public Property Get DisplayName
		If Len(pstrcshpaddrMiddleInitial) > 0 Then
			DisplayName = pstrcshpaddrFirstName & " " & pstrcshpaddrMiddleInitial & " " & pstrcshpaddrLastName
		Else
			DisplayName = pstrcshpaddrFirstName & " " & pstrcshpaddrLastName
		End If
	End Property

	'***********************************************************************************************

	Public Property Get countryName
		If Len(pstrcshpaddrCountryName) = 0 And Len(pstrcshpaddrCountry) > 0 Then pstrcshpaddrCountryName = Trim(getNameWithID("sfLocalesCountry", pstrcshpaddrCountry, "loclctryAbbreviation", "loclctryName", 1))
		countryName = pstrcshpaddrCountryName
	End Property

	'***********************************************************************************************

	Public Property Get stateName
		If Len(pstrcshpaddrStateName) = 0 And Len(pstrcshpaddrState) > 0 Then
			pstrcshpaddrStateName = Trim(getNameWithID("sfLocalesState", pstrcshpaddrState, "loclstAbbreviation", "loclstName", 1))
			If Len(pstrcshpaddrStateName) = 0 Then
				pstrcshpaddrStateName = pstrcshpaddrState
				pblnStatePreDefined = False
			Else
				pblnStatePreDefined = True
			End If
		End If
		stateName = pstrcshpaddrStateName
	End Property

	'***********************************************************************************************

	Public Property Get StatePreDefined
		StatePreDefined = pblnStatePreDefined
	End Property

	'***********************************************************************************************

	Private Function ValidateValues
	
	Dim pblnResult

		Call checkLength(pstrcshpaddrFirstName, 50, True, True)
		Call checkLength(pstrcshpaddrMiddleInitial, 1, True, True)
		Call checkLength(pstrcshpaddrLastName, 50, True, True)
		Call checkLength(pstrcshpaddrCompany, 60, True, True)
		Call checkLength(pstrcshpaddrAddr1, 50, True, True)
		Call checkLength(pstrcshpaddrAddr2, 50, True, True)
		Call checkLength(pstrcshpaddrCity, 50, True, True)
		Call checkLength(pstrcshpaddrState, 25, True, True)
		Call checkLength(pstrcshpaddrZip, 12, True, True)
		Call checkLength(pstrcshpaddrCountry, 50, True, True)
		Call checkLength(pstrcshpaddrPhone, 25, True, True)
		Call checkLength(pstrcshpaddrFax, 25, True, True)
		
		If Len(plngcshpaddrID) > 0 And Not isNumeric(plngcshpaddrID) Then plngcshpaddrID = ""
		If Len(plngcshpaddrCustID) > 0 And Not isNumeric(plngcshpaddrCustID) Then plngcshpaddrCustID = ""
		
		If Len(pstrcshpaddrEmail) > 0 Then
			If fncEmailValid(pstrcshpaddrEmail) Then
				pblnResult = checkLength(pstrcshpaddrEmail, 100, False, True)
			Else
				pstrMessage = "Invalid email"
				pblnResult = False
			End If
		Else
			pblnResult = True
		End If

		ValidateValues = pblnResult

	End Function	'ValidateValues

	'***********************************************************************************************

	Private Sub LoadValues(byRef objRS, byVal lngRow)
	
		If isArray(objRS) Then
			plngcshpaddrID = trim(objRS(0, lngRow) & "")
			plngcshpaddrCustID = trim(objRS(1, lngRow) & "")
			pstrcshpaddrFirstName = trim(objRS(2, lngRow) & "")
			pstrcshpaddrMiddleInitial = trim(objRS(3, lngRow) & "")
			pstrcshpaddrLastName = trim(objRS(4, lngRow) & "")
			pstrcshpaddrCompany = trim(objRS(5, lngRow) & "")
			pstrcshpaddrAddr1 = trim(objRS(6, lngRow) & "")
			pstrcshpaddrAddr2 = trim(objRS(7, lngRow) & "")
			pstrcshpaddrCity = trim(objRS(8, lngRow) & "")
			pstrcshpaddrState = trim(objRS(9, lngRow) & "")
			pstrcshpaddrZip = trim(objRS(10, lngRow) & "")
			pstrcshpaddrCountry = trim(objRS(11, lngRow) & "")
			pstrcshpaddrPhone = trim(objRS(12, lngRow) & "")
			pstrcshpaddrFax = trim(objRS(13, lngRow) & "")
			pstrcshpaddrEmail = trim(objRS(14, lngRow) & "")
			pbytcshpaddrIsActive = trim(objRS(15, lngRow) & "")
		ElseIf isObject(objRS) Then
			With objRS
				If Not .EOF Then
					plngcshpaddrID = trim(.Fields("cshpaddrID").Value & "")
					plngcshpaddrCustID = trim(.Fields("cshpaddrCustID").Value & "")
					pstrcshpaddrFirstName = trim(.Fields("cshpaddrShipFirstName").Value & "")
					pstrcshpaddrMiddleInitial = trim(.Fields("cshpaddrShipMiddleInitial").Value & "")
					pstrcshpaddrLastName = trim(.Fields("cshpaddrShipLastName").Value & "")
					pstrcshpaddrCompany = trim(.Fields("cshpaddrShipCompany").Value & "")
					pstrcshpaddrAddr1 = trim(.Fields("cshpaddrShipAddr1").Value & "")
					pstrcshpaddrAddr2 = trim(.Fields("cshpaddrShipAddr2").Value & "")
					pstrcshpaddrCity = trim(.Fields("cshpaddrShipCity").Value & "")
					pstrcshpaddrState = trim(.Fields("cshpaddrShipState").Value & "")
					pstrcshpaddrZip = trim(.Fields("cshpaddrShipZip").Value & "")
					pstrcshpaddrCountry = trim(.Fields("cshpaddrShipCountry").Value & "")
					pstrcshpaddrPhone = trim(.Fields("cshpaddrShipPhone").Value & "")
					pstrcshpaddrFax = trim(.Fields("cshpaddrShipFax").Value & "")
					pstrcshpaddrEmail = trim(.Fields("cshpaddrShipEmail").Value & "")
					pbytcshpaddrIsActive = trim(.Fields("cshpaddrIsActive").Value & "")
				End If
			End With
		End If
		
	End Sub	'LoadValues

	'***********************************************************************************************

	Private Function CreateEmptyAddress

	Dim pblnResult
	Dim pobjCmd
	Dim pobjRS
	Dim pstrSQL
	Dim pstrTempID
	
	'On Error Resume Next

		pstrTempID = SessionID & CStr(Now())

		pstrSQL = "Insert Into sfCShipAddresses (cshpaddrShipEmail) Values (?)"

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtext = pstrSQL
			.Commandtype = adCmdText
			'.Commandtype = adCmdStoredProc
			Set .ActiveConnection = pobjcnn

			addParameter pobjCmd, "cshpaddrShipEmail", adWChar, pstrTempID, 100, 2
			.Execute , , adExecuteNoRecords
		
			'plngcshpaddrID = .Parameters("cshpaddrID")
			
			pstrSQL = "Select cshpaddrID From sfCShipAddresses Where cshpaddrShipEmail=?"
			.Commandtext = pstrSQL
			Set pobjRS = .Execute			
			If pobjRS.EOF Then
				plngcshpaddrID = -1
				pblnResult = False
			Else
				plngcshpaddrID = pobjRS.Fields("cshpaddrID").Value
				pblnResult = True
			End If
			Call closeObj(pobjRS)

		End With	'pobjCmd
		Call closeObj(pobjCmd)

		CreateEmptyAddress = pblnResult

	End Function	'CreateEmptyAddress

	'***********************************************************************************************

	Public Function addAddress

	Dim pblnAlreadyExists
	Dim pblnResult
	Dim pobjCmd
	Dim pstrSQL
	
	'On Error Resume Next

		If Not ValidateValues Then
			addAddress = False
			Exit Function
		End If
		
		'See if customer exists
		If len(plngcshpaddrID) > 0 Then
			pblnAlreadyExists = LoadAddress(plngcshpaddrID)
		Else
			pblnAlreadyExists = False
		End If
		
		If pblnError Then
			pblnResult = False
		Else
			If Not pblnAlreadyExists Then pblnAlreadyExists = CreateEmptyAddress 
			If pblnAlreadyExists Then
				'debugprint "pstrSQL", pstrSQL

				pstrSQL = "Update sfCShipAddresses Set " _
						& "cshpaddrCustID=?," _
						& "cshpaddrShipFirstName=?," _
						& "cshpaddrShipMiddleInitial=?," _
						& "cshpaddrShipLastName=?," _
						& "cshpaddrShipCompany=?," _
						& "cshpaddrShipAddr1=?," _
						& "cshpaddrShipAddr2=?," _
						& "cshpaddrShipCity=?," _
						& "cshpaddrShipState=?," _
						& "cshpaddrShipZip=?," _
						& "cshpaddrShipCountry=?," _
						& "cshpaddrShipPhone=?," _
						& "cshpaddrShipFax=?," _
						& "cshpaddrShipEmail=?" _
						& " Where cshpaddrID=?"

				Set pobjCmd  = CreateObject("ADODB.Command")
				With pobjCmd
					.Commandtext = pstrSQL
					.Commandtype = adCmdText
					'.Commandtype = adCmdStoredProc
					Set .ActiveConnection = pobjcnn

					.Parameters.Append .CreateParameter("cshpaddrCustID", adInteger, adParamInput, 4, plngcshpaddrCustID)
					addParameter pobjCmd, "cshpaddrShipFirstName", adWChar, pstrcshpaddrFirstName, 100, 2
					addParameter pobjCmd, "cshpaddrShipMiddleInitial", adWChar, pstrcshpaddrMiddleInitial, 1, 2
					addParameter pobjCmd, "cshpaddrShipLastName", adWChar, pstrcshpaddrLastName, 50, 2
					addParameter pobjCmd, "cshpaddrShipCompany", adWChar, pstrcshpaddrCompany, 60, 2
					addParameter pobjCmd, "cshpaddrShipAddr1", adWChar, pstrcshpaddrAddr1, 50, 2
					addParameter pobjCmd, "cshpaddrShipAddr2", adWChar, pstrcshpaddrAddr2, 50, 2
					addParameter pobjCmd, "cshpaddrShipCity", adWChar, pstrcshpaddrCity, 50, 2
					addParameter pobjCmd, "cshpaddrShipState", adWChar, pstrcshpaddrState, 25, 2
					addParameter pobjCmd, "cshpaddrShipZip", adWChar, pstrcshpaddrZip, 12, 2
					addParameter pobjCmd, "cshpaddrShipCountry", adWChar, pstrcshpaddrCountry, 50, 2
					addParameter pobjCmd, "cshpaddrShipPhone", adWChar, pstrcshpaddrPhone, 25, 2
					addParameter pobjCmd, "cshpaddrShipFax", adWChar, pstrcshpaddrFax, 25, 2
					addParameter pobjCmd, "cshpaddrShipEmail", adWChar, pstrcshpaddrEmail, 100, 2

					.Parameters.Append .CreateParameter("cshpaddrID", adInteger, adParamInput, 4, plngcshpaddrID)

					.Execute , , adExecuteNoRecords
					pblnResult = CBool(Err.number = 0)
				
				End With	'pobjCmd
				Call closeObj(pobjCmd)

			Else
				pblnResult = False
			End If
		End If

		addAddress = pblnResult

	End Function	'addAddress

	'***********************************************************************************************

	Public Function Update

	Dim pblnResult
	Dim pobjCmd
	Dim pstrSQL
	
	'On Error Resume Next

		If ValidateValues Then
			pstrSQL = "Update sfCShipAddresses Set " _
					& "cshpaddrShipFirstName=?," _
					& "cshpaddrShipMiddleInitial=?," _
					& "cshpaddrShipLastName=?," _
					& "cshpaddrShipCompany=?," _
					& "cshpaddrShipAddr1=?," _
					& "cshpaddrShipAddr2=?," _
					& "cshpaddrShipCity=?," _
					& "cshpaddrShipState=?," _
					& "cshpaddrShipZip=?," _
					& "cshpaddrShipCountry=?," _
					& "cshpaddrShipPhone=?," _
					& "cshpaddrShipFax=?," _
					& "cshpaddrShipEmail=?" _
					& " Where cshpaddrID=?"

			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtext = pstrSQL
				.Commandtype = adCmdText
				'.Commandtype = adCmdStoredProc
				Set .ActiveConnection = pobjcnn

				addParameter pobjCmd, "cshpaddrShipFirstName", adWChar, pstrcshpaddrFirstName, 100, 2
				addParameter pobjCmd, "cshpaddrShipMiddleInitial", adWChar, pstrcshpaddrMiddleInitial, 1, 2
				addParameter pobjCmd, "cshpaddrShipLastName", adWChar, pstrcshpaddrLastName, 50, 2
				addParameter pobjCmd, "cshpaddrShipCompany", adWChar, pstrcshpaddrCompany, 60, 2
				addParameter pobjCmd, "cshpaddrShipAddr1", adWChar, pstrcshpaddrAddr1, 50, 2
				addParameter pobjCmd, "cshpaddrShipAddr2", adWChar, pstrcshpaddrAddr2, 50, 2
				addParameter pobjCmd, "cshpaddrShipCity", adWChar, pstrcshpaddrCity, 50, 2
				addParameter pobjCmd, "cshpaddrShipState", adWChar, pstrcshpaddrState, 25, 2
				addParameter pobjCmd, "cshpaddrShipZip", adWChar, pstrcshpaddrZip, 12, 2
				addParameter pobjCmd, "cshpaddrShipCountry", adWChar, pstrcshpaddrCountry, 50, 2
				addParameter pobjCmd, "cshpaddrShipPhone", adWChar, pstrcshpaddrPhone, 25, 2
				addParameter pobjCmd, "cshpaddrShipFax", adWChar, pstrcshpaddrFax, 25, 2
				addParameter pobjCmd, "cshpaddrShipEmail", adWChar, pstrcshpaddrEmail, 100, 2

				.Parameters.Append .CreateParameter("cshpaddrID", adInteger, adParamInput, 4, plngcshpaddrID)

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

	'********************************************************************************

	Public Function LoadAddress(byVal lngcshpaddrID)

	Dim pblnResult
	Dim pobjCmd
	Dim pobjRS
	Dim pstrSQL
	
	'On Error Resume Next

		If validNumber(lngcshpaddrID) Then
			If Not isObject(pobjcnn) Then
				If Not InitializeConnection Then Exit Function
			End If
			
			pstrSQL = "Select * From sfCShipAddresses where cshpaddrID=?"
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtext = pstrSQL
				.Commandtype = adCmdText
				'.Commandtype = adCmdStoredProc
				Set .ActiveConnection = pobjcnn

				.Parameters.Append .CreateParameter("cshpaddrID", adInteger, adParamInput, 4, lngcshpaddrID)
				Set pobjRS = .Execute			
				If pobjRS.EOF Then
					pblnResult = False
				Else
					Call LoadValues(pobjRS.GetRows(), 0)
					pblnResult = True
				End If
				Call closeObj(pobjRS)

			End With	'pobjCmd
			Call closeObj(pobjCmd)
		Else
			pstrMessage = "Invalid Customer ID"
			pblnError = True
			pblnResult = False
		End If	'validNumber(lngcshpaddrID)

		LoadAddress = pblnResult

	End Function	'LoadAddress

	'********************************************************************************

	Public Function getPriorShippingAddresses(byVal lngcshpaddrID, byVal lngCustID, byRef aryPriorShippingAddresses)

	Dim i
	Dim pblnResult
	Dim pobjCmd
	Dim pobjRS
	Dim pstrSQL
	
	'0 - cshpaddrID
	'1 - cshpaddrCustID
	'2 - cshpaddrShipFirstName
	'3 - cshpaddrShipMiddleInitial
	'4 - cshpaddrShipLastName
	'5 - cshpaddrShipCompany
	'6 - cshpaddrShipAddr1
	'7 - cshpaddrShipAddr2
	'8 - cshpaddrShipCity
	'9 - cshpaddrShipState
	'10 - cshpaddrShipZip
	'11 - cshpaddrShipCountry
	'12 - cshpaddrShipPhone
	'13 - cshpaddrShipFax
	'14 - cshpaddrShipEmail
	'15 - cshpaddrIsActive

	'On Error Resume Next

		If validNumber(lngCustID) Then
			If Not isObject(pobjcnn) Then
				If Not InitializeConnection Then Exit Function
			End If
			
			pstrSQL = "Select cshpaddrID, cshpaddrCustID, cshpaddrShipFirstName, cshpaddrShipMiddleInitial, cshpaddrShipLastName, cshpaddrShipCompany, cshpaddrShipAddr1, cshpaddrShipAddr2, cshpaddrShipCity, cshpaddrShipState, cshpaddrShipZip, cshpaddrShipCountry, cshpaddrShipPhone, cshpaddrShipFax, cshpaddrShipEmail, cshpaddrIsActive From sfCShipAddresses where cshpaddrCustID=? Order By cshpaddrShipLastName, cshpaddrShipFirstName, cshpaddrShipCity"

			pstrSQL = "SELECT Count(sfCShipAddresses.cshpaddrID) AS CountOfcshpaddrID, sfCShipAddresses.cshpaddrCustID, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipMiddleInitial, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipPhone, sfCShipAddresses.cshpaddrShipFax, sfCShipAddresses.cshpaddrShipEmail, sfCShipAddresses.cshpaddrIsActive" _
					& " FROM sfCShipAddresses" _
					& " GROUP BY sfCShipAddresses.cshpaddrCustID, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipMiddleInitial, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipPhone, sfCShipAddresses.cshpaddrShipFax, sfCShipAddresses.cshpaddrShipEmail, sfCShipAddresses.cshpaddrIsActive, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipCity" _
					& " HAVING sfCShipAddresses.cshpaddrCustID=?" _
					& " ORDER BY sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipCity"

			pstrSQL = "SELECT Count(sfCShipAddresses.cshpaddrID) AS CountOfcshpaddrID, sfCShipAddresses.cshpaddrCustID, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipMiddleInitial, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipPhone, sfCShipAddresses.cshpaddrShipFax, sfCShipAddresses.cshpaddrShipEmail, sfCShipAddresses.cshpaddrIsActive" _
					& " FROM sfCShipAddresses" _
					& " WHERE sfCShipAddresses.cshpaddrCustID=?" _
					& " GROUP BY sfCShipAddresses.cshpaddrCustID, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipMiddleInitial, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipPhone, sfCShipAddresses.cshpaddrShipFax, sfCShipAddresses.cshpaddrShipEmail, sfCShipAddresses.cshpaddrIsActive, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipCity" _
					& " ORDER BY sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipCity"

			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtext = pstrSQL
				.Commandtype = adCmdText
				'.Commandtype = adCmdStoredProc
				Set .ActiveConnection = pobjcnn

				.Parameters.Append .CreateParameter("cshpaddrCustID", adInteger, adParamInput, 4, lngCustID)
				Set pobjRS = .Execute			
				If pobjRS.EOF Then
					pblnResult = False
				Else
  					aryPriorShippingAddresses = pobjRS.GetRows()
					plngNumAddresses = UBound(aryPriorShippingAddresses, 2)

					If Len(lngcshpaddrID) > 0 And isNumeric(lngcshpaddrID) And CStr(lngcshpaddrID)<>"0" Then
  						For i = 0 to plngNumAddresses
							If aryPriorShippingAddresses(0,i) = CLng(lngcshpaddrID) Then
								Call LoadValues(aryPriorShippingAddresses, i)
								Exit For
							End If
  						Next
					Else
  						For i = 0 to plngNumAddresses
							If aryPriorShippingAddresses(15,i) = 1 Then
								Call LoadValues(aryPriorShippingAddresses, i)
								Exit For
							End If
  						Next
					End If

					pblnResult = True
					
				End If

			End With	'pobjCmd
			Call closeObj(pobjCmd)
		Else
			pstrMessage = "Invalid Customer ID"
			pblnError = True
			pblnResult = False
		End If	'validNumber(lngCustID)

		getPriorShippingAddresses = pblnResult

	End Function	'getPriorShippingAddresses

	'********************************************************************************
	
	Private Function checkInput(byRef strEmail, byRef strPassword)
	'Purpose: protect username (email address) and password from SQL Injection attacks
	
		checkInput = fncEmailValid(strEmail) And validatepassword(strPassword)
	
	End Function	'checkInput

	'********************************************************************************
	
End Class	'clsCustomerShipAddress

Dim mclsCustomerShipAddress
'Set mclsCustomerShipAddress = New clsCustomerShipAddress
%>