<% Option Explicit 
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version:	2.00.002		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		March 7, 2004											*
'*                                                                              *
'*   Release 2.00.002 (September 5, 2003)										*
'*	   - Updated link to point to new carrier admin page						*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clsAdmin
	'Assumptions:
	'   cnn: defines a previously opened connection to the database

	'class variables
	Private cstrDelimeter
	Private pstrMessage
	Private pblnError
	Private pRS
	'database variables

Private pstradminDomainName
Private pstradminStoreName
Private pstradminSSLPath
Private pblnadminEzeeActive
Private pstradminEzeeLogin
Private pblnadminSFActive
Private pstradminSFID
Private pblnadminSaveCartActive
Private pblnadminEmailActive

Private pstradminMailMethod
Private pstradminMailServer
Private pstradminMailServerUserName
Private pstradminMailServerPassword

Private pstradminPrimaryEmail
Private pstradminSecondaryEmail
Private pstradminEmailMessage
Private pstradminEmailSubject
Private pblnadminSubscribeMailIsActive

Private pstradminTransMethod
Private pstradminMerchantType
Private pstradminDeletePolicy
Private pintadminDeleteSchedule
Private pblnadminEncodeCCIsActive
Private pstradminPaymentServer
Private pstradminLogin
Private pstradminPassword

Private pstradminLCID
Private pstradminOriginCountry
Private pstradminOriginState
Private pstradminOriginZip
Private pstradminOandaID
Private pblnadminActivateOanda

Private pintadminShipType
Private pintadminShipType2
Private pstradminSpcShipAmt
Private pblnadminPrmShipIsActive
Private pstradminCODAmount
Private pstradminHandling
Private pbytadminHandlingType
Private pblnadminHandlingIsActive
Private pstradminShipMin
Private pblnadminTaxShipIsActive
Private pdbladminFreeShippingAmount
Private pblnadminFreeShippingIsActive

Private pdbladminGlobalSaleAmt
Private pblnadminGlobalSaleIsActive
Private pbytadminUpdDesign
Private pbytadminUpdList
Private pbytadminUpdTxt

Private pblnadminGlobalConfirmationMessageIsactive
Private pstradminGlobalConfirmationMessage
Private pstrhoursBetweenCleanings
Private pstrdaysToSaveIncompleteOrders
Private pstrdaysToSaveTempOrders
Private pstrdaysToSaveSavedOrders
Private pstrdaysToKeepVisitors
Private pstrCCReplace

Private pstradminTechnicalEmail
Private pstradminTermsAndConditions
Private pblnadminTermsAndConditionsIsactive
Private pdbladminMinOrderAmount
Private pstradminMinOrderMessage
Private plngStoreID

	'***********************************************************************************************

	Private Sub class_Initialize()
		cstrDelimeter = ";"
		plngStoreID = 1
	End Sub

	Private Sub class_Terminate()
		On Error Resume Next
		pRS.Close
		set pRS = nothing
	End Sub

	'***********************************************************************************************

	Public Property Let Recordset(oRS)
	    set pRS = oRS
	End Property

	Public Property Get Recordset()
	    set Recordset = pRS
	End Property


	Public Property Get Message()
	    Message = pstrMessage
	End Property

	Public Sub OutputMessage()

	Dim i
	Dim aError

	    aError = Split(pstrMessage, cstrDelimeter)
	    For i = 0 To UBound(aError)
	        If pblnError Then
	            Response.Write "<P align='center'><H4><FONT color=Red>" & aError(i) & "</FONT></H4></P>"
	        Else
	            Response.Write "<P align='center'><H4>" & aError(i) & "</H4></P>"
	        End If
	    Next 'i

	End Sub 'OutputMessage

	'***********************************************************************************************

	Public Property Get adminActivateOanda()
	    adminActivateOanda = pblnadminActivateOanda
	End Property

	Public Property Get adminCODAmount()
	    adminCODAmount = pstradminCODAmount
	End Property

	Public Property Get adminDeletePolicy()
	    adminDeletePolicy = pstradminDeletePolicy
	End Property

	Public Property Get adminDeleteSchedule()
	    adminDeleteSchedule = pintadminDeleteSchedule
	End Property

	Public Property Get adminDomainName()
	    adminDomainName = pstradminDomainName
	End Property

	Public Property Get adminEmailActive()
	    adminEmailActive = pblnadminEmailActive
	End Property

	Public Property Get adminEmailMessage()
	    adminEmailMessage = pstradminEmailMessage
	End Property

	Public Property Get adminEmailSubject()
	    adminEmailSubject = pstradminEmailSubject
	End Property

	Public Property Get adminEncodeCCIsActive()
	    adminEncodeCCIsActive = pblnadminEncodeCCIsActive
	End Property

	Public Property Get adminEzeeActive()
	    adminEzeeActive = pblnadminEzeeActive
	End Property

	Public Property Get adminEzeeLogin()
	    adminEzeeLogin = pstradminEzeeLogin
	End Property

	Public Property Get adminFreeShippingAmount()
	    adminFreeShippingAmount = pdbladminFreeShippingAmount
	End Property

	Public Property Get adminFreeShippingIsActive()
	    adminFreeShippingIsActive = pblnadminFreeShippingIsActive
	End Property

	Public Property Get adminGlobalSaleAmt()
	    adminGlobalSaleAmt = pdbladminGlobalSaleAmt
	End Property

	Public Property Get adminGlobalSaleIsActive()
	    adminGlobalSaleIsActive = pblnadminGlobalSaleIsActive
	End Property

	Public Property Get adminHandling()
	    adminHandling = pstradminHandling
	End Property

	Public Property Get hoursBetweenCleanings()
	    hoursBetweenCleanings = pstrhoursBetweenCleanings
	End Property

	Public Property Get daysToSaveIncompleteOrders()
	    daysToSaveIncompleteOrders = pstrdaysToSaveIncompleteOrders
	End Property

	Public Property Get daysToSaveTempOrders()
	    daysToSaveTempOrders = pstrdaysToSaveTempOrders
	End Property

	Public Property Get daysToSaveSavedOrders()
	    daysToSaveSavedOrders = pstrdaysToSaveSavedOrders
	End Property

	Public Property Get daysToKeepVisitors()
	    daysToKeepVisitors = pstrdaysToKeepVisitors
	End Property

	Public Property Get CCReplace()
	    CCReplace = pstrCCReplace
	End Property

	Public Property Get adminTechnicalEmail()
	    adminTechnicalEmail = pstradminTechnicalEmail
	End Property

	Public Property Get adminTermsAndConditions()
	    adminTermsAndConditions = pstradminTermsAndConditions
	End Property

	Public Property Get adminTermsAndConditionsIsactive()
	    adminTermsAndConditionsIsactive = pblnadminTermsAndConditionsIsactive
	End Property

	Public Property Get adminMinOrderAmount()
	    adminMinOrderAmount = pdbladminMinOrderAmount
	End Property

	Public Property Get adminMinOrderMessage()
	    adminMinOrderMessage = pstradminMinOrderMessage
	End Property

	Public Property Get adminGlobalConfirmationMessageIsactive()
	    adminGlobalConfirmationMessageIsactive = pblnadminGlobalConfirmationMessageIsactive
	End Property

	Public Property Get adminGlobalConfirmationMessage()
	    adminGlobalConfirmationMessage = pstradminGlobalConfirmationMessage
	End Property

	Public Property Get adminHandlingIsActive()
	    adminHandlingIsActive = pblnadminHandlingIsActive
	End Property

	Public Property Get adminHandlingType()
	    adminHandlingType = pbytadminHandlingType
	End Property

	Public Property Get adminLCID()
	    adminLCID = pstradminLCID
	End Property

	Public Property Get adminLogin()
	    adminLogin = pstradminLogin
	End Property

	Public Property Get adminMailMethod()
	    adminMailMethod = pstradminMailMethod
	End Property

	Public Property Get adminMailServer()
	    adminMailServer = pstradminMailServer
	End Property

	Public Property Get adminMailServerUserName()
	    adminMailServerUserName = pstradminMailServerUserName
	End Property

	Public Property Get adminMailServerPassword()
	    adminMailServerPassword = pstradminMailServerPassword
	End Property

	Public Property Get adminMerchantType()
	    adminMerchantType = pstradminMerchantType
	End Property

	Public Property Get adminOandaID()
	    adminOandaID = pstradminOandaID
	End Property

	Public Property Get adminOriginCountry()
	    adminOriginCountry = pstradminOriginCountry
	End Property

	Public Property Get adminOriginState()
	    adminOriginState = pstradminOriginState
	End Property

	Public Property Get adminOriginZip()
	    adminOriginZip = pstradminOriginZip
	End Property

	Public Property Get adminPassword()
	    adminPassword = pstradminPassword
	End Property

	Public Property Get adminPaymentServer()
	    adminPaymentServer = pstradminPaymentServer
	End Property

	Public Property Get adminPrimaryEmail()
	    adminPrimaryEmail = pstradminPrimaryEmail
	End Property

	Public Property Get adminPrmShipIsActive()
	    adminPrmShipIsActive = pblnadminPrmShipIsActive
	End Property

	Public Property Get adminSaveCartActive()
	    adminSaveCartActive = pblnadminSaveCartActive
	End Property

	Public Property Get adminSecondaryEmail()
	    adminSecondaryEmail = pstradminSecondaryEmail
	End Property

	Public Property Get adminSFActive()
	    adminSFActive = pblnadminSFActive
	End Property

	Public Property Get adminSFID()
	    adminSFID = pstradminSFID
	End Property

	Public Property Get adminShipMin()
	    adminShipMin = pstradminShipMin
	End Property

	Public Property Get adminShipType()
	    adminShipType = pintadminShipType
	End Property

	Public Property Get adminShipType2()
	    adminShipType2 = pintadminShipType2
	End Property

	Public Property Get adminSpcShipAmt()
	    adminSpcShipAmt = pstradminSpcShipAmt
	End Property

	Public Property Get adminSSLPath()
	    adminSSLPath = pstradminSSLPath
	End Property

	Public Property Get adminStoreName()
	    adminStoreName = pstradminStoreName
	End Property

	Public Property Get adminSubscribeMailIsActive()
	    adminSubscribeMailIsActive = pblnadminSubscribeMailIsActive
	End Property

	Public Property Get adminTaxShipIsActive()
	    adminTaxShipIsActive = pblnadminTaxShipIsActive
	End Property

	Public Property Get adminTransMethod()
	    adminTransMethod = pstradminTransMethod
	End Property

	Public Property Get adminUpdDesign()
	    adminUpdDesign = pbytadminUpdDesign
	End Property

	Public Property Get adminUpdList()
	    adminUpdList = pbytadminUpdList
	End Property

	Public Property Get adminUpdTxt()
	    adminUpdTxt = pbytadminUpdTxt
	End Property

	Public Property Let StoreID(byVal lngValue)
	    plngStoreID = lngValue
	End Property
	Public Property Get StoreID()
	    StoreID = plngStoreID
	End Property

	'***********************************************************************************************

	Private Sub LoadValues(rs)

		pblnadminActivateOanda = trim(rs("adminActivateOanda"))
		pstradminCODAmount = trim(rs("adminCODAmount"))
		pstradminDeletePolicy = trim(rs("adminDeletePolicy"))
		pintadminDeleteSchedule = trim(rs("adminDeleteSchedule"))
		pstradminDomainName = trim(rs("adminDomainName"))
		pblnadminEmailActive = trim(rs("adminEmailActive"))
		pstradminEmailMessage = trim(rs("adminEmailMessage"))
		pstradminEmailSubject = trim(rs("adminEmailSubject"))
		pblnadminEncodeCCIsActive = trim(rs("adminEncodeCCIsActive"))
		pblnadminEzeeActive = trim(rs("adminEzeeActive"))
		pstradminEzeeLogin = trim(rs("adminEzeeLogin"))
		pdbladminFreeShippingAmount = trim(rs("adminFreeShippingAmount"))
		pblnadminFreeShippingIsActive = trim(rs("adminFreeShippingIsActive"))
		pdbladminGlobalSaleAmt = trim(rs("adminGlobalSaleAmt"))
		pblnadminGlobalSaleIsActive = trim(rs("adminGlobalSaleIsActive"))

		pblnadminGlobalConfirmationMessageIsactive = trim(rs("adminGlobalConfirmationMessageIsactive"))
		pstradminGlobalConfirmationMessage = trim(rs("adminGlobalConfirmationMessage"))

		pstrhoursBetweenCleanings = trim(rs.Fields("hoursBetweenCleanings").Value)
		pstrdaysToSaveIncompleteOrders = trim(rs.Fields("daysToSaveIncompleteOrders").Value)
		pstrdaysToSaveTempOrders = trim(rs.Fields("daysToSaveTempOrders").Value)
		pstrdaysToSaveSavedOrders = trim(rs.Fields("daysToSaveSavedOrders").Value)
		pstrdaysToKeepVisitors = trim(rs.Fields("daysToKeepVisitors").Value)
		pstrCCReplace = trim(rs.Fields("CCReplace").Value)

		pstradminTechnicalEmail = trim(rs.Fields("adminTechnicalEmail").Value)
		pstradminTermsAndConditions = trim(rs.Fields("adminTermsAndConditions").Value)
		pblnadminTermsAndConditionsIsactive = trim(rs.Fields("adminTermsAndConditionsIsactive").Value)
		pdbladminMinOrderAmount = trim(rs.Fields("adminMinOrderAmount").Value)
		pstradminMinOrderMessage = trim(rs.Fields("adminMinOrderMessage").Value)

		pstradminHandling = trim(rs("adminHandling"))
		pblnadminHandlingIsActive = trim(rs("adminHandlingIsActive"))
		pbytadminHandlingType = trim(rs("adminHandlingType"))
		pstradminLCID = trim(rs("adminLCID"))
		pstradminLogin = trim(rs("adminLogin"))
		pstradminMailMethod = trim(rs("adminMailMethod"))
		pstradminMailServer = trim(rs("adminMailServer"))
		pstradminMerchantType = trim(rs("adminMerchantType"))
		pstradminOandaID = trim(rs("adminOandaID"))
		pstradminOriginCountry = trim(rs("adminOriginCountry"))
		pstradminOriginState = trim(rs("adminOriginState"))
		pstradminOriginZip = trim(rs("adminOriginZip"))
		pstradminPassword = trim(rs("adminPassword"))
		pstradminPaymentServer = trim(rs("adminPaymentServer"))
		pstradminPrimaryEmail = trim(rs("adminPrimaryEmail"))
		pblnadminPrmShipIsActive = trim(rs("adminPrmShipIsActive"))
		pblnadminSaveCartActive = trim(rs("adminSaveCartActive"))
		pstradminSecondaryEmail = trim(rs("adminSecondaryEmail"))
		pblnadminSFActive = trim(rs("adminSFActive"))
		pstradminSFID = trim(rs("adminSFID"))
		pstradminShipMin = trim(rs("adminShipMin"))
		pintadminShipType = trim(rs("adminShipType"))
		pintadminShipType2 = trim(rs("adminShipType2"))
		pstradminSpcShipAmt = trim(rs("adminSpcShipAmt"))
		pstradminSSLPath = trim(rs("adminSSLPath"))
		pstradminStoreName = trim(rs("adminStoreName"))
		pblnadminSubscribeMailIsActive = trim(rs("adminSubscribeMailIsActive"))
		pblnadminTaxShipIsActive = trim(rs("adminTaxShipIsActive"))
		pstradminTransMethod = trim(rs("adminTransMethod"))
		pbytadminUpdDesign = trim(rs("adminUpdDesign"))
		pbytadminUpdList = trim(rs("adminUpdList"))
		pbytadminUpdTxt = trim(rs("adminUpdTxt"))

		'added to support mail server login
		Dim paryMailServer
		If InStr(1, pstradminMailServer, "|") > 1 Then
			paryMailServer = Split(pstradminMailServer, "|")
			pstradminMailServer = paryMailServer(0)
			pstradminMailServerUserName = paryMailServer(1)
			pstradminMailServerPassword = paryMailServer(2)
		End If
	    
	End Sub 'LoadValues

	'***********************************************************************************************

	Private Sub LoadFromRequest

	    With Request.Form
			pblnadminActivateOanda = (UCase(.Item("adminActivateOanda")) = "ON")
			pstradminCODAmount = Trim(.Item("adminCODAmount"))
			pstradminDeletePolicy = Trim(.Item("adminDeletePolicy"))
			pintadminDeleteSchedule = Trim(.Item("adminDeleteSchedule"))
			pstradminDomainName = Trim(.Item("adminDomainName"))
			pblnadminEmailActive = (UCase(.Item("adminEmailActive")) = "ON")
			pstradminEmailMessage = Trim(.Item("adminEmailMessage"))
			pstradminEmailSubject = Trim(.Item("adminEmailSubject"))
			pblnadminEncodeCCIsActive = (UCase(.Item("adminEncodeCCIsActive")) = "ON")
			pblnadminEzeeActive = (UCase(.Item("adminEzeeActive")) = "ON")
			pstradminEzeeLogin = Trim(.Item("adminEzeeLogin"))
			pdbladminFreeShippingAmount = Trim(.Item("adminFreeShippingAmount"))
			pblnadminFreeShippingIsActive = (UCase(.Item("adminFreeShippingIsActive")) = "ON")
			pdbladminGlobalSaleAmt = Trim(.Item("adminGlobalSaleAmt"))
			pblnadminGlobalSaleIsActive = (UCase(.Item("adminGlobalSaleIsActive")) = "ON")

			pstradminGlobalConfirmationMessage = Trim(.Item("adminGlobalConfirmationMessage"))
			pblnadminGlobalConfirmationMessageIsactive = (UCase(.Item("adminGlobalConfirmationMessageIsactive")) = "ON")

			pstrhoursBetweenCleanings = correctEmptyValue(Trim(.Item("hoursBetweenCleanings")), 24)
			pstrdaysToSaveIncompleteOrders = correctEmptyValue(Trim(.Item("daysToSaveIncompleteOrders")), 7)
			pstrdaysToSaveTempOrders = correctEmptyValue(Trim(.Item("daysToSaveTempOrders")), 3)
			pstrdaysToSaveSavedOrders = correctEmptyValue(Trim(.Item("daysToSaveSavedOrders")), 30)
			pstrdaysToKeepVisitors = correctEmptyValue(Trim(.Item("daysToKeepVisitors")), 3)
			pstrCCReplace = Trim(.Item("CCReplace"))

			pstradminTechnicalEmail = Trim(.Item("adminTechnicalEmail"))
			pstradminTermsAndConditions = Trim(.Item("adminTermsAndConditions"))
			pdbladminMinOrderAmount = Trim(.Item("adminMinOrderAmount"))
			pstradminMinOrderMessage = Trim(.Item("adminMinOrderMessage"))
			pblnadminTermsAndConditionsIsactive = (UCase(.Item("adminTermsAndConditionsIsactive")) = "ON")

			pstradminHandling = Trim(.Item("adminHandling"))
			'pblnadminHandlingIsActive = (UCase(.Item("adminHandlingIsActive")) = "ON")
			pblnadminHandlingIsActive = .Item("adminHandlingIsActive")
			pbytadminHandlingType = Trim(.Item("adminHandlingType"))
			pstradminLCID = Trim(.Item("adminLCID"))
			pstradminLogin = Trim(.Item("adminLogin"))
			pstradminMailMethod = Trim(.Item("adminMailMethod"))
			pstradminMailServer = Trim(.Item("adminMailServer"))
			pstradminMailServerUserName = Trim(.Item("adminMailServerUserName"))
			pstradminMailServerPassword = Trim(.Item("adminMailServerPassword"))
			
			pstradminMerchantType = Trim(.Item("adminMerchantType"))
			pstradminOandaID = Trim(.Item("adminOandaID"))
			pstradminOriginCountry = Trim(.Item("adminOriginCountry"))
			pstradminOriginState = Trim(.Item("adminOriginState"))
			pstradminOriginZip = Trim(.Item("adminOriginZip"))
			pstradminPassword = Trim(.Item("adminPassword"))
			pstradminPaymentServer = Trim(.Item("adminPaymentServer"))
			pstradminPrimaryEmail = Trim(.Item("adminPrimaryEmail"))
			pblnadminPrmShipIsActive = (UCase(.Item("adminPrmShipIsActive")) = "ON")
			pblnadminSaveCartActive = (UCase(.Item("adminSaveCartActive")) = "ON")
			pstradminSecondaryEmail = Trim(.Item("adminSecondaryEmail"))
			pblnadminSFActive = (UCase(.Item("adminSFActive")) = "ON")
			pstradminSFID = Trim(.Item("adminSFID"))
			pstradminShipMin = Trim(.Item("adminShipMin"))
			pintadminShipType = Trim(.Item("adminShipType"))
			pintadminShipType2 = Trim(.Item("adminShipType2"))
			pstradminSpcShipAmt = Trim(.Item("adminSpcShipAmt"))
			pstradminSSLPath = Trim(.Item("adminSSLPath"))
			pstradminStoreName = Trim(.Item("adminStoreName"))
			pblnadminSubscribeMailIsActive = (UCase(.Item("adminSubscribeMailIsActive")) = "ON")
			pblnadminTaxShipIsActive = (UCase(.Item("adminTaxShipIsActive")) = "ON")
			pstradminTransMethod = Trim(.Item("adminTransMethod"))
			pbytadminUpdDesign = Trim(.Item("adminUpdDesign"))
			pbytadminUpdList = Trim(.Item("adminUpdList"))
			pbytadminUpdTxt = Trim(.Item("adminUpdTxt"))
			
			'Set some defaults
			If Len(pblnadminHandlingIsActive) = 0 Then pblnadminHandlingIsActive = 0
			If Len(pdbladminFreeShippingAmount) = 0 Then pdbladminFreeShippingAmount = 0
			If Len(pstradminShipMin) = 0 Then pstradminShipMin = 0
			If Len(pstradminCODAmount) = 0 Then pstradminCODAmount = 0
			If Len(pstradminSpcShipAmt) = 0 Then pstradminSpcShipAmt = 0
			If Len(pbytadminHandlingType) = 0 Then pbytadminHandlingType = 0
			
	    End With

	End Sub 'LoadFromRequest

	'***********************************************************************************************

	Public Function Load()

	'On Error Resume Next

	    Set pRS = server.CreateObject("adodb.Recordset")
		set pRS = GetRS("Select * from sfAdmin Where adminID=" & plngStoreID)
	    If Not (pRS.EOF Or pRS.BOF) Then
			Call LoadValues(pRS)
			Load = True
		End If

	End Function    'Load

	'***********************************************************************************************

	Public Function deleteStore

		If plngStoreID = 1 Then
			pstrMessage = "This store cannot be deleted."
		Else
			cnn.Execute "Delete From sfAdmin where adminID=" & plngStoreID,,128
			plngStoreID = 1
			If Err.Number = 0 Then
				pstrMessage = "The store configuration was successfully deleted."
			Else
				pstrMessage = "Error: " & Err.Number & " - " & Err.Description
			End If
		End If

	End Function    'deleteStore

	'***********************************************************************************************

	Public Function newStore

	Dim pobjRS

		cnn.Execute "Insert Into sfAdmin (adminStoreName) Values ('NewStoreName')",,128
		
		Set pobjRS = GetRS("Select adminID From sfAdmin Where adminStoreName='NewStoreName'")
		If pobjRS.EOF Then
			newStore = False
            pstrMessage = "Error creating new store."
		Else
			plngStoreID = pobjRS.Fields("adminID").Value
            pstrMessage = "New store created."
		End If

	End Function    'newStore

	'***********************************************************************************************

	Public Function Update

	Dim pblnError
	Dim rs
	Dim i
	
	'On Error Resume Next

    pblnError = False
    Call LoadFromRequest
    If ValidateValues Then

        Set rs = server.CreateObject("adodb.Recordset")
        With rs
			.open "Select * from sfAdmin where adminID=" & plngStoreID, cnn, 1, 3
			If not .EOF Then
				.Fields("adminActivateOanda") = pblnadminActivateOanda * -1
				.Fields("adminCODAmount") = pstradminCODAmount
				.Fields("adminDeletePolicy") = pstradminDeletePolicy
				.Fields("adminDeleteSchedule") = pintadminDeleteSchedule
				.Fields("adminDomainName") = pstradminDomainName
				.Fields("adminEmailActive") = pblnadminEmailActive * -1
				.Fields("adminEmailMessage") = pstradminEmailMessage
				.Fields("adminEmailSubject") = pstradminEmailSubject
				.Fields("adminEncodeCCIsActive") = pblnadminEncodeCCIsActive * -1
				.Fields("adminEzeeActive") = pblnadminEzeeActive * -1
				.Fields("adminEzeeLogin") = pstradminEzeeLogin
				.Fields("adminFreeShippingAmount") = pdbladminFreeShippingAmount
				.Fields("adminFreeShippingIsActive") = pblnadminFreeShippingIsActive * -1
				.Fields("adminGlobalSaleAmt") = pdbladminGlobalSaleAmt
				.Fields("adminGlobalSaleIsActive") = pblnadminGlobalSaleIsActive * -1

				.Fields("adminGlobalConfirmationMessage") = pstradminGlobalConfirmationMessage
				.Fields("adminGlobalConfirmationMessageIsactive") = pblnadminGlobalConfirmationMessageIsactive * -1

				.Fields("hoursBetweenCleanings") = pstrhoursBetweenCleanings
				.Fields("daysToSaveIncompleteOrders") = pstrdaysToSaveIncompleteOrders
				.Fields("daysToSaveTempOrders") = pstrdaysToSaveTempOrders
				.Fields("daysToSaveSavedOrders") = pstrdaysToSaveSavedOrders
				.Fields("daysToKeepVisitors") = pstrdaysToKeepVisitors
				.Fields("CCReplace") = pstrCCReplace

				.Fields("adminTechnicalEmail") = pstradminTechnicalEmail
				.Fields("adminTermsAndConditionsIsactive") = pblnadminTermsAndConditionsIsactive * -1
				.Fields("adminTermsAndConditions") = pstradminTermsAndConditions
				.Fields("adminMinOrderAmount") = pdbladminMinOrderAmount
				.Fields("adminMinOrderMessage") = pstradminMinOrderMessage

				.Fields("adminHandling") = pstradminHandling
				'.Fields("adminHandlingIsActive") = pblnadminHandlingIsActive * -1
				.Fields("adminHandlingIsActive") = pblnadminHandlingIsActive
				.Fields("adminHandlingType") = pbytadminHandlingType
				.Fields("adminLCID") = pstradminLCID
				.Fields("adminLogin") = pstradminLogin
				.Fields("adminMailMethod") = pstradminMailMethod
				.Fields("adminMailServer") = pstradminMailServer & "|" & pstradminMailServerUserName & "|" & pstradminMailServerPassword
				.Fields("adminMerchantType") = pstradminMerchantType
				.Fields("adminOandaID") = pstradminOandaID
				.Fields("adminOriginCountry") = pstradminOriginCountry
				.Fields("adminOriginState") = pstradminOriginState
				.Fields("adminOriginZip") = pstradminOriginZip
				.Fields("adminPassword") = pstradminPassword
				.Fields("adminPaymentServer") = pstradminPaymentServer
				.Fields("adminPrimaryEmail") = pstradminPrimaryEmail
				.Fields("adminPrmShipIsActive") = pblnadminPrmShipIsActive * -1
				.Fields("adminSaveCartActive") = pblnadminSaveCartActive * -1
				.Fields("adminSecondaryEmail") = pstradminSecondaryEmail
				.Fields("adminSFActive") = pblnadminSFActive * -1
				.Fields("adminSFID") = pstradminSFID
				.Fields("adminShipMin") = pstradminShipMin
				.Fields("adminShipType") = pintadminShipType
				.Fields("adminShipType2") = pintadminShipType2
				.Fields("adminSpcShipAmt") = pstradminSpcShipAmt
				.Fields("adminSSLPath") = pstradminSSLPath
				.Fields("adminStoreName") = pstradminStoreName
				.Fields("adminSubscribeMailIsActive") = pblnadminSubscribeMailIsActive * -1
				.Fields("adminTaxShipIsActive") = pblnadminTaxShipIsActive * -1
				.Fields("adminTransMethod") = pstradminTransMethod
'				.Fields("adminUpdDesign") = pbytadminUpdDesign
'				.Fields("adminUpdList") = pbytadminUpdList
'				.Fields("adminUpdTxt") = pbytadminUpdTxt

				.Update
			
				If Err.number <> 0 Then Err.Clear
				'On Error Resume Next
				Dim pstrFieldName
				For i = 0 To rs.Fields.Count - 1
					pstrFieldName = .Fields(i).Name
					If rs.Fields(i).Type <> 128 Then Application(pstrFieldName) = Trim(rs.Fields(pstrFieldName).Value & "")
					If Err.number <> 0 Then
						Response.Write "Error in update with <em>" & pstrFieldName & "</em>"
						Err.Clear
					End If
				Next 'i
				On Error Goto 0
				
				cnn.execute "Update sfTransactionTypes Set transIsActive=" & ((UCase(Request.Form("COD")) = "ON") * -1) & " Where transType='COD'",,128

				If Err.Number = 0 Then
					pstrMessage = "The website configuration was successfully saved."
				Elseif Err.Number = -2147217887 Then
					If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
						pstrMessage = "<H4>The data you entered is already in use.<br />Please enter a different data.</H4><br />"
						pblnError = True
					End If
				ElseIf Err.Number <> 0 Then
					Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br />"
				End If
				
			Else
				pstrMessage = "<H4>The admin record does not exist.</H4><br />"
				pblnError = True
			End If	'not .EOF
		End With 'rs
		Call ReleaseObject(rs)

    Else
        pblnError = True
    End If	'ValidateValues

    Update = (not pblnError)

	End Function    'Update

	'***********************************************************************************************

	Function ValidateValues()

	Dim strError

		strError = ""

		If Not IsNumeric(pintadminDeleteSchedule) And Len(pintadminDeleteSchedule) <> 0 Then
			strError = strError & "Please enter a number for the adminDeleteSchedule." & cstrDelimeter
		ElseIf Len(pintadminDeleteSchedule) = 0 Then
			strError = strError & "Please enter a value for the adminDeleteSchedule." & cstrDelimeter
		End If

		If Not IsNumeric(pdbladminFreeShippingAmount) And Len(pdbladminFreeShippingAmount) <> 0 Then
			strError = strError & "Please enter a number for the adminFreeShippingAmount." & cstrDelimeter
		ElseIf Len(pdbladminFreeShippingAmount) = 0 Then
			strError = strError & "Please enter a value for the adminFreeShippingAmount." & cstrDelimeter
		End If

		If Not IsNumeric(pdbladminGlobalSaleAmt) And Len(pdbladminGlobalSaleAmt) <> 0 Then
			strError = strError & "Please enter a number for the First Time Customer Discount." & cstrDelimeter
		ElseIf Len(pdbladminGlobalSaleAmt) = 0 Then
			strError = strError & "Please enter a value for the First Time Customer Discount." & cstrDelimeter
		End If

 '   If Not IsNumeric(pbytadminUpdDesign) And Len(pbytadminUpdDesign) <> 0 Then
 '       strError = strError & "Please enter a number for the adminUpdDesign." & cstrDelimeter
 '   ElseIf Len(pbytadminUpdDesign) = 0 Then
 '       strError = strError & "Please enter a value for the adminUpdDesign." & cstrDelimeter
 '   End If

 '   If Not IsNumeric(pbytadminUpdList) And Len(pbytadminUpdList) <> 0 Then
 '       strError = strError & "Please enter a number for the adminUpdList." & cstrDelimeter
 '   ElseIf Len(pbytadminUpdList) = 0 Then
 '       strError = strError & "Please enter a value for the adminUpdList." & cstrDelimeter
 '   End If

 '   If Not IsNumeric(pbytadminUpdTxt) And Len(pbytadminUpdTxt) <> 0 Then
 '       strError = strError & "Please enter a number for the adminUpdTxt." & cstrDelimeter
 '   ElseIf Len(pbytadminUpdTxt) = 0 Then
 '       strError = strError & "Please enter a value for the adminUpdTxt." & cstrDelimeter
 '   End If
    	   
		pstrMessage = strError
	    ValidateValues = (len(strError) = 0)

	End Function 'ValidateValues

End Class   'clsAdmin

'***********************************************************************************************
'***********************************************************************************************

Class clsCustomConfiguration
%><!--#include file="ssLibrary/ssmodCommonError.asp"--><%
	'Assumptions:
	'   cnn: defines a previously opened connection to the database

	'class variables
	Private plngStoreID
	Private pRS
	'database variables

	'***********************************************************************************************

	Private Sub class_Initialize()

	End Sub

	Private Sub class_Terminate()
		On Error Resume Next
		pRS.Close
		set pRS = nothing
	End Sub

	'***********************************************************************************************

	Public Property Get Recordset()
	    set Recordset = pRS
	End Property

	'***********************************************************************************************

	Public Property Let StoreID(byVal lngValue)
	    plngStoreID = lngValue
	End Property
	Public Property Get StoreID()
	    StoreID = plngStoreID
	End Property

	'***********************************************************************************************

	Public Function Load()

	'On Error Resume Next

	    Set pRS = server.CreateObject("adodb.Recordset")
		set pRS = GetRS("Select * from ssConfigurationSettings Where storeID=" & plngStoreID & " Order By configCategory,configTitle")
	    If Not (pRS.EOF Or pRS.BOF) Then
			Load = True
		End If

	End Function    'Load

	'***********************************************************************************************

	Public Function Update

	Dim pblnError
	Dim pstrSQL
	Dim pstrLocalError
	Dim i
	Dim pstrField
	Dim isDirty
	Dim pstrNewValue
	Dim pstrOrigValue
	Dim pblnUpdated
	
	'On Error Resume Next

    pblnError = False
    If Load Then
		Do While Not pRS.EOF
			pstrField = pRS.Fields("configName").Value
			isDirty = CBool(Len(Request.Form(pstrField & "_IsDirty")) > 0)

			If isDirty Then
				pstrOrigValue = pRS.Fields("configValue").Value
				pstrNewValue = LoadRequestValue(pstrField)
				
				If isValidType(pstrNewValue, pRS.Fields("configDataType").Value) Then
					'Validate Value
					If True Then
						pstrSQL = "Update ssConfigurationSettings Set configValue=" & wrapSQLValue(pstrNewValue, False, pRS.Fields("configDataType").Value) & " where configID=" & pRS.Fields("configID").Value
						If Execute_NoReturn(pstrSQL, pstrLocalError) Then
							Call addMessage("<em>" & pRS.Fields("configCategory").Value & ": " & pRS.Fields("configTitle").Value & "</em> updated.")
							pblnUpdated = True
						Else
							Call addError("<em>" & pRS.Fields("configCategory").Value & ": " & pRS.Fields("configTitle").Value & "</em> was not updated.<br />" & pstrLocalError)
						End If
					End If
				
				Else
					Call addError("<em>" & pstrNewValue & "</em> is not a valid value for " & pRS.Fields("configCategory").Value & ": " & pRS.Fields("configTitle").Value)
				End If


			Else
				pblnError = True
			End If	'ValidateValues
			pRS.MoveNext
		Loop
    End If	'Load
    
    If pblnUpdated Then ResetConfigurationSettingsInCache
    
    Update = (not pblnError)

	End Function    'Update

	'***********************************************************************************************

	Function ValidateValues()

	Dim strError

		strError = ""

 '   If Not IsNumeric(pbytadminUpdDesign) And Len(pbytadminUpdDesign) <> 0 Then
 '       strError = strError & "Please enter a number for the adminUpdDesign." & cstrDelimeter
 '   ElseIf Len(pbytadminUpdDesign) = 0 Then
 '       strError = strError & "Please enter a value for the adminUpdDesign." & cstrDelimeter
 '   End If

 '   If Not IsNumeric(pbytadminUpdList) And Len(pbytadminUpdList) <> 0 Then
 '       strError = strError & "Please enter a number for the adminUpdList." & cstrDelimeter
 '   ElseIf Len(pbytadminUpdList) = 0 Then
 '       strError = strError & "Please enter a value for the adminUpdList." & cstrDelimeter
 '   End If

 '   If Not IsNumeric(pbytadminUpdTxt) And Len(pbytadminUpdTxt) <> 0 Then
 '       strError = strError & "Please enter a number for the adminUpdTxt." & cstrDelimeter
 '   ElseIf Len(pbytadminUpdTxt) = 0 Then
 '       strError = strError & "Please enter a value for the adminUpdTxt." & cstrDelimeter
 '   End If
    	   
		pstrMessage = strError
	    ValidateValues = (len(strError) = 0)

	End Function 'ValidateValues

End Class   'clsCustomConfiguration

'***********************************************************************************************
'***********************************************************************************************

mstrPageTitle = "Website Configuration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="../SFLib/storeAdminSettings.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsAdmin
Dim mclsCustomConfiguration
Dim mstrShow
Dim prsTransactionTypes
Dim mblntranCOD

    mstrShow = Request.QueryString("Show")
    If Len(mstrShow) = 0 Then mstrShow = Request.Form("Show")

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    mlngStoreID = LoadRequestValue("StoreID")
    If Len(mlngStoreID) = 0 Then mlngStoreID = 1
    
    Set mclsAdmin = New clsAdmin
    Set mclsCustomConfiguration = New clsCustomConfiguration
	mclsAdmin.StoreID = mlngStoreID
	mclsCustomConfiguration.StoreID = mlngStoreID
	Select Case mAction
		Case "Update"
			mclsAdmin.Update
			mclsCustomConfiguration.Update
		Case "setStore":		 
		Case "newStore":		mclsAdmin.newStore
		Case "deleteStore":		mclsAdmin.deleteStore
		Case Else
	End Select

	mclsAdmin.Load
	mclsCustomConfiguration.Load

	With mclsAdmin
		
Call WriteHeader("body_onload();",True)
'On Error Resume Next
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--

function ValidInput(theForm)
{
  if (!isInteger(theForm.adminDeleteSchedule,true,"Please enter a value for the adminDeleteSchedule.")) {return(false);}
  if (!isNumeric(theForm.adminFreeShippingAmount,true,"Please enter a value for the adminFreeShippingAmount.")) {return(false);}
  if (!isNumeric(theForm.adminGlobalSaleAmt,true,"Please enter a value for the First Time Customer Discount.")) {return(false);}
//  if (!isInteger(theForm.adminUpdDesign,true,"Please enter a value for the adminUpdDesign.")) {return(false);}
//  if (!isInteger(theForm.adminUpdList,true,"Please enter a value for the adminUpdList.")) {return(false);}
//  if (!isInteger(theForm.adminUpdTxt,true,"Please enter a value for the adminUpdTxt.")) {return(false);}

//  if (isEmpty(theForm.COD_AMOUNT,"Please enter a value for the COD Amount.")) {return(false);}
//  if (!isInteger(theForm.DELETE_SCHEDULE,true,"Please enter a integer for the deletion schedule.")) {return(false);}

//  if (!isNumeric(theForm.HANDLING,true,"Please enter a value for the Handling charge.")) {return(false);}
//  if (!isNumeric(theForm.SHIP_MIN,true,"Please enter a value for the minimum shipping charge.")) {return(false);}

    return(true);
}

function DisplaySection(strSection)
{
  frmData.Show.value = strSection;

  if (strSection == "Application") {
     document.all("tblApplication").style.display = "";
     document.all("tdApplication").className = "hdrSelected";
  } else {
     document.all("tblApplication").style.display = "none";
     document.all("tdApplication").className = "hdrNonSelected";
  }
  if (strSection == "Mail") {
     document.all("tblMail").style.display = "";
     document.all("tdMail").className = "hdrSelected";
  } else {
     document.all("tblMail").style.display = "none";
     document.all("tdMail").className = "hdrNonSelected";
  }
  if (strSection == "Transaction") {
     document.all("tblTransaction").style.display = "";
     document.all("tdTransaction").className = "hdrSelected";
  } else {
     document.all("tblTransaction").style.display = "none";
     document.all("tdTransaction").className = "hdrNonSelected";
  }
  if (strSection == "Geographical") {
     document.all("tblGeographical").style.display = "";
     document.all("tdGeographical").className = "hdrSelected";
  } else {
     document.all("tblGeographical").style.display = "none";
     document.all("tdGeographical").className = "hdrNonSelected";
  }
  if (strSection == "Shipping") {
     document.all("tblShipping").style.display = "";
     document.all("tdShipping").className = "hdrSelected";
  } else {
     document.all("tblShipping").style.display = "none";
     document.all("tdShipping").className = "hdrNonSelected";
  }

  if (strSection == "Discount") {
     document.all("tblDiscount").style.display = "";
     document.all("tdDiscount").className = "hdrSelected";
  } else {
     document.all("tblDiscount").style.display = "none";
     document.all("tdDiscount").className = "hdrNonSelected";
  }

  if (strSection == "DBCleanup") {
     document.all("tblDBCleanup").style.display = "";
     document.all("tdDBCleanup").className = "hdrSelected";
  } else {
     document.all("tblDBCleanup").style.display = "none";
     document.all("tdDBCleanup").className = "hdrNonSelected";
  }

  if (strSection == "CustomConfigurationSettings") {
     document.all("tblCustomConfigurationSettings").style.display = "";
     document.all("tdCustomConfigurationSettings").className = "hdrSelected";
  } else {
     document.all("tblCustomConfigurationSettings").style.display = "none";
     document.all("tdCustomConfigurationSettings").className = "hdrNonSelected";
  }

return(false);
}

function body_onload()
{
<%
If len(mstrShow)>0 then 
	Response.Write "DisplaySection(" & chr(34) & mstrShow & chr(34) & ");"
else
	Response.Write "DisplaySection(" & chr(34) & "Application" & chr(34) & ");"
end if
%>
	return true;
}

function customAction(strAction)
{
	frmData.Action.value = strAction;
	frmData.submit();
}

function makeCustomConfigurationDirty(strID)
{
var e = document.getElementById(strID);
e.value = 1;
}

//-->
</SCRIPT>
<BODY onload="body_onload();">
<CENTER>
<div class="pagetitle "><%= mstrPageTitle %></div>
<p>
	<%= mclsAdmin.OutputMessage %>
	<%= mclsCustomConfiguration.writeErrorMessages %>
</p>
<FORM action='sfAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<input type=hidden id=Action name=Action value='Update'>
<input type=hidden id=Show name=Show value=''>
<table class="tbl" width="95%" border="1" cellpadding=0 cellspacing=0><tr><td>
<p>
Select a Store:&nbsp;<select size=1 name=StoreID id=StoreID onchange="customAction('setStore');"><% Call MakeCombo("Select adminID, adminStoreName from sfAdmin Order by adminStoreName","adminStoreName","adminID",.StoreID) %></select>
<input type=button name=newStore id=newStore class=butn value="New" onclick="customAction('newStore');">&nbsp;
<input type=button name=deleteStore id="deleteStore" class=butn value="Delete" onclick="customAction('deleteStore');">
</p>
<TABLE class="tbl" width="100%" cellpadding="3" cellspacing="0" border="0" rules="none">
	  <tr align=center>
		<th ID='tdApplication' class="hdrNonSelected" onclick='return DisplaySection("Application");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Application Settings" >Application</th>
		<th nowrap width="2pt"></th>
		<th ID='tdMail' class="hdrNonSelected" onclick='return DisplaySection("Mail");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Mail Settings" >Mail</th>
		<th nowrap width="2pt"></th>
		<th ID='tdTransaction' class="hdrNonSelected" onclick='return DisplaySection("Transaction");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Transaction Settings" >Transaction</th>
		<th nowrap width="2pt"></th>
		<th ID='tdGeographical' class="hdrNonSelected" onclick='return DisplaySection("Geographical");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Geographical Settings" >Geographical</th>
		<th nowrap width="2pt"></th>
		<th ID='tdShipping' class="hdrNonSelected" onclick='return DisplaySection("Shipping");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View Shipping Settings" >Shipping</th>
		<th nowrap width="2pt"></th>
		<th ID='tdDiscount' class="hdrNonSelected" onclick='return DisplaySection("Discount");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View global discount settings" >Discount</th>
		<th nowrap width="2pt"></th>
		<th ID='tdDBCleanup' class="hdrNonSelected" onclick='return DisplaySection("DBCleanup");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View database cleanup settings" >Database Cleanup</th>
		<th nowrap width="2pt"></th>
		<th ID='tdCustomConfigurationSettings' class="hdrNonSelected" onclick='return DisplaySection("CustomConfigurationSettings");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View custom configuration settings" >Custom Configuration</th>
		<th width="90%" align=right>&nbsp;</th>
	  </tr>
  <tr>
	<td colspan="16" class="hdrSelected" height="1px"></td>
  </tr>
</TABLE>

<TABLE class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%' id='tblApplication'>
<colgroup width="25%">
<colgroup width="75%">
      <tr>
        <TD class="SectionHeading" colspan=2>&nbsp;</TD>
      </tr>
      <tr>
        <TD class="label">Store ID:&nbsp;</LABEL></TD>
        <TD>&nbsp;<%= .StoreID %></TD>
      </tr>
      <tr>
        <TD class="label"><LABEL id=lbladminStoreName for=adminStoreName>Store Name:&nbsp;</LABEL></TD>
        <TD>&nbsp;<INPUT id=adminStoreName name=adminStoreName Value="<%= .adminStoreName %>" maxlength=50 size=50></TD>
      </tr>
      <tr>
        <TD class="label"><LABEL id=lbladminDomainName for=adminDomainName>Domain Root Path:&nbsp;</LABEL></TD>
        <TD>&nbsp;<INPUT id=adminDomainName name=adminDomainName Value="<%= .adminDomainName %>" maxlength=255 size=60></TD>
      </tr>
      <tr>
        <TD class="label"><LABEL id=lbladminSSLPath for=adminSSLPath>SSL Path:&nbsp;</LABEL></TD>
        <TD>&nbsp;<INPUT id=adminSSLPath name=adminSSLPath Value="<%= .adminSSLPath %>" maxlength=255 size=60></TD>
      </tr>
      <tr>
        <TD class="label" valign=top>ezeehelp:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminEzeeLogin name=adminEzeeLogin Value="<%= .adminEzeeLogin %>" maxlength=50 size=50><br />
          <INPUT id=adminEzeeActive name=adminEzeeActive type="checkbox" <% if .adminEzeeActive then Response.Write "checked" %>>
          Check here to activate ezeehelp  
        </TD>
      </tr>
      <tr>
        <TD class="label" valign=top>Oanada Username:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminOandaID name=adminOandaID Value="<%= .adminOandaID %>" maxlength=50 size=50><br />
          <INPUT id=adminActivateOanda name=adminActivateOanda type="checkbox" <% if .adminActivateOanda then Response.Write "checked" %>>
          Check here to activate Oanada  
        </TD>
      </tr>
      <tr>
        <TD class="label" valign=top>StoreFront Affiliate Partner:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminSFID name=adminSFID Value="<%= .adminSFID %>" maxlength=255 size=60><br />
          <INPUT id=adminSFActive name=adminSFActive type="checkbox" <% if .adminSFActive then Response.Write "checked" %>>
          Check here to activate StoreFront affiliate partner 
        </TD>
      </tr>
      <tr>
        <TD>&nbsp;</TD>
        <TD>
          <INPUT id=adminSaveCartActive name=adminSaveCartActive type="checkbox" <% if .adminSaveCartActive then Response.Write "checked" %>>
		  Check here to activate Save Cart 
        </TD>
      </tr>
      <tr>
        <TD>&nbsp;</TD>
        <TD>
          <INPUT id=adminEmailActive name=adminEmailActive type="checkbox" <% if .adminEmailActive then Response.Write "checked" %>>
		  Check here to activate email a friend
        </TD>
      </tr>
      <TR>
        <TD class="label">Minimum Order Amount:&nbsp;</TD>
        <TD>
          <INPUT id="adminMinOrderAmount" name=adminMinOrderAmount Value="<%= .adminMinOrderAmount %>" maxlength=10 size=10>
        </TD>
	  </TR>
      <tr>
        <td class="Label"><label for="adminMinOrderMessage">Min. Order Message:</label><br /><fieldset style="display: inline; text-align:center;"><font size="1"><legend>Replacement Codes</legend>{MinimumOrderAmount}<br/>{amountShort}</fieldset></font></td>
        <td>
          <textarea id="adminMinOrderMessage" onchange="MakeDirty(this);" name=adminMinOrderMessage rows="5" cols="50" title="Message"><%= .adminMinOrderMessage %></textarea>
          <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.adminMinOrderMessage);" title="Edit this field with the HTML Editor">
          <img SRC="images/prop.bmp" BORDER=0></a>
        </td>
      </tr>
      <tr>
        <td class="Label"><label for="adminTermsAndConditions">Terms & Conditions:</label></td>
        <td>
          <textarea id=adminTermsAndConditions onchange="MakeDirty(this);" name=adminTermsAndConditions rows="5" cols="50" title="Global Confirmation Message"><%= .adminTermsAndConditions %></textarea>
          <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.adminTermsAndConditions);" title="Edit this field with the HTML Editor">
          <img SRC="images/prop.bmp" BORDER=0></a>
        </td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>
          <input type="checkbox" name="adminTermsAndConditionsIsactive" id="adminTermsAndConditionsIsactive" <%= isChecked(.adminTermsAndConditionsIsactive) %>><LABEL for="adminTermsAndConditionsIsactive">Activate Terms & Conditions</LABEL>
        </td>
	  </tr>

</TABLE>
<TABLE  class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%' id='tblMail'>
	<colgroup width="25%">
	<colgroup width="75%">
      <tr>
        <TD class="SectionHeading" colspan=2>&nbsp;</TD>
      </tr>
      <tr>
        <TD class="label">Mail Method:&nbsp;</TD>
        <TD>&nbsp;<select size="1" name="adminMailMethod">
<% Call MakeCombo("Select slctvalMailMethod from sfSelectValues where slctvalMailMethod<>'' Order by slctvalMailMethod","slctvalMailMethod","slctvalMailMethod",.adminMailMethod) %>
          </select></TD>
      </tr>
      <tr>
        <TD class="label">Mail Server Address:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminMailServer name=adminMailServer Value="<%= .adminMailServer %>" maxlength=255 size=60></TD>
      </tr>
      <tr>
        <TD class="label">Mail Server Username:&nbsp;</TD>
        <TD>&nbsp;<INPUT id="adminMailServerUserName" name=adminMailServerUserName Value="<%= .adminMailServerUserName %>" maxlength=255 size=60></TD>
      </tr>
      <tr>
        <TD class="label">Mail Server Password:&nbsp;</TD>
        <TD>&nbsp;<INPUT id="adminMailServerPassword" name=adminMailServerPassword Value="<%= .adminMailServerPassword %>" maxlength=255 size=60></TD>
      </tr>
      <tr>
        <TD colspan=2>&nbsp;</TD>
      </tr>
      <tr>
        <TD>&nbsp;</TD>
        <TD align=left>&nbsp;<b><i>Technical Contact Email</i></b></TD>
      </tr>
      <tr>
        <TD class="label">Email Recipient(s):&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminTechnicalEmail name=adminTechnicalEmail Value="<%= .adminTechnicalEmail %>" maxlength=100 size=60></TD>
      </tr>
      <tr>
        <TD>&nbsp;</TD>
        <TD align=left>&nbsp;<b><i>Customer Confirmation email</i></b></TD>
      </tr>
      <tr>
        <TD class="label">Primary email recipient:&nbsp;</TD>
        <TD>&nbsp;<INPUT id="adminPrimaryEmail" name=adminPrimaryEmail Value="<%= .adminPrimaryEmail %>" maxlength=100 size=60></TD>
      </tr>
      <tr>
        <TD class="label">Secondary email recipient:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminSecondaryEmail name=adminSecondaryEmail Value="<%= .adminSecondaryEmail %>" maxlength=100 size=60></TD>
      </tr>
      <tr>
        <TD class="label">Subject Line:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminEmailSubject name=adminEmailSubject Value="<%= .adminEmailSubject %>" maxlength=255 size=60></TD>
      </tr>
      <tr>
        <TD class="label" valign=top>Mail Message:&nbsp;</TD>
        <TD>&nbsp;<TEXTAREA id=adminEmailMessage name=adminEmailMessage rows="3" cols="51"><%= .adminEmailMessage %></TEXTAREA></TD>
      </tr>
      <tr>
        <TD>&nbsp;</TD>
        <TD >
          <INPUT id=adminSubscribeMailIsActive name=adminSubscribeMailIsActive type="checkbox" <% if .adminSubscribeMailIsActive then Response.Write "checked" %>>
          <label for="adminSubscribeMailIsActive">Check here to show Subscribe mail checkbox on Order form</label>
        </TD>
      </tr>
</TABLE>
<TABLE  class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%' id='tblTransaction'>
	<colgroup width="25%">
	<colgroup width="75%">
      <tr>
        <TD class="SectionHeading" colspan=2>&nbsp;</TD>
      </tr>
        <TD class="label">Transaction Method:&nbsp;</TD>
        <TD>&nbsp;<select size="1" name="adminTransMethod">
<% Call MakeCombo("Select trnsmthdID,trnsmthdName from sfTransactionMethods Order by trnsmthdName","trnsmthdName","trnsmthdID",.adminTransMethod) %>
          </select></TD>
      </tr>
        <TD class="label" valign=top>Processing Method:&nbsp;</TD>
        <TD>
           <input type="radio" value="authonly" id="authonly0" name="adminMerchantType" <% if .adminMerchantType="authonly" then Response.Write "checked" %>>
           <label for="authonly0">Authorization Only</label><br />
           <input type="radio" value="authcapture" name="adminMerchantType" id="authonly1" <% if .adminMerchantType="authcapture" then Response.Write "checked" %>>
           <label for="authonly1">Authorization and Capture</label>
        </TD>
      </tr>
      <tr>
        <TD>&nbsp;</TD>
        <TD>
          <INPUT type="checkbox" id=adminDeletePolicy name=adminDeletePolicy value=2 <%= isChecked(.adminDeletePolicy = "2") %>>
          <label for="adminDeletePolicy">Check here to delete credit card data every</label> <INPUT id=adminDeleteSchedule name=adminDeleteSchedule Value="<%= .adminDeleteSchedule %>" size=6> days</TD>
		</TD>
      </tr>
      <tr>
        <TD class="label">Payment Server Path:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminPaymentServer name=adminPaymentServer Value="<%= .adminPaymentServer %>" maxlength=255 size=60></TD>
      </tr>
      <tr>
        <TD class="label">Payment Server Login:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminLogin name=adminLogin Value="<%= .adminLogin %>" size=50></TD>
      </tr>
      <tr>
        <TD class="label">Payment Server Password:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=adminPassword name=adminPassword Value="<%= .adminPassword %>" size=50></TD>
      </tr>
      <tr>
        <TD>&nbsp;</TD>
        <TD>&nbsp;
			<INPUT id=adminEncodeCCIsActive name=adminEncodeCCIsActive type="checkbox" <% if .adminEncodeCCIsActive then Response.Write "checked" %>>
			Encode Credit Card Numbers
		</TD>
      </tr>
      <tr>
        <TD>&nbsp;</TD>
        <TD>
          <table class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%'>
            <tr>
              <TD valign=top>
                <b><i><a href="sfTransactionTypesAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Payment Methods">Accepted Payment Types</a></i></b>
                <ul>
<%
Set prsTransactionTypes = server.CreateObject("adodb.Recordset")
set prsTransactionTypes = GetRS("Select * from sfTransactionTypes")
prsTransactionTypes.Filter = "transType = 'COD'"
if prsTransactionTypes.EOF Then
	mblntranCOD = False
Else
	mblntranCOD = prsTransactionTypes("transIsActive") = 1
End If

prsTransactionTypes.Filter = "transType <> 'Credit Card'"
prsTransactionTypes.MoveFirst
do while not prsTransactionTypes.EOF
    If len(trim(prsTransactionTypes("transName"))) > 0 Then
		Response.Write "<li>" & prsTransactionTypes("transType") & " / " & prsTransactionTypes("transName") & vbcrlf
    Else
		Response.Write "<li>" & prsTransactionTypes("transType") & vbcrlf
    End If
	prsTransactionTypes.MoveNext
loop
%>
				</ul>
              </TD>
              <TD valign=top>
                <b><i><a href="sfTransactionTypesAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Payment Methods">Accepted Credit Cards</a></i></b>
                <ul>
<%
prsTransactionTypes.Filter = "transType = 'Credit Card'"
prsTransactionTypes.MoveFirst
do while not prsTransactionTypes.EOF
    Response.Write "<li>" & prsTransactionTypes("transName") & vbcrlf
	prsTransactionTypes.MoveNext
loop
Set prsTransactionTypes = Nothing
%>
				</ul>
              </TD>
            </tr>
          </table>
        </TD>
      </tr>
</TABLE>
<TABLE  class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%' id='tblGeographical'>
	<colgroup width="25%">
	<colgroup width="75%">
      <tr>
        <TD class="SectionHeading" colspan=2>&nbsp;</TD>
      </tr>
      <tr>
        <TD class="label">Origin State:&nbsp;</TD>
        <TD><select size="1" name="adminOriginState" ID="adminOriginState"><% Call MakeCombo("Select loclstName,loclstAbbreviation from sfLocalesState Order by loclstName","loclstName","loclstAbbreviation",.adminOriginState) %></select></TD>
      </tr>
      <tr>
        <TD class="label">Origin Zip:&nbsp;</TD>
        <TD><INPUT id=adminOriginZip name=adminOriginZip Value="<%= .adminOriginZip %>" maxlength=15 size=15></TD>
      </tr>
      <tr>
        <TD class="label">Origin Country:&nbsp;</TD>
        <TD><select size="1" name="adminOriginCountry" ID="Select2"><% Call MakeCombo("Select loclctryName,loclctryAbbreviation from sfLocalesCountry Order by loclctryName","loclctryName","loclctryAbbreviation",.adminOriginCountry) %></select></TD>
      </tr>
      <tr>
        <TD class="label">Currency Type - LCID:&nbsp;</TD>
        <TD><select size="1" name="adminLCID"><% Call MakeCombo("Select slctvalLCIDLabel,slctvalLCID from sfSelectValues where slctvalLCID<>'' Order by slctvalLCIDLabel","slctvalLCIDLabel","slctvalLCID",.adminLCID) %></select></TD>
      </tr>
</TABLE>
<TABLE  class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%' id='tblShipping'>
	<colgroup width="25%">
	<colgroup width="75%">
      <tr>
        <TD class="SectionHeading" colspan=2>&nbsp;</TD>
      </tr>
      <TR>
      <td colspan=2>
        <table width=100% border=0>
        <tr>
        <td width=10%>&nbsp;</td>
       <TD width=40% valign=top>
          <p><b>Primary Shipping Method</b></p>
          <input type="radio" value="2" name="adminShipType" id="adminShipType2" <% if .adminShipType="2" then Response.Write "checked" %>><a href="ssPostageRate_shippingMethodsAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Carrier Based Rates">Real Time Rates (by weight)</a><br />
          <input type="radio" value="3" name="adminShipType" id="adminShipType3" <% if .adminShipType="3" then Response.Write "checked" %>>Per Product Based Rates<br />
          <input type="radio" value="1" name="adminShipType" id="adminShipType1" <% if .adminShipType="1" then Response.Write "checked" %>><a href="sszbsShippingRatesAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Value Based Rates">Zone Based Rates, By Subtotal</a><br />
          <input type="radio" value="4" name="adminShipType" id="adminShipType4" <% if .adminShipType="4" then Response.Write "checked" %>><a href="sszbsShippingRatesAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Value Based Rates">Zone Based Rates, By Weight</a><br />
          <input type="radio" value="5" name="adminShipType" id="adminShipType5" <% if .adminShipType="5" then Response.Write "checked" %>><a href="sszbsShippingRatesAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Value Based Rates">Zone Based Rates, By Item Count</a>
        </TD>
       <TD width=50% valign=top>
          <p><b>Back-up Shipping Method</b></p>
          <input type="radio" value="3" name="adminShipType2" <% if .adminShipType2="3" then Response.Write "checked" %>>Product Based Rates<br />
          <input type="radio" value="1" name="adminShipType2" <% if .adminShipType2="1" then Response.Write "checked" %>><a href="sszbsShippingRatesAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Value Based Rates">Zone Based Rates, By Subtotal</a><br />
          <input type="radio" value="4" name="adminShipType2" <% if .adminShipType2="4" then Response.Write "checked" %>><a href="sszbsShippingRatesAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Value Based Rates">Zone Based Rates, By Weight</a><br />
          <input type="radio" value="5" name="adminShipType2" <% if .adminShipType2="5" then Response.Write "checked" %>><a href="sszbsShippingRatesAdmin.asp" onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title="Configure Value Based Rates">Zone Based Rates, By Item Count</a>
        </TD>
        </tr>
        </table>
      </td>
	  </TR>
      <TR>
        <TD class="label">&nbsp;</TD>
        <TD>
		  <INPUT id=adminTaxShipIsActive name=adminTaxShipIsActive type="checkbox" <% if .adminTaxShipIsActive then Response.Write "checked" %>>
		  <LABEL id=lbladminTaxShipIsActive for=adminTaxShipIsActive>Check here to apply tax to shipping charges</LABEL>
        </TD>
	  </TR>
      <TR>
        <TD class="label">Minimum Shipping Charge:&nbsp;</TD>
        <TD>
          <INPUT id=adminShipMin name=adminShipMin Value="<%= .adminShipMin %>" maxlength=10 size=10>
        </TD>
	  </TR>
      <TR>
        <TD class="label">Premium Shipping Amount:&nbsp;</TD>
        <TD>
          <INPUT id=adminSpcShipAmt name=adminSpcShipAmt Value="<%= .adminSpcShipAmt %>" maxlength=10 size=10>
          <INPUT id=adminPrmShipIsActive name=adminPrmShipIsActive type="checkbox" <% if .adminPrmShipIsActive then Response.Write "checked" %>><LABEL id=lbladminPrmShipIsActive for=adminPrmShipIsActive>Check here to activate premium shipping</LABEL> <em>Applies to Product Based shipping only</em>
        </TD>
	  </TR>
      <TR>
        <TD class="label">C.O.D. Amount:&nbsp;</TD>
        <TD>
          <INPUT id=adminCODAmount name=adminCODAmount Value="<%= .adminCODAmount %>" maxlength=10 size=10>
          <INPUT id=COD name=COD type="checkbox" <% if mblntranCOD then Response.Write "checked" %>><LABEL id=lblCOD for=COD>Check here to allow COD orders</LABEL>
        </TD>
	  </TR>
      <TR>
        <TD class="label">No handling charge for orders over:&nbsp;</TD>
        <TD><INPUT id=adminHandlingIsActive name=adminHandlingIsActive value="<%= .adminHandlingIsActive %>"></TD>
      </tr>
      <TR>
        <TD class="label" valign="top">&nbsp;</TD>
        <TD>
          <p>Apply a handling charge of <INPUT id=adminHandling name=adminHandling Value="<%= .adminHandling %>" maxlength=10 size=10> to:<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="adminHandlingType" id="adminHandlingType1" value="1" <%= isChecked(.adminHandlingType=1) %>><label for="adminHandlingType1">All orders</label><br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="adminHandlingType" id="adminHandlingType2" value="2" <%= isChecked(.adminHandlingType=2) %>><label for="adminHandlingType2">Shipped orders only</label><br />
            <!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="adminHandlingType" id="adminHandlingType3" value="3" <%= isChecked(.adminHandlingType=3) %>><label for="adminHandlingType3">International orders only</label>-->
          </p>
        </TD>
	  </TR>
      <TR>
        <TD class="label">Free Shipping for orders over:&nbsp;</TD>
        <TD>
          <INPUT id=adminFreeShippingAmount name=adminFreeShippingAmount Value="<%= .adminFreeShippingAmount %>" maxlength=10 size=10>&nbsp;&nbsp;
          <INPUT id=adminFreeShippingIsActive name=adminFreeShippingIsActive type="checkbox" <% if .adminFreeShippingIsActive then Response.Write "checked" %>><LABEL id=lbladminPrmShipIsActive for=adminPrmShipIsActive>Check here to activate free shipping</LABEL>
        </TD>
	  </TR>
</TABLE>

<TABLE class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%' id="tblDiscount">
<colgroup width="25%">
<colgroup width="75%">
      <tr>
        <TD class="SectionHeading" colspan=2>&nbsp;</TD>
      </tr>
      <TR>
        <TD class="label">First Time Customer Discount:&nbsp;</TD>
        <TD>
          <INPUT id=adminGlobalSaleAmt name=adminGlobalSaleAmt Value="<%= .adminGlobalSaleAmt %>" maxlength=10 size=5>&nbsp;&nbsp;(Set to 0 to make inactive)
        </td>
	  </tr>
      <tr>
        <td>&nbsp;</td>
        <td>
          <INPUT id=adminGlobalSaleIsActive name=adminGlobalSaleIsActive type="checkbox" <%= isChecked(.adminGlobalSaleIsActive) %>><LABEL for=adminGlobalSaleIsActive>First Time Customer Discount is a percent off</LABEL>
        </td>
	  </tr>
      <tr>
        <td class="Label"><label for="adminGlobalConfirmationMessage">Global Confirmation Message:</label></td>
        <td>
          <textarea id=adminGlobalConfirmationMessage onchange="MakeDirty(this);" name=adminGlobalConfirmationMessage rows="5" cols="50" title="Global Confirmation Message"><%= .adminGlobalConfirmationMessage %></textarea>
          <a HREF="javascript:doNothing()" onClick="return openACE(document.frmData.adminGlobalConfirmationMessage);" title="Edit this field with the HTML Editor">
          <img SRC="images/prop.bmp" BORDER=0></a>
        </td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>
          <input type="checkbox" name="adminGlobalConfirmationMessageIsactive" id="adminGlobalConfirmationMessageIsactive" <%= isChecked(.adminGlobalConfirmationMessageIsactive) %>><LABEL for="adminGlobalConfirmationMessageIsactive">Activate Global Confirmation Message</LABEL>
        </td>
	  </tr>
</TABLE>

<script language="javascript">

tipMessage['hoursBetweenCleanings']=["Data Entry Help","Number of hours between cleanings."]
tipMessage['daysToSaveIncompleteOrders']=["Data Entry Help","Number of days to keep incomplete orders before they are automatically deleted."]
tipMessage['daysToSaveTempOrders']=["Data Entry Help","Number of days to keep visitors' carts."]
tipMessage['daysToSaveSavedOrders']=["Data Entry Help","Number of days to keep visitors' saved carts."]
tipMessage['daysToKeepVisitors']=["Data Entry Help","Number of days to keep visitors."]
tipMessage['CCReplace']=["Data Entry Help","Text to replace credit card numbers with."]

</script>
<TABLE class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%' id="tblDBCleanup">
<colgroup>
  <col  width="25%">
  <col  width="75%">
</colgroup>
      <tr>
        <TD class="SectionHeading" colspan=2>&nbsp;</TD>
      </tr>
      <tr>
        <td class="label"><label for="hoursBetweenCleanings" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Hours Between Cleanings</label>:&nbsp;</td>
        <td><INPUT id="hoursBetweenCleanings" name="hoursBetweenCleanings" Value="<%= .hoursBetweenCleanings %>" maxlength=10 size=5></td>
	  </tr>
      <tr>
        <td class="label"><label for="daysToSaveIncompleteOrders" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Days to Keep Incomplete Orders</label>:&nbsp;</td>
        <td><INPUT id="daysToSaveIncompleteOrders" name="daysToSaveIncompleteOrders" Value="<%= .daysToSaveIncompleteOrders %>" maxlength=10 size=5></td>
	  </tr>
      <tr>
        <td class="label"><label for="daysToSaveTempOrders" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Days to Keep Shopping Carts</label>:&nbsp;</td>
        <td><INPUT id="daysToSaveTempOrders" name="daysToSaveTempOrders" Value="<%= .daysToSaveTempOrders %>" maxlength=10 size=5></td>
	  </tr>
      <tr>
        <td class="label"><label for="daysToSaveSavedOrders" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Days to Keep Saved Orders</label>:&nbsp;</td>
        <td><INPUT id="daysToSaveSavedOrders" name="daysToSaveSavedOrders" Value="<%= .daysToSaveSavedOrders %>" maxlength=10 size=5></td>
	  </tr>
      <tr>
        <td class="label"><label for="daysToKeepVisitors" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Days to Keep Visitors</label>:&nbsp;</td>
        <td><INPUT id="daysToKeepVisitors" name="daysToKeepVisitors" Value="<%= .daysToKeepVisitors %>" maxlength=10 size=5></td>
	  </tr>
      <tr>
        <td class="label"><label for="CCReplace" onmouseover="showDataEntryTip(this);" onmouseout="htm();">Credit Card Replacement Text</label>:&nbsp;</td>
        <td><INPUT id="CCReplace" name="CCReplace" Value="<%= .CCReplace %>" maxlength=20 size=20></td>
	  </tr>
</TABLE>
<%
end with	'mclsAdmin
%>
<TABLE class="tbl" border=1 cellPadding=2 cellSpacing=0 id='tblCustomConfigurationSettings'>
<%
Dim rsConfig
Dim pstrPrevCategory
Dim strOptions

With mclsCustomConfiguration
Set rsConfig = .Recordset
Do While Not rsConfig.EOF
	If pstrPrevCategory <> Trim(rsConfig.Fields("configCategory").Value & "") Then
		pstrPrevCategory = Trim(rsConfig.Fields("configCategory").Value & "")
		Response.Write "<tr><th class=""tblhdr"" colspan=""3"" align=""left"">&nbsp;" & pstrPrevCategory & "</td></tr>"
	End If
	
	strOptions = deserializeArray(Trim(rsConfig.Fields("configOptions").Value & ""), ",", ";")
	Response.Write "<tr>"
	Response.Write "<td class=""label"">" & rsConfig.Fields("configTitle").Value & ":&nbsp;<input type=""hidden"" name=""" & rsConfig.Fields("configName").Value & "_isDirty"" id=""" & rsConfig.Fields("configName").Value & "_isDirty""></td>"
	Response.Write "<td>" & writeHTMLFormElement(rsConfig.Fields("configDisplayType").Value, "", rsConfig.Fields("configName").Value, rsConfig.Fields("configName").Value, rsConfig.Fields("configValue").Value, strOptions, " onchange=""makeCustomConfigurationDirty('" & rsConfig.Fields("configName").Value & "_isDirty');""") & "</td>"
	Response.Write "<td>" & rsConfig.Fields("configDescription").Value & "</td>"
	
	Response.Write "</tr>"
	rsConfig.MoveNext
Loop
Set rsConfig = Nothing
End With	'mclsCustomConfiguration
%>
</TABLE>

<TABLE class="tbl" border=0 cellPadding=2 cellSpacing=0 width='100%'>
	<colgroup width="25%">
	<colgroup width="75%">
      <TR>
        <TD colspan="2" class="SectionHeading">&nbsp;
      </TR>
  <TR>
    <TD colspan="2" class="SectionHeading" align=center>
        &nbsp; <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset>
        &nbsp;<INPUT class='butn' id=btnUpdate name=btnUpdate type=submit value='Save Changes'>
    </TD>
  </TR>
</TABLE>
</td></tr></table>
</FORM>

</CENTER>
<% 

Set mclsAdmin = Nothing
Set mclsCustomConfiguration = Nothing
Set cnn = Nothing
Response.Flush
%>
</BODY>
</HTML>