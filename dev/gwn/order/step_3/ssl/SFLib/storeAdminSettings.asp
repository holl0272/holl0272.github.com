<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Revision Date: November 26, 2004											*
'*   Version 1.01.001                                                           *
'*                                                                              *
'*   1.01.001 (November 26, 2004)                                               *
'*   - general cleanup
'*   - Moved incGeneral.asp variable declarations to this file
'*                                                                              *
'*   1.00.002 (April 21, 2004)                                                  *
'*   - added check if Application settings are incorrect                        *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	'Noene

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim maryApplicationSettings
Dim maryConfigurationSettings

Dim mstrBodyStyle
Dim cstrTDBottomNavStyle
Dim cstrCartDisplay_MfgName
Dim cstrCartDisplay_ProductID
Dim cstrCCV_Optional
Dim cstrCCV_SaveToDB
Dim cstrCCVFieldName
Dim cstrCODTerm
Dim cstrTDContentStyle
Dim cssDefaultMaxRecords
Dim cssDefaultPageSize
Dim cblnDisableLogin
Dim cblnDisplayCategoryDescriptions
Dim cblnDisplayProductIDs
Dim cblnExpandCategoriesByDefault
Dim cstrECheckTerm
Dim C_FORMDESIGN
Dim cstrGoogleAdwords_conversion_id
Dim cstrGoogleAnalytics_uacct
Dim cstrHighlightSearchTermClass
Dim cstrImageDetailInstructions
Dim cblnIncludeEmailVerification
Dim cstrLargeWindowImageField
Dim cstrOvertureID
Dim cstrTDLeftNavStyle
Dim clngMaxLengthDescription
Dim cstrNewProductField
Dim clngNewProductsDaysSinceAdded
Dim clngNumEntries
Dim cbytOrderViewImageSrc
Dim C_BNRBKGRND
Dim C_WIDTH
Dim cstrPhoneFaxTerm
Dim cstrPOTerm
Dim cstrPrimaryEmailToSendErrorTo
Dim cstrTDRightNavStyle
Dim cblnSearchAttributes
Dim cbytSearchResultsDisplayType
Dim cstrSearchResultsDisplayTypeFixed
Dim cblnShowOptionToTurnOffImages
Dim cblnShowSearchCustomizationOption
Dim cstrSubWebPath
Dim cstrTextLinkToLargerWindow
Dim cstrTimeoutRedirectPage
Dim cstrTDTopNavStyle
Dim cblnTrackPageViews
Dim cstrURLTemplate_Manufacturer
Dim cstrURLTemplate_Vendor

'**********************************************************
'*	Functions
'**********************************************************

'Function CurrencyISO()
'Function getAppSetting(byVal strAppSetting)
'Function IsSaveCartActive()
'Function LoadApplicationSettingsToCache()
'Function RetrieveFromApplication(byVal strField, byVal vntDefault)

'**********************************************************
'*	Begin Page Code
'**********************************************************

Function CurrencyISO()

Dim pstrSQL
Dim pobjRS
Dim pstrLCID

	If Len(Application("CurrencyISO")) > 0 Then
		pstrLCID = Application("CurrencyISO")
	Else
		pstrSQL = "Select slctvalCurrencyISO From sfSelectValues Where slctvalLCID = '" &  makeInputSafe(Session.LCID) & "'" 
		set pobjRS = GetRS(pstrSQL)
		pstrLCID = trim(pobjRS.Fields("slctvalCurrencyISO").Value) 
		Application("CurrencyISO") = pstrLCID
		Call closeObj(pobjRS)
	End If

    getCurrencyISO = pstrLCID

End Function	'getCurrencyISO

'********************************************************************************

Function getAppSetting(byVal strAppSetting)
	getAppSetting = RetrieveFromApplication(strAppSetting, "")
End Function	' getAppSetting

Function C_STORENAME
	C_STORENAME = RetrieveFromApplication("adminStoreName", "")
End Function

Function C_HomePath
	C_HomePath = RetrieveFromApplication("adminDomainName", "")
End Function

Function C_SecurePath
	C_SecurePath = RetrieveFromApplication("adminSSLPath", "")
End Function

Function sUserName
	sUserName = RetrieveFromApplication("adminOandaID", "")
End Function

Function iConverion
	iConverion = RetrieveFromApplication("adminActivateOanda", 0)
End Function

Function sEzeeHelp
	sEzeeHelp = RetrieveFromApplication("adminEzeeLogin", "")
End Function

Function iEzeeHelp
	iEzeeHelp = RetrieveFromApplication("adminEzeeActive", 0)
End Function

Function iSaveCartActive
	iSaveCartActive = IsSaveCartActive
End Function

Function iEmailActive
	iEmailActive = RetrieveFromApplication("adminEmailActive", 0)
End Function

Function iBrandActive
	iBrandActive = RetrieveFromApplication("adminSFActive", 0)
End Function

Function sAffID
	sAffID = RetrieveFromApplication("adminSFID", "")
End Function

'********************************************************************************

Function IsSaveCartActive()

Dim pbytIsActive

	pbytIsActive = RetrieveFromApplication("adminSaveCartActive", 0)
	If isNumeric(pbytIsActive) Then
		pbytIsActive = CLng(pbytIsActive)
	Else
		pbytIsActive = 0
	End If

	IsSaveCartActive = pbytIsActive

End Function	' IsSaveCartActive

'********************************************************************************

Function StoreID

Dim plngStoreID

	plngStoreID = Application("StoreID")
	If Not isNumeric(plngStoreID) Then plngStoreID = 0
	
	StoreID = plngStoreID

End Function

'********************************************************************************

Function LoadApplicationSettingsToCache()

Dim pobjRS
Dim pstrSQL
Dim pblnSuccess
Dim i
Dim pstrFieldName
Dim pvntTemp

	'Response.Write "<h4>Loading Application Settings to cache . . .</h4>"

	pblnSuccess = False
	
	If Err.number <> 0 Then Err.Clear
	
	If StoreID = 0 Then
		pstrSQL = "SELECT * FROM sfAdmin"
	Else
		pstrSQL = "SELECT * FROM sfAdmin Where adminID = " & StoreID
	End If
	
	Set	pobjRS = CreateObject("adodb.recordset")
	With pobjRS
		.CursorLocation = 2 'adUseClient
	    
		On Error Resume Next
		.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		If Err.number <> 0 Then
			Response.Write "<font color=red>Error in LoadApplicationSettingsToCache: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
			Response.Write "<font color=red>LoadApplicationSettingsToCache: sql = " & pstrSQL & "</font><br />" & vbcrlf
			Response.Flush
			Err.Clear
		ElseIf Not .EOF Then
			On Error Goto 0
			For i = 0 To .Fields.Count - 1
				If .Fields(i).Type <> 128 Then	'ts_upsize
					pstrFieldName = .Fields(i).Name
					pvntTemp = Trim(.Fields(i).Value & "")
					'handle special cases
					Select Case LCase(pstrFieldName)
						Case "adminhandling", "adminshipmin", "admincodamount", "adminspcshipamt", "adminminorderamount"
							If Len(pvntTemp) = 0 Or Not isNumeric(pvntTemp) Then pvntTemp = 0
						Case Else
							'do nothing
					End Select
					Application(pstrFieldName) = pvntTemp
				End If
			Next 'i
			
			If cblnSF5AE Then
				pvntTemp = Trim(.Fields("adminBackOrderBilling").Value)
				If Len(pvntTemp) = 0 Or Not isNumeric(pvntTemp) Then pvntTemp = 0
			Else
				pvntTemp = 0
			End If
			Application("adminBackOrderBilling") = pvntTemp

			If Right(Application("adminDomainName") & "", 1) <> "/" Then Application("adminDomainName") = Application("adminDomainName") & "/"

			If False Then
				Application("adminStoreName") = Trim(.Fields("adminStoreName").Value)
				Application("adminDomainName") = Trim(.Fields("adminDomainName").Value)
				Application("adminSSLPath") = Trim(.Fields("adminSSLPath").Value)
				Application("adminOandaID") = Trim(.Fields("adminOandaID").Value)
				Application("adminActivateOanda") = Trim(.Fields("adminActivateOanda").Value)
				Application("adminEzeeLogin") = Trim(.Fields("adminEzeeLogin").Value)
				Application("adminEzeeActive") = Trim(.Fields("adminEzeeActive").Value)
				Application("adminSaveCartActive") = Trim(.Fields("adminSaveCartActive").Value)
				Application("adminEmailActive") = Trim(.Fields("adminEmailActive").Value)
				Application("adminSFActive") = Trim(.Fields("adminSFActive").Value)
				Application("adminSFID") = Trim(.Fields("adminSFID").Value)
				Application("adminLCID") = Trim(.Fields("adminLCID").Value)
				Application("adminShipType") = Trim(.Fields("adminShipType").Value)
				Application("adminFreeShippingIsActive") = Trim(.Fields("adminFreeShippingIsActive").Value)
				Application("adminFreeShippingAmount") = Trim(.Fields("adminFreeShippingAmount").Value)
				Application("adminPrmShipIsActive") = Trim(.Fields("adminPrmShipIsActive").Value)
				Application("adminEncodeCCIsActive") = Trim(.Fields("adminEncodeCCIsActive").Value)
				Application("adminTaxShipIsActive") = Trim(.Fields("adminTaxShipIsActive").Value)

				Application("adminPrimaryEmail") = Trim(.Fields("adminPrimaryEmail").Value)
				Application("adminSecondaryEmail") = Trim(.Fields("adminSecondaryEmail").Value)
				Application("adminEmailSubject") = Trim(.Fields("adminEmailSubject").Value)
				Application("adminEmailMessage") = Trim(.Fields("adminEmailMessage").Value)
				Application("adminMailMethod") = Trim(.Fields("adminMailMethod").Value)
				Application("adminMailServer") = Trim(.Fields("adminMailServer").Value)
				
				Application("adminGlobalSaleIsActive") = Trim(.Fields("adminGlobalSaleIsActive").Value)
				Application("adminGlobalSaleIsActive") = Trim(.Fields("adminGlobalSaleIsActive").Value)
			End If
			
			pblnSuccess = True
			
		End If
		.Close
	End with
	Set pobjRS = Nothing

	If Len(Session("LCID")) = 0 Then Session("LCID") = adminLCID
	Session.LCID = Session("LCID")

	LoadApplicationSettingsToCache = pblnSuccess

End Function	' LoadApplicationSettingsToCache

'********************************************************************************

Function RetrieveFromApplication(byVal strField, byVal vntDefault)

Dim vntValue

	'Call LoadApplicationSettingsToCache	'For Testing
	If Len(Application(strField)) > 0 Then
		vntValue = Application(strField)
	ElseIf Len(Application("adminStoreName")) = 0 Then
		If LoadApplicationSettingsToCache Then
			vntValue = Application(strField)
		Else
			vntValue = vntDefault
		End If
	Else
		vntValue = ""
	End If

	RetrieveFromApplication = vntValue

End Function	' RetrieveFromApplication

'--------------------------------------------------------------------------------------------------
' Properties from admin table
'--------------------------------------------------------------------------------------------------

Function adminSFID()
	adminSFID= getAppSetting("adminSFID")
End Function

Function adminStoreName()
	adminStoreName= getAppSetting("adminStoreName")
End Function

Function adminDomainName()
	adminDomainName= getAppSetting("adminDomainName")
End Function

Function adminSSLPath()
	adminSSLPath= getAppSetting("adminSSLPath")
End Function

Function adminOandaID()
	adminOandaID= getAppSetting("adminOandaID")
End Function

Function adminActivateOanda()
	adminActivateOanda= getAppSetting("adminActivateOanda")
End Function

Function adminEzeeLogin()
	adminEzeeLogin= getAppSetting("adminEzeeLogin")
End Function

Function adminEzeeActive()
	adminEzeeActive= getAppSetting("adminEzeeActive")
End Function

Function adminSaveCartActive()
	adminSaveCartActive= getAppSetting("adminSaveCartActive")
End Function

Function adminEmailActive()
	adminEmailActive= getAppSetting("adminEmailActive")
End Function

Function adminSFActive()
	adminSFActive= getAppSetting("adminSFActive")
End Function

Function adminSFActive()
	adminSFActive= getAppSetting("adminSFActive")
End Function

Function adminLCID()
	adminLCID= getAppSetting("adminLCID")
End Function

Function adminFreeShippingIsActive()
	adminFreeShippingIsActive= getAppSetting("adminFreeShippingIsActive")
End Function

Function adminFreeShippingAmount()
	adminFreeShippingAmount= getAppSetting("adminFreeShippingAmount")
End Function

Function adminShipType()
	adminShipType= getAppSetting("adminShipType")
End Function

Function adminShipType2()
	adminShipType2= getAppSetting("adminShipType2")
End Function

Function adminPrmShipIsActive()
	adminPrmShipIsActive= getAppSetting("adminPrmShipIsActive")
End Function

Function adminTransMethod()
	adminTransMethod= getAppSetting("adminTransMethod")
End Function

Function adminPaymentServer()
	adminPaymentServer= getAppSetting("adminPaymentServer")
End Function

Function adminLogin()
	adminLogin= getAppSetting("adminLogin")
End Function

Function adminPassword()
	adminPassword= getAppSetting("adminPassword")
End Function

Function adminMerchantType()
	adminMerchantType= getAppSetting("adminMerchantType")
End Function

Function adminEncodeCCIsActive()
	adminEncodeCCIsActive= getAppSetting("adminEncodeCCIsActive")
End Function

Function adminPrimaryEmail()
	adminPrimaryEmail= getAppSetting("adminPrimaryEmail")
End Function

Function adminSecondaryEmail()
	adminSecondaryEmail= getAppSetting("adminSecondaryEmail")
End Function

Function adminEmailSubject()
	adminEmailSubject= getAppSetting("adminEmailSubject")
End Function

Function adminEmailMessage()
	adminEmailMessage= getAppSetting("adminEmailMessage")
End Function

Function adminEmailSubject()
	adminEmailSubject= getAppSetting("adminEmailSubject")
End Function

Function adminEmailMessage()
	adminEmailMessage= getAppSetting("adminEmailMessage")
End Function

Function adminMailMethod()
	adminMailMethod= getAppSetting("adminMailMethod")
End Function

Function adminMailServer()
	adminMailServer= getAppSetting("adminMailServer")
End Function

Function adminGlobalSaleIsActive()
	adminGlobalSaleIsActive= getAppSetting("adminGlobalSaleIsActive")
End Function

Function adminGlobalSaleAmt()
	adminGlobalSaleAmt= getAppSetting("adminGlobalSaleAmt")
End Function

Function adminTaxShipIsActive()
	adminTaxShipIsActive= getAppSetting("adminTaxShipIsActive")
End Function

Function adminShipMin()
	adminShipMin= getAppSetting("adminShipMin")
End Function

Function adminSpcShipAmt()
	adminSpcShipAmt = getAppSetting("adminSpcShipAmt")
End Function

Function adminHandling()
Dim pvnt: pvnt = getAppSetting("adminHandling")
	If Len(pvnt) > 0 And isNumeric(pvnt) Then
		adminHandling = CDbl(pvnt)
	Else
		adminHandling = 0
	End If
End Function

Function adminHandlingIsActive()
	adminHandlingIsActive = getAppSetting("adminHandlingIsActive")
End Function

Function adminHandlingType()
	adminHandlingType = getAppSetting("adminHandlingType")
End Function

Function adminCODAmount()
	adminCODAmount = getAppSetting("adminCODAmount")
End Function

Function adminOriginState()
	adminOriginState= getAppSetting("adminOriginState")
End Function

Function adminOriginCountry()
	adminOriginCountry= getAppSetting("adminOriginCountry")
End Function

Function adminOriginZip()
	adminOriginZip= getAppSetting("adminOriginZip")
End Function

Function adminUsPsUserName()
	adminUsPsUserName= getAppSetting("adminUsPsUserName")
End Function

Function adminUsPsPassword()
	adminUsPsPassword= getAppSetting("adminUsPsPassword")
End Function

Function ltlEmail()
	ltlEmail= getAppSetting("ltlEmail")
End Function

Function ltlUN()
	ltlUN= getAppSetting("ltlUN")
End Function

Function UPSUserName()
	UPSUserName= getAppSetting("UPSUserName")
End Function

Function UPSPassword()
	UPSPassword= getAppSetting("UPSPassword")
End Function

Function UPSAccessKey()
	UPSAccessKey= getAppSetting("UPSAccessKey")
End Function

Function adminBackOrderBilling()
	adminBackOrderBilling= getAppSetting("adminBackOrderBilling")
End Function

Function adminTechnicalEmail()
	adminBackOrderBilling= getAppSetting("adminTechnicalEmail")
End Function

Function adminTermsAndConditions()
	adminTermsAndConditions= getAppSetting("adminTermsAndConditions")
End Function

Function adminTermsAndConditionsIsactive()
	adminTermsAndConditionsIsactive= getAppSetting("adminTermsAndConditionsIsactive")
End Function

Function adminMinOrderAmount()
Dim pvnt: pvnt = getAppSetting("adminMinOrderAmount")
	If Len(pvnt) > 0 And isNumeric(pvnt) Then
		adminMinOrderAmount = CDbl(pvnt)
	Else
		adminMinOrderAmount = 0
	End If
End Function

Function adminMinOrderMessage()
	adminMinOrderMessage= getAppSetting("adminMinOrderMessage")
End Function

Function FirstTimeCustomerDiscount()
	FirstTimeCustomerDiscount= getAppSetting("adminGlobalSaleAmt")
End Function

Function FirstTimeCustomerDiscount_IsPercent()
	FirstTimeCustomerDiscount_IsPercent= getAppSetting("adminGlobalSaleIsActive")
End Function

Function adminGlobalConfirmationMessage()
	adminGlobalConfirmationMessage= getAppSetting("adminGlobalConfirmationMessage")
End Function

Function adminGlobalConfirmationMessageIsactive()
	adminGlobalConfirmationMessageIsactive= ConvertToBoolean(getAppSetting("adminGlobalConfirmationMessageIsactive"), False)
End Function

'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------

Function CartName()
	CartName= Application("CartName")
End Function

'--------------------------------------------------------------------------------------------------
' Properties from configuration table
'--------------------------------------------------------------------------------------------------

'********************************************************************************

Sub LoadCustomStoreConfigurationSettings

	mstrBodyStyle = getConfigurationSettingFromCache("BodyStyle", "")
	cstrTDBottomNavStyle = getConfigurationSettingFromCache("BottomNavStyle", " id=""tdBottomNav""")
	cstrCartDisplay_MfgName = getConfigurationSettingFromCache("CartDisplay_MfgName", "")
	cstrCartDisplay_ProductID = getConfigurationSettingFromCache("CartDisplay_ProductID", "")
	cstrCCV_Optional = ConvertToBoolean(getConfigurationSettingFromCache("CCV_Optional", True), False)
	cstrCCV_SaveToDB = ConvertToBoolean(getConfigurationSettingFromCache("CCV_SaveToDB", True), False)
	cstrCCVFieldName = getConfigurationSettingFromCache("CCVFieldName", "")
	cstrCODTerm = getConfigurationSettingFromCache("CODTerm", "COD")
	cstrTDContentStyle = getConfigurationSettingFromCache("ContentStyle", " id=""tdContent""")
	cssDefaultMaxRecords = CLng(getConfigurationSettingFromCache("DefaultMaxRecords", 0))
	cssDefaultPageSize = CLng(getConfigurationSettingFromCache("DefaultPageSize", 12))
	cblnDisableLogin = ConvertToBoolean(getConfigurationSettingFromCache("DisableLogin", True), False)
	cblnDisplayCategoryDescriptions = ConvertToBoolean(getConfigurationSettingFromCache("DisplayCategoryDescriptions", True), False)
	cblnDisplayProductIDs = ConvertToBoolean(getConfigurationSettingFromCache("DisplayProductIDs", True), False)
	cblnExpandCategoriesByDefault = ConvertToBoolean(getConfigurationSettingFromCache("ExpandCategoriesByDefault", True), False)
	cstrECheckTerm = getConfigurationSettingFromCache("ECheckTerm", "eCheck")
	C_FORMDESIGN = getConfigurationSettingFromCache("FormFieldStyle", "BACKGROUND-COLOR:; FONT-FAMILY:; FONT-SIZE:pt;")
	cstrGoogleAdwords_conversion_id = getConfigurationSettingFromCache("GoogleAdwords", "")
	cstrGoogleAnalytics_uacct = getConfigurationSettingFromCache("GoogleAnalytics", "")
	cstrHighlightSearchTermClass = getConfigurationSettingFromCache("HighlightSearchTermClass", "highlightSearch")
	cstrImageDetailInstructions = getConfigurationSettingFromCache("ImageDetailInstructions", "")
	cblnIncludeEmailVerification = ConvertToBoolean(getConfigurationSettingFromCache("IncludeEmailVerification", True), False)
	cstrLargeWindowImageField = getConfigurationSettingFromCache("LargeWindowImageField", "")
	cstrOvertureID = getConfigurationSettingFromCache("OvertureID", "")
	cstrTDLeftNavStyle = getConfigurationSettingFromCache("LeftNavStyle", " id=""tdLeftNav""")
	clngMaxLengthDescription = CLng(getConfigurationSettingFromCache("MaxLengthDescription", 255))
	cstrNewProductField = getConfigurationSettingFromCache("NewProductField", "prodDateModified")
	clngNewProductsDaysSinceAdded = getConfigurationSettingFromCache("NewProductsDaysSinceAdded", 30)
	clngNumEntries = CLng(getConfigurationSettingFromCache("NumEntries", 10))
	
	If Len(getConfigurationSettingFromCache("OrderViewImageSrc", 8)) > 0 Then
		cbytOrderViewImageSrc = CLng(getConfigurationSettingFromCache("OrderViewImageSrc", 8))
	Else
		cbytOrderViewImageSrc = ""
	End If
	
	C_BNRBKGRND = getConfigurationSettingFromCache("PageBackground", "")
	C_WIDTH = getConfigurationSettingFromCache("PageWidth", "100%")
	cstrPhoneFaxTerm = getConfigurationSettingFromCache("PhoneFaxTerm", "PhoneFax")
	cstrPOTerm = getConfigurationSettingFromCache("POTerm", "PO")
	cstrPrimaryEmailToSendErrorTo = getConfigurationSettingFromCache("PrimaryEmailToSendErrorTo", "")
	cstrTDRightNavStyle = getConfigurationSettingFromCache("RightNavStyle", " id=""tdRightNav""")
	cblnSearchAttributes = ConvertToBoolean(getConfigurationSettingFromCache("SearchAttributes", True), False)

	If Len(getConfigurationSettingFromCache("SearchResultsDisplayType", 0)) > 0 Then
		cbytSearchResultsDisplayType = CLng(getConfigurationSettingFromCache("SearchResultsDisplayType", 0))
	Else
		cbytSearchResultsDisplayType = 0
	End If

	cstrSearchResultsDisplayTypeFixed = Trim(getConfigurationSettingFromCache("SearchResultsDisplayTypeFixed", "") & "")
	cblnShowOptionToTurnOffImages = ConvertToBoolean(getConfigurationSettingFromCache("ShowOptionToTurnOffImages", True), False)
	cblnShowSearchCustomizationOption = ConvertToBoolean(getConfigurationSettingFromCache("ShowSearchCustomizationOption", True), False)
	cstrSubWebPath = Trim(getConfigurationSettingFromCache("SubWebPath", "") & "")
	cstrTextLinkToLargerWindow = getConfigurationSettingFromCache("TextLinkToLargerWindow", "")
	cstrTimeoutRedirectPage = getConfigurationSettingFromCache("TimeoutRedirectPage", "order.asp")
	cstrTDTopNavStyle = getConfigurationSettingFromCache("TopNavStyle", " id=""tdTopNav""")
	cblnTrackPageViews = ConvertToBoolean(getConfigurationSettingFromCache("TrackPageViews", True), False)
	cstrURLTemplate_Manufacturer = getConfigurationSettingFromCache("URLTemplate_Manufacturer", "")
	cstrURLTemplate_Vendor = getConfigurationSettingFromCache("URLTemplate_Vendor", "")

	'Response.Write "cstrTDTopNavStyle: " & cstrTDTopNavStyle & "<br />"

End Sub	'LoadCustomStoreConfigurationSettings

'********************************************************************************

Function LoadConfigurationSettingsToCache()

Dim pobjRS
Dim pstrSQL
Dim pblnSuccess
Dim i
Dim pstrFieldName
Dim pvntTemp

	'ResetConfigurationSettingsInCache	'for testing
	maryConfigurationSettings = Application("ConfigurationSettings")
	If isArray(maryConfigurationSettings) Then
		'Response.Write "<h4>Application Settings loaded from cache . . .</h4>"
		LoadConfigurationSettingsToCache = True
		Exit Function
	End If

	'Response.Write "<h4>Loading Application Settings to cache . . .</h4>"

	pblnSuccess = False
	
	If Err.number <> 0 Then Err.Clear
	
	If StoreID = 0 Then
		pstrSQL = "SELECT configName, configValue FROM ssConfigurationSettings Where storeID = 1"
	Else
		pstrSQL = "SELECT configName, configValue FROM ssConfigurationSettings Where storeID = " & StoreID
	End If
	
	Set	pobjRS = CreateObject("adodb.recordset")
	With pobjRS
		.CursorLocation = 2 'adUseClient
	    
		'On Error Resume Next
		.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		If Err.number <> 0 Then
			Response.Write "<font color=red>Error in LoadConfigurationSettingsToCache: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
			Response.Write "<font color=red>LoadConfigurationSettingsToCache: sql = " & pstrSQL & "</font><br />" & vbcrlf
			Response.Flush
			Err.Clear
		ElseIf Not .EOF Then
			maryConfigurationSettings = .GetRows()
			pblnSuccess = True
		End If
		.Close
	End with
	Set pobjRS = Nothing
	
	Application.Lock()
	Application("ConfigurationSettings") = maryConfigurationSettings
	Application.UnLock
	
	LoadConfigurationSettingsToCache = pblnSuccess

End Function	' LoadConfigurationSettingsToCache

'********************************************************************************

Sub ResetConfigurationSettingsInCache
	Application.Lock()
	Application.Contents.Remove("ConfigurationSettings")	'for testing
	Application.UnLock
	Set maryConfigurationSettings = Nothing
	Call LoadConfigurationSettingsToCache
End Sub

'********************************************************************************

Function getConfigurationSettingFromCache(byVal strField, byVal vntDefault)

Dim i

	If Not isArray(maryConfigurationSettings) Then
		If Not LoadConfigurationSettingsToCache Then
			'Couldn't load configuration settings so use default and exit
			getConfigurationSettingFromCache = vntDefault
			Exit Function
		End If
	End If
	
	For i = 0 To UBound(maryConfigurationSettings,2)
		If maryConfigurationSettings(0,i) = strField Then
			If isNull(maryConfigurationSettings(1,i)) Then
				getConfigurationSettingFromCache = ""
			Else
				getConfigurationSettingFromCache = maryConfigurationSettings(1,i)
			End If
			Exit Function
		End If
	Next 'i
	getConfigurationSettingFromCache = vntDefault

End Function	' getConfigurationSettingFromCache
%>