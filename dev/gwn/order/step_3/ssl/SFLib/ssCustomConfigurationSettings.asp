<%
'********************************************************************************
'*   Custom Configuration Settings.asp for StoreFront 5.0 	                    *
'*   Release Version:	2.00.001 												*
'*   Release Date:		October 1, 2006											*
'*   Revision Date:		October 1, 2006											*
'*																				*
'*   Release 2.00.001 (October 1, 2006)											*
'*	   - Initial Release														*
'*                                                                              *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.				*
'********************************************************************************

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
Dim cstrECheckTerm
Dim C_FORMDESIGN
Dim cstrGoogleAdwords_conversion_id
Dim cstrGoogleAnalytics_uacct
Dim cstrHighlightSearchTermClass
Dim cstrImageDetailInstructions
Dim cblnIncludeEmailVerification
Dim cstrLargeWindowImageField
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
	cstrECheckTerm = getConfigurationSettingFromCache("ECheckTerm", "eCheck")
	C_FORMDESIGN = getConfigurationSettingFromCache("FormFieldStyle", "BACKGROUND-COLOR:; FONT-FAMILY:; FONT-SIZE:pt;")
	cstrGoogleAdwords_conversion_id = getConfigurationSettingFromCache("GoogleAdwords", "")
	cstrGoogleAnalytics_uacct = getConfigurationSettingFromCache("GoogleAnalytics", "")
	cstrHighlightSearchTermClass = getConfigurationSettingFromCache("HighlightSearchTermClass", "highlightSearch")
	cstrImageDetailInstructions = getConfigurationSettingFromCache("ImageDetailInstructions", "")
	cblnIncludeEmailVerification = ConvertToBoolean(getConfigurationSettingFromCache("IncludeEmailVerification", True), False)
	cstrLargeWindowImageField = getConfigurationSettingFromCache("LargeWindowImageField", "")
	cstrTDLeftNavStyle = getConfigurationSettingFromCache("LeftNavStyle", " id=""tdLeftNav""")
	clngMaxLengthDescription = CLng(getConfigurationSettingFromCache("MaxLengthDescription", 255))
	cstrNewProductField = getConfigurationSettingFromCache("NewProductField", "prodDateModified")
	clngNewProductsDaysSinceAdded = getConfigurationSettingFromCache("NewProductsDaysSinceAdded", 30)
	clngNumEntries = CLng(getConfigurationSettingFromCache("NumEntries", 10))
	cbytOrderViewImageSrc = CLng(getConfigurationSettingFromCache("OrderViewImageSrc", 8))
	C_BNRBKGRND = getConfigurationSettingFromCache("PageBackground", "")
	C_WIDTH = getConfigurationSettingFromCache("PageWidth", "100%")
	cstrPhoneFaxTerm = getConfigurationSettingFromCache("PhoneFaxTerm", "PhoneFax")
	cstrPOTerm = getConfigurationSettingFromCache("POTerm", "PO")
	cstrPrimaryEmailToSendErrorTo = getConfigurationSettingFromCache("PrimaryEmailToSendErrorTo", "")
	cstrTDRightNavStyle = getConfigurationSettingFromCache("RightNavStyle", " id=""tdRightNav""")
	cblnSearchAttributes = ConvertToBoolean(getConfigurationSettingFromCache("SearchAttributes", True), False)
	cbytSearchResultsDisplayType = CLng(getConfigurationSettingFromCache("SearchResultsDisplayType", 0))
	cstrSearchResultsDisplayTypeFixed = Trim(getConfigurationSettingFromCache("SearchResultsDisplayTypeFixed", "") & "")
	cblnShowOptionToTurnOffImages = ConvertToBoolean(getConfigurationSettingFromCache("ShowOptionToTurnOffImages", True), False)
	cblnShowSearchCustomizationOption = ConvertToBoolean(getConfigurationSettingFromCache("ShowSearchCustomizationOption", True), False)
	cstrSubWebPath = getConfigurationSettingFromCache("SubWebPath", "")
	cstrTextLinkToLargerWindow = getConfigurationSettingFromCache("TextLinkToLargerWindow", "")
	cstrTimeoutRedirectPage = getConfigurationSettingFromCache("TimeoutRedirectPage", "order.asp")
	cstrTDTopNavStyle = getConfigurationSettingFromCache("TopNavStyle", " id=""tdTopNav""")
	cblnTrackPageViews = ConvertToBoolean(getConfigurationSettingFromCache("TrackPageViews", True), False)
	cstrURLTemplate_Manufacturer = getConfigurationSettingFromCache("URLTemplate_Manufacturer", "")
	cstrURLTemplate_Vendor = getConfigurationSettingFromCache("URLTemplate_Vendor", "")

	'Response.Write "cstrTDTopNavStyle: " & cstrTDTopNavStyle & "<BR>"

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
			Response.Write "<font color=red>Error in LoadConfigurationSettingsToCache: Error " & Err.number & ": " & Err.Description & "</font><BR>" & vbcrlf
			Response.Write "<font color=red>LoadConfigurationSettingsToCache: sql = " & pstrSQL & "</font><BR>" & vbcrlf
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
			getConfigurationSettingFromCache = maryConfigurationSettings(1,i)
			Exit Function
		End If
	Next 'i
	getConfigurationSettingFromCache = vntDefault

End Function	' getConfigurationSettingFromCache

%>