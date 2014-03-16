<%
'********************************************************************************
'*   Postage Rate Component	for StoreFront 2000/5.0								*
'*   Release Version   2.00.017													*
'*   Release Date      October 27, 2002											*
'*   Revision Date     June 10, 2006											*
'*																				*
'*   Release Notes:                                                             *
'*																				*
'*   Release 2.00.017 (June 10, 2006)											*
'*	   - Added support for special, product/mfg specific ship methods			*
'*	   - Added support for product specific shipping price						*
'*																				*
'*   Release 2.00.016 (January 17, 2005)										*
'*	   - Added framework for custom shipping methods - NOT HOOKED UP			*
'*																				*
'*   Release 2.00.015 (October 5, 2004)											*
'*	   - Added additional debugging code										*
'*																				*
'*   Release 2.00.014 (May 22, 2004)											*
'*	   - Reviewed code for SQL Injection vulnerabilities						*
'*																				*
'*   Release 2.00.013 (August 10, 2003)											*
'*	   - Added Canada Post as a supported carrier								*
'*	   - Updated insurance calculation routines for multiple packages			*
'*																				*
'*   Release 2.0.12 (July 11, 2003)												*
'*	   - Updated FedEx module for orders over $1000								*
'*	   - Bug Fix: Updated FedEx module for international rates					*
'*																				*
'*   Release 2.0.11 (June 16, 2003)												*
'*	   - Updated preventing PO Boxes to include APO/AE/AA						*
'*	   - Added check for when IIS "loses" shipping address information			*
'*																				*
'*   Release 2.0.10 (May 11, 2003)												*
'*	   - Added support for preventing PO Boxes - requires custom implementation	*
'*																				*
'*   Release 2.0.9 (March 20, 2003)                                             *
'*	   - Added U.S.P.S. Media Mail Support										*
'*	   - Added support to break down large items below carrier limits			*
'*																				*
'*   Release 2.0.8 (March 1, 2003)                                              *
'*	   - General code cleanup													*
'*																				*
'*   Release 2.0.7 (February 28, 2003)                                          *
'*	   - Updated FedEx interface to match their revised API						*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.				*
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

        Const cblnGenericUSPSMethods = True
		Const cblnAutomaticallyFindAnyAvailable = True		'
		Const cstrLocalPickupCode = ""					'

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim maryShippingMethods
Dim mstrShippingCode
	
'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'*	Begin Page Code
'**********************************************************

	mstrShippingCode = Trim(Request.Form("Shipping"))
	
'***********************************************************************************************
'***********************************************************************************************

Class clsShipping

	'Internal Variables
	Private paryAvailableShippingMethods
	Private pobjDic
	Private pobjConn
	Private prsssShippingMethods
	
	Private pblnCustomSize	'Used for custom sizing
	
	'Input Variables
	
	'Origin
	Private pstrOriginStateAbb
	Private pstrOriginZip
	Private pstrOriginCountryAbb
	Private pblnInsidePickup
	Private pblnLoadingDockPickup
	
	'Destination
	Private pstrDestinationStateAbb
	Private pstrDestinationZIP
	Private pstrDestinationCountryAbb
	Private pstrDestinationCountryName
	Private pblnInsideDelivery
	Private pblnResidentialDelivery
	Private pblnLoadingDockDelivery
	
	'Package Features
	Private psngLength
	Private psngWidth
	Private psngHeight
	Private psngWeight
	Private pblnFixedSize		'i.e. if this is one item it cannot be broken into multiple products
	Private pblnDoNotCombine	'not applicable except for array below
	Private pdblMaxItemWeight	'used to limit available shipping methods
	Private pdblTotalOrderWeight	'used to limit available shipping methods
	
	'Array of packages (array of arrays with items from Package Features above)
	Private paryOrderItems
	
	'Order Paramaters
	Private psngOrderSubtotal		'used for free shipping determination
	Private psngDeclaredValue		'used for insurance calculation
	Private pblnInsured
	
	'Shipping Carrier choices
	Private pstrShippingSelection		'used for Freight Quote
	Private pstrssShippingMethodCode
	Private pstrssShippingMethodName
	
	Private pblnExceedsUPSMaxWeight
	Private pblnExceedsUSPSMaxWeight
	Private pblnExceedsFedExMaxWeight
	
	Private pblnShowFreightQuoteRates
	Private pblnFreightQuoteEnabled
	
	Private pbytMaxFreightQuoteRates
	Private pblnShowTransitTimes
	Private pblnUseOrderWeightToLimitCarrierChoices
	Private pstrProductBasedShippingCode
	Private pdblProductBasedShipping
	Private cblnAllowPackageBreakdown
	Private pstrEnabledSpecialShippingMethods
	Private pdblFixedShippingAmount
	
	Private pblnDebug
	
	Private enShippingCarrier_USPS
	Private enShippingCarrier_UPS
	Private enShippingCarrier_FedEx
	Private enShippingCarrier_CanadaPost
	Private enShippingCarrier_FreightQuote
	Private enShippingCarrier_Unknown
	Private enShippingCarrier_FlatRate
	Private enShippingCarrier_PerItem
	Private enShippingCarrier_PerPound

	Private Sub Class_Initialize

		pblnInsured = False
		pblnCustomSize = False
		pblnDebug = False
		pblnExceedsUPSMaxWeight = True
		pblnExceedsUSPSMaxWeight = True
		pblnExceedsFedExMaxWeight = True
		pdblProductBasedShipping = 0
		psngOrderSubtotal = 0
		psngDeclaredValue = 0
		pdblFixedShippingAmount = 0
		pblnResidentialDelivery = True
		pblnInsideDelivery = True
		
	'////////////////////////////////////////////////////////////////////////////////
	'//
	'//		USER CONFIGURATION

		pbytMaxFreightQuoteRates = 10
		pstrProductBasedShippingCode = ""		'Code of shipping method to use if any item exceed max weights
		pblnUseOrderWeightToLimitCarrierChoices = False	'set to true to limit carrier choices by total order weight, set to false to limit carrier choices by the maximum item weight
		pblnShowTransitTimes = False
		pblnFreightQuoteEnabled = False
		pblnShowFreightQuoteRates = False	'Only Show FreightQuote rates if selected or no other options
		cblnAllowPackageBreakdown = True

	'//
	'////////////////////////////////////////////////////////////////////////////////

	End Sub

	Public Property Let Connection(objConn)
		Set pobjConn = objConn
	End Property
	
	'Origin
	Public Property Let OriginStateAbb(strOriginStateAbb)
		pstrOriginStateAbb = strOriginStateAbb
	End Property

	Public Property Let OriginZip(strOriginZip)
		pstrOriginZip = strOriginZip
	End Property

	Public Property Let OriginCountryAbb(strOriginCountryAbb)
		pstrOriginCountryAbb = strOriginCountryAbb
	End Property

	Public Property Let InsidePickup(vntValue)
		pblnInsidePickup = vntValue
	End Property

	Public Property Let LoadingDockPickup(vntValue)
		pblnLoadingDockPickup = vntValue
	End Property

	'Destination
	Public Property Let DestinationStateAbb(strDestinationStateAbb)
		pstrDestinationStateAbb = strDestinationStateAbb
	End Property
	
	Public Property Let DestinationZIP(strDestinationZIP)
		pstrDestinationZIP = strDestinationZIP
	End Property
	
	Public Property Let DestinationCountryAbb(strDestinationCountryAbb)
		pstrDestinationCountryAbb = strDestinationCountryAbb
	End Property

	Public Property Let DestinationCountryName(vntValue)
		pstrDestinationCountryName = vntValue
	End Property

	Public Property Let InsideDelivery(vntValue)
		pblnInsideDelivery = vntValue
	End Property

	Public Property Let ResidentialDelivery(vntValue)
		pblnResidentialDelivery = vntValue
	End Property

	Public Property Let LoadingDockDelivery(vntValue)
		pblnLoadingDockDelivery = vntValue
	End Property

	'Package Features
	Public Property Let Length(sngLength)
		psngLength = sngLength
	End Property

	Public Property Let Width(sngWidth)
		psngWidth = sngWidth
	End Property

	Public Property Let Height(sngHeight)
		psngHeight = sngHeight
	End Property

	Public Property Let Weight(sngWeight)
		psngWeight = sngWeight
	End Property

	Public Property Let OrderItems(vntValue)
		paryOrderItems = vntValue
	End Property

	Public Property Let MaxItemWeight(dblMaxItemWeight)
		pdblMaxItemWeight = dblMaxItemWeight
	End Property

	Public Property Let TotalOrderWeight(dblTotalOrderWeight)
		pdblTotalOrderWeight = dblTotalOrderWeight
	End Property

	'Order Paramaters
	Public Property Let OrderSubtotal(sngOrderSubtotal)
		If isNumeric(sngOrderSubtotal) Then psngOrderSubtotal = sngOrderSubtotal
	End Property

	Public Property Let DeclaredValue(sngDeclaredValue)
		If isNumeric(sngDeclaredValue) Then psngDeclaredValue = sngDeclaredValue
	End Property

	Public Property Let Insured(blnInsured)
		pblnInsured = blnInsured
	End Property

	'Shipping Carrier choices
	Public Property Let ShippingSelection(strShippingSelection)
		pstrShippingSelection = strShippingSelection
	End Property
	
	Public Property Let ssShippingMethodCode(strssShippingMethodCode)
		pstrssShippingMethodCode = strssShippingMethodCode
	End Property
	Public Property Get ssShippingMethodCode()
		ssShippingMethodCode = pstrssShippingMethodCode
	End Property
	
	Public Property Let ssShippingMethodName(vntValue)
		pstrssShippingMethodName = vntValue
	End Property
	Public Property Get ssShippingMethodName()
		ssShippingMethodName = pstrssShippingMethodName
	End Property
	
	Public Property Let ShowFreightQuoteRates(vntValue)
		pblnShowFreightQuoteRates = vntValue
	End Property
	Public Property Get ShowFreightQuoteRates
		ShowFreightQuoteRates = pblnShowFreightQuoteRates
	End Property

	Public Property Get FreightQuoteEnabled
		FreightQuoteEnabled = pblnFreightQuoteEnabled
	End Property
	
	Public Property Get availableRates
		availableRates = paryAvailableShippingMethods
	End Property
	
	Public Property Let EnabledSpecialShippingMethods(byVal strValue)
		pstrEnabledSpecialShippingMethods = strValue
	End Property

	'***********************************************************************************************

	Public Function USPSInsurance()
	
		If psngDeclaredValue = 0 then
			USPSInsurance = 0
		Elseif psngDeclaredValue <= 50 then
			USPSInsurance = 1.30
		Elseif psngDeclaredValue <= 100 then
			USPSInsurance = 2.20
		Else
			USPSInsurance = int((psngDeclaredValue - 100) / 100 + 1.00) + 2.20
		End If
		
	End Function

	'***********************************************************************************************

	Public Function UPSInsurance()
		If psngDeclaredValue <= 100 then
			UPSInsurance = 0
		Else
			UPSInsurance = int((psngDeclaredValue - 100) / 100 + 0.5) * 0.35
		End If
	End Function

	'****************************************************************************************************************

	Public Sub checkForDefaultDestination()
	
		If Len(pstrDestinationCountryAbb) = 0 Then pstrDestinationCountryAbb = pstrOriginCountryAbb
		If pstrDestinationCountryAbb = "US" Then
			If Len(pstrDestinationStateAbb) = 0 Then pstrDestinationStateAbb = pstrOriginStateAbb
			If Len(pstrDestinationZIP) = 0 Then pstrDestinationZIP = pstrOriginZip
		Else
			If Len(pstrDestinationStateAbb) = 0 Then pstrDestinationStateAbb = pstrOriginStateAbb
			If Len(pstrDestinationZIP) = 0 Then pstrDestinationZIP = pstrOriginZip
		End If
		
	End Sub	'checkForDefaultDestination
	
	'***********************************************************************************************

	Private Function SetIndividualItem()
	
		'paryOrderItems(numItems)(8) decoder
		'0 - Length - default None
		'1 - Width - default None
		'2 - Height - default None
		'3 - Weight - default None
		'4 - Quantity - default None
		'5 - FixedSize - default True
		'6 - DoNotCombine - default False
		'7 - ProductBasedShipping - default None
		'8 - MustShipFreight - default False
		'9 - FixedShipping - default 0
		'10 - SpecialShipping - default None

		'psngLength, psngWidth, psngHeight, psngWeight, p_intQuantity, pblnFixedSize, pblnDoNotCombine, p_dblProductBasedShipping, pblnMustShipFreight, p_dblFixedShipping, p_strSpecialShipping

		If cblnDebugPostageRateAddon And True Then 
			Response.Write "<br /><hr><strong><font size='+2'>Setting Individual item . . .</font></strong><br />" & vbcrlf
			Response.Write "psngWeight =" & psngWeight & "<br />" & vbcrlf
			Response.Write "psngHeight =" & psngHeight & "<br />" & vbcrlf
		End If
		
		If Len(psngWeight) > 0 Then
			ReDim paryOrderItems(1)
			paryOrderItems(1) = Array (psngLength, psngWidth, psngHeight, psngWeight, 1, True, False, 0, False, 0, "")
			SetIndividualItem = True
		Else
			SetIndividualItem = False
		End If
		
	End Function	'SetIndividualItem

	'***********************************************************************************************
	
	Private Function poundsToKilos(byVal dblWeight)
		poundsToKilos = dblWeight/2.20462
	End Function

	Private Function inchesToCentimeters(byVal dblDimension)
		inchesToCentimeters = dblDimension*2.54
	End Function

	Public Sub GetCanadaPostRates(aryCarrierInfo)
	
		Dim pstrURL, pstrData
		Dim pclsShippingOption
		Dim pstrKey
		Dim pblnLoadError
		Dim pstrTrackingType
		Dim pstrRawData
		Dim i,j
		Dim pstrLanguage
		Dim pstrTurnAroundTime
		Dim pobjXMLDoc
		Dim nodeList,n,e
		Dim psngPostage
		Dim pblnError
		Dim pstrErrorDescription

		'//////////////////////////////////////////////////////////////////
		'/
		'/	set Canada Post Defaults
			
				pstrLanguage = "EN"
				pstrTurnAroundTime = 0
				
		'/
		'//////////////////////////////////////////////////////////////////

		Dim pstrTempState
		Dim p_strZip
		Dim pstrTempCountry
		
		pstrTempState = pstrDestinationStateAbb
		pstrTempCountry = pstrDestinationCountryAbb
		If pstrDestinationCountryAbb = "US" Then
			If Len(pstrDestinationZIP) > 5 Then
				p_strZip = left(pstrDestinationZIP,5)
			Else
				p_strZip = Trim(pstrDestinationZIP & "")
			End If
		Else
			If (pstrDestinationCountryAbb = "UK") OR (pstrDestinationCountryAbb = "EN") OR (pstrDestinationCountryAbb = "IE") OR (pstrDestinationCountryAbb = "SF") OR (pstrDestinationCountryAbb = "WL")  Then
				pstrTempCountry = "GB"
			End If
			If pstrDestinationCountryAbb <> "CA" Then pstrTempState = "xxx"
			p_strZip = Trim(pstrDestinationZIP & "")
		End If

		Dim pstrUsername
		Dim pstrPassword

		pstrUsername = aryCarrierInfo(1)
		pstrPassword = aryCarrierInfo(2)

		'Figure out oversize weight packages/ valid for ground only
		Dim psngTempWeight
		Dim plngNumPackages
		Dim psng_Length
		Dim psng_Height
		Dim psng_Width

'			'aryCarrierInfo
			'4 - ssShippingMethodPrefWeight
			'5 - ssShippingMethodMaxWeight
			'6 - ssShippingMethodMaxLength
			'7 - ssShippingMethodMaxWidth
			'8 - ssShippingMethodMaxHeight
			'9 - ssShippingMethodMaxGirth
			
		'check if any item exceeds max allowable weight
		If cblnDebugPostageRateAddon Then Response.Write "<br /><hr><strong><font size='+2'>Calculating <em>Canada Post</em> rates . . .</font></strong><br />" & vbcrlf
		pblnExceedsUPSMaxWeight = CalculatePackageSize(aryCarrierInfo(6),aryCarrierInfo(7),aryCarrierInfo(8),aryCarrierInfo(5),aryCarrierInfo(4),aryCarrierInfo(9), psng_Length, psng_Width, psng_Height, psngTempWeight, plngNumPackages)
		If pblnExceedsUPSMaxWeight Then 
			If cblnDebugPostageRateAddon Then Response.Write "Cannot ship via UPS - order size exceeds limits<br />"
			Exit Sub
		Else
			If cblnDebugPostageRateAddon Then Response.Write "<br /><strong><font size='+1'>Contacting <em>Canada Post</em> . . .</font></strong><br />" & vbcrlf
		End If
		
		'psngTempWeight = Round((psngTempWeight + 0.4999), 0)
		
		Dim psngTempInsuredAmount
		If pblnInsured Then
			psngTempInsuredAmount = FormatNumber(psngDeclaredValue / plngNumPackages,2,False, False, False)
		Else
			psngTempInsuredAmount = 0
		End If

		pstrURL = "http://206.191.4.228:30000"		'for testing
		'pstrURL = "http://216.191.36.73:30000"		'live
		pstrData = "<?xml version=" & chr(34) & "1.0" & chr(34) & " ?><!DOCTYPE eparcel SYSTEM " & chr(34) & "eParcel.dtd" & chr(34) & " >" _
						& "<eparcel>" _
						& "   <language>" & pstrLanguage & "</language> " _
						& "   <ratesAndServicesRequest>" _
						& "     <merchantCPCID>" & pstrUsername & "</merchantCPCID>" _
						& "		<fromPostalCode>" & pstrOriginZip & "</fromPostalCode>" _
						& "		<turnAroundTime>" & pstrTurnAroundTime & "</turnAroundTime>" _
						& "      <itemsPrice>" & psngDeclaredValue & "</itemsPrice>" _
						& "      <lineItems>" _
						& "         <item>" _
						& "            <quantity>1</quantity>" _
						& "            <weight>" & poundsToKilos(psngTempWeight) & "</weight>" _
						& "            <length>" & inchesToCentimeters(psng_Length) & "</length>" _
						& "            <width>" & inchesToCentimeters(psng_Width) & "</width>" _
						& "            <height>" & inchesToCentimeters(psng_Height) & "</height>" _
						& "            <description>Description</description>" _
						& "			   <imageURL>" & "@IMAGE_URL@" & "</imageURL>" _
						& "         </item>" _
						& "      </lineItems>" _
						& "      <provOrState>" & pstrTempState & "</provOrState>" _
						& "      <country>" & pstrTempCountry & "</country>" _
						& "      <postalCode>" & p_strZip & "</postalCode>" _
						& "   </ratesAndServicesRequest>" _
						& "</eparcel>"
'						& "      <city>" & pstrCity & "</city>" _

		If cblnDebugPostageRateAddon Then Response.Write "<fieldset><legend>Canada Post Data</legend>" & pstrData & "</fieldset>" & vbcrlf
		Call DebugRecordSplitTime("Contacting Canada Post . . .")
		pstrRawData = RetrieveRemoteData(pstrURL,pstrData,True)
		Call DebugRecordSplitTime("Canada Post responded")
		If cblnDebugPostageRateAddon Then Response.Write "<fieldset><legend>Canada Post Response</legend><pre>" & pstrRawData & "</pre></fieldset>" & vbcrlf

		set pobjXMLDoc = CreateObject("MSXML.DOMDocument")
		pobjXMLDoc.async = false
		pobjXMLDoc.resolveExternals = false
		pblnLoadError = not pobjXMLDoc.loadXML(pstrRawData)
			
		if pblnLoadError then
			pblnError = True
			pstrErrorDescription = "Unknown Error"
			Exit Sub
		End If
		
		Set nodeList = pobjXMLDoc.getElementsByTagName("error")
		If nodeList.length = 0 then
			set pobjDic = CreateObject("Scripting.Dictionary")
			Set nodeList = pobjXMLDoc.getElementsByTagName("product")
			For i =0 To nodeList.length - 1
				Set n = nodeList.Item(i)
				pstrKey = n.attributes.item(0).nodeValue
				For j =0 To n.childNodes.length - 1
					Set e = n.childNodes.Item(j)
					'If e.nodeName = "name" Then pstrKey = e.firstChild.nodeValue
					If e.nodeName = "rate" Then psngPostage = e.firstChild.nodeValue
				Next 'j

				if len(pstrKey) > 0 then
					set pclsShippingOption = New clsShippingOption
					pclsShippingOption.CarrierCode = enShippingCarrier_CanadaPost
					pclsShippingOption.ShippingCode = pstrKey
					pclsShippingOption.ShippingType = ShipCodeToMethod(pstrKey)
					pclsShippingOption.ShippingID = ShipCodeToID(pstrKey)
					pclsShippingOption.Postage = cCur(psngPostage)
					pclsShippingOption.Insurance = CPInsurance(psngTempInsuredAmount)
					pclsShippingOption.NumPackages = plngNumPackages
					pobjDic.Add pstrKey,pclsShippingOption
					If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & psngPostage & "<br />"
				end if
			Next 'i
		Else
			Set nodeList = pobjXMLDoc.getElementsByTagName("statusMessage")
			pblnError = True
			pstrErrorDescription = nodeList(0).text
			Response.Write "<fieldset><legend>Error Contacting Canada Post</legend><font color=red>" & pstrErrorDescription & "</font></fieldset><br />"
		End If

		If cblnDebugPostageRateAddon Then Response.Write "<br /><strong>Done with <em>Canada Post</em></strong><br />"

	End Sub	'GetCanadaPostRates

	'***********************************************************************************************

	Public Function CPInsurance(byVal sngDeclaredValue)
		If sngDeclaredValue <= 100 then
			CPInsurance = 0
		Else
			CPInsurance = int((sngDeclaredValue - 100) / 100 + 0.5) * 0.55
		End If
	End Function

	'***********************************************************************************************

	Public Sub GetFedExRates(aryCarrierInfo)
	
		Dim pstrUsername
		Dim pstrPassword

		pstrUsername = aryCarrierInfo(1)
		pstrPassword = aryCarrierInfo(2)

		Dim pstrURL, pstrData
		Dim pclsShippingOption
		Dim pstrKey
		Dim pstrFedExCode
		Dim pstrResponseText
		Dim psngPostage,psngInsurance
		Dim plngStart,plngEnd
			
		Dim pstrRateChart
		
'//////////////////////////////////////////////////////////////////
'/
'/	set FedEx Defaults
	
		pstrRateChart = "HomeD"
'/
'//////////////////////////////////////////////////////////////////

		Dim p_strZip
		If pstrDestinationCountryAbb = "US" Then
			If Len(pstrDestinationZIP) > 5 Then
				p_strZip = left(pstrDestinationZIP,5)
			Else
				p_strZip = Trim(pstrDestinationZIP & "")
			End If
		Else
			p_strZip = Trim(pstrDestinationZIP & "")
		End If

		Dim paryShipCodes
		paryShipCodes = Split(aryCarrierInfo(3),",")

		'correct for England/U.K.
		If (pstrDestinationCountryAbb = "UK") OR (pstrDestinationCountryAbb = "EN") OR (pstrDestinationCountryAbb = "IE") OR (pstrDestinationCountryAbb = "SF") OR (pstrDestinationCountryAbb = "WL")  Then pstrDestinationCountryAbb = "GB"

		If pstrDestinationCountryAbb = "US" Then
			For i = 0 To UBound(paryShipCodes)
				If Instr(1,paryShipCodes(i),"i") > 0 Then paryShipCodes(i) = ""
			Next 'i
		Else
			For i = 0 To UBound(paryShipCodes)
				If Instr(1,paryShipCodes(i),"i") < 1 Then paryShipCodes(i) = ""
			Next 'i
		End If
		
		'Figure out oversize weight packages/ valid for ground only
		Dim psngTempWeight
		Dim plngNumPackages
		Dim psng_Length
		Dim psng_Height
		Dim psng_Width

		'check if any item exceeds max allowable weight
		If cblnDebugPostageRateAddon Then Response.Write "<br /><hr><strong><font size='+2'>Calculating <em>FedEx</em> rates . . .</font></strong><br />" & vbcrlf
		pblnExceedsFedExMaxWeight = CalculatePackageSize(aryCarrierInfo(6),aryCarrierInfo(7),aryCarrierInfo(8),aryCarrierInfo(5),aryCarrierInfo(4),aryCarrierInfo(9), psng_Length, psng_Width, psng_Height, psngTempWeight, plngNumPackages)
		'Response.Write "pblnExceedsFedExMaxWeight: " & pblnExceedsFedExMaxWeight & "<br />" & vbcrlf
		If pblnExceedsFedExMaxWeight Then 
			If cblnDebugPostageRateAddon Then Response.Write "Cannot ship via FedEx - order size exceeds limits<br />"
			Exit Sub
		Else
			If cblnDebugPostageRateAddon Then Response.Write "<br /><strong><font size='+1'>Contacting <em>FedEx</em> . . .</font></strong><br />" & vbcrlf
		End If

		'this was brought in because FedEx only supports secure comm which requires the ServerXMLHTTP
		Dim pobjXMLHTTP
		
		'set timeouts in milliseconds
		Const resolveTimeout = 1000
		Const connectTimeout = 1000
		Const sendTimeout = 1000
		Const receiveTimeout = 10000
		Dim plngCounter
		Dim pblnEven
		Dim plngPos1
		Dim paryResults
		Dim pstrMeterNumber
		Dim pstrSubscription
		Dim i,j
		Dim pstrFedExServiceCode
		
		'for testing
		'pstrURL = "https://gatewaybeta.fedex.com:443/GatewayDC"
		
		'live
		pstrURL = "https://gateway.fedex.com:443/GatewayDC"
		
		'For i = 0 To UBound(paryShipCodes)
		'	If Len(paryShipCodes(i)) > 0 Then debugprint i,paryShipCodes(i)
		'Next 'i
		
		'Set minimum sizes
		If psngTempWeight < 1 Then psngTempWeight = 1
		If psng_Height < 2 Then psng_Height = 7
		If psng_Width < 4 Then psng_Width = 7
		If psng_Length < 7 Then psng_Length = 7
		
		'Round dimensions
		psng_Height = Round(psng_Height,0)
		psng_Width = Round(psng_Width,0)
		psng_Length = Round(psng_Length,0)

		'debugprint "psngTempWeight",psngTempWeight

		For i = 0 To UBound(paryShipCodes)
			If Len(paryShipCodes(i)) > 0 Then
				pstrKey =  Trim(paryShipCodes(i))
				'debugprint "pstrKey",pstrKey & ": " & ShipCodetoMethod(pstrKey)
				
				Select Case pstrKey
					Case "01","03","05","06","20","70","80","83","86"
						pstrFedExServiceCode = "FDXE"
						pstrFedExCode = pstrKey
					Case "01i","03i","06i","70i","86i"
						pstrFedExServiceCode = "FDXE"
						pstrFedExCode = Replace(pstrKey, "i", "")
					Case "92i"
						pstrFedExServiceCode = "FDXG"
						pstrFedExCode = Replace(pstrKey, "i", "")
					Case "90","92"
						pstrFedExServiceCode = "FDXG"
						pstrFedExCode = pstrKey
					Case Else
						pstrFedExServiceCode = "FDXE"
						pstrFedExCode = pstrKey
				End Select
'Notes:
'0: Can be FDXG or FDXE
'3025: Can be FDXG or FDXE
'1274
				pstrData = "0," & (Chr(34) & "022" & Chr(34)) _
						& "1," & (Chr(34) & "Rate Transaction" & Chr(34)) _
						& "10," & (Chr(34) & pstrUsername & Chr(34)) _
						& "23," & (Chr(34) & "1" & Chr(34)) _
						& "498," & (Chr(34) & pstrPassword & Chr(34)) _
						& "3025," & (Chr(34) & pstrFedExServiceCode & Chr(34)) _
						& "8," & (Chr(34) & pstrOriginStateAbb & Chr(34)) _
						& "9," & (Chr(34) & pstrOriginZip & Chr(34)) _
						& "117," & (Chr(34) & pstrOriginCountryAbb & Chr(34)) _
						& "16," & (Chr(34) & pstrDestinationStateAbb & Chr(34)) _
						& "17," & (Chr(34) & Replace(p_strZip, " ", "") & Chr(34)) _
						& "50," & (Chr(34) & pstrDestinationCountryAbb & Chr(34)) _
						& "57," & (Chr(34) & psng_Height & Chr(34)) _
						& "58," & (Chr(34) & psng_Width & Chr(34)) _
						& "59," & (Chr(34) & psng_Length & Chr(34)) _
						& "68," & (Chr(34) & "USD" & Chr(34)) _
						& "1415," & (Chr(34) & FormatNumber(psngDeclaredValue,2,,,False) & Chr(34)) _
						& "1401," & (Chr(34) & FormatNumber(psngTempWeight,1) & Chr(34)) _
						& "75," & (Chr(34) & "LBS" & Chr(34)) _
						& "116," & (Chr(34) & "1" & Chr(34)) _
						& "1116," & (Chr(34) & "I" & Chr(34)) _
						& "1273," & (Chr(34) & "01" & Chr(34)) _
						& "1274," & (Chr(34) & pstrFedExCode & Chr(34))
						  
				Select Case pstrKey
					Case "90","92"
						pstrData = pstrData & "3020," & (Chr(34) & "1" & Chr(34))
					Case Else
						pstrFedExServiceCode = "FDXE"
				End Select

				If pblnResidentialDelivery Then pstrData = pstrData & "440," & (Chr(34) & "Y" & Chr(34))

				'If pblnCODInd Then pstrData = pstrData & "27," & (Chr(34) & "Y" & Chr(34))
				'If pblnInsidePickup Then pstrData = pstrData & "440," & (Chr(34) & "Y" & Chr(34))
				'If pblnInsideDelivery Then pstrData = pstrData & "440," & (Chr(34) & "Y" & Chr(34))
				
				pstrData = pstrData & "99," & (Chr(34) & "" & Chr(34))
				'debugprint "pstrData",pstrData
			
				On Error Resume Next

				If Err.number <> 0 Then	Err.Clear
				'Use MSXML2 if possible - must have the Microsoft XML Parser v3 or later installed
				Set pobjXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
				If Err.number <> 0 Then
					Err.Clear
					Set pobjXMLHTTP = CreateObject("Microsoft.XMLHTTP")
				End If
				For plngCounter = 0 To 1	'added because of unexplained error on the first call
					With pobjXMLHTTP
						'.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
						.open "POST", pstrURL, False
						.setRequestHeader "Referer", pstrUsername
						.setRequestHeader "Host", "gateway.fedex.com:443"
						.setRequestHeader "Accept","image/gif, image/jpeg, image/pjpeg, text/plain,text/html, */*"
						.setRequestHeader "Content-Type","image/gif"
						.setRequestHeader "Content-Length", CStr(Len(pstrData))
						Call DebugRecordSplitTime("Contacting FedEx . . .")
						.send pstrData
						Call DebugRecordSplitTime("FedEx responded")
						pstrResponseText  = .responseText
						'Response.Write "() Error:" & Err.number & " - " & Err.Description & " (" & Err.Source & ")" & "<br />"	
						If cblnDebugPostageRateAddon And True Then 
							Response.Write "<b>" & ShipCodetoMethod(pstrKey) & "</b> (" & pstrKey & ")<br />" & vbcrlf
							Response.Write "FedExServiceCode =" & pstrFedExServiceCode & "<br />" & vbcrlf
							Response.Write "pstrData =" & pstrData & "<br />" & vbcrlf
							Response.Write "responseText =" & .responseText & "<br />" & vbcrlf
							Response.Flush
						End If
						'Response.Write "pstrResponseText =" & pstrResponseText & "<br />" & vbcrlf
						If cblnDebugPostageRateAddon Then Response.Write "<fieldset><legend>FedEx Response</legend><pre>" & pstrResponseText & "</pre></fieldset><br />"
					End With
					If Err.number <> -2147467259 Then Exit For
				Next 'plngCounter
				set pobjXMLHTTP = nothing

				If Err.number <> 0 Then
					Select Case Err.number
						Case -2147467259:	'Unspecified error 
							'This has only been seen when the server permissions are set incorrectly (access is denied msxml3.dll)
							Response.Write "<h3><font color=red>Server permissions error: <br /><i>msxml3.dll error '80070005'<br />Access is denied.</i><br />Please contact your server administrator</font></h3>"
							Response.Write "<h3><font color=red>Error " & Err.number & ": " & Err.Description & "</font></h3>"
						Case 438: 'Object doesn't support this property or method
							'This is from the set timeouts, no action required
						Case Else
							Response.Write "<h3><font color=red>Error " & Err.number & ": " & Err.Description & "</font></h3>"
					End Select
					Err.Clear
				End If
				
				'On Error Goto 0

				'Now split out the results
				plngPos1 = 1
				Do While plngPos1 > 0
					plngPos1 = Instr(plngPos1,pstrResponseText,Chr(34))
					If plngPos1 > 0 Then
						If pblnEven Then
							plngPos1 = plngPos1 + 1
						Else
							pstrResponseText = Left(pstrResponseText,plngPos1 - 1) & Right(pstrResponseText,Len(pstrResponseText)-plngPos1)
						End If
						pblnEven = Not pblnEven
					End If
					'Response.Write "pstrResponseText =" & pstrResponseText & "<br />" & vbcrlf
				Loop
				
				paryResults = Split(pstrResponseText,Chr(34))
				
				For j = 0 To UBound(paryResults)-1
					paryResults(j) = Split(paryResults(j),",")
				Next 'i

				psngPostage = ""
				For j = 0 To UBound(paryResults)-1
					Select Case paryResults(j)(0)
						Case "3":		'Error
							If Len(paryResults(j)(1)) > 0 Then Response.Write "<fieldset><legend>Error Contacting FedEx</legend><font color=red>" & paryResults(j)(1) & "</font></fieldset><br />"
						Case "1416":		'Base Rate
						Case "1417": 		'Surcharge
						Case "1418":		'Discount
						Case "1419": psngPostage = paryResults(j)(1)		'Net Amount
					End Select
					'Response.Write paryResults(j)(0) & "=" & paryResults(j)(1) & "<br />" & vbcrlf
					'Response.Flush
				Next 'j

				'Response.Flush
				'FedEx will return a 0 shipping rate for an invalid request
				If psngPostage = "0.00" Then psngPostage = ""
				psngInsurance = 0

				If (Len(pstrKey) > 0) And (Len(psngPostage) > 0) then
					set pclsShippingOption = New clsShippingOption
					pclsShippingOption.CarrierCode = enShippingCarrier_FedEx
					pclsShippingOption.ShippingCode = pstrKey
					pclsShippingOption.ShippingType = ShipCodeToMethod(pstrKey)
					pclsShippingOption.ShippingID = ShipCodeToID(pstrKey)
					pclsShippingOption.Postage = psngPostage
					pclsShippingOption.Insurance = psngInsurance
					pclsShippingOption.NumPackages = plngNumPackages
		'			pclsShippingOption.shipRate = ShipMultiple(pstrKey)
					pobjDic.Add pstrKey,pclsShippingOption
					If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & psngPostage & "<br />"
				end if
			End If
		Next 'i
		If cblnDebugPostageRateAddon Then Response.Write "<br /><strong>Done with <em>FedEx</em></strong><br />"

	End Sub	'GetFedExRates
	
	'***********************************************************************************************

	Public Sub GetUPSRates(aryCarrierInfo)
	
		Dim pstrURL, pstrData
		Dim pclsShippingOption
		Dim pstrKey
		Dim pblnLoadError
		Dim pstrTrackingType
		Dim pstrRawData
		Dim paryResult
		Dim i,j
		Dim aDetail
		Dim pintNumDetails
		Dim pbytService
	
		Dim pbytActionCode
		Dim pstrServiceLevelCode
		Dim pstrRateChart
		Dim pblnOversizeInd
		Dim pblnCODInd
		Dim pblnHazMat
		Dim pblnAdditionalHandlingInd
		Dim pbytCallTagARSInd
		Dim pblnSatDeliveryInd
		Dim pblnSatPickupInd
		Dim pbytDCISInd
		Dim pblnVerbalConfirmationInd
		Dim pbytSNDestinationInd1
		Dim pbytSNDestinationInd2
		Dim pblnReturnLabelInd
		Dim pblnResidentialInd
		Dim pstrPackagingType

		pbytActionCode = 4
		
'//////////////////////////////////////////////////////////////////
'/
'/	set UPS Defaults

		'Rate Chart Options
		'Regular+Daily+Pickup
		'On+Call+Air
		'One+Time+Pickup
		'Letter+Center
		'Customer+Counter

		pstrRateChart = "Regular+Daily+Pickup"
		pblnOversizeInd = False
		pblnCODInd = False
		pblnHazMat = False
		pblnAdditionalHandlingInd = False
		pbytCallTagARSInd = 0
		pblnSatDeliveryInd = False
		pblnSatPickupInd = False
		pbytDCISInd = 0
		pblnVerbalConfirmationInd = False
		pbytSNDestinationInd1 = 0
		pbytSNDestinationInd2 = 0
		pblnReturnLabelInd = False

		'Use the below line if you always want the Residential Rate used by default
		pblnResidentialInd = True
		pblnResidentialInd = pblnResidentialDelivery
		
		pstrPackagingType = "00"
		
'/
'//////////////////////////////////////////////////////////////////

		If Len(pstrssShippingMethodCode) = 0 then
		
			Select Case pstrDestinationCountryAbb
				Case "US": pstrServiceLevelCode = "3DS"	
				Case "CA": pstrServiceLevelCode = "STD"
				Case Else: pstrServiceLevelCode = "XPD"
			End Select
			pbytActionCode = 4
		Else
			'pstrServiceLevelCode = ShipIDtoCode(pstrssShippingMethodCode)
			pstrServiceLevelCode = pstrssShippingMethodCode
			pbytActionCode = 3
		End If

		Dim p_strZip
		If pstrDestinationCountryAbb = "US" Then
			If Len(pstrDestinationZIP) > 5 Then
				p_strZip = left(pstrDestinationZIP,5)
			Else
				p_strZip = Trim(pstrDestinationZIP & "")
			End If
		Else
			p_strZip = Trim(pstrDestinationZIP & "")
		End If

		Dim pstrUsername
		Dim pstrPassword

		pstrUsername = aryCarrierInfo(1)
		pstrPassword = aryCarrierInfo(2)

		'Figure out oversize weight packages/ valid for ground only
		Dim psngTempWeight
		Dim plngNumPackages
		Dim psng_Length
		Dim psng_Height
		Dim psng_Width

'			'aryCarrierInfo
			'4 - ssShippingMethodPrefWeight
			'5 - ssShippingMethodMaxWeight
			'6 - ssShippingMethodMaxLength
			'7 - ssShippingMethodMaxWidth
			'8 - ssShippingMethodMaxHeight
			'9 - ssShippingMethodMaxGirth
			
		'check if any item exceeds max allowable weight
		If cblnDebugPostageRateAddon Then Response.Write "<br /><hr><strong><font size='+2'>Calculating <em>UPS</em> rates . . .</font></strong><br />" & vbcrlf
		pblnExceedsUPSMaxWeight = CalculatePackageSize(aryCarrierInfo(6),aryCarrierInfo(7),aryCarrierInfo(8),aryCarrierInfo(5),aryCarrierInfo(4),aryCarrierInfo(9), psng_Length, psng_Width, psng_Height, psngTempWeight, plngNumPackages)
		If pblnExceedsUPSMaxWeight Then 
			If cblnDebugPostageRateAddon Then Response.Write "Cannot ship via UPS - order size exceeds limits<br />"
			Exit Sub
		Else
			If cblnDebugPostageRateAddon Then Response.Write "<br /><strong><font size='+1'>Contacting <em>UPS</em> . . .</font></strong><br />" & vbcrlf
		End If
		
		psngTempWeight = Round((psngTempWeight + 0.4999), 0)
		
		Dim psngTempInsuredAmount
		If pblnInsured Then
			psngTempInsuredAmount = FormatNumber(psngDeclaredValue / plngNumPackages,2,False, False, False)
		Else
			psngTempInsuredAmount = 0
		End If

		pstrURL = "http://www.ups.com/using/services/rave/qcost_dss.cgi"
		pstrData = "AppVersion=" & server.URLEncode("1.2") _
				  & "&AcceptUPSLicenseAgreement=yes" _
				  & "&ResponseType=" & server.URLEncode("application/x-ups-rss") _
				  & "&ActionCode=" & pbytActionCode _
				  & "&ServiceLevelCode=" & server.URLEncode(pstrServiceLevelCode) _
				  & "&RateChart=" & pstrRateChart _
				  & "&ShipperPostalCode=" & server.URLEncode(pstrOriginZip) _
				  & "&ConsigneePostalCode=" & server.URLEncode(p_strZip) _
				  & "&ConsigneeCountry=" & server.URLEncode(pstrDestinationCountryAbb) _
				  & "&PackageActualWeight=" & server.URLEncode(psngTempWeight) _
				  & "&DeclaredValueInsurance=" & server.URLEncode(psngTempInsuredAmount) _
				  & "&Length=" & server.URLEncode(psng_Length) _
				  & "&Width=" & server.URLEncode(psng_Width) _
				  & "&Height=" & server.URLEncode(psng_Height) _
				  & "&OversizeInd=" & cstr(pblnOversizeInd * -1) _
				  & "&CODInd=" & cstr(pblnCODInd * -1) _
				  & "&HazMat=" & cstr(pblnHazMat * -1) _
				  & "&AdditionalHandlingInd=" & cstr(pblnAdditionalHandlingInd * -1) _
				  & "&CallTagARSInd=" & pbytCallTagARSInd _
				  & "&SatDeliveryInd=" & cstr(pblnSatDeliveryInd * -1) _
				  & "&SatPickupInd=" & cstr(pblnSatPickupInd * -1) _
				  & "&DCISInd=" & server.URLEncode(pbytDCISInd) _
				  & "&VerbalConfirmationInd=" & cstr(pblnVerbalConfirmationInd * -1) _
				  & "&SNDestinationInd1=" & server.URLEncode(pbytSNDestinationInd1) _
				  & "&SNDestinationInd2=" & server.URLEncode(pbytSNDestinationInd2) _
				  & "&ReturnLabelInd=" & cstr(pblnReturnLabelInd * -1) _
				  & "&ResidentialInd=" & cstr(pblnResidentialInd * -1) _
				  & "&PackagingType=" & server.URLEncode(pstrPackagingType)

		'Response.Write pstrData & "<br />" & vbcrlf
		Call DebugRecordSplitTime("Contacting UPS . . .")
		pstrRawData = RetrieveRemoteData(pstrURL,pstrData,True)
		Call DebugRecordSplitTime("UPS responded")
		If cblnDebugPostageRateAddon Then Response.Write "<fieldset><legend>UPS. Response</legend><pre>" & pstrRawData & "</pre></fieldset><br />"
		paryResult = Split(pstrRawData, "--UPSBOUNDARY") 

		On Error Resume Next
		Dim pstrCommitTime

		For i = 2 to ubound(paryResult) - 1
			if instr(1,paryResult(i),"application/x-ups-error",1) = 0 then
				aDetail = split(paryResult(i),"%")
				set pclsShippingOption = New clsShippingOption
				pstrKey = aDetail(5)
				pclsShippingOption.CarrierCode = enShippingCarrier_UPS
				pclsShippingOption.ShippingCode = pstrKey
				pclsShippingOption.ShippingType = ShipCodeToMethod(pstrKey)
				pclsShippingOption.ShippingID = ShipCodeToID(pstrKey)
				pclsShippingOption.Postage = (cCur(aDetail(12)) + cCur(aDetail(13)))
				pclsShippingOption.Insurance = UPSInsurance
				pclsShippingOption.NumPackages = plngNumPackages
				pstrCommitTime = CDbl(Left(aDetail(15),2))
				If pstrCommitTime > 0 Then 
					If pstrCommitTime < 24 Then
						pclsShippingOption.TransitTime = 1
					Else
						pclsShippingOption.TransitTime = pstrCommitTime/24
					End If
				End If
				pobjDic.Add pstrKey,pclsShippingOption
				If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & pclsShippingOption.Postage & "<br />"
			end if
		Next 'i
		If cblnDebugPostageRateAddon Then Response.Write "<br /><strong>Done with <em>UPS</em></strong><br />"
		
	End Sub	'GetUPSRates

	'***********************************************************************************************

	Public Sub GetFlatRate(aryCarrierInfo)
		Call GetPerPound(aryCarrierInfo)
	End Sub	'GetFlatRate

	'***********************************************************************************************

	Public Sub GetPerItem(aryCarrierInfo)
		Call GetPerPound(aryCarrierInfo)
	End Sub	'GetPerItem

	'***********************************************************************************************

	Public Sub GetPerPound(aryCarrierInfo)
	
		Dim pclsShippingOption
		Dim pstrKey
		
		Dim psngPostage
			
		'check if any item exceeds max allowable weight
		If cblnDebugPostageRateAddon Then Response.Write "<br /><hr><strong><font size='+2'>Calculating <em>Special Carrier</em> rates . . .</font></strong><br />" & vbcrlf
		pblnExceedsUPSMaxWeight = CalculatePackageSize(aryCarrierInfo(6),aryCarrierInfo(7),aryCarrierInfo(8),aryCarrierInfo(5),aryCarrierInfo(4),aryCarrierInfo(9), psng_Length, psng_Width, psng_Height, psngTempWeight, plngNumPackages)
		If pblnExceedsUPSMaxWeight Then 
			If cblnDebugPostageRateAddon Then Response.Write "Cannot ship via Special - order size exceeds limits<br />"
			Exit Sub
		End If
		
		pstrKey = aryCarrierInfo(3)
		psngPostage = 1

		If len(pstrKey) > 0 then
			set pclsShippingOption = New clsShippingOption
			pclsShippingOption.CarrierCode = enShippingCarrier_Special
			pclsShippingOption.ShippingCode = pstrKey
			pclsShippingOption.ShippingType = ShipCodeToMethod(pstrKey)
			pclsShippingOption.ShippingID = ShipCodeToID(pstrKey)
			pclsShippingOption.Postage = cCur(psngPostage)
			pclsShippingOption.Insurance = 0
			pclsShippingOption.NumPackages = plngNumPackages
			pobjDic.Add pstrKey, pclsShippingOption
			If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & psngPostage & "<br />"
		end if

		If cblnDebugPostageRateAddon Then Response.Write "<br /><strong>Done with <em>Special</em></strong><br />"

	End Sub	'GetPerPound

	'***********************************************************************************************

	Public Sub GetUnknownRates(aryCarrierInfo)
	
		Dim pstrURL, pstrData
		Dim pclsShippingOption
		Dim pstrKey
		Dim pblnLoadError
		Dim pstrTrackingType
		Dim pstrRawData
		Dim i,j
		Dim pstrLanguage
		Dim pstrTurnAroundTime
		Dim pobjXMLDoc
		Dim nodeList,n,e
		Dim psngPostage
		Dim pblnError
		Dim pstrErrorDescription

		'Figure out oversize weight packages/ valid for ground only
		Dim psngTempWeight
		Dim plngNumPackages
		Dim psng_Length
		Dim psng_Height
		Dim psng_Width

'			'aryCarrierInfo
			'4 - ssShippingMethodPrefWeight
			'5 - ssShippingMethodMaxWeight
			'6 - ssShippingMethodMaxLength
			'7 - ssShippingMethodMaxWidth
			'8 - ssShippingMethodMaxHeight
			'9 - ssShippingMethodMaxGirth
			
		'check if any item exceeds max allowable weight
		If cblnDebugPostageRateAddon Then Response.Write "<br /><hr><strong><font size='+2'>Calculating <em>Unknown Carrier</em> rates . . .</font></strong><br />" & vbcrlf
		pblnExceedsUPSMaxWeight = CalculatePackageSize(aryCarrierInfo(6),aryCarrierInfo(7),aryCarrierInfo(8),aryCarrierInfo(5),aryCarrierInfo(4),aryCarrierInfo(9), psng_Length, psng_Width, psng_Height, psngTempWeight, plngNumPackages)
		If pblnExceedsUPSMaxWeight Then 
			If cblnDebugPostageRateAddon Then Response.Write "Cannot ship via Unknown - order size exceeds limits<br />"
			Exit Sub
		End If
		
		'psngTempWeight = Round((psngTempWeight + 0.4999), 0)
		
		pstrKey = aryCarrierInfo(3)
		psngPostage = 1

		if len(pstrKey) > 0 then
			set pclsShippingOption = New clsShippingOption
			pclsShippingOption.CarrierCode = enShippingCarrier_Unknown
			pclsShippingOption.ShippingCode = pstrKey
			pclsShippingOption.ShippingType = ShipCodeToMethod(pstrKey)
			pclsShippingOption.ShippingID = ShipCodeToID(pstrKey)
			pclsShippingOption.Postage = cCur(psngPostage)
			pclsShippingOption.Insurance = 0
			pclsShippingOption.NumPackages = plngNumPackages
			pobjDic.Add pstrKey, pclsShippingOption
			If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & psngPostage & "<br />"
		end if

		If cblnDebugPostageRateAddon Then Response.Write "<br /><strong>Done with <em>Unknown</em></strong><br />"

	End Sub	'GetUnknownRates

	'***********************************************************************************************
	Public Sub GetFreightQuoteRates(aryCarrierInfo, strShippingSelection)
	
		Dim pstrUsername
		Dim pstrPassword
		Dim p_bytMaxFreightQuoteRates

		pstrUsername = aryCarrierInfo(1)
		pstrPassword = aryCarrierInfo(2)
	
		dim i
		dim mstrURL
		dim mstrRequestData
		dim plngGirth
		dim pstrKey
		dim psngPostage
		dim pclsShippingOption
		Dim pstrSaveFQInfo
		Dim cstrServer
		
		If Len(strShippingSelection) > 0 Then
			Dim paryShippingChoices
			Dim paryShippingSelection
			paryShippingChoices = Split(Session("SaveFQInfo"),"|")
'debugprint "strShippingSelection",strShippingSelection			
			For i = 1 To UBound(paryShippingChoices)
				paryShippingSelection = Split(paryShippingChoices(i),",")
				If strShippingSelection = paryShippingSelection(0) Then
					pstrKey = paryShippingSelection(0)
					psngPostage = paryShippingSelection(2)
'debugprint "Carrier",paryShippingSelection(1)
'debugprint "Rate",psngPostage
					set pclsShippingOption = New clsShippingOption
					pclsShippingOption.CarrierCode = enShippingCarrier_FreightQuote
					pclsShippingOption.ShippingCode = "FreightQuote"
					pclsShippingOption.ShippingType = paryShippingSelection(1)
					pclsShippingOption.ShippingID = pstrKey
					pclsShippingOption.Postage = cCur(psngPostage)
					pclsShippingOption.Insurance = 0
					pclsShippingOption.NumPackages = 1
					pobjDic.Add pstrKey,pclsShippingOption
					If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & psngPostage & "<br />"
				End If
			Next 'i
			Exit Sub
		End If
	
		'user input
'//////////////////////////////////////////////////////////////////
'/
'/	set FreightQuote Defaults
	
		Dim pblnCODInd
		Dim pblnOriginLoadingDock
		Dim pblnOriginResidence
		Dim pblnOriginConstructionSite
		Dim pblnOriginInsidePickup
		Dim pblnDestLoadingDock
		Dim pblnDestResidence
		Dim pblnDestConstructionSite
		Dim pbytSNDestinationInd1
		Dim pbytSNDestinationInd2
		Dim pblnDestInsideDelivery

		cstrServer = "http://B2b.freightquote.com:4000"
		
		pblnOriginLoadingDock = "True"			'Indicates whether or not the shipping location has a loading dock
		pblnOriginResidence = "False"			'Indicates whether or not the shipping location is a residence
		pblnOriginConstructionSite = "False"	'Indicates whether or not the shipping location is a construction site
		pblnOriginInsidePickup = "False"		'Indicates whether or not the shipping location needs the shipment retrieved from inside the pickup location

		pblnDestLoadingDock = "False"			'Indicates whether or not the receiving location has a loading dock
		pblnDestResidence = "False"				'Indicates whether or not the receiving location is a residence
		pblnDestConstructionSite = "False"		'Indicates whether or not the receiving location is a construction site
		pblnDestInsideDelivery = "False"		'Indicates whether or not the receiving location needs the shipment retrieved from inside the pickup location

		pblnDestInsideDelivery = pblnInsideDelivery
		pblnDestResidence = pblnResidentialDelivery

'/
'//////////////////////////////////////////////////////////////////
	
		Dim p_strZip
		If Len(pstrDestinationZIP) > 5 Then
			p_strZip = left(pstrDestinationZIP,5)
		Else
			p_strZip = Trim(pstrDestinationZIP & "")
		End If

		mstrRequestData = "<FREIGHTQUOTE REQUEST=""QUOTE"" EMAIL=" & Chr(34) & pstrUsername & Chr(34) & " PASSWORD=" & Chr(34) & pstrPassword & Chr(34) & " BILLTO="">" _
						& "  <ORIGIN>" _
						& "    <ZIPCODE>" & pstrOriginZip & "</ZIPCODE>" _
						& "    <LOADINGDOCK>" & pblnOriginLoadingDock & "</LOADINGDOCK>" _
						& "    <RESIDENCE>" & pblnOriginLoadingDock & "</RESIDENCE>" _
						& "    <CONSTRUCTIONSITE>" & pblnOriginConstructionSite & "</CONSTRUCTIONSITE>" _
						& "    <INSIDEPICKUP>" & pblnOriginInsidePickup & "</INSIDEPICKUP>" _
						& "  </ORIGIN>" _
						& "  <DESTINATION>" _
						& "    <ZIPCODE>" & p_strZip & "</ZIPCODE>" _
						& "    <LOADINGDOCK>" & pblnDestLoadingDock & "</LOADINGDOCK>" _
						& "    <RESIDENCE>" & pblnDestResidence & "</RESIDENCE>" _
						& "    <CONSTRUCTIONSITE>" & pblnDestConstructionSite & "</CONSTRUCTIONSITE>" _
						& "    <INSIDEDELIVERY>" & pblnDestInsideDelivery & "</INSIDEDELIVERY>" _
						& "  </DESTINATION>" _
						& "  <SHIPMENT>" _
						& "    <WEIGHT>" & psngWeight & "</WEIGHT>" _
						& "    <CLASS></CLASS>" _
						& "    <DIMENSIONS>" _
						& "      <LENGTH>" & psngLength & "</LENGTH>" _
						& "      <WIDTH>" & psngWidth & "</WIDTH>" _
						& "      <HEIGHT>" & psngHeight & "</HEIGHT>" _
						& "    </DIMENSIONS>" _
						& "    <NMFC></NMFC>" _
						& "    <PRODUCTDESC>Misc Products</PRODUCTDESC>" _
						& "    <HZMT></HZMT>" _
						& "    <PACKAGETYPE></PACKAGETYPE>" _
						& "    <NUMBEROFPALLETS>1</NUMBEROFPALLETS>" _
						& "    <DECLAREDVALUE></DECLAREDVALUE>" _
						& "    <COMMODITYTYPE></COMMODITYTYPE>" _
						& "  </SHIPMENT>" _
						& "  <SERVICE>" _
						& "    <COD>" _
						& "      <AMOUNTTOCOLLECT></AMOUNTTOCOLLECT>" _
						& "      <REMITTONAME></REMITTONAME>" _
						& "      <REMITTOADDRESS></REMITTOADDRESS>" _
						& "      <REMITTOCITY></REMITTOCITY>" _
						& "      <REMITTOSTATE></REMITTOSTATE>" _
						& "      <REMITTOZIP></REMITTOZIP>" _
						& "    </COD>" _
						& "    <BLIND></BLIND>" _
						& "    <CONTENTTYPE></CONTENTTYPE>" _
						& "  </SERVICE>" _
						& "  <PALLETS_STACKABLE></PALLETS_STACKABLE>" _
						& "</FREIGHTQUOTE>"

		Dim pobjFreightQuote
		
		'Figure out oversize weight packages/ valid for ground only
		Dim psngTempWeight
		Dim plngNumPackages
		Dim psng_Length
		Dim psng_Height
		Dim psng_Width

		If cblnDebugPostageRateAddon Then Response.Write "<br /><hr><strong><font size='+2'>Calculating <em>Freight Quote</em> rates . . .</font></strong><br />" & vbcrlf
		Call CalculatePackageSize(1, 1, 1, 9999999999, 9999999999, 9999999999, psng_Length, psng_Width, psng_Height, psngTempWeight, plngNumPackages)
		If cblnDebugPostageRateAddon Then Response.Write "<br /><strong><font size='+1'>Contacting <em>FedEx</em> . . .</font></strong><br />" & vbcrlf

		On Error Resume Next
		Set pobjFreightQuote = CreateObject("FQRating.cFQRating")
		If err.number <> 0 Then
			Response.Write "<h3><font color=red>You need to install the Freight Quote COM object on the server</font></h3>"
			Exit Sub		
		End If
		On Error Goto 0

		With pobjFreightQuote
			.Email = pstrUsername
			.Password =  pstrPassword
			.oaddress.zip = pstrOriginZip
			.oAddress.LoadingDock = CStr(pblnOriginLoadingDock)
			.oAddress.Residence = CStr(pblnOriginResidence)
			
			.daddress.zip = p_strZip
			.dAddress.LoadingDock = CStr(Not pblnDestInsideDelivery)
			.dAddress.Residence = CStr(pblnDestResidence)

			'Populate the required Product Properties
			.FQProds.Class1 = 50
			.FQProds.Description1 = "Internet Order"
			.FQProds.PackageType1 = "Box"
			.FQProds.Pieces1 = 1
			.FQProds.Weight1 = psngTempWeight
			.BillTo = "SITE"

			.GetQuote

			If False Then	
				response.write "<table border=1 width=100% cellspacing=0 cellpadding=3 bgcolor=#ffffff0>"
				response.write "<tr>"
 				response.write "<td width=33% align=center><b><font color=#000000>Carrier Option</font></b></td>"
				response.write "<td width=33% align=center><b><font color=#000000>Rate</font></b></td>"
				response.write "<td width=34% align=center><b><font color=#000000>Transit Time (days)</font></b></td>"
				response.write "<td width=34% align=center><b><font color=#000000>Select Carrier</font></b></td>"
				response.write "</tr>"
				For i = 1 TO .FQResults.Count
					response.write "<tr>"
					response.write "<td width=33% align=center><i><font color=#000000>" & .FQResults.item(i).Carrier & "</font></i></td>"
					response.write "<td width=33% align=center><i><font color=#000000>" & .FQResults.item(i).Rate & "</font></i></td>"
					response.write "<td width=33% align=center><i><font color=#000000>" & .FQResults.item(i).Transit & "</font></i></td>"
					response.write "<td width=5% align=center><input type=checkbox name=Carrier onClick=selectcarrier(" & .FQResults.item(1).QuoteID & "," &  .FQResults.item(i).OptionID & ")></td>"
					response.write "</tr>"
				Next
				response.write "</table>"
			End If
			
			'pstrRawData = .XMLQuoteResponse
			
			If pblnDebug Then
				Response.Write ".XMLQuoteResponse = " & .XMLQuoteResponse & "<br />"
				Response.Write ".FQResults.Count = " & .FQResults.Count & "<br />"
				Response.Flush
			End If
			
			'limit the return
			If pbytMaxFreightQuoteRates < .FQResults.Count Then
				p_bytMaxFreightQuoteRates = pbytMaxFreightQuoteRates
			Else
				p_bytMaxFreightQuoteRates = .FQResults.Count
			End If

			For i = 1 TO p_bytMaxFreightQuoteRates
				pstrKey = "FreightQuotez" & .FQResults.item(1).QuoteID & "z" &  .FQResults.item(i).OptionID
				If pblnDebug Then Response.Write i & ": " & pstrKey & "<br />"
				psngPostage = .FQResults.item(i).Rate
				If (Len(pstrKey) > 0) And (Len(psngPostage) > 0) then
					If Not pobjDic.Exists(pstrKey) Then
						set pclsShippingOption = New clsShippingOption
						pclsShippingOption.CarrierCode = enShippingCarrier_FreightQuote
						pclsShippingOption.ShippingCode = "FreightQuote"
						pclsShippingOption.ShippingType = .FQResults.item(i).Carrier
						pclsShippingOption.ShippingID = pstrKey
						pclsShippingOption.TransitTime = .FQResults.item(i).Transit
						pclsShippingOption.Postage = cCur(psngPostage)
						pclsShippingOption.Insurance = 0
						pclsShippingOption.NumPackages = 1
						pobjDic.Add pstrKey,pclsShippingOption
						If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & psngPostage & "<br />"
						pstrSaveFQInfo = pstrSaveFQInfo & "|" & pstrKey & "," & .FQResults.item(i).Carrier & "," & .FQResults.item(i).Rate
						If pblnDebug Then Response.Write "ShippingCode = " & pclsShippingOption.ShippingCode & "<br />"
					End If
				End If
			Next
			Session("SaveFQInfo") = pstrSaveFQInfo
			
		End With
		If cblnDebugPostageRateAddon Then Response.Write "<br /><strong>Done with <em>Freight Quote</em></strong><br />"

	End Sub	'GetFreightQuoteRates

	'***********************************************************************************************

	Public Sub GetUSPSRates(aryCarrierInfo)
	
		dim pobjXMLDoc
		dim i,j,k
		dim nodeList,n,e

		dim pstrRawData
		dim mstrURL
		dim mstrRequestData
		dim plngGirth
		dim pstrKey
		dim psngPostage
		dim pclsShippingOption
		
		Dim cstrServer
	
		'user input
		Dim pstrGirth
		Dim pstrPackageID
		Dim pstrContainer
		Dim plngPounds
		Dim plngOunces
		Dim pstrSize
		Dim pblnMachinable
		Dim pstrMailType
	
		cstrServer = "Production.ShippingApis.com/ShippingAPI.dll?API="
		
		Dim pstrUsername
		Dim pstrPassword

		pstrUsername = aryCarrierInfo(1)
		pstrPassword = aryCarrierInfo(2)
	
		'set USPS Defaults
	
		pstrSize = "Regular"
		pstrContainer = "None"
		pblnMachinable = "True"
		pstrMailType = "Package"

		'Figure out oversize weight packages/ valid for ground only
		Dim psngTempWeight
		Dim plngNumPackages
		Dim psng_Length
		Dim psng_Height
		Dim psng_Width

		'check if any item exceeds max allowable weight
		If cblnDebugPostageRateAddon Then Response.Write "<br /><hr><strong><font size='+2'>Calculating <em>U.S.P.S.</em> rates . . .</font></strong><br />" & vbcrlf
		pblnExceedsUSPSMaxWeight = CalculatePackageSize(aryCarrierInfo(6),aryCarrierInfo(7),aryCarrierInfo(8),aryCarrierInfo(5),aryCarrierInfo(4),aryCarrierInfo(9), psng_Length, psng_Width, psng_Height, psngTempWeight, plngNumPackages)
		If pblnExceedsUSPSMaxWeight Then 
			If cblnDebugPostageRateAddon Then Response.Write "Cannot ship via U.S.P.S. - order size exceeds limits<br />"
			Exit Sub
		Else
			If cblnDebugPostageRateAddon Then Response.Write "<br /><strong><font size='+1'>Contacting <em>U.S.P.S.</em> . . .</font></strong><br />" & vbcrlf
		End If

		plngPounds = int(psngTempWeight)
		plngOunces = int((psngTempWeight-int(psngTempWeight)) * 16 + 0.99)
		
		plngGirth = psng_Length + 2 * (psng_Height + psng_Width)
		If plngGirth <= 84 then
			pstrSize = "Regular"
		Elseif plngGirth <= 108 then
			pstrSize = "Large"
		Elseif plngGirth <= 130 then
			pstrSize = "Oversize"
		End if

		if pstrDestinationCountryAbb = "US" then

			mstrURL = "http://" & cstrServer & "RateV4"
			
			Dim p_strZip
			If Len(pstrDestinationZIP) > 5 Then
				p_strZip = left(pstrDestinationZIP,5)
			Else
				p_strZip = Trim(pstrDestinationZIP & "")
			End If

			mstrRequestData = "<RateV4Request USERID=""" & pstrUsername & """>" _
					& "<Revision>4</Revision>" _
                    & "<Package ID=" & Chr(34) & "0" & Chr(34) & ">" _
					& "<Service>ALL</Service>" _
					& "<ZipOrigination>" & pstrOriginZip & "</ZipOrigination>" _
					& "<ZipDestination>" & p_strZip & "</ZipDestination>" _
					& "<Pounds>" & plngPounds & "</Pounds><Ounces>" & plngOunces & "</Ounces>" _
					& "<Container/>" _
					& "<Size>" & pstrSize & "</Size>" _
					& "<Width>" & FormatNumber(psng_Width, 1) & "</Width>" _
					& "<Length>" & FormatNumber(psng_Length, 1) & "</Length>" _
					& "<Height>" & FormatNumber(psng_Height, 1) & "</Height>" _
					& "<Girth>" & CLng(plngGirth) & "</Girth>" _
					& "<Machinable>true</Machinable>" _
					& "</Package>" _
					& "</RateV4Request>"

			Call DebugRecordSplitTime("Contacting U.S.P.S. . . .")
			pstrRawData = RetrieveRemoteData(mstrURL & "&XML=" & server.URLEncode(mstrRequestData),"",False)
			Call DebugRecordSplitTime("U.S.P.S. responded")
			If cblnDebugPostageRateAddon Then Response.Write "<fieldset><legend>U.S.P.S. Response</legend><pre>" & pstrRawData & "</pre></fieldset><br />"

			set pobjXMLDoc = CreateObject("MSXML.DOMDocument")
			pobjXMLDoc.async = false
			pobjXMLDoc.resolveExternals = false
			
			If pobjXMLDoc.loadXML(pstrRawData) Then
				If pobjXMLDoc.documentElement.nodeName <> "Error" then 'Top-level Error
					Set nodeList = pobjXMLDoc.getElementsByTagName("Package/Postage")
					For i = 0 To nodeList.length - 1
						Set n = nodeList.Item(i)
						For j = 0 To n.childNodes.length - 1
							Set e = n.childNodes.Item(j)
							If e.nodeName = "MailService" Then
                                pstrKey = Replace(e.firstChild.nodeValue, "&lt;sup&gt;&amp;reg;&lt;/sup&gt;", "")
                                pstrKey = Replace(pstrKey, "&lt;sup&gt;&#174;&lt;/sup&gt;", "")
                                pstrKey = Replace(pstrKey, "&lt;sup&gt;&#8482;&lt;/sup&gt;", "")
                            End If
							If e.nodeName = "Rate" Then psngPostage = e.firstChild.nodeValue
						Next 'j
						If (Len(pstrKey) > 0) And (Len(psngPostage) > 0) then
							If Not pobjDic.Exists(pstrKey) Then
								set pclsShippingOption = New clsShippingOption
								pclsShippingOption.CarrierCode = enShippingCarrier_USPS
								pclsShippingOption.ShippingCode = pstrKey
								pclsShippingOption.ShippingType = ShipCodeToMethod(pstrKey)
								pclsShippingOption.ShippingID = ShipCodeToID(pstrKey)
								pclsShippingOption.Postage = cCur(psngPostage)
								pclsShippingOption.Insurance = USPSInsurance
								pclsShippingOption.NumPackages = plngNumPackages
								pobjDic.Add pstrKey,pclsShippingOption
								If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & psngPostage & "<br />"
							End If
						end if
					Next 'i
				Else
					Set nodeList = pobjXMLDoc.getElementsByTagName("Description")
					Response.Write "<fieldset><legend>Error Contacting U.S.P.S.</legend><font color=red>" & nodeList(0).text & "</font></fieldset><br />"
					Set nodeList = Nothing
				End If	'pobjXMLDoc.documentElement.nodeName <> "Error"
			End If	'pobjXMLDoc.loadXML(pstrRawData)
		else
		
			Dim pstrTempCountry
			If (pstrDestinationCountryAbb = "UK") OR (pstrDestinationCountryAbb = "GB") OR (pstrDestinationCountryAbb = "EN") OR (pstrDestinationCountryAbb = "IE") OR (pstrDestinationCountryAbb = "SF") OR (pstrDestinationCountryAbb = "WL")  Then
				pstrTempCountry = "Great Britain and Northern Ireland"
			Else
				pstrTempCountry = pstrDestinationCountryName
			End If
		
			mstrURL = "http://" & cstrServer & "IntlRateV2"
			mstrRequestData = "<IntlRateV2Request USERID=" & Chr(34) & pstrUsername & Chr(34) & ">" _
							& "<Revision>2</Revision>" _
							& "<Package ID=""0"">" _
							& "<Pounds>" & plngPounds & "</Pounds><Ounces>" & plngOunces & "</Ounces>" _
							& "<Machinable>True</Machinable>" _
							& "<MailType>All</MailType>" _
							& "<GXG><POBoxFlag>N</POBoxFlag><GiftFlag>N</GiftFlag></GXG>" _
							& "<ValueOfContents>0</ValueOfContents>" _
							& "<Country>" & pstrTempCountry & "</Country>" _
							& "<Container>RECTANGULAR</Container>" _
					        & "<Size>" & pstrSize & "</Size>" _
							& "<Width>" & FormatNumber(psng_Width, 1) & "</Width>" _
							& "<Length>" & FormatNumber(psng_Length, 1) & "</Length>" _
							& "<Height>" & FormatNumber(psng_Height, 1) & "</Height>" _
							& "<Girth>" & CDbl(plngGirth) & "</Girth>" _
							& "</Package>" _
							& "</IntlRateV2Request>"

			Call DebugRecordSplitTime("Contacting U.S.P.S. . . .")
			pstrRawData = RetrieveRemoteData(mstrURL & "&XML=" & server.URLEncode(mstrRequestData),"",False)
			Call DebugRecordSplitTime("U.S.P.S. responded")
			If cblnDebugPostageRateAddon Then Response.Write "<fieldset><legend>U.S.P.S. Response</legend><pre>" & pstrRawData & "</pre></fieldset><br />"

			set pobjXMLDoc = CreateObject("MSXML.DOMDocument")
			pobjXMLDoc.async = false
			
			If pobjXMLDoc.loadXML(pstrRawData) Then
				Set nodeList = pobjXMLDoc.getElementsByTagName("Error")
				If nodeList.length = 0 Then 'Top-level Error
					Set nodeList = pobjXMLDoc.getElementsByTagName("Package")
					For i =0 To nodeList.length - 1
						Set n = nodeList.Item(i)
						For j = 0 To n.childNodes.length - 1
							Set e = n.childNodes.Item(j)
							if e.nodeName = "Service" Then Call ParseService(e, plngNumPackages)
						Next 'j
					Next 'i
				Else
					Set nodeList = pobjXMLDoc.getElementsByTagName("Description")
					Response.Write "<fieldset><legend>Error Contacting U.S.P.S.</legend><font color=red>" & nodeList(0).text & "</font></fieldset><br />"
					Set nodeList = Nothing
				End If
			End If
		end if
		set nodeList = Nothing
		set n = Nothing
		set e = Nothing

		If cblnDebugPostageRateAddon Then Response.Write "<br /><strong>Done with <em>U.S.P.S.</em></strong><br />"

	End Sub	'GetUSPSRates

	Private Sub ParseService(ochild, lngNumPackages)

	Dim i
	Dim s
	Dim psngPostage
	Dim pstrSvcDescription
	Dim pclsShippingOption
	Dim pstrKey
	
		For i =0 To ochild.childNodes.length - 1
			Set s = ochild.childNodes.Item(i)
			Select Case s.nodeName
				Case "Postage": psngPostage = s.firstChild.nodeValue
				Case "SvcDescription": 
                    pstrSvcDescription = Replace(s.firstChild.nodeValue, "&lt;sup&gt;&amp;reg;&lt;/sup&gt;", "")
                    pstrSvcDescription = Replace(pstrSvcDescription, "&lt;sup&gt;&amp;trade;&lt;/sup&gt;", "")

                    pstrSvcDescription = Replace(pstrSvcDescription, "&lt;sup&gt;&#174;&lt;/sup&gt;", "")
                    pstrSvcDescription = Replace(pstrSvcDescription, "&lt;sup&gt;&#8482;&lt;/sup&gt;", "")

                    'Remove ** if notes about customs
                    If Len(pstrSvcDescription) > 2 Then
                        If Right(pstrSvcDescription, 2) = "**" Then pstrSvcDescription = Left(pstrSvcDescription, Len(pstrSvcDescription) - 2)
                    End If
			End Select
		Next
		
		pstrKey = pstrSvcDescription
		set pclsShippingOption = New clsShippingOption
		pclsShippingOption.CarrierCode = 0
		pclsShippingOption.ShippingCode = pstrKey
		pclsShippingOption.ShippingType = ShipCodeToMethod(pstrKey)
		pclsShippingOption.ShippingID = ShipCodeToID(pstrKey)
		pclsShippingOption.Postage = cCur(psngPostage)
		pclsShippingOption.Insurance = USPSInsurance
		pclsShippingOption.NumPackages = lngNumPackages
		If cblnDebugPostageRateAddon Then Response.Write "&nbsp;&nbsp;" & pclsShippingOption.ShippingType & " (" & pstrKey & "): " & psngPostage & "<br />"
		pobjDic.Add pstrKey,pclsShippingOption

	End Sub	'ParseService

	'***********************************************************************************************

	Public Function GetRates(byVal strssShippingMethodCode)

		Dim pRS
		Dim pobjTempDic
		Dim pstrKey
		Dim pstrTempFedExIDs
		
		Dim i,j
		Dim plngCarrierID
		Dim paryCheckCarrier(11)
		Dim p_strGetRates
		
		'Check to make sure the product is loaded
		If Not isArray(paryOrderItems) Then
			If Not SetIndividualItem Then
				GetRates = "FAIL"
				Exit Function
			End If
		End If

		For i = 1 To UBound(paryCheckCarrier)
			paryCheckCarrier(i) = Array (False,"","","","","","","","","","")	'decoder: 0 - active, 1 - username, 2 - password, 3 - active carrier codes
		Next 'i
		
		If Len(psngLength) = 0 Or (psngLength = "0") then psngLength = 12
		If Len(psngHeight) = 0 Or (psngHeight = "0") then psngHeight = 12
		If Len(psngWidth) = 0 Or (psngWidth = "0") then psngWidth = 12
		If Len(psngDeclaredValue) = 0 then psngDeclaredValue = 0
		
		Set pobjDic = CreateObject("Scripting.Dictionary")

		pstrssShippingMethodCode = Trim(strssShippingMethodCode)
		
		'Added to include special shipping methods
		Call CalculatePackageSize(1,1,1,1,1,1,1,1,1,1,1)

		Call InitializeShipMethods(pstrssShippingMethodCode)
		
		If cblnDebugPostageRateAddon Then
			Response.Write "<br /><strong>Origin</strong><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;State: " & pstrOriginStateAbb & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;ZIP: " & pstrOriginZip & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Country: " & pstrOriginCountryAbb & "<br />" & vbcrlf
			Response.Write "<br /><strong>Destination</strong><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;State: " & pstrDestinationStateAbb & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;ZIP: " & pstrDestinationZIP & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Country: " & pstrDestinationCountryName & " (" & pstrDestinationCountryAbb & ")<br />" & vbcrlf
			
			Response.Write "<br /><strong>Order Information</strong><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Order Subtotal: " & psngOrderSubtotal & " (for free shipping)<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Declared Value: " & psngDeclaredValue & " (for insurance)<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;pblnInsured: " & pblnInsured & "<br />" & vbcrlf
		End If

		For i = 1 To prsssShippingMethods.RecordCount
			plngCarrierID = prsssShippingMethods.Fields("ssShippingCarrierID").Value

			'paryCheckCarrier(plngCarrierID) = Array (True,Trim(prsssShippingMethods.Fields("ssShippingCarrierUserName").Value & ""),Trim(prsssShippingMethods.Fields("ssShippingCarrierPassword").Value & ""))
			paryCheckCarrier(plngCarrierID)(0) = True
			paryCheckCarrier(plngCarrierID)(1) = Trim(prsssShippingMethods.Fields("ssShippingCarrierUserName").Value & "")
			paryCheckCarrier(plngCarrierID)(2) = Trim(prsssShippingMethods.Fields("ssShippingCarrierPassword").Value & "")

			If Len(paryCheckCarrier(plngCarrierID)(3)) > 0 Then
				paryCheckCarrier(plngCarrierID)(3) = paryCheckCarrier(plngCarrierID)(3) & "," & Trim(prsssShippingMethods.Fields("ssShippingCode").Value & "")
			Else
				paryCheckCarrier(plngCarrierID)(3) = Trim(prsssShippingMethods.Fields("ssShippingCode").Value & "")
			End If
			
			paryCheckCarrier(plngCarrierID)(4) = Trim(prsssShippingMethods.Fields("ssShippingMethodPrefWeight").Value & "")
			paryCheckCarrier(plngCarrierID)(5) = Trim(prsssShippingMethods.Fields("ssShippingMethodMaxWeight").Value & "")
			paryCheckCarrier(plngCarrierID)(6) = Trim(prsssShippingMethods.Fields("ssShippingMethodMaxLength").Value & "")
			paryCheckCarrier(plngCarrierID)(7) = Trim(prsssShippingMethods.Fields("ssShippingMethodMaxWidth").Value & "")
			paryCheckCarrier(plngCarrierID)(8) = Trim(prsssShippingMethods.Fields("ssShippingMethodMaxHeight").Value & "")
			paryCheckCarrier(plngCarrierID)(9) = Trim(prsssShippingMethods.Fields("ssShippingMethodMaxGirth").Value & "")
			paryCheckCarrier(plngCarrierID)(10) = Trim(prsssShippingMethods.Fields("ssShippingMethodMinWeight").Value & "")

			prsssShippingMethods.MoveNext
		Next 'i
		If prsssShippingMethods.RecordCount > 0 Then prsssShippingMethods.MoveFirst

		If pblnDebug Then
			Response.Write "prsssShippingMethods.RecordCount =" & prsssShippingMethods.RecordCount & "<br />" & vbcrlf
			Response.Write "pstrssShippingMethodCode =" & pstrssShippingMethodCode & "<br />" & vbcrlf
			Response.Write "pstrOriginZip =" & pstrOriginZip & "<br />" & vbcrlf
			Response.Write "pstrOriginCountryAbb =" & pstrOriginCountryAbb & "<br />" & vbcrlf
			Response.Write "pstrDestinationZIP =" & pstrDestinationZIP & "<br />" & vbcrlf
			Response.Write "pstrDestinationCountryAbb =" & pstrDestinationCountryAbb & "<br />" & vbcrlf
			
			Response.Write "psngLength =" & psngLength & "<br />" & vbcrlf
			Response.Write "psngWidth =" & psngWidth & "<br />" & vbcrlf
			Response.Write "psngHeight =" & psngHeight & "<br />" & vbcrlf
			Response.Write "psngWeight =" & psngWeight & "<br />" & vbcrlf
			Response.Write "pdblMaxItemWeight =" & pdblMaxItemWeight & "<br />" & vbcrlf

			For i = 1 To UBound(paryCheckCarrier)
				Response.Write i & ": " & paryCheckCarrier(i)(0) & "<br />" & vbcrlf
			Next 'i
			If Response.Buffer Then Response.Flush
		End If

		enShippingCarrier_USPS = 2
		enShippingCarrier_UPS = 3
		enShippingCarrier_FedEx = 4
		enShippingCarrier_CanadaPost = 5
		enShippingCarrier_FreightQuote = 8
		enShippingCarrier_Unknown = 1
		enShippingCarrier_FlatRate = 9
		enShippingCarrier_PerItem = 10
		enShippingCarrier_PerPound = 11

		If paryCheckCarrier(enShippingCarrier_USPS)(0) Then Call GetUSPSRates(paryCheckCarrier(enShippingCarrier_USPS))
		If paryCheckCarrier(enShippingCarrier_UPS)(0) Then Call GetUPSRates(paryCheckCarrier(enShippingCarrier_UPS))
		If paryCheckCarrier(enShippingCarrier_FedEx)(0) Then Call GetFedExRates(paryCheckCarrier(enShippingCarrier_FedEx))
		If paryCheckCarrier(enShippingCarrier_CanadaPost)(0) Then Call GetCanadaPostRates(paryCheckCarrier(enShippingCarrier_CanadaPost))
		If paryCheckCarrier(enShippingCarrier_Unknown)(0) Then Call GetUnknownRates(paryCheckCarrier(enShippingCarrier_Unknown))
		If paryCheckCarrier(enShippingCarrier_FlatRate)(0) Then Call GetFlatRate(paryCheckCarrier(enShippingCarrier_FlatRate))
		If paryCheckCarrier(enShippingCarrier_PerItem)(0) Then Call GetPerItem(paryCheckCarrier(enShippingCarrier_PerItem))
		If paryCheckCarrier(enShippingCarrier_PerPound)(0) Then Call GetPerPound(paryCheckCarrier(enShippingCarrier_PerPound))

'		If paryCheckCarrier(6)(0) Then Call GetAirborneRates(paryCheckCarrier(6))
'		If paryCheckCarrier(7)(0) Then Call GetDHLRates(paryCheckCarrier(7))

		'Freight Quote is kind of slow as of this writing
		'There is no reason to quote them unless there is a particularly heavy item/order
		If Not paryCheckCarrier(enShippingCarrier_FreightQuote)(0) Then pblnShowFreightQuoteRates = CBool(pobjDic.Count = 0)
		If Not pblnShowFreightQuoteRates Then paryCheckCarrier(enShippingCarrier_FreightQuote)(0) = CBool(pblnExceedsUPSMaxWeight AND pblnExceedsUSPSMaxWeight AND pblnExceedsFedExMaxWeight) AND paryCheckCarrier(enShippingCarrier_FreightQuote)(0)

		If paryCheckCarrier(enShippingCarrier_FreightQuote)(0) Then Call GetFreightQuoteRates(paryCheckCarrier(enShippingCarrier_FreightQuote), pstrShippingSelection)
		'debugprint "GetFedExRates",paryCheckCarrier(enShippingCarrier_FedEx)(0)
		'debugprint "GetUPSRates",paryCheckCarrier(enShippingCarrier_UPS)(0)
		'debugprint "GetFreightQuoteRates",paryCheckCarrier(enShippingCarrier_FreightQuote)(0)

		If False Then

			Dim vItem
			Dim paryKeys

			Set pobjTempDic = CreateObject("Scripting.Dictionary")
			If prsssShippingMethods.RecordCount > 0 Then prsssShippingMethods.MoveFirst

			paryKeys = pobjDic.Keys()
			For j = 0 To UBound(paryKeys)
				pstrKey = paryKeys(j)
				'Response.Write j & " = " & pstrKey & "<br />"
				prsssShippingMethods.Filter = "ssShippingCode='" & pobjDic.Item(pstrKey).ShippingCode & "'"
				If Not prsssShippingMethods.EOF Then
					pobjDic.Item(pstrKey).Postage = SumRate(pobjDic.Item(pstrKey), prsssShippingMethods)
					pobjTempDic.Add pstrKey, pobjDic.Item(pstrKey)
				End If
			Next 'j
			Set pobjDic = pobjTempDic
			Set pobjTempDic = Nothing
		Else
			Dim pblnValidCountryCheck

			Set pobjTempDic = CreateObject("Scripting.Dictionary")
			If prsssShippingMethods.RecordCount > 0 Then prsssShippingMethods.MoveFirst
			For i = 1 to prsssShippingMethods.RecordCount
				pstrKey = Trim(prsssShippingMethods.Fields("ssShippingCode").value)
				Select Case Trim(prsssShippingMethods.Fields("ssShippingMethodCountryRule").value & "")
					Case "1":	'U.S. Only
						pblnValidCountryCheck = CBool(pstrDestinationCountryAbb = "US")
					Case "2":	'Intl Only
						pblnValidCountryCheck = CBool(pstrDestinationCountryAbb <> "US")
					Case Else	'the rest
						pblnValidCountryCheck = True
				End Select
				'Response.Write i & " = " & pstrKey & "<br />"
				If pobjDic.Exists(pstrKey) And pblnValidCountryCheck Then
					pobjDic.Item(pstrKey).Postage = SumRate(pobjDic.Item(pstrKey), prsssShippingMethods)
					pobjTempDic.Add pstrKey, pobjDic.Item(pstrKey)
				End If
				prsssShippingMethods.movenext
			Next 'i
			Set pobjDic = pobjTempDic
			Set pobjTempDic = Nothing
		End If
				
		If pobjDic.Exists(pstrssShippingMethodCode) Then
			p_strGetRates = pobjDic.Item(pstrssShippingMethodCode).Postage
			pstrssShippingMethodName = pobjDic.Item(pstrssShippingMethodCode).ShippingType
		ElseIf pobjDic.Exists(pstrShippingSelection) Then
			p_strGetRates = pobjDic.Item(pstrShippingSelection).Postage
			pstrssShippingMethodName = pobjDic.Item(pstrShippingSelection).ShippingType
		Else
			p_strGetRates = "FAIL"
		End If

		If cblnDebugPostageRateAddon Then
			Response.Write "&nbsp;&nbsp;strssShippingMethodCode " & strssShippingMethodCode & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;pobjDic.Exists(pstrssShippingMethodCode) " & pobjDic.Exists(pstrssShippingMethodCode) & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;p_strGetRates " & p_strGetRates & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;pobjDic.Count " & pobjDic.Count & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;pblnExceedsUPSMaxWeight " & pblnExceedsUPSMaxWeight & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;pblnExceedsUSPSMaxWeight " & pblnExceedsUSPSMaxWeight & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;pblnExceedsFedExMaxWeight " & pblnExceedsFedExMaxWeight & "<br />" & vbcrlf
		End If

		If (p_strGetRates <> "FAIL") And (pobjDic.Count = 0) And (pblnExceedsUPSMaxWeight OR pblnExceedsUSPSMaxWeight OR pblnExceedsFedExMaxWeight) Then
		'If (pobjDic.Count = 0) And (pblnExceedsUPSMaxWeight OR pblnExceedsUSPSMaxWeight OR pblnExceedsFedExMaxWeight) Then
			Dim pclsShippingOption
			set pclsShippingOption = New clsShippingOption
			pclsShippingOption.CarrierCode = 2
			pclsShippingOption.ShippingCode = pstrProductBasedShippingCode
			pclsShippingOption.ShippingType = ShipCodeToMethod(pstrProductBasedShippingCode)
			pclsShippingOption.ShippingID = ShipCodeToID(pstrProductBasedShippingCode)
			pclsShippingOption.Postage = pdblProductBasedShipping
			pclsShippingOption.Insurance = 0
'			pclsShippingOption.shipRate = ShipMultiple(pstrProductBasedShippingCode)
			If Not pobjDic.Exists(pstrKey) Then pobjDic.Add pstrKey,pclsShippingOption
			p_strGetRates = pdblProductBasedShipping
			pstrssShippingMethodName = pclsShippingOption.ShippingType
		End If
		
		Dim plngCounter
		ReDim paryAvailableShippingMethods(pobjDic.count - 1)
		plngCounter = -1
		For each vItem in pobjDic
			plngCounter = plngCounter + 1
			With pobjDic.Item(vItem)
				paryAvailableShippingMethods(plngCounter) = Array(Trim(.ShippingCode), Trim(.ShippingType), .Postage)
			End With
		next

		GetRates = p_strGetRates

		'sort array
		If False Then
		Dim p_objDic
		Set p_objDic = CreateObject("SCRIPTING.DICTIONARY")
		prsssShippingMethods.Filter = ""
		prsssShippingMethods.MoveFirst
		For i = 1 to prsssShippingMethods.RecordCount
			pstrKey = Trim(prsssShippingMethods("ssShippingCode").value)

			Response.Write pstrKey & "<br />"
			If pobjDic.Exists(pstrKey) Then 
				p_objDic.Add pstrKey, pobjDic.Item(pstrKey)
			End If
			prsssShippingMethods.movenext
		Next 'i
		Set pobjDic = p_objDic
		Set p_objDic = Nothing
		End If
	
	End Function	'GetRates

	'***********************************************************************************************

	Private Function SumRate(ByRef objPostageItem, ByRef objrsShipMethod)

	Dim p_sngPackageCost
	
	Dim p_sngPostage
	Dim p_sngMultiple
	Dim p_sngPerPackageFee
	Dim p_sngMinCharge
	Dim p_sngPerShipmentFee
	Dim p_sngOfferFreeShippingAbove
	Dim p_sngLimitFreeShippingByWeight
	Dim p_lngssShippingMethodOrderBy

	Dim pblnFreeShipping
	
		With objPostageItem	

			p_sngPostage = Trim(.Postage & "")
			p_sngMultiple = Trim(prsssShippingMethods("ssShippingMethodMultiple").value & "")
			p_sngPerPackageFee = Trim(prsssShippingMethods("ssShippingMethodPerPackageFee").value & "")
			p_sngMinCharge = Trim(prsssShippingMethods("ssShippingMethodMinCharge").value & "")
			p_sngPerShipmentFee = Trim(prsssShippingMethods("ssShippingMethodPerShipmentFee").value & "")
			p_sngOfferFreeShippingAbove = Trim(prsssShippingMethods("ssShippingMethodOfferFreeShippingAbove").value & "")
			p_sngLimitFreeShippingByWeight = Trim(prsssShippingMethods("ssShippingMethodLimitFreeShippingByWeight").value & "")
			p_lngssShippingMethodOrderBy = Trim(prsssShippingMethods("ssShippingMethodOrderBy").value & "")
			
			.sortOrder = p_lngssShippingMethodOrderBy

			If Len(p_sngPostage) = 0 Then p_sngPostage = 0
			If Len(p_sngMultiple) = 0 Then p_sngMultiple = 1
			If Len(p_sngPerPackageFee) = 0 Then p_sngPerPackageFee = 0
			If Len(p_sngMinCharge) = 0 Then p_sngMinCharge = 0
			If Len(p_sngPerShipmentFee) = 0 Then p_sngPerShipmentFee = 0
			If Len(p_sngOfferFreeShippingAbove) = 0 Then p_sngOfferFreeShippingAbove = 0
			If Len(p_sngLimitFreeShippingByWeight) = 0 Then p_sngLimitFreeShippingByWeight = 0

			If cblnDebugPostageRateAddon Then	'pblnDebug
				Response.Write "<br /><hr><strong>Summing rate for <em>" & .ShippingType & "</em> (" & .ShippingCode & ")</strong><br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;Base Postage: " & p_sngPostage & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;Multiple: " & p_sngMultiple & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;PerPackageFee: " & p_sngPerPackageFee & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;MinCharge: " & p_sngMinCharge & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;PerShipmentFee: " & p_sngPerShipmentFee & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;OfferFreeShippingAbove: " & p_sngOfferFreeShippingAbove & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;FixedShippingAmount: " & pdblFixedShippingAmount & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;OrderSubtotal: " & psngOrderSubtotal & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;Insured: " & pblnInsured & "<br />" & vbcrlf
				Response.Flush
			End If

			p_sngPackageCost = p_sngPostage * p_sngMultiple + p_sngPerPackageFee + pdblFixedShippingAmount
			If pblnInsured Then p_sngPackageCost = p_sngPackageCost + .Insurance
			If CDbl(p_sngPackageCost) < CDbl(p_sngMinCharge) Then p_sngPackageCost = p_sngMinCharge
			
			'Check for free shipping
			pblnFreeShipping = CBool(CDbl(psngOrderSubtotal) >= CDbl(p_sngOfferFreeShippingAbove))
			If cblnDebugPostageRateAddon Then	'pblnDebug
				Response.Write "&nbsp;&nbsp;pblnFreeShipping (for order amount): " & pblnFreeShipping & "<br />" & vbcrlf
				Response.Flush
			End If
			
			'Additional weight check for free shipping
			If pblnFreeShipping Then
				pblnFreeShipping = CBool(p_sngLimitFreeShippingByWeight = 0) Or CBool(CDbl(pdblTotalOrderWeight) <= CDbl(p_sngLimitFreeShippingByWeight))
			End If
			If cblnDebugPostageRateAddon Then	'pblnDebug
				Response.Write "&nbsp;&nbsp;pblnFreeShipping (for weight): " & pblnFreeShipping & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;p_sngLimitFreeShippingByWeight: " & p_sngLimitFreeShippingByWeight & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;pdblTotalOrderWeight: " & pdblTotalOrderWeight & "<br />" & vbcrlf
				Response.Flush
			End If
			
			If pblnFreeShipping Then
				p_sngPostage = 0
			Else
				p_sngPostage = FormatNumber(p_sngPerShipmentFee + .NumPackages * p_sngPackageCost, 2)
			End If
			.Postage = p_sngPostage
			
			If cblnDebugPostageRateAddon Then 
				Response.Write "&nbsp;&nbsp;PackageCost: " & p_sngPackageCost & "<br />"
				Response.Write "&nbsp;&nbsp;Insured: " & pblnInsured & "<br />"
				Response.Write "&nbsp;&nbsp;Free Shipping: " & pblnFreeShipping & "<br />"
				Response.Write "&nbsp;&nbsp;Postage: " & p_sngPostage & "<br />"
			End If

			SumRate = .Postage
		
		End With

	End Function	'SumRate

	'***********************************************************************************************

	Public Sub GetAnyAvailableShipping(byRef vntShipCode, byRef strShipMethodName, byRef vntShipping)
	
	Dim vItem
	Dim pvntTempShipCode, pstrTempShipMethodName, pvntTempShipping
	
		'No reason to requery the carriers if no code was set as they will already have been queried
		If Len(pstrssShippingMethodCode) > 0 Then
			prsssShippingMethods.Filter = ""
			Call GetRates("")
		End If
		
		If pobjDic.Count = 0 Then
			vntShipping = "FAIL"
		Else
			pvntTempShipping = 9999999999
			For each vItem in pobjDic
				With pobjDic.Item(vItem)
					If CDbl(.Postage) < CDbl(pvntTempShipping) And (.ShippingCode <> cstrLocalPickupCode) Or isGenericUSPSCodeMatch(vntShipCode, vItem) Then

						'pvntTempShipCode = .ShippingID
						pvntTempShipCode = .ShippingCode
						pstrTempShipMethodName = .ShippingType
						pvntTempShipping = .Postage
						
						If False Then
							Response.Write "<fieldset><legend>GetAnyAvailableShipping</legend>"
							Response.Write ".ShippingID: " & .ShippingID & "<br />"
							Response.Write ".ShippingCode: " & .ShippingCode & "<br />"
							Response.Write ".ShippingType: " & .ShippingType & "<br />"
							Response.Write ".Postage: " & .Postage & "<br />"
							Response.Write "</fieldset>"
						End If
						pstrssShippingMethodCode = pvntTempShipCode
						pstrssShippingMethodName = pstrTempShipMethodName
					End If
				End With
			Next
			
			If pvntTempShipping = 9999999999 Then
				pvntTempShipping = 0
			End If
			
			vntShipCode = pvntTempShipCode
			strShipMethodName = pstrTempShipMethodName
			vntShipping = pvntTempShipping
		End If	'pobjDic.Count = 0
	
	End Sub	'GetAnyAvailableShipping

	'***********************************************************************************************

	Public Function Postage(vntssShippingMethodCode)
	
	Dim p_strssShippingMethodCode

'		On Error Resume Next

		p_strssShippingMethodCode = vntssShippingMethodCode
		'If isNumeric(vntssShippingMethodCode) Then
		'	p_strssShippingMethodCode = ShipIDtoCode(vntssShippingMethodCode)
		'Else
		'	p_strssShippingMethodCode = vntssShippingMethodCode
		'End If

		If pobjDic.Exists(p_strssShippingMethodCode) Then
			Postage = pobjDic.Item(p_strssShippingMethodCode).Postage
		Else
			Postage = "FAIL"
		End If
		
	End Function	'Postage

	'***********************************************************************************************

	Private Function ShipMultiple(strShipCode)
	
		If isNumeric(strShipCode) Then
			prsssShippingMethods.Filter = "ShipID=" & strShipCode
			If not prsssShippingMethods.EOF Then 
				ShipMultiple = Trim(prsssShippingMethods("ssShippingMethodMultiple").value)
			Else
				ShipMultiple = 99
			End If
		ElseIf Len(strShipCode) > 0 Then
			prsssShippingMethods.Filter = "ssShippingCode='" & strShipCode & "'"
'			prsssShippingMethods.Filter = "ssShippingCode='" & strShipCode & Space(prsssShippingMethods.Fields("ssShippingCode").DefinedSize - len(strShipCode))& "'"
			If not prsssShippingMethods.EOF Then 
				ShipMultiple = Trim(prsssShippingMethods("ssShippingMethodMultiple").value)
			Else
				ShipMultiple = 99
			End If
		End If
		prsssShippingMethods.Filter = ""
		
	End Function	'ShipMultiple

	'***********************************************************************************************

	Public Function ShipIDtoCode(strssShippingMethodCode)
	
'		On Error Resume Next
		
		If isNumeric(strssShippingMethodCode) Then
			If not isObject(prsssShippingMethods) Then Call InitializeShipMethods
			prsssShippingMethods.Filter = "ssShippingMethodID=" & strssShippingMethodCode
			If not prsssShippingMethods.EOF Then 
				ShipIDtoCode = Trim(prsssShippingMethods("ssShippingMethodCode").value)
			Else
				ShipIDtoCode = ""		'Should never see this condition
			End If
		Else
			ShipIDtoCode = strssShippingMethodCode
		End If
		
	End Function	'ShipIDtoCode

	Public Function ShipCodetoID(strShipCode)
	
'		On Error Resume Next

		If not isObject(prsssShippingMethods) Then Call InitializeShipMethods
		prsssShippingMethods.Filter = "ssShippingCode='" & strShipCode & "'"
		If not prsssShippingMethods.EOF Then 
			ShipCodetoID = prsssShippingMethods("ssShippingCarrierID")
		Else
			ShipCodetoID = ""		'Should never see this condition
		End If
		prsssShippingMethods.Filter = ""
		
	End Function	'ShipCodetoID

	Public Function ShipCodetoMethod(strShipCode)
	
'		On Error Resume Next

		If not isObject(prsssShippingMethods) Then Call InitializeShipMethods(strShipCode)
		prsssShippingMethods.Filter = "ssShippingCode='" & strShipCode & "'"
		If not prsssShippingMethods.EOF Then 
			ShipCodetoMethod = Trim(prsssShippingMethods.Fields("ssShippingMethodName").Value & "")
		Else
			ShipCodetoMethod = ""		'Should never see this condition
		End If
		prsssShippingMethods.Filter = ""
		
	End Function	'ShipCodetoMethod

	'***********************************************************************************************

	Public Function Insurance(strssShippingMethodCode)
	
		On Error Resume Next
		Insurance = pobjDic.Item(strssShippingMethodCode).Insurance

	End Function

	'***********************************************************************************************

	Public Property Get dicRates()
		If isObject(pobjDic) Then Set dicRates = pobjDic
	End Property

	'***********************************************************************************************

	Public Sub DisplayShippingOptions(blnRequired)
	
	Dim vItem
	
	Response.Write "<center>"
	If pobjDic.Count > 0 Then Response.Write "<h3>Select Shipping Method</h3>"
	Response.Write "<table cellpadding='5' width='90%' border=1 cellspacing=0>"
	Response.Write "<colgroup><col align='left' /><col align='right' /><col align='right' /></colgroup>"

	If pobjDic.Count > 0 Then
		Response.Write "<tr>"
		Response.Write "<th>Shipping Method</th>"
		Response.Write "<th>Postage</th>"
		'Response.Write "<th>Insurance</th>"
		Response.Write "</tr>"
	End If

	for each vItem in pobjDic
		With pobjDic.Item(vItem)
			If blnRequired Then
				Response.Write "<tr><td>"
				Response.Write "<a href='' onclick='return MakeSelection_required(""" & .ShippingCode & """,""" & .ShippingType & """);' title='Select this shipping option'>" & .ShippingType & "</a>"
				If pblnShowTransitTimes Then Response.Write "<br />Transit Time " & .TransitTime & ""
				Response.Write "</td>"
			Else
				Response.Write "<tr><td>"
				Response.Write "<a href='' onclick='return MakeSelection(" & Chr(34) & .ShippingCode & Chr(34) & "," _
																		   & Chr(34) & .ShippingType & Chr(34) & "," _
																		   & Chr(34) & vItem & Chr(34) & ");' title='Select this shipping option'>" & .ShippingType & "</a>"
				If Len(.TransitTime) > 0 And pblnShowTransitTimes Then
					If .TransitTime = 1 Then
						Response.Write "<br />Transit Time: 1 day"
					Else
						Response.Write "<br />Transit Time: " & .TransitTime & " days"
					End If
				End If
				Response.Write "</td>"
				'Response.Write "<tr><td><a href='' onclick='return MakeSelection(""" & .ShippingCode & """,""" & .ShippingType & """);' title='Select this shipping option'>" & .ShippingType & "</a></td>"
			End If
			Response.Write "<td><a href='' onclick='return MakeSelection(" & Chr(34) & .ShippingCode & Chr(34) & "," _
																		& Chr(34) & .ShippingType & Chr(34) & "," _
																		& Chr(34) & vItem & Chr(34) & ");' title='Select this shipping option'>" & FormatCurrency(.Postage,2) & "</a></td>"
			'Response.Write "<td>" & FormatCurrency(.Postage,2) & "</td>"
'			Response.Write "<td>" & FormatCurrency(.Insurance,2) & "</td>"
			Response.Write "</tr>" & vbcrlf
		End With
	next
	If pobjDic.Count = 0 Then 
		Response.Write "<tr><td colspan=2>" & vbcrlf
		Response.Write "There are no available shipping methods to the specified shipping address.<br />"
		Response.Write "Please verify your shipping address and try again.<br />"
	'	Response.Write "If you are having difficulty, please contact our order desk at: <br />"
		Response.Write "</td></tr>" & vbcrlf
		Response.Write "<tr><td colspan=2><a href='' onclick='window.close();'>Close window</a></td></tr>"		
	End If
	Response.Write "</table>"
	Response.Write "</center>"
	End Sub

	'***********************************************************************************************

	Public Sub altDisplayShippingOptions
	
	Dim vItem

	Response.Write "<center>"
	Response.Write "<table cellpadding='5' width='90%'>"
	Response.Write "<colgroup align='left'><colgroup align='right'><colgroup align='right'>"

	If pobjDic.Count > 0 Then
		Response.Write "<tr colspan=2><td align=center>Select Shipping Method</td></tr>"
		Response.Write "<tr>"
		Response.Write "<th>Shipping Method</th>"
		Response.Write "<th>Postage</th>"
	'	Response.Write "<th>Insurance</th>"
		Response.Write "</tr>"
	End If
	
	for each vItem in pobjDic
		With pobjDic.Item(vItem)
			Response.Write "<tr><td><a href='' onclick='return MakeSelection(" & chr(34) & .ShippingID & chr(34) & "," & chr(34) & .ShippingType & chr(34) & ");' title='Select this shipping option'>" & .ShippingType & "</a></td>"
			Response.Write "<td>" & FormatCurrency(.Postage,2) & "</td>"
'			Response.Write "<td>" & FormatCurrency(.Insurance,2) & "</td>"
			Response.Write "</tr>" & vbcrlf
		End With
	next
	If pobjDic.Count = 0 Then 
		Response.Write "<tr><td colspan=2>" & vbcrlf
		Response.Write "There are no available shipping methods to the specified shipping address. Please correct the shipping address and click the " _
					 & "button to Select Shipping Method again, or if you are having difficulty, please contact our order desk at: <br />"
		Response.Write "</td></tr>" & vbcrlf
		Response.Write "<tr><td colspan=2><a href='' onclick='window.close();'>Close window</a></td></tr>"		
	End If
	Response.Write "</table>"
	Response.Write "</center>"
	End Sub

	'***********************************************************************************************

	Public Sub DisplayShippingOptionsAsRadio
	
	Dim vItem
	Dim pCount
	
		Response.Write "<center>"
		Response.Write "<form name='getRates' onsubmit='return MakeSelectionRadio(this);'>"
		Response.Write "<table cellpadding='0' cellspacing='0' width='90%'>"
		Response.Write "<colgroup align='left'><colgroup align='right'><colgroup align='right'>"

		If pobjDic.Count > 0 Then
			Response.Write "<tr>"
		'	Response.write("<FONT COLOR=yellow SIZE=2 FACE='Verdana, Arial, Helvetica, sans-serif'>")
			Response.Write "<th bgcolor='#FFFFCC'>Service</th>"
		'	Response.Write "<th bgcolor='#FFFFCC'>Business Days</th>"
			Response.Write "<th bgcolor='#FFFFCC'>Rate</th>"
			Response.Write "</tr>"
			Response.Write "<tr><td colspan=2><hr></td></tr>"		
		End If

		pCount = 0
		for each vItem in pobjDic
			With pobjDic.Item(vItem)
				Response.Write "<tr>"
				If cblnGenericUSPSMethods Then
                    Response.Write "<td><input type='radio' value='" & setGenericUSPSCodeMatch(Trim(.ShippingCode)) & "' name='radio1' id='radio" & pCount & "'><label for='radio" & pCount & "'>" & Trim(.ShippingType) & "</label></td>" & vbcrlf
				Else
                    Response.Write "<td><input type='radio' value='" & Trim(.ShippingCode) & "' name='radio1' id='radio" & pCount & "'><label for='radio" & pCount & "'>" & Trim(.ShippingType) & "</label></td>" & vbcrlf
				End If
				If cDbl(.Postage) = 0 Then
					Response.Write "<td><b>Free</b></td>"
				Else
					Response.Write "<td>" & FormatCurrency(.Postage,2) & "</td>"
				End If
				Response.Write "</tr>" & vbcrlf
			End With
			If pCount < pobjDic.Count Then
				Response.Write "<tr><td colspan=2><hr></td></tr>"		
			End If
			pCount = pCount + 1
		next
		If pobjDic.Count = 0 Then 
			Response.Write "<tr><td colspan=2>" & vbcrlf
			Response.Write "<center>"
			Response.Write "There are no available shipping methods to the specified shipping address.<br /><br />"
			Response.Write "Please verify your Postal Code and try again.<br /><br />"
		'	Response.Write "If you are having difficulty, please contact customer service at <br />"
			Response.Write "</center>"
			Response.Write "<tr><td colspan=2><hr></td></tr>"		
			Response.Write "</td></tr>" & vbcrlf
			Response.Write "<tr><td colspan=2><a href='' onclick='window.close();'>Close window</a></td></tr>"		
		Else
			Response.Write "<tr><td colspan=2  bgcolor='#FFFFCC'>" & vbcrlf
			Response.Write "<center>Please select your shipping method and click <a href='' onclick='MakeSelectionRadio(document.getRates); return false;'>CONTINUE</a>.</center>"
			Response.Write "<tr><td colspan=2><hr></td></tr>"		
'			Response.Write "<tr><td colspan=2 align='right'><input type=image name='btnGetShipping' id='btnGetShipping' src='/images/continue.gif' value='Select Shipping Method'></td></tr>"
		End If
	
		Response.Write "</table>"
		Response.Write "</form>"
		Response.Write "</center>"
	
	End Sub	'DisplayShippingOptionsAsRadio

	'***********************************************************************************************

	Public Sub DisplayShippingOptionsAsSelect
	
	Dim vItem
	
	'sort dictionary
	Dim i, j
	
	Response.Write "<select name='Shipping' size='" & pobjDic.Count & "'>" & vbcrlf
	for each vItem in pobjDic
		With pobjDic.Item(vItem)
			Response.Write "<option value='" & .ShippingID & "'>" & .ShippingType & ": " & FormatCurrency(.Postage,2) & "</option>" & vbcrlf
		End With
	next
	Response.Write "</select>" & vbcrlf
'	If Len(vItem) = 0 then Response.Write "<tr><td colspan=2>There are no available shipping methods</td></tr>"
	
	End Sub

	'***********************************************************************************************

	Private Sub InitializeShipMethods(byVal strssShippingMethodCode)
	
	Dim i
	Dim pstrSQL
	Dim pstrSQL_Enabled
	Dim pstrSQL_Special
	Dim pstrSQLWhere
	Dim pdblTempWeight
	Dim arySpecialShipCodes

		If pblnUseOrderWeightToLimitCarrierChoices Then
			pdblTempWeight = pdblTotalOrderWeight
		Else
			pdblTempWeight = pdblMaxItemWeight
		End If	
	
		pstrSQL_Enabled = "(ssShippingMethodEnabled<>0)"
		
		If Len(pstrEnabledSpecialShippingMethods) > 0 Then
			arySpecialShipCodes = Split(pstrEnabledSpecialShippingMethods, ",")
			For i = 0 To UBound(arySpecialShipCodes)
				If Len(pstrSQL_Special) = 0 Then
					pstrSQL_Special = "ssShippingMethodID=" & Trim(arySpecialShipCodes(i))
				Else
					pstrSQL_Special = pstrSQL_Special & " Or ssShippingMethodID=" & Trim(arySpecialShipCodes(i))
				End If
			Next 'i
			pstrSQL_Special = " And (ssShippingMethodIsSpecial=0 Or " & pstrSQL_Special & ")"
		Else
			pstrSQL_Special = " And (ssShippingMethodIsSpecial=0 Or ssShippingMethodIsSpecial is Null)"
		End If
		pstrSQL_Enabled = pstrSQL_Enabled & pstrSQL_Special

		set prsssShippingMethods = CreateObject("ADODB.RECORDSET")
		with prsssShippingMethods
			.ActiveConnection = pobjConn
			.CursorLocation = 2 'adUseClient
			.CursorType = 3 'adOpenStatic
			.LockType = 1 'adLockReadOnly

			If Len(pdblTempWeight) > 0 Then
				If Len(strssShippingMethodCode) > 0 Then
					pstrSQLWhere = " WHERE " & pstrSQL_Enabled & " AND (ssShippingMethodCode = '" & Replace(strssShippingMethodCode, "'", "''") & "') AND (ssShippingMethodMinWeight<=" & pdblTempWeight & ") AND ((ssShippingMethodMaxWeight>=" & pdblTempWeight & ") OR (ssShippingMethodMaxWeight=0))"
				Else
					pstrSQLWhere = " WHERE " & pstrSQL_Enabled & " AND (ssShippingMethodMinWeight<=" & pdblTempWeight & ") AND ((ssShippingMethodMaxWeight>=" & pdblTempWeight & ") OR (ssShippingMethodMaxWeight=0))"
				End If
			Else
				If Len(strssShippingMethodCode) > 0 Then
					pstrSQLWhere = " WHERE " & pstrSQL_Enabled & " AND (ssShippingMethodCode = '" & Replace(strssShippingMethodCode, "'", "''") & "')"
				Else
					pstrSQLWhere = " WHERE " & pstrSQL_Enabled & ""
				End If
			End If

			If False Then	'True for Access, False for SQL Server
				pstrSQL = "SELECT ssShippingMethods.ssShippingCarrierID, ssShippingCarriers.ssShippingCarrierName, Trim(ssShippingMethodCode) as ssShippingCode, ssShippingMethodOrderBy, ssShippingMethodName, ssShippingMethodMinCharge, ssShippingMethodMultiple, ssShippingMethodPerPackageFee, ssShippingMethodPerShipmentFee, ssShippingMethodOfferFreeShippingAbove, ssShippingMethodLimitFreeShippingByWeight, ssShippingCarriers.ssShippingCarrierUserName, ssShippingCarriers.ssShippingCarrierPassword, ssShippingMethods.ssShippingMethodMaxWeight, ssShippingMethods.ssShippingMethodPrefWeight, ssShippingMethods.ssShippingMethodMaxLength, ssShippingMethods.ssShippingMethodMaxWidth, ssShippingMethods.ssShippingMethodMaxHeight, ssShippingMethods.ssShippingMethodMaxGirth, ssShippingMethodMinWeight, ssShippingMethodDefault, ssShippingMethodCountryRule" _
						& " FROM ssShippingMethods INNER JOIN ssShippingCarriers ON ssShippingMethods.ssShippingCarrierID = ssShippingCarriers.ssShippingCarrierID" _
						& pstrSQLWhere _
						& " ORDER BY ssShippingMethodOrderBy, ssShippingMethodName"
			Else
				pstrSQL = "SELECT ssShippingMethods.ssShippingCarrierID, ssShippingCarriers.ssShippingCarrierName, RTrim(ssShippingMethodCode) as ssShippingCode, ssShippingMethodOrderBy, ssShippingMethodName, ssShippingMethodMinCharge, ssShippingMethodMultiple, ssShippingMethodPerPackageFee, ssShippingMethodPerShipmentFee, ssShippingMethodOfferFreeShippingAbove, ssShippingMethodLimitFreeShippingByWeight, ssShippingCarriers.ssShippingCarrierUserName, ssShippingCarriers.ssShippingCarrierPassword, ssShippingMethods.ssShippingMethodMaxWeight, ssShippingMethods.ssShippingMethodPrefWeight, ssShippingMethods.ssShippingMethodMaxLength, ssShippingMethods.ssShippingMethodMaxWidth, ssShippingMethods.ssShippingMethodMaxHeight, ssShippingMethods.ssShippingMethodMaxGirth, ssShippingMethodMinWeight, ssShippingMethodDefault, ssShippingMethodCountryRule" _
						& " FROM ssShippingMethods INNER JOIN ssShippingCarriers ON ssShippingMethods.ssShippingCarrierID = ssShippingCarriers.ssShippingCarrierID" _
						& pstrSQLWhere _
						& " ORDER BY ssShippingMethodOrderBy, ssShippingMethodName"
			End If
			If pblnDebug Then debugprint "pstrSQL",pstrSQL
			.Source = pstrSQL
			
			If Err.number <> 0 Then Err.Clear
			On Error Resume Next
			.Open
			If Err.number <> 0 Then
				Response.Write "<h3><font color=red>The Postage Rate add-on upgrade does not appear to have been performed</font></h3>"
				Response.Write "<h4>Error " & Err.number & ": " & Err.Description & "</h4>"
				Response.Write "<h4>SQL: " & pstrSQL & "</h4>"
				Response.Flush
				Err.Clear
			Else

				ReDim maryShippingMethods(.RecordCount - 1)
				For i = 1 To .RecordCount
					Select Case Trim(.Fields("ssShippingCarrierName").Value & "")
						Case "U.S.P.S.":		enShippingCarrier_USPS = .Fields("ssShippingCarrierID").Value
						Case "UPS":				enShippingCarrier_UPS = .Fields("ssShippingCarrierID").Value
						Case "FedEx":			enShippingCarrier_FedEx = .Fields("ssShippingCarrierID").Value
						Case "Canada Post":		enShippingCarrier_CanadaPost = .Fields("ssShippingCarrierID").Value
						Case "Freight Quote":	enShippingCarrier_FreightQuote = .Fields("ssShippingCarrierID").Value
						Case "Unknown":			enShippingCarrier_Unknown = .Fields("ssShippingCarrierID").Value
						Case "FlatRate":		enShippingCarrier_FlatRate = .Fields("ssShippingCarrierID").Value
						Case "PerItem":			enShippingCarrier_PerItem = .Fields("ssShippingCarrierID").Value
						Case "PerPound":		enShippingCarrier_PerPound = .Fields("ssShippingCarrierID").Value
					End Select

					maryShippingMethods(i-1) = Array(Trim(.Fields("ssShippingCode").Value & ""), Trim(.Fields("ssShippingMethodName").Value & ""), Trim(.Fields("ssShippingMethodDefault").Value & ""), Trim(.Fields("ssShippingMethodOfferFreeShippingAbove").Value & ""))
					.MoveNext
				Next
				.MoveFirst
			End If
			
			If cblnDebugPostageRateAddon Then
				Response.Write "<fieldset><legend>Active Shipping Methods</legend>"
				Response.Write "SQL: " & pstrSQL & "<hr />"
				If isArray(maryShippingMethods) Then
					For i = 0 To UBound(maryShippingMethods)
						Response.Write maryShippingMethods(i)(0) & ": " & maryShippingMethods(i)(1) & "<br />"
					Next 'i
				Else
					Response.Write "No Active Shiping Methods"
				End If
				Response.Write "</fieldset>"
			End If
			
		End With
		
	End Sub	'InitializeShipMethods

	'***********************************************************************************************

	Private Sub class_Terminate()
		On Error Resume Next
		set pobjDic = nothing
		set prsssShippingMethods = nothing
	End Sub

	'***********************************************************************************************

	Private Function RetrieveRemoteData(strURL,strFormData,blnPostData)
	
	Dim pobjXMLHTTP
	
	'set timeouts in milliseconds
	Const resolveTimeout = 10000
	Const connectTimeout = 10000
	Const sendTimeout = 10000
	Const receiveTimeout = 100000
	Dim plngCounter
	Dim pstrResult
	
	On Error Resume Next

		If Err.number <> 0 Then	Err.Clear
		'Use MSXML2 if possible - must have the Microsoft XML Parser v3 or later installed
		Set pobjXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		If Err.number <> 0 Then
			Err.Clear
			Set pobjXMLHTTP = CreateObject("Microsoft.XMLHTTP")
		End If
		For plngCounter = 0 To 1	'added because of unexplained error on the first call
			With pobjXMLHTTP
				If blnPostData Then
					.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
					.open "POST", strURL, False
					.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
					.send strFormData
				Else
					.open "GET", strURL, False
					.send
				End If
				pstrResult  = .responseText
				RetrieveRemoteData  = pstrResult
				'Response.Write "() Error:" & Err.number & " - " & Err.Description & " (" & Err.Source & ")" & "<br />"	
				If pblnDebug Then Response.Write "responseText =" & .responseText & "<br />" & vbcrlf
			End With
			If Err.number <> -2147467259 Then Exit For
		Next 'plngCounter
		set pobjXMLHTTP = nothing

		If Err.number <> 0 Then
			If cblnDebugPostageRateAddon Then
				Response.Write "<fieldset><legend>RetrieveRemoteData</legend>"
				Response.Write "<h4><font color=red>Error " & Err.number & ": " & Err.Description & "</font></h4>"
				Response.Write "strURL: " & strURL & "<br />" & vbcrlf
				Response.Write "responseText: " & pstrResult & vbcrlf
				Response.Write "</fieldset>" & vbcrlf
			End If
			If pblnDebug Then 
			Select Case Err.number
				Case -2147467259:	'Unspecified error 
					'This has only been seen when the server permissions are set incorrectly (access is denied msxml3.dll)
					Response.Write "<h3><font color=red>Server permissions error: <br /><i>msxml3.dll error '80070005'<br />Access is denied.</i><br />Please contact your server administrator</font></h3>"
					Response.Write "<h3><font color=red>Error " & Err.number & ": " & Err.Description & "</font></h3>"
				Case 438: 'Object doesn't support this property or method
					'This is from the set timeouts, no action required
				Case Else
					Response.Write "<h3><font color=red>Error " & Err.number & ": " & Err.Description & "</font></h3>"
			End Select
			End If
			Err.Clear
		End If

	End Function	'RetrieveRemoteData

	'****************************************************************************************************************

	Private Function CalculatePackageSize(ByVal sngMaxLength, _
										  ByVal sngMaxWidth, _
										  ByVal sngMaxHeight, _
										  ByVal sngMaxWeight, _
										  ByVal sngRecommendedWeight, _
										  ByVal sngMaxGirth, _
										  ByRef sngLength, _
										  ByRef sngWidth, _
										  ByRef sngHeight, _
										  ByRef sngWeight, _
										  ByRef lngNumPackages _
										 )
	
	Dim psng_Length
	Dim psng_Width
	Dim psng_Height
	Dim psng_Weight
	Dim plng_NumPackages
	Dim pdblMaxItemWeight
	Dim i,j
	
	Dim p_blnExceedsMaxWeight: p_blnExceedsMaxWeight = False
	Dim p_blnExceedsPrefWeight: p_blnExceedsPrefWeight = False
	Dim p_lngNumItems: p_lngNumItems = 0
	
	'Merge the rest into a superbox and divide by volume
	'Compare at the end and pick the largest to rate
	'returns length, width, height, weight, number of packages
	Dim paryShipAloneItem	'length, width, height, weight, qty
	Dim parySuperItem		'length, width, height, weight, volume
	Dim paryMaxItem			'length, width, height, weight, qty
	
		paryShipAloneItem = Array (0, 0, 0, 0, 0)
		parySuperItem = Array (0, 0, 0, 0, 1)
		paryMaxItem = Array (0, 0, 0, 0, 1)
		pdblMaxItemWeight = 0
		pdblFixedShippingAmount = 0
		pstrEnabledSpecialShippingMethods = paryOrderItems(1)(10)

		If cblnDebugPostageRateAddon Then Response.Write "<hr>Calculating Package Size . . .<br />" & vbcrlf

		'Look for items which are not combineable - this sets the base package
		'paryOrderItems(numItems)(8) decoder
		'0 - Length - default None
		'1 - Width - default None
		'2 - Height - default None
		'3 - Weight - default None
		'4 - Quantity - default None
		'5 - FixedSize - default True
		'6 - DoNotCombine - default False
		'7 - ProductBasedShipping - default None
		'8 - MustShipFreight - default False
		'9 - FixedShipping - default 0
		'10 - SpecialShipping - default None

		'psngLength, psngWidth, psngHeight, psngWeight, p_intQuantity, pblnFixedSize, pblnDoNotCombine, p_dblProductBasedShipping, pblnMustShipFreight, p_dblFixedShipping, p_strSpecialShipping

		For i = 1 to UBound(paryOrderItems)
		
			If cblnDebugPostageRateAddon Then
				Response.Write "<fieldset><legend>Item " & i & "</legend><br />" & vbcrlf
				Response.Write "Length: " & paryOrderItems(i)(0) & "<br />" & vbcrlf
				Response.Write "Width: " & paryOrderItems(i)(1) & "<br />" & vbcrlf
				Response.Write "Height: " & paryOrderItems(i)(2) & "<br />" & vbcrlf
				Response.Write "Weight: " & paryOrderItems(i)(3) & "<br />" & vbcrlf
				Response.Write "Qty: " & paryOrderItems(i)(4) & "<br />" & vbcrlf
				Response.Write "Fixed Shipping: " & paryOrderItems(i)(9) & "<br />" & vbcrlf
				Response.Write "Special Shipping: " & paryOrderItems(i)(10) & "<br />" & vbcrlf
				Response.Write "</fieldset>"
				Response.Flush
			End If	'cblnDebugPostageRateAddon
			
			paryOrderItems(i)(9) = Trim(paryOrderItems(i)(9) & "")
			If Len(paryOrderItems(i)(9)) = 0 Or Not isNumeric(paryOrderItems(i)(9)) Then paryOrderItems(i)(9) = 0
			If paryOrderItems(i)(9) = 0 Then
				p_lngNumItems = p_lngNumItems + paryOrderItems(i)(4)
				If paryOrderItems(i)(6) Then
					If paryShipAloneItem(3) < paryOrderItems(i)(3) Then
						For j = 0 To 3
							paryShipAloneItem(j) = paryOrderItems(i)(j)
						Next 'j
					End If
					paryShipAloneItem(4) = paryShipAloneItem(4) + paryOrderItems(i)(4)
				Else
					'since length, width, and height are ordered, package will be built by max length & width with heights added
					If parySuperItem(0) < paryOrderItems(i)(0) Then parySuperItem(0) = paryOrderItems(i)(0)
					If parySuperItem(1) < paryOrderItems(i)(1) Then parySuperItem(1) = paryOrderItems(i)(1)
					parySuperItem(2) = parySuperItem(2) + paryOrderItems(i)(2) * paryOrderItems(i)(4)
					
					parySuperItem(3) = parySuperItem(3) + paryOrderItems(i)(3) * paryOrderItems(i)(4)
					parySuperItem(4) = parySuperItem(4) + paryOrderItems(i)(0) * paryOrderItems(i)(1) * paryOrderItems(i)(2) * paryOrderItems(i)(4)
					'Now figure out the biggest item
					For j = 0 To 3
						If paryMaxItem(j) < paryOrderItems(i)(j) Then paryMaxItem(j) = paryOrderItems(i)(j)
					Next 'j
				End If
				
				If pdblMaxItemWeight < paryOrderItems(i)(3) Then pdblMaxItemWeight = paryOrderItems(i)(3)
			Else
				'If FixedShipping < 0 it indicates free shipping
				If paryOrderItems(i)(9) > 0 Then pdblFixedShippingAmount = pdblFixedShippingAmount + paryOrderItems(i)(9)
			End If	'paryOrderItems(i)(9) = 0
			
			'Check for special shipping; it must be an exact match
			If pstrEnabledSpecialShippingMethods <> paryOrderItems(i)(10) Then pstrEnabledSpecialShippingMethods = ""

		Next 'i
		
		If cblnDebugPostageRateAddon Then
			Response.Write "<fieldset><legend>Superpackage (pre-minimum)</legend><br />" & vbcrlf
			Response.Write "Length: " & parySuperItem(0) & "<br />" & vbcrlf
			Response.Write "Width: " & parySuperItem(1) & "<br />" & vbcrlf
			Response.Write "Height: " & parySuperItem(2) & "<br />" & vbcrlf
			Response.Write "Weight: " & parySuperItem(3) & "<br />" & vbcrlf
			Response.Write "Volume: " & parySuperItem(4) & "<br />" & vbcrlf
			Response.Write "</fieldset>"
			Response.Flush
		End If	'pblnDebug

		'Set minimums in case there are none
		For j = 0 To 2
			If parySuperItem(j) = 0 Then parySuperItem(j) = 1
			If paryShipAloneItem(j) = 0 Then paryShipAloneItem(j) = 1
			If paryMaxItem(j) = 0 Then paryMaxItem(j) = 1
		Next 'j
		
		If cblnDebugPostageRateAddon Then
			Response.Write "<fieldset><legend>Superpackage (post-minimum)</legend><br />" & vbcrlf
			Response.Write "Length: " & parySuperItem(0) & "<br />" & vbcrlf
			Response.Write "Width: " & parySuperItem(1) & "<br />" & vbcrlf
			Response.Write "Height: " & parySuperItem(2) & "<br />" & vbcrlf
			Response.Write "Weight: " & parySuperItem(3) & "<br />" & vbcrlf
			Response.Write "Volume: " & parySuperItem(4) & "<br />" & vbcrlf
			Response.Write "</fieldset>"
			Response.Flush
		End If	'pblnDebug

		'Now break the SuperItem into a shippable package
		Dim psngTempWeight
		Dim plngNumPackages
		
		p_blnExceedsMaxWeight = CBool(CDbl(parySuperItem(3)) > CDbl(sngMaxWeight))
		p_blnExceedsPrefWeight = CBool(CDbl(parySuperItem(3)) > CDbl(sngRecommendedWeight))

		Dim pblnAllowPackageBreakdown: pblnAllowPackageBreakdown = cblnAllowPackageBreakdown
		If pblnAllowPackageBreakdown Or (p_lngNumItems > 1) Then pblnAllowPackageBreakdown = True
		
		If cblnDebugPostageRateAddon Then
			Response.Write "<br /><strong>Check for exceeding preferred weight</strong><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Weight: " & CDbl(parySuperItem(3)) & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Recommended Wt: " & sngRecommendedWeight & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Exceeds recommended weight: " & p_blnExceedsPrefWeight & "<br />" & vbcrlf
			
			Response.Write "<br /><b>Check for exceeding max weight</b><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Max Item Weight: " & pdblMaxItemWeight & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Max Weight: " & sngMaxWeight & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Exceeds max weight: " & p_blnExceedsMaxWeight & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;pblnAllowPackageBreakdown: " & pblnAllowPackageBreakdown & "<br />" & vbcrlf
			
			Response.Write "<br /><b>Special Shipping Characteristics</b><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Special Shipping Methods: " & pstrEnabledSpecialShippingMethods & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Fixed Shipping: " & pdblFixedShippingAmount & "<br />" & vbcrlf
			Response.Flush
		End If	'pblnDebug
		
		If p_blnExceedsMaxWeight And pblnAllowPackageBreakdown Then
			If p_blnExceedsPrefWeight Then 
				plngNumPackages = int((parySuperItem(3)/sngRecommendedWeight) + 0.99)
				psngTempWeight = parySuperItem(3)/plngNumPackages
			Else
				plngNumPackages = int((parySuperItem(3)/sngMaxWeight) + 0.99)
				psngTempWeight = parySuperItem(3)/plngNumPackages
			End If

			If cblnDebugPostageRateAddon Then
				Response.Write "<br /><b>Redid package weight</b><br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;plngNumPackages: " & plngNumPackages & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;psngTempWeight: " & psngTempWeight & "<br />" & vbcrlf
			End If

			p_blnExceedsMaxWeight = False
			'cut down the package height by the number of packages - this will get cleaned up in the final step
			parySuperItem(2) = parySuperItem(2)/plngNumPackages
		ElseIf p_blnExceedsPrefWeight And pblnAllowPackageBreakdown Then
			If pdblMaxItemWeight > sngRecommendedWeight Then 
				plngNumPackages = int((parySuperItem(3)/pdblMaxItemWeight) + 0.99)
				psngTempWeight = parySuperItem(3)/plngNumPackages
			Else
				plngNumPackages = int((parySuperItem(3)/sngRecommendedWeight) + 0.99)
				psngTempWeight = parySuperItem(3)/plngNumPackages
			End If
			
			'cut down the package height by the number of packages - this will get cleaned up in the final step
			parySuperItem(2) = parySuperItem(2)/plngNumPackages
		Else
			plngNumPackages = 1
			psngTempWeight = parySuperItem(3)
		End If
		
		'Now check for packages over girth
		Dim psng_Girth
		Dim p_sngVolume
		
		psng_Length = parySuperItem(0)
		psng_Width = parySuperItem(1)
		psng_Height = parySuperItem(2)
		psng_Girth = psng_Length + 2 * (psng_Height + psng_Width)
		
		If Not isNumeric(sngMaxGirth) Then
			sngMaxGirth = 0
		Else
			sngMaxGirth = CDbl(sngMaxGirth)
		End If
		
		If cblnDebugPostageRateAddon Then
			Response.Write "<br /><b>Girth Check</b><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;psng_Girth: " & psng_Girth & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;sngMaxGirth: " & sngMaxGirth & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;CBool((psng_Girth > sngMaxGirth)): " & CBool((psng_Girth > sngMaxGirth)) & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;CBool((sngMaxGirth <> 0)): " & CBool((sngMaxGirth <> 0)) & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Exceeds Max Girth Check: " & CBool((psng_Girth > sngMaxGirth) And (sngMaxGirth <> 0)) & "<br />" & vbcrlf
		End If
		
		If (psng_Girth > sngMaxGirth) And (sngMaxGirth <> 0)	Then
			plng_NumPackages = int((psng_Girth/sngMaxGirth) + 0.99)
			
			'now repackage into cubes
			p_sngVolume = psng_Length * psng_Height * psng_Width
			psng_Length = paryMaxItem(0)
			psng_Height = Round(sqr(p_sngVolume / psng_Length) + .5)
			psng_Width = psng_Height
			
			If psng_Width > paryMaxItem(1) Then psng_Width = paryMaxItem(1)
			If psng_Height > paryMaxItem(2) Then psng_Height = paryMaxItem(2)

			psng_Length = paryMaxItem(0)
			psng_Width = paryMaxItem(1)
			psng_Height = paryMaxItem(2)

			psngTempWeight = psngTempWeight / plng_NumPackages
			plngNumPackages = plngNumPackages * plng_NumPackages

		End If
		
		'Check to force package weight to be at least the max item weight
		If Not cblnAllowPackageBreakdown And CBool(CDbl(psngTempWeight) < CDbl(pdblMaxItemWeight)) Then psngTempWeight = pdblMaxItemWeight
		
		parySuperItem(0) = psng_Length
		parySuperItem(1) = psng_Width
		parySuperItem(2) = psng_Height
		parySuperItem(3) = psngTempWeight
		parySuperItem(4) = plngNumPackages
		
		'Compare the ship alone item against the super package
		'Use weight as the determinent - ignores dimensional weight
		
		If cblnDebugPostageRateAddon Then
			Response.Write "<br /><b>Ship Alone Item</b><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Length: " & paryShipAloneItem(0) & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Width: " & paryShipAloneItem(1) & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Height: " & paryShipAloneItem(2) & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Weight: " & paryShipAloneItem(3) & "<br />" & vbcrlf
			Response.Flush
		End If	'pblnDebug

		If parySuperItem(3) > paryShipAloneItem(3) Then
			psng_Length = parySuperItem(0)
			psng_Width = parySuperItem(1)
			psng_Height = parySuperItem(2)
			psng_Weight = parySuperItem(3)
		Else
			psng_Length = paryShipAloneItem(0)
			psng_Width = paryShipAloneItem(1)
			psng_Height = paryShipAloneItem(2)
			psng_Weight = paryShipAloneItem(3)
		End If
		plng_NumPackages = parySuperItem(4) + paryShipAloneItem(4)

		If cblnDebugPostageRateAddon Then
			Response.Write "<br /><b>Package(s) to be rated</b><br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Length: " & psng_Length & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Width: " & psng_Width & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Height: " & psng_Height & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;Weight: " & psng_Weight & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;NumPackages: " & plngNumPackages & "<br />" & vbcrlf
			Response.Flush
		End If	'pblnDebug
		
		sngLength = psng_Length
		sngWidth = psng_Width
		sngHeight = psng_Height
		sngWeight = psng_Weight
		lngNumPackages = plng_NumPackages
		
		CalculatePackageSize = p_blnExceedsMaxWeight
		
	End Function	'CalculatePackageSize

	'****************************************************************************************************************

End Class

'****************************************************************************************************************

Class clsShippingOption

	Private plngShippingID
	Private pstrShippingType
	Private pstrShippingCode
	Private pstrCarrierCode
	Private pstrshipRate
	Private pcurPostage
	Private pcurInsurance
	Private plngNumPackages
	Private plngTransitTime
	Private pbytsortOrder
	
	Public Property Get TransitTime()
		TransitTime = plngTransitTime
	End Property

	Public Property Let TransitTime(lngTransitTime)
		plngTransitTime = lngTransitTime
	End Property

	Public Property Get NumPackages()
		NumPackages = plngNumPackages
	End Property

	Public Property Let NumPackages(lngNumPackages)
		plngNumPackages = lngNumPackages
	End Property

	Public Property Get ShippingID()
		ShippingID = plngShippingID
	End Property

	Public Property Let ShippingID(lngShippingID)
		plngShippingID = lngShippingID
	End Property

	Public Property Get ShippingType()
		ShippingType = pstrShippingType
	End Property

	Public Property Let ShippingType(strShippingType)
		pstrShippingType = strShippingType
	End Property

	Public Property Get ShippingCode()
		ShippingCode = pstrShippingCode
	End Property

	Public Property Let ShippingCode(strShippingCode)
		pstrShippingCode = strShippingCode
	End Property

	Public Property Get CarrierCode()
		CarrierCode = pstrCarrierCode
	End Property

	Public Property Let CarrierCode(strCarrierCode)
		pstrCarrierCode = strCarrierCode
	End Property

	Public Property Get shipRate()
		shipRate = pstrshipRate
	End Property

	Public Property Let shipRate(strshipRate)
		pstrshipRate = strshipRate
	End Property

	Public Property Get Postage()
		Postage = pcurPostage
	End Property

	Public Property Let Postage(curPostage)
		pcurPostage = curPostage
	End Property

	Public Property Get Insurance()
		Insurance = pcurInsurance
	End Property

	Public Property Let Insurance(curInsurance)
		pcurInsurance = curInsurance
	End Property

	Public Property Get sortOrder()
		sortOrder = pbytsortOrder
	End Property

	Public Property Let sortOrder(bytsortOrder)
		pbytsortOrder = bytsortOrder
	End Property

End Class

'****************************************************************************************************************
'****************************************************************************************************************

	Function LoadOrderItems_SF5(byRef aryOrderItems)
	
	Dim pstrSQL
	Dim pstrSQL_withAttrWeight
	Dim pobjRS
	
	Dim p_dblWeight
	Dim p_dblAttrWeight
	Dim p_dblLength
	Dim p_dblHeight
	Dim p_dblWidth
	Dim p_intQuantity
	Dim p_dblProductBasedShipping
	Dim pblnMustShipFreight
	Dim p_dblFixedShipping
	Dim p_strSpecialShipping
	
	Dim i
	Dim pstrPrevID
	Dim pstrNewID
	
	Dim plngItemCount
	Dim plngSessionID
	Dim pblnIsInitialized
	
		If isArray(maryOrderItems) Then
			On Error Resume Next
			pblnIsInitialized = CBool(UBound(maryOrderItems) < 1)
			If Err.number <> 0 Then
				Err.Clear
			End If
			On Error Goto 0
		Else
			pblnIsInitialized = False
		End If

		If Not pblnIsInitialized Then
			If isArray(aryOrderItems) Then
				plngItemCount = UBound(aryOrderItems) + 1
				ReDim maryOrderItems(UBound(aryOrderItems) + 1)
				For i = 0 To UBound(aryOrderItems)
					'aryOrderItems(numItems)(8) decoder
					'0 - Length - default None
					'1 - Width - default None
					'2 - Height - default None
					'3 - Weight - default None
					'4 - Quantity - default None
					'5 - FixedSize - default True
					'6 - DoNotCombine - default False
					'7 - ProductBasedShipping - default None
					'8 - MustShipFreight - default False
					'9 - FixedShipping - default 0
					'10 - SpecialShipping - default None

					'psngLength, psngWidth, psngHeight, psngWeight, p_intQuantity, pblnFixedSize, pblnDoNotCombine, p_dblProductBasedShipping, pblnMustShipFreight, p_dblFixedShipping, p_strSpecialShipping
					If aryOrderItems(i)(enOrderItem_prodShipIsActive) Then
						p_dblLength = aryOrderItems(i)(enOrderItem_UnitLength)
						p_dblWidth = aryOrderItems(i)(enOrderItem_UnitWidth)
						p_dblHeight = aryOrderItems(i)(enOrderItem_UnitHeight)
						p_dblWeight = aryOrderItems(i)(enOrderItem_prodWeight)
						p_intQuantity = aryOrderItems(i)(enOrderItem_odrdttmpQuantity)
						p_dblProductBasedShipping = aryOrderItems(i)(enOrderItem_prodShip)
						pblnMustShipFreight = aryOrderItems(i)(enOrderItem_MustShipFreight)
						p_dblFixedShipping = aryOrderItems(i)(enOrderItem_prodFixedShippingCharge)
						p_strSpecialShipping = aryOrderItems(i)(enOrderItem_prodSpecialShippingMethods)
					Else
						p_dblLength = 0
						p_dblWidth = 0
						p_dblHeight = 0
						p_dblWeight = 0
						p_intQuantity = 0
						p_dblProductBasedShipping = 0
						pblnMustShipFreight = False
						p_dblFixedShipping = 0
						p_strSpecialShipping = ""
					End If

					'check 
					If Len(p_dblWeight & "") = 0 Then p_dblWeight = 0
					If Len(p_dblLength & "") = 0 Then p_dblLength = 0
					If Len(p_dblWidth & "") = 0 Then p_dblWidth = 0
					If Len(p_dblHeight & "") = 0 Then p_dblHeight = 0
					If Len(p_dblProductBasedShipping & "") = 0 Then p_dblProductBasedShipping = 0
					
					maryOrderItems(i+1) = Array (p_dblLength, p_dblWidth, p_dblHeight, p_dblWeight, p_intQuantity, True, False, p_dblProductBasedShipping, pblnMustShipFreight, p_dblFixedShipping, p_strSpecialShipping)
					
				Next 'i
				pblnIsInitialized = True
			End If
		End If	'Not pblnIsInitialized

		'this next section should be dead code now
		If Not pblnIsInitialized Then
			plngSessionID = Session("SessionID")
			If Len(plngSessionID) = 0 Or Not isNumeric(plngSessionID) Then
				LoadOrderItems_SF5 = False
				Exit Function
			End If
		
			plngItemCount = 0
			If Len(Session("SessionID")) = 0 Then 
				LoadOrderItems_SF5 = False
				Exit Function
			End If
			
			pstrSQL_withAttrWeight = "SELECT sfTmpOrderDetails.odrdttmpID, sfProducts.prodWeight, sfProducts.prodShip, sfProducts.prodLength, sfProducts.prodWidth, sfProducts.prodHeight, sfTmpOrderDetails.odrdttmpQuantity, sfAttributeDetail." & cstrAttrWeight & "" _
									& " FROM (sfTmpOrderAttributes RIGHT JOIN (sfTmpOrderDetails INNER JOIN sfProducts ON sfTmpOrderDetails.odrdttmpProductID = sfProducts.prodID) ON sfTmpOrderAttributes.odrattrtmpOrderDetailId = sfTmpOrderDetails.odrdttmpID) LEFT JOIN sfAttributeDetail ON sfTmpOrderAttributes.odrattrtmpAttrID = sfAttributeDetail.attrdtID" _
									& " WHERE (((sfProducts.prodShipIsActive)=1) AND ((sfTmpOrderDetails.odrdttmpSessionID)=" & plngSessionID & "))"
				'					& " WHERE (((sfProducts.prodShipIsActive)=1) AND ((sfTmpOrderDetails.odrdttmpShipping)>=1) AND ((sfTmpOrderDetails.odrdttmpSessionID)=" & plngSessionID & "))"
			pstrSQL = "SELECT prodShip, sfProducts.prodWeight, prodLength, prodWidth, prodHeight, odrdttmpQuantity" _
					& " FROM sfTmpOrderDetails INNER JOIN sfProducts ON sfTmpOrderDetails.odrdttmpProductID = sfProducts.prodID" _
					& " WHERE (((sfProducts.prodShipIsActive)=1) AND ((sfTmpOrderDetails.odrdttmpSessionID)=" & plngSessionID & "))"
	'				& " WHERE (((sfProducts.prodShipIsActive)=1) AND ((sfTmpOrderDetails.odrdttmpShipping)>=1) AND ((sfTmpOrderDetails.odrdttmpSessionID)=" & plngSessionID & "))"

			If Len(cstrMustShipFreight) > 0 Then pstrSQL = Replace(pstrSQL,"sfProducts.prodWeight,", "sfProducts.prodWeight, " & cstrMustShipFreight & ",",1,1)

			Set pobjRS = CreateObject("ADODB.RECORDSET")
			With pobjRS
				.ActiveConnection = cnn
				.CursorLocation = 2 'adUseClient
				.CursorType = 3 'adOpenStatic
				.LockType = 1 'adLockReadOnly

				On Error Resume Next
				If Err.number <> 0 Then Err.Clear
				If Len(cstrAttrWeight) > 0 Then 
					.Source = pstrSQL_withAttrWeight
					.Open
					If Err.number <> 0 Then
						Err.Clear
						On Error Goto 0
						cstrAttrWeight = ""
						.Source = pstrSQL
						.Open
					End If
				Else
					.Source = pstrSQL
					.Open
					On Error Goto 0
				End If

				If cblnDebugPostageRateAddon Then 
					Response.Write "<strong>Order Loaded</strong><br />"
					Response.Write "&nbsp;&nbsp;There are " & .RecordCount & " item(s) in the order.<br />"
					Response.Flush
				End If
				plngItemCount = 0
				
				For i = 1 To .RecordCount
				
					If Len(cstrAttrWeight) > 0 Then 
						pstrNewID = .Fields("odrdttmpID").value
					Else
						pstrNewID = "x" & pstrPrevID
					End If
					
					If pstrPrevID = pstrNewID Then 
						If Len(cstrAttrWeight) > 0 Then 
							p_dblAttrWeight = Trim(.Fields(cstrAttrWeight) & "")
							If isNumeric(p_dblAttrWeight) Then p_dblWeight = p_dblWeight + p_dblAttrWeight
						End If
					Else
						plngItemCount = plngItemCount + 1
						If Not isArray(aryOrderItems) Then Set aryOrderItems = Nothing

						If plngItemCount = 1 Then
							ReDim aryOrderItems(1)
						Else
							ReDim Preserve aryOrderItems(plngItemCount)
						End If
						
						p_dblWeight = 0
						p_dblWidth = 0
						p_dblLength = 0
						p_dblHeight = 0
						p_dblProductBasedShipping = 0
					
						If Not isNull(.Fields("prodWeight")) Then p_dblWeight = .Fields("prodWeight").value
						If Not isNull(.Fields("prodLength")) Then p_dblLength = .Fields("prodLength").value
						If Not isNull(.Fields("prodHeight")) Then p_dblHeight = .Fields("prodHeight").value
						If Not isNull(.Fields("prodWidth")) Then p_dblWidth = .Fields("prodWidth").value

						If Len(cstrAttrWeight) > 0 Then 
							p_dblAttrWeight = Trim(.Fields(cstrAttrWeight) & "")
							If isNumeric(p_dblAttrWeight) Then p_dblWeight = p_dblWeight + p_dblAttrWeight
						End If

						p_intQuantity = .Fields("odrdttmpQuantity").value
						
						If Len(.Fields("prodShip").Value & "") > 0 Then p_dblProductBasedShipping = .Fields("prodShip").value
						
						If Len(cstrAttrWeight) > 0 Then pstrPrevID = .Fields("odrdttmpID").value
						
						If cblnDebugPostageRateAddon Then
							Response.Write "&nbsp;&nbsp;" & plngItemCount & ": Length =" & p_dblLength & "<br />" & vbcrlf
							Response.Write "&nbsp;&nbsp;" & plngItemCount & ": Width =" & p_dblWidth & "<br />" & vbcrlf
							Response.Write "&nbsp;&nbsp;" & plngItemCount & ": Height =" & p_dblHeight & "<br />" & vbcrlf
							Response.Write "&nbsp;&nbsp;" & plngItemCount & ": Weight =" & p_dblWeight & "<br />" & vbcrlf
							Response.Write "&nbsp;&nbsp;" & plngItemCount & ": Quantity =" & p_intQuantity & "<br />" & vbcrlf
							Response.Write "&nbsp;&nbsp;" & plngItemCount & ": ProductBasedShipping =" & p_dblProductBasedShipping & "<br />" & vbcrlf
							If Len(cstrMustShipFreight) > 0 Then Response.Write "&nbsp;&nbsp;" & plngItemCount & ": MustShipFreight =" & .Fields(cstrMustShipFreight).value & "<br />" & vbcrlf
							Response.Flush
						End If	'cblnDebugPostageRateAddon
					End If	'pstrPrevID = pstrNewID

					'aryOrderItems(numItems)(8) decoder
					'0 - Length - default None
					'1 - Width - default None
					'2 - Height - default None
					'3 - Weight - default None
					'4 - Quantity - default None
					'5 - FixedSize - default True
					'6 - DoNotCombine - default False
					'7 - ProductBasedShipping - default None
					'8 - MustShipFreight - default False
					'9 - FixedShipping - default 0
					'10 - SpecialShipping - default None

					'psngLength, psngWidth, psngHeight, psngWeight, p_intQuantity, pblnFixedSize, pblnDoNotCombine, p_dblProductBasedShipping, pblnMustShipFreight, p_dblFixedShipping, p_strSpecialShipping
					aryOrderItems(plngItemCount) = Array (p_dblLength, p_dblWidth, p_dblHeight, p_dblWeight, p_intQuantity, True, False, p_dblProductBasedShipping, False, 0, "")
					
					'check 
					If Len(aryOrderItems(plngItemCount)(0) & "") = 0 Then aryOrderItems(plngItemCount)(0) = 0
					If Len(aryOrderItems(plngItemCount)(1) & "") = 0 Then aryOrderItems(plngItemCount)(1) = 0
					If Len(aryOrderItems(plngItemCount)(2) & "") = 0 Then aryOrderItems(plngItemCount)(2) = 0
					If Len(aryOrderItems(plngItemCount)(7) & "") = 0 Then aryOrderItems(plngItemCount)(7) = 0

					.MoveNext
				Next 'i
				.Close
			End With
			Set pobjRS = Nothing
		End If	'Not isArray(aryOrderItems)
		
		mdblOrderWeight = 0
		mdblMaxItemWeight = 0
		
		If plngItemCount > 0 Then Call SortOrderItems(maryOrderItems)
		'Now sort the sizes by length, width, height
		
		If cblnDebugPostageRateAddon Then
			Response.Write "&nbsp;&nbsp;Order Weight: " & mdblOrderWeight & "<br />" & vbcrlf
			Response.Write "&nbsp;&nbsp;MaxItemWeight: " & mdblMaxItemWeight & "<br />" & vbcrlf
					If plngItemCount > 0 Then Response.Write "&nbsp;&nbsp;Item(s): " & UBound(aryOrderItems) & "<br />" & vbcrlf
			Response.Flush
		End If	'cblnDebugPostageRateAddon
		
		LoadOrderItems_SF5 = CBool(plngItemCount > 0)

	End Function	'LoadOrderItems_SF5

	'***********************************************************************************************
	
	Sub SortOrderItems(ByRef aryOrderItems)
	
	Dim p_dblWeight
	Dim p_dblAttrWeight
	Dim p_dblLength
	Dim p_dblHeight
	Dim p_dblWidth
	Dim p_intQuantity
	Dim p_dblProductBasedShipping
	
	Dim i
	
		'aryOrderItems(numItems)(10) decoder
		'0 - Length - default None
		'1 - Width - default None
		'2 - Height - default None
		'3 - Weight - default None
		'4 - Quantity - default None
		'5 - FixedSize - default True
		'6 - DoNotCombine - default False
		'7 - ProductBasedShipping - default None
		'8 - MustShipFreight - default False
		'9 - FixedShipping - default 0
		'10 - SpecialShipping - default None

		'psngLength, psngWidth, psngHeight, psngWeight, p_intQuantity, pblnFixedSize, pblnDoNotCombine, p_dblProductBasedShipping, pblnMustShipFreight, p_dblFixedShipping, p_strSpecialShipping
		'aryOrderItems(plngItemCount) = Array (p_dblLength, p_dblWidth, p_dblHeight, p_dblWeight, p_intQuantity, True, False, p_dblProductBasedShipping, False, 0, "")
		
		mdblOrderWeight = 0
		mdblMaxItemWeight = 0
		
		'Now sort the sizes by length, width, height
		Dim p_dblTemp
		For i = 1 To UBound(aryOrderItems)
			p_dblLength = aryOrderItems(i)(0)
			p_dblWidth = aryOrderItems(i)(1)
			p_dblHeight = aryOrderItems(i)(2)
			
			If p_dblWidth > p_dblLength Then
				p_dblTemp = p_dblLength
				p_dblLength = p_dblWidth
				p_dblWidth = p_dblTemp
			End If
			
			If p_dblHeight > p_dblLength Then
				p_dblTemp = p_dblLength
				p_dblLength = p_dblHeight
				p_dblHeight = p_dblTemp
			End If

			If p_dblHeight > p_dblWidth Then
				p_dblTemp = p_dblHeight
				p_dblHeight = p_dblWidth
				p_dblWidth = p_dblTemp
			End If

			aryOrderItems(i)(0) = p_dblLength
			aryOrderItems(i)(1) = p_dblWidth
			aryOrderItems(i)(2) = p_dblHeight

			'Calculate order weight and max Item weight
			mdblOrderWeight = mdblOrderWeight + aryOrderItems(i)(3) * aryOrderItems(i)(4)
			If mdblMaxItemWeight < aryOrderItems(i)(3) Then mdblMaxItemWeight = aryOrderItems(i)(3)
		Next 'i
		
	End Sub	'SortOrderItems

	'***********************************************************************************************

	Sub InitializeOrigin(strOriginStateAbb, strOriginZip, strOriginCountryAbb)
	
	Dim prs
	
		If Len(strOriginZip) > 0 And Len(strOriginCountryAbb) > 0 Then Exit Sub
		
		set prs = CreateObject("ADODB.RECORDSET")
		with prs
			If isObject(cnn) Then
				.ActiveConnection = cnn
			Else
				.ActiveConnection = Connection
			End If
			.CursorLocation = 2 'adUseClient
			.CursorType = 3 'adOpenStatic
			.LockType = 1 'adLockReadOnly
			.Source = "SELECT adminOriginState, adminOriginCountry, adminOriginZip FROM sfAdmin"
			.Open
			
			strOriginStateAbb = trim(prs.Fields("adminOriginState").value)
			strOriginZip	= trim(prs.Fields("adminOriginZip").value)
			strOriginCountryAbb = trim(prs.Fields("adminOriginCountry").value)
			
			If False Then
				Response.Write "<br /><strong>Origin</strong><br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;State: " & strOriginStateAbb & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;ZIP: " & strOriginZip & "<br />" & vbcrlf
				Response.Write "&nbsp;&nbsp;Country: " & strOriginCountryAbb & "<br />" & vbcrlf
			End If
			
			.Close
	
		end with
		Set prs = Nothing

	End Sub	'InitializeOrigin

	'***********************************************************************************************

	Function GetDestinationCountryName(byVal strDestinationCountryAbb)
	
	Dim prs
	
		If Len(strDestinationCountryAbb) = 0 Or isNull(strDestinationCountryAbb) Then Exit Function
		
		set prs = CreateObject("ADODB.RECORDSET")
		with prs
			If isObject(cnn) Then
				.ActiveConnection = cnn
			Else
				.ActiveConnection = Connection
			End If
			.CursorLocation = 2 'adUseClient
			.CursorType = 3 'adOpenStatic
			.LockType = 1 'adLockReadOnly
			.Source = "Select loclctryName from sfLocalesCountry where loclctryAbbreviation='" & Replace(strDestinationCountryAbb, "'", "''") & "'"
			.Open
			
			If Not .EOF Then GetDestinationCountryName = trim(prs.Fields("loclctryName").value)
			
			.Close
	
		end with
		Set prs = Nothing

	End Function	'GetDestinationCountryName

	'****************************************************************************************************************

	Function ssPostageRate_SF5(byRef strShipCode, byVal strDestinationState, byVal strDestinationZip, byVal strDestinationCountryAbb, byVal dblTotalPrice)

	Dim pclsShipping
	Dim pstrDestinationStateAbb
	Dim pstrOriginStateAbb
	Dim pdblShipping
	
		Call LoadOrderItems_SF5(maryOrderItems)
		
		Set pclsShipping= New clsShipping
		With pclsShipping

			.Connection = cnn
			.OriginStateAbb = pstrOriginStateAbb
			.OriginZip = adminOriginZip
			.OriginCountryAbb = adminOriginCountry
			
			.DestinationStateAbb = strDestinationState
			.DestinationZIP = strDestinationZip
			.DestinationCountryAbb = strDestinationCountryAbb
			.DestinationCountryName = GetDestinationCountryName(strDestinationCountryAbb)
			
			'It is up to you to define/calculate the following
			'Comment out if you do not use
			'.ResidentialDelivery = mblnShipResidential
			'.InsideDelivery = mblnIndoorDeliver

			.OrderItems = maryOrderItems
			.MaxItemWeight = mdblMaxItemWeight
			.TotalOrderWeight = mdblOrderWeight
			
			.OrderSubtotal = FormatNumber(dblTotalPrice,2,,,false)
			.DeclaredValue = FormatNumber(dblTotalPrice,2,,,false)
			.Insured = False
			
			If cblnDebugPostageRateAddon And True Then 
				Response.Write "<br /><hr><strong><font size='+2'>Setting Individual item . . .</font></strong><br />" & vbcrlf
				Response.Write "Session(sTotalPrice) =" & Session("sTotalPrice") & "<br />" & vbcrlf
				Response.Write "Session(persistTotalPrice) =" & Session("persistTotalPrice") & "<br />" & vbcrlf
			End If
			
			.ShippingSelection = Request.Form("ShippingSelection")
			pdblShipping = .GetRates(strShipCode)
			sShipMethodName = .ssShippingMethodName
			
			'debugprint "strShipping",strShipping
			'debugprint "sShipMethodName",sShipMethodName
			If CStr(pdblShipping) = "FAIL" And cblnAutomaticallyFindAnyAvailable Then .GetAnyAvailableShipping strShipCode, sShipMethodName, pdblShipping
		End With
		Set pclsShipping= Nothing

		'use the following line for versions PRIOR to v50.03
		'If strShipping = "FAIL" Then strShipping = getShipping(iTotalPur, iPremiumShipping, "FAIL" & "|" & strShipping, strDestinationZip, strDestinationCountryAbb, sTotalPrice)

		'use the following line for versions v50.03 and later
		If strShipping = "FAIL" Then strShipping = getShipping(iTotalPur, iPremiumShipping, "FAIL" & "|" & strShipping, "", "", strDestinationZip, strDestinationCountryAbb, sTotalPrice,"")
		
	End Function	'ssPostageRate_SF5

	'****************************************************************************************************************

	Function getssShippingOptions()

	Dim pstrTemp
	Dim i
	
		If loadssShippingMethods Then
			For i = 0 To UBound(maryShippingMethods)
				If maryShippingMethods(i)(2) Then
					pstrTemp = pstrTemp & "<option value=""" & maryShippingMethods(i)(0) & """ selected>" & maryShippingMethods(i)(1) & "</option>" & vbcrlf
				Else
					pstrTemp = pstrTemp & "<option value=""" & maryShippingMethods(i)(0) & """>" & maryShippingMethods(i)(1) & "</option>" & vbcrlf
				End If
			Next
		End If
		
		getssShippingOptions = pstrTemp
		
	End Function	'getssShippingOptions
	
	'****************************************************************************************************************

	Function isGenericUSPSCodeMatch(byRef strGenericCode, byVal strAvailableCode)

        Select Case strGenericCode
            Case "First-Class", "Priority Mail", "Express Mail": If cblnGenericUSPSMethods And InStr(1, strAvailableCode, strGenericCode) > 0 Then isGenericUSPSCodeMatch = True
            Case Else:
                'Response.Write strGenericCode & " = " & strAvailableCode & "? " & isGenericUSPSCodeMatch & "<br>"
        End Select
		
	End Function	'isGenericUSPSCodeMatch
	
	'****************************************************************************************************************

	Function setGenericUSPSCodeMatch(byVal strAvailableCode)

        Select Case True
            Case InStr(1, strAvailableCode, "Express Mail") > 0, InStr(1, strAvailableCode, "Priority Mail Express") > 0:
                setGenericUSPSCodeMatch = "Express Mail"
            Case InStr(1, strAvailableCode, "Priority Mail") > 0:
                setGenericUSPSCodeMatch = "Priority Mail"
            Case InStr(1, strAvailableCode, "First-Class") > 0:
                setGenericUSPSCodeMatch = "First-Class"
            Case Else:
                setGenericUSPSCodeMatch = strAvailableCode
        End Select

    End Function    'setGenericUSPSCodeMatch

	'****************************************************************************************************************

	Sub adjustUSPSMethods()

    Dim i

		For i = 0 To UBound(maryShippingMethods)
            Select Case True
                Case InStr(1, maryShippingMethods(i)(0), "Express Mail") > 0, InStr(1, maryShippingMethods(i)(0), "Priority Mail Express") > 0:
                    If maryShippingMethods(i)(0) <> "Express Mail" Then maryShippingMethods(i)(0) = ""
                Case InStr(1, maryShippingMethods(i)(0), "Priority Mail") > 0:
                    If maryShippingMethods(i)(0) <> "Priority Mail" Then maryShippingMethods(i)(0) = ""
                Case InStr(1, maryShippingMethods(i)(0), "First-Class") > 0:
                    If maryShippingMethods(i)(0) <> "First-Class" Then maryShippingMethods(i)(0) = ""
                Case Else:
                    'Response.Write maryShippingMethods(i)(0) & "<BR>"
            End Select
        Next 'i

    End Sub

	'****************************************************************************************************************

	Function getssShippingOptions_new(byVal strShipCode)

	Dim pstrTemp
	Dim i
	Dim pblnSelected
	
		If loadssShippingMethods Then
            If cblnGenericUSPSMethods Then Call adjustUSPSMethods
			For i = 0 To UBound(maryShippingMethods)
				If Len(maryShippingMethods(i)(0)) > 0 Then
				If Len(strShipCode) > 0 Then
					pblnSelected = CBool(strShipCode = maryShippingMethods(i)(0))
				Else
					pblnSelected = maryShippingMethods(i)(2)
				End If
				
				If pblnSelected Then
					pstrTemp = pstrTemp & "<option value=""" & maryShippingMethods(i)(0) & """ selected>" & maryShippingMethods(i)(1) & "</option>" & vbcrlf
				Else
					pstrTemp = pstrTemp & "<option value=""" & maryShippingMethods(i)(0) & """>" & maryShippingMethods(i)(1) & "</option>" & vbcrlf
				End If
				End If
			Next
		End If
		
		getssShippingOptions_new = pstrTemp
		
	End Function	'getssShippingOptions_new

	'****************************************************************************************************************

	Function includeShippingMethod(byVal strShipCode)
	
    Dim pblnResult
    Dim pdtNextShip

        'pblnResult = True
        pblnResult = CBool(Len(strShipCode) > 0)
        includeShippingMethod = pblnResult
        Exit Function
       
        Select Case strShipCode
            Case "1DA", "1DP", "1DM"
	            If Time() < CDate("12:00:00 PM") Then
		            pdtNextShip = Date()
	            Else
		            pdtNextShip = DateAdd("d", 1, Date())
	            End If
            
	            'Check date of next shipment and adjust for weekends
	            Select Case DatePart("w", pdtNextShip)
		            Case 1, 7	'Sunday, Saturday
                        pblnResult = False
	            End Select
        End Select

        includeShippingMethod = pblnResult

	End Function    'includeShippingMethod

	'****************************************************************************************************************

	Function getssShippingOptions_Radio(byVal strShipCode)

	Dim pstrTemp
	Dim i
	Dim pblnSelected
	
		If loadssShippingMethods Then
			For i = 0 To UBound(maryShippingMethods)
				If Len(strShipCode) > 0 Then
					pblnSelected = CBool(strShipCode = maryShippingMethods(i)(0))
				Else
					pblnSelected = maryShippingMethods(i)(2)
				End If
				
				If i > 0 Then pstrTemp = pstrTemp & "<br />"
				If pblnSelected Then
					pstrTemp = pstrTemp & "<input type=radio name=Shipping id=Shipping" & i & " value=""" & maryShippingMethods(i)(0) & """ checked><label for=Shipping" & i & ">" & maryShippingMethods(i)(1) & "</label>" & vbcrlf
				Else
					pstrTemp = pstrTemp & "<input type=radio name=Shipping id=Shipping" & i & " value=""" & maryShippingMethods(i)(0) & """><label for=Shipping" & i & ">" & maryShippingMethods(i)(1) & "</label>" & vbcrlf
				End If
			Next
		End If
		
		getssShippingOptions_Radio = pstrTemp
		
	End Function	'getssShippingOptions_Radio

'Const enOrderItem_prodFixedShippingCharge = 45
'Const enOrderItem_prodSpecialShippingMethods = 46
	'****************************************************************************************************************

	Function loadssShippingMethods()

	Dim pblnResult
	Dim prsssShippingMethods
	Dim pstrSQL
	Dim pstrTemp
	Dim i

		pblnResult = False
		
		If isArray(maryShippingMethods) Then
			pblnResult = True
		Else
			If Len(mdblMaxItemWeight) > 0 Then
				pstrSQL = "SELECT ssShippingMethodCode, ssShippingMethodName, ssShippingMethodDefault, ssShippingMethodOfferFreeShippingAbove " _
						& " FROM ssShippingMethods" _
						& " WHERE (ssShippingMethodIsSpecial=0) AND (ssShippingMethodEnabled<>0) AND (ssShippingMethodMinWeight<=" & Replace(mdblMaxItemWeight, "'", "''") & ") AND ((ssShippingMethodMaxWeight>=" & Replace(mdblMaxItemWeight, "'", "''") & ") OR (ssShippingMethodMaxWeight=0))" _
						& " ORDER BY ssShippingMethodOrderBy, ssShippingMethodName"
						'& " WHERE (ssShippingMethodEnabled<>0) AND (ssShippingMethodMinWeight<=" & Replace(mdblMaxItemWeight, "'", "''") & ") AND ((ssShippingMethodMaxWeight>=" & Replace(mdblMaxItemWeight, "'", "''") & ") OR (ssShippingMethodMaxWeight=0))" _
			Else
				pstrSQL = "SELECT ssShippingMethodCode, ssShippingMethodName, ssShippingMethodDefault, ssShippingMethodOfferFreeShippingAbove FROM ssShippingMethods WHERE ssShippingMethodEnabled<>0 Order By ssShippingMethodOrderBy, ssShippingMethodName"
			End If

			Set prsssShippingMethods = CreateObject("ADODB.RECORDSET")
			With prsssShippingMethods
				.CursorLocation = 2 'adUseClient
				If Err.number <> 0 Then Err.Clear
				On Error Resume Next
				.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
				
				If Err.number <> 0 Then
					pstrTemp = "<option value="""">The Postage Rate add-on upgrade does not appear to have been performed</option>"
					pstrTemp = pstrTemp & "<option value="""">Error " & Err.number & ": " & Err.Description & "</option>"
					pstrTemp = pstrTemp & "<option value="""">SQL: " & pstrSQL & "</option>"
					Err.Clear
				Else
					ReDim maryShippingMethods(.RecordCount - 1)
					For i = 1 To .RecordCount
						maryShippingMethods(i-1) = Array(Trim(.Fields("ssShippingMethodCode").Value & ""), Trim(.Fields("ssShippingMethodName").Value & ""), Trim(.Fields("ssShippingMethodDefault").Value & ""), Trim(.Fields("ssShippingMethodOfferFreeShippingAbove").Value & ""))
						.MoveNext
					Next
					pblnResult = True
				End If
				.Close
				On Error Goto 0
				
			End With
			Set prsssShippingMethods = Nothing
		End If
		
		loadssShippingMethods = pblnResult
		
	End Function	'loadssShippingMethods

	'****************************************************************************************************************

	Function getShippingCode()
	
		If cblnssCheckForPO Then
			getShippingCode = mstrShippingCode
		Else
			getShippingCode = Trim(Request.Form("Shipping"))
		End If
	
	End Function	'getShippingCode

	'****************************************************************************************************************

	Function determineSpecialShippingCodes()
	
	Dim pstrEnabledSpecialShippingMethods
	
		If cblnssCheckForPO Then
			getShippingCode = mstrShippingCode
		Else
			getShippingCode = Trim(Request.Form("Shipping"))
		End If
	
	End Function	'getShippingCode
	
	'****************************************************************************************************************
	
	Sub checkForPOBox()

	'Instructions for use:
	'
	'File: verify.asp
	'Find: 	OVC_SaveSubTotalWOD 'SFAE
	'Insert immediately afterwards:	Call checkForPOBox	'Inserted for Sandshot Software P.O. Box check

	'File: incVerify.asp
	'Find:	sShipCode 		= Request.Form("Shipping")
	'Insert immediately afterwards:	sShipCode 		= getShippingCode	'Inserted for Sandshot Software P.O. Box check

	Const pblnLocalDebug = False
	Dim pstrAddressToCheck
	Dim pblnIsPOBox
	Dim pstrNewShipMethod
	
		If Not cblnssCheckForPO Then Exit Sub

		'Supported Shipping Methods
		'<option value="03">FedEx 2day</option>
		'<option value="90">FedEx Home Delivery</option>
		'<option value="01">FedEx Priority</option>
		'<option value="92" selected>U.S. Domestic FedEx Ground Package</option>
		'<option value="Priority">Priority Mail</option>
	
		'Check if shipMethod does not allow PO Boxes
		If (iShipMethod = "01") OR (iShipMethod = "03") OR (iShipMethod = "90") OR (iShipMethod = "92") Then
			If pblnLocalDebug Then Response.Write iShipMethod & " does not permit PO Boxes<br />"
			pblnIsPOBox = False
			pstrAddressToCheck = LCase(Trim(sShipCustAddress1) & Trim(sShipCustAddress2))
			
			'Check for PO Box/variations of same in shipping address
			If Not pblnIsPOBox Then pblnIsPOBox = CBool(Instr(1, pstrAddressToCheck, "pob") > 0)
			If Not pblnIsPOBox Then pblnIsPOBox = CBool(Instr(1, pstrAddressToCheck, "po box") > 0)
			If Not pblnIsPOBox Then pblnIsPOBox = CBool(Instr(1, pstrAddressToCheck, "p.o. box ") > 0)

			If pblnLocalDebug Then Response.Write "pblnIsPOBox: " & pblnIsPOBox & "<br />"
			If pblnIsPOBox Then
				Select Case iShipMethod
					Case "": pstrNewShipMethod = "Priority"	'
					Case Else:
						pstrNewShipMethod = "Priority"	'Priority Mail
						'sShipMethodName = "Priority Mail"	'Priority Mail
				End Select
				iShipMethod = pstrNewShipMethod
				mstrShippingCode = pstrNewShipMethod
				sShipMethodName = getNameWithID("ssShippingMethods",iShipMethod,"ssShippingMethodCode","ssShippingMethodName",1)
			End If
		End If

		If pblnLocalDebug Then Response.Write "iShipMethod: " & iShipMethod & "<br />"
		If pblnLocalDebug Then Response.Write "sShipMethodName: " & sShipMethodName & "<br />"

	End Sub	'checkForPOBox

'****************************************************************************************************************
'****************************************************************************************************************

Const cstrAttrWeight = ""					'Set to field name for using attribute sensitive weights
'Const cstrAttrWeight = "attrdtWeight"	'Example

Const cstrMustShipFreight = ""				'Set to field name if some products must ship by Freight
'Const cstrMustShipFreight = "Shiptruck"	'Example

Const cblnssUsePostageRate = True
Const cblnssCheckForPO = False

Dim cblnDebugPostageRateAddon
cblnDebugPostageRateAddon = Len(Session("ssDebug_PostageRate")) > 0
'cblnDebugPostageRateAddon = True

Dim maryOrderItems
Dim mdblOrderWeight
Dim mdblMaxItemWeight









































'------------------------------------------------------------------
' Returns shipping amount
' Returns a string
'------------------------------------------------------------------
Function GetShipping(byVal iTotalPur, byVal iPremiumShipping, byVal aCheck, byVal dCity, byVal dState, byVal dZip, byVal dCountry, byVal sTotalPrice, byVal sType)

Dim SQL, sShipCode, sProdID, sShipping
Dim oCountry, oZip, oCity, oState
Dim iLength, iWidth, iHeight, iWeight
Dim iQuantity, iShipType, iShipMin, iSpcShipAmt,uspsUsername,uspsPassword, CanadaPostRefNum  'JF added for Canada
Dim rsProdShipping, rsShipping, ups, arrCheck, sCheck, sErrMsg, obj, posit
Dim sFreeship,TotaL_with_Attributes
Dim boQty,shpQty,allQty 'SFAE
dim noShip
dim bUsingBackup,lmtPrice
    
	bUsingBackup=false
	iShipMin 		= adminShipMin 
	iSpcShipAmt 	= adminSpcShipAmt
	
	sShipCode 		= Request.Form("Shipping")

	'if the order from didn't have the shipping combo box +JF
	if trim(sShipCode) = "" then
		GetShipping="0"
		exit function
	end if

	posit =instr(sShipCode,",")
	If posit > 0 Then
		sShipCode = left(trim(sShipCode),posit-1) 'FreeShipping''''''''''''''''''''''
	End If
    
	oCity 			= "x"  'JF Seems to work with bogus City as long as Zip is correct, no state in DB
	oState 			= "x"  'JF Seems to work with bogus State as long as Zip is correct, no state in DB
	oZip 			= adminOriginZip
	oCountry 		= adminOriginCountry

    If cblnSF5AE Then
		If NOT isnull(ltlUN)then  
		  dim ltlID,ltlEmail
		  ltlID = ltlUN
		  ltlEmail = ltlEmail
		end if   
    End If 
    
   	If iTotalPur = 0 and posit > 0 Then
    	getShipping = "0"
    	Exit Function
	elseIf ((adminFreeShippingIsActive = "1") AND (cDbl(adminFreeShippingAmount) <= cDbl(sTotalPrice))) Then
	  if posit > 0 then	
		getShipping = "0"'
		sShipping = "0"
	    Exit Function
	  end if  
	end if
  
	'Collect UPS Error Message
	arrCheck 	= split(aCheck, "|")
	If aCheck 	<> "" Then
		sCheck = arrCheck(0)
		sErrMsg = arrCheck(1)
	Else
		sCheck = ""
	End If
	
	'if Carrier Failed, use secondary Shipping method
	If sCheck <> "FAIL" Then
		iShipType = adminShipType
	Else
		Response.Write "<font face=verdana size=2><b><center>Carrier Based Shipping has failed and the secondary Shipping Method is being used<br />Error Description: " & sErrMsg & "</center></b></font>"
		iShipType = adminShipType2
		sShipMethodName = "	Regular Shipping" 
		'Guard against an infanite loop
		If iShipType = 2 Then iShipType = 1  'changes it to valuebased shipping

	End If

	Set rsShipping = CreateObject("ADODB.RecordSet")

	If iShip > 0 Then
		Select Case iShipType
			Case 1	'Zone Based Shipping
			
			Case 2	'Carrier shipping
				Call ssPostageRate_SF5(sShipping, sShipCode, oZip, oCountry, dZip, dCountry, iTotalPur, iPremiumShipping, sTotalPrice)
			Case 3	'Product Based
				'handled in ssCartContents and/or ssmodShipping.asp

		End Select

	End If
	
	'Default Minimum Shipping and premium shipping
	If iShipMin = "" Then iShipMin = 0
	If iSpcShipAmt = "" Then iSpcShipAmt = 0
	
	If bUsingBackup Then
		getShipping = formatNumber(sShipping, 2)  
	Else
		'Add in Premium shipping, check minimum shipping, apply shipping sale
		If iPremiumShipping = "1" Then 								
			sShipping = CDbl(iSpcShipAmt) + CDbl(sShipping)
		End If	
	
		If isnumeric(sShipping) then '#313
    		 If CDbl(sShipping) < CDbl(iShipMin) Then sShipping = iShipMin
		End If
	
		If bLtl Then
			getShipping = "@" & sShipping & "|" & arrLtl(1)
			session("sltl") = "@" & sShipping & "|" & arrLtl(1)
		else
			 getShipping = formatNumber(sShipping, 2)  
		End if   
	End If

	closeObj(rsShipping)
	closeObj(rsProdShipping)
	 
End Function

Function get_ltl(dZip, sContainer, iLength, iWidth, iHeight, sProdId, iQuantity, ltlUN, ltlEmail, oZip, iWeight, sType, boQty, shpQty, bShiprates, iTotalpur, iPremiumShipping, aCheck, dCity, dState, dCountry, sTotalPrice, iship, sShipmethodname)
dim sReturn,iClass
dim objFQRating
dim i,dblLtlPRice, ltlIndex
'check user
 Dim arrltl()
    Dim arrLTL2()
    Dim arrltla()
    Dim arrLTLa2()
    Dim arrltlb()
    Dim arrLTLb2()
    Dim arrltlc()
    Dim arrLTLc2()
    Dim TempRate
    Dim TempBillRate
    Dim TempBORate 
    Dim ii
    Dim f

if trim(ltlUn) = "" or trim(ltlEmail) = "" then
 get_LTL = "Not Registered for LTL Carriers Rate Service"
 exit function
end if  



on error resume next
'Create The Rating Object
Set objFQRating = CreateObject("FQRating.cFQRating")



 if err.number <> 0 then

  get_LTL = "LTL component is not properly installed."
 exit function
end if

 iClass = 50

'Populate the Email/Password Properties to Log-In
'set index for backorders...

if not isnull(Session("LTLIndex")) or Session("LTLIndex") <> "" then
	ltlIndex = cint(Session("LTLIndex")) 
	if ltlIndex < 1 then ltlIndex = 1
elseif ltlIndex < 1 or not isnumeric(ltlIndex) then 
	ltlIndex = 1
end if


if Session("SpecialBilling") =1 AND sType="All" THEN			

	If boQty <> 0 And Trim(boQty) <> "" Then
		Session("BackOrderPrices") = ""
		Session("BackOrderCarriers") = ""
		Session("BackOrderOptionIDs") = ""
		Session("BackOrderTransits") = ""
		Session("backordershipping") = getShipping(iTotalpur, iPremiumShipping, aCheck, dCity, dState, dZip, dCountry, sTotalPrice, "BackOrder", iship, sShipmethodname)
	End If

	If shpQty <> 0 And Trim(boQty) <> "" Then
		Session("ShippedPrices") = ""
		Session("ShippedCarriers") = ""
		Session("ShippedOptionIDs") = ""
		Session("ShippedTransits") = ""
		Session("BillShipping") = getShipping(iTotalpur, iPremiumShipping, aCheck, dCity, dState, dZip, dCountry, sTotalPrice, "Shipped", iship)
	End If
end if		
If Session("specialbilling") = 1 And LCase(sType) = "all" Then
'do nothing
Else

'NOTE: TOTALLY BOGUS RESULTS WHEN USING THE TEST ACCOUNT''''''''''''''''''
objfqrating.Email = ltlEmail  '"xmltest@freightquote.com"
objfqrating.Password =  ltlUN  ' "xml"

'Set the Origin/Destination Zip Codes
	objFQRating.oaddress.zip =ozip
	objFQRating.daddress.zip =dzip

	'Populate the required Product Properties
	objFQRating.FQProds.Class1 =iClass
	objFQRating.FQProds.Description1 = sprodid
	objFQRating.FQProds.PackageType1 = sContainer
	objFQRating.FQProds.Pieces1 = iQuantity
	'---------------------------------------------------------------------------------
	'Issue #255 
	if iWeight > 0 and iWeight <1 then 
		iWeight=1
	end if
	'---------------------------------------------------------------------------------

	objFQRating.FQProds.Weight1 = iWeight
	objFQRating.BILLTO = "SITE"

	'Run Get Quote to get the quote
    objFQRating.GetQuote
End If


If Session("specialbilling") = 1 And LCase(sType) = "all" Then
'go through this to combine billing and backorder
    
If Trim(Session("BackOrderPrices")) <> "" And Trim(Session("ShippedPrices")) <> "" Then
    arrltl = Split(Session("BackOrderPrices"), "|")
    arrLTL2 = Split(Session("ShippedPrices"), "|")
    arrltla = Split(Session("BackOrderCarriers"), "|")
    arrLTLa2 = Split(Session("ShippedCarriers"), "|")
    arrltlb = Split(Session("BackOrderOptionIDs"), "|")
    arrLTLb2 = Split(Session("ShippedOptionIDs"), "|")
    arrltlc = Split(Session("BackOrderTransits"), "|")
    arrLTLc2 = Split(Session("ShippedTransits"), "|")
    Dim tempcount
    
    
        'close existing tables
        'sReturn = sReturn & "</table></td></tr></table></td></tr>" & vbCrLf
        sReturn = sReturn & "</table></td></tr>" & vbCrLf
        sReturn = sReturn & "<script>" & Chr(13) & vbCrLf
        '''''
        'build java
        sReturn = sReturn & "<!--" & Chr(13) & vbCrLf
        sReturn = sReturn & "function selectcarrier(chkindex,pQuoteID,pOptionID,sLTLCarrier) {" & vbCrLf
        sReturn = sReturn & "var ichkcount =" & objfqRating.FQResults.Count & ";" & vbCrLf
        sReturn = sReturn & "var e;" & vbCrLf
       ' sReturn = sReturn & "for (var i = 0; i < ichkcount - 1; i++) {" & vbCrlf
    '   sReturn = sReturn & "  document.frmLTL.Carrier[i].checked =false; " & vbCrlf
     '   sReturn = sReturn & "  }" & vbCrlf
      '  sReturn = sReturn & "  document.frmLTL.Carrier[chkindex -1 ].checked = true; " & vbCrlf
         sReturn = sReturn & "document.frmLTL.action = " & Chr(34) & "verify.asp?OptionID=" & Chr(34) & " + chkindex ;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.OptionID.value = chkindex;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.ltlrate.value = pOptionID;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.ltlcarrier.value = sLTLCarrier;" & vbCrLf
        'sReturn = sReturn & "alert(sLTLCarrier);"  & vbCrlf
        sReturn = sReturn & "document.frmLTL.submit();" & vbCrLf
        sReturn = sReturn & "}" & Chr(13) & vbCrLf
        sReturn = sReturn & "//-->" & Chr(13) & vbCrLf
        sReturn = sReturn & "</script>" & vbCrLf
        'build form and new tables
        sReturn = sReturn & "<form method=post name=frmLTL action=verify.asp>"
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=100% align=center class='tdContent2'>" & vbCrLf
        sReturn = sReturn & "<table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=100% align=left class='tdMiddleTopBanner'><font class='Middle_Top_Banner_Small'><B>Select a shipping Option:</font></B></td>" & vbCrLf
        sReturn = sReturn & "</tr>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td>" & vbCrLf
        sReturn = sReturn & "<table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=10% align=center class='tdContentBar' ><B>Select</B></td>" & vbCrLf
        sReturn = sReturn & "<td width=40% align=center class='tdContentBar' ><B>Method</B></td>" & vbCrLf
      If iConverion = 1 Then
          sReturn = sReturn & "<td width=10% align=center class='tdContentBar' ><B>Transit Time (days)</B></td>" & vbCrLf
          sReturn = sReturn & "<td width=40% align=center class='tdContentBar' ><B>Rate</B></td>" & vbCrLf
      Else
          sReturn = sReturn & "<td width=25% align=center class='tdContentBar' ><B>Transit Time (days)</B></td>" & vbCrLf
          sReturn = sReturn & "<td width=25% align=center class='tdContentBar' ><B>Rate</B></td>" & vbCrLf
       
      End If
        sReturn = sReturn & "</tr>" & vbCrLf

    'build results table
    tempcount = 0
     For i = 1 To UBound(arrltl)
        For ii = 1 To UBound(arrLTL2)
               If Trim(arrltla(i)) = Trim(arrLTLa2(ii)) Then
               tempcount = tempcount + 1

               sReturn = sReturn & "<tr>" & vbCrLf

               If i = ltlIndex Then
                 sShipmethodname = arrltla(i)
'                    arrltl = Split(Session("BackOrderPrices"), "|")
'                    arrLTL2 = Split(Session("ShippedPrices"), "|")

                     If Trim(Session("BackOrderPrices")) <> "" And Trim(Session("ShippedPrices")) <> "" Then
                        TempRate = "$" & CDBL(Mid(arrltl(i), 1)) + CDBL(Mid(arrLTL2(i), 1))
                     ElseIf Trim(Session("BackOrderPrices")) <> "" Then
                        TempRate = arrltl(i)
                     ElseIf Trim(Session("ShippedPrices")) <> "" Then
                        TempRate = arrLTL2(i)
                     End If
                     dblLTLPrice = TempRate * CDBL(bShiprates)
                     'sReturn = sReturn & "<td width=5% align=center><input type=radio name=Carrier checked =true onClick=selectcarrier(" & i & "," & objFQRating.FQResults.item(1).OptionID & "," & chr(34) & TempRate & chr(34) & "," & chr(34) & objFQRating.FQResults.item(i).Carrier & chr(34) & ")></td>" & vbCrlf
                     sReturn = sReturn & "<td width=5% align=center  class='tdContent' ><input type=radio name=Carrier value=" & Chr(34) & arrltla(i) & Chr(34) & " checked =true onClick=selectcarrier(" & i & "," & arrltlb(i) & "," & Chr(34) & TempRate & Chr(34) & ",this.value)></td>" & vbCrLf
                  
             Else

                   arrltl = Split(Session("BackOrderPrices"), "|")
                   arrLTL2 = Split(Session("ShippedPrices"), "|")
                    If Trim(Session("BackOrderPrices")) <> "" And Trim(Session("ShippedPrices")) <> "" Then
                     TempRate = "$" & CDBL(Mid(arrltl(i), 1)) + CDBL(Mid(arrLTL2(i), 1))
                    ElseIf Trim(Session("BackOrderPrices")) <> "" Then
                     TempRate = arrltl(i)
                    ElseIf Trim(Session("ShippedPrices")) <> "" Then
                     TempRate = arrLTL2(i)
                    End If
                sReturn = sReturn & "<td width=5% align=center class='tdContent' ><input type=radio name=Carrier value=" & Chr(34) & arrltla(i) & Chr(34) & " onClick=selectcarrier(" & i & "," & arrltlb(i) & "," & Chr(34) & TempRate & Chr(34) & ",this.value)></td>" & vbCrLf
             End If
             sReturn = sReturn & "<td width=45% align=center class='tdContent' >" & arrltla(i) & "</td>" & vbCrLf
             sReturn = sReturn & "<td width=2% align=center class='tdContent' >" & arrltlc(i) & "</td>" & vbCrLf
             If iConverion = 1 Then
               sReturn = sReturn & "<td width=45% align=center class='tdContent' >" & "<script> document.write(""" & TempRate & " = ("" + OANDAconvert(" & CDBL(TempRate) & ", " & Chr(34) & CurrencyISO & Chr(34) & ") + "")"");</script></td>" & vbCrLf
             Else
               sReturn = sReturn & "<td width=33% align=center  class='tdContent' >" & "<B>" & TempRate & "</B></td>" & vbCrLf
             End If
            sReturn = sReturn & "</tr>" & vbCrLf
            End If
       Next
    Next
      ' get request form variable for re-submit
      
     ' ERR.Clear
     ' On Error Resume Next
     
      For f = 1 To Request.Form.Count
        sReturn = sReturn & "<input type=hidden name= " & Request.Form.Key(f) & " value= " & Request.Form.Item(f) & ">"
      Next
    'close all tables
    sReturn = sReturn & "</table>" & vbCrLf
    sReturn = sReturn & "</td>"
    sReturn = sReturn & "</tr>"
    sReturn = sReturn & "</table>" & vbCrLf
    
    'attach hidden variables to form
    sReturn = sReturn & "<input type=hidden name=QuoteID>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=OptionID>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=FromVerify value =" & Chr(34) & "1" & Chr(34) & ">" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=ltlrate>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=ltlcarrier>" & vbCrLf
    sReturn = sReturn & "</td>"
    sReturn = sReturn & "</tr>"
    sReturn = sReturn & "</form>" & vbCrLf
    
    'build java for re-submit
    sReturn = sReturn & "<script>" & Chr(13) & vbCrLf
    sReturn = sReturn & "<!--" & Chr(13) & vbCrLf
    sReturn = sReturn & "function resetme(chkID) {" & vbCrLf
'    sReturn = sReturn & "alert(chkID);" & vbCrlf
    sReturn = sReturn & "var ichkcount =" & tempcount & ";" & vbCrLf
    sReturn = sReturn & "for (var i = 0; i < ichkcount - 1; i++) {" & vbCrLf
    sReturn = sReturn & "  document.frmLTL.Carrier[i].checked =false; " & vbCrLf
    sReturn = sReturn & "  }" & vbCrLf
    sReturn = sReturn & "document.frmLTL.Carrier[chkID -1].checked =true;  " & vbCrLf
    sReturn = sReturn & "}" & Chr(13) & vbCrLf
    sReturn = sReturn & "//-->" & Chr(13) & vbCrLf
    sReturn = sReturn & "</script>" & vbCrLf
    'reopen old tables
    sReturn = sReturn & "<td width=100% align=center class='tdContent2'>" & vbCrLf
    sReturn = sReturn & "<table border= 0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
    sReturn = sReturn & "<tr><td class='tdContent2'>" & vbCrLf
    sReturn = sReturn & " <table border=0 width=100% cellspacing=0 cellpadding=2 class='tdContent2'>" & vbCrLf




    get_ltl = CDBL(dblLTLPrice) & "|" & sReturn
    Session("sLTL") = FormatCurrency(dblLTLPrice) & "|" & sReturn
    
Else

    sReturn = sReturn & "No Carriers Found." & vbCrLf
'    sReturn = sReturn & objfqRating.XMLQuoteResponse & vbCrLf
    get_ltl = sReturn
End If


Else

If objfqRating.FQResults.Count > 0 Then
        'close existing tables
        'sReturn = sReturn & "</table></td></tr></table></td></tr>" & vbCrLf
        sReturn = sReturn & "</table></td></tr>" & vbCrLf
        sReturn = sReturn & "<script>" & Chr(13) & vbCrLf
        '''''
        'build java
        sReturn = sReturn & "<!--" & Chr(13) & vbCrLf
        sReturn = sReturn & "function selectcarrier(chkindex,pQuoteID,pOptionID,sLTLCarrier) {" & vbCrLf
        sReturn = sReturn & "var ichkcount =" & objfqRating.FQResults.Count & ";" & vbCrLf
        sReturn = sReturn & "var e;" & vbCrLf
       ' sReturn = sReturn & "for (var i = 0; i < ichkcount - 1; i++) {" & vbCrlf
    '   sReturn = sReturn & "  document.frmLTL.Carrier[i].checked =false; " & vbCrlf
     '   sReturn = sReturn & "  }" & vbCrlf
      '  sReturn = sReturn & "  document.frmLTL.Carrier[chkindex -1 ].checked = true; " & vbCrlf
         sReturn = sReturn & "document.frmLTL.action = " & Chr(34) & "verify.asp?OptionID=" & Chr(34) & " + chkindex ;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.OptionID.value = chkindex;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.ltlrate.value = pOptionID;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.ltlcarrier.value = sLTLCarrier;" & vbCrLf
        'sReturn = sReturn & "alert(sLTLCarrier);"  & vbCrlf
        sReturn = sReturn & "document.frmLTL.submit();" & vbCrLf
        sReturn = sReturn & "}" & Chr(13) & vbCrLf
        sReturn = sReturn & "//-->" & Chr(13) & vbCrLf
        sReturn = sReturn & "</script>" & vbCrLf
        
        'build form and new tables
        sReturn = sReturn & "<form method=post name=frmLTL action=verify.asp>"
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=100% align=center class='tdContent2' >" & vbCrLf
        sReturn = sReturn & "<table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=100% align=left class='tdContentBar'><font class='Middle_Top_Banner_Small'><B>Select a shipping Option:</font></B><br />&nbsp;</td>" & vbCrLf
        sReturn = sReturn & "</tr>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td>" & vbCrLf
        sReturn = sReturn & "<table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=10% align=center class='tdContentBar' ><B>Select</B></td>" & vbCrLf
        sReturn = sReturn & "<td width=40% align=center class='tdContentBar' ><B>Method</B></td>" & vbCrLf
      If iConverion = 1 Then
          sReturn = sReturn & "<td width=10% align=center class='tdContentBar' ><B>Transit Time (days)</B></td>" & vbCrLf
          sReturn = sReturn & "<td width=40% align=center class='tdContentBar' ><B>Rate</B></td>" & vbCrLf
      Else
          sReturn = sReturn & "<td width=25% align=center class='tdContentBar' ><B>Transit Time (days)</B></td>" & vbCrLf
          sReturn = sReturn & "<td width=25% align=center class='tdContentBar' ><B>Rate</B></td>" & vbCrLf
       
      End If
        sReturn = sReturn & "</tr>" & vbCrLf

    'build results table
    
     For i = 1 To objfqRating.FQResults.Count
               
               sReturn = sReturn & "<tr>" & vbCrLf
               If sType = "BackOrder" Then
                   Session("BackOrderPrices") = Session("BackOrderPrices") & "|" & objfqRating.FQResults.Item(i).Rate
                   Session("BackOrderCarriers") = Session("BackOrderCarriers") & "|" & objfqRating.FQResults.Item(i).Carrier
                   Session("BackOrderOptionIDs") = Session("BackOrderOptionIDs") & "|" & objfqRating.FQResults.Item(i).OptionID
                   Session("BackOrderTransits") = Session("BackOrderTransits") & "|" & objfqRating.FQResults.Item(i).Transit
               ElseIf sType = "Shipped" Then
                   
                   Session("ShippedPrices") = Session("ShippedPrices") & "|" & objfqRating.FQResults.Item(i).Rate
                   Session("ShippedCarriers") = Session("ShippedCarriers") & "|" & objfqRating.FQResults.Item(i).Carrier
                   Session("ShippedOptionIDs") = Session("ShippedOptionIDs") & "|" & objfqRating.FQResults.Item(i).OptionID
                   Session("ShippedTransits") = Session("ShippedTransits") & "|" & objfqRating.FQResults.Item(i).Transit
               End If
               If i = ltlIndex Then
                 sShipmethodname = objfqRating.FQResults.Item(i).Carrier
 
                     dblLTLPrice = objfqRating.FQResults.Item(i).Rate
                     dblLTLPrice = dblLTLPrice * CDBL(bShiprates)
                     dblLTLPrice = FormatCurrency(CStr(dblLTLPrice))
                     TempRate = objfqRating.FQResults.Item(i).Rate
                     TempRate = TempRate * CDBL(bShiprates)
                     TempRate = FormatCurrency(TempRate)
                     sReturn = sReturn & "<td class='tdContent' width=5% align=center> <input type=radio name=Carrier value=" & Chr(34) & objfqRating.FQResults.Item(i).Carrier & Chr(34) & " checked =true onClick=selectcarrier(" & i & "," & objfqRating.FQResults.Item(i).OptionID & "," & Chr(34) & TempRate & Chr(34) & ",this.value)></td>" & vbCrLf
   '              End If
             Else
    
                   TempRate = objfqRating.FQResults.Item(i).Rate
                   TempRate = CDBL(TempRate) * CDBL(bShiprates)
                   TempRate = FormatCurrency(TempRate)
    '             End If
                sReturn = sReturn & "<td class='tdContent' width=5% align=center><input type=radio name=Carrier value=" & Chr(34) & objfqRating.FQResults.Item(i).Carrier & Chr(34) & " onClick=selectcarrier(" & i & "," & objfqRating.FQResults.Item(i).OptionID & "," & Chr(34) & TempRate & Chr(34) & ",this.value)></td>" & vbCrLf
             End If
             sReturn = sReturn & "<td class='tdContent' width=45% align=center>" & objfqRating.FQResults.Item(i).Carrier & "</td>" & vbCrLf
             sReturn = sReturn & "<td class='tdContent' width=2% align=center>" & objfqRating.FQResults.Item(i).Transit & "</td>" & vbCrLf
             
             If iConverion = 1 Then
                        
             sReturn=sReturn & "<td class='td Content' width=45% align=center>" & "<script> document.write(""" & TempRate & " = ("" + OANDAconvert(" & cDbl(TempRate) & ", """ & CurrencyISO & """) + "")"");</script></td>" & vbCrLf
             '  sReturn = sReturn & "<td class='tdContent' width=45% align=center>HHH<font class='ContentBar_Small'>" & "<script> document.write(""" & TempRate & " = ("" + OANDAconvert(" & cDbl(TempRate) & ", " & Chr(34) & CurrencyISO & Chr(34) & ") + "")"");</script></font></i></td>" & vbCrLf
             
             Else
             
               sReturn = sReturn & "<td class='tdContent' width=33% align=center>" & "<B>" & TempRate & "</B></td>" & vbCrLf
             End If
            sReturn = sReturn & "</tr>" & vbCrLf
            
    Next
      ' get request form variable for re-submit
      
     ' ERR.Clear
     ' On Error Resume Next
      For f = 1 To Request.Form.Count
        sReturn = sReturn & "<input type=hidden name= " & Request.Form.Key(f) & " value= " & Request.Form.Item(f) & ">"
      Next
      
    'close all tables
    sReturn = sReturn & "</table>" & vbCrLf
    sReturn = sReturn & "</td>"
    sReturn = sReturn & "</tr>"
    sReturn = sReturn & "</table>" & vbCrLf
    
    'attach hidden variables to form
    sReturn = sReturn & "<input type=hidden name=QuoteID>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=OptionID>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=FromVerify value =" & Chr(34) & "1" & Chr(34) & ">" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=ltlrate>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=ltlcarrier>" & vbCrLf
    sReturn = sReturn & "</td>"
    sReturn = sReturn & "</tr>"
    sReturn = sReturn & "</form>" & vbCrLf
    'build java for re-submit
    sReturn = sReturn & "<script>" & Chr(13) & vbCrLf
    sReturn = sReturn & "<!--" & Chr(13) & vbCrLf
    sReturn = sReturn & "function resetme(chkID) {" & vbCrLf
'    sReturn = sReturn & "alert(chkID);" & vbCrlf
    sReturn = sReturn & "var ichkcount =" & objfqRating.FQResults.Count & ";" & vbCrLf
    sReturn = sReturn & "for (var i = 0; i < ichkcount - 1; i++) {" & vbCrLf
    sReturn = sReturn & "  document.frmLTL.Carrier[i].checked =false; " & vbCrLf
    sReturn = sReturn & "  }" & vbCrLf
    sReturn = sReturn & "document.frmLTL.Carrier[chkID -1].checked =true;  " & vbCrLf
    sReturn = sReturn & "}" & Chr(13) & vbCrLf
    sReturn = sReturn & "//-->" & Chr(13) & vbCrLf
    sReturn = sReturn & "</script>" & vbCrLf
    'reopen old tables
    sReturn = sReturn & "<td width=100% align=center  class='tdContent2'>" & vbCrLf
    sReturn = sReturn & "<table border= 0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
    sReturn = sReturn & "<tr><td class='tdContent2'>" & vbCrLf
    sReturn = sReturn & " <table border=0 width=100% cellspacing=0 cellpadding=2 class='tdContent2'>" & vbCrLf




    get_ltl = cDbl(dblLTLPrice) & "|" & sReturn
    Session("sLTL") = FormatCurrency(dblLTLPrice) & "|" & sReturn
    'Response.Write "S"
Else

    sReturn = sReturn & "No Carriers Found." & vbCrLf
    sReturn = sReturn & objfqRating.XMLQuoteResponse & vbCrLf
    get_ltl = sReturn
End If
End If
Set objfqRating = Nothing
End Function

'***********************************************************************************************

Function GetExpectedDelivery(byVal strCountry, byVal strLocale, byVal bytMethod)

Dim pstrTemp
Dim pdtNextShip
Dim pdtDelivery

	'If time is after xxx then can't ship until tomorrow
	If Time() < CDate("12:00:00 PM") Then
		pdtNextShip = Date()
	Else
		pdtNextShip = DateAdd("d", 1, Date())
	End If

	'Check date of next shipment and adjust for weekends
	Select Case DatePart("w", pdtNextShip)
		Case 1	'Sunday
			pdtNextShip = DateAdd("d", 1, pdtNextShip)
		Case 7	'Saturday
			pdtNextShip = DateAdd("d", 2, pdtNextShip)
		Case Else
	End Select
	
	Select Case bytMethod
		Case 0: 'Ground
			pstrTemp = "3 - 7 days"
		Case 1: '4 day
			pdtDelivery = DateAdd("d", 4, pdtNextShip)
		Case 2: '2 day
			pdtDelivery = DateAdd("d", 2, pdtNextShip)
		Case 3: 'Next day
			pdtDelivery = DateAdd("d", 1, pdtNextShip)
	End Select
		
	'Correct for weekend deliveries
	If isDate(pdtDelivery) Then
		Select Case DatePart("w", pdtDelivery)
			Case 1	'Sunday
				pdtDelivery = DateAdd("d", 1, pdtDelivery)
				pstrTemp = WeekdayName(DatePart("w",pdtDelivery))
			Case 7	'Saturday
'				pdtDelivery = DateAdd("d", 2, pdtDelivery)
				pstrTemp = WeekdayName(DatePart("w",pdtDelivery))
			Case Else
				pstrTemp = WeekdayName(DatePart("w",pdtDelivery))
		End Select
	End If
	
	GetExpectedDelivery = pstrTemp

End Function	'GetExpectedDelivery

'***********************************************************************************************

Function getShipmentDate()

Const orderCutoffTime = "5:00:00 PM"

Dim pdtNextShip

	Select Case DatePart("w", Date())
		Case 4	'Wednesday
			If Time() < CDate(orderCutoffTime) Then
				pdtNextShip = DateAdd("d", 2, Date())
			Else
				pdtNextShip = DateAdd("d", 5, Date())
			End If
		Case 5	'Thursday
			If Time() < CDate(orderCutoffTime) Then
				pdtNextShip = DateAdd("d", 4, Date())
			Else
				pdtNextShip = DateAdd("d", 5, Date())
			End If
		Case 6	'Friday
			pdtNextShip = DateAdd("d", 4, Date())
		Case 7	'Saturday
			pdtNextShip = DateAdd("d", 4, Date())
		Case 1	'Sunday
			pdtNextShip = DateAdd("d", 3, Date())
		Case Else
			If Time() < CDate(orderCutoffTime) Then
				pdtNextShip = DateAdd("d", 2, Date())
			Else
				pdtNextShip = DateAdd("d", 3, Date())
			End If
	End Select
	
	getShipmentDate = pdtNextShip
	
End Function	'getShipmentDate

'***********************************************************************************************

Sub DisplayShippingTimeMessage

	Response.Write "This order should ship on " & FormatDateTime(getShipmentDate, 1) & "<br />"
	'Response.Write WeekdayName(DatePart("w",pdtNextShip))
	
End Sub	'DisplayShippingTimeMessage

%>
