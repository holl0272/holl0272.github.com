<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Release Version   2.00.002	                                                *
'*   Release Date      December 16, 2003										*
'*   Revision Date     March 13, 2005											*
'*                                                                              *
'*	 2.00.002 (March 13, 2005)													*
'*	 - Updated U.S.P.S. tracking link											*
'*	 - Added multiple tracking number per order capability						*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Const cblnUseNewShipMethodIDs = True	'Set to False if you have a prior version of Order Manager and used UPS as item 1
Dim maryShipMethods(3)

	'array
	'0 - Name
	'1 - Tracking URL
	If cblnUseNewShipMethodIDs Then
		maryShipMethods(0) = Array("UPS","http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=1&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=")
		'maryShipMethods(1) = Array("U.S.P.S.","http://www.framed.usps.com/cgi-bin/cttgate/ontrack.cgi?tracknbr=")
		'maryShipMethods(1) = Array("U.S.P.S.","http://trkcnfrm1.smi.usps.com/netdata-cgi/db2www/cbd_243.d2w/output?CAMEFROM=OK&strOrigTrackNum=")
		'Updated 3 Sep 05
		maryShipMethods(1) = Array("U.S.P.S.","http://trkcnfrm1.smi.usps.com/PTSInternetWeb/InterLabelInquiry.do?origTrackNum=")
	Else
		maryShipMethods(1) = Array("UPS","http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=1&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=")
		'maryShipMethods(0) = Array("U.S.P.S.","http://www.framed.usps.com/cgi-bin/cttgate/ontrack.cgi?tracknbr=")
		'maryShipMethods(0) = Array("U.S.P.S.","http://trkcnfrm1.smi.usps.com/netdata-cgi/db2www/cbd_243.d2w/output?CAMEFROM=OK&strOrigTrackNum=")
		'Updated 3 Sep 05
		maryShipMethods(0) = Array("U.S.P.S.","http://trkcnfrm1.smi.usps.com/PTSInternetWeb/InterLabelInquiry.do?origTrackNum=")
	End If
	'maryShipMethods(2) = Array("FedEx","http://www.fedex.com/cgi-bin/tracking?action=track&language=english&cntry_code=us&tracknumbers=")
	'Updated to new FedEx tracking link on 11/28/2006
	maryShipMethods(2) = Array("FedEx","http://fedex.com/Tracking?ascend_header=1&clienttype=dotcom&cntry_code=us&language=english&tracknumbers=")

	maryShipMethods(3) = Array("Canada Post","http://204.104.133.7/scripts/tracktrace.dll?MfcIsApiCommand=TraceE&referrer=CPCNewPop&i_num=")

	'***********************************************************************************************

	Function ShipIDToName(lngID)
		If lngID <= UBound(maryShipMethods) Then ShipIDToName = maryShipMethods(lngID)(0)
	End Function	'ShipIDToName

	'***********************************************************************************************

	Function ShipNameToID(strName)
        For i=0 to UBound(maryShipMethods)
			If Trim(strName) = Trim(maryShipMethods(i)(0)) Then
				ShipNameToID = i
				Exit Function
			End If
        Next
        ShipNameToID = -1
	End Function	'ShipNameToID

	'***********************************************************************************************

	Function ShipCodeToCarrierID(strShipMethod)
	
	Dim pbytCarrierID

		Select Case strShipMethod
			Case "1DM","1DA","1DAPI","1DP","2DM","2DA","3DS","GND","STD","XPR","XDM","XPD","Next Day Air", _
				 "UPS"
				pbytCarrierID = ShipNameToID("UPS")
			Case "Express", _
				 "Priority", _
				 "Parcel", _
				 "First Class", _
				 "Global Express Guaranteed (GXG) Document Service", _
				 "Global Express Guaranteed (GXG) Non-Document Service", _
				 "Global Express Mail (EMS)", _
				 "Global Priority Mail - Flat-rate Envelope (large):", _
				 "Global Priority Mail - Flat-rate Envelope (small)", _
				 "Global Priority Mail - Variable Weight Envelope (single)", _
				 "Airmail Letter Post", _
				 "Airmail Parcel Post", _
				 "Economy (Surface) Letter Post", _
				 "Economy (Surface) Parcel Post", _
				 "USPS", "U.S.P.S."
				pbytCarrierID = ShipNameToID("U.S.P.S.")
			Case "U.S. Domestic FedEx Home Delivery Package", "FedEx"
				pbytCarrierID = ShipNameToID("FedEx")
			Case "Canada Post"
				pbytCarrierID = ShipNameToID("Canada Post")
			Case Else
				pbytCarrierID = -1
		End Select
		
		ShipCodeToCarrierID = pbytCarrierID
		
	End Function	'ShipCodeToCarrierID

	'***********************************************************************************************

	Function ShipMethodsAsOptions(strValue)

	Dim i
	Dim pstrTemp

		pstrTemp = "<option value=''></option>"
		For i=0 to UBound(maryShipMethods)
			If strValue = i Then
				pstrTemp = pstrTemp & "<option value='" & i & "' selected>" & maryShipMethods(i)(0) & "</option>"
			Else
				pstrTemp = pstrTemp & "<option value='" & i & "'>" & maryShipMethods(i)(0) & "</option>"
			End If
		Next
		
		ShipMethodsAsOptions = pstrTemp

	End Function	'ShipMethodsAsOptions

'***********************************************************************************************

	Function TrackingLink(byVal vntCarrierID, byVal strTrackingNumber, byVal strFormName)

	Dim p_TrackingURL
	Dim p_bytCarrierID
	Dim pstrOut
	
		
		If isNumeric(vntCarrierID) Then
			p_bytCarrierID = vntCarrierID
		Else
			p_bytCarrierID = ShipNameToID(vntCarrierID)
		End If
		If Len(strTrackingNumber & "") > 0 Then
			If p_bytCarrierID > -1 Then
				Select Case p_bytCarrierID
					Case 4:	'DHL
						pstrOut = "<form name='" & strFormName & "' id='" & strFormName & "' action='http://track.dhl-usa.com/TrackByNbr.asp' method='post' target='DHLTracking'>" _
								& "<input type='hidden' name='hdnTrackMode' value='nbr'>" _
								& "<input type='hidden' name='hdnPostType' value='init'>" _
								& "<input type='hidden' name='hdnRefPage' value='0'>" _
								& "<input type='hidden' name='txtTrackNbrs' value='" & Server.HTMLEncode(Replace(strTrackingNumber, ",", vbcrlf)) & "'>" _
								& "</form><a href='' onclick='document." & strFormName & ".submit(); return false;'>" & Replace(strTrackingNumber, ",", ", ") & "</a>" _
								& ""
					Case Else
						p_TrackingURL = maryShipMethods(p_bytCarrierID)(1) & strTrackingNumber
						pstrOut = ShipIDToName(p_bytCarrierID) & " <a href=" & Chr(34) & p_TrackingURL & Chr(34) & " target=_blank>" & strTrackingNumber & "</a>"
				End Select
			End If
		End If

		TrackingLink = pstrOut
		
	End Function	'TrackingLink
%>
