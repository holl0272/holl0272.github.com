<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Release Version   2.00.001	                                                *
'*   Release Date      December 16, 2003										*
'*   Revision Date     December 16, 2003										*
'*                                                                              *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Const cblnUseNewShipMethodIDs = True	'Set to False if you have a prior version of Order Manager and used UPS as item 1
Dim maryShipMethods(4)

	'array
	'0 - Name
	'1 - Tracking URL
	If cblnUseNewShipMethodIDs Then
		maryShipMethods(0) = Array("UPS","http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=1&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=")
		maryShipMethods(1) = Array("U.S.P.S.","http://trkcnfrm1.smi.usps.com/netdatacgi/db2www/cbd_243.d2w/output?CAMEFROM=OK&strOrigTrackNum=")
	Else
		maryShipMethods(1) = Array("UPS","http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=1&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=")
		maryShipMethods(0) = Array("U.S.P.S.","http://trkcnfrm1.smi.usps.com/netdatacgi/db2www/cbd_243.d2w/output?CAMEFROM=OK&strOrigTrackNum=")
	End If
	maryShipMethods(2) = Array("FedEx","http://www.fedex.com/cgi-bin/tracking?action=track&language=english&cntry_code=us&tracknumbers=")
	'maryShipMethods(3) = Array("Canada Post","http://204.104.133.7/scripts/tracktrace.dll?MfcIsApiCommand=TraceE&referrer=CPCNewPop&i_num=")	'old method
	maryShipMethods(3) = Array("Canada Post","https://obc.canadapost.ca/emo/basicPin.do?trackingCode=PIN&action=query&language=en&trackingId=")
	maryShipMethods(4) = Array("DHL","")

	'***********************************************************************************************

	Function ShipIDToName(byVal lngID)
	
		If Len(lngID) > 0 And isNumeric(lngID) Then
			If CLng(lngID) <= UBound(maryShipMethods) Then ShipIDToName = maryShipMethods(CLng(lngID))(0)
		End If
		
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

	Function TrackingLink(byVal vntCarrierID, byVal strTrackingNumber)

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
						pstrOut = "<form name='frmTrackByNbr' id='frmTrackByNbr' action='http://track.dhl-usa.com/TrackByNbr.asp' method='post' target='DHLTracking'>" _
								& "<input type='hidden' name='hdnTrackMode' value='nbr'>" _
								& "<input type='hidden' name='hdnPostType' value='init'>" _
								& "<input type='hidden' name='hdnRefPage' value='0'>" _
								& "<input type='hidden' name='txtTrackNbrs' value='" & strTrackingNumber & "'>" _
								& "<a href='' onclick='document.frmTrackByNbr.submit(); return false;'>" & strTrackingNumber & "</a>" _
								& "</form>"
					Case Else
						p_TrackingURL = maryShipMethods(p_bytCarrierID)(1) & strTrackingNumber
						pstrOut = ShipIDToName(p_bytCarrierID) & " <a href=" & Chr(34) & p_TrackingURL & Chr(34) & " target=_blank>" & strTrackingNumber & "</a>"
				End Select
			End If
		End If

		TrackingLink = pstrOut
		
	End Function	'TrackingLink
%>
