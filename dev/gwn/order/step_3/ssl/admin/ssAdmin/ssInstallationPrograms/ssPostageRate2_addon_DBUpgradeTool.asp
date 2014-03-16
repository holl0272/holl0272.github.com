<%Option Explicit 
'********************************************************************************
'*   Postage Rate Administration						                        *
'*   Release Version: 2.0			                                            *
'*   Release Date: September 15, 2002											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

sub debugprint(sField1,sField2)
Response.Write "<H3>" & sField1 & ": " & sField2 & "</H3><BR>"
end sub

Function stepCounter()

	mlngCounter = mlngCounter + 1
	stepCounter = mlngCounter
	
End Function	'stepCounter

dim conn,strConn,strDBpath,sql
dim mstrFilePath
dim mrsTest
dim mblnUpgraded
dim strTableName
dim mstrAction
dim mblnError
Dim mstrMessage
dim mblnValidConnection
dim mblnSQLServer
dim mlngCounter
Dim i

'Set the Carriers
Dim paryShippingCarriers(8)
paryShippingCarriers(1) = Array ("Unknown", "RateURL", "Username", "Password", "TrackingURL", "Imagepath")
paryShippingCarriers(2) = Array ("U.S.P.S.","http://Production.ShippingApis.com/ShippingAPI.dll?API=","Username","Password","http://www.framed.usps.com/cgi-bin/cttgate/ontrack.cgi?tracknbr=[TrackingNumber]","Imagepath")
paryShippingCarriers(3) = Array ("UPS","http://www.ups.com/using/services/rave/qcost_dss.cgi", "Username", "Password", "http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=1&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=[TrackingNumber]","Imagepath")
paryShippingCarriers(4) = Array ("FedEx","http://grd.fedex.com/cgi-bin/rrr2010.exe","Username","Password","http://www.fedex.com/cgi-bin/tracking?action=track&language=english&cntry_code=us&tracknumbers=[TrackingNumber]","Imagepath")
paryShippingCarriers(5) = Array ("Canada Post","http://206.191.4.228:30000","Username","Password","http://204.104.133.7/scripts/tracktrace.dll?MfcIsApiCommand=TraceE&referrer=CPCNewPop&i_num=[TrackingNumber]","Imagepath")
paryShippingCarriers(6) = Array ("Airborne","RateURL","Username","Password","TrackingURL","Imagepath")
paryShippingCarriers(7) = Array ("DHL","RateURL","Username","Password","TrackingURL","Imagepath")
paryShippingCarriers(8) = Array ("Freight Quote","RateURL","Username","Password","TrackingURL","Imagepath")

'Set the Methods
Dim paryShippingMethods(60)
'paryShippingMethods(0) = Array ("ssShippingCarrierID", "ssShippingMethodName", "ssShippingMethodCode", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

'U.S.P.S. Methods
'domestic
mlngCounter = 0
paryShippingMethods(stepCounter) = Array (2, "Express Mail", "Express", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Priority Mail", "Priority", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Parcel Post", "Parcel", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "First Class", "First Class", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

'international
paryShippingMethods(stepCounter) = Array (2, "Global Express Guaranteed Document Service", "Global Express Guaranteed Document Service", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Global Express Guaranteed Non-Document Service", "Global Express Guaranteed Non-Document Service", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Global Express Mail (EMS)", "Global Express Mail (EMS)", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Global Priority Mail - Flat-rate Envelope (large)", "Global Priority Mail - Flat-rate Envelope (large)", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Global Priority Mail - Flat-rate Envelope (small)", "Global Priority Mail - Flat-rate Envelope (small)", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Global Priority Mail - Variable Weight Envelope (single)", "Global Priority Mail - Variable Weight Envelope (single)", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Airmail Letter Post", "Airmail Letter Post", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Airmail Parcel Post", "Airmail Parcel Post", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Economy (Surface) Letter Post", "Economy (Surface) Letter Post", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (2, "Economy (Surface) Parcel Post", "Economy (Surface) Parcel Post", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

'UPS Methods
paryShippingMethods(stepCounter) = Array (3, "UPS Next Day AM", "1DM", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS Next Day Air", "1DA", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS Next Day Air Saver", "1DP", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS 2nd Day Air AM", "2DM", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS 2nd Day Air", "2DA", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS 3 Day Select", "3DS", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS Standard Ground", "GND", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS Canada Standard", "STD", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS WorldWide Express", "XPR", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS WorldWide Express Plus", "XDM", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (3, "UPS WorldWide Expedited", "XPD", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

'FedEx Methods
paryShippingMethods(stepCounter) = Array (4, "FedEx Priority", "01", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx 2day", "03", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx Standard Overnight", "5", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx First Overnight", "06", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx Express Saver", "20", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx Overnight Freight", "70", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx 2day Freight", "80", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx Express Saver Freight", "83", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx International Priority", "01i", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx International Economy", "03i", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx International First", "06i", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx International Priority Freight", "70i", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx International Economy Freight", "86i", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "FedEx Home Delivery", "90", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "U.S. Domestic FedEx Ground Package", "92", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (4, "International FedEx Ground Package", "92i", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

'Canada Post Methods
paryShippingMethods(stepCounter) = Array (5, "Domestic - Regular", "1010", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "Domestic - Expedited", "1020", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "Domestic - Xpresspost", "1030", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "Domestic - Priority Courier", "1040", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "Domestic - Expedited Evening", "1120", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "Domestic - Xpresspost Evening", "1130", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "Domestic - Expedited Saturday", "1220", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "Domestic - Xpresspost Saturday", "1230", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "USA - Surface", "2010", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "USA - Air", "2020", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "USA - Xpresspost", "2030", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "USA - Purolator", "2040", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "USA - Puropak", "2050", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "International - Surface", "3010", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "International - Air", "3020", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "International - Purolator", "3040", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
paryShippingMethods(stepCounter) = Array (5, "International - Puropak", "3050", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)

'Freight Quote
paryShippingMethods(stepCounter) = Array (8, "Freight Quote", "FreightQuote", 1, 1, 1, 0, 0, 999999, 0, 1, 0, 0)
'debugprint "stepCounter",stepCounter

'For i = 1 To mlngCounter
'	Response.Write i & ": " & paryShippingMethods(i)(1) & "<BR>"
'Next 'i

On Error Resume Next

	mblnError = False
	mstrAction = Request.Form("Action")

	if len(mstrAction) > 0 then

		Set conn = Server.CreateObject ("ADODB.Connection")
		conn.Open session("DSN_NAME")
		mstrFilePath = Request.Form("FilePath")
		mblnSQLServer = (lCase(Request.Form("SQLServer")) = "on")

	End If

	'Test Connection to the database
	If len(session("DSN_NAME")) > 0 then
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open session("DSN_NAME")
		If (conn.State = 1) Then
			mblnValidConnection = True
		Else
			mblnValidConnection = False
			mstrMessage = "<H3><Font Color='Red'>Could not connect to the database. Error: " & Err.number & " - " & Err.Description & "</FONT></H3>"
			Err.Clear
		End If
	Else
		mblnValidConnection = False
		mstrMessage = "<H3><Font Color='Red'>Could not connect to the database</FONT></H3>"
	End If

	'Test to see if DB has already been upgraded
	If mblnValidConnection Then
		sql = "Select ssShippingCarrierID from ssShippingCarriers"
		Set mrsTest = conn.Execute(sql)

		mblnUpgraded = (Err.number=0)
		mrsTest.Close
		Set mrsTest = Nothing
		Err.Clear
		
	Else
		mblnUpgraded = False
	End If

	If (mstrAction = "Install Upgrade") then

		'------------------------------------------------------------------------------'
		'UPGRADE sfOrders TABLE														   '
		'------------------------------------------------------------------------------'


		SQL = "ALTER TABLE sfOrders ALTER COLUMN orderShipMethod Char (65)"
		conn.Execute(SQL)

		if Err.number = 0 then
			mstrMessage = mstrMessage & "<H3><B>sfOrders Table Successfully Upgraded</B></H3><BR>"
		else
			mblnError = True
			mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error upgrading sfOrders Table: " & Err.description	& "</FONT></H3><BR>"	
		end if

		'------------------------------------------------------------------------------'
		' ADD ssShippingCarriers TABLE														   '
		'------------------------------------------------------------------------------'

		strTableName = "ssShippingCarriers"
		If mblnSQLServer Then
			SQL = "CREATE TABLE " & strTableName & " " _
					& "(ssShippingCarrierID int Identity PRIMARY KEY," _
					& " ssShippingCarrierName char (65)," _
					& " ssShippingCarrierUserName char (65)," _
					& " ssShippingCarrierPassword char (65)," _
					& " ssShippingCarrierRateURL char (255)," _
					& " ssShippingCarrierTrackingURL char (255)," _
					& " ssShippingCarrierImagePath char (255)" _
					& " )"
		Else
			SQL = "CREATE TABLE " & strTableName & " " _
					& "(ssShippingCarrierID COUNTER PRIMARY KEY," _
					& " ssShippingCarrierName char (65)," _
					& " ssShippingCarrierUserName char (65)," _
					& " ssShippingCarrierPassword char (65)," _
					& " ssShippingCarrierRateURL char (255)," _
					& " ssShippingCarrierTrackingURL char (255)," _
					& " ssShippingCarrierImagePath char (255)" _
					& " )"
		End If

		conn.Execute SQL,, 128

		if Err.number = 0 then
			For i = 1 To UBound(paryShippingCarriers)
				SQL = "Insert Into " & strTableName & " (ssShippingCarrierID, ssShippingCarrierName, ssShippingCarrierUserName, ssShippingCarrierPassword, ssShippingCarrierRateURL, ssShippingCarrierTrackingURL, ssShippingCarrierImagePath)" _
					& " Values (" & i & ", '" & paryShippingCarriers(i)(0) & "', '" & paryShippingCarriers(i)(2) & "', '" & paryShippingCarriers(i)(3) & "', '" & paryShippingCarriers(i)(1) & "', '" & paryShippingCarriers(i)(4) & "', '" & paryShippingCarriers(i)(5) & "')"
				conn.Execute SQL,, 128

			Next	'i
			conn.Execute "Update ssShippingCarriers Set ssShippingCarrierUserName='229830988', ssShippingCarrierPassword='1019067' Where ssShippingCarrierID=4",, 128
			conn.Execute "Update ssShippingCarriers Set ssShippingCarrierUserName='chuck@northamericanpackaging.com', ssShippingCarrierPassword='931820' Where ssShippingCarrierID=8",, 128
		
			mblnUpgraded = True
			mstrMessage = mstrMessage & "<H3><B>" & strTableName & " table successfully added.</B></H3><BR>"
		else
			mblnError = True
			mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error adding " & strTableName & ": " & Err.description	& "</FONT></H3><BR>"	
			Err.Clear
		end if

		'------------------------------------------------------------------------------'
		' ADD ssShippingMethods TABLE														   '
		'------------------------------------------------------------------------------'

		strTableName = "ssShippingMethods"
		If mblnSQLServer Then
			SQL = "CREATE TABLE " & strTableName & " " _
					& "(ssShippingMethodID int Identity PRIMARY KEY," _
					& " ssShippingCarrierID int," _
					& " ssShippingMethodCode char (65)," _
					& " ssShippingMethodName char (65)," _
					& " ssShippingMethodEnabled bit," _
					& " ssShippingMethodLocked bit," _
					& " ssShippingMethodMinCharge Decimal," _
					& " ssShippingMethodMultiple Decimal," _
					& " ssShippingMethodPerPackageFee Decimal," _
					& " ssShippingMethodPerShipmentFee Decimal," _
					& " ssShippingMethodOfferFreeShippingAbove Decimal," _
					& " ssShippingMethodLimitFreeShippingByWeight Decimal," _
					& " ssShippingMethodClass Decimal," _
					& " ssShippingMethodOrderBy Decimal," _
					& " ssShippingMethodDefault bit," _
					& " ssShippingMethodMinWeight Decimal," _
					& " ssShippingMethodPrefWeight Decimal," _
					& " ssShippingMethodMaxLength Decimal," _
					& " ssShippingMethodMaxWidth Decimal," _
					& " ssShippingMethodMaxHeight Decimal," _
					& " ssShippingMethodMaxWeight Decimal," _
					& " ssShippingMethodMaxGirth Decimal" _
					& " )"
		Else
			SQL = "CREATE TABLE " & strTableName & " " _
					& "(ssShippingMethodID COUNTER PRIMARY KEY," _
					& " ssShippingCarrierID long," _
					& " ssShippingMethodCode char (65)," _
					& " ssShippingMethodName char (65)," _
					& " ssShippingMethodEnabled YESNO," _
					& " ssShippingMethodLocked YESNO," _
					& " ssShippingMethodMinCharge double," _
					& " ssShippingMethodMultiple double," _
					& " ssShippingMethodPerPackageFee double," _
					& " ssShippingMethodPerShipmentFee double," _
					& " ssShippingMethodOfferFreeShippingAbove double," _
					& " ssShippingMethodLimitFreeShippingByWeight double," _
					& " ssShippingMethodClass double," _
					& " ssShippingMethodOrderBy double," _
					& " ssShippingMethodDefault YESNO," _
					& " ssShippingMethodMinWeight double," _
					& " ssShippingMethodPrefWeight double," _
					& " ssShippingMethodMaxLength double," _
					& " ssShippingMethodMaxWidth double," _
					& " ssShippingMethodMaxHeight double," _
					& " ssShippingMethodMaxWeight double," _
					& " ssShippingMethodMaxGirth double" _
					& " )"
		End If

		conn.Execute (SQL)

		if Err.number = 0 then
			mblnUpgraded = True
			mstrMessage = mstrMessage & "<H3><B>" & strTableName & " table successfully added.</B></H3><BR>"

			For i = 1 To mlngCounter
				SQL = "Insert Into " & strTableName & " (ssShippingMethodID, ssShippingCarrierID, ssShippingMethodCode, ssShippingMethodName, ssShippingMethodEnabled, ssShippingMethodLocked, ssShippingMethodMultiple, ssShippingMethodPerPackageFee, ssShippingMethodPerShipmentFee, ssShippingMethodOfferFreeShippingAbove, ssShippingMethodLimitFreeShippingByWeight, ssShippingMethodClass, ssShippingMethodOrderBy, ssShippingMethodDefault, ssShippingMethodMinCharge, ssShippingMethodMinWeight, ssShippingMethodPrefWeight, ssShippingMethodMaxLength, ssShippingMethodMaxWidth, ssShippingMethodMaxHeight, ssShippingMethodMaxWeight, ssShippingMethodMaxGirth)" _
					& " Values (" & i & ", '" & paryShippingMethods(i)(0) & "', '" & paryShippingMethods(i)(2) & "', '" & paryShippingMethods(i)(1) & "', " & paryShippingMethods(i)(3) & ", " & paryShippingMethods(i)(4) & ", " & paryShippingMethods(i)(5) & ", " & paryShippingMethods(i)(6) & ", " & paryShippingMethods(i)(7) & ", " & paryShippingMethods(i)(8) & ", " & paryShippingMethods(i)(9) & ", " & paryShippingMethods(i)(10) & ", " & paryShippingMethods(i)(11) & ", " & paryShippingMethods(i)(12) & ",0,0,50,0,0,0,170,0)"
				conn.Execute SQL,, 128
				If Err.number <> 0 Then
					Response.Write "SQL: " & SQL & "<BR>"
					Err.Clear
				End If
			Next	'i

		else
			mblnError = True
			mstrMessage = mstrMessage & "<H3><Font Color='Red'>Error adding " & strTableName & ": " & Err.description	& "</FONT></H3><BR>"	
			Err.Clear
		end if

	ElseIf (mstrAction = "Uninstall Upgrade") then

		'------------------------------------------------------------------------------'
		' REMOVE ssShippingCarriers TABLE														   '
		'------------------------------------------------------------------------------'

		strTableName = "ssShippingCarriers"
		SQL = "DROP TABLE " & strTableName
		conn.Execute (SQL)
		if Err.number = 0 then
			mblnUpgraded = False
			mstrMessage = mstrMessage & "<LI><B>Table " & strTableName & " successfully removed.</B></LI>"
		else
			mblnError = True
			mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & strTableName & ": " & Err.description	& "</FONT></LI>"	
		end if

		'------------------------------------------------------------------------------'
		' REMOVE ssShippingMethods TABLE														   '
		'------------------------------------------------------------------------------'

		strTableName = "ssShippingMethods"
		SQL = "DROP TABLE " & strTableName
		conn.Execute (SQL)
		if Err.number = 0 then
			mblnUpgraded = False
			mstrMessage = mstrMessage & "<LI><B>Table " & strTableName & " successfully removed.</B></LI>"
		else
			mblnError = True
			mstrMessage = mstrMessage & "<LI><Font Color='Red'>Error removing " & strTableName & ": " & Err.description	& "</FONT></LI>"	
		end if

	End If

	On Error Resume Next
	
	conn.Close 
	Set conn = Nothing
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<META http-equiv="Content-Type" content="text/html; charset=windows-1252">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<TITLE>Sandshot Sofware Database Upgrade Utility</TITLE>
</HEAD>
<BODY>
<H2>Sandshot Sofware Database Upgrade Utility</H2>
<H4>Welcome to the Postage Rate Component Add-on upgrade utility for Lagarde's Storefront 5.0</H4>
<P>This utility upgrades the StoreFront 5.0 database to use the Postage Rate Component.&nbsp; It accomplishes the following:</P>
<OL>
  <LI>adds the ssShippingCarriers table</LI>           
  <LI>adds the ssShippingMethods table</LI>           
</OL>
<P>Instructions for use:</P>
<OL>
  <LI>This file must be located in your active Storefront web. 
  <LI>Run it from your web browser.</LI>
</OL>
  
<P>Disclaimer: This utility is provided without warranty. While it has been 
successfully tested using the standard Storefront database, no guarantee regarding fitness for use in your application is 
made. Always make a backup of your database prior to making changes to it.</P>
<P><%= mstrMessage %></P>
<FORM action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="Post" id=form1 name=form1>
<% If mblnValidConnection Then %>
<%	  If mblnUpgraded Then %>
<P><INPUT type=submit value="Uninstall Upgrade" id=submit2 name=Action></P>
<%    End If %>
<%    If Not mblnUpgraded Then %>
<INPUT type=checkbox name="SQLSERVER" ID="Checkbox1">&nbsp;Check if this is a SQL Server database<BR>
<P><INPUT type=submit value="Install Upgrade" id=submit1 name=Action></P>
<%    End If %>
<% End If %>

</FORM>
<p><a href="../ssPostageRate_shippingMethodsAdmin.asp" title="Edit your shipping methods">Edit Shipping Methods</a></p>

</BODY></HTML>
