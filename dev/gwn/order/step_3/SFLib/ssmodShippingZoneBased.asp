<%
'********************************************************************************
'*   Zone Based Shipping					                                    *
'*   Release Version:   2.00.002												*
'*   Release Date:		September 5, 2003										*
'*   Revision Date:		May 22, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*																				*
'*   Release 2.00.002 (May 22, 2004)											*
'*	   - Reviewed code for SQL Injection routines								*
'*	   - Added option to determine shipCost										*
'*																				*
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'****************************************************************************************************************

Class clsZoneBasedShipping

Private pobjConn

Private plngZoneID
Private pbytZoneType
Private paryAvailableShippingMethods


	Public Property Let Connection(objConn)
		Set pobjConn = objConn
	End Property

	Public Property Let ZoneType(bytZoneType)
		pbytZoneType = bytZoneType
	End Property

	Public Property Let ZoneID(lngZoneID)
		plngZoneID = lngZoneID
	End Property
	Public Property Get ZoneID
		ZoneID = plngZoneID
	End Property

	Public Property Get availableRates
		availableRates = paryAvailableShippingMethods
	End Property
	
	Public Function SetZone(strZone)

	dim sql
	dim p_rs

	'On Error Resume Next

		If len(strZone) = 0 Then Exit Function
		
		Select Case pbytZoneType
			Case 0: 'Country
				sql = "Select ZoneID from ssShipZones where ZoneCountries like '%;" & Replace(strZone, "'", "''") & ";%'"
				If cblnDebugZBS Then Response.Write "Searching for Zone by Country: " & strZone & "<br />"
			Case 1: 'State
				sql = "Select ZoneID from ssShipZones where ZoneStates like '%;" & Replace(strZone, "'", "''") & ";%'"
				If cblnDebugZBS Then Response.Write "Searching for Zone by State: " & strZone & "<br />"
			Case 2: 'ZIP
				sql = "Select ZoneID from ssShipZones where ZoneZips like '%;" & Replace(strZone, "'", "''") & ";%'"
				If cblnDebugZBS Then Response.Write "Searching for Zone by ZIP: " & strZone & "<br />"
		End Select
		
		Set p_rs = CreateObject("ADODB.RECORDSET")
		with p_rs
			.ActiveConnection = pobjConn
		    .CursorLocation = 2 'adUseClient
		    .CursorType = 3 'adOpenStatic
		    .LockType = 1 'adLockReadOnly
			.Source = sql
			.Open
				
			If p_rs.EOF Then
				plngZoneID = "FAIL"
				If cblnDebugZBS Then Response.Write "<font color=red>Could not find a zone corresponding to <b>" & strZone & "</b></font><br />"
			Else
				plngZoneID = p_rs("ZoneID")
				If cblnDebugZBS Then Response.Write "Found Zone for " & strZone & ": " & plngZoneID & "<br />"
			End If
			.Close
		End With
		Set p_rs = Nothing
		SetZone = plngZoneID

	End Function	'SetZone

	'***********************************************************************************************

	Function GetRate(byRef lngShipMethod, byVal dblShipWeight)

	dim sql
	dim p_rs
	Dim p_blnShipRatePercentage
	Dim pdblRate

	'On Error Resume Next

		If cblnDebugZBS Then 
			Response.Write "Searching for Rate . . .<br />"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;ZoneID: " & plngZoneID & "<br />"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;Weight: " & dblShipWeight & "<br />"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;Ship Method ID: " & lngShipMethod & "<br />"
		End If
		
		If len(plngZoneID) = 0 Or Not isNumeric(plngZoneID) Then
			GetRate = "FAIL"
			Exit Function
		End If
		
		If len(dblShipWeight) = 0 Or Not isNumeric(dblShipWeight) Then
			GetRate = "FAIL"
			Exit Function
		End If

		If len(lngShipMethod) = 0 Then
			sql = "Select Distinct ShipMethod, ShipRate, ShipRatePercentage from ssShippingRates where " _
				& " ShipZone=" & plngZoneID _
				& " And ShipWeight>" & dblShipWeight _
				& " Order By ShipRate Asc"
			Set p_rs = CreateObject("ADODB.RECORDSET")
			with p_rs
				.ActiveConnection = pobjConn
			    .CursorLocation = 2 'adUseClient
			    .CursorType = 3 'adOpenStatic
			    .LockType = 1 'adLockReadOnly
				.Source = sql
				.Open
				
				If p_rs.EOF Then
					GetRate = "FAIL"
				Else
					'GetRate = p_rs("ShipRate").value
					'GetRate = p_rs.GetRows()
					'GetRate = p_rs.GetString(,,",",",")

					If Len(p_rs.Fields("ShipRatePercentage").Value & "") = 0 Then
						p_blnShipRatePercentage = False
					ElseIf CBool(p_rs.Fields("ShipRatePercentage").Value) Then
						p_blnShipRatePercentage = True
					Else
						p_blnShipRatePercentage = False
					End If
					
					If p_blnShipRatePercentage Then
						pdblRate = Round(CDbl(p_rs.Fields("ShipRate").Value) * CDbl(Trim(dblShipWeight)) / 100,2)
					Else	
						pdblRate = p_rs.Fields("ShipRate").Value
					End If

					lngShipMethod = p_rs.Fields("ShipMethod").Value
					
					GetRate = CDbl(pdblRate)
					
					If cblnDebugZBS Then 
						Response.Write "Rate Found: " & CDbl(pdblRate) & "<br />"
						Response.Write "ShipRate: " & p_rs.Fields("ShipRate").Value & "<br />"
						Response.Write "ShipMethod: " & p_rs.Fields("ShipMethod").Value & "<br />"
						Response.Write "ShipRatePercentage: " & p_rs.Fields("ShipRatePercentage").Value & "<br />"
						Response.Write "p_blnShipRatePercentage: " & p_blnShipRatePercentage & "<br />"
						Response.Write "dblShipWeight: " & dblShipWeight & "<br />"
					End If

				End If
				.Close
			End With
			Set p_rs = Nothing
		Else
			If Not isNumeric(lngShipMethod) Then
				GetRate = "FAIL"
				Exit Function
			End If
		
			sql = "Select ShipRate, ShipRatePercentage from ssShippingRates where " _
				& "ShipMethod=" & lngShipMethod _
				& " And ShipZone=" & plngZoneID _
				& " And ShipWeight>" & dblShipWeight _
				& " Order By ShipWeight"

			Set p_rs = CreateObject("ADODB.RECORDSET")
			with p_rs
				.MaxRecords = 1
				.ActiveConnection = pobjConn
			    .CursorLocation = 2 'adUseClient
			    .CursorType = 3 'adOpenStatic
			    .LockType = 1 'adLockReadOnly
				.Source = sql
				.Open
				
				If p_rs.EOF Then
					GetRate = "FAIL"
					If cblnDebugZBS Then 
						Response.Write "<font color=red>Could not find a rate corresponding to the above information</font><br />"
						Response.Write "sql: " & sql & "<br />"
					End If
				Else
					If Len(p_rs.Fields("ShipRatePercentage").Value & "") = 0 Then
						p_blnShipRatePercentage = False
					ElseIf CBool(p_rs.Fields("ShipRatePercentage").Value) Then
						p_blnShipRatePercentage = True
					Else
						p_blnShipRatePercentage = False
					End If
					
					If p_blnShipRatePercentage Then
						pdblRate = Round(CDbl(p_rs.Fields("ShipRate").Value) * CDbl(Trim(dblShipWeight)) / 100,2)
					Else	
						pdblRate = p_rs.Fields("ShipRate").Value
					End If
					
					GetRate = CDbl(pdblRate)
					
					If cblnDebugZBS Then 
						Response.Write "Rate Found: " & CDbl(pdblRate) & "<br />"
						Response.Write "ShipRate: " & p_rs.Fields("ShipRate").Value & "<br />"
						Response.Write "ShipRatePercentage: " & p_rs.Fields("ShipRatePercentage").Value & "<br />"
						Response.Write "p_blnShipRatePercentage: " & p_blnShipRatePercentage & "<br />"
						Response.Write "dblShipWeight: " & dblShipWeight & "<br />"
					End If
				End If
				.Close
			End With
			Set p_rs = Nothing
		End If

	End Function	'GetRate
	
	'***********************************************************************************************

	Function GetAnyAvailableShipping(byRef lngShipMethod, byVal dblShipWeight)

	dim sql
	dim p_rs
	Dim p_blnShipRatePercentage
	Dim pdblRate
	Dim pdblRateOut
	Dim pstrCodeOut
	Dim plngCounter
	Dim pstrPrevCode
	Dim pblnFound
	
	'On Error Resume Next

		If cblnDebugZBS Then 
			Response.Write "Searching for Any available rate . . .<br />"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;ZoneID: " & plngZoneID & "<br />"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;Weight: " & dblShipWeight & "<br />"
		End If
		
		If len(plngZoneID) = 0 Or Not isNumeric(plngZoneID) Then
			GetAnyAvailableShipping = "FAIL"
			Exit Function
		ElseIf len(dblShipWeight) = 0 Or Not isNumeric(dblShipWeight) Then
			GetAnyAvailableShipping = "FAIL"
			Exit Function
		End If

		sql = "Select Distinct ShipMethod, ShipRate, ShipRatePercentage from ssShippingRates where " _
			& " ShipZone=" & plngZoneID _
			& " And ShipWeight>" & dblShipWeight _
			& " Order By ShipRate Asc"
			
		sql = "SELECT  ssShippingRates.ShipMethod, ssShippingRates.ShipRate, ssShippingRates.ShipRatePercentage, sfShipping.shipMethod As shipMethodName" _
			& " FROM sfShipping INNER JOIN ssShippingRates ON sfShipping.shipID = ssShippingRates.ShipMethod " _
			& " WHERE ((ssShippingRates.ShipWeight>" & dblShipWeight & ") AND (ssShippingRates.ShipZone=" & plngZoneID & ")) " _
			& " ORDER BY sfShipping.shipMethod, ssShippingRates.ShipWeight, ssShippingRates.ShipRate"
			
		Set p_rs = CreateObject("ADODB.RECORDSET")
		with p_rs
			.ActiveConnection = pobjConn
			.CursorLocation = 2 'adUseClient
			.CursorType = 3 'adOpenStatic
			.LockType = 1 'adLockReadOnly
			.Source = sql
			.Open
			
			If p_rs.EOF Then
				GetAnyAvailableShipping = "FAIL"
			Else
				plngCounter = -1
				pdblRateOut = 99999999999	'need a big number, code below will find smallest
				pblnFound = False

				ReDim paryAvailableShippingMethods(0)
				Do While Not .EOF
					If pstrPrevCode <> .Fields("ShipMethod").Value Then
						pstrPrevCode = .Fields("ShipMethod").Value
						plngCounter = plngCounter + 1
						ReDim Preserve paryAvailableShippingMethods(plngCounter)
				
						If Len(.Fields("ShipRatePercentage").Value & "") = 0 Then
							p_blnShipRatePercentage = False
						ElseIf CBool(.Fields("ShipRatePercentage").Value) Then
							p_blnShipRatePercentage = True
						Else
							p_blnShipRatePercentage = False
						End If
						
						If p_blnShipRatePercentage Then
							pdblRate = Round(CDbl(.Fields("ShipRate").Value) * CDbl(Trim(dblShipWeight)) / 100,2)
						Else	
							pdblRate = Round(.Fields("ShipRate").Value, 2)
						End If

						If pdblRate <= pdblRateOut And Not pblnFound Then
							pdblRateOut = pdblRate
							pstrCodeOut = Trim(.Fields("ShipMethod").Value)
						End If
						
						If lngShipMethod = Trim(.Fields("ShipMethod").Value) Then
							pdblRateOut = pdblRate
							pstrCodeOut = lngShipMethod
							pblnFound = True
						End If
					
						paryAvailableShippingMethods(plngCounter) = Array(Trim(.Fields("ShipMethod").Value), Trim(.Fields("shipMethodName").Value), pdblRate)
						
					End If
					.MoveNext
				Loop

				'GetRate = p_rs("ShipRate").value
				'GetRate = p_rs.GetRows()
				'GetRate = p_rs.GetString(,,",",",")

				lngShipMethod = pstrCodeOut
				GetAnyAvailableShipping = CDbl(pdblRateOut)
				
				If cblnDebugZBS Then 
					Response.Write "ShipRate: " & pdblRate & "<br />"
					Response.Write "p_blnShipRatePercentage: " & p_blnShipRatePercentage & "<br />"
					Response.Write "dblShipWeight: " & dblShipWeight & "<br />"
					Response.Write "Rate Found: " & CDbl(pdblRate) & "<br />"
				End If

			End If
			.Close
		End With
		Set p_rs = Nothing

	End Function	'GetAnyAvailableShipping
	
End Class	'clsZoneBasedShipping

	'*******************************************************************************************************************************

	Function OrderWeight()

	Dim pstrSQL
	Dim pobjRS
	Dim plngSessionID
	
		plngSessionID = Session("SessionID")
		If Len(plngSessionID) = 0 Or Not isNumeric(plngSessionID) Then
			OrderWeight = 0
			Exit Function
		End If
	
		pstrSQL = "SELECT Sum([odrdttmpQuantity]*[prodWeight]) AS orderWeight" _
				& " FROM sfProducts INNER JOIN sfTmpOrderDetails ON sfProducts.prodID = sfTmpOrderDetails.odrdttmpProductID" _
				& " GROUP BY sfProducts.prodShipIsActive, sfTmpOrderDetails.odrdttmpSessionID" _
				& " HAVING (((sfProducts.prodShipIsActive)=1) AND ((sfTmpOrderDetails.odrdttmpSessionID)=" & plngSessionID & "))"

		Set pobjRS = CreateObject("ADODB.RecordSet")
		pobjRS.Open pstrSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If pobjRS.EOF Then
			OrderWeight = 0
		Else
			OrderWeight = pobjRS.Fields("orderWeight").Value
		End If
		pobjRS.Close
		Set pobjRS = Nothing

	End Function	'OrderWeight

	'*******************************************************************************************************************************

	Function OrderShipCost()

	Dim pstrSQL
	Dim pobjRS
	Dim plngSessionID
	
		plngSessionID = Session("SessionID")
		If Len(plngSessionID) = 0 Or Not isNumeric(plngSessionID) Then
			OrderShipCost = 0
			Exit Function
		End If
	
		pstrSQL = "SELECT Sum([odrdttmpQuantity]*[prodShip]) AS orderWeight" _
				& " FROM sfProducts INNER JOIN sfTmpOrderDetails ON sfProducts.prodID = sfTmpOrderDetails.odrdttmpProductID" _
				& " GROUP BY sfProducts.prodShipIsActive, sfTmpOrderDetails.odrdttmpSessionID" _
				& " HAVING (((sfProducts.prodShipIsActive)=1) AND ((sfTmpOrderDetails.odrdttmpSessionID)=" & plngSessionID & "))"

		Set pobjRS = CreateObject("ADODB.RecordSet")
		pobjRS.Open pstrSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If pobjRS.EOF Then
			OrderShipCost = 0
		Else
			OrderShipCost = pobjRS.Fields("orderWeight").Value
		End If
		pobjRS.Close
		Set pobjRS = Nothing

	End Function	'OrderShipCost

	'*******************************************************************************************************************************

	Function OrderCount()

	Dim pstrSQL
	Dim pobjRS
	Dim plngSessionID
	
		plngSessionID = Session("SessionID")
		If Len(plngSessionID) = 0 Or Not isNumeric(plngSessionID) Then
			OrderWeight = 0
			Exit Function
		End If
	
		pstrSQL = "SELECT Sum([odrdttmpQuantity]) AS orderCount" _
				& " FROM sfProducts INNER JOIN sfTmpOrderDetails ON sfProducts.prodID = sfTmpOrderDetails.odrdttmpProductID" _
				& " GROUP BY sfProducts.prodShipIsActive, sfTmpOrderDetails.odrdttmpSessionID" _
				& " HAVING (((sfProducts.prodShipIsActive)=1) AND ((sfTmpOrderDetails.odrdttmpSessionID)=" & plngSessionID & "))"

		Set pobjRS = CreateObject("ADODB.RecordSet")
		pobjRS.Open pstrSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If pobjRS.EOF Then
			OrderCount = 0
		Else
			OrderCount = pobjRS.Fields("orderCount").Value
		End If
		pobjRS.Close
		Set pobjRS = Nothing

	End Function	'OrderCount

	'*******************************************************************************************************************************

Dim cblnDebugZBS
cblnDebugZBS = CBool(Len(Session("ssDebug_ZBS")) > 0)

Const cblnUseZoneBasedShipping = True
%>