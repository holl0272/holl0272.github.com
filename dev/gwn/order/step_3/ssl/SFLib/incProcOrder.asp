<%

'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.3

'@FILENAME: incProcOrder.asp
	 


'@DESCRIPTION: Process the customers order

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws  and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

'Modified 10/23/01 
'Storefront Ref#'s: 147 'JF

'****************************************************************************************************************

Dim maryPaymentTypes

'****************************************************************************************************************

Function getPaymentList(byVal strSelected)

Dim pstrTemp
Dim i

	If loadPaymentTypes Then
		For i = 0 To UBound(maryPaymentTypes, 2)
			If maryPaymentTypes(0, i) = strSelected Then
				pstrTemp = pstrTemp & "<option value=""" & maryPaymentTypes(0, i) & """ selected>" & maryPaymentTypes(1, i) & "</option>" & vbcrlf
			Else
				pstrTemp = pstrTemp & "<option value=""" & maryPaymentTypes(0, i) & """>" & maryPaymentTypes(1, i) & "</option>" & vbcrlf
			End If
		Next
	End If
	
	getPaymentList = pstrTemp

End Function	'getPaymentList

'****************************************************************************************************************

Function getPaymentList_radio(byVal strSelected)

Dim pstrTemp
Dim i

	If loadPaymentTypes Then
		If UBound(maryPaymentTypes, 2) = 0 Then
			pstrTemp = pstrTemp & "<input type=hidden name=PaymentMethod id=PaymentMethod value=""" & maryPaymentTypes(0, i) & """>" & maryPaymentTypes(1, i) & vbcrlf
		Else
			For i = 0 To UBound(maryPaymentTypes, 2)
				If i > 0 Then pstrTemp = pstrTemp & "<br />"
				If maryPaymentTypes(0, i) = strSelected Then
					pstrTemp = pstrTemp & "<input type=radio name=PaymentMethod id=PaymentMethod" & i & " value=""" & maryPaymentTypes(0, i) & """ checked>&nbsp;<label for=PaymentMethod" & i & ">" & maryPaymentTypes(1, i) & "</label>" & vbcrlf
				Else
					pstrTemp = pstrTemp & "<input type=radio name=PaymentMethod id=PaymentMethod" & i & " value=""" & maryPaymentTypes(0, i) & """>&nbsp;<label for=PaymentMethod" & i & ">" & maryPaymentTypes(1, i) & "</label>" & vbcrlf
				End If
			Next
		End If
	End If
	
	getPaymentList_radio = pstrTemp

End Function	'getPaymentList_radio

'****************************************************************************************************************

Function loadPaymentTypes()

Dim pblnResult
Dim pobjRS
Dim pstrSQL
Dim pstrTemp
Dim i

	pblnResult = False
	
	Application.Contents.Remove "PaymentTypesArray"
	maryPaymentTypes = Application("PaymentTypesArray")
	If isArray(maryPaymentTypes) Then
		pblnResult = True
	Else
		pstrSQL = "SELECT DISTINCT transtype, transtype as transText FROM sfTransactionTypes WHERE transIsActive = 1 ORDER BY transtype"

		Set pobjRS = CreateObject("ADODB.RECORDSET")
		With pobjRS
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			
			On Error Resume Next
			If Err.number = 0 Then
				maryPaymentTypes = .GetRows()
				For i = 0 To UBound(maryPaymentTypes, 2)
					Select Case Trim(maryPaymentTypes(0, i) & "")
						Case cstrECheckTerm:
							maryPaymentTypes(0, i) = Trim(maryPaymentTypes(0, i) & "")
							maryPaymentTypes(1, i) = "Check"
						Case cstrPOTerm:
							maryPaymentTypes(0, i) = Trim(maryPaymentTypes(0, i) & "")
							maryPaymentTypes(1, i) = "Purchase Order"
						Case cstrPhoneFaxTerm:
							maryPaymentTypes(0, i) = Trim(maryPaymentTypes(0, i) & "")
							maryPaymentTypes(1, i) = "Phone/Fax"
						Case Else:
							maryPaymentTypes(0, i) = Trim(maryPaymentTypes(0, i) & "")
							maryPaymentTypes(1, i) = Trim(maryPaymentTypes(1, i) & "")
					End Select
				Next
				pblnResult = True
				Application("PaymentTypesArray") = maryPaymentTypes
			Else
				Err.Clear
			End If
			.Close
			On Error Goto 0
			
		End With
		Set pobjRS = Nothing
	End If
	
	loadPaymentTypes = pblnResult
	
End Function	'loadPaymentTypes

'****************************************************************************************************************

'-------------------------------------------------------------------
' Subroutine setUpdateSavedCartCustID
'-------------------------------------------------------------------
Sub setUpdateSavedCartCustID(iCustID,iDeletedCustID)
	Dim sSQL, rsTmpCust
	sSQL = "Select odrdtsvdCustID FROM sfSavedOrderDetails WHERE odrdtsvdCustID=" & makeInputSafe(iDeletedCustID)
	Set rsTmpCust = CreateObject("ADODB.RecordSet")		
		rsTmpCust.Open sSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText					
			Do While NOT rsTmpCust.EOF
					rsTmpCust.Fields("odrdtsvdCustID")	= makeInputSafe(trim(iCustID))
					rsTmpCust.Update	
					rsTmpCust.MoveNext
			Loop
		closeobj(rsTmpCust)		
End Sub

'--------------------------------------------------------------------
' Function to update sessionid
'--------------------------------------------------------------------
Sub setUpdateTmpOrdersSessionID(OldSessionID,NewSessionID)
Dim sSQL, rsTmp 
	sSQL = "SELECT odrdttmpSessionID FROM sfTmpOrderDetails WHERE odrdttmpSessionID=" & OldSessionID
		Set rsTmp = CreateObject("ADODB.RecordSet")		
		rsTmp.Open sSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText					
			Do While NOT rsTmp.EOF
					rsTmp.Fields("odrdttmpSessionID")	= makeInputSafe(trim(NewSessionID))
					rsTmp.Update	
					rsTmp.MoveNext
			Loop
		closeobj(rsTmp)		
End Sub

'-----------------------------------------------------------------------
' Deletes saved customer row
'-----------------------------------------------------------------------
Sub DeleteCustRow(iCustID)
	Dim rsDelete, sSQL
	
	sSQL = "DELETE FROM sfCustomers WHERE custID= " & makeInputSafe(iCustID) & " AND custFirstName = 'SavedCartCustomer'"
	Set rsDelete = cnn.Execute(sSQL)
	closeObj(rsDelete)
End Sub

'--------------------------------------------------------------------
' Function : getShippingList
' This returns the list for shipping options in HTML format for dropdown box.
'--------------------------------------------------------------------	
Function getShippingList(byVal vntShipCode, byVal blnFree)

Dim sShipList, rsShipList, sLocalSQL, iCounter
Dim sSql
	
'debugprint "adminShipType", adminShipType
	If adminShipType = 1 Then
		sLocalSQL = "SELECT shipID, shipMethod FROM sfShipping WHERE shipIsActive = 1"	
		
		Set rsShipList = CreateObject("ADODB.RecordSet")
		if blnFree = true then
			sSql = "SELECT * FROM sfShipping WHERE shipMethod = 'Free Shipping' "
			rsShipList.Open sSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
			sShipList = sShipList &"<option value=" & chr(34) & rsShipList.Fields("shipID")& ",400" & chr(34) & "style=""WEIGHT: bold;COLOR: red"">Free Shipping</option>" 
			rsShipList.Close 
		end if  
					
		rsShipList.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
		If rsShipList.EOF Or rsShipList.BOF Then
			sShipList = ""
		Else					
			For iCounter = 1 to rsShipList.RecordCount
				If CStr(vntShipCode & "") = Trim(rsShipList.Fields("shipID")) Then
					sShipList = sShipList & "<option value=""" & Trim(rsShipList.Fields("shipID"))& """ selected>" & Trim(rsShipList.Fields("shipMethod")) & "</option>"
				Else
					sShipList = sShipList & "<option value=""" & Trim(rsShipList.Fields("shipID"))& """>" & Trim(rsShipList.Fields("shipMethod")) & "</option>"
				End If
				rsShipList.MoveNext
			Next	
		End If
	ElseIf adminShipType = 2 Then		
		sLocalSQL = "SELECT shipID, shipMethod FROM sfShipping WHERE shipIsActive = 1"	
		
		Set rsShipList = CreateObject("ADODB.RecordSet")
		if blnFree = true then
'			
			    sSql = "SELECT * FROM sfShipping WHERE shipMethod = 'Free Shipping' "
			    rsShipList.Open sSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText

			sShipList = sShipList &"<option value=" & chr(34) & rsShipList.Fields("shipID")& ",400" & chr(34) & "style=""WEIGHT: bold;COLOR: red"">Free Shipping</option>" 
			rsShipList.Close 
		end if  
		rsShipList.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
					
		If rsShipList.EOF Or rsShipList.BOF Then
			sShipList = ""
		Else					
			For iCounter = 1 to rsShipList.RecordCount
				sShipList = sShipList & "<option value=""" & Trim(rsShipList.Fields("shipID"))& """>" & Trim(rsShipList.Fields("shipMethod")) & "</option>"
				rsShipList.MoveNext
			Next	
		End If
				
	ElseIf adminShipType = 3 Then
			if blnFree = true then
				sShipList = "<option value=""3,400"" style=""WEIGHT: bold;COLOR: red"">Free Shipping</option>"
			Else
				sShipList =  "<option value=""0"">Regular Shipping</option>"
			End If
			
			If adminShipType = 1 Then
					sShipList = sShipList & "<option value=""1"">Premium Shipping</option>"
			End If
			
	End If

	closeobj(rsShipList)
	getShippingList = sShipList
	
End Function

%>








