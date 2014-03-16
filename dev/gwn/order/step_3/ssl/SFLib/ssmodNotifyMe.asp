<%
'/////////////////////////////////////////////////
'/
'/  User Parameters
'/


'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************


'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'*	Begin Page Code
'**********************************************************

'**********************************************************
'**********************************************************

Function createNotifyMe(byVal notifyProdID, byVal notifyLastName, byVal notifyFirstName, byVal notifyEmail, byVal notifyType, byVal InventoryID)

Dim pobjCmd
Dim pobjRS

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		Set .ActiveConnection = cnn
		'.Parameters.Append .CreateParameter("notifyMeID", adInteger, adParamInputOutput, 4, NULL)
		
		If Len(visitorLoggedInCustomerID) > 0 Then
			.Parameters.Append .CreateParameter("notifyCustID", adInteger, adParamInput, 4, visitorLoggedInCustomerID)
		ElseIf Len(custID_cookie) > 0 Then
			.Parameters.Append .CreateParameter("notifyCustID", adInteger, adParamInput, 4, custID_cookie)
		Else
			.Parameters.Append .CreateParameter("notifyCustID", adInteger, adParamInput, 4, 0)
		End If
		.Parameters.Append .CreateParameter("notifyProdID", adVarChar, adParamInput, 50, notifyProdID)
		.Parameters.Append .CreateParameter("notifyStoreID", adInteger, adParamInput, 4, StoreID)
		.Parameters.Append .CreateParameter("notifyLastName", adVarChar, adParamInput, 50, notifyLastName)
		.Parameters.Append .CreateParameter("notifyFirstName", adVarChar, adParamInput, 50, notifyFirstName)
		.Parameters.Append .CreateParameter("notifyEmail", adVarChar, adParamInput, 100, notifyEmail)
		.Parameters.Append .CreateParameter("notifyType", adInteger, adParamInput, 4, notifyType)
		.Parameters.Append .CreateParameter("notifyInventoryID", adInteger, adParamInput, 4, InventoryID)

		.Commandtext = "Select notifyMeID From notifyMe Where notifyCustID=? And notifyProdID=? And notifyStoreID=? And notifyLastName=? And notifyFirstName=? And notifyEmail=? And notifyType=? And notifyInventoryID=?"
		Set pobjRS = .Execute
		If pobjRS.EOF Then
			.Parameters.Append .CreateParameter("notifyDateCreated", adDBTimeStamp, adParamInput, 16, Now())
			.Commandtext = "Insert Into notifyMe (notifyCustID, notifyProdID, notifyStoreID, notifyLastName, notifyFirstName, notifyEmail, notifyType, notifyInventoryID, notifyDateCreated, notifyNotifyCount) Values (?, ?, ?, ?, ?, ?, ?, ?, ?, 0)"
			.Execute , , adExecuteNoRecords
			
			'.Parameters.Delete "notifyLastName"
			'.Parameters.Delete "notifyFirstName"
			.Commandtext = "Select notifyMeID From notifyMe Where notifyCustID=? And notifyProdID=? And notifyStoreID=? And notifyLastName=? And notifyFirstName=? And notifyEmail=? And notifyType=? And notifyInventoryID=? And notifyDateCreated=?"

			Set pobjRS = .Execute
			If pobjRS.EOF Then
				createNotifyMe = -1
			Else
				createNotifyMe = pobjRS.Fields(0).Value
			End If
		Else
			createNotifyMe = pobjRS.Fields(0).Value
		End If	'pobjRS.EOF

		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Function	'createNotifyMe

'***********************************************************************************************

Function DeleteNotifyMeByID(byVal lngnotifyMeID)

Dim pobjCmd

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		Set .ActiveConnection = cnn
		.Commandtext = "Delete From notifyMe Where notifyMeID=?"
		.Parameters.Append .CreateParameter("notifyMeID", adInteger, adParamInput, 4, lngnotifyMeID)
		
		On Error Resume Next
		.Execute,,128
		If Err.number = 0 Then
			DeleteNotifyMeByID = True
		Else
			DeleteNotifyMeByID = False
			Err.Clear
		End If

	End With	'pobjCmd
	Set pobjCmd = Nothing

End Function	'DeleteNotifyMeByID

'**********************************************************

Function LoadNotifyMeByInventoryID(byVal InventoryID, byRef aryNotifyMe)

Dim pobjCmd
Dim pobjRS

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		Set .ActiveConnection = cnn
		.Commandtext = "Select notifyMeID, notifyCustID, notifyProdID, notifyStoreID, notifyLastName, notifyFirstName, notifyEmail, notifyType, notifyDateCreated, notifyDateNotified, notifyNotifyCount From notifyMe Where notifyInventoryID=?"
		.Parameters.Append .CreateParameter("notifyInventoryID", adInteger, adParamInput, 4, InventoryID)
		
		Set pobjRS = .Execute
		If Not pobjRS.EOF Then
			ReDim aryNotifyMe(10)
			
			aryNotifyMe(0) = pobjRS.Fields("notifyMeID").Value
			aryNotifyMe(1) = pobjRS.Fields("notifyCustID").Value
			aryNotifyMe(2) = pobjRS.Fields("notifyProdID").Value
			aryNotifyMe(3) = pobjRS.Fields("notifyStoreID").Value
			aryNotifyMe(4) = Trim(pobjRS.Fields("notifyLastName").Value & "")
			aryNotifyMe(5) = Trim(pobjRS.Fields("notifyFirstName").Value & "")
			aryNotifyMe(6) = Trim(pobjRS.Fields("notifyEmail").Value & "")
			aryNotifyMe(7) = pobjRS.Fields("notifyType").Value
			aryNotifyMe(8) = pobjRS.Fields("notifyDateCreated").Value
			aryNotifyMe(9) = pobjRS.Fields("notifyDateNotified").Value
			aryNotifyMe(10) = pobjRS.Fields("notifyNotifyCount").Value
			
			LoadNotifyMeByInventoryID = True
		Else
			LoadNotifyMeByInventoryID = False
		End If	'pobjRS.EOF

		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

End Function	'LoadNotifyMeByInventoryID

'**********************************************************

Function saveNotifyMe(byVal tmpOrderDetailId, byVal notifyLastName, byVal notifyFirstName, byVal notifyEmail)

Dim pobjCmd
Dim pobjRS
Dim pobjRSAttributes
Dim pstrProductID
Dim pstrAttributeDetailID
Dim pstrProductName
Dim plngNotifyMeID
Dim plngInventoryID

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd

		.Commandtype = adCmdText
		.Commandtext = "SELECT sfTmpOrderDetails.odrdttmpProductID, sfTmpOrderAttributes.odrattrtmpAttrID, sfProducts.prodName FROM (sfTmpOrderDetails INNER JOIN sfProducts ON sfTmpOrderDetails.odrdttmpProductID = sfProducts.prodID) LEFT JOIN sfTmpOrderAttributes ON sfTmpOrderDetails.odrdttmpID = sfTmpOrderAttributes.odrattrtmpOrderDetailId WHERE sfTmpOrderDetails.odrdttmpID=?"
		.Parameters.Append .CreateParameter("odrdttmpID", adInteger, adParamInput, 4, tmpOrderDetailId)
		Set .ActiveConnection = cnn
		Set pobjRS = .Execute

		If pobjRS.EOF Then
			saveNotifyMe = False
		Else
			.Parameters.Append .CreateParameter("AttrdtID", adInteger, adParamInput, 4, tmpOrderDetailId)
			Do While Not pobjRS.EOF
				If pstrProductID <> Trim(pobjRS.Fields("odrdttmpProductID").Value & "") Then
					pstrProductID = Trim(pobjRS.Fields("odrdttmpProductID").Value & "")
					pstrProductName = Trim(pobjRS.Fields("prodName").Value & "")
					plngInventoryID = getInventoryRecordID(pstrProductID, GetAttDetailID (tmpOrderDetailId, "tmp"))
					plngNotifyMeID = createNotifyMe(pstrProductID, notifyLastName, notifyFirstName, notifyEmail, 0, plngInventoryID)
				End If

				If Not isNull(pobjRS.Fields("odrattrtmpAttrID").Value) Then
					.Parameters("odrdttmpID").Value = plngNotifyMeID
					.Parameters("AttrdtID").Value = Trim(pobjRS.Fields("odrattrtmpAttrID").Value & "")
					.Commandtext = "Select notifyMeAttributeID From notifyMeAttributes Where notifyMeID=? And AttrdtID=?"
					Set pobjRSAttributes = .Execute
					If pobjRSAttributes.EOF Then
						.Commandtext = "Insert Into notifyMeAttributes (notifyMeID, AttrdtID) Values (?, ?)"
						.Execute , , adExecuteNoRecords
					End If
				End If
				
				pobjRS.MoveNext
			Loop
		
			saveNotifyMe = True
		End If
	End With

End Function	'saveNotifyMe

'**********************************************************
%>
