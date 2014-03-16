<%
'********************************************************************************
'*
'*   incGeneral.asp
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins is incGeneral.asp APPVERSION: 50.4014.0.3
'*   from LaGarde, Incorporated. It has been heavily modified by Sandshot Software
'*   with permission from LaGarde, Incorporated who has NOT released any of the 
'*   original copyright protections. As such, this is a derived work covered by
'*   the copyright provisions of both companies.
'*
'*   LaGarde, Incorporated Copyright Statement                                                                           *
'*   The contents of this file is protected under the United States copyright
'*   laws and is confidential and proprietary to LaGarde, Incorporated.  Its 
'*   use ordisclosure in whole or in part without the expressed written 
'*   permission of LaGarde, Incorporated is expressly prohibited.
'*   (c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'*   
'*   Sandshot Software Copyright Statement
'*   The contents of this file are protected by United States copyright laws 
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************
%>
<!--#include file="storeAdminSettings.asp"-->
<!--#include file="mail.asp"-->
<!--#include file="ssmodAttributeExtender.asp"-->
<!--#include file="ssmodVisitor.asp"-->
<!--#include file="ssmodProduct.asp"-->
<!--#include file="ssPricingLevels.asp"-->
<!--#include file="incAddProduct.asp"-->
<!--#include file="ssclsCartContents.asp"-->
<%
'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

Const en_CatFields_uid = 0
Const en_CatFields_ParentLevel = 1
Const en_CatFields_Name = 2
Const en_CatFields_ParentID = 3
Const en_CatFields_IsActive = 4
Const en_CatFields_Description = 5
Const en_CatFields_URL = 6
Const en_CatFields_ImagePath = 7
Const en_CatFields_CategoryID = 8
Const en_CatFields_IsBottom = 9
Const en_CatFields_InTrail = 10
Const en_CatFields_NumFields = 10

Dim maryCountry
Dim maryStates
Dim maryCreditCards
Dim maryManufacturers
Dim maryVendors

'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'*	Begin Page Code
'**********************************************************

'**********************************************************
'**********************************************************

'***********************************************************************************************

Function bubbleSort2DArray(byVal ary, byVal columnToSort, byVal sortType, byVal blnSortAsc)
'use the bubble sort since it is generally the best to use for smaller (ie < 25) items)

Dim i
Dim j
Dim paryList
Dim pblnSwap
Dim paryTemp

	paryList = ary

	For i = UBound(paryList) - 1 To 0 Step -1
		For j = 0 To i
			pblnSwap = False
			Select Case sortType
				Case "number":
					If blnSortAsc Then
						pblnSwap = CBool(CDbl(paryList(j)(columnToSort)) > CDbl(paryList(j+1)(columnToSort)))
					Else
						pblnSwap = CBool(CDbl(paryList(j)(columnToSort)) < CDbl(paryList(j+1)(columnToSort)))
					End If
				Case "date":
					If blnSortAsc Then
						pblnSwap = CBool(CDate(paryList(j)(columnToSort)) > CDate(paryList(j+1)(columnToSort)))
					Else
						pblnSwap = CBool(CDate(paryList(j)(columnToSort)) < CDate(paryList(j+1)(columnToSort)))
					End If
				Case Else
					If blnSortAsc Then
						pblnSwap = CBool(CStr(paryList(j)(columnToSort)) > CStr(paryList(j+1)(columnToSort)))
					Else
						pblnSwap = CBool(CStr(paryList(j)(columnToSort)) < CStr(paryList(j+1)(columnToSort)))
					End If
			End Select
			
			If pblnSwap Then
				If blnSortAsc Then
					paryTemp = paryList(j+1)
					paryList(j+1) = paryList(j)
					paryList(j) = paryTemp
				Else
					paryTemp = paryList(j+1)
					paryList(j+1) = paryList(j)
					paryList(j) = paryTemp
				End If
			End If
		Next 'j
	Next 'i

	bubbleSort2DArray = paryList
	
End Function	'bubbleSort2DArray

'***********************************************************************************************

Function bubbleSortList(byVal strList, byVal strDelimiter)
'use the bubble sort since it is generally the best to use for smaller (ie < 25) items)

Dim i
Dim j
Dim paryList
Dim pstrTemp

	If Len(strList) > 0 Then
		paryList = Split(strList, strDelimiter)
		For i = UBound(paryList) - 1 To 0 Step -1
			For j = 0 To i
				If CLng(paryList(j)) > CLng(paryList(j+1)) Then
					pstrTemp = paryList(j+1)
					paryList(j+1) = paryList(j)
					paryList(j) = pstrTemp
				End If
			Next 'j
		Next 'i
		
		strList = paryList(0)
		For i = 1 To UBound(paryList)
			strList = strList & strDelimiter & paryList(i)
		Next 'i
	End If	'Len(strList) > 0
	
	bubbleSortList = strList
	
End Function	'bubbleSortList

'***********************************************************************************************

Function getLastSearch()

Dim i
Dim paryNameValue
Dim paryQuerystring
Dim paryTemp
Dim pstrLastSearch
Dim pstrNameValue
Dim pstrTemp

	On Error Resume Next
	
	pstrLastSearch = visitorLastSearch
	
	If Len(pstrLastSearch) = 0 Or InStr(LCase(pstrLastSearch), "login.asp") <> 0 Then
		pstrLastSearch = "search.asp"
	Else
		'Need to strip out an back-order items
		paryQuerystring = Split(pstrLastSearch, "?")
		If isArray(paryQuerystring) Then
			If UBound(paryQuerystring) > 0 Then
				paryTemp = Split(paryQuerystring(1), "&")
				For i = 0 To UBound(paryTemp)
					paryNameValue = Split(paryTemp(i), "=")
					If isArray(paryNameValue) Then
						pstrNameValue = ""
						If UBound(paryNameValue) >= 0 Then pstrNameValue = paryNameValue(0)
						Select Case pstrNameValue
							Case "btnAction", "BackOrderPos", "BackOrderCount", "OrderQty", "notifyLastName", "notifyFirstName", "notifyEmail"
								'do nothing
							Case Else
								If Len(pstrTemp) = 0 Then
									If UBound(paryNameValue) = 1 Then
										pstrTemp = "?" & paryNameValue(0) & "=" & paryNameValue(1)
									ElseIf UBound(paryNameValue) >= 0 Then
										pstrTemp = "?" & "&amp;" & paryNameValue(0)
									Else
										pstrTemp = ""
									End If
								Else
									If UBound(paryNameValue) = 1 Then
										pstrTemp = pstrTemp & "&amp;" & paryNameValue(0) & "=" & paryNameValue(1)
									Else
										pstrTemp = pstrTemp & "&amp;" & paryNameValue(0)
									End If
								End If
						End Select
					End If	'isArray(paryNameValue)
				Next 'i
			End If	'UBound(paryTemp) > 0
			If Len(pstrTemp) = 0 Then
				pstrTemp = paryQuerystring(0)
			Else
				pstrTemp = paryQuerystring(0) & pstrTemp
			End If
		End If	'isArray(paryQuerystring)
	End If 
	
	getLastSearch = pstrTemp
	
End Function	'getLastSearch

'*******************************************************************************************************

Function isOrderPage()
	isOrderPage = CBool(CurrentPage = "order.asp")
End Function	'isOrderPage

'*******************************************************************************************************

Function isCheckoutPage()
	
	Select Case CurrentPage
		Case "customerLogin.asp", "process_order.asp", "verify.asp", "confirm.asp"
			isCheckoutPage = True
		Case Else
			isCheckoutPage = False
	End Select

End Function	'isCheckoutPage

'**********************************************************

Function MakeUSDate(byVal InDate)
	If IsDate(InDate) Then
		MakeUSDate = Month(InDate) & "/" & Day(InDate) & "/" & Right(Year(InDate),2)
	End If
End Function	'MakeUSDate

'---------------------------------------------------------------------
' Purpose: Deletes recordset from TmpOrders and associated child relations
'---------------------------------------------------------------------
Function setDeleteOrder(byVal sPrefix, byVal iOrderDetailId)

Dim pblnResult
Dim pobjCmd

	pblnResult = False

	If Len(iOrderDetailId) > 0 And isNumeric(iOrderDetailId) Then

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			'.Commandtype = adCmdStoredProc
			Set .ActiveConnection = cnn

			.Parameters.Append .CreateParameter("iOrderDetailId", adInteger, adParamInput, 4, iOrderDetailId)
		
			Select Case sPrefix
				Case "odrdttmp"
					.Commandtext = "Delete From sfTmpOrderAttributes Where odrattrtmpOrderDetailId=?"
					.Execute , , adExecuteNoRecords
								
					If cblnSF5AE Then
						.Commandtext = "Delete From sfOrderDetailsAE Where odrdtAEID=?"
						.Execute , , adExecuteNoRecords
	  				End If

					.Commandtext = "Delete From sfTmpOrderDetails Where odrdttmpID=?"
					.Execute , , adExecuteNoRecords
					
					pblnResult = CBool(Err.number = 0)
				Case "odrdtsvd"
					.Commandtext = "Delete From sfSavedOrderDetails Where odrdtsvdID=?"
					.Execute , , adExecuteNoRecords
								
					.Commandtext = "Delete From sfSavedOrderAttributes Where odrattrsvdOrderDetailId=?"
					.Execute , , adExecuteNoRecords

					pblnResult = CBool(Err.number = 0)
			End Select	

		End With
		Set pobjCmd  = Nothing
	End If
	
	setDeleteOrder = pblnResult
		
End Function	'setDeleteOrder 	

'---------------------------------------------------------------------
' Collect Attribute IDs
'---------------------------------------------------------------------
Function getProdAttr(byVal sPrefix, byVal sOrderID, byRef iProdAttrNum)

Dim paryOut
Dim paryTemp
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim i

	Set pobjCmd  = CreateObject("ADODB.Command")

	Select Case sPrefix
		Case "odrattrtmp"
			pstrSQL = "SELECT odrattrtmpAttrID, odrattrtmpAttrText FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=?"
			pobjCmd.Parameters.Append pobjCmd.CreateParameter("FindKey", adInteger, adParamInput, 4, sOrderID)
		Case "odrattrsvd"	
			pstrSQL = "SELECT odrattrsvdAttrID, odrattrsvdAttrText FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailId=?"
			pobjCmd.Parameters.Append pobjCmd.CreateParameter("FindKey", adInteger, adParamInput, 4, sOrderID)
		Case "odr"
			pstrSQL = "SELECT odrattrID, odrattrAttribute FROM sfOrderAttributes WHERE odrattrOrderDetailId=?"
			pobjCmd.Parameters.Append pobjCmd.CreateParameter("FindKey", adInteger, adParamInput, 4, sOrderID)
	End Select 
	
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
  		Set pobjRS = .Execute
	End With
	
  	If pobjRS.EOF Then
  		'If vDebug = 1 Then Response.Write "Either the recordset doesn't exit or the field name is not typed correctly :<br />" & pstrSQL			
  	Else
  		paryTemp = pobjRS.GetRows()
  		Redim paryOut(UBound(paryTemp, 2))	
  		For i = 0 to UBound(paryOut)
			paryOut(i) = BuildAttribute(paryTemp(0, i), paryTemp(1, i))
  			If vDebug = 1 Then Response.Write "<br />AttrID: " & paryOut(i)
  		Next
  	End If
  	closeObj(pobjRS)
	Set pobjCmd = Nothing
	
	getProdAttr = paryOut
  	  
End Function	'getProdAttr

'-------------------------------------------------------
' Update saved cart customers' info in sfCustomers
'-------------------------------------------------------
Sub setUpdateCustomer(sNewEmail,sFirstName,sMiddleInitial,sLastName,sCompany,sAddress1,sAddress2,sCity,sState,sZip,sCountry,sPhone,sFax,bSubscribed)
	Dim	sLocalSQl, rsUpdate, iOldNum
	
	sLocalSQL = "Select custFirstName, custMiddleInitial, custLastName, custCompany, custAddr1, custAddr2, custCity, custState, custZip, custCountry, "_
				& "custPhone, custFax, custTimesAccessed, custLastAccess, custEmail, custIsSubscribed FROM sfCustomers WHERE custID = " & custID_cookie
	
	Set rsUpdate = CreateObject("ADODB.RecordSet")
		rsUpdate.Open sLocalSQL,cnn,adOpenDynamic,adLockOptimistic,adCmdText
		
		If Not rsUpdate.EOF Then
				iOldNum = (rsUpdate.Fields("custTimesAccessed"))
				If iOldNum = "" or isnull(iOldNum) Then 
					iOldNum = 1 
				Else
					iOldNum = cInt(iOldNum)	
				End If		
				rsUpdate.Fields("custFirstName")		= sFirstName
				rsUpdate.Fields("custMiddleInitial")	= sMiddleInitial
				rsUpdate.Fields("custLastName")		= sLastName
				rsUpdate.Fields("custCompany")			= sCompany
				rsUpdate.Fields("custAddr1")			= sAddress1
				rsUpdate.Fields("custAddr2")			= sAddress2
				rsUpdate.Fields("custCity")				= sCity
				rsUpdate.Fields("custState")			= sState
				rsUpdate.Fields("custZip")				= sZip
				rsUpdate.Fields("custCountry")			= sCountry
				rsUpdate.Fields("custPhone")			= sPhone
				rsUpdate.Fields("custFax")				= sFax	
				rsUpdate.Fields("custTimesAccessed")	= iOldNum + 1
				rsUpdate.Fields("custLastAccess")		= Date()
				If sNewEmail <> "" Then
					rsUpdate.Fields("custEmail")		= sNewEmail
				End If		
				If CStr(bSubscribed) = "" Or CStr(bSubscribed) = "0" Then
                             rsUpdate.Fields("custissubscribed") = 0
                Else
                             rsUpdate.Fields("custissubscribed") = 1
                End If
				rsUpdate.Update		
		End If
		closeObj(rsUpdate)	
End Sub

'---------------------------------------------------------------------
' This function returns one specific value associated with a single id
' Used for lookup of VendorID, ManufacturerID, CategoryID, etc
'---------------------------------------------------------------------
Function getNameWithID(byVal sLocalTableName, byVal sLocalFindKey, byVal sLocalFindKeyLabel,byVal sLocalSearchName, byVal bStringOrNot)

Dim pblnIsNumeric
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim pvntValue

	If Len(sLocalFindKey & "") = 0 Then Exit Function
	If Not (bStringOrNot = 0 Or bStringOrNot = 1)  Then Exit Function

	pstrSQL = "SELECT " & sLocalSearchName & " FROM " & sLocalTableName & " WHERE " & sLocalFindKeyLabel & "=?"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn

		pblnIsNumeric = CBool(bStringOrNot = 0)
		If pblnIsNumeric And Not isNumeric(sLocalFindKey) Then pblnIsNumeric = False
		If pblnIsNumeric Then
			.Parameters.Append .CreateParameter("FindKey", adInteger, adParamInput, 4, sLocalFindKey)
		Else
			.Parameters.Append .CreateParameter("FindKey", adVarChar, adParamInput, Len(sLocalFindKey), sLocalFindKey)
  		End If

  		Set pobjRS = .Execute
  		If pobjRS.EOF Then
  			'If vDebug = 1 Then Response.Write "Either the recordset doesn't exit or the field name is not typed correctly :<br />" & sLocalSQL			
  		Else
  		  pvntValue = pobjRS.Fields("" &sLocalSearchName& "").Value
  		End If
  		closeObj(pobjRS)
	End With
	Set pobjCmd = Nothing
	
    getNameWithID = pvntValue

End Function	'getNameWithID

'---------------------------------------------------------------------
' This function returns one specific value associated with a single id
' Used for lookup of VendorID, ManufacturerID, CategoryID, etc
'---------------------------------------------------------------------
Function getSavedAttributeText(byVal odrattrsvdOrderDetailId, byVal odrattrsvdAttrID)

Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim pvntValue

	If Len(odrattrsvdOrderDetailId & "") = 0 Then Exit Function
	If Len(odrattrsvdAttrID & "") = 0 Then Exit Function

	pstrSQL = "SELECT odrattrsvdAttrText FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailId=? AND odrattrsvdAttrID=?"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("odrattrsvdOrderDetailId", adInteger, adParamInput, 4, odrattrsvdOrderDetailId)
		.Parameters.Append .CreateParameter("odrattrsvdAttrID", adInteger, adParamInput, 4, odrattrsvdAttrID)
		
  		Set pobjRS = .Execute
  		If pobjRS.EOF Then
			If vDebug = 1 Then
				Response.Write "<fieldset><legend>getSavedAttributeText</legend>"
				Response.Write ".Commandtext: " & .Commandtext & "<br />"
				Response.Write "odrattrsvdOrderDetailId: " & .Parameters("odrattrsvdOrderDetailId").Value & "<br />"
				Response.Write "odrattrsvdAttrID: " & .Parameters("odrattrsvdAttrID").Value & "<hr>"
				Response.Write "Either the recordset doesn't exit or the field name is not typed correctly<br />"
				Response.Write "</fieldset>"
			End If
  		Else
  			pvntValue = pobjRS.Fields("odrattrsvdAttrText").Value
			If vDebug = 1 Then
				Response.Write "<fieldset><legend>getSavedAttributeText</legend>"
				Response.Write ".Commandtext: " & .Commandtext & "<br />"
				Response.Write "odrattrsvdOrderDetailId: " & .Parameters("odrattrsvdOrderDetailId").Value & "<br />"
				Response.Write "odrattrsvdAttrID: " & .Parameters("odrattrsvdAttrID").Value & "<hr>"
				Response.Write "odrattrsvdAttrText: " & pvntValue & "<br />"
				Response.Write "</fieldset>"
			End If
  		End If
  		closeObj(pobjRS)
	End With
	Set pobjCmd = Nothing
	
    getSavedAttributeText = pvntValue

End Function	'getSavedAttributeText

'---------------------------------------------------------------------
' Enters record svdOrders, returns the ID of the SvdOrder
'---------------------------------------------------------------------
Function getSavedTable(byVal aProdAttr, byVal sProdID, byVal iNewQuantity, byVal iCustID, byVal sReferer)

Dim i
Dim plngID
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim pstrtmpAttrID
Dim pstrtmpAttrValue

	pstrSQL = "Insert Into sfSavedOrderDetails (odrdtsvdQuantity, odrdtsvdProductID, odrdtsvdSessionID, odrdtsvdCustID, odrdtsvdDate, odrdtsvdHttpReferer) Values (?, ?, ?, ?, ?, ?)"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		'.Commandtype = adCmdStoredProc
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("odrdtsvdQuantity", adDouble, adParamInput, 4, iNewQuantity)
		.Parameters.Append .CreateParameter("odrdtsvdProductID", adWChar, adParamInput, 50, sProdID)
		.Parameters.Append .CreateParameter("odrdtsvdSessionID", adInteger, adParamInput, 4, SessionID)
		.Parameters.Append .CreateParameter("odrdtsvdCustID", adInteger, adParamInput, 4, iCustID)
		.Parameters.Append .CreateParameter("odrdtsvdDate", adDBTimeStamp, adParamInput, 16, Now())

		If Len(sReferer) > 255 Then
			.Parameters.Append .CreateParameter("odrdtsvdHttpReferer", adVarChar, adParamInput, 255, Left(sReferer, 255))
		ElseIf Len(sReferer) = 0 Then
			.Parameters.Append .CreateParameter("odrdtsvdHttpReferer", adVarChar, adParamInput, 255, NULL)
		Else
			.Parameters.Append .CreateParameter("odrdtsvdHttpReferer", adVarChar, adParamInput, 255, sReferer)
		End If
		
		.Execute , , adExecuteNoRecords

		.Parameters.Delete "odrdtsvdHttpReferer"

		pstrSQL = "Select odrdtsvdID From sfSavedOrderDetails Where odrdtsvdQuantity=? And odrdtsvdProductID=? And odrdtsvdSessionID=? And odrdtsvdCustID=? And odrdtsvdDate=? Order By odrdtsvdID Desc"
		.Commandtext = pstrSQL

		Set pobjRS = .Execute
		plngID  = pobjRS.Fields("odrdtsvdID").Value
  		closeObj(pobjRS)
  		
		If vDebug = 1 Then Response.Write "<p><font size=4><b>odrdtsvdID = " & plngID & "</b></font>"
		' Copy Attributes
			
		' Collect Attribute Info from sfSavedOrderAttributes
		If IsArray(aProdAttr) Then
			.Parameters.Delete "odrdtsvdQuantity"
			.Parameters.Delete "odrdtsvdProductID"
			.Parameters.Delete "odrdtsvdSessionID"
			.Parameters.Delete "odrdtsvdCustID"
			.Parameters.Delete "odrdtsvdDate"

			pstrSQL = "Insert Into sfSavedOrderAttributes (odrattrsvdOrderDetailId, odrattrsvdAttrID, odrattrsvdAttrText) Values (?, ?, ?)"
			.Commandtext = pstrSQL
			
			'i = 0
			'Do While Len(aProdAttr(i)) > 0
			For i = 0 To UBound(aProdAttr)
				pstrtmpAttrID = GetAttributeID(aProdAttr(i))
				If Len(pstrtmpAttrID) > 0 Then
					pstrtmpAttrValue = GetAttributeValue(aProdAttr(i))
					If vDebug = 1 Then
						Response.Write "<fieldset><legend>getTmpTable - Attribute " & i & " = " & aProdAttr(i) & "</legend>"
						Response.Write "plngID: " & plngID & "<br />"
						Response.Write "pstrtmpAttrID: " & pstrtmpAttrID & "<br />"
						Response.Write "pstrtmpAttrValue: " & pstrtmpAttrValue & "<br />"
						Response.Write "</fieldset>"
					End If
					If Len(pstrtmpAttrValue) = 0 Then pstrtmpAttrValue = Null
					
					If .Parameters.Count = 0 Then
						.Parameters.Append .CreateParameter("odrattrsvdOrderDetailId", adInteger, adParamInput, 4, plngID)
						.Parameters.Append .CreateParameter("odrattrsvdAttrID", adInteger, adParamInput, 4, pstrtmpAttrID)
						.Parameters.Append .CreateParameter("odrattrsvdAttrText", adLongVarWChar, adParamInput, 2147483646, pstrtmpAttrValue)
					Else
						.Parameters("odrattrsvdOrderDetailId").Value = plngID
						.Parameters("odrattrsvdAttrID").Value = pstrtmpAttrID
						.Parameters("odrattrsvdAttrText").Value = pstrtmpAttrValue
					End If
					.Execute , , adExecuteNoRecords
				End If	'Len(pstrtmpAttrID) > 0
			Next 'i
			'	i = i + 1			
			'Loop
		' End IsArray If 	
		End If
  		
	End With	'pobjCmd	
  	closeobj(pobjCmd)
	
	If vDebug = 1 Then 	Response.Write "<p><font color=""red"" face=""verdana"" size=""2"">Copied Record To TmpOrder</font>"			
	
  	getSavedTable = plngID
  	
End Function	'getSavedTable


'---------------------------------------------------------------------
' Enters record TmpOrders, returns the ID of the TmpOrder
'---------------------------------------------------------------------
Function getTmpTable(byVal aProdAttr, byVal sProdID, byVal iNewQuantity, byVal sReferer, byVal iShip)

Dim i
Dim plngID
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim pstrtmpAttrID
Dim pstrtmpAttrValue

	pstrSQL = "Insert Into sfTmpOrderDetails (odrdttmpQuantity, odrdttmpProductID, odrdttmpSessionID) Values (?, ?, ?)"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		'.Commandtype = adCmdStoredProc
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("odrdttmpQuantity", adDouble, adParamInput, 4, iNewQuantity)
		.Parameters.Append .CreateParameter("odrdttmpProductID", adWChar, adParamInput, 50, sProdID)
		.Parameters.Append .CreateParameter("odrdttmpSessionID", adInteger, adParamInput, 4, SessionID)
		.Execute , , adExecuteNoRecords

		pstrSQL = "Select odrdttmpID From sfTmpOrderDetails Where odrdttmpQuantity=? And odrdttmpProductID=? And odrdttmpSessionID=? Order By odrdttmpID Desc"
		.Commandtext = pstrSQL

		Set pobjRS = .Execute
		plngID  = pobjRS.Fields("odrdttmpID").Value
  		closeObj(pobjRS)
  		
		If vDebug = 1 Then Response.Write "<b>TmpCart Key ID = " & plngID & "</b><br /></font>"
		If vDebug = 1 Then Response.Write "Has Attributes: " & IsArray(aProdAttr) & "<br />"
			
		' Collect Attribute Info from sfTmpOrderAttributes
		If IsArray(aProdAttr) Then
			.Parameters.Delete "odrdttmpQuantity"
			.Parameters.Delete "odrdttmpProductID"
			.Parameters.Delete "odrdttmpSessionID"

			pstrSQL = "Insert Into sfTmpOrderAttributes (odrattrtmpOrderDetailId, odrattrtmpAttrID, odrattrtmpAttrText) Values (?, ?, ?)"
			.Commandtext = pstrSQL
			
			i = 0
			Do While Len(aProdAttr(i)) > 0
			'For i = 0 To UBound(aProdAttr)
				pstrtmpAttrID = GetAttributeID(aProdAttr(i))
				If Len(pstrtmpAttrID) = 0 Then pstrtmpAttrID = 0
				pstrtmpAttrValue = GetAttributeValue(aProdAttr(i))
				If vDebug = 1 Then
					Response.Write "<fieldset><legend>getTmpTable - Attribute " & i & " = " & aProdAttr(i) & "</legend>"
					Response.Write "plngID: " & plngID & "<br />"
					Response.Write "pstrtmpAttrID: " & pstrtmpAttrID & "<br />"
					Response.Write "pstrtmpAttrValue: " & pstrtmpAttrValue & "<br />"
					Response.Write "</fieldset>"
				End If
				If Len(pstrtmpAttrValue & "") = 0 Then pstrtmpAttrValue = Null
				
				If .Parameters.Count = 0 Then
					.Parameters.Append .CreateParameter("odrattrtmpOrderDetailId", adInteger, adParamInput, 4, plngID)
					.Parameters.Append .CreateParameter("odrattrtmpAttrID", adInteger, adParamInput, 4, pstrtmpAttrID)
					.Parameters.Append .CreateParameter("odrattrtmpAttrText", adLongVarWChar, adParamInput, 2147483646, pstrtmpAttrValue)
				Else
					.Parameters("odrattrtmpOrderDetailId").Value = plngID
					.Parameters("odrattrtmpAttrID").Value = pstrtmpAttrID
					.Parameters("odrattrtmpAttrText").Value = pstrtmpAttrValue
				End If
				.Execute , , adExecuteNoRecords
			'Next 'i
				i = i + 1
				If i > UBound(aProdAttr) Then Exit Do
			Loop
		' End IsArray If 	
		End If
  		
	End With	'pobjCmd	
  	closeobj(pobjCmd)
	
	If vDebug = 1 Then 	Response.Write "<p><font color=""red"" face=""verdana"" size=""2"">Copied Record To TmpOrder</font>"			
	
  	getTmpTable = plngID
  	
End Function	'getTmpTable

'---------------------------------------------------------------------
' Purpose: Updates the Quantity field with associated prodId and CartID
'---------------------------------------------------------------------
Sub setUpdateQuantity(byVal sPrefix, byVal iQuantity, byVal iTmpOrderID)

Dim iOldQuantity, iNewQuantity
Dim pobjCmd
Dim pobjRS

	If Len(iTmpOrderID) > 0 And isNumeric(iTmpOrderID) Then

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			Set .ActiveConnection = cnn
			.Commandtype = adCmdText
			.Parameters.Append .CreateParameter("odrdttmpID", adInteger, adParamInput, 4, iTmpOrderID)
			Select Case sPrefix
				Case "odrdttmp"	
					.Commandtext = "SELECT odrdttmpQuantity FROM sfTmpOrderDetails WHERE odrdttmpID=? AND odrdttmpSessionID=?"
					If vDebug = 1 Then Response.Write "<br /> setUpdateQuantity SQL : " & .Commandtext & "<br />"
					.Parameters.Append .CreateParameter("SessionID", adInteger, adParamInput, 4, SessionID)
				Case "odrdtsvd" 
					.Commandtext = "SELECT odrdtsvdQuantity FROM sfSavedOrderDetails WHERE odrdtsvdID=? AND odrdtsvdCustID=?"
					.Parameters.Append .CreateParameter("SessionID", adInteger, adParamInput, 4, custID_cookie)
			End Select

			Set pobjRS = .Execute
			If pobjRS.EOF Then
				Call abandonSession
			Else
				iOldQuantity = pobjRS.Fields(0).Value
			End If
  			closeObj(pobjRS)

			iNewQuantity = cInt(iOldQuantity) + cInt(iQuantity)

			.Parameters.Delete "odrdttmpID"
			.Parameters.Delete "SessionID"
			
			.Parameters.Append .CreateParameter("qty", adDouble, adParamInput, 4, iNewQuantity)
			.Parameters.Append .CreateParameter("odrdttmpID", adInteger, adParamInput, 4, iTmpOrderID)
			Select Case sPrefix
				Case "odrdttmp"	
					.Commandtext = "Update sfTmpOrderDetails Set odrdttmpQuantity=? WHERE odrdttmpID=? AND odrdttmpSessionID=?"
					.Parameters.Append .CreateParameter("SessionID", adInteger, adParamInput, 4, SessionID)
				Case "odrdtsvd" 
					.Commandtext = "Update sfSavedOrderDetails Set odrdtsvdQuantity=? WHERE odrdtsvdID=? AND odrdtsvdCustID=?"
					.Parameters.Append .CreateParameter("SessionID", adInteger, adParamInput, 4, custID_cookie)
			End Select
			
			.Execute , , adExecuteNoRecords
		End With
		Set pobjCmd = Nothing
	End If	'Len(iTmpOrderID) > 0 And isNumeric(iTmpOrderID)

End Sub	'setUpdateQuantity

'---------------------------------------------------------------------
' Purpose: Updates the Quantity field with associated prodId and CartID
'---------------------------------------------------------------------
Sub setReplaceQuantity(byVal sPrefix, byVal iQuantity, byVal iTmpOrderID)

Dim pobjCmd

	If Len(iTmpOrderID) > 0 And isNumeric(iTmpOrderID) Then

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			Set .ActiveConnection = cnn
			.Commandtype = adCmdText
			
			.Parameters.Append .CreateParameter("qty", adDouble, adParamInput, 4, iQuantity)
			.Parameters.Append .CreateParameter("odrdttmpID", adInteger, adParamInput, 4, iTmpOrderID)
			Select Case sPrefix
				Case "odrdttmp"	
					.Commandtext = "Update sfTmpOrderDetails Set odrdttmpQuantity=? WHERE odrdttmpID=? AND odrdttmpSessionID=?"
					.Parameters.Append .CreateParameter("SessionID", adInteger, adParamInput, 4, SessionID)
				Case "odrdtsvd" 
					.Commandtext = "Update sfSavedOrderDetails Set odrdtsvdQuantity=? WHERE odrdtsvdID=? AND odrdtsvdCustID=?"
					.Parameters.Append .CreateParameter("SessionID", adInteger, adParamInput, 4, custID_cookie)
			End Select
			.Execute , , adExecuteNoRecords

		End With
		Set pobjCmd = Nothing
	End If	'Len(iTmpOrderID) > 0 And isNumeric(iTmpOrderID)

End Sub	'setReplaceQuantity
 

'---------------------------------------------------------------------
' Checks for existence of same product and attributes (if any)
' Returns the OrderDetail ID or -1 if record DNE
'---------------------------------------------------------------------
Function getOrderID(byVal sPrefix, byVal sAttrPrefix, byVal sProdID, byVal aProdAttr, byVal iProdAttrNum, byRef lngQuantity)

Dim sTmpVar, rsSelectProd, sTmpPrefixID, sTmpAttrName, sTmpAttr
Dim sLocalSQL, sAttrName, bMatch, iUpperBound

Dim i
Dim plngCounter
Dim plngID
Dim pobjCmd
Dim pobjRS
Dim pstrSQL
Dim pstrtmpAttrID
Dim pstrtmpAttrValue
	
	plngID = 0
	bMatch = 0

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd

		Select Case sPrefix
			Case "odrdttmp"	
				pstrSQL = "SELECT odrdttmpQuantity as Quantity, odrdttmpID, odrattrtmpAttrID, odrattrtmpAttrText" _
						& " FROM sfTmpOrderAttributes RIGHT JOIN sfTmpOrderDetails ON sfTmpOrderAttributes.odrattrtmpOrderDetailId = sfTmpOrderDetails.odrdttmpID" _
						& " WHERE odrdttmpSessionID=" & wrapSQLValue(SessionID, False, enDatatype_number) & " AND odrdttmpProductID=" & wrapSQLValue(sProdID, False, enDatatype_string)
				pstrSQL = "SELECT odrdttmpQuantity as Quantity, odrdttmpID, odrattrtmpAttrID, odrattrtmpAttrText" _
						& " FROM sfTmpOrderAttributes RIGHT JOIN sfTmpOrderDetails ON sfTmpOrderAttributes.odrattrtmpOrderDetailId = sfTmpOrderDetails.odrdttmpID" _
						& " WHERE odrdttmpSessionID=? AND odrdttmpProductID=?"
				.Parameters.Append .CreateParameter("SessionID", adInteger, adParamInput, 4, SessionID)
				.Parameters.Append .CreateParameter("ProductID", adWChar, adParamInput, 50, sProdID)
			Case "odrdtsvd" 
				pstrSQL = "SELECT odrdtsvdQuantity as Quantity, odrdtsvdID, odrattrsvdAttrID, odrattrsvdAttrText" _
						& " FROM sfSavedOrderDetails RIGHT JOIN sfSavedOrderAttributes ON sfSavedOrderDetails.odrdtsvdID = sfSavedOrderAttributes.odrattrsvdOrderDetailId " _
						& " WHERE odrdtsvdCustID=? AND odrdtsvdProductID=?"
				.Parameters.Append .CreateParameter("SessionID", adInteger, adParamInput, 4, custID_cookie)
				.Parameters.Append .CreateParameter("ProductID", adWChar, adParamInput, 50, sProdID)
			Case Else
				plngID = -1
		End Select
		
		If plngID = 0 Then
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			'.Commandtype = adCmdStoredProc
			Set .ActiveConnection = cnn
			Set pobjRS = .Execute
			'Set pobjRS = GetRS(pstrSQL)
		
			If vDebug = 1 Then 
				Response.Write "<fieldset><legend>getOrderID - Attributes</legend>"
				Response.Write "pstrSQL: " & pstrSQL & "<br />"
				Response.Write "SessionID: " & .Parameters("SessionID").Value & "<br />"
				Response.Write "sProdID: " & .Parameters("ProductID").Value & "<br />"
				Response.Write "Found? " & CStr(Not pobjRS.EOF) & "<br />"
				If isArray(aProdAttr) Then Response.Write "UBound(aProdAttr): " & UBound(aProdAttr) & "<br />"
				Response.Write "iProdAttrNum: " & iProdAttrNum & "<br />"
				If Not pobjRS.EOF Then Response.Write "isNull: " & isNull(pobjRS.Fields(sAttrPrefix & "AttrID").Value) & "<br />"
				Response.Write "</fieldset>"
			End If
			
			If pobjRS.EOF Then
				plngID = -1
				lngQuantity = 0
			Else
				lngQuantity = pobjRS.Fields("Quantity").Value
				If isNull(pobjRS.Fields(sAttrPrefix & "AttrID").Value) Then
					plngID = pobjRS.Fields(sPrefix & "ID")			
				Else 
					If vDebug = 1 And Not cblnSQLDatabase Then	'Rowset position cannot be restarted error in SQL Server
						Response.Write "<fieldset><legend>getOrderID - Attributes</legend>"
						Do While Not pobjRS.EOF
							Response.Write "ID : " & pobjRS.Fields(sPrefix & "Id").Value & " AttrID :" & pobjRS.Fields(sAttrPrefix & "AttrID").Value & "-" & pobjRS.Fields(sAttrPrefix & "AttrText").Value & "<br />"
							pobjRS.MoveNext
						Loop
						Response.Write "</fieldset>"
						pobjRS.MoveFirst
					End If

					If isArray(aProdAttr) Then
						iUpperBound = UBound(aProdAttr)
					Else
						iUpperBound = -1
					End If
					plngCounter = 0
					Do While Not pobjRS.EOF
						plngCounter = plngCounter + 1
						If vDebug = 1 Then
							Response.Write "<br />Position " & plngCounter & ", iUpperBound = " & iUpperBound 			
						End If
						For i = 0 to iUpperBound'-1 					
							sTmpAttr = aProdAttr(i)   

							If Len(sTmpAttr) = 0 Or pobjRS.EOF Then 
								plngID = -1
							Else  								
								pstrtmpAttrID = pobjRS.Fields(sAttrPrefix & "AttrID").Value
								pstrtmpAttrValue = Trim(pobjRS.Fields(sAttrPrefix & "AttrText").Value & "")	'could be a null

								'debugprint "isAttributeMatch", isAttributeMatch(sTmpAttr, pstrtmpAttrID, pstrtmpAttrValue)
								'debugprint "sTmpAttr", sTmpAttr
								If isAttributeMatch(sTmpAttr, pstrtmpAttrID, pstrtmpAttrValue) Then bMatch = bMatch + 1

								If vDebug = 1 Then
									Response.Write "<p>" & sTmpAttr & " VS " & pobjRS.Fields(sAttrPrefix & "AttrID").Value
									Response.Write "<br />bMatch = " & bMatch 			
								End If

								If bMatch = cInt(iProdAttrNum) Then
									If vDebug = 1 Then Response.Write "<br />Return the Found Record: " & pobjRS.Fields(sPrefix & "ID").Value & "<br />"
									plngID = pobjRS.Fields(sPrefix & "ID").Value
								End If					
							End If	'Len(sTmpAttr) = 0 Or pobjRS.EOF
							If Not pobjRS.EOF Then pobjRS.MoveNext
						Next	'i
						'This check added since not all products have attributes
						If i = 0 Then pobjRS.MoveNext
						
						If vDebug = 1 Then
							If bMatch = cInt(iProdAttrNum) Then
								Response.Write "<h4>Product Match Found in cart: " & plngID & "</h4>"
							Else
								Response.Write "<h4>NO Product Match Found in cart: </h4>"
							End If
						End If
						bMatch = 0	' Reset Match at end of Recordset
					Loop

				End If	'Not isNull(pobjRS.Fields(sAttrPrefix & "AttrID").Value)

			End If	'pobjRS.EOF
			closeObj(pobjRS)
			
		End If	'plngID = 0

	End With
	closeObj(pobjCmd)

	If plngID = 0 Then plngID = -1
	If plngID = -1 Then lngQuantity = 0
	
	getOrderID = plngID
	
End Function	'getOrderID

'---------------------------------------------------------------------
' Returns the name, price, and type associated with the attribute ID of Old Order
'---------------------------------------------------------------------
Function getAttrDetailsRetriveOrder(iAttrID)
Dim sLocalSQL, rsFindAttr, aLocalAttr

	sLocalSQL = "SELECT odrattrAttribute, odrattrName, odrattrPrice, odrattrType FROM sfOrderAttributes WHERE odrattrID = " & makeInputSafe(iAttrID)
		Set rsFindAttr = CreateObject("ADODB.RecordSet")
			rsFindAttr.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			If rsFindAttr.BOF Or rsFindAttr.EOF Then
				If vDebug = 1 Then Response.Write "<br />Empty Recordset in getAttrNames"
			Else
				Redim aLocalAttr(4)
				aLocalAttr(0) = rsFindAttr.Fields("odrattrAttribute")
				aLocalAttr(1) = rsFindAttr.Fields("odrattrPrice")
				aLocalAttr(2) = rsFindAttr.Fields("odrattrType")
				aLocalAttr(3) = rsFindAttr.Fields("odrattrName")
			End If
		
	closeObj(rsFindAttr)
	getAttrDetailsRetriveOrder = aLocalAttr
End Function

'---------------------------------------------------------------------
' This function calculates the subtotal for attributes
'---------------------------------------------------------------------
Function getAttrUnitPrice (dAttrTotal,sAttrPrice,iAttrType)
	' Recalculate Price
	If iAttrType = 1 Then
		dAttrTotal = dAttrTotal + cDbl(sAttrPrice)
	ElseIf iAttrType = 2  Then
		dAttrTotal = dAttrTotal + cDbl(sAttrPrice)*(-1)	
	End If
getAttrUnitPrice = dAttrTotal
End Function
'-------------------------------------------------------------------
' Returns the recordset corresponding to a custId identifier
'-------------------------------------------------------------------
Function getRow(sTableName,sIdName,iID,cnn)
	Dim sLocalSQL, rsSet
		
	sLocalSQL = "SELECT * FROM " & sTableName & " WHERE " & sIdName & " = " & makeInputSafe(iID)
		
	' Object Creation
	Set rsSet = CreateObject("ADODB.RecordSet")
	rsSet.Open sLocalSQL, cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
				
	Set getRow = rsSet
End Function
'-------------------------------------------------------------------
' Gets records for tables with multiple records for one customer ID
' Returns the recordset
'-------------------------------------------------------------------
Function getRowActive(sTableName,sIdName,sActiveName,iID,cnn)
	Dim sLocalSQL, rsSet
		
	sLocalSQL = "SELECT * FROM " & sTableName & " WHERE " & sIdName & " = " & makeInputSafe(iID) & " AND " & sActiveName  & " = 1"
		
	' Object Creation
	Set rsSet = CreateObject("ADODB.RecordSet")
	rsSet.Open sLocalSQL, cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
				
	Set getRowActive = rsSet
End Function


'-------------------------------------------------------
' Compares email and password, then returns the ID of the customer
' Returns -1 for failed authentication
'-------------------------------------------------------
Function customerAuth(byVal sEmail, byVal sPassword, byVal sType)

Dim plngCustID
Dim pclsCustomer

	If Len(sEmail) = 0 Or Len(sPassword) = 0 Then		
		customerAuth = -1
		Exit Function
	End If
		
	Select Case sType
		Case "strict_NEVERUSED"
			sLocalSQL = "SELECT custID FROM sfCustomers WHERE custEmail = '" & makeInputSafe(sEmail) & "' AND custPasswd = '" & makeInputSafe(sPassword) & "' AND custID = " & visitorLoggedInCustomerID 
		Case "loose"
			Set mclsLogin = New clsLogin
			If Len(mclsLogin.ValidUserName(sEmail, sPassword)) = 0 Then
				plngCustID = VisitorLoggedInCustomerID
				If plngCustID = 0 Then
					plngCustID = -1
				Else
				
				End If
			Else
				plngCustID = -1
			End If
			Set mclsLogin = Nothing
		Case Else
			Set pclsCustomer = New clsCustomer
			Set pclsCustomer.Connection = cnn
			If pclsCustomer.LoadCustomerByEmail(sEmail) Then
				plngCustID = pclsCustomer.CustID
				If Len(CStr(plngCustID)) = 0 Then plngCustID = -1
			Else
				plngCustID = -1
			End If
			Set pclsCustomer = Nothing
	End Select

	customerAuth = plngCustID
	
End Function	'customerAuth

'-----------------------------------------------------------------------
' Generates a unique password
'-----------------------------------------------------------------------
Function generatePassword
	Dim sPassword,Random_Number_Min,Random_Number_Max
  	
  	Randomize
	Random_Number_Min = 10000000
	Random_Number_Max = 99999999

	sPassword = Int(((Random_Number_Max-Random_Number_Min+1) * Rnd) + Random_Number_Min)
	generatePassword = sPassword
End Function

'------------------------------------------------------------------
' Gets the InternetCash Merchant ID
'------------------------------------------------------------------
Function getICashMercID()
	Dim sLocalSQL, rsICash, iID
	
	sLocalSQL = "SELECT trnsmthdLogin FROM sfTransactionMethods WHERE trnsmthdName = 'InternetCash'"
	Set rsICash = CreateObject("ADODB.RecordSet")
	rsICash.Open sLocalSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	If rsICash.EOF or rsICash.BOF Then
		Response.Write "Error: No merchant ID set for Internet Cash in table sfTransactionMethods"
	Else
		iID = trim(rsICash.Fields("trnsmthdLogin"))
	End If
	
	closeobj(rsICash)
	getICashMercID = iID	
End Function

'------------------------------------------------------------------
' Gets shipping types
'------------------------------------------------------------------
Function getShipped(sProdID)
	Dim rsProdShipped, SQL
	SQL = "SELECT prodShipIsActive FROM sfProducts WHERE prodID = '" & makeInputSafe(sProdID) & "'"
	Set rsProdShipped = CreateObject("ADODB.Recordset")
	rsProdShipped.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	getShipped = rsProdShipped(0)
	closeObj(rsProdShipped)
End Function


'---------------------------------------------------------------
' To see if it is a saved cart customer
' Returns a boolean value
'---------------------------------------------------------------
Function CheckSavedCartCustomer(iCustID)
	Dim sSQL, rsTmp, bTruth
	sSQL = "SELECT custFirstName FROM sfCustomers WHERE custID=" & makeInputSafe(iCustID)
	
	bTruth = false
	
	Set rsTmp = CreateObject("ADODB.RecordSet")
		 rsTmp.Open sSQL,cnn,adOpenDynamic,adLockOptimistic,adCmdText
		 If NOT rsTmp.EOF Then
		 	If trim(rsTmp.Fields("custFirstName")) = "Saved Cart Customer" Then
		 		bTruth = true
		 		
		 	Else
		 		bTruth = false
		 	End If		
		 End If
		
	closeobj(rsTmp)	
	CheckSavedCartCustomer = bTruth
End Function

'--------------------------------------------------------

Function DeleteOrder(sID)

Dim rsDelete 'As New ADODB.Recordset
Dim rsDelete1 'As New ADODB.Recordset
Dim rsDelete2 'As New ADODB.Recordset
Dim rsDelete3 'As New ADODB.Recordset
Dim vOrderId 'As Variant
Dim sSql
On Error Resume Next
Set rsDelete = CreateObject("ADODB.RecordSet")
Set rsDelete1 = CreateObject("ADODB.RecordSet")
Set rsDelete2 = CreateObject("ADODB.RecordSet")
Set rsDelete3 = CreateObject("ADODB.RecordSet")

sSql = "SELECT * FROM sfOrders" _
        & " WHERE orderID = " & sID
    rsDelete.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
vOrderId = rsDelete("orderAddrId")
sSql = "SELECT * FROM sfOrderDetails WHERE odrdtOrderId = " & rsDelete.Fields("orderID")
    rsDelete2.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
'    '''''rsOrderCredit
sSql = "SELECT * FrOM sfCPayments WHERE payID = " & Trim(rsDelete.Fields("orderPayId"))
rsDelete3.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
rsDelete.Delete 
'rsDelete1.Delete adAffectCurrent
rsDelete2.Delete 
rsDelete3.Delete 

Set rsDelete = Nothing
Set rsDelete1 = Nothing
Set rsDelete2 = Nothing
Set rsDelete3 = Nothing
End Function

Function Reset_Shipping()
	Dim sSql,RstProd,rsttmpOrder
	Set rstProd = CreateObject("ADODB.RecordSet")
	Set rsttmpOrder = CreateObject("ADODB.RecordSet")
	sSql = "SELECT * FROM sfTmpOrderDetails" _
	        & " WHERE odrdttmpSessionID = " & SessionID
	rsttmpOrder.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
	
While rsttmpOrder.EOF =False
	 sSql = "SELECT prodShipIsActive FROM sfProducts " _
	        & " WHERE prodID = '" & rsttmpOrder("odrdttmpProductID") & "'"
	RstProd.Open sSql,cnn,adOpenStatic ,adLockReadOnly ,1
	If Not isNull(RstProd("prodShipIsActive")) then
	  rsttmpOrder("odrdttmpShipping") = RstProd("prodShipIsActive")
	Else
	  rsttmpOrder("odrdttmpShipping") = 0
	end if
	rsttmpOrder.Update 
	rsttmpOrder.MoveNext 
	rstProd.Close '#309
Wend	        

On error Resume Next
rsttmpOrder.Close 

Set rstProd =Nothing
Set rsttmpOrder = Nothing

End Function

Function get_Invalid_eMail(sData)
Dim iLoop,rst,sSql,sTemp,aCHK
aCHK = split(sData,",")
Set rst = CreateObject("ADODB.RecordSet")
sTemp = ""
For iLoop = 0 to uBound(aCHK)
  sSql = "Select custId From sfCustomers Where CustEmail = '" & aCHK(iLoop) & "'"
 If rst.State = 1 then rst.Close 
  rst.Open sSql,cnn,adOpenStatic ,adLockReadOnly ,1
   If rst.EOF AND rst.BOF Then
    sTemp = sTemp & aCHK(iLoop) & ","
   
   End if
     
next
 If right(stemp,1) = ";" then
   sTemp = left(len(sTemp)-1)
 End If   
 
get_Invalid_eMail = sTemp
on error resume next
rst.Close 
set rst = nothing
End Function

'--------------------------------------------------------------------------------------------------

Function HighlightSearchText(byVal strText, byVal strSearch, byVal strSearchType)

Dim pstrSearchFor
Dim pstrReplaceWith
Dim pstrOut
Dim i
Dim parySearchText

	If Len(cstrHighlightSearchTermClass) = 0 Then
		HighlightSearchText = strText
		Exit Function
	End If
	
	pstrOut = Trim(strText & "")
    'Response.Write "<fieldset><legend>HighlightSearchText</legend>Text: " & strText & "<br />Search: " & strSearch & "</fieldset>"
	If Len(strSearch) > 0 And strSearch <> "*" Then
		If strSearchType="Exact" Then				'ALL, ANY, Exact
			parySearchText = Array(strSearch)
		Else
			parySearchText = Split(strSearch, " ")
		End If
		
		For i = 0 To UBound(parySearchText)
			pstrSearchFor = "" & Trim(parySearchText(i) & "") & ""
			pstrReplaceWith = "{" & i & "}"
			'pstrOut = Replace(pstrOut, pstrSearchFor, pstrReplaceWith, 1, -1, 1)	'this is case insensitive but may display wrong case
			pstrOut = Replace(pstrOut, pstrSearchFor, pstrReplaceWith, 1, -1, 0)
		Next 'i

		For i = 0 To UBound(parySearchText)
			pstrSearchFor = "{" & i & "}"
			pstrReplaceWith = "<span class=highlightSearch>" & parySearchText(i) & "</span>"
			'pstrOut = Replace(pstrOut, pstrSearchFor, pstrReplaceWith, 1, -1, 1)
			pstrOut = Replace(pstrOut, pstrSearchFor, pstrReplaceWith, 1, -1, 0)
		Next 'i
	End If

	HighlightSearchText = pstrOut
	
End Function	'HighlightSearchText

'--------------------------------------------------------------------------------------------------

Sub writeCurrencyConverterOpeningScript
	If iConverion = 1 Then Response.Write  "<script language=""JavaScript"" src=""http://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>"
End Sub

'--------------------------------------------------------------------------------------------------

Sub writeCustomCurrency(byVal strCurrency)
	Response.Write customCurrency(strCurrency)
End Sub	'writeCustomCurrency

'--------------------------------------------------------------------------------------------------

Function customCurrency(byVal strCurrency)

Dim pstrOut

	If iConverion = 1 Then
		pstrOut = "<script>var ihomecurrency;ihomecurrency=OANDAconvert(1, " & chr(34) & CurrencyISO & chr(34) & ");ihomecurrency=ihomecurrency.substring(ihomecurrency.length-7,ihomecurrency.length-4);document.write(""" & FormatCurrency(strCurrency) & " = "" + OANDAconvert(" & strCurrency & ", " & chr(34) & CurrencyISO & chr(34) & ", ihomecurrency) + "" "" + ihomecurrency)</script>"
	Else
		pstrOut = FormatCurrency(strCurrency)
	End If
	
	customCurrency = pstrOut
	
End Function	'customCurrency

'--------------------------------------------------------------------------------------------------

Function getCountyList(byVal strSelected)

Dim pstrTemp
Dim i

	If loadCountyArray Then
		For i = 0 To UBound(maryCounties)
			If maryCounties(i)(0) = strSelected Then
				pstrTemp = pstrTemp & "<option value=""" & maryCounties(i)(0) & """ selected>" & maryCounties(i)(0) & "</option>" & vbcrlf
			Else
				pstrTemp = pstrTemp & "<option value=""" & maryCounties(i)(0) & """>" & maryCounties(i)(0) & "</option>" & vbcrlf
			End If
		Next
	Else
		pstrTemp = pstrTemp & "<option value="""">Error Loading Countries</option>" & vbcrlf
	End If
	
	getCountyList = pstrTemp

End Function	'getCountyList

'--------------------------------------------------------------------------------------------------

Function getStateList(byVal strSelected)

Dim pstrTemp
Dim i

	If loadStateArray Then
		For i = 0 To UBound(maryStates)
			If maryStates(i)(0) = strSelected Then
				pstrTemp = pstrTemp & "<option value=""" & maryStates(i)(0) & """ selected>" & maryStates(i)(1) & "</option>" & vbcrlf
			Else
				pstrTemp = pstrTemp & "<option value=""" & maryStates(i)(0) & """>" & maryStates(i)(1) & "</option>" & vbcrlf
			End If
		Next
	Else
		pstrTemp = pstrTemp & "<option value="""">Error Loading States</option>" & vbcrlf
	End If
	
	getStateList = pstrTemp

End Function	'getStateList

'--------------------------------------------------------------------------------------------------

Function getStateTaxRate(byVal strAbbr)

Dim pdblRate
Dim i

	pdblRate = 0
	If loadStateArray Then
		For i = 0 To UBound(maryStates)
			If maryStates(i)(0) = strAbbr Then
				If maryStates(i)(3) = 1 Then pdblRate = maryStates(i)(2)
				Exit For
			End If
		Next
	End If
	
	getStateTaxRate = pdblRate

End Function	'getStateTaxRate

'--------------------------------------------------------------------------------------------------

Function getCountryList(byVal strSelected, byVal strDefaultCountry)

Dim pstrTemp
Dim i
Dim pblnSelected
Dim pstrDefaultCountryName

	pblnSelected = False
	If loadCountryArray Then
		For i = 0 To UBound(maryCountry)
			If maryCountry(i)(0) = strDefaultCountry Then pstrDefaultCountryName = maryCountry(i)(1)
			If maryCountry(i)(0) = strSelected Then
				pstrTemp = pstrTemp & "<option value=""" & maryCountry(i)(0) & """ selected>" & maryCountry(i)(1) & "</option>" & vbcrlf
				pblnSelected = True
			Else
				pstrTemp = pstrTemp & "<option value=""" & maryCountry(i)(0) & """>" & maryCountry(i)(1) & "</option>" & vbcrlf
			End If
		Next
		
		If Not pblnSelected And Len(strDefaultCountry) > 0 Then
			pstrTemp = "<option value=""" & strDefaultCountry & """ selected>" & pstrDefaultCountryName & "</option>" & vbcrlf & pstrTemp
		End If
	
	End If
	
	getCountryList = pstrTemp

End Function	'getCountryList

'--------------------------------------------------------------------------------------------------

Function getCountryTaxRate(byVal strAbbr)

Dim pdblRate
Dim i

	pdblRate = 0
	If loadCountryArray Then
		For i = 0 To UBound(maryCountry)
			If maryCountry(i)(0) = strAbbr Then
				If maryCountry(i)(3) = 1 Then pdblRate = maryCountry(i)(2)
				Exit For
			End If
		Next
	End If
	
	getCountryTaxRate = pdblRate

End Function	'getCountryTaxRate

'--------------------------------------------------------------------------------------------------

Function getCountryFraudScore(byVal strAbbr)

Dim pdblRate
Dim i

	pdblRate = 0
	If loadCountryArray Then
		For i = 0 To UBound(maryCountry)
			If maryCountry(i)(0) = strAbbr Then
				If maryCountry(i)(3) = 1 Then pdblRate = maryCountry(i)(3)
				Exit For
			End If
		Next
	End If
	
	getCountryFraudScore = pdblRate

End Function	'getCountryFraudScore

'--------------------------------------------------------------------------------------------------

Function getCreditCardList(byVal strSelected)

Dim pstrTemp
Dim i

	If loadCreditCardArray Then
		For i = 0 To UBound(maryCreditCards)
			If maryCreditCards(i)(0) = strSelected Then
				pstrTemp = pstrTemp & "<option value=""" & maryCreditCards(i)(0) & """ selected>" & maryCreditCards(i)(1) & "</option>" & vbcrlf
			Else
				pstrTemp = pstrTemp & "<option value=""" & maryCreditCards(i)(0) & """>" & maryCreditCards(i)(1) & "</option>" & vbcrlf
			End If
		Next
	End If
	
	getCreditCardList = pstrTemp

End Function	'getCreditCardList

'--------------------------------------------------------------------------------------------------

Function getTransactionName(byVal lngTranID)

Dim pblnFound
Dim pstrTransactionName
Dim pstrTemp
Dim i

	pblnFound = False
	If loadCreditCardArray Then
		For i = 0 To UBound(maryCreditCards)
			If maryCreditCards(i)(0) = lngTranID Then
				pstrTransactionName = maryCreditCards(i)(1)
				pblnFound = True
				Exit For
			End If
		Next
	End If
	
	If Not pblnFound Then pstrTransactionName = Trim(getNameWithID("sfTransactionTypes", lngTranID, "transID", "transName", 0))
	
	getTransactionName = pstrTransactionName

End Function	'getTransactionName

'****************************************************************************************************************

Dim maryCounties

Function loadCountyArray()

Dim pblnResult
Dim pobjRS
Dim pstrSQL
Dim pstrTemp
Dim i

	pblnResult = False
	
	maryCounties = Application("CountyArray")
	If isArray(maryCounties) Then
		pblnResult = True
	Else
		pstrSQL = "SELECT County, LocaleAbbr FROM ssTaxTable ORDER BY LocaleAbbr, County"

		Set pobjRS = CreateObject("ADODB.RECORDSET")
		With pobjRS
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			
			'On Error Resume Next
			If Err.number = 0 Then
				ReDim maryCounties(.RecordCount - 1)
				For i = 1 To .RecordCount
					maryCounties(i-1) = Array(Trim(.Fields("County").Value & ""), Trim(.Fields("LocaleAbbr").Value & ""))
					.MoveNext
				Next
				pblnResult = True
				Application("CountyArray") = maryCounties
			Else
				Err.Clear
			End If
			.Close
			
		End With
		Set pobjRS = Nothing
	End If
	
	loadCountyArray = pblnResult
	
End Function	'loadCountyArray

'****************************************************************************************************************

Function loadStateArray()

Dim pblnResult
Dim pobjRS
Dim pstrSQL
Dim pstrTemp
Dim i

	pblnResult = False
	
	maryStates = Application("StateArray")
	If isArray(maryStates) Then
		pblnResult = True
	Else
		pstrSQL = "SELECT loclstAbbreviation, loclstName, loclstTax, loclstTaxIsActive FROM sfLocalesState WHERE loclstLocaleIsActive=1 ORDER BY loclstName"

		Set pobjRS = CreateObject("ADODB.RECORDSET")
		With pobjRS
			.CursorLocation = 2 'adUseClient
			On Error Resume Next
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			
			If Err.number = 0 Then
				ReDim maryStates(.RecordCount - 1)
				For i = 1 To .RecordCount
					maryStates(i-1) = Array(Trim(.Fields("loclstAbbreviation").Value & ""), Trim(.Fields("loclstName").Value & ""), Trim(.Fields("loclstTax").Value & ""), Trim(.Fields("loclstTaxIsActive").Value & ""))
					.MoveNext
				Next
				pblnResult = True
				Application("StateArray") = maryStates
			Else
				Err.Clear
			End If
			.Close
			On Error Goto 0
			
		End With
		Set pobjRS = Nothing
	End If
	
	loadStateArray = pblnResult
	
End Function	'loadStateArray

'****************************************************************************************************************

Function loadCountryArray()

Dim pblnResult
Dim pobjRS
Dim pstrSQL
Dim pstrTemp
Dim i

	pblnResult = False
	
	maryCountry = Application("CountryArray")
	If isArray(maryCountry) Then
		pblnResult = True
	Else
		pstrSQL = "SELECT loclctryAbbreviation, loclctryName, loclctryTax, loclctryTaxIsActive, loclctryFraudRating FROM sfLocalesCountry WHERE loclctryLocalIsActive=1 ORDER BY loclctryName"

		Set pobjRS = CreateObject("ADODB.RECORDSET")
		With pobjRS
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			
			On Error Resume Next
			If Err.number = 0 Then
				ReDim maryCountry(.RecordCount - 1)
				For i = 1 To .RecordCount
					maryCountry(i-1) = Array(Trim(.Fields("loclctryAbbreviation").Value & ""), Trim(.Fields("loclctryName").Value & ""), Trim(.Fields("loclctryTax").Value & ""), Trim(.Fields("loclctryTaxIsActive").Value & ""), Trim(.Fields("loclctryFraudRating").Value & ""))
					If Len(maryCountry(i-1)(3)) > 0 Then maryCountry(i-1)(3) = Abs(maryCountry(i-1)(3))
					If Len(maryCountry(i-1)(4)) = 0 Then maryCountry(i-1)(4) = 0
					.MoveNext
				Next
				pblnResult = True
				Application("CountryArray") = maryCountry
			Else
				Err.Clear
			End If
			.Close
			On Error Goto 0
			
		End With
		Set pobjRS = Nothing
	End If
	
	loadCountryArray = pblnResult
	
End Function	'loadCountryArray

'****************************************************************************************************************

Function loadCreditCardArray()

Dim pblnResult
Dim pobjRS
Dim pstrSQL
Dim pstrTemp
Dim i

	pblnResult = False
	
	maryCreditCards = Application("CreditCardArray")
	If isArray(maryCreditCards) Then
		pblnResult = True
	Else
		pstrSQL = "Select transID, transName From sfTransactionTypes WHERE transType = 'Credit Card' AND transIsActive = 1 ORDER BY transName"

		Set pobjRS = CreateObject("ADODB.RECORDSET")
		With pobjRS
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			
			On Error Resume Next
			If Err.number = 0 Then
				ReDim maryCreditCards(.RecordCount - 1)
				For i = 1 To .RecordCount
					maryCreditCards(i-1) = Array(Trim(.Fields("transID").Value & ""), Trim(.Fields("transName").Value & ""))
					.MoveNext
				Next
				pblnResult = True
				Application("CreditCardArray") = maryCreditCards
			Else
				Err.Clear
			End If
			.Close
			On Error Goto 0
			
		End With
		Set pobjRS = Nothing
	End If
	
	loadCreditCardArray = pblnResult
	
End Function	'loadCreditCardArray

'****************************************************************************************************************

Function loadManufacturerArray

Dim pblnResult
Dim pdicTemp
Dim i
Dim plngKey
Dim pobjRS
Dim pobjRSContent
Dim pstrKey
Dim pstrSQL
Dim vItem

	If Err.number <> 0 Then Err.Clear
	
	'Only need to load if not previously loaded
	If isArray(maryManufacturers) Then
		pblnResult = True
	Else
		'Application.Contents.Remove("sfManufacturers")
		If Not isCacheItemExpired("sfManufacturers") Then maryManufacturers = getFromCache("sfManufacturers")
		
		Set pdicTemp = CreateObject("scripting.dictionary")
		
		pstrSQL = "Select mfgID, mfgName, mfgNotes, mfgHttpAdd, mfgMetaTitle, mfgMetaDescription, mfgMetaKeywords, mfgDescription from sfManufacturers ORDER BY mfgName"
		Set pobjRS = GetRS(pstrSQL)
		With pobjRS
			If Not .EOF Then
				For i = 1 To .RecordCount
					plngKey = .Fields("mfgID").Value
					pstrKey = "mfg" &  plngKey
					If Not pdicTemp.Exists(pstrKey) Then
						vItem = Array(plngKey, _
									  Trim(.Fields("mfgName").Value), _
									  Trim(.Fields("mfgNotes").Value & ""), _
									  Trim(.Fields("mfgHttpAdd").Value & ""), _
									  Trim(.Fields("mfgMetaTitle").Value & ""), _
									  Trim(.Fields("mfgMetaDescription").Value & ""), _
									  Trim(.Fields("mfgMetaKeywords").Value & ""), _
									  Trim(.Fields("mfgDescription").Value & ""), _
									  False, _
									  False)
						pdicTemp.Add pstrKey, vItem
					End If
					.MoveNext
				Next 'i
			End If
			.Close
		End With	'pobjRS
		Set pobjRS = Nothing
		
		pstrSQL = "Select contentID, contentReferenceID, contentTitle, contentPageTitle, contentMetaDescription, contentMetaKeywords, contentPageName, mfgName FROM content LEFT JOIN sfManufacturers ON content.contentReferenceID = sfManufacturers.mfgID  Where contentContentType=3 And contentApprovedForDisplay <> 0 ORDER BY contentSortOrder, contentTitle"

		Set pobjRSContent = GetRS(pstrSQL)
		With pobjRSContent
			If Not .EOF Then
				For i = 1 To .RecordCount
					plngKey = .Fields("contentID").Value
					pstrKey = "cms" &  plngKey

					'if primary entry exists, then replace
					If pdicTemp.Exists("mfg" & .Fields("contentReferenceID").Value) Then
						vItem = pdicTemp("mfg" & .Fields("contentReferenceID").Value)
						vItem(9) = True
						pdicTemp("mfg" & .Fields("contentReferenceID").Value) = vItem
						'pdicTemp.Remove "mfg" & .Fields("contentReferenceID").Value
					End If
						
					vItem = Array(plngKey, _
								  Trim(.Fields("contentTitle").Value & ""), _
								  Trim(.Fields("contentMetaDescription").Value & ""), _
								  Trim(.Fields("contentPageName").Value & ""), _
								  Trim(.Fields("contentPageTitle").Value & ""), _
								  Trim(.Fields("contentMetaDescription").Value & ""), _
								  Trim(.Fields("contentMetaKeywords").Value & ""), _
								  Trim(.Fields("contentMetaDescription").Value & ""), _
								  True, _
								  False)
					If Len(vItem(1)) = 0 Then vItem(1) = vItem(4)
					If Len(vItem(1)) = 0 Then vItem(1) = Trim(.Fields("mfgName").Value & "")
					If Not pdicTemp.Exists(pstrKey) Then pdicTemp.Add pstrKey, vItem

					'ID: 0
					'Name: 1
					'Notes: 2
					'HttpAddr: 3
					'MetaTitle: 4
					'MetaDescription: 5
					'MetaKeywords: 6
					'Description: 7
					'IsContentPage: 8
					'HasContentReplacement: 9

					.MoveNext
				Next 'i
			End If
			.Close
		End With	'pobjRSContent
		Set pobjRSContent = Nothing

		'move dictionary to array
		ReDim maryManufacturers(pdicTemp.Count - 1)
		i = -1
		For Each vItem in pdicTemp
			i = i + 1
			maryManufacturers(i) = pdicTemp(vItem)
		Next

		'Now sort it
		maryManufacturers = bubbleSort2DArray(maryManufacturers, 1, "string", True)
		
		Call saveToCache("sfManufacturers", maryManufacturers, DateAdd("s", 3600, Now()))
		pblnResult = True
	End If

	loadManufacturerArray = pblnResult
	
End Function	'loadManufacturerArray

'****************************************************************************************************************

Function loadManufacturerArray_v0

Dim pblnResult
Dim i
Dim pobjRS
Dim pstrSQL

	If Err.number <> 0 Then Err.Clear
	
	'Only need to load if not previously loaded
	If isArray(maryManufacturers) Then
		pblnResult = True
	Else
		If Not isCacheItemExpired("sfManufacturers") Then maryManufacturers = getFromCache("sfManufacturers")
		
		'pstrSQL = "Select mfgID, mfgName, mfgNotes from sfManufacturers ORDER BY mfgName"
		pstrSQL = "Select mfgID, mfgName, mfgNotes, mfgHttpAdd, mfgMetaTitle, mfgMetaDescription, mfgMetaKeywords, mfgDescription from sfManufacturers ORDER BY mfgName"
		Set pobjRS = GetRS(pstrSQL)
		With pobjRS
			If Not .EOF Then
				ReDim maryManufacturers(.RecordCount - 1)
				For i = 1 To .RecordCount
					'maryManufacturers(i - 1) = Array(.Fields("mfgID").Value, Trim(.Fields("mfgName").Value), Trim(.Fields("mfgNotes").Value))
					maryManufacturers(i - 1) = Array(.Fields("mfgID").Value, Trim(.Fields("mfgName").Value), Trim(.Fields("mfgNotes").Value & ""), Trim(.Fields("mfgHttpAdd").Value & ""), Trim(.Fields("mfgMetaTitle").Value & ""), Trim(.Fields("mfgMetaDescription").Value & ""), Trim(.Fields("mfgMetaKeywords").Value & ""), Trim(.Fields("mfgDescription").Value & ""))
					.MoveNext
				Next 'i
			End If
			.Close
		End With	'pobjRS
		Set pobjRS = Nothing

		Call saveToCache("sfManufacturers", maryManufacturers, DateAdd("s", 3600, Now()))
		pblnResult = True
	End If

	loadManufacturerArray = pblnResult
	
End Function	'loadManufacturerArray

'****************************************************************************************************************

Function stripHTML(byVal strHTML)
'Strips the HTML tags from strHTML

Dim objRegExp
Dim strOutput
	
	Set objRegExp = New Regexp

	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "<(.|\n)+?>"

	'Replace all HTML tag matches with the empty string
	strOutput = objRegExp.Replace(strHTML & "", "")

	'Replace all < and > with &lt; and &gt;
	strOutput = Replace(strOutput, "<", "&lt;")
	strOutput = Replace(strOutput, ">", "&gt;")

	stripHTML = strOutput    'Return the value of strOutput

	Set objRegExp = Nothing
	
End Function 'stripHTML

'****************************************************************************************************************

Function loadVendorArray

Dim pblnResult
Dim i
Dim pobjRS
Dim pstrSQL

	If Err.number <> 0 Then Err.Clear
	
	'Only need to load if not previously loaded
	If isArray(maryVendors) Then
		pblnResult = True
	Else
		If Not isCacheItemExpired("sfVendors") Then maryVendors = getFromCache("sfVendors")
		
		'pstrSQL = "Select vendID, vendName, vendNotes from sfVendors ORDER BY vendName"
		pstrSQL = "Select vendID, vendName, vendNotes, vendHttpAddr, vendMetaTitle, vendMetaDescription, vendMetaKeywords, vendDescription from sfVendors ORDER BY vendName"
		Set pobjRS = GetRS(pstrSQL)
		With pobjRS
			If Not .EOF Then
				ReDim maryVendors(.RecordCount - 1)
				For i = 1 To .RecordCount
					maryVendors(i - 1) = Array(.Fields("vendID").Value, _
											   Trim(.Fields("vendName").Value), _
											   Trim(.Fields("vendNotes").Value & ""), _
											   Trim(.Fields("vendHttpAddr").Value & ""), _
											   Trim(.Fields("vendMetaTitle").Value & ""), _
											   Trim(.Fields("vendMetaDescription").Value & ""), _
											   Trim(.Fields("vendMetaKeywords").Value & ""), _
											   Trim(.Fields("vendDescription").Value & ""), _
											   False _
											  )
					.MoveNext
				Next 'i
			End If
			.Close
		End With	'pobjRS
		Set pobjRS = Nothing

		Call saveToCache("sfVendors", maryVendors, DateAdd("s", 3600, Now()))
		pblnResult = True
	End If

	loadVendorArray = pblnResult
	
End Function	'loadVendorArray

'****************************************************************************************************************

Function hasManufacturers(byVal excludeNoManufacturer)

	If loadManufacturerArray Then
		If UBound(maryManufacturers) > 1 Then
			hasManufacturers = True
		ElseIf UBound(maryManufacturers) = 0 Then
			If excludeNoManufacturer Then
				hasManufacturers = CBool(maryManufacturers(0)(0) <> 1)
			Else
				hasManufacturers = True
			End If
		Else
			hasManufacturers = False
		End If
	Else
		hasManufacturers = False
	End If

End Function	'hasManufacturers

'****************************************************************************************************************

Function hasVendors(byVal excludeNoVendor)

	If loadVendorArray Then
		If UBound(maryVendors) > 1 Then
			hasVendors = True
		ElseIf UBound(maryVendors) = 0 Then
			If excludeNoVendor Then
				hasVendors = CBool(maryVendors(0)(0) <> 1)
			Else
				hasVendors = True
			End If
		Else
			hasVendors = False
		End If
	Else
		hasVendors = False
	End If

End Function	'hasVendors

'****************************************************************************************************************

Function getMfgVendItem(byVal mfgVendID, byVal strItem, byVal blnMfg)

Dim i
Dim plngNumItems
Dim pstrTemp
Dim pblnFound

	If Len(mfgVendID) > 0 And isNumeric(mfgVendID) Then
		mfgVendID = CLng(mfgVendID)
	Else
		Exit Function
	End If
	
	pblnFound = False
	If blnMfg Then
		If loadManufacturerArray Then
			plngNumItems = UBound(maryManufacturers)
			For i = 0 To plngNumItems
				If maryManufacturers(i)(0) = mfgVendID And Not maryManufacturers(i)(8) Then
					Select Case strItem
						Case "Name": pstrTemp = maryManufacturers(i)(1)
						Case "Notes": pstrTemp = maryManufacturers(i)(2)
						Case "HttpAddr": pstrTemp = maryManufacturers(i)(3)
						Case "MetaTitle": pstrTemp = maryManufacturers(i)(4)
						Case "MetaDescription": pstrTemp = maryManufacturers(i)(5)
						Case "MetaKeywords": pstrTemp = maryManufacturers(i)(6)
						Case "Description": pstrTemp = maryManufacturers(i)(7)
						Case "URL":
							If Len(maryManufacturers(i)(3)) = 0 Then
								pstrTemp = Replace(cstrURLTemplate_Manufacturer, "{ID}", maryManufacturers(i)(0))
							Else
								pstrTemp = maryManufacturers(i)(3)
							End If
					End Select
					pblnFound = True
				End If
			Next 'i
		Else
			'Response.Write "Failed to load manufacturers<br />"
		End If	'loadManufacturerArray
		
		If Not pblnFound Then
			Select Case strItem
				Case "Name": pstrTemp = getNameWithID("sfManufacturers", mfgVendID, "mfgID", "mfgName", 0)
				Case "Notes": pstrTemp = getNameWithID("sfManufacturers", mfgVendID, "mfgID", "mfgNotes", 0)
				Case "HttpAddr": pstrTemp = getNameWithID("sfManufacturers", mfgVendID, "mfgID", "mfgHttpAdd", 0)
				Case "MetaTitle": pstrTemp = getNameWithID("sfManufacturers", mfgVendID, "mfgID", "mfgMetaTitle", 0)
				Case "MetaDescription": pstrTemp = getNameWithID("sfManufacturers", mfgVendID, "mfgID", "mfgMetaDescription", 0)
				Case "MetaKeywords": pstrTemp = getNameWithID("sfManufacturers", mfgVendID, "mfgID", "mfgMetaKeywords", 0)
				Case "Description": pstrTemp = getNameWithID("sfManufacturers", mfgVendID, "mfgID", "mfgDescription", 0)
				Case "URL":	pstrTemp = Replace(cstrURLTemplate_Manufacturer, "{ID}", mfgVendID)
			End Select
		End If
	Else
		If loadVendorArray Then
			plngNumItems = UBound(maryVendors)
			For i = 0 To plngNumItems
				If maryVendors(i)(0) = mfgVendID Then
					Select Case strItem
						Case "Name": pstrTemp = maryVendors(i)(1)
						Case "Notes": pstrTemp = maryVendors(i)(2)
						Case "HttpAddr": pstrTemp = maryVendors(i)(3)
						Case "MetaTitle": pstrTemp = maryVendors(i)(4)
						Case "MetaDescription": pstrTemp = maryVendors(i)(5)
						Case "MetaKeywords": pstrTemp = maryVendors(i)(6)
						Case "Description": pstrTemp = maryVendors(i)(7)
						Case "URL":
							If Len(maryVendors(i)(3)) = 0 Then
								pstrTemp = Replace(cstrURLTemplate_Vendor, "{ID}", maryVendors(i)(0))
							Else
								pstrTemp = maryVendors(i)(3)
							End If
					End Select
					pblnFound = True
				End If
			Next 'i
		End If	'loadVendorArray
		
		If Not pblnFound Then
			Select Case strItem
				Case "Name": pstrTemp = getNameWithID("sfVendors", mfgVendID, "vendID", "vendName", 0)
				Case "Notes": pstrTemp = getNameWithID("sfVendors", mfgVendID, "vendID", "vendNotes", 0)
				Case "HttpAddr": pstrTemp = getNameWithID("sfVendors", mfgVendID, "vendID", "vendHttpAddr", 0)
				Case "MetaTitle": pstrTemp = getNameWithID("sfVendors", mfgVendID, "vendID", "vendMetaTitle", 0)
				Case "MetaDescription": pstrTemp = getNameWithID("sfVendors", mfgVendID, "vendID", "vendMetaDescription", 0)
				Case "MetaKeywords": pstrTemp = getNameWithID("sfVendors", mfgVendID, "vendID", "vendMetaKeywords", 0)
				Case "Description": pstrTemp = getNameWithID("sfVendors", mfgVendID, "vendID", "vendDescription", 0)
				Case "URL":	pstrTemp = Replace(cstrURLTemplate_Vendor, "{ID}", mfgVendID)
			End Select
		End If
	End If	'blnMfg
	
	getMfgVendItem = pstrTemp

End Function	'getMfgVendItem

'****************************************************************************************************************

Function getCategoryItem(byVal catID, byVal strItem)

Dim i
Dim paryCategories
Dim plngNumItems
Dim pstrTemp

	If Len(catID) > 0 And isNumeric(catID) Then
		catID = CLng(catID)
	Else
		Exit Function
	End If
	
	paryCategories = getFromCache("ssCategorySearch")
	If isArray(paryCategories) Then
		plngNumItems = UBound(paryCategories, 1)
		For i = 0 To plngNumItems
			If paryCategories(i, en_CatFields_uid) = catID Then
				Select Case strItem
					Case "Name": pstrTemp = paryCategories(i, en_CatFields_Name)
					Case "HttpAddr": pstrTemp = paryCategories(i, en_CatFields_URL)
					Case "MetaTitle": pstrTemp = paryCategories(i, en_CatFields_Name)
					Case "MetaDescription": pstrTemp = paryCategories(i, en_CatFields_Description)
					Case "MetaKeywords": pstrTemp = paryCategories(i, en_CatFields_Description)
					Case "Description": pstrTemp = paryCategories(i, en_CatFields_Description)
					Case "URL":
						If Len(paryCategories(i, en_CatFields_URL)) = 0 Then
							pstrTemp = Replace(cstrURLTemplate_Manufacturer, "{ID}", paryCategories(i, en_CatFields_URL))
						Else
							pstrTemp = paryCategories(i, en_CatFields_URL)
						End If
				End Select
			End If
		Next 'i
	End If	'isArray(paryCategories)
	
	getCategoryItem = pstrTemp

End Function	'getCategoryItem

'****************************************************************************************************************

Function getPageFragmentByID(byVal cmsID)

Dim pobjCmd
Dim pobjRS
Dim pstrContent

	If Len(cmsID) > 0 And isNumeric(cmsID) Then
		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = "SELECT contentContent, contentApprovedForDisplay FROM content WHERE contentID=?"
			Set .ActiveConnection = cnn

			.Parameters.Append .CreateParameter("contentID", adInteger, adParamInput, 4, cmsID)

  			Set pobjRS = .Execute
  			If pobjRS.EOF Then
				If isAdminLoggedIn Then Response.Write "<div class=""adminCMS""><font color=red><strong>Warning: " & cmsID & " content does not exist!</strong></font></div>"
  			Else
  				Call DisplayCMSEditLink(cmsID, Trim(pobjRS.Fields("contentContent").Value & ""), ConvertToBoolean(pobjRS.Fields("contentApprovedForDisplay").Value, False))
  				If ConvertToBoolean(pobjRS.Fields("contentApprovedForDisplay").Value, False) Then pstrContent = Trim(pobjRS.Fields("contentContent").Value & "")
  			End If
  			closeObj(pobjRS)
		End With
		Set pobjCmd = Nothing
	
	End If

	getPageFragmentByID = pstrContent

End Function	'getPageFragmentByID

'****************************************************************************************************************

Function getPageFragmentByKey(byVal strCMSKey)

Dim pobjCmd
Dim pobjRS
Dim pstrContent
Dim pstrCacheKey

	If Len(strCMSKey) > 0 Then
		pstrCacheKey = "cmsFragment_" & strCMSKey
		pstrContent = getFromCache(pstrCacheKey)
		If Len(pstrContent) = 0 Then
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = "SELECT contentID, contentContent, contentApprovedForDisplay FROM content WHERE contentMetaCustom1=?"
				Set .ActiveConnection = cnn

				.Parameters.Append .CreateParameter("key", adVarChar, adParamInput, Len(strCMSKey), strCMSKey)

  				Set pobjRS = .Execute
  				If pobjRS.EOF Then
  					'If vDebug = 1 Then Response.Write "Either the recordset doesn't exit or the field name is not typed correctly :<br />" & sLocalSQL			
					If isAdminLoggedIn Then Response.Write "<div class=""adminCMS""><font color=red><strong>Warning: " & strCMSKey & " content does not exist!</strong></font><br />Please create content using (" & strCMSKey & ") as the page reference. <a class=""adminCMS"" href=""ssl/ssAdmin/ssCMS_PageFragmentAdmin.asp?Action=FixMissingFragment&amp;ViewID=" & strCMSKey & """>Fix</a></div>"
  				Else
					If isAdminLoggedIn Then
						pstrContent = "<div class=""adminCMS""><a class=""adminCMS"" href=""ssl/ssAdmin/ssCMS_PageFragmentAdmin.asp?Action=viewItem&amp;ViewID=" & pobjRS.Fields("contentID").Value & """>Edit Content</a>" & Trim(pobjRS.Fields("contentContent").Value & "")
  						If Not ConvertToBoolean(pobjRS.Fields("contentApprovedForDisplay").Value, False) Then pstrContent = pstrContent & "<br /><font color=red><strong>Not Approved for Display!</strong></font>"
						pstrContent = pstrContent & "</div>"
  					Else
  						If ConvertToBoolean(pobjRS.Fields("contentApprovedForDisplay").Value, False) Then pstrContent = Trim(pobjRS.Fields("contentContent").Value & "")
  					End If
					Call saveToCache(pstrCacheKey, pstrContent, DateAdd("s", 600, Now()))
  				End If
  				closeObj(pobjRS)
			End With
		End If	'Len(pstrContent) = 0
		Set pobjCmd = Nothing
	
	End If	'Len(strCMSKey) > 0

	getPageFragmentByKey = pstrContent

End Function	'getPageFragmentByKey

'*************************************************************************************************************************

Function BreadCrumbsTrail(byVal aryTrail)

Dim i
Dim pstrHRef
Dim pstrLinkText
Dim pstrTitle
Dim pstrBreadcrumbsText
Dim pstrCrumb

	If isArray(aryTrail) Then
		For i = 0 To UBound(aryTrail) Step 3
			pstrCrumb = ""
			pstrHRef = ""
			pstrTitle = ""
			pstrLinkText = aryTrail(i)
			
			If Len(pstrLinkText) > 0 Then
				If UBound(aryTrail) >= i+1 Then pstrHRef = aryTrail(i+1)
				If UBound(aryTrail) >= i+2 Then pstrTitle = aryTrail(i+2)
				If Len(pstrTitle) = 0 Then pstrTitle = pstrLinkText
				pstrCrumb = "<span class=""categoryTrail"">" & pstrLinkText & "</span>"
				
				If Len(pstrHRef) > 0 Then
					pstrCrumb = "<a href=""" & pstrHRef & """ class=""categoryTrail"" title=""" & pstrTitle & """>" & pstrCrumb & "</a>"
				End If
			End If

			If Len(pstrBreadcrumbsText) > 0 Then
				pstrBreadcrumbsText = pstrBreadcrumbsText & "&nbsp;>>&nbsp;" & pstrCrumb
			Else
				pstrBreadcrumbsText = pstrCrumb
			End If
		Next 'i
	End If
	
	BreadCrumbsTrail = "<div class=""clsCategoryTrail"" id=""divCategoryTrail"">" & pstrBreadcrumbsText & "</div>"

End Function 'BreadCrumbsTrail

%>
