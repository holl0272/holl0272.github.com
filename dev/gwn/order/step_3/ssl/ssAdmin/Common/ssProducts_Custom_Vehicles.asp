<%
'********************************************************************************
'*   Common Support File For StoreFront 6.0 add-ons
'*   Custom Product Management Routines for Vehicle
'*
'*   This file must be included from ssProducts_Common
'*
'*   File Version:		1.00.001
'*   Revision Date:		September 19, 2005
'*
'*   1.00.001 - September 19, 2005
'*	 ' Initial Release
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************

'**********************************************************
'	Developer notes
'**********************************************************

'Note: Functions below are hooked from ssProducts_Custom functions of similar name

'Function DeleteProduct_Custom(byVal lngUID)
'Function DeleteAllProducts_Custom()
'Function updateProducts_Custom(byVal strProductID)

'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------

'***********************************************************************************************
'	Custom section for Vehicle
'***********************************************************************************************

Function DeleteProduct_Custom_Vehicle(byVal lngUID)

Dim pblnResult
Dim pstrSQL

'On Error Resume Next

	If len(lngUID) = 0 Then
		DeleteProduct_Custom = False
		Exit Function
	End If
	
	pblnResult = True
	
	pstrSQL = "Delete From ProductVehicles Where ProductID=" & lngUID
	cnn.Execute pstrSQL,,128

    If (Err.Number = 0) Then

    Else
        Call addMessageItem("Error deleting the product catalog: " & Err.Description, True)
        pblnResult = False
    End If
    
    DeleteProduct_Custom_Vehicle = pblnResult
    
End Function    'DeleteProduct_Custom_Vehicle

'***********************************************************************************************

Function DeleteProductVehicle_Custom_Vehicle(byVal lngUID)

Dim pblnResult
Dim pstrSQL

'On Error Resume Next

	If len(lngUID) = 0 Then
		DeleteProduct_Custom = False
		Exit Function
	End If
	
	pblnResult = True
	
	pstrSQL = "Delete From ProductVehicles Where uid=" & lngUID
	cnn.Execute pstrSQL,,128

    If (Err.Number = 0) Then
        Call addMessageItem("Vehicle assignment deleted for this product (" & lngUID & ")", False)
    Else
        Call addMessageItem("Error deleting the product catalog: " & Err.Description, True)
        pblnResult = False
    End If
    
    DeleteProductVehicle_Custom_Vehicle = pblnResult
    
End Function    'DeleteProductVehicle_Custom_Vehicle

'***********************************************************************************************

Function DeleteAllProducts_Custom_Vehicle()

Dim pblnResult
Dim pstrSQL

'On Error Resume Next

	pblnResult = True
	
	pstrSQL = "Delete From ProductVehicles"
	cnn.Execute pstrSQL,,128

    If (Err.Number = 0) Then

    Else
        Call addMessageItem("Error deleting the product catalog: " & Err.Description, True)
        pblnResult = False
    End If
    
	DeleteAllProducts_Custom_Vehicle = pblnResult	

End Function    'DeleteAllProducts_Custom_Vehicle

'***********************************************************************************************

Function createVehicle(byVal strMake, byVal strModel, byVal strYear)

Dim pobjCmd
Dim pobjRS

	'Check Data
	If Len(strMake) = 0 Or Len(strModel) = 0 Or Len(strYear) = 0 Then
		createVehicle = -1
		Exit Function
	End If
	
	'On Error Resume Next
		
	Set pobjCmd  = Server.CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Insert Into Vehicles ([Make], [Model], [Year]) Values (?, ?, ?)"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("Make", adVarChar, adParamInput, 100, strMake)
		.Parameters.Append .CreateParameter("Model", adVarChar, adParamInput, 100, strModel)
		.Parameters.Append .CreateParameter("Year", adVarChar, adParamInput, 50, strYear)
		.Execute , , adExecuteNoRecords
		
		.Commandtext = "Select uid From Vehicles Where Make=? And Model=? And Year=?"
		Set pobjRS = .Execute
		
		If pobjRS.EOF Then
			createVehicle = -1
		Else
			createVehicle = pobjRS.Fields(0).Value
		End If

		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing
	
End Function	'createVehicle

'***********************************************************************************************

Function getAssignedVehicles(byVal strProductID, byRef aryVehicles)

Dim pobjRS
Dim pstrSQL

	'On Error Resume Next
		
	pstrSQL = "SELECT ProductVehicles.uid AS ProductVehicleID, Vehicles.Make, Vehicles.Model, Vehicles.Year" _
			& " FROM ProductVehicles LEFT JOIN Vehicles ON ProductVehicles.VehicleID = Vehicles.uid" _
			& " WHERE ProductVehicles.ProductID=" & strProductID _
			& " ORDER BY Vehicles.Make, Vehicles.Model, Vehicles.Year"
	Set pobjRS = GetRS(pstrSQL)
		
	If pobjRS.EOF Then
		getAssignedVehicles = False
	Else
		aryVehicles = pobjRS.GetRows()
		getAssignedVehicles = True
	End If
	pobjRS.Close
	Set pobjRS = Nothing
	
End Function	'getAssignedVehicles

'***********************************************************************************************

Function getVehicle(byVal strMake, byVal strModel, byVal strYear)

Dim pobjCmd
Dim pobjRS

	'Check Data
	If Len(strMake) = 0 Or Len(strModel) = 0 Or Len(strYear) = 0 Then
		getVehicle = -1
		Exit Function
	End If
	
	'On Error Resume Next
		
	Set pobjCmd  = Server.CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select uid From Vehicles Where Make=? And Model=? And Year=?"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("Make", adVarChar, adParamInput, 100, strMake)
		.Parameters.Append .CreateParameter("Model", adVarChar, adParamInput, 100, strModel)
		.Parameters.Append .CreateParameter("Year", adVarChar, adParamInput, 50, strYear)
		Set pobjRS = .Execute
		
		If pobjRS.EOF Then
			getVehicle = -1
		Else
			getVehicle = pobjRS.Fields(0).Value
		End If

		If False Then
			Response.Write "<fieldset><legend>getVehicle</legend>"
			Response.Write "strMake: " & strMake & "<br />"
			Response.Write "strModel: " & strModel & "<br />"
			Response.Write "strYear: " & strYear & "<br />"
			Response.Write "pobjRS.EOF: " & pobjRS.EOF & "<br />"
			Response.Write "</fieldset>"
		End If
	
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing
		
End Function	'getVehicle

'***********************************************************************************************

Function getVehicles(byRef aryVehicles)

Dim pobjRS
Dim pstrSQL

	'On Error Resume Next
		
	pstrSQL = "Select uid, Make, Model, Year From Vehicles Order By Make, Model, Year"
	Set pobjRS = GetRS(pstrSQL)
		
	If pobjRS.EOF Then
		getVehicles = False
	Else
		aryVehicles = pobjRS.GetRows()
		getVehicles = True
	End If
	pobjRS.Close
	Set pobjRS = Nothing
	
End Function	'getVehicles

'***********************************************************************************************

Function setProductVehicle(byVal strProductID, byVal strMake, byVal strModel, byVal strYear)

Dim plngVehicleID
Dim pobjCmd
Dim pobjRS
Dim pstrMakeModelYear

	strMake = Trim(strMake)
	strModel = Trim(strModel)
	strYear = Trim(strYear)
	
	'Check Data
	If Len(strMake) = 0 Or Len(strModel) = 0 Or Len(strYear) = 0 Or Len(strProductID) = 0 Then
		setProductVehicle = False
		Exit Function
	End If
	
	pstrMakeModelYear = strMake & "~" & strModel & "~" & strYear
	plngVehicleID = getVehicle(strMake, strModel, strYear)
	If plngVehicleID = -1 Then plngVehicleID = createVehicle(strMake, strModel, strYear)
	If plngVehicleID = -1 Then
		setProductVehicle = False
		Exit Function
	End If
	
	Set pobjCmd  = Server.CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select uid From ProductVehicles Where ProductID=? And VehicleID=?"
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("ProductID", adInteger, adParamInput, 4, strProductID)
		.Parameters.Append .CreateParameter("VehicleID", adInteger, adParamInput, 4, plngVehicleID)
		Set pobjRS = .Execute
		
		If pobjRS.EOF Then
			.Commandtext = "Insert Into ProductVehicles (ProductID, VehicleID) Values (?, ?)"
			.Execute , , 128	'adExecuteNoRecords
		End If

		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing

	setProductVehicle = CBool(Err.number = 0)
		
End Function	'setProductVehicle

'***********************************************************************************************

Function updateProducts_Custom_Vehicle(byVal strProductID)

Dim pblnResult
Dim pstrSQL
Dim paryDeletions
Dim paryProductVehicles
Dim i
Dim pstrMake
Dim pstrModel
Dim pstrVehicle
Dim pstrYear

'On Error Resume Next

	pblnResult = True
	
	'Inbound parameters
	'productVehiclesToDelete
	'productVehiclesIsDirty
	'productVehicles
	
	Dim productVehiclesToDelete:	productVehiclesToDelete = LoadRequestValue("productVehiclesToDelete")
	Dim productVehiclesIsDirty:		productVehiclesIsDirty = LoadRequestValue("productVehiclesIsDirty")
	Dim productVehicles:			productVehicles = LoadRequestValue("productVehicles")
	
	If False Then
		Response.Write "<fieldset><legend>updateProducts_Custom_Vehicle</legend>"
		Response.Write "productVehiclesIsDirty: " & productVehiclesIsDirty & "<br />"
		Response.Write "productVehiclesToDelete: " & productVehiclesToDelete & "<br />"
		Response.Write "productVehicles: " & productVehicles & "<br />"
		Response.Write "strProductID: " & strProductID & "<br />"
		Response.Write "</fieldset>"
	End If
	
	If Len(productVehiclesToDelete) > 0 Then
		paryDeletions = Split(productVehiclesToDelete, ",")
		For i = 0 To UBound(paryDeletions)
			Call DeleteProductVehicle_Custom_Vehicle(Trim(paryDeletions(i)))
		Next 'i
	End If
	
	If Len(productVehicles) > 0 Then
		paryProductVehicles = Split(productVehicles, ",")
		For i = 0 To UBound(paryProductVehicles)
			pstrVehicle = vehicleKey(pstrMake, pstrModel, pstrYear, Trim(paryProductVehicles(i)))

			If setProductVehicle(strProductID, pstrMake, pstrModel, pstrYear) Then
				Call addMessageItem("Created vehicle <em>" & pstrVehicle & "</em>", False)
			End If
		Next 'i
	End If
	
	Call updateProducts_Custom_BulkUpload_Vehicle
	
	updateProducts_Custom_Vehicle = pblnResult	

End Function    'updateProducts_Custom_Vehicle

'***********************************************************************************************

Function updateProducts_Custom_BulkUpload_Vehicle()

Dim i
Dim paryProductCodes
Dim paryVehicle
Dim paryYears
Dim p_lngProductUID
Dim pstrProductCode
Dim pstrProductCodes
Dim pstrMake
Dim pstrModel
Dim pstrVehicle
Dim pstrVehicles
Dim pstrYear
Dim pstrYear_Start
Dim pstrYear_End
Dim Year
Dim plngSplitCount

	pstrProductCodes = LoadRequestValue("bulkLoadProductIDs")
	pstrVehicles = LoadRequestValue("bulkLoadVehicles")
	
	'decipher the vehicle: ex. Mercury/Mariner 100 1989-1991
	paryVehicle = Split(pstrVehicles, " ")
	plngSplitCount = UBound(paryVehicle)
	If plngSplitCount = 2 Then
		pstrMake = paryVehicle(0)
		pstrModel = paryVehicle(1)
		pstrYear = paryVehicle(2)
	ElseIf plngSplitCount > 2 Then
		pstrMake = paryVehicle(0)
		pstrModel = paryVehicle(1)
		For i = 2 To plngSplitCount - 1
			pstrModel = pstrModel & " " & paryVehicle(i)
		Next 'i
		pstrYear = paryVehicle(plngSplitCount)
	Else
		If UBound(paryVehicle) >= 0 Then pstrMake = paryVehicle(0)
		If UBound(paryVehicle) >= 1 Then pstrModel = paryVehicle(1)
		If UBound(paryVehicle) >= 2 Then pstrYear = paryVehicle(2)
	End If
	
	If InStr(1, pstrYear, "-") > 0 Then
		paryYears = Split(pstrYear, "-")
		pstrYear_Start = CLng(paryYears(0))
		pstrYear_End = CLng(paryYears(1))
	Else
		pstrYear_Start = CLng(pstrYear)
		pstrYear_End = CLng(pstrYear)
	End If
	
	
	If Len(pstrProductCodes) > 0 And Len(pstrVehicles) > 0 Then
		paryProductCodes = Split(pstrProductCodes, vbcrlf)
		For i = 0 To UBound(paryProductCodes)
			pstrProductCode = Trim(paryProductCodes(i))
			p_lngProductUID = getProductUIDByCode(pstrProductCode)
			If p_lngProductUID <> -1 Then
				For Year = pstrYear_Start To pstrYear_End
					pstrVehicle = vehicleKey(pstrMake, pstrModel, Year, "")
					If setProductVehicle(p_lngProductUID, pstrMake, pstrModel, Year) Then
						Call addMessageItem("Created vehicle <em>" & pstrVehicle & "</em> for product <em>" & pstrProductCode & "</em>", False)
					End If
				Next 'Year
			Else
					Call addMessageItem("Unable to create vehicle <em>" & pstrVehicle & "</em> for product <em>" & pstrProductCode & "</em>, product does not exist", True)
			End If
		Next 'i
	End If
	
End Function	'updateProducts_Custom_BulkUpload_Vehicle

'***********************************************************************************************

Function vehicleKey(byRef strMake, byRef strModel, byRef strYear, byVal strVehicleKey)

Dim paryVehicle
Dim pstrVehicle

	If Len(strVehicleKey) = 0 Then
		pstrVehicle = strMake & "_" & strModel & "_" & strYear
	Else
		pstrVehicle = strVehicleKey
		paryVehicle = Split(pstrVehicle, "_")
		If UBound(paryVehicle) >= 0 Then strMake = paryVehicle(0)
		If UBound(paryVehicle) >= 1 Then strModel = paryVehicle(1)
		If UBound(paryVehicle) >= 2 Then strYear = paryVehicle(2)
	End If
	
	vehicleKey = pstrVehicle

End Function	'vehicleKey

%>
