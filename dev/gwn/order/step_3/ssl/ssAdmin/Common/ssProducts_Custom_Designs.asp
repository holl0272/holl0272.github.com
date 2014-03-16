<%
'********************************************************************************
'*   Common Support File For StoreFront 6.0 add-ons
'*   Custom Product Management Routines for Design
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
'Function updateProducts_Custom(byVal strProductID)

'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------

'***********************************************************************************************
'	Custom section for Design
'***********************************************************************************************

Function DeleteProduct_Custom_Design(byVal lngUID)

Dim pblnResult
Dim pstrSQL

'On Error Resume Next

	If len(lngUID) = 0 Then
		DeleteProduct_Custom_Design = False
		Exit Function
	End If
	
	pblnResult = True
	
	pstrSQL = "Delete From contentProductAssignments Where contentProductAssignmentProductID=" & lngUID
	cnn.Execute pstrSQL,,128

    If (Err.Number = 0) Then

    Else
        Call addMessageItem("Error deleting the product catalog: " & Err.Description, True)
        pblnResult = False
    End If
    
    DeleteProduct_Custom_Design = pblnResult
    
End Function    'DeleteProduct_Custom_Design

'***********************************************************************************************

Function updateProducts_Custom_Design(byVal lngProductUID)

Dim paryOptions
Dim pblnDeletedOption
Dim pblnResult
Dim pobjRSSelections
Dim pstrOptions
Dim pstrSQL
	
	pstrOptions = Request.Form("itemOption")
	paryOptions = Split(pstrOptions, ",")
	pblnResult = True

	pstrSQL = "Select contentProductAssignmentID, contentProductAssignmentProductID, contentProductAssignmentContentID From contentProductAssignments Where contentProductAssignmentProductID=" & lngProductUID & " Order By contentProductAssignmentContentID Asc"
	Set pobjRSSelections = GetRS(pstrSQL)
	For i = 0 To UBound(paryOptions)
		pstrOptions = Trim(paryOptions(i))
		pobjRSSelections.Filter = "contentProductAssignmentContentID=" & pstrOptions
		If pobjRSSelections.EOF Then
			pstrSQL = "Insert Into contentProductAssignments (contentProductAssignmentProductID, contentProductAssignmentContentID) Values (" & lngProductUID & "," & pstrOptions & ")"
			cnn.Execute pstrSQL,,128
		End If
	Next 'i
	
	'Now check for deletions
	pobjRSSelections.Filter = ""
	If pobjRSSelections.RecordCount > 0 Then pobjRSSelections.MoveFirst
	Do While Not pobjRSSelections.EOF
		pblnDeletedOption = True
		For i = 0 To UBound(paryOptions)
			pstrOptions = Trim(paryOptions(i))
			If CStr(pstrOptions) = CStr(pobjRSSelections.Fields("contentProductAssignmentContentID").Value) Then
				pblnDeletedOption = False
				Exit For
			End If
		Next 'i
		If pblnDeletedOption Then
			pstrSQL = "Delete From contentProductAssignments Where contentProductAssignmentID=" & pobjRSSelections.Fields("contentProductAssignmentID").Value
			cnn.Execute pstrSQL,,128
		End If
		pobjRSSelections.MoveNext
	Loop
	Call ReleaseObject(pobjRSSelections)
	
	updateProducts_Custom_Design = pblnResult	

End Function    'updateProducts_Custom_Design

'***********************************************************************************************

%>
