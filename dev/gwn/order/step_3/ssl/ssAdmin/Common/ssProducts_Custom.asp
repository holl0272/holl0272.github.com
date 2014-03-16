<%
'********************************************************************************
'*   Common Support File For StoreFront 6.0 add-ons
'*   Custom Product Management Routines
'*   Release Version:	2.00.001		
'*   Release Date:		January 1, 2006
'*   Revision Date:		January 1, 2006
'*
'*   This file must be included from ssProducts_Common
'*
'*   Release Notes:
'*
'*   2.00.001 (January 1, 2006)
'*	 - Initial Release
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

'**********************************************************
'*	Global variables
'**********************************************************

Const cblnUseCustom_Design = False
Const cblnUseCustom_Vehicle = False

'**********************************************************
'*	Functions
'**********************************************************

'Function DeleteProduct_Custom(byVal lngUID)
'Function DeleteAllProducts_Custom()
'Function updateProducts_Custom(byVal lngProductUID)


'--------------------------------------------------------------------------------------------------

Function useCheckboxForFieldDisplay(aryFieldsToDisplay)

Dim pblnResult

	If aryFieldsToDisplay(1) = enDatatype_boolean Then
		pblnResult = True
	Else
		Select Case aryFieldsToDisplay(0)
			Case "IsActive", _
					"IsOnSale", _
					"IsShipable", _
					"HasCountryTax", _
					"HasStateTax", _
					"HasLocalTax", _
					"Inventory_Tracked", _
					"DropShip", _
					"DealTimeIsActive", _
					"MMIsActive", _
					"FroogleIsEnabled", _
					"NextagIsEnabled", _
					"ShoppingcomIsEnabled", _
					"YahooSubmitIsEnabled", _
					"PriceGrabberIsEnabled", _
					"ShopzillaIsEnabled"
				pblnResult = True
			Case Else
				'nothing
		End Select
	End If

	useCheckboxForFieldDisplay = pblnResult

End Function	'useCheckboxForFieldDisplay

'--------------------------------------------------------------------------------------------------

Function DeleteProduct_Custom(byVal lngUID)

Dim pblnResult
Dim pstrSQL

'On Error Resume Next

	If len(lngUID) = 0 Then
		DeleteProduct_Custom = False
		Exit Function
	End If
	
	pblnResult = True
	
	If cblnUseCustom_Design Then pblnResult = pblnResult And DeleteProduct_Custom_Design(lngUID)
	If cblnUseCustom_Vehicle Then pblnResult = pblnResult And DeleteProduct_Custom_Vehicle(lngUID)

    If (Err.Number = 0) Then

    Else
        Call addMessageItem("Error deleting the product catalog: " & Err.Description, True)
        pblnResult = False
    End If
    
    DeleteProduct_Custom = pblnResult
    
End Function    'DeleteProduct_Custom

'--------------------------------------------------------------------------------------------------

Function updateProducts_Custom(byVal lngProductUID, byVal strProductID)

Dim pblnResult
Dim pstrSQL

'On Error Resume Next

	pblnResult = True
	
	If cblnUseCustom_Design Then pblnResult = pblnResult And updateProducts_Custom_Design(lngProductUID)
	If cblnUseCustom_Vehicle Then pblnResult = pblnResult And updateProducts_Custom_Vehicle(strProductID)
    	
    If (Err.Number = 0) Then

    Else
        Call addMessageItem("Error deleting the product catalog: " & Err.Description, True)
        pblnResult = False
    End If
    
	updateProducts_Custom = pblnResult	

End Function    'updateProducts_Custom

%>
<!--#include file="ssProducts_Custom_Designs.asp"-->
<!--#include file="ssProducts_Custom_Vehicles.asp"-->

