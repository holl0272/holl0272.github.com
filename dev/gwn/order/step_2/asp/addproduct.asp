<%@ Language=VBScript %>
<%	Option Explicit
	Response.Buffer = True
	Server.ScriptTimeout = 300

'********************************************************************************
'*
'*   addproduct.asp -
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins are addproduct.asp
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
<!--#include file="incCoreFiles.asp"-->
<script language="javascript" src="SFLib/sfCookieCheck.js"></script>
<%

'**********************************************************
'	Developer notes
'**********************************************************


'Cart additions can come in the following form:

'SF2k compatibility

'SF5 default

'SF5 MPO

'SF5 Attribute Extender

'SF5 MPO + Attribute Extender

'SF5 - Master Template
'Quantity.ProductID.AttributeDetailID
'Note AttributeID only present if

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim mblnRedirectToLogin
Dim mblnSaveCart

'**********************************************************
'*	Functions
'**********************************************************

	Sub cleanupPageObjects

	On Error Resume Next

		Call cleanup_ssclsCartContents
		Call cleanup_dbconnopen

	End Sub

'**********************************************************
'*	Begin Page Code
'**********************************************************

If CBool(Len(Session("ssDebug_AddToCart")) > 0) Then
	vDebug = 1
	'Response.Flush
	On Error Goto 0
End If

mblnRedirectToLogin = False

' Collect variables passed from the Form

mblnSaveCart = CBool(Len(Trim(Request.Form("SaveCart.x"))) > 0)

If addProductsToCart Then Call removeCartFromSession

Call cleanupPageObjects

If mblnRedirectToLogin Then
	' Write to a cookie for thank you redirection
	Response.Cookies("sfThanks")("PreviousAction") = "SaveCart"
	Response.Redirect("login.asp")
End If

'added for debugging
If vDebug = 1 Then
	debug.enabled = True
	debug.DisplayInScreen = True
'	Response.Write "<fieldset><legend>CartContents</legend>" & CartContents & "</fieldset>"
	Response.Write "<a href='" & redirectURL & "'>Return to " & redirectURL & "</a><br />"
	Response.Write "<a href='' onclick='javascript:window.history.back(-1); return false;'>history.back(-1)</a><br />"
	Response.Write "redirectURL: " & redirectURL & "<br />"
	Response.Write "<a href='order.asp'>Checkout</a><br />"
	Response.End
Else
	Response.Redirect(redirectURL)
End If

'**********************************************************
'	Functions
'**********************************************************

Function addProductsToCart()

Dim aProdAttr, aProdValues
Dim i
Dim iCounter
Dim iProdAttrNum
Dim iQuantity
Dim iSvdCartID
Dim iTmpOrderDetailId
Dim iUpperBound
Dim pblnOldPage
Dim pblnSuccess
Dim plngCurrentQty
Dim plngGiftWrapQuantity
Dim plngNumProductsToAdd
Dim pstrReturningFromLogin
Dim sProdID
Dim sReferer
Dim sTmpAttrName

	Call DebugRecordSplitTime("Beginning add products to cart")
	pblnOldPage = CBool(Len(Trim(Request.Form("Order_Flag"))) > 0)
	pstrReturningFromLogin = Request.QueryString("logedin")
	sReferer = visitor_REFERER & "," & vistor_HTTP_REFERER & "," & visitor_REMOTE_ADDR
	pblnSuccess = False

	plngNumProductsToAdd = numberProductsToAddToCart()
	For i = 0 To plngNumProductsToAdd
		Call DebugRecordSplitTime("Getting product " & i & " (" & sProdID & ") information to add to cart . . .")
		Call getProductToAdd(sProdID, iQuantity, i, pblnOldPage)
		Call DebugRecordSplitTime("Product " & i & " (" & sProdID & ") information retrieved")

		If Len(iQuantity) > 0 Then
			pblnSuccess = True

			If Not IsNumeric(iQuantity) Then iQuantity = 1
			maryCartAdditions(i)(enCartItem_QtyToAdd) = iQuantity

			'Now determine the Gift Wrap quantities
			If CBool(Request.Form("chkGiftWrap." & sProdID) = "1") Then
				plngGiftWrapQuantity = iQuantity
			ElseIf CBool(Request.Form("chkGiftWrap") = "1") Then
				plngGiftWrapQuantity = iQuantity
			Else
				plngGiftWrapQuantity = 0
			End If
			maryCartAdditions(i)(enCartItem_QtyToGW) = iQuantity

			Call setCookie_sfSearch
			Call setCookie_sfAddProduct

			If Len(pstrReturningFromLogin) = 0 Then
				ReDim aProdValues(3)

  				If Len(getProductInfo(sProdID, enProduct_NamePlural)) > 0 And iQuantity > 1 Then
  					aProdValues(0) = getProductInfo(sProdID, enProduct_NamePlural)
  				Else
  					aProdValues(0) = getProductInfo(sProdID, enProduct_Name)
  				End If

				If Len(Trim(aProdValues(0))) > 0 Then

					If getProductInfo(sProdID, enProduct_Exists) Then
						aProdValues(1) = getProductInfo(sProdID, enProduct_Message)
						aProdValues(2) = getProductInfo(sProdID, enProduct_AttrNum)
						aProdValues(3) = getProductInfo(sProdID, enProduct_ShipIsActive)

						If getProductInfo(sProdID, enProduct_IsActive) = 0 Then
							maryCartAdditions(i)(enCartItem_QtyAdded) = -1
							maryCartAdditions(i)(enCartItem_ResponseMessage) = "<em>" & aProdValues(0) & "</em> is not currently available."
						Else
  							If Len(aProdValues(2)) > 0 Then iProdAttrNum = aProdValues(2)
							If mblnSaveCart Then
								maryCartAdditions(i)(enCartItem_AddType) = True
								maryCartAdditions(i)(enCartItem_QtyAdded) = iQuantity
							Else
								maryCartAdditions(i)(enCartItem_AddType) = False
								Call DebugRecordSplitTime("Getting product " & i & " (" & sProdID & ") inventory level . . .")
								Call getProductInventoryLevel(CleanAttributeArray(maryCartAdditions(i)(enCartItem_AttributeArray)))
								Call DebugRecordSplitTime("Product " & i & " (" & sProdID & ") inventory level retrieved")
							End If

							maryCartAdditions(i)(enCartItem_prodID) = sProdID
							maryCartAdditions(i)(enCartItem_prodName) = aProdValues(0)
							If adminGlobalConfirmationMessageIsactive Then
								maryCartAdditions(i)(enCartItem_Upsell) = getProductInfo(sProdID, enProduct_Message) & adminGlobalConfirmationMessage
							Else
								maryCartAdditions(i)(enCartItem_Upsell) = getProductInfo(sProdID, enProduct_Message)
							End If
							aProdAttr = maryCartAdditions(i)(enCartItem_AttributeArray)

							maryCartAdditions(i)(enCartItem_QtyInStock) = getProductInfo(sProdID, enProduct_invenInStock)
							maryCartAdditions(i)(enCartItem_invenbBackOrder) = getProductInfo(sProdID, enProduct_invenbBackOrder)
							maryCartAdditions(i)(enCartItem_invenbTracked) = getProductInfo(sProdID, enProduct_invenbTracked)

							'-------------------------------------------------------------------------------
							' This shows whether there is a previous order or a new order.
							' New Products are treated like new orders but can be gathered together through
							' the session variable SessionCartID
							'-------------------------------------------------------------------------------
							Call DebugRecordSplitTime("Checking for existing temp order . . .")
							iTmpOrderDetailId = getOrderID("odrdttmp", "odrattrtmp", sProdID, aProdAttr, iProdAttrNum, plngCurrentQty)
							Call DebugRecordSplitTime("Temp order retrieved")
							maryCartAdditions(i)(enCartItem_QtyInCart) = plngCurrentQty

							If vDebug = 1 Then
								If Len(CStr(iTmpOrderDetailID)) > 0  Then
									If iTmpOrderDetailID = -1 Then
										Response.Write "Setting temporary cart detail -- <b>Product doesn't exist in current cart</b><br />"
									Else
										Response.Write "Setting temporary cart detail -- <b>Matching product already exists in cart</b> iTmpOrderDetailID " & iTmpOrderDetailID & "<br />"
									End If
								Else
									Response.Write "<h4><font color=red>Error setting temporary cart detail</font></h4>"
								End If

								Response.Write "<fieldset><legend>Product " & maryCartAdditions(i)(enCartItem_prodID) & "</legend>"
								Response.Write "enCartItem_prodName: " & maryCartAdditions(i)(enCartItem_prodName) & "<br />"
								Response.Write "enCartItem_QtyToAdd: " & maryCartAdditions(i)(enCartItem_QtyToAdd) & "<br />"
								Response.Write "enCartItem_QtyInCart: " & maryCartAdditions(i)(enCartItem_QtyInCart) & "<br />"
								Response.Write "enCartItem_QtyInStock: " & maryCartAdditions(i)(enCartItem_QtyInStock) & "<br />"
								Response.Write "enCartItem_QtyToGW: " & maryCartAdditions(i)(enCartItem_QtyToGW) & "<br />"
								Response.Write "enCartItem_invenbTracked: " & maryCartAdditions(i)(enCartItem_invenbTracked) & "<br />"
								Response.Write "enCartItem_invenbBackOrder: " & maryCartAdditions(i)(enCartItem_invenbBackOrder) & "<br />"
								Response.Write "enCartItem_Upsell: " & maryCartAdditions(i)(enCartItem_Upsell) & "<br />"
								Response.Write "enCartItem_AddType: " & maryCartAdditions(i)(enCartItem_AddType) & "<br />"

								Response.Write "</fieldset>"
							End If	'vDebug = 1

							If Len(iTmpOrderDetailID) > 0  Then

								If mblnSaveCart Then
									If Not isLoggedIn Then
										' If no cookie with custID, direct to Login
										' Write to saved with custID of 0
										Call DebugRecordSplitTime("Checking for existing temp order . . .")
										Call DebugRecordSplitTime("Temp order retrieved")
										Call getSavedTable(aProdAttr, sProdID, iQuantity, 0, sReferer)
										mblnRedirectToLogin = True
									Else
										' Check for existing SessionCartId, -1 is returned if not found
										iSvdCartID = getOrderID("odrdtsvd", "odrattrsvd", sProdID, aProdAttr, iProdAttrNum, plngCurrentQty)
										If vDebug = 1 Then Response.Write "<p>Saved Cart Found or Not Found -- Record " & iSvdCartID

										If Len(CStr(iSvdCartID)) > 0 Then
											If iSvdCartID < 0 Then
												If vDebug = 1 Then Response.Write "<h2><font face=verdana color=#334455>Prior Record Not Found</h2></font>"
												Call getSavedTable(aProdAttr,sProdID,iQuantity,custID_cookie,sReferer)
											Else
												If vDebug = 1 Then Response.Write "<h2><font face=verdana color=#334455>Adding to Existing Saved Cart</h2> <br />SvdCartID = " & iSvdCartID
												Call setUpdateQuantity("odrdtsvd",iQuantity,iSvdCartID)
											End If
										Else
											Response.Write "<p>Number of attributes not equal to the product specs or database writing error"
										End If	'Len(CStr(iSvdCartID)) > 0
									End If	' End No Cookie If
								Else

									'Process Inventory
									If maryCartAdditions(i)(enCartItem_invenbTracked) = 1 Then
									If CBool(maryCartAdditions(i)(enCartItem_QtyToAdd) + maryCartAdditions(i)(enCartItem_QtyInCart) > maryCartAdditions(i)(enCartItem_QtyInStock)) AND CBool(maryCartAdditions(i)(enCartItem_invenbTracked) = 1) Then
										If maryCartAdditions(i)(enCartItem_invenbBackOrder) <> 1 Then
											iQuantity = maryCartAdditions(i)(enCartItem_QtyInStock) - maryCartAdditions(i)(enCartItem_QtyInCart)
										Else
											'If the qty in the cart is > available this implies acceptance of the backorder
											If CBool(maryCartAdditions(i)(enCartItem_QtyInCart) > maryCartAdditions(i)(enCartItem_QtyInStock)) Then
												maryCartAdditions(i)(enCartItem_AskForBackOrder) = False
											Else
												iQuantity = maryCartAdditions(i)(enCartItem_QtyInStock) - maryCartAdditions(i)(enCartItem_QtyInCart)
												maryCartAdditions(i)(enCartItem_AskForBackOrder) = True
											End If
										End If
									End If
									End If
									maryCartAdditions(i)(enCartItem_QtyAdded) = iQuantity

  									If Len(getProductInfo(sProdID, enProduct_NamePlural)) > 0 And maryCartAdditions(i)(enCartItem_QtyAdded) > 1 Then
  										maryCartAdditions(i)(enCartItem_prodName) = getProductInfo(sProdID, enProduct_NamePlural)
  									Else
  										maryCartAdditions(i)(enCartItem_prodName) = getProductInfo(sProdID, enProduct_Name)
  									End If

									If (iTmpOrderDetailID < 1) Then
										If vDebug = 1 Then Response.Write "<h2><font face=verdana color=#334455>New Order Item</h2>"
										Call DebugRecordSplitTime("getTmpTable . . .")
										iTmpOrderDetailID = getTmpTable(aProdAttr, sProdID, iQuantity, sReferer, aProdValues(3))
										Call DebugRecordSplitTime("getTmpTable complete")
									ElseIf (iTmpOrderDetailID > 0 AND iTmpOrderDetailID <> "") Then
										If vDebug = 1 Then Response.Write "<h2><font face=""verdana"" color=""#334455"">Adding to Existing Cart: Found OrderID = " & iTmpOrderDetailId & "</font></h2>"
										Call DebugRecordSplitTime("setUpdateQuantity . . .")
										Call setUpdateQuantity("odrdttmp", iQuantity, iTmpOrderDetailId)
										Call DebugRecordSplitTime("setUpdateQuantity complete")
									End If	' End Add Product If
									maryCartAdditions(i)(enCartItem_tmpOrderDetailId) = iTmpOrderDetailId

									'Now for the gift wrap
									If cblnSF5AE Then
										Call DebugRecordSplitTime("AddProduct_WriteTmpOrderDetailsAE . . .")
										Call AddProduct_WriteTmpOrderDetailsAE(iTmpOrderDetailId, iQuantity, plngGiftWrapQuantity)
										Call DebugRecordSplitTime("AddProduct_WriteTmpOrderDetailsAE complete")
									End If
								End If	' mblnSaveCart
							Else
								Response.Write "<br />Unknown ActionType Occurred or Database writing error"
								Response.End
							End If	'Len(iTmpOrderDetailID) > 0
						End If	'getProductInfo(sProdID, enProduct_IsActive) = 0
					Else
						maryCartAdditions(i)(enCartItem_QtyAdded) = -1
						maryCartAdditions(i)(enCartItem_ResponseMessage) = "<em>" & sProdID & "</em> is not a valid product."
					End If	'getProductInfo(sProdID, enProduct_Exists)
				Else
					maryCartAdditions(i)(enCartItem_QtyAdded) = 0
					maryCartAdditions(i)(enCartItem_ResponseMessage) = "This product is not currently available."
				End If	'Len(Trim(aProdValues(0))) > 0
			Else
				Response.Cookies("sfAddProduct").Expires = Now()
				Response.Cookies("sfThanks").Expires = Now()
			End If	'Len(pstrReturningFromLogin) = 0

			'added to store cart to session variable
			'Call SetCartSummaryToSession

		End	If 'Len(iQuantity) > 0
	Next	'i

	addProductsToCart = pblnSuccess

	If Not pblnSuccess Then maryCartAdditions = "<font class=""Content_Large""><b>No quantity was selected!</b></font><p>Please enter a quantity for at least one product.</p>"
	Call setCartAdditionResultsToSession

End Function	'addProductsToCart

'**********************************************************

Function redirectURL()

Dim pstrURL

	pstrURL = Request.ServerVariables("HTTP_REFERER")
	If Len(pstrURL) = 0 Then
		'Some browsers will have this blocked so . . .
		If mblnSaveCart Then
			pstrURL = "savecart.asp"
		Else
			pstrURL = "order.asp"
		End If
	ElseIf Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
		pstrURL = pstrURL & "?" & Request.ServerVariables("QUERY_STRING")
	End If
	'added for ISAPI Rewrite
	pstrURL = Replace(pstrURL, "%3D", "=")

	'redirectURL = pstrURL
	redirectURL = "http://dev.gamewearnow.com/order.asp"

End Function	'redirectURL

'**********************************************************
%>