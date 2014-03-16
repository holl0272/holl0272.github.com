<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4013.0.2

'@FILENAME: 
	


'@DESCRIPTION: 

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'Modified 10/23/01 
'Storefront Ref#'s: 157 'JF
'Modified 10/31/01 
'Storefront Ref#'s: 193 'JF

' Public Variables needed for AE functions in this module
Dim gTmpSQL

' Public Initialization
gtmpSQL = "Select * FROM sfTmpOrderDetails as A LEFT JOIN sfTmpOrderDetailsAE as B ON A.odrdttmpID = B.odrdttmpaeID "

Sub LockApp
'For inventory tracking purpose
	IF Application("InventoryLock") = "Locked"   Then
		JS "window.history.go(-1)"	
		Response.End 
	Else
		Application.Lock 
		Application("InventoryLock") = "Locked"
		Application("LockTime") = Now()  'save date time of lock
		Application("LockSessionID") = SessionID
		Application.UnLock
	End IF

End Sub

Sub ShowLockValues
	'DEBug mode only
	Response.Write "<br />" & vbcrlf
	Response.Write "<br /> InventoryLock:" & Application("InventoryLock") & vbcrlf
	Response.Write "<br /> LockTime:" & Application("LockTime") & vbcrlf
	Response.Write "<br /> LockSessionID:" & Application("LockSessionID") & vbcrlf
	Response.Write "<br /> My SessionID:" & SessionID & vbcrlf
	'Response.write "DateDiff:" & DateDiff("s",Application("LockTime"), Now)
End Sub

Sub ReleaseAppLock
'This routine prevents from infinite or bad locks 
	IF Application("InventoryLock") = "Locked" Then
		If  DateDiff("s",Application("LockTime"), Now) > 120 Or Application("LockSessionID") = SessionID  Then
			Application.Lock 
			Application("InventoryLock") = "Unlocked"
			Application("LockTime") = ""  
			Application("LockSessionID") = ""
			Application.UnLock
		End If
	End If
End Sub

Sub UnlockApp
	Application.Lock 
	Application("InventoryLock") = "Unlocked"
	Application("LockTime") = ""  
	Application("LockSessionID") = ""
	Application.UnLock
End Sub

'******************************************* CONFIRM PAGE ***************************

Sub Confirm_SaveAmounts(byVal lngOrderID)

Dim pobjCmd
Dim pobjRS

Dim orderBackOrderAmount
Dim orderBillAmount
Dim orderCouponDiscount

	If Len(lngOrderID & "") = 0 Or Not isNumeric(lngOrderID) Then Exit Sub

	orderBackOrderAmount = CorrectEmptyValue(Session("BackOrderAmount"), 0)
	orderBillAmount = CorrectEmptyValue(Session("BillAmount"), 0)
	
	If cdbl(Session("CouponDiscountPercent")) > 0 then
		orderCouponDiscount = CorrectEmptyValue(Session("CouponDiscountPercent"), 0)
	Else
		orderCouponDiscount = CorrectEmptyValue(Session("CouponDiscountAmount"), 0)
	End If

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select orderAEID FROM sfOrdersAE where orderAEID=?"
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("orderAEID", adInteger, adParamInput, 4, lngOrderID)
  		Set pobjRS = .Execute
		
  		If pobjRS.EOF Then
			.Commandtext = "Insert Into sfOrdersAE (orderAEID,orderBackOrderAmount,orderBillAmount,orderCouponDiscount) Values (?,?,?,?)"
			.Parameters.Append .CreateParameter("orderBackOrderAmount", adWChar, adParamInput, 50, orderBackOrderAmount)
			.Parameters.Append .CreateParameter("orderBillAmount", adWChar, adParamInput, 50, orderBillAmount)
			.Parameters.Append .CreateParameter("orderCouponDiscount", adWChar, adParamInput, 50, orderCouponDiscount)
			.Execute , , adExecuteNoRecords
  		Else
  			.Parameters.Remove("orderAEID")
			.Commandtext = "Update sfOrdersAE Set orderBackOrderAmount=?, orderBillAmount=?, orderCouponDiscount=? Where orderAEID=?"
			.Parameters.Append .CreateParameter("orderBackOrderAmount", adWChar, adParamInput, 50, orderBackOrderAmount)
			.Parameters.Append .CreateParameter("orderBillAmount", adWChar, adParamInput, 50, orderBillAmount)
			.Parameters.Append .CreateParameter("orderCouponDiscount", adWChar, adParamInput, 50, orderCouponDiscount)
 			.Parameters.Append .CreateParameter("orderAEID", adInteger, adParamInput, 4, lngOrderID)
			.Execute , , adExecuteNoRecords
 		End If
  		closeObj(pobjRS)
	End With
	Set pobjCmd = Nothing
	
End Sub	'Confirm_SaveAmounts

'***************************************************************************************************************************************************************************

Sub Confirm_CheckCartAndRedirect

	If  CheckCartInventory = 0 then 'stock depleted !  
		Session("ShowInventoryMessage") = "1"
		Call cleanup_dbconnopen	'This line needs to be included to close database connection
		response.Redirect(C_HomePath & "order.asp")
	End If

End Sub

'***************************************************************************************************************************************************************************

Sub Confirm_UpdateInventory

dim sql
dim rst
dim i

	Set rst = CreateObject("ADODB.RecordSet")		
	sql = gtmpSQL &  " WHERE odrdttmpSessionId=" & SessionID
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	If rst.recordcount > 0 then
		rst.movefirst
		For i = 1 to rst.recordcount
		
			UpdateAvailableQTY rst("odrdttmpProductId"),GetAttDetailID(rst("odrdttmpID"),"tmp"),clng(rst("odrdttmpQuantity"))
			If not rst.eof then rst.movenext
		Next
	End If
	CloseObj (rst)

End Sub

'***************************************************************************************************************************************************************************

'-------------------------------------------------------------------------------
'Purpose: gets attdetailid field based on order attributes - for inventory purposes
'Accepts: 
'Returns: a formulated invenAttDetailID field e.g. 48,51,89 
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetAttDetailID (iDetailID,sType)
dim sql
dim rst
dim i
dim sAttID
	
	If Len(Trim(iDetailID  & "")) = 0 Then Exit Function

	Select Case sType
		Case "svd":
			sql = "Select * FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailID=" & makeInputSafe(iDetailID) & " ORDER BY odrattrsvdAttrID" 
		Case "tmp":
			If cblnSQLDatabase Then
				sql = "Select * FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=" & makeInputSafe(iDetailID) & " ORDER BY Cast(int, odrattrtmpAttrID)" 
			Else
				sql = "Select * FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=" & makeInputSafe(iDetailID) & " ORDER BY CLng(odrattrtmpAttrID)" 
			End If

			'Note: SF changed this in 50.5 but didn't document why; this changes it back; necessary because of Attribute Extender
			sql = "Select * FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=" & makeInputSafe(iDetailID) & " ORDER BY odrattrtmpAttrID"
		Case "odr":
			sql = "Select * FROM sfOrderDetailsAE WHERE odrdtAEID=" & makeInputSafe(iDetailID) 'b2
		Case Else
			Exit Function
	End Select

	Set rst = CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount <= 0 then
		GetAttDetailID = 0
		closeobj(rst)
		exit function
	End if
	
	sAttId = ""
	
	If sType ="odr" then   'b2
		sAttId = rst("odrdtAttDetailID")
		GetAttDetailID = sAttID						
		CloseObj (rst)
		Exit Function
	End If

	rst.movefirst
	For i = 1 to rst.recordcount
		If sType ="svd" then 
			If sAttID <> "" then
				sAttId = sAttId &  "," & rst("odrattrsvdAttrID") 
			else
				sAttId = rst("odrattrsvdAttrID")
			End if
				
		Elseif  sType ="tmp" then 
			if sAttID <> "" then
				sAttId = sAttId & "," & rst("odrattrtmpAttrID") 
			else
				sAttId = rst("odrattrtmpAttrID")
			end if
				
		End if
	
		If not rst.eof then rst.movenext
	Next
		
	GetAttDetailID = sAttID						
	CloseObj (rst)

End Function

'***************************************************************************************************************************************************************************

Function GetAttName (iDetailID,sType)
dim sql
dim rst
dim i
dim sAttID
dim sAttName

	Select Case sType
		Case "svd":
			sql = "Select odrattrsvdAttrID FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailID=" & makeInputSafe(iDetailID)
		Case "tmp":
			sql = "Select odrattrtmpAttrID FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=" & makeInputSafe(iDetailID)
		Case "odr":
			sql = "Select odrattrAttribute FROM sfOrderAttributes WHERE odrattrOrderDetailId=" & makeInputSafe(iDetailID)
	End Select
	
	Set rst = CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.EOF Then
		sAttId = ""
	Else
		For i = 1 to rst.recordcount
			If Len(sAttID) = 0 then
				sAttId = rst.Fields(0).Value
			else
				sAttId = sAttId &  cstrMultipleAttributeDelimiter & rst.Fields(0).Value
			End if
					
			If not rst.eof then rst.movenext
		Next
	End if
	closeobj(rst)
	
	'inserted for Attribute Extender
	If vDebug = 1 Then Response.Write "GetAttName - sAttId: " & sAttId & "<br />"
	Call CleanAttributeID(sAttId)
	If vDebug = 1 Then Response.Write "GetAttName - sAttId: " & sAttId & "<br />"
	
	If Len(sAttId) > 0 Then
		Set rst = CreateObject("ADODB.RecordSet")
		sql = "Select * FROM sfAttributeDetail WHERE attrdtID in (" & makeInputSafe(sAttId) & ") order by attrdtAttributeId" 
		
		rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
		
		if rst.recordcount > 0 then
			for i = 1 to rst.recordcount 
				sAttName = sAttName & " " & rst("attrdtName")
				rst.movenext
			next 
		else
			sAttName = ""
		end if
	End If
	
	GetAttName = sAttName
	CloseObj (rst)
	
End Function	'GetAttName

'******************************************* SAVE CART PAGE *********************************

Sub SaveCart_WritesvdtmpAERecord(byVal iSvdOrderDetailID, byVal iTmpOrderDetailID)

Dim rst2
Dim sql
	
	sql = "Select odrdttmpAEID FROM sfTmpOrderDetailsAE WHERE odrdttmpAEID=" & makeInputSafe(iTmpOrderDetailID)
	Set rst2 = CreateObject("ADODB.RecordSet")		
    rst2.CursorLocation = 2 'adUseClient
	rst2.Open sql, cnn, 3, 1	'adOpenStatic, adLockReadOnly
	If rst2.eof then
		sql = "Insert Into sfTmpOrderDetailsAE (odrdttmpaeID) Values (" & makeInputSafe(iTmpOrderDetailID) & ")"
		cnn.Execute sql,,128
	End if
	CloseObj (rst2)

End Sub	'SaveCart_WritesvdtmpAERecord

'********************************** SEARCH_RESULTS PAGE *************************************

Sub SearchResults_GetGiftWrap(byVal strProductID)

Dim ret
	
	ret = GetGiftWrapPrice(strProductID)
	
	select case ret
		case "X" 'no gift wrap for this product
			
		case 0 'gift wrap for free if price is 0
			Response.Write "<p align=""left""><INPUT name=chkGiftWrap type=checkbox value = 1 >Gift wrap (free of charge!)</p>" & vbcrlf
		case else
			Response.Write "<p align=""left""><INPUT name=chkGiftWrap type=checkbox value = 1 >Gift wrap (add " & FormatCurrency(ret)  & " per item)</p>" & vbcrlf
    end select
    
End Sub	'SearchResults_GetGiftWrap

'***************************************************************************************************************************************************************************

Sub SearchResults_ShowMTPricesLink(byVal strProdID) 

	If hasMTP(strProdID) Then Response.Write "<br /><a href=""javascript:show_page('MTPrices.asp?sProdId=" & strProdID & "')"">Check Volume Discounts</a>" & vbcrlf	

End Sub	'SearchResults_ShowMTPricesLink

'***************************************************************************************************************************************************************************

Sub SearchResults_GetProductInventory(byVal strProduct)

Dim ret

	ret = CheckInventoryTracked (strProduct)

	If ret = 1 then 
		
		If CheckShowStatus(strProduct) <> 1 Then Exit Sub
			
		Select Case CheckInStock(strProduct)
			Case "X" 'inventory not tracked for this product
				
			Case 0 'inventory tracked
				Response.Write "<br /> " & vbcrlf
				Response.Write "Out of Stock!" & vbcrlf
				If checkbackorder(strProduct) = 1 then	
					Response.Write "<br /> Click ""Add to Cart"" to BackOrder!" & vbcrlf
				Else
					If False Then
						Response.Write "<br /> Click <a href=""detail_NotifyMe.asp?product_id=" & strProduct & """>here</a> to be notified when it arrives!" & vbcrlf
					End If
				End If
													
			Case Else
				Response.Write "<br /> <a href=""javascript:show_stockinfo('StockInfo.asp?sProdId=" & strProduct & "')"">Check Stock</a>" & vbcrlf	
		End  Select
	end if
	
End Sub	'SearchResults_GetProductInventory

'********************************* ORDER PAGE (inventory stuff) ***********************************************

Sub Order_ShowInventoryMessage

	Order_AdjustCart
	If Session("ShowInventoryMessage") = "1" then
		'Order_AdjustCart
		js "show_page(" & chr(34) & "invenMessage.asp" & chr(34) & ")"
	End if
	
	Session("ShowInventoryMessage") = "0"

End Sub

'***************************************************************************************************************************************************************************

Sub Order_AdjustCart

	Dim ret
	dim rstAll,sql,i,sPath,sProdName,sAttName,bo
	dim inv,ordqty,itmporderdetailid,sprodid,avlqty,sattdetailid,sResponseMessage
	dim boqty
	
	If  CheckCartInventory = 0  then
	
		SQL = "Select * FROM sfTmpOrderDetails as A LEFT JOIN sfTmpOrderDetailsAE as B ON A.odrdttmpID = B.odrdttmpaeID WHERE odrdttmpSessionID = " & SessionID
	
		Set rstAll = CreateObject("ADODB.RecordSet")
		rstAll.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
	
		For i = 1 to rstall.RecordCount 
		
			inv =  CheckInventoryTracked(rstall("odrdttmpProductID"))  'inventory tracked for  this product?
			'bo = CheckBackOrder(rstall("odrdttmpProductID")) ' backorder allowed ?
			ordQTY = rstall("odrdttmpQuantity")
			boqty = rstAll("odrdttmpBackOrderQTY")
			iTmpOrderDetailID = rstall("odrdttmpID")
			sProdID = rstall("odrdttmpProductID")
			sAttDetailID = GetAttDetailID(iTmpOrderDetailID,"tmp")
			avlqty = GetAvailableQTY(sProdId, sAttDetailID) 
			sProdName = GetProductName(sProdId)
			sAttName = getattname(iTmpOrderDetailId,"tmp")
			
			If ordqty > 1 Then
				sResponseMessage = "have been added to your order."
			Else
				sResponseMessage = "has been added to your order."
			End If
			
			If  ordqty > avlqty AND boqty <= 0 then
				sPath = "InventoryOD.asp?sProdID=" & sProdID & "&iTmpOrderDetailID=" & iTmpOrderDetailID & "&iQuantity=" & iQuantity & "&sProdName=" & sProdName & "&sResponseMessage="& Server.URLEncode(sResponseMessage) 'AE
				js "show_page(" & chr(34) & sPath & chr(34) & ")"
			End if
			If not  rstAll.EOF then	rstAll.movenext
		
		next
	
	Else
		ret = DeleteBadItems
		ValidateCartItems   ' corrects items qty according to backorder-flag and stock-qty
		ret = DeleteBadItems
	End If

End Sub	'Order_AdjustCart

'***************************************************************************************************************************************************************************

Sub Order_Update_GiftWrapsBackOrder(byVal lngTmpOrderID)

dim rst
dim sql
dim Price
dim gwqty
dim boflag
dim ProdId,avlqty
	
	sql = "Select * FROM sfTmpOrderDetailsAE WHERE odrdttmpaeID=" & makeInputSafe(lngTmpOrderID)
	Set rst = CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	If Not rst.eof then	
		gwqty = Request.Form("GWQTY" & iCounter)
		prodid = Request.Form("sProdID" & iCounter)
		If gwqty <> "" AND Not IsNull(gwqty) then
			If GetGiftWrapPrice(ProdId) <> "X" then
				If gwqty > iNewQuantity then gwqty = iNewQuantity
				rst.Fields("odrdttmpGiftWrapQTY") = gwqty
			End if
		End if
		
		avlqty = getavailableqty(prodid, getattdetailid(rst.Fields("odrdttmpaeID").Value, "tmp"))
		If avlqty <> "X" Then
			IF clng(avlqty) => clng(iNewQuantity) then
				rst.Fields("odrdttmpBackOrderQty") = 0 'beta 2
			End if
		End IF
		
		rst.update
	End IF
	CloseObj (rst)
	'response.Redirect(C_HomePath & "order.asp")	
	
End Sub	'Order_Update_GiftWrapsBackOrder

'************************************** ADD PRODUCT PAGE ***************************************************

'*****************************************************************************************

Sub AddProduct_WriteTmpOrderDetailsAE(byVal lngTmpOrderDetailID, byVal lngQuantity, byVal lngGiftWrapQuantity)

Dim rst
Dim sql
	
	'Write to sfTmpOrderDetailsAE
	sql = "Select * FROM sfTmpOrderDetailsAE WHERE odrdttmpaeID=" & makeInputSafe(lngTmpOrderDetailID)
	Set rst = CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	If not rst.recordcount > 0 then
		rst.AddNew 
		rst.Fields("odrdttmpaeID") = lngTmpOrderDetailID
	End If
	if rst.Fields("odrdttmpGiftWrapQty") > 0 then
		rst.Fields("odrdttmpGiftWrapQty") = rst.Fields("odrdttmpGiftWrapQty") + lngGiftWrapQuantity
	else
		rst.Fields("odrdttmpGiftWrapQty") = lngGiftWrapQuantity
	end if
		
	
	rst.update
	CloseObj (rst)
	
End Sub	'AddProduct_WriteTmpOrderDetailsAE


'*****************************************************************************************
'*****************************************************************************************
'********************** Independent procedures and functions *****************************
'*****************************************************************************************
'*****************************************************************************************


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub AdjustTmpOrderDetails(byVal iTmpOrderDetailID)

Dim sql,rst
dim gwqty
dim avlqty
dim ordqty

	If Len(iTmpOrderDetailID & "") = 0 Or Not isNumeric(iTmpOrderDetailID) Then Exit Sub

	sql = gtmpSQL & " where odrdttmpID=" & makeInputSafe(iTmpOrderDetailID) 'rsAllOrders("odrdttmpProductID")
	Set rst = CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenKeySet, adLockOptimistic, adCmdText
	
	If rst.RecordCount <= 0 then exit sub 'error this should never happen
	
	avlqty = GetAvailableQTY(rst.Fields("odrdttmpProductID").Value, GetAttDetailID (rst.Fields("odrdttmpId").Value, "tmp"))
	gwqty =  rst.Fields("odrdttmpGiftWrapQty").Value
	ordqty = rst.Fields("odrdttmpQuantity").Value
	
	If CheckInventoryTracked(rst.Fields("odrdttmpProductID").Value) <> 1 or Not IsNumeric(avlqty) then 'b2
		closeobj(rst)
		Exit Sub
	End If
	
	If avlqty = 0  AND CheckBackOrder(rst.Fields("odrdttmpProductID").Value) <> 1 then
		' Delete item if out of stock with no backorder
		DeletetmpOrderDetailsAE(rst.Fields("odrdttmpId").Value)
		closeobj(rst)
		exit sub	
	End If
	
	'If avlqty < ordqty then ordqty = avlqty
	
	'beta 2
	If avlqty < ordqty then 
		rst.Fields("odrdttmpBackOrderQty").Value = clng(ordqty)- clng(avlqty)
	else
		rst.Fields("odrdttmpBackOrderQty").Value = 0
	end if
		
	If gwqty > ordqty then gwqty =ordqty
	
	rst.Fields("odrdttmpGiftWrapQty").Value = gwqty
	
	On Error Resume Next
	rst.update
	
	rst.Fields("odrdttmpQuantity").Value = ordqty
	rst.update
	Call CloseObj (rst)
	
	If Err.number <> 0 Then Err.Clear

End Sub	'AdjustTmpOrderDetails



Function GetProdGiftWrapPrice(sProdID) '8/15
	GetProdGiftWrapPrice = 0
	GetProdGiftWrapPrice = getgiftwrapprice(sProdID)
	If GetProdGiftWrapPrice = "X" or GetProdGiftWrapPrice = "" then 
		GetProdGiftWrapPrice = 0
	elseIf GetProdGiftWrapPrice < 0 then 
		GetProdGiftWrapPrice = 0
	
	else
		GetProdGiftWrapPrice=  GetProdGiftWrapPrice 'in porgress
	End IF
End Function

'-------------------------------------------------------------------------------
'Purpose: Runs any javascript function
'Accepts: name of javascript
'Returns: nothing
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub JS(sfunction)
	sfunction = replace(sfunction,";","")
	Response.Write "<script language=" & chr(34) & "javascript" & chr(34) & " type=""text/javascript"">" & vbcrlf & vbcrlf
	Response.write sfunction & ";" & vbcrlf
	Response.Write "</SCRIPT>" & vbcrlf
End sub

'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckLogin

Dim iCustID,  iSessionID	
	    
	' Request Cookie for custID 
	iCustID		= custID_cookie
	iSessionID	= getCookie_SessionID
	    
	If Len(iCustID) = 0 or iSessionID <> SessionID Then			
		CheckLogin = 0 'no login 
	Else
		CheckLogin = 1 
	End if
	
End Function

'-------------------------------------------------------------------------------
'Purpose: deletes any items from the cart with qty = 0 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function DeleteBadItems 'b2

dim rst
dim sql
dim i
Dim plngID

	DeleteBadItems = 0

	sql= "Select * FROM sfTmpOrderDetails WHERE odrdttmpSessionID=" & SessionID
	
	Set rst = CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenKeySet, adLockOptimistic,adcmdText
	
	If rst.recordcount > 0 then
		rst.movefirst
		For i = 1 to rst.recordcount
			
			If  rst.Fields("odrdttmpQuantity").Value <= 0 then
				plngID = rst.Fields("odrdttmpID").Value
				rst.delete   'b2
				rst.update   'b2
				DeletetmpOrderDetailsAE plngID
				DeleteBadItems = 1	
			End if
		if not rst.eof then rst.movenext
		Next 
				
	end if
	CloseObj (rst)

End Function

Sub DeleteTmpOrderDetailsAE (byVal iTmpOrderDetailID)
	Call setDeleteOrder("odrdttmp", iTmpOrderDetailID)
End Sub




'************************************************************************************
'**********************  MULTI-TIER PRICING **********************************************
'************************************************************************************

'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:price for a single product based on volume discount (MTP)
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetMTPrice (strProductID,sProdPrice,lOrderId) 
Dim rst,sql
Dim totqty
dim diff	
	If lOrderId > 0 then 
		'get from order details
		sql = "Select SUM(odrdtQuantity) as totqty FROM sfOrderDetails WHERE odrdtProductID= '" & makeInputSafe(strProductID) & "' AND odrdtOrderId= " & makeInputSafe(lOrderID)
	else 
		'get from temp order details
		sql = "Select SUM(odrdttmpQuantity) as totqty FROM sftmpOrderDetails WHERE odrdttmpProductID= '" & makeInputSafe(strProductID) & "' AND odrdttmpSessionID= " & SessionID
	End If
	
	Set rst = CreateObject("ADODB.RecordSet")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	if rst.recordcount <= 0 then
		GetMTPrice = sProdPrice 'no mtp
		CloseObj (rst)
		exit function
	end if
	totqty = Trim(rst.Fields("totqty").Value & "")
	If Len(totqty) = 0 Then Exit Function
	CloseObj (rst)
	
	sql = "Select * FROM sfMTPrices WHERE mtprodid= '" & makeInputSafe(strProductID) & "' AND mtQUANTITY  <= " & makeInputSafe(totqty) & " ORDER BY mtValue DESC"
	Set rst = CreateObject("ADODB.RecordSet")
'	modified for Sandshot Software pricing levels
'	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	SQL = AdjustSQLPricingLevel(SQL,"SearchProduct")
	With rst
        .CursorLocation = 3 'adUseClient
		.Open SQL, cnn, adOpenStatic, adLockBatchOptimistic, adCmdText
	End With
	Call AdjustRecordPricingLevel(rst, "mtprices")
	
	If rst.recordcount <= 0 then  
		GetMTPrice = sProdPrice 'no mtp
	else
		rst.movefirst
		If rst("mtType") = "Amount" then
			GetMTPrice = cdbl(sProdPrice) - cdbl(rst("mtValue"))
		else
			diff = cdbl(sProdPrice) * (cdbl(rst("mtValue"))/100)
			GetMTPrice = cdbl(sProdPrice) - cdbl(diff) 
		end if
	End If
	
	If cdbl(GetMTPrice) > cdbl(sProdPrice) then 
		GetMTPrice = sProdPrice
	End IF
	
	
	CloseObj (rst)	
	If GetMTPrice < 0 then GetMTPrice = 0

End Function

Function GetMTPrice2(byVal strProductID, byVal lngQty, byVal dblBasePrice)

Dim pstrSQL
Dim pobjCmd
Dim pobjRS
Dim pdblPriceOut

Dim totqty
dim diff	
	
	pstrSQL = "Select Top 1 mtValue, mtType, mtPLValue From sfMTPrices where mtProdID=? AND mtQUANTITY<=? ORDER BY mtQUANTITY DESC"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtext = pstrSQL
		.Commandtype = adCmdText
		'.Commandtype = adCmdStoredProc
		Set .ActiveConnection = cnn

		.Parameters.Append .CreateParameter("mtProdID", adWChar, adParamInput, 50, checkFieldLength(strProductID, 50, 0))
		.Parameters.Append .CreateParameter("mtQUANTITY", adInteger, adParamInput, 4, lngQty)
		Set pobjRS = .Execute
		
		pdblPriceOut = CDbl(dblBasePrice)
		If Not pobjRS.EOF Then
			If pobjRS.Fields("mtType").Value = "Amount" then
				pdblPriceOut = pdblPriceOut - CDbl(GetPricingLevelPrice(pobjRS.Fields("mtValue").Value, pobjRS.Fields("mtPLValue").Value))
			Else
				pdblPriceOut = pdblPriceOut * (1 - CDbl(GetPricingLevelPrice(pobjRS.Fields("mtValue").Value, pobjRS.Fields("mtPLValue").Value))/100)
			End If
		End If
		Call closeObj(pobjRS)

	End With	'pobjCmd
	Call closeObj(pobjCmd)

	If pdblPriceOut < 0 Then pdblPriceOut = 0
	pdblPriceOut = Round(pdblPriceOut, 2)

	GetMTPrice2 = pdblPriceOut

End Function	'GetMTPrice2


'************************************************************************************
'**********************  INVENTORY TRACKING ************************************************
'************************************************************************************
	

'-------------------------------------------------------------------------------
'Purpose: checks all cart items availability
'Accepts: 
'Returns: 1 if all items in the cart are available, else returns 0
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckCartInventory

dim sql
dim rst
dim i
dim avlqty
dim pbytResult
	
	If Len(SessionID) = 0 then 
		pbytResult = 1
	Else
		Set rst = CreateObject("ADODB.RecordSet")		
		sql = gtmpSQL & " WHERE odrdttmpSessionID=" & SessionID
		rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
		if rst.recordcount > 0 then
			rst.movefirst
			For i = 1 to rst.recordcount
				'If rst("odrdttmpBackOrderQty") = 0 then 'skip back ordered items
				If rst("odrdttmpBackOrderQty") <= 0 then ' beta 2
					avlqty = GetAvailableqty(rst("odrdttmpproductid"),GetAttDetailID(rst("odrdttmpID"),"tmp"))
					If avlqty <> "X" AND avlqty < rst("odrdttmpQuantity") then
						pbytResult = 0 'cart items need adjustment 
						Exit For
					End If
				End if	
				
				If not rst.eof then rst.movenext
			Next
				
			pbytResult = 1 'cart items ok 	
		
		Else
			pbytResult = 1
		End if
		
		CloseObj (rst)
	End If
	'Response.Write "CheckCartInventory: " & pbytResult & "<br />"

	CheckCartInventory = pbytResult

End Function	'CheckCartInventory

'************************************************************************************

Sub ValidateCartItems
	
dim sql
dim rst
dim i
	
	Set rst = CreateObject("ADODB.RecordSet")		
	sql = gtmpSQL & " WHERE odrdttmpSessionID=" & SessionID
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	If rst.recordcount > 0 then
		rst.movefirst
		For i = 1 to rst.recordcount
			AdjustTmpOrderDetails(rst.Fields("odrdttmpID").Value)
			If not rst.eof then rst.movenext
		Next
		
	End if
	
	CloseObj (rst)

End Sub

'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub UpdateAvailableQty(strProductID,AttIDs,subQTY)
Dim sql, rst
dim ret,avlqty
Dim bLow
Dim nProdID
Dim nSubject
Dim nBody
Dim nProdName
	ret = CheckInventoryTracked (strProductID)
	
	bLow = 0
 
    'Response.write "<br /> subqty:" & subqty
    'Response.write "<br /> prod-att:" & strProductID & "-" &  attids
    'Response.write "<br /> ret:" & ret
	
	If ret = 0 then Exit sub 'no inventory tracked so exit
	
	sql = "Select * FROM sfInventory WHERE invenProdID= '" & makeInputSafe(strProductID) & "' AND  invenAttDetailID='" & makeInputSafe(AttIDs) & "'"
	Set rst = CreateObject("ADODB.RecordSet")		
	'rst.CursorLocation = adUseClient
	rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	'rst.Open sql, cnn, adOpenDynamic, adLockOptimistic, adCmdText		

	if rst.recordcount <= 0 then
		closeobj(rst)
		exit sub
	end if


	avlqty = rst("invenInstock") 
	'Response.write "<br /> avlqty:" & avlqty
	'Response.write "<br /> rst(instock) after diff:" & rst("invenInstock")
	If (avlqty - subqty) < 0 then
		rst("invenInstock") = 0
		'rst("invenInstock") = (avlqty - subqty)
	else
		rst("invenInstock") = (avlqty - subqty)
	End if
		
	If rst("invenInstock") <= rst("invenLowFlag") then 
		bLow = 1
				
		nProdID = rst("InvenProdID")
		nProdName = GetProductName(nProdId)
		
		If  rst("invenInstock") <= 0 then
			nSubject = "Product Out of Stock!"
			nBody = vbcrlf & "Store: " & C_STORENAME
			nBody = nBody & vbcrlf & "Notification: Product Out of Stock!"  
			nBody = nBody & vbcrlf & "Product: " & nProdName
			nBody = nBody & vbcrlf & "Product Attributes: " & rst("invenAttName")
		
		Else 'qty is low 
			nSubject = "Product Stock Low!"
			nBody = vbcrlf & "Store: " & C_STORENAME
			nBody = nBody & vbcrlf & "Notification: Product Stock Low!"  
			nBody = nBody & vbcrlf & "Product: " & nProdName
			nBody = nBody & vbcrlf & "Product Attributes: " & rst("invenAttName")
			nBody = nBody & vbcrlf & "Quantity Remaining:" &  rst("invenInStock") 
		End If
		
	End if
'	DoEvents
	rst.update
	CloseObj (rst)
	
	
	'send notification if low
	If bLow = 1 AND CStr(CheckNotification(nProdID)) = "1"  then
	'If bLow =1 AND CheckNotification(nProdID)  = 1  then
		'sProdInfo = sProdInfo  & " " & GetProductName(strProductID)
		'Session("sNotification") = sProdInfo
		CreateMail "InvenNotification",nSubject & "|" & nBody 
	End if
	
	
	
	
End Sub



'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub sfReports1_ShowCouponDiscount

End Sub

Sub sfReports1_SQL1
	If cblnSF5AE Then sSQL = "Select * FROM sfOrders as A LEFT JOIN sfOrdersAE as B ON A.orderID = B.orderaeID " & "WHERE orderID = " & makeInputSafe(sOrderId)  & " and A.orderIsComplete = 1"
End Sub

Sub sfReports1_SQL2
	If cblnSF5AE Then sSQL = "Select * FROM sfOrderDetails as A LEFT JOIN sfOrderDetailsAE as B ON A.odrdtID = B.odrdtaeID " & "WHERE odrdtOrderId = " & makeInputSafe(sOrderId)
End Sub

Sub sfReports1_sql3
	If cblnSF5AE Then sSQL = "Select * FROM sfOrderDetails as A LEFT JOIN sfOrderDetailsAE as B ON A.odrdtID = B.odrdtaeID " & "WHERE odrdtOrderId = " & rsOrderDetail.Fields("orderID")
End Sub

Sub sfReports1_ShowProductDetails

Dim plngBOQty
Dim plngGWQty
Dim pdblGWPrice

	If Not cblnSF5AE Then Exit Sub

	plngBOQty = Trim(rsOrderProducts.Fields("odrdtBackOrderQTY").Value & "")
	If Len(plngBOQty) > 0 And isNumeric(plngBOQty) Then
		plngBOQty =  cDbl(plngBOQty)
	Else
		plngBOQty =  0
	End If
	
	plngGWQty = Trim(rsOrderProducts.Fields("odrdtGiftWrapQTY").Value & "")
	If Len(plngGWQty) > 0 And isNumeric(plngGWQty) Then
		plngGWQty =  cDbl(plngGWQty)
	Else
		plngGWQty =  0
	End If
	
	pdblGWPrice = Trim(rsOrderProducts.Fields("odrdtGiftWrapPrice").Value & "")
	If Len(pdblGWPrice) > 0 And isNumeric(pdblGWPrice) Then
		pdblGWPrice =  cDbl(pdblGWPrice)
	Else
		pdblGWPrice =  0
	End If
	
	Response.Write "<tr>" & vbcrlf
	Response.Write "<td></td>" & vbcrlf
	Response.Write "<td valign=""top"" align=""left"">Gift Wrap</td>" & vbcrlf
	Response.Write "<td valign=""top"" align=""right"">" & plngGWQty & "</td>" & vbcrlf

	If plngGWQty <> 0 Then 
		Response.Write "<td valign=""top"" align=""right"">" & FormatCurrency(pdblGWPrice/ plngGWQty) & "</td>" & vbcrlf
	else
   		Response.Write "<td valign=""top"" align=""right"">" & FormatCurrency(pdblGWPrice) & "</td>" & vbcrlf
	end if
	
	Response.Write "<td valign=""top"" align=""right"">" & FormatCurrency(pdblGWPrice) & "</td>" & vbcrlf
	Response.Write "</tr>" & vbcrlf
	Response.Write "<tr>" & vbcrlf
	Response.Write "<td> </td>" & vbcrlf
	Response.Write "<td valign=""top"" colspan = ""4"" align=""left"">BackOrdered Quantity: " & plngBOQty & "</td>" & vbcrlf
	Response.Write "</tr>" & vbcrlf

	If Err.number <> 0 Then Err.Clear
	
End Sub	'sfReports1_ShowProductDetails

'***************************************************************************************************************************************************************************

Sub sfReports1_Coupon
On Error REsume Next
Dim sql, rst

	If Not cblnSF5AE Then Exit Sub
   
	sql = "Select * FROM sfOrdersAE WHERE orderAEID= " & makeInputSafe(sOrderID)
	Set rst = CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	If rst.recordcount > 0 then

		Response.Write "<tr>" & vbcrlf
		Response.Write "<td>Coupon Code:&nbsp;&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "<td align=""right"">" & rst("orderCouponCode") & "&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		Response.Write "<tr>" & vbcrlf
		Response.Write "<td>Coupon Discount:&nbsp;&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "<td align=""right"">- " & FormatCurrency(rst("orderCouponDiscount")) & "&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		

		iTempDisc = iTempDisc - cDbl(rst("orderCouponDiscount"))
	End If 
	
	rst.close
	set rst = nothing
End Sub

'***************************************************************************************************************************************************************************

Sub sfReports1_Billing
Dim sql, rst

	If Not cblnSF5AE Then Exit Sub

	sql = "Select * FROM sfOrdersAE WHERE orderAEID= " & makeInputSafe(sOrderID)
	Set rst = CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	If rst.recordcount > 0 then
		

		Response.Write "<tr>" & vbcrlf
		Response.Write "<td><B>Billed Amount:&nbsp;&nbsp;&nbsp;</B></td>" & vbcrlf
		Response.Write "<td align=""right"">" & FormatCurrency(rst("orderBillAmount")) & "&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		Response.Write "<tr>" & vbcrlf
		Response.Write "<td><b>Remaining Amount:&nbsp;&nbsp;&nbsp;</B></td>" & vbcrlf
		Response.Write "<td align=""right"">" & FormatCurrency(rst("orderBackOrderAmount")) & "&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		

	End If 
	rst.close
	set rst = nothing
End Sub

'***************************************************************************************************************************************************************************

Sub OVC_AddProductGiftWrapPrice
		dProductSubtotal = cdbl(dProductSubtotal) + cdbl(Session("gwprice"))
		sTotalPrice = cdbl(sTotalPrice) + cdbl(Session("gwprice"))
End Sub

'***************************************************************************************************************************************************************************

%>