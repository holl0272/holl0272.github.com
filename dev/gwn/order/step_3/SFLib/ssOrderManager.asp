<!--#include file="ssShippingMethods_common.asp"-->
<%
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   2.00.002												*
'*   Release Date:		December 6, 2002										*
'*   Revision Date:     May 8, 2004												*
'*                                                                              *
'*   2.00.002 (May 8, 2004)														*
'*   - SQL Injection review                                                     *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	'Order Status Messages
	'This section must match the section in ssOrderAdmin_common
	Dim maryOrderStatuses(7)
	maryOrderStatuses(0) = "Order Placed, Awaiting Payment"
	maryOrderStatuses(0) = ""	'this is set to empty so as not to display anything unless explicity set
	maryOrderStatuses(1) = "Payment Received, Awaiting Shipment"
	maryOrderStatuses(2) = "Payment Received, Will Ship when payment clears"
	maryOrderStatuses(3) = "Awaiting Payment, Awaiting Shipment"
	maryOrderStatuses(4) = "Awaiting Payment, Order Shipped"
	maryOrderStatuses(5) = "Order Shipped"
	maryOrderStatuses(6) = "Order Complete"
	maryOrderStatuses(7) = "Order Cancelled"
	
	Const mblnShowOrderSummaries = True	'True 
	Const cstrAwaitPayMsg = "Awaiting Payment"

'/
'/////////////////////////////////////////////////

'***********************************************************************************************

Function getOrderStatusMessage(byVal lngOrderStatus)

Dim plngStatus

	plngStatus = Trim(lngOrderStatus & "")
	If Len(plngStatus) = 0 Or Not isNumeric(plngStatus) Then plngStatus = 0

	If UBound(maryOrderStatuses) >= CLng(plngStatus) Then 
		getOrderStatusMessage = maryOrderStatuses(plngStatus)
	Else
		getOrderStatusMessage = maryOrderStatuses(0)
	End If
	
End Function	'getOrderStatusMessage

'***********************************************************************************************

Sub ShowLogin(strEmail, lngOrderID, strMessage, bytDisplay)
%>
<script language="javascript" type="text/javascript">

function isInteger(theField, emptyOK, theMessage)
{

return true;

  if (theField.value == "")
  {
	if (emptyOK)
	{
		return(true);
	}
	{
		alert(theMessage);
		theField.focus();
		theField.select();
	  return (false);
	}
  }

    var i;
    var s = theField.value;
    for (i = 0; i < s.length; i++)
    {
        var c = s.charAt(i);
        if (!((c >= "0") && (c <= "9")))
        {
			alert(theMessage);
			theField.focus();
			theField.select();
            return (false);
        }
    }

  return (true);
}

function Validate(theForm)
{
if (theForm.email.value == "")
{
alert("You must enter an email address.");
theForm.email.focus();
return false;
}

<% If bytDisplay = 0 Then %>
if (theForm.OrderID.value == "")
{
alert("Please enter a Order Number.");
theForm.OrderID.focus();
return false;
}
<% ElseIf bytDisplay = 1 Then %>
if (theForm.password.value == "")
{
alert("You must enter a password.");
theForm.password.focus();
return false;
}
<% Else %>
if ((theForm.OrderID.value == "") && (theForm.password.value == ""))
{
alert("You must enter either an Order Number or a password.");
theForm.password.focus();
return false;
}
<% End If %>

return true;

}

</script>
<FORM action="OrderHistory.asp" method=post id=form1 name=form1 onsubmit="return Validate(this);">
<INPUT type=HIDDEN name=Action value=Login>
<table border=0 cellspacing=0>
<colgroup align=right>
<colgroup align=left>  
<colgroup align=left>  
  <TR>
    <TH colSpan=3 align=center>Customer Login</TH>
  </TR>
  <TR>
    <TD colSpan=3 align=center><font color=red><b><%= strMessage %></b></font></TD>
  </TR>
  <TR>
    <TD><LABEL for=email>email 
      address:</LABEL></TD>
    <TD><INPUT name=email value="<%= strEmail %>"></TD>
    <TD><font color=red size=-1><i>(*required)</i></font></TD>
  </TR>
<% If bytDisplay = 0 Then %>
  <TR>
    <TD><LABEL for=OrderID>Order Number:</LABEL></TD>
    <TD><INPUT name=OrderID value="<%= lngOrderID %>" onblur="return isInteger(this, true, 'Please enter a number');"></TD>
    <TD></TD>
  </TR>
<% ElseIf bytDisplay = 1 Then %>
  <INPUT type=hidden name=OrderID value="<%= lngOrderID %>">
  <TR>
    <TD><LABEL for=password>password:</LABEL></TD>
    <TD><INPUT type=password name=password></TD>
    <TD><font color=red size=-1><i>(*required to view order history)</i></font></TD>
  </TR>
<% Else %>
  <TR>
    <TD><LABEL for=OrderID>Order Number:</LABEL></TD>
    <TD><INPUT name=OrderID value="<%= lngOrderID %>" onblur="return isInteger(this, true, 'Please enter a number');"></TD>
    <TD></TD>
  </TR>
  <TR>
    <TD><LABEL for=password>password:</LABEL></TD>
    <TD><INPUT type=password name=password></TD>
    <TD><font color=red size=-1><i>(*required to view order history)</i></font></TD>
  </TR>
<% End If %>
  <TR>
	<TD></TD>
	<TD align="center"><input type="image" class="inputImage" src="<%= C_BTN16 %>" name=btnLogin></TD>
  </TR>
</TABLE>
</FORM>
<%
End Sub	'ShowLoginForm

'***********************************************************************************************
	
Function checkInput(byRef strEmail, byRef strPassword)
'Purpose: protect username (email address) and password from SQL Injection attacks

	checkInput = fncEmailValid(strEmail) And validatepassword(strPassword)

End Function	'checkInput

'Generic function to validate strings using regular expressions
Function fncStringValid(strInput, strPattern)

Dim MyRegExp

	Set MyRegExp = New RegExp
	MyRegExp.Pattern = strPattern
	fncStringValid = MyRegExp.Test(strInput)

End Function

Function fncEmailValid(strInput)
	fncEmailValid = fncStringValid(strInput, "^[A-Za-z0-9](([_\.\-]?[a-zA-Z0-9]+)*)@([A-Za-z0-9]+)(([\.\-]?[a-zA-Z0-9]+)*)\.([A-Za-z]{2,})$")
End Function

Function validatepassword(strPassword)

Dim good_password_chars
Dim i
Dim c

	good_password_chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789" 
	
	validatepassword = true
	
	For i = 1 to len(strPassword)
		c = mid(strPassword, i, 1 )
		if (InStr(good_password_chars, c ) = 0) then
			validatepassword = false
			exit function
		end if
	next
	
End Function	'validatepassword

'***********************************************************************************************
	
Function Login(byVal strEmail, byVal strPassword, byVal lngOrderID, byRef strMessage)
'returns login status

Dim pobjRS
Dim pstrSQL

	If Not checkInput(strEmail, strPassword) Then
		strMessage = "There was a problem with your username/password. Please contact the system administrator."
		Login = 0
		Exit Function
	ElseIf Len(lngOrderID) > 0  AND Not isNumeric(lngOrderID) Then
		strMessage = "You have entered an invalid order number."
		Login = 0
		Exit Function
	End If

	If Len(strEmail) > 0 Then
		If Len(strPassword) > 0 Then
			Set mclsLogin = New clsLogin
			strMessage = mclsLogin.ValidUserName(mstrEmail, mstrPassword)
			Set mclsLogin = Nothing
				If strMessage = "True" Or Len(mstrLoginMessage) = 0 Then
					Login = 2
				Else
					Login = 0
				End If
		ElseIf Len(lngOrderID) > 0 Then
			If isNumeric(lngOrderID) Then
				pstrSQL = "SELECT orderID FROM sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId where custEmail='" & strEmail & "' AND orderID=" & lngOrderID
				Set pobjRS = CreateObject("ADODB.RECORDSET")
				pobjRS.Open pstrSQL,cnn
				If pobjRS.EOF Then
					strMessage = strMessage & "I could not locate an order for the specified email address.<br />Please verify your entries."
					Login = 0
				Else
'					Session("ssLoginStatus") = 1
					Login = 1
				End If
				pobjRS.Close
				Set pobjRS = Nothing
			Else
				strMessage = strMessage & "You must enter a numeric value for the order number."
				Login = 0
			End If
		Else
			strMessage = strMessage & "You must enter a password or order number."
			Login = 0
		End If
	Else
		strMessage = strMessage & "You must enter an email address."
		Login = 0
	End If
	
End Function	'Login

'***********************************************************************************************

Function LoadOrderHistory(byVal lngCustomerID, byRef objRS)

Dim pstrSQL

	If Len(lngCustomerID) = 0 Then
		LoadOrderHistory = False
	ElseIf Not isNumeric(lngCustomerID) Then
		LoadOrderHistory = False
	Else
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderGrandTotal, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, sfOrders.orderStatus" _
				& " FROM sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
				& " GROUP BY sfOrders.orderID, sfOrders.orderDate, sfOrders.orderGrandTotal, sfOrders.orderIsComplete, sfOrders.orderCustId, sfOrders.orderStatus" _
				& " HAVING sfOrders.orderCustId=" & lngCustomerID _
				& " ORDER BY sfOrders.orderDate DESC"
				
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderGrandTotal, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, ssOrderManager.ssOrderStatus" _
				& " FROM ssOrderManager RIGHT JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON ssOrderManager.ssorderID = sfOrders.orderID" _
				& " GROUP BY sfOrders.orderID, sfOrders.orderDate, sfOrders.orderGrandTotal, sfOrders.orderIsComplete, sfOrders.orderCustId, ssOrderManager.ssOrderStatus" _
				& " HAVING sfOrders.orderIsComplete=1 AND sfOrders.orderCustId=" & lngCustomerID _
				& " ORDER BY sfOrders.orderDate DESC"
				
				'requirement removed for completed orders since some are completed but not recorded properly
				'& " HAVING sfOrders.orderIsComplete=1 AND sfOrders.orderCustId=" & lngCustomerID _
				'& " HAVING sfOrders.orderCustId=" & lngCustomerID _

		Set objRS = CreateObject("ADODB.RECORDSET")
		objRS.CursorLocation = 2 'adUseClient
		objRS.CursorType = 3 'adOpenStatic
		objRS.LockType = 1 'adLockReadOnly
		objRS.Open pstrSQL,cnn
		
		LoadOrderHistory = Not objRS.EOF
	End If

End Function	'LoadOrderHistory

'***********************************************************************************************

Sub ShowOrderDetail(lngOrderID) 

Dim pstrSQL
Dim pobjRS
Dim pcurRealSubTotal
Dim pstrCustName
Dim pstrCustAddr
Dim pstrProdID
Dim pblnOddRow
Dim pstrBackground
Dim plngodrdtID
Dim pstrOrderStatusMsg
Dim pstrAttributeName
Dim pblnAEDisplayed

	pblnOddRow = False

	If Len(lngOrderID) > 0 And isNumeric(lngOrderID) Then

		If cblnSF5AE Then
			pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderAmount, sfOrders.orderComments, sfOrders.orderShipMethod, sfOrders.orderCTax, sfOrders.orderSTax, sfOrders.orderShippingAmount, sfOrders.orderHandling, sfOrders.orderGrandTotal, sfOrders.orderPaymentMethod, sfOrders.orderCheckAcctNumber, sfOrders.orderCheckNumber, sfOrders.orderBankName, sfOrders.orderRoutingNumber, sfOrders.orderPurchaseOrderName, sfOrders.orderPurchaseOrderNumber, sfOrders.orderStatus, sfOrders.orderProcessed, sfOrders.orderTracking, sfOrders.orderIsComplete, " _
					& "       sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtPrice, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, " _
					& "       sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, " _
					& "       sfCPayments.payCardType, sfCPayments.payCardName, sfCPayments.payCardNumber, sfCPayments.payCardExpires, " _
					& "       sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custLastName, sfCustomers.custCompany, sfCustomers.custAddr1, sfCustomers.custAddr2, sfCustomers.custCity, sfCustomers.custState, sfCustomers.custZip, sfCustomers.custCountry, sfCustomers.custPhone, sfCustomers.custFax, sfCustomers.custEmail, " _
					& "       sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipMiddleInitial, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipPhone, sfCShipAddresses.cshpaddrShipFax, sfCShipAddresses.cshpaddrShipEmail, " _
					& "       ssOrderManager.ssExternalNotes, ssOrderManager.ssInternalNotes, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssShippedVia, ssOrderManager.ssTrackingNumber, ssOrderManager.ssOrderStatus, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssDateEmailSent, ssOrderManager.ssPaidVia, ssOrderManager.ssExported, ssOrderManager.ssOrderFlagged, ssOrderManager.ssBackOrderDateNotified, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssBackOrderMessage, ssOrderManager.ssBackOrderInternalMessage, ssOrderManager.ssBackOrderTrackingNumber, " _
					& "       sfOrdersAE.orderCouponCode, sfOrdersAE.orderBillAmount, sfOrdersAE.orderBackOrderAmount, sfOrdersAE.orderCouponDiscount, " _
					& "       sfOrderDetailsAE.odrdtGiftWrapPrice, sfOrderDetailsAE.odrdtGiftWrapQTY, sfOrderDetailsAE.odrdtBackOrderQTY, sfOrderDetailsAE.odrdtAttDetailID" _
					& " FROM sfOrderDetailsAE RIGHT JOIN (sfOrdersAE RIGHT JOIN ((((((sfCustomers RIGHT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) LEFT JOIN sfCShipAddresses ON sfOrders.orderAddrId = sfCShipAddresses.cshpaddrID) LEFT JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) ON sfOrdersAE.orderAEID = sfOrders.orderID) ON sfOrderDetailsAE.odrdtAEID = sfOrderDetails.odrdtID"
		Else
			pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, sfOrders.orderAmount, sfOrders.orderComments, sfOrders.orderShipMethod, sfOrders.orderCTax, sfOrders.orderSTax, sfOrders.orderShippingAmount, sfOrders.orderHandling, sfOrders.orderGrandTotal, sfOrders.orderPaymentMethod, sfOrders.orderCheckAcctNumber, sfOrders.orderCheckNumber, sfOrders.orderBankName, sfOrders.orderRoutingNumber, sfOrders.orderPurchaseOrderName, sfOrders.orderPurchaseOrderNumber, sfOrders.orderStatus, sfOrders.orderProcessed, sfOrders.orderTracking, sfOrders.orderIsComplete, " _
					& "       sfOrderDetails.odrdtID, sfOrderDetails.odrdtProductID, sfOrderDetails.odrdtProductName, sfOrderDetails.odrdtPrice, sfOrderDetails.odrdtQuantity, sfOrderDetails.odrdtSubTotal, " _
					& "       sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName, " _
					& "       sfCPayments.payCardType, sfCPayments.payCardName, sfCPayments.payCardNumber, sfCPayments.payCardExpires, " _
					& "       sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custLastName, sfCustomers.custCompany, sfCustomers.custAddr1, sfCustomers.custAddr2, sfCustomers.custCity, sfCustomers.custState, sfCustomers.custZip, sfCustomers.custCountry, sfCustomers.custPhone, sfCustomers.custFax, sfCustomers.custEmail, " _
					& "       sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipMiddleInitial, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipPhone, sfCShipAddresses.cshpaddrShipFax, sfCShipAddresses.cshpaddrShipEmail, " _
					& "       ssOrderManager.ssExternalNotes, ssOrderManager.ssInternalNotes, ssOrderManager.ssDatePaymentReceived,ssOrderManager.ssShippedVia, ssOrderManager.ssTrackingNumber, ssOrderManager.ssOrderStatus, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssDateEmailSent, ssOrderManager.ssPaidVia, ssOrderManager.ssExported, ssOrderManager.ssOrderFlagged, ssOrderManager.ssBackOrderDateNotified, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssBackOrderMessage, ssOrderManager.ssBackOrderInternalMessage, ssOrderManager.ssBackOrderTrackingNumber " _
					& " FROM (((((sfCustomers RIGHT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) LEFT JOIN sfCShipAddresses ON sfOrders.orderAddrId = sfCShipAddresses.cshpaddrID) LEFT JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID"
		End If
		
		If Len(mstrEmail) > 0 Then
			pstrSQL = pstrSQL & " WHERE orderID=" & lngOrderID & " AND sfCustomers.custEmail='" & mstrEmail & "'"
		ElseIf Len(visitorLoggedInCustomerID) > 0 And isNumeric(visitorLoggedInCustomerID) Then
			pstrSQL = pstrSQL & " WHERE orderID=" & lngOrderID & " AND sfOrders.orderCustId=" & visitorLoggedInCustomerID
		Else
			pstrSQL = pstrSQL & " WHERE orderID=" & lngOrderID
		End If

		Select Case 1
			Case 0:	'Sort by product name only
					pstrSQL = pstrSQL & " Order By odrdtProductName Asc, odrattrOrderDetailId, odrattrID"	'odrattrOrderDetailId necessary to keep the attributes in the right place
			Case 1: 'Sort by product name, attribute name
					'pstrSQL = pstrSQL & " Order By odrdtProductName Asc, odrattrOrderDetailId, odrattrName"	'odrattrOrderDetailId necessary to keep the attributes in the right place
					pstrSQL = pstrSQL & " Order By odrdtProductName Asc, odrattrOrderDetailId, odrattrID"	'odrattrOrderDetailId necessary to keep the attributes in the right place
			Case Else	'Sort by the order they were added to the cart
		End Select
		
		'debugprint "pstrSQL",pstrSQL 
		Set pobjRS = CreateObject("ADODB.RECORDSET")
		pobjRS.CursorLocation = 2 'adUseClient
		pobjRS.CursorType = 3 'adOpenStatic
		pobjRS.LockType = 1 'adLockReadOnly
		pobjRS.Open pstrSQL,cnn

If pobjRS.EOF Then
	Response.Write "<h4>Invalid Order Number</h4>"
Else
		pstrCustName = Replace(pobjRS.Fields("custFirstName").Value & " " & pobjRS.Fields("custMiddleInitial").Value & " " & pobjRS.Fields("custLastName").Value,"  "," ")
		pstrCustAddr = pobjRS.Fields("custCity").Value & ", " & pobjRS.Fields("custState").Value & " " & pobjRS.Fields("custZip").Value
%>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <colgroup align=left width=50%>
  <colgroup align=left width=50%>
  <tr>
    <td width="100%" colspan="2">Order Number:&nbsp;&nbsp; <%=lngOrderID %></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <table border="0" cellspacing="0" cellpadding="3" width="100%">
		<colgroup>
		  <col align="center" />
		  <col align="left" />
		  <col align="center" />
		  <col align="center" />
		  <col align="center" />
		</colgroup>
        <tr class="tdContentBar">
          <th align=center>Product ID</th>
          <th align=center>Product Name</th>
          <th align=center>Quantity</th>
          <th align=center>Unit Price</th>
          <th align=center>Extended Price</th>
        </tr>
<%
pcurRealSubTotal = 0

On Error Resume Next
Do While Not pobjRS.EOF

	If plngodrdtID <> Trim(pobjRS.Fields("odrdtID").Value) Then
'	If pstrProdID <> Trim(pobjRS.Fields("odrdtProductID").Value) Then
		plngodrdtID = Trim(pobjRS.Fields("odrdtID").Value)
		pstrProdID = Trim(pobjRS.Fields("odrdtProductID").Value)
		pblnOddRow = Not pblnOddRow
		If pblnOddRow Then
			pstrBackground = "tdAltFont1"		'set this color for the odd rows
		Else
			pstrBackground = "tdAltFont2"	'set this color for the even rows
		End If
		pblnAEDisplayed = False

%>
        <tr class="<%= pstrBackground %>">
          <td align="center"><%= pstrProdID %></td>
          <td align="left"><%= pobjRS.Fields("odrdtProductName").Value %></td>
          <td align="center"><%= pobjRS.Fields("odrdtQuantity").Value %></td>
          <td align="center"><%= FormatCurrency(pobjRS.Fields("odrdtPrice").Value,2) %></td>
          <td align="center"><%= FormatCurrency(pobjRS.Fields("odrdtSubTotal").Value,2) %></td>
        </tr>
<%
		pcurRealSubTotal = pcurRealSubTotal + CDbl(pobjRS.Fields("odrdtSubTotal").Value)
	End If

	pstrAttributeName = Trim(pobjRS.Fields("odrattrName").Value & " " & pobjRS.Fields("odrattrAttribute").Value)
	If Len(pstrAttributeName) > 0 Then
		If InStr(1, pobjRS.Fields("odrattrName").Value, Left(pobjRS.Fields("odrattrAttribute").Value, Len(pobjRS.Fields("odrattrName").Value))) > 0 Then 
			pstrAttributeName = Replace(pstrAttributeName, pobjRS.Fields("odrattrName").Value, "", 1, 1)
		Else
			pstrAttributeName = Trim(pobjRS.Fields("odrattrName").Value & ": " & pobjRS.Fields("odrattrAttribute").Value)
		End If
%>
        <tr class="<%= pstrBackground %>">
          <td></td>
          <td><font size="-1">&nbsp;&nbsp;<%= pstrAttributeName %></font></td>
          <td></td>
          <td></td>
          <td></td>
        </tr>
<%	
	End If

	If cblnSF5AE And Not pblnAEDisplayed Then
		'Check for Gift Wrap
		If pobjRS.Fields("odrdtGiftWrapQTY").Value > 0 Then %>
			<tr class="<%= pstrBackground %>">
			  <td></td>
			  <td align="left">&nbsp;<i>Gift Wrap</i></td>
			  <td align="center"><%= pobjRS.Fields("odrdtGiftWrapQTY").Value %></td>
			  <td align="center"><%= FormatCurrency(pobjRS.Fields("odrdtGiftWrapPrice").Value/pobjRS.Fields("odrdtGiftWrapQTY").Value,2) %></td>
			  <td align="center"><%= FormatCurrency(pobjRS.Fields("odrdtGiftWrapPrice").Value,2) %></td>
			</tr>
<%			pcurRealSubTotal = pcurRealSubTotal + CDbl(pobjRS.Fields("odrdtGiftWrapPrice").Value)
		End If

		'Check for Back Orders
		If pobjRS.Fields("odrdtBackOrderQTY").Value > 0 Then %>
			<tr class="<%= pstrBackground %>">
			  <td></td>
			  <td colspan=4 align="left">&nbsp;<b><i>Quantity on Back Order: <%= pobjRS.Fields("odrdtBackOrderQTY").Value %></i></b></td>
			</tr>
<%		End If

		pblnAEDisplayed= True
	End If

	pobjRS.MoveNext
Loop
pobjRS.MoveFirst
%>
        <tr>
          <td width="100%" colspan="5">&nbsp;</td>
        </tr>
        <tr>
          <td colspan=3 align=left valign=bottom>&nbsp;&nbsp;</td>
          <td width="40%" colspan="2" align=center>
<%
Dim mcurDiscount

Dim mcurCoupon: mcurCoupon = 0
If cblnSF5AE Then
	If isNumeric(pobjRS.Fields("orderCouponDiscount").Value) Then mcurCoupon = CDbl(pobjRS.Fields("orderCouponDiscount").Value)
End If

mcurDiscount = Round(Abs(CDbl(pobjRS.Fields("orderGrandTotal").Value) - pcurRealSubTotal + mcurCoupon - CDbl(pobjRS.Fields("orderHandling").Value) - CDbl(pobjRS.Fields("orderCTax").Value) - CDbl(pobjRS.Fields("orderSTax").Value) - CDbl(pobjRS.Fields("orderShippingAmount").Value)), 2)
%>
            <table border="0" cellspacing="0" cellpadding="3">
			  <colgroup align=left>
			  <colgroup align=right>
              <tr>
                <td>Subtotal:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                <td><%= FormatCurrency(pcurRealSubTotal,2) %></td>
              </tr>
			  <% 
			  If cblnSF5AE Then
				If Len(pobjRS.Fields("orderCouponDiscount").Value & "") > 0 Then 
				  If CDbl(pobjRS.Fields("orderCouponDiscount").Value) > 0 Then 
			  %>
              <tr>
                <td>Coupon (<%= pobjRS.Fields("orderCouponCode").Value %>):</td>
                <td><%= FormatCurrency(-1 * pobjRS.Fields("orderCouponDiscount").Value,2) %></td>
              </tr>
			  <% 
				  End If 
				End If 
			  End If 
			  %>
			  <% If CDbl(mcurDiscount) < -0.009 Then %>
              <tr>
                <td>Discount:</td>
                <td><%= FormatCurrency(mcurDiscount,2) %></td>
              </tr>
			  <% End If %>
			  <% If CDbl(pobjRS.Fields("orderSTax").Value & "") > 0 Then %>
              <tr>
                <td><%= pobjRS.Fields("cshpaddrShipState").Value %> Tax:</td>
                <td><%= FormatCurrency(pobjRS.Fields("orderSTax").Value) %></td>
              </tr>
			  <% End If %>
			  <% If CDbl(pobjRS.Fields("orderCTax").Value & "") > 0 Then %>
              <tr>
                <td><%= pobjRS.Fields("cshpaddrShipCountry").Value %> Tax:</td>
                <td><%= FormatCurrency(pobjRS.Fields("orderCTax").Value) %></td>
              </tr>
			  <% End If %>
              <tr>
                <td><%= pobjRS.Fields("orderShipMethod").Value %>:</td>
                <td><%= FormatCurrency(pobjRS.Fields("orderShippingAmount").Value) %></td>
              </tr>
			  <% If CDbl(pobjRS.Fields("orderHandling").Value & "") > 0 Then %>
              <tr>
                <td>Handling:</td>
                <td><%= FormatCurrency(pobjRS.Fields("orderHandling").Value) %></td>
              </tr>
			  <% End If %>
              <tr>
                <td colspan="2">
                  <hr>
                </td>
              </tr>
              <tr>
                <td>Total:</td>
                <td><%= FormatCurrency(pobjRS.Fields("orderGrandTotal").Value) %></td>
              </tr>
            <%
            If False Then 
				If GC_LoadByOrder(lngOrderID, pobjRS.Fields("orderGrandTotal").Value) Then
            %>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
				<tr>
					<td align="left">Certificate (<%= mstrCertificate %>):</td>
					<td align="right"><%= FormatCurrency(mdblssCertificateAmount, 2) %></td>
				</tr>
				<tr>
					<td align="left">Amount Billed:</td>
					<td align="right"><%= FormatCurrency(mdblssGCNewTotalDue, 2) %></td>
				</tr>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
            <% 
				End If	'GC_LoadByOrder 
            End If	' 
            %>
            </table>
          </td>
        </tr>
		<% If Len(pobjRS.Fields("orderComments").Value & "") > 0 Then %>
        <tr>
          <td colspan=5 align=left>&nbsp;&nbsp;Special Instructions:&nbsp;&nbsp;<%= pobjRS.Fields("orderComments").Value %></td>
        </tr>
        <% End If %>
      </table>
    </td>
  </tr>
  <tr><th colspan="2">&nbsp;</th></tr>
  <tr>
    <th colspan="2">Customer Information</th>
  </tr>
  <tr>
    <td valign="top">
      <table border="0" cellspacing="0" cellpadding="3" width="100%">
        <tr>
          <td align="right">Sold to:&nbsp;&nbsp;</td>
          <td><%= pstrCustName %></td>
        </tr>
		<% If Len(pobjRS.Fields("custCompany").Value & "") > 0 Then %>
        <tr>
          <td align="right">Company:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("custCompany").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td align="right">Address:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("custAddr1").Value %></td>
        </tr>
		<% If Len(pobjRS.Fields("custAddr2").Value & "") > 0 Then %>
        <tr>
          <td align="right">&nbsp;</td>
          <td><%= pobjRS.Fields("custAddr2").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td align="right">&nbsp;</td>
          <td><%= pstrCustAddr %></td>
        </tr>
        <tr>
          <td align="right">&nbsp;</td>
          <td><%= pobjRS.Fields("custCountry").Value %></td>
        </tr>
		<% If Len(pobjRS.Fields("custPhone").Value & "") > 0 Then %>
        <tr>
          <td align="right">Phone:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("custPhone").Value %></td>
        </tr>
        <% End If %>
		<% If Len(pobjRS.Fields("custFax").Value & "") > 0 Then %>
        <tr>
          <td align="right">Fax:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("custFax").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td align="right">email:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("custEmail").Value %></td>
        </tr>
      </table>
    </td>
<%
pstrCustName = Replace(pobjRS.Fields("cshpaddrShipFirstName").Value & " " & pobjRS.Fields("cshpaddrShipMiddleInitial").Value & " " & pobjRS.Fields("cshpaddrShipLastName").Value,"  "," ")
pstrCustAddr = pobjRS.Fields("cshpaddrShipCity").Value & ", " & pobjRS.Fields("cshpaddrShipState").Value & " " & pobjRS.Fields("cshpaddrShipZip").Value
%>
    <td valign="top">
      <table border="0" cellspacing="0" cellpadding="3" width="100%">
        <tr>
          <td align="right">Shipped to:&nbsp;&nbsp;</td>
          <td><%= pstrCustName %></td>
        </tr>
		<% If Len(pobjRS.Fields("cshpaddrShipCompany").Value & "") > 0 Then %>
        <tr>
          <td align="right">Company:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("cshpaddrShipCompany").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td align="right">Address:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("cshpaddrShipAddr1").Value %></td>
        </tr>
		<% If Len(pobjRS.Fields("cshpaddrShipAddr2").Value & "") > 0 Then %>
        <tr>
          <td align="right">&nbsp;</td>
          <td><%= pobjRS.Fields("cshpaddrShipAddr2").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td align="right">&nbsp;</td>
          <td><%= pstrCustAddr %></td>
        </tr>
        <tr>
          <td align="right">&nbsp;</td>
          <td><%= pobjRS.Fields("cshpaddrShipCountry").Value %></td>
        </tr>
		<% If Len(pobjRS.Fields("cshpaddrShipPhone").Value & "") > 0 Then %>
        <tr>
          <td align="right">Phone:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("cshpaddrShipPhone").Value %></td>
        </tr>
        <% End If %>
		<% If Len(pobjRS.Fields("cshpaddrShipFax").Value & "") > 0 Then %>
        <tr>
          <td align="right">Fax:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("cshpaddrShipFax").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td align="right">email:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("cshpaddrShipEmail").Value %></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr><th colspan="2">&nbsp;</th></tr>
<% 
'Figure out the payment
Dim pstrPaymentType, pstrPaymentMsg
pstrPaymentType = Trim(pobjRS.Fields("orderPaymentMethod").Value)
If pstrPaymentType = "Credit Card" Then
'	Dim pobjRSPayment
'	Set pobjRSPayment = CreateObject("ADODB.RECORDSET")
'	With
'		.CursorLocation = 2 'adUseClient
'		.CursorType = 3 'adOpenStatic
'		.LockType = 1 'adLockReadOnly
'		.Open "Select ",cnn
'		.Close
'	End With
'	Set pobjRSPayment = Nothing

End If

If Len(Trim(pobjRS.Fields("ssOrderStatus").Value & "")) > 0 Then
	pstrOrderStatusMsg = maryOrderStatuses(pobjRS.Fields("ssOrderStatus").Value)
Else
	pstrOrderStatusMsg = maryOrderStatuses(0)
End If

If Len(pobjRS.Fields("ssDatePaymentReceived").Value & "") > 0 Then
	pstrPaymentMsg = "Payment received on " & FormatDateTime(pobjRS.Fields("ssDatePaymentReceived").Value,1)
Else
	pstrPaymentMsg = cstrAwaitPayMsg
End If

'Figure out the shipping
Dim pstrShippingType, pstrShippingMessage
Const cstrAwaitShipMsg = "Awaiting Shipment"
If isNull(pobjRS.Fields("ssDateOrderShipped").Value) Then
	If pstrPaymentMsg <> cstrAwaitPayMsg Then 
		pstrShippingMessage = cstrAwaitShipMsg
	Else
		pstrShippingMessage = ""
	End If
ElseIf pobjRS.Fields("ssDateOrderShipped").Value > DateAdd("h",0,Date()) Then
	pstrShippingType = "Order projected to be shipped on"
	pstrShippingMessage = FormatDateTime(pobjRS.Fields("ssDateOrderShipped").Value,1)
Else
	pstrShippingType = "Order shipped on"
	pstrShippingMessage = FormatDateTime(pobjRS.Fields("ssDateOrderShipped").Value,1)
End If

%>
  <tr>
    <th width="100%" colspan="2">Order Status</th>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <table border="0" cellspacing="0">
		<% If Len(pstrOrderStatusMsg) > 0 Then %>
        <tr>
          <td align="right">Status:&nbsp;&nbsp;</td>
          <td><%= pstrOrderStatusMsg %></td>
        </tr>
		<% End If %>
        <tr>
          <td align="right">Order Placed On:&nbsp;&nbsp;</td>
          <td><%= FormatDateTime(pobjRS.Fields("orderDate").Value,1) & " at " & FormatDateTime(pobjRS.Fields("orderDate").Value,3) %></td>
        </tr>
        <tr>
          <td align="right">Payment Type:&nbsp;&nbsp;</td>
          <td><%= pstrPaymentType %></td>
        </tr>
        <tr>
          <td align="right">Payment Status:&nbsp;&nbsp;</td>
          <td><%= pstrPaymentMsg %></td>
        </tr>
        <tr>
          <td align="right">Order&nbsp;Status:&nbsp;&nbsp;</td>
          <td><%= getOrderStatusMessage(pobjRS.Fields("ssorderStatus").Value) %></td>
        </tr>
		<% If Len(pstrShippingType) > 0 Then %>
        <tr>
          <td align="right"><%= pstrShippingType %>:&nbsp;&nbsp;</td>
          <td><%= pstrShippingMessage %></td>
        </tr>
		<% End If %>
		<% If Len(pobjRS.Fields("ssTrackingNumber").Value & "") > 0 Then %>
        <tr>
          <td align="right">Track this order:&nbsp;&nbsp;</td>
          <td><%= splitTrackingLinks(Trim(pobjRS.Fields("ssTrackingNumber").Value & ""), pobjRS.Fields("ssShippedVia").Value) %>&nbsp;</td>
        </tr>
		<% End If %>
		<% If Len(pobjRS.Fields("ssExternalNotes").Value & "") > 0 Then %>
        <tr>
          <td align="right">Notes:&nbsp;&nbsp;</td>
          <td><%= pobjRS.Fields("ssExternalNotes").Value %></td>
        </tr>
		<% End If %>

		<% If Len(pobjRS.Fields("ssBackOrderDateExpected").Value & "") > 0 Then %>
			<tr><th colspan="2">&nbsp;</th></tr>
			<tr>
			  <th width="100%" colspan="2" align="left">Back Order Information</th>
			</tr>
			<tr>
			  <td align="right">Date Expected In:&nbsp;&nbsp;</td>
			  <td align="left"><%= pobjRS.Fields("ssBackOrderDateExpected").Value %></td>
			</tr>
		<% End If %>
		<% If Len(pobjRS.Fields("ssBackOrderMessage").Value & "") > 0 Then %>
			<tr>
			  <td align="right">Back Order Message:&nbsp;&nbsp;</td>
			  <td align="left"><%= pobjRS.Fields("ssBackOrderMessage").Value %></td>
			</tr>
		<% End If %>
		<% If Len(pobjRS.Fields("ssBackOrderTrackingNumber").Value & "") > 0 Then %>
        <tr>
          <td>Track this shipment:&nbsp;&nbsp;</td>
          <td><%= splitTrackingLinks(Trim(pobjRS.Fields("ssBackOrderTrackingNumber").Value & ""), pobjRS.Fields("ssShippedVia").Value) %>&nbsp;</td>
        </tr>
		<% End If %>

        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
		pobjRS.Close
		Set pobjRS = Nothing
End If	
	Else	'check for Len(lngOrderID) > 0
%>
<table><tr><td>Select an Order Number below to view your order details</td></tr></table>
<%	
	End If
%>

<% End Sub 'ShowOrderDetail%>

<% Sub ShowOrderHistory(lngOrderID, blnLoggedIn, objrsOrders)

Dim pstrOrderLink
Dim pstrOrderTrackLink

	If Not mblnShowOrderSummaries Then Exit Sub

	'No reason to display order history if there is only the one order
	If isObject(objrsOrders) Then
		If objrsOrders.RecordCount > 0 Then 
			objrsOrders.MoveFirst
			If Len(lngOrderID) > 0 Then objrsOrders.Filter = "orderID<>" & lngOrderID
			If objrsOrders.EOF Then
				objrsOrders.Filter = ""
				Exit Sub
			End If
			objrsOrders.Filter = ""
		Else
			Exit Sub
		End If
	End If
 %>

<hr>
<table border="0" cellspacing="0" width="100%">
  <colgroup>
  <col width="10%" align="center" />
  <col width="20%" align="center" />
  <col width="15%" align="center" />
  <col width="15%" align="center" />
  <col width="25%" align="center" />
  <col width="15%" align="center" />
  </colgroup>
  <tr>
    <td width="100%" colspan="6">Your Order History</td>
  </tr>
<% If Not blnLoggedIn Then %>
  <tr>
    <td width="100%" colspan="6" align="center">You must login to view your order history.</td>
  </tr>
  <tr>
    <td width="100%" colspan="6" align="center"><% Call ShowLogin(mstrEmail, mlngOrderID, mstrMessage, 1) %></td>
  </tr>
<% Else %>
  <tr>
    <th>Order Number</th>
    <th>Order Date</th>
    <th>Items Ordered</th>
    <th>Order Total</th>
    <th>Status</th>
    <th>Tracking</th>
  </tr>
<%
'		pstrSQL = "SELECT sfOrders., sfOrders., sfOrders., sfOrders.orderIsComplete, Sum(sfOrderDetails.odrdtQuantity) AS " _

Do While Not objrsOrders.EOF

	If CStr(objrsOrders.Fields("orderID").Value) = CStr(lngOrderID) Then
		pstrOrderLink = objrsOrders.Fields("orderID").Value
	Else
		pstrOrderLink = "<a href='OrderHistory.asp?OrderID=" & objrsOrders.Fields("orderID").Value & "'>" & objrsOrders.Fields("orderID").Value & "</a>"
		pstrOrderTrackLink = "<a href='OrderHistory.asp?OrderID=" & objrsOrders.Fields("orderID").Value & "#tracking'>Track</a>"
	End If
%>
  <tr>
    <td><%= pstrOrderLink %></td>
    <td><%= FormatDateTime(objrsOrders.Fields("orderDate").Value,2) %></td>
    <td><%= objrsOrders.Fields("SumOfodrdtQuantity").Value %></td>
    <td><%= FormatCurrency(objrsOrders.Fields("orderGrandTotal").Value,2) %></td>
    <td><%= getOrderStatusMessage(objrsOrders.Fields("ssOrderStatus").Value) %></td>
    <td><%= pstrOrderTrackLink %></td>
  </tr>
<%
	objrsOrders.MoveNext
Loop

End If
%>
</table>

<% End Sub 'ShowOrderHistory %>

<%
Function splitTrackingLinks(byVal strTracking, byVal strShippedVia)

'Figure out the tracking
Dim pstrOut
Dim paryTrackingNumbers
Dim pstrTrackingItem
Dim paryTrackingItem
Dim i

	pstrOut = strTracking
	If instr(1, pstrOut, vbcrlf) > 0 Then
		paryTrackingNumbers = Split(pstrOut, vbcrlf)
		pstrOut = ""
		For i = 0 To UBound(paryTrackingNumbers)
			pstrTrackingItem = Trim(paryTrackingNumbers(i))
			If Len(pstrTrackingItem) > 0 Then
				paryTrackingItem = Split(pstrTrackingItem,";")
				If UBound(paryTrackingItem) > 0 Then
					If Len(pstrOut) = 0 Then
						pstrOut = TrackingLink(paryTrackingItem(0), paryTrackingItem(1), "frmTrackPackage" & i)
					Else
						pstrOut = pstrOut & "<br />" & TrackingLink(paryTrackingItem(0),paryTrackingItem(1), "frmTrackPackage" & i)
					End If
				Else
					If Len(pstrOut) = 0 Then
						pstrOut = TrackingLink(strShippedVia, paryTrackingItem(0), "frmTrackPackage" & i)
					Else
						pstrOut = pstrOut & "<br />" & TrackingLink(strShippedVia, paryTrackingItem(0), "frmTrackPackage" & i)
					End If
				End If
			End If	'Len(pstrTrackingItem) > 0
		Next 'i
	Else
		pstrOut = TrackingLink(strShippedVia, pstrOut, "frmTrackPackage")
	End If
	
	splitTrackingLinks = pstrOut
	
End Function	'splitTrackingLinks

'***********************************************************************************************
' Added for Gift Certificates
'***********************************************************************************************

	Function GC_LoadByOrder(lngOrderID, dblGrandTotal)

	Dim pstrSQL
	Dim p_objRS

	'On Error Resume Next
	
		If Len(lngOrderID) And isNumeric(lngOrderID) Then

			pstrSQL = "SELECT ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
					& " FROM ssGiftCertificateRedemptions" _
					& " WHERE ssGCRedemptionType=1 AND ssGiftCertificateRedemptions.ssGCRedemptionOrderID=" & lngOrderID

			Set p_objRS = CreateObject("adodb.Recordset")
			p_objRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			If Not (p_objRS.EOF Or p_objRS.BOF) Then
				If Trim(p_objRS.Fields("ssGCRedemptionType").Value & "") <> cStr(enGiftCertificate) Then
					mstrCertificate = Trim(p_objRS.Fields("ssGCRedemptionCGCode").Value & "")
					mdblssCertificateAmount = Trim(p_objRS.Fields("ssGCRedemptionAmount").Value & "")
					If isNumeric(mdblssCertificateAmount) Then
						mdblssCertificateAmount = CDbl(mdblssCertificateAmount)
						mdblssGCNewTotalDue = dblGrandTotal + mdblssCertificateAmount
					End If
					GC_LoadByOrder = True
				Else
					GC_LoadByOrder = False
				End If
			Else
				GC_LoadByOrder = False
			End If
			p_objRS.Close
			Set p_objRS = Nothing
		End If

	End Function    'GC_LoadByOrder

%>