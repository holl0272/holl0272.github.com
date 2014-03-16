<%Option Explicit
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   2.00.0001												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		November 15, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

'--------------------------------------------------------------------------------------------------

Sub closeObj(objItem)
	ReleaseObject objItem
End Sub

'--------------------------------------------------------------------------------------------------
%>
<!--#include virtual="ssl/sfLib/mail.asp"-->
<!--#include virtual="ssl/SFLib/adovbs.inc"-->
<!--#include virtual="ssl/SFLib/incCC.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_class.asp"-->
<%
'**************************************************
'
'	Start Code Execution
'

On Error Goto 0

mstrPageTitle = "Order Administration"

'page variables
Dim mAction
Dim mlngOrderID

Dim mblnShowFilter, mblnShowSummary
Dim mstrsqlWhere, mstrSortOrder,mstrOrderBy

'Display setting
Dim mbytDisplay

'Filter Elements
Dim mbytText_Filter
Dim mstrText_Filter

Dim mstrStartDate, mstrEndDate

Dim mbytShipment_Filter
Dim mbytPayment_Filter
Dim mbytDate_Filter
Dim mbytoptFlag_Filter

'Paging Elements
Dim mlngPageCount,mlngAbsolutePage
Dim mlngMaxRecords

    Set mclsOrder = New clsOrder
    With mclsOrder
    
	mAction = LoadRequestValue("Action")
	mlngOrderID = Request.QueryString("OrderID")

	Call .Load(mlngOrderID)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<LINK href="ssLibrary/ssStyleSheet.css" type="text/css" rel="stylesheet">
<title>Order <%= mlngOrderID %></title>

</head>

<BODY>
<CENTER>

<% Call ShowOrderDetail(.rsOrders) %>

</CENTER>
</BODY>
</HTML>
<%
    End With

    
    Set cnn = Nothing
    Response.Flush
%>

<% Sub ShowOrderDetail(objRS) 

Dim plngOrderID
Dim pcurRealSubTotal
Dim pstrCustName
Dim pstrCustAddr
Dim pstrProdID
Dim pblnOddRow
Dim pstrBackground
Dim pstrTempID
Dim plngodrdtID

	If Not isObject(objRS) Then Exit Sub
	plngOrderID = objRS.Fields("orderID").Value
	pblnOddRow = False

	pstrCustName = Replace(objRS.Fields("custFirstName").Value & " " & objRS.Fields("custMiddleInitial").Value & " " & objRS.Fields("custLastName").Value,"  "," ")
	pstrCustAddr = objRS.Fields("custCity").Value & ", " & objRS.Fields("custState").Value & " " & objRS.Fields("custZip").Value
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblOrderDetail">
  <colgroup align=left width=50%>
  <colgroup align=left width=50%>
  <tr>
    <th colspan="2" align="center"><h4>Order <%= mlngOrderID %></h4></th>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <table border="0" cellspacing="0" cellpadding="3" width="100%">
		<colgroup align=center>
		<colgroup align=left>
		<colgroup align=center>
		<colgroup align=center>
		<colgroup align=center>
        <tr>
          <th align=center>Product ID</th>
          <th align=center>Product Name</th>
          <th align=center>Quantity</th>
          <th align=center>Unit Price</th>
          <th align=center>Extended Price</th>
        </tr>
<%
Dim pblnWriteGiftWrapBackorder
Dim pblnStepBack

pcurRealSubTotal = 0
pblnWriteGiftWrapBackorder = False
pblnStepBack = False

Do While Not objRS.EOF

	If plngodrdtID <> Trim(objRS.Fields("odrdtID").Value) Then
'	If pstrProdID <> Trim(objRS.Fields("odrdtProductID").Value) Then
		plngodrdtID = Trim(objRS.Fields("odrdtID").Value)
		pstrProdID = Trim(objRS.Fields("odrdtProductID").Value)
		pblnOddRow = Not pblnOddRow
		If pblnOddRow Then
			pstrBackground = "lightgrey"	'set this color for the odd rows
		Else
			pstrBackground = ""				'set this color for the even rows
		End If

%>
        <tr bgcolor="<%= pstrBackground %>">
          <td><%= pstrProdID %></td>
          <td><%= objRS.Fields("odrdtProductName").Value %></td>
          <td><%= objRS.Fields("odrdtQuantity").Value %></td>
          <td><%= FormatCurrency(objRS.Fields("odrdtSubTotal").Value/objRS.Fields("odrdtQuantity").Value,2) %></td>
          <td><%= FormatCurrency(objRS.Fields("odrdtSubTotal").Value,2) %></td>
        </tr>
<%
		pcurRealSubTotal = pcurRealSubTotal + CDbl(objRS.Fields("odrdtSubTotal").Value)
	End If
	If Len(objRS.Fields("odrattrName").Value & "" & objRS.Fields("odrattrAttribute").Value) > 0 Then
%>
        <tr bgcolor="<%= pstrBackground %>">
          <td></td>
          <td><font size="-1">&nbsp;&nbsp;<%= objRS.Fields("odrattrName").Value & ": " & objRS.Fields("odrattrAttribute").Value %></font></td>
          <td></td>
          <td></td>
          <td></td>
        </tr>
<%	
	End If

	objRS.MoveNext
	
	If cblnSF5AE Then
		If objRS.EOF Then
			pblnWriteGiftWrapBackorder = True
		Else
			pblnWriteGiftWrapBackorder = CBool(plngodrdtID <> Trim(objRS.Fields("odrdtID").Value))
		End If
	End If
	
	If pblnWriteGiftWrapBackorder Then
		'Check for Gift Wrap
		If objRS.EOF Then
			pblnStepBack = True
			objRS.MovePrevious
		End If
		If objRS.Fields("odrdtGiftWrapQTY").Value > 0 Then %>
			<tr bgcolor="<%= pstrBackground %>">
			  <td></td>
			  <td>&nbsp;<i>Gift Wrap</i></td>
			  <td><%= objRS.Fields("odrdtGiftWrapQTY").Value %></td>
			  <td><%= FormatCurrency(objRS.Fields("odrdtGiftWrapPrice").Value/objRS.Fields("odrdtGiftWrapQTY").Value,2) %></td>
			  <td><%= FormatCurrency(objRS.Fields("odrdtGiftWrapPrice").Value,2) %></td>
			</tr>
<%			pcurRealSubTotal = pcurRealSubTotal + CDbl(objRS.Fields("odrdtGiftWrapPrice").Value)
		End If

		'Check for Back Orders
		If objRS.Fields("odrdtBackOrderQTY").Value > 0 Then %>
			<tr bgcolor="<%= pstrBackground %>">
			  <td></td>
			  <td colspan=4>&nbsp;<b><i>Quantity on Back Order: <%= objRS.Fields("odrdtBackOrderQTY").Value %></i></b></td>
			</tr>
<%		End If
		If pblnStepBack Then objRS.MoveNext
	End If
	pblnWriteGiftWrapBackorder = False
	
Loop
objRS.MoveFirst
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
	If isNumeric(objRS.Fields("orderCouponDiscount").Value) Then mcurCoupon = CDbl(objRS.Fields("orderCouponDiscount").Value)
End If

mcurDiscount = CDbl(objRS.Fields("orderGrandTotal").Value) - pcurRealSubTotal + mcurCoupon - CDbl(objRS.Fields("orderHandling").Value) - CDbl(objRS.Fields("orderCTax").Value) - CDbl(objRS.Fields("orderSTax").Value) - CDbl(objRS.Fields("orderShippingAmount").Value)
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
				If Len(objRS.Fields("orderCouponDiscount").Value & "") > 0 Then 
				  If CDbl(objRS.Fields("orderCouponDiscount").Value) > 0 Then 
			  %>
              <tr>
                <td>Coupon (<%= objRS.Fields("orderCouponCode").Value %>):</td>
                <td><%= FormatCurrency(-1 * objRS.Fields("orderCouponDiscount").Value,2) %></td>
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
			  <% If CDbl(objRS.Fields("orderSTax").Value & "") > 0 Then %>
              <tr>
                <td><%= objRS.Fields("cshpaddrShipState").Value %> Tax:</td>
                <td><%= FormatCurrency(objRS.Fields("orderSTax").Value) %></td>
              </tr>
			  <% End If %>
			  <% If CDbl(objRS.Fields("orderCTax").Value & "") > 0 Then %>
              <tr>
                <td><%= objRS.Fields("cshpaddrShipCountry").Value %> Tax:</td>
                <td><%= FormatCurrency(objRS.Fields("orderCTax").Value) %></td>
              </tr>
			  <% End If %>
              <tr>
                <td><%= objRS.Fields("orderShipMethod").Value %>:&nbsp;&nbsp;</td>
                <td><%= FormatCurrency(objRS.Fields("orderShippingAmount").Value) %></td>
              </tr>
			  <% If CDbl(objRS.Fields("orderHandling").Value & "") > 0 Then %>
              <tr>
                <td>Handling:</td>
                <td><%= FormatCurrency(objRS.Fields("orderHandling").Value) %></td>
              </tr>
			  <% End If %>
              <tr>
                <td colspan="2">
                  <hr>
                </td>
              </tr>
              <tr>
                <td>Total:</td>
                <td><%= FormatCurrency(objRS.Fields("orderGrandTotal").Value) %></td>
              </tr>
            <% 
				If GC_LoadByOrder(mlngOrderID, objRS.Fields("orderGrandTotal").Value) Then
            %>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
				<tr>
					<td align="left">Certificate (<a href="ssGiftCertificateAdmin.asp?Action=ViewByCode&ssGCCode><%= mstrCertificate %>"><%= mstrCertificate %></a>):</td>
					<td align="right"><%= FormatCurrency(mdblssCertificateAmount, 2) %></td>
				</tr>
				<tr>
					<td align="left">Amount Billed:</td>
					<td align="right"><%= FormatCurrency(mdblssGCNewTotalDue, 2) %></td>
				</tr>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
            <% 
				End If	'GC_LoadByOrder 
            %>
            </table>
          </td>
        </tr>
		<% If Len(objRS.Fields("orderComments").Value & "") > 0 Then %>
        <tr>
          <td colspan=5 align=left>&nbsp;&nbsp;Special Instructions:&nbsp;&nbsp;<%= objRS.Fields("orderComments").Value %></td>
        </tr>
        <% End If %>
      </table>
    </td>
  </tr>
  <tr class="tblhdr">
    <th colspan="2">Customer Information</th>
  </tr>
  <tr>
    <td valign="top">
      <table border="0" cellspacing="0" cellpadding="3" width="100%">
		<colgroup align=right>
		<colgroup align=left>
        <tr>
          <td>Sold to:&nbsp;&nbsp;</td>
          <td><%= pstrCustName %></td>
        </tr>
		<% If Len(objRS.Fields("custCompany").Value & "") > 0 Then %>
        <tr>
          <td>Company:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custCompany").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>Address:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custAddr1").Value %></td>
        </tr>
		<% If Len(objRS.Fields("custAddr2").Value & "") > 0 Then %>
        <tr>
          <td>&nbsp;</td>
          <td><%= objRS.Fields("custAddr2").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>&nbsp;</td>
          <td><%= pstrCustAddr %></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td><%= objRS.Fields("custCountry").Value %></td>
        </tr>
		<% If Len(objRS.Fields("custPhone").Value & "") > 0 Then %>
        <tr>
          <td>Phone:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custPhone").Value %></td>
        </tr>
        <% End If %>
		<% If Len(objRS.Fields("custFax").Value & "") > 0 Then %>
        <tr>
          <td>Fax:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custFax").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>email:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("custEmail").Value %></td>
        </tr>
      </table>
    </td>
<%
pstrCustName = Replace(objRS.Fields("cshpaddrShipFirstName").Value & " " & objRS.Fields("cshpaddrShipMiddleInitial").Value & " " & objRS.Fields("cshpaddrShipLastName").Value,"  "," ")
pstrCustAddr = objRS.Fields("cshpaddrShipCity").Value & ", " & objRS.Fields("cshpaddrShipState").Value & " " & objRS.Fields("cshpaddrShipZip").Value
%>
    <td valign="top">
      <table border="0" cellspacing="0" cellpadding="3" width="100%">
		<colgroup align=right>
		<colgroup align=left>
        <tr>
          <td>Shipped to:&nbsp;&nbsp;</td>
          <td><%= pstrCustName %></td>
        </tr>
		<% If Len(objRS.Fields("cshpaddrShipCompany").Value & "") > 0 Then %>
        <tr>
          <td>Company:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipCompany").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>Address:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipAddr1").Value %></td>
        </tr>
		<% If Len(objRS.Fields("cshpaddrShipAddr2").Value & "") > 0 Then %>
        <tr>
          <td>&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipAddr2").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>&nbsp;</td>
          <td><%= pstrCustAddr %></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipCountry").Value %></td>
        </tr>
		<% If Len(objRS.Fields("cshpaddrShipPhone").Value & "") > 0 Then %>
        <tr>
          <td>Phone:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipPhone").Value %></td>
        </tr>
        <% End If %>
		<% If Len(objRS.Fields("cshpaddrShipFax").Value & "") > 0 Then %>
        <tr>
          <td>Fax:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipFax").Value %></td>
        </tr>
        <% End If %>
        <tr>
          <td>email:&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("cshpaddrShipEmail").Value %></td>
        </tr>
      </table>
    </td>
  </tr>

   <tr class="tblhdr">
    <th width="100%" colspan="2">Payment Method</th>
   </tr>
   <tr>
     <td valign=top style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
		<%
		Dim pblnPhoneFax
		Dim pblnPayPal
		Dim mstrOrderPaymentMethod
		Dim mstrOrderPaymentMethodText
		Dim mstrOrderPaymentMethodAutoFill

		mstrOrderPaymentMethod = Trim(objRS.Fields("orderPaymentMethod").Value & "")
		
		'Now PayPal may or may not just be PayPal
		pblnPayPal = CBool(Instr(1, mstrOrderPaymentMethod, "PayPal") > 0)
		
		'Now PhoneFax will/may have a payment type extension. Ex. PhoneFax_Credit Card, PhoneFax_eCheck, PhoneFax_PO
		pblnPhoneFax = CBool(Instr(1, mstrOrderPaymentMethod, "PhoneFax") > 0)
		mstrOrderPaymentMethod = Replace(mstrOrderPaymentMethod, "PhoneFax_", "")	'Now remove the prefix

		Select Case mstrOrderPaymentMethod
			Case "Credit Card"
								mstrOrderPaymentMethodText = mstrOrderPaymentMethod
								mstrOrderPaymentMethodAutoFill = mstrOrderPaymentMethod
			Case "eCheck"
								mstrOrderPaymentMethodText = mstrOrderPaymentMethod & objRS.Fields("orderCheckNumber").Value
								mstrOrderPaymentMethodAutoFill = "eCheck " & objRS.Fields("orderCheckNumber").Value
			Case "PO"
								mstrOrderPaymentMethodText = "P.O. " & objRS.Fields("orderPurchaseOrderNumber").Value & "<br />" _
														   & "Name: " & objRS.Fields("orderPurchaseOrderName").Value
								mstrOrderPaymentMethodAutoFill = "P.O. " & objRS.Fields("orderPurchaseOrderNumber").Value
			Case "PayPal"
								mstrOrderPaymentMethodText = mstrOrderPaymentMethod
								mstrOrderPaymentMethodAutoFill = mstrOrderPaymentMethod
			Case Else
				'Catch the unique ones here
								mstrOrderPaymentMethodText = mstrOrderPaymentMethod				
								mstrOrderPaymentMethodAutoFill = mstrOrderPaymentMethod
		End Select
		If pblnPhoneFax Then mstrOrderPaymentMethodText = "Phone/Fax - " & mstrOrderPaymentMethodText
		
		%>
		<table ID="Table11">
		  <tr><th align="left"><%= mstrOrderPaymentMethodText %></th></tr>
		  <%
			Select Case mstrOrderPaymentMethod
				Case "Credit Card":
					Response.Write "<tr><td>"
					Call ShowCC(plngOrderID)
					Response.Write "</td></tr>"
				Case "eCheck":
				Case "PO"
				Case "PayPal"
				Case Else
					'Catch the unique ones here
			End Select
		  %>
	    </table>
     </td>
     <td valign=top style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
		<%
			Select Case mstrOrderPaymentMethod
				Case "Credit Card":		Call ShowTransactionResponse(plngOrderID)
				Case "eCheck":			Response.Write "<p align=center><a href='../printEcheck.asp?orderid=" & plngOrderID & "'>Print Check</a></p>"
				Case "PO"
				Case "PayPal"
				Case Else
					'Catch the unique ones here
			End Select
		%>
	  </td>
	</tr>
<script language="javascript">
function SetPaymentMethod()
{
document.frmData.ssPaidVia.value = "<%= mstrOrderPaymentMethodAutoFill %>";
}
</script>
 <tr class="tblhdr">
    <th width="100%" colspan="2">Order Status</th>
  </tr>
  <tr>
    <td>

      <table border="0" cellspacing="0">
		<colgroup align=right>
		<colgroup align=left>
        <tr>
          <td>Order Placed On:&nbsp;&nbsp;</td>
          <td><%= FormatDateTime(objRS.Fields("orderDate").Value,1) & " at " & FormatDateTime(objRS.Fields("orderDate").Value,3) %></td>
        </tr>
<% 
'Figure out the payment
Const cstrAwaitPayMsg = "Awaiting Payment"
Dim pstrPaymentType, pstrPaymentMsg
pstrPaymentType = Trim(objRS.Fields("orderPaymentMethod").Value)
If pstrPaymentType = "Credit Card" Then
'	Dim objRSPayment
'	Set objRSPayment = Server.CreateObject("ADODB.RECORDSET")
'	With
'		.CursorLocation = 2 'adUseClient
'		.CursorType = 3 'adOpenStatic
'		.LockType = 1 'adLockReadOnly
'		.Open "Select ",cnn
'		.Close
'	End With
'	Set objRSPayment = Nothing

End If

%>
        <tr>
          <td><LABEL for="ssDatePaymentReceived">Date Payment Received:</LABEL>&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("ssDatePaymentReceived").Value %></td>
        </tr>
        <tr>
          <td><LABEL for="ssPaidVia" ondblclick="SetPaymentMethod();" title="Double click to set payment method">Paid By:</LABEL>&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("ssPaidVia").Value %></td>
        </tr>

        <tr>
          <td><LABEL for="ssDateOrderShipped">Date Order Shipped:</LABEL>&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("ssDateOrderShipped").Value %></td>
        </tr>
		
        <tr>
          <td><LABEL for="ssShippedVia">Order Shipped via:&nbsp;&nbsp;</LABEL></td>
          <td>
              <%
              Dim i
              
              For i=0 to UBound(maryShipMethods)
				If CStr(objRS.Fields("ssShippedVia").Value & "") = CStr(i) Then
					Response.Write maryShipMethods(i)(0)
				End If
              Next
			  %>
            </select>
          </td>
        </tr>
        
        <tr>
          <td><LABEL for="ssTrackingNumber">Tracking Number:</LABEL>&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("ssTrackingNumber").Value %></td>
        </tr>
        <tr>
          <td><LABEL for="ssInternalNotes">Internal Notes:</LABEL>&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("ssInternalNotes").Value %></td>
        </tr>
        <tr>
          <td><LABEL for="ssExternalNotes">External Notes:</LABEL>&nbsp;&nbsp;</td>
          <td><%= objRS.Fields("ssExternalNotes").Value %></td>
        </tr>

<% If Len(objRS.Fields("ssDateEmailSent").Value & "") = 0 Then %>
        <tr>
          <td>Order Fulfilment Email Sent on:&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
<% Else %>
        <tr>
          <td>Order Fulfilment Email Sent on:&nbsp;&nbsp;</td>
          <td><%= FormatDateTime(objRS.Fields("ssDateEmailSent").Value,1) %></td>
        </tr>
<% End If %>
      </table>
    </td>
    <td>
		<table>
			<tr>
				<td align="right">&nbsp;</td>
				<td align="left"><INPUT type="checkbox" id=ssOrderFlagged name=ssOrderFlagged value='1' <% If objRS.Fields("ssOrderFlagged").Value Then Response.Write "checked"%>><LABEL for"ssOrderFlagged">Flag this order</LABEL>&nbsp;</td>
			</tr>
			<tr>
			  <td align="right"><LABEL for="ssBackOrderDateNotified">Date Back Order Notification Sent:</LABEL>&nbsp;&nbsp;</td>
			  <td align="left"><%= objRS.Fields("ssBackOrderDateNotified").Value %></td>
			</tr>
			<tr>
			  <td align="right"><LABEL for="ssBackOrderDateExpected">Date Expected In:</LABEL>&nbsp;&nbsp;</td>
			  <td align="left"><%= objRS.Fields("ssBackOrderDateExpected").Value %></td>
			</tr>
			<tr>
			  <td align="right"><LABEL for="ssBackOrderMessage">Back Order Message:</LABEL>&nbsp;&nbsp;</td>
			  <td align="left"><%= objRS.Fields("ssBackOrderMessage").Value %></td>
			</tr>
			<tr>
			  <td align="right"><LABEL for="ssBackOrderInternalMessage">Internal Back Order Message:</LABEL>&nbsp;&nbsp;</td>
			  <td align="left"><%= objRS.Fields("ssBackOrderInternalMessage").Value %></td>
			</tr>
	  </TABLE>
    </td>
  </tr>
</table>
<% 

End Sub 'ShowOrderDetail

'**************************************************************************************************************************************************

Sub ShowTransactionResponse(lngOrderID)

Dim pstrSQL
Dim mobjRSCC
Dim i

	pstrSQL = "SELECT trnsrspID, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode, trnsrspRetrievalCode, trnsrspAuthNo, trnsrspErrorMsg, trnsrspErrorLocation, trnsrspSuccess FROM sfTransactionResponse WHERE trnsrspOrderId=" & wrapSQLValue(lngOrderID, False, enDatatype_number)
	Set mobjRSCC = GetRS(pstrSQL)
	If Not mobjRSCC.EOF Then
		For i = 1 To mobjRSCC.RecordCount
%>
		<table>
		  <tr>
			<th colspan=2>Transaction Response</th>
		  </tr>
		  <tr>
			<td align="right">Authorization&nbsp;#:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspAuthNo").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Success:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspSuccess").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Customer Tx #:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspCustTransNo").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Merchant Tx #:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspMerchTransNo").Value %></td>
		  </tr>
		  <tr>
			<td align="right">AVS Code:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspAVSCode").Value %></td>
		  </tr>
		  <tr>
			<td align="right">AUX Message:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspAUXMsg").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Action Code:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspActionCode").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Retrieval Code:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspRetrievalCode").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Error Message:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspErrorMsg").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Error Location:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("trnsrspErrorLocation").Value %></td>
		  </tr>
	    </table>
<%
		Next 'i
	End If	'mobjRSCC.EOF
	
	Call ReleaseObject(mobjRSCC)

End Sub	'ShowTransactionResponse

'**************************************************************************************************************************************************

Sub ShowCC(lngOrderID)

Dim pstrSQL
Dim mobjRSCC
Dim mstrCCNumber

	If cblnSQLDatabase Then
		pstrSQL = "SELECT sfCPayments.payCardType, sfCPayments.payCardName, sfCPayments.payCardNumber, sfCPayments.payCardExpires, sfTransactionTypes.transName " _
				& " FROM (sfOrders INNER JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) INNER JOIN sfTransactionTypes ON convert(Integer,sfCPayments.payCardType) = sfTransactionTypes.transID" _
				& " Where OrderID=" & lngOrderID
	Else
		pstrSQL = "SELECT sfCPayments.payCardType, sfCPayments.payCardName, sfCPayments.payCardNumber, sfCPayments.payCardExpires, sfTransactionTypes.transName " _
				& " FROM (sfOrders INNER JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) INNER JOIN sfTransactionTypes ON CLng(sfCPayments.payCardType) = sfTransactionTypes.transID" _
				& " Where OrderID=" & lngOrderID
	End If
	If Len(cstrCCV) > 0 Then pstrSQL = Replace(pstrSQL,"sfTransactionTypes.transName", "sfTransactionTypes.transName, " & cstrCCV)
	Set mobjRSCC = GetRS(pstrSQL)

	If mobjRSCC.EOF Then
%>
		<table ID="Table10">
		  <tr>
			<th>No transaction record for <%= lngOrderID %></th>
		  </tr>
	    </table>
<%
	Else
		mstrCCNumber = DecryptCardNumber(mobjRSCC.Fields("payCardNumber").Value, False)
%>
		<table ID="tblCCInfo">
		  <tr>
			<th colspan=2>Credit Card Information</th>
		  </tr>
		  <tr>
			<td align="right">Name on Credit Card:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("payCardName").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Credit Card Type:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields("transName").Value %></td>
		  </tr>
		  <tr>
			<td align="right">Credit Card Number:</td>
			<td align="left">&nbsp;<%= mstrCCNumber %></td>
		  </tr>
		  <% If Len(cstrCCV) > 0 Then %>
		  <tr>
			<td align="right">CCV:</td>
			<td align="left">&nbsp;<%= mobjRSCC.Fields(cstrCCV).Value %></td>
		  </tr>
		  <% End If 'Len(cstrCCV) > 0 %>
		  <tr>
		    <td align="right">Expiration Date:</td>
		    <td align="left">&nbsp;<%= mobjRSCC.Fields("payCardExpires").Value %></td>
		  </tr>
	    </table>
<%
	End If	'mobjRSCC.EOF
	mobjRSCC.Close 
	Set mobjRSCC = nothing

End Sub 'ShowCC
%>