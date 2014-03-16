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

Function ConvertToBoolean(vntValue)

On Error Resume Next

	vntValue = cBool(vntValue)
	If Err.number <> 0 Then vntValue = False
	ConvertToBoolean = vntValue

End Function	'ConvertToBoolean

Sub closeObj(objItem)
	ReleaseObject objItem
End Sub


'--------------------------------------------------------------------------------------------------
%>
<!--#include file="../SFLib/mail.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/incCC.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
<!--#include file="ssOrderAdmin_class.asp"-->
<%
'**************************************************
'
'	Start Code Execution
'

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

<title>Order <%= mlngOrderID %></title>
<style>

body         { font-size: 12pt; font-family: Times New Roman }
.PaymentText { font-size: 12pt }
.PaymentHeader { font-size: 12pt; font-weight: bold }
.ShippingText { font-size: 12pt }
.ShippingHeader { font-size: 12pt; font-weight: bold }
.orderItemText { font-size: 10pt; }
.orderItemAttrText { font-size: 9pt; }
.CostText { font-size: 12pt }
.CostHeader { font-size: 12pt; font-weight: bold }
.AmountDueText { font-size: 12pt }
.AmountDueHeader { font-size: 12pt; font-weight: bold }
.style1 {
	color: #0000FF;
}
.style2 {
	font-family: "Times New Roman";
	font-size: small;
}
</style>
</head>

<BODY >
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
<table class="tbl" cellpadding="3" cellspacing="0" border="0" id="tblOrderDetail">
  <colgroup align=left width=50%>
  <colgroup align=left width=50%>
   <tr>
    <th width="100%" colspan="2" align="center">
<span lang="en-us"><span class="style1">GameWearNow</span></span><br />
<font color="blue" style="font-size: 9pt; font-family: Book Antiqua">http://<span lang="en-us"><span class="style2">www.gamewearnow.com</span></span></font> 
    </th>
   </tr>
  <tr>
    <td><strong><font size="3">Order ID:&nbsp;&nbsp;<%= mlngOrderID %></font></strong></td>
    <td align="right"><strong>Order Date:&nbsp;&nbsp;<%= FormatDateTime(objRS.Fields("orderDate").Value,2) %></strong></td>
  </tr>
   <tr>
    <th width="100%" colspan="2">&nbsp;</th>
   </tr>
  <tr>
    <td valign="top">
      <table border="0" cellspacing="0" cellpadding="3" width="100%">
		<colgroup align=right>
		<colgroup align=left>
        <tr>
          <td align="right"><span class="ShippingHeader">Sold To:</span>&nbsp;&nbsp;</td>
          <td align="left">&nbsp;</td>
        </tr>
        <tr>
          <td>Name:</strong>&nbsp;&nbsp;</td>
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
          <td><span class="ShippingHeader">Ship To:</strong>&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>Name:&nbsp;&nbsp;</td>
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
		<% If Len(objRS.Fields("orderComments").Value & "") > 0 Then %>
        <tr>
          <td colspan=5 align=left>&nbsp;&nbsp;Special Instructions:&nbsp;&nbsp;<%= objRS.Fields("orderComments").Value %></td>
        </tr>
        <% End If %>

   <TR>
    <TH width="100%" colspan="2">&nbsp;</TH>
   </TR>
   <TR>
     <TD colspan="2">
<%
Dim mstrOrderPaymentMethod
Dim mobjRSCC
Dim mstrCCNumber
Dim pstrSQL

	mstrOrderPaymentMethod = Trim(objRS.Fields("orderPaymentMethod").Value & "")

	If (mstrOrderPaymentMethod = "Credit Card") OR (mstrOrderPaymentMethod = "PhoneFax") Then
		If cblnSQLDatabase Then
			pstrSQL = "SELECT sfCPayments.payCardType, sfCPayments.payCardName, sfCPayments.payCardNumber, sfCPayments.payCardExpires, sfTransactionTypes.transName " _
					& " FROM (sfOrders INNER JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) INNER JOIN sfTransactionTypes ON convert(Integer,sfCPayments.payCardType) = sfTransactionTypes.transID" _
					& " Where OrderID=" & plngOrderID
		Else
			pstrSQL = "SELECT sfCPayments.payCardType, sfCPayments.payCardName, sfCPayments.payCardNumber, sfCPayments.payCardExpires, sfTransactionTypes.transName " _
					& " FROM (sfOrders INNER JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) INNER JOIN sfTransactionTypes ON CLng(sfCPayments.payCardType) = sfTransactionTypes.transID" _
					& " Where OrderID=" & plngOrderID
		End If
		Set mobjRSCC = GetRS(pstrSQL)

		'Decrypt card - protect against error
		On Error Resume Next
		Dim pobjCCEncrypt		
		Set pobjCCEncrypt = Server.CreateObject("SFServer.CCEncrypt")
		If Err.number = 0 Then
			pobjCCEncrypt.putSeed(iCC)
			mstrCCNumber = pobjCCEncrypt.decrypt(mobjRSCC.Fields("payCardNumber").Value)
		Else
			mstrCCNumber = mobjRSCC.Fields("payCardNumber").Value
		End If
		Set pobjCCEncrypt = Nothing
		On Error Goto 0
		
		'Now Mask the card Number
		mstrCCNumber = Replace(mstrCCNumber," ","")
		mstrCCNumber = String(Len(mstrCCNumber) - 4,"*") & "-" & Right(mstrCCNumber,4)

%>
		<table>
		  <tr>
			<td align="left"><span class="PaymentHeader">Payment Method:</span>&nbsp;<span class="PaymentText">Credit Card</span>&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td align="left"><span class="PaymentHeader">Name on Credit Card:</span>&nbsp;<span class="PaymentText"><%= mobjRSCC.Fields("payCardName").Value %></span></td>
		  </tr>
		  <TR>
			<TD align="left"><span class="PaymentHeader">Credit Card Type:</span>&nbsp;<span class="PaymentText"><%= mobjRSCC.Fields("transName").Value %></span></TD>
			<TD align="left"><span class="PaymentHeader">Credit Card Number:</span>&nbsp;<span class="PaymentText"><%= mstrCCNumber %>&nbsp;&nbsp;&nbsp;<span class="PaymentHeader">Expiration Date:</span>&nbsp;<%= mobjRSCC.Fields("payCardExpires").Value %></span></TD>
		  </TR>
	  </TABLE>
	  </TD>
	</tr>
<%
		mobjRSCC.Close 
		Set mobjRSCC = nothing
	End If
%> 
   <tr>
    <th width="100%" colspan="2">&nbsp;</th>
   </tr>
  <tr>
    <td width="100%" colspan="2">
      <table border="0" cellspacing="0" cellpadding="4" width="100%">
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
Dim p_strProdIDLink
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
			pstrBackground = ""	'set this color for the odd rows
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
          <td colspan="6">&nbsp;&nbsp;<span class="orderItemText"><%= objRS.Fields("odrattrName").Value & ": " & objRS.Fields("odrattrAttribute").Value %></span></td>
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
			  <td>&nbsp;</td>
			  <td>&nbsp;<i>Gift Wrap</i></td>
			  <td><span class="orderItemAttrText"><%= objRS.Fields("odrdtGiftWrapQTY").Value %></span></td>
			  <td><span class="orderItemAttrText"><%= FormatCurrency(objRS.Fields("odrdtGiftWrapPrice").Value/objRS.Fields("odrdtGiftWrapQTY").Value,2) %></span></td>
			  <td><span class="orderItemAttrText"><%= FormatCurrency(objRS.Fields("odrdtGiftWrapPrice").Value,2) %></span></td>
			</tr>
<%			pcurRealSubTotal = pcurRealSubTotal + CDbl(objRS.Fields("odrdtGiftWrapPrice").Value)
		End If

		'Check for Back Orders
		If objRS.Fields("odrdtBackOrderQTY").Value > 0 Then %>
			<tr bgcolor="<%= pstrBackground %>">
			  <td colspan=1></td>
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
          <td width="100%" colspan="7">&nbsp;</td>
        </tr>
        <tr>
          <td colspan=3 align=left valign=bottom>&nbsp;&nbsp;</td>
          <td width="40%" colspan="4" align=right>
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
                <td><span class="CostHeader">Subtotal:</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                <td><span class="CostText"><%= FormatCurrency(pcurRealSubTotal,2) %></span>&nbsp;&nbsp;</td>
              </tr>
			  <% 
			  If cblnSF5AE Then
				If Len(objRS.Fields("orderCouponDiscount").Value & "") > 0 Then 
				  If CDbl(objRS.Fields("orderCouponDiscount").Value) > 0 Then 
			  %>
              <tr>
                <td><span class="CostHeader">Coupon (<%= objRS.Fields("orderCouponCode").Value %>):</span></td>
                <td><span class="CostText"><%= FormatCurrency(-1 * objRS.Fields("orderCouponDiscount").Value,2) %></span>&nbsp;&nbsp;</td>
              </tr>
			  <% 
				  End If 
				End If 
			  End If 
			  %>
			  <% If CDbl(mcurDiscount) < -0.009 Then %>
              <tr>
                <td><span class="CostHeader">Discount:</span></td>
                <td><span class="CostText"><%= FormatCurrency(mcurDiscount,2) %></span>&nbsp;&nbsp;</td>
              </tr>
			  <% End If %>
			  <% If CDbl(objRS.Fields("orderSTax").Value & "") > 0 Then %>
              <tr>
                <td><span class="CostHeader"><%= objRS.Fields("cshpaddrShipState").Value %> Tax:</span></td>
                <td><span class="CostText"><%= FormatCurrency(objRS.Fields("orderSTax").Value) %></span>&nbsp;&nbsp;</td>
              </tr>
			  <% End If %>
			  <% If CDbl(objRS.Fields("orderCTax").Value & "") > 0 Then %>
              <tr>
                <td><span class="CostHeader"><%= objRS.Fields("cshpaddrShipCountry").Value %> Tax:</span></td>
                <td><span class="CostText"><%= FormatCurrency(objRS.Fields("orderCTax").Value) %></span>&nbsp;&nbsp;</td>
              </tr>
			  <% End If %>
              <tr>
                <td><span class="CostHeader"><%= objRS.Fields("orderShipMethod").Value %>:</span>&nbsp;&nbsp;</td>
                <td><span class="CostText"><%= FormatCurrency(objRS.Fields("orderShippingAmount").Value) %></span>&nbsp;&nbsp;</td>
              </tr>
			  <% If CDbl(objRS.Fields("orderHandling").Value & "") > 0 Then %>
              <tr>
                <td><span class="CostHeader">Handling:</span></td>
                <td><span class="CostText"><%= FormatCurrency(objRS.Fields("orderHandling").Value) %></span>&nbsp;&nbsp;</td>
              </tr>
			  <% End If %>
              <tr>
                <td colspan="2">
                  <hr>
                </td>
              </tr>
              <tr>
                <td><span class="AmountDueHeader">Total:</span></td>
                <td><span class="AmountDueText"><%= FormatCurrency(objRS.Fields("orderGrandTotal").Value) %></span>&nbsp;&nbsp;</td>
              </tr>
            <% 
				If GC_LoadByOrder(mlngOrderID, objRS.Fields("orderGrandTotal").Value) Then
            %>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
				<tr>
					<td align="left"><span class="AmountDueHeader">Certificate (<a href="ssGiftCertificateAdmin.asp?Action=ViewByCode&ssGCCode=<%= mstrCertificate %>"><%= mstrCertificate %></a>):</span></td>
					<td align="right"><span class="AmountDueText"><%= FormatCurrency(mdblssCertificateAmount, 2) %></span>&nbsp;&nbsp;</td>
				</tr>
				<tr>
					<td align="left"><span class="AmountDueHeader">Amount Billed:</span></td>
					<td align="right"><span class="AmountDueText"><%= FormatCurrency(mdblssGCNewTotalDue, 2) %></span>&nbsp;&nbsp;</td>
				</tr>
				<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
            <% 
				End If	'GC_LoadByOrder 
            %>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
   <tr>
    <th width="100%" colspan="2" align="center">
<font color="red" style="font-size: 14pt; font-family: Book Antiqua">Thanks!</font><br />
<font color="red" style="font-size: 14pt; font-family: Book Antiqua">We appreciate your business!</font>
    </th>
   </tr>

</table>



<% End Sub 'ShowOrderDetail%>