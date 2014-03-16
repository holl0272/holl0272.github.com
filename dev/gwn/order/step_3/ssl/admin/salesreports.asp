<%@ LANGUAGE="VBSCRIPT" %><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront v.2.5.1
'
'   Author     :   LaGarde, Incorporated

'   Description:   This file produces the various sales reports that are built-in
'                  to the StoreFront product.
'
'   Notes      :  There are no configurable elements in this file.
'                  
'
'                         COPYRIGHT NOTICE
'
'   The contents of this file is protected under the United States
'   copyright laws as an unpublished work, and is confidential and
'   proprietary to LaGarde, Incorporated.  Its use or disclosure in 
'   whole or in part without the expressed written permission of 
'   LaGarde, Incorporated is expressely prohibited.
'
'   (c) Copyright 1998 by LaGarde, Incorporated.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

%><%


	DSN_Name = Session("DSN_Name")
	StartDate = Request("StartDate")&" 12:00:01 AM"
	EndDate = Request("EndDate")&" 11:59:59 PM"
	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""
	'Connection.Open "driver={Microsoft Access Driver (*.mdb)};dbq=e:\inetpub\vsroot\lagarde\thenaturestore\_private\naturestore.mdb"

	'Report 1 is the Order Summary Report for the requested time period.

	If Request("Report") = "1" Then
	
	SQLStmt = "SELECT CUSTOMER_ID, NAME, ORDER_DATE, GRAND_TOTAL FROM Customer "
	SQLStmt = SQLStmt & "WHERE ((ORDER_DATE >= #" &StartDate & "#) "
	SQLStmt = SQLStmt & "AND (ORDER_DATE <= #" & EndDate & "#)) "
	SQLStmt = SQLStmt & "AND GRAND_TOTAL <> '$0.00' "
	
	Set RSSumDate = Connection.Execute(SQLStmt)
	
	'Report 2 is the Order Detail for a transaction reported 
	'in the Order Summary report.
	
	ElseIf Request("Report") = "2" Then
	
	SQLStmt = "SELECT * FROM Customer WHERE CUSTOMER_ID = "
	SQLStmt = SQLStmt & "" & Request("OrderID") & " "
	'Response.Write (SQLStmt)
	Set RSCustDetail = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT * FROM orders, product WHERE ORDER_ID = "
	SQLStmt = SQLStmt & "" & Request("OrderID") & " AND "
	SQLStmt = SQLStmt & "orders.Product_ID = product.Product_ID "
	'Response.Write (SQLStmt)
	Set RSOrderDetail = Connection.Execute(SQLStmt)
	
	'Report 3 is the Customer Invoice Report.

	ElseIf Request("Report") = "3" Then

	SQLStmt = "SELECT * FROM Customer WHERE CUSTOMER_ID = "
	SQLStmt = SQLStmt & "" & Request("OrderID") & " "

	Set RSCustDetail = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT * FROM orders, product WHERE ORDER_ID = "
	SQLStmt = SQLStmt & "" & Request("OrderID") & " AND "
	SQLStmt = SQLStmt & "orders.Product_ID = product.Product_ID "

	Set RSOrderDetail = Connection.Execute(SQLStmt)

	'Report 4 is the Sales Transaction Report for summarizing Tax, 
	'Shipping, Net and Gross Sales for the selected period.
	
	ElseIf Request("Report") = "4" Then

	
	SQLStmt = "SELECT Sum(CCur(SUB_TOTAL)) AS SumSubTotal, Sum(CCur(TAX)) "
	SQLStmt = SQLStmt & "AS SumTax, Sum(CCur(SHIPPING_UPS)) AS SumShipUPS, "
	SQLStmt = SQLStmt & "Sum(CCur(SHIPPING_AIR)) AS SumShipAIR, "
	SQLStmt = SQLStmt & "Sum(CCur(GRAND_TOTAL)) AS SumGrandTotal "
	SQLStmt = SQLStmt & "FROM Customer WHERE ((ORDER_DATE >= "
	SQLStmt = SQLStmt & "#" & StartDate & "#) AND "
	SQLStmt = SQLStmt & "(ORDER_DATE <= #" & EndDate & "#)) "
	'Response.Write (SQLStmt)
	Set RSSummary = Connection.Execute(SQLStmt)


	
	If IsNull (RSSummary("SumGrandTotal")) Then
	SumGrandTotal = 0
	SumSubTotal = 0
	SumShipAIR = 0
	SumShipUPS = 0
	SumTax = 0
	Else 
	SumGrandTotal = RSSummary("SumGrandTotal")
	SumSubTotal = RSSummary("SumSubTotal")
	SumShipAIR = RSSummary("SumShipAIR")
	SumShipUPS = RSSummary("SumShipUPS")
	SumTax = RSSummary("SumTax")
	

	End If
	ElseIf Request("Report") = "5" Then

	SQLStmt = "DELETE * FROM customer WHERE CUSTOMER_ID = " & Request("OrderID") & " "

	Set RSDelete = Connection.Execute(SQLStmt)

	ElseIf Request("Report") = "6" Then

	SQLStmt ="SELECT * FROM customer WHERE CUSTOMER_ID = " & Request("OrderID") & " "
	RSChck = Connection.Execute(SQLStmt)
	
	ElseIf Request("Report") = "7" Then

	SQLStmt = "Select DISTINCT REFERER, HTTP_REFERER FROM customer "
	SQLStmt = SQLStmt & "WHERE ((ORDER_DATE >= #" & StartDate & "#) AND "
	SQLStmt = SQLStmt & "(ORDER_DATE <= #" & EndDate & "#)) "
	
	Set RSRef = Connection.Execute(SQLStmt)
%>

<table border="0" cellpadding="2" cellspacing="2" width="95%" align="center">
	<% If RSRef.EOF Then %>
  <tr>
    <td colspan="4" align="center" bgcolor="#A09A8B"> <big><strong>There were no Referer sales
    reported for the period <%= Request("StartDate") %> to <%= Request("EndDate") %></strong></big> </td>
  </tr>
  <% Else %>


  <tr>
    <td colspan="3" align="center" bgcolor="#A09A8B"> <big><strong>Referer Sales For the Period <%= Request("StartDate") %> to <%= Request("EndDate") %></strong></big>
    </td>
  </tr>

<tr>
    <td align="center" bgcolor="#A09A8B"><strong>Referring Partner</strong></td>
    <td align="center" bgcolor="#A09A8B"><strong>Referring Domain</strong></td>
    <td align="center" bgcolor="#A09A8B"><strong>Total Sales</strong></td>
  </tr>

<%

	RSRef.MoveFirst
	Do While NOT RSRef.EOF
%>
<%

	SQLSTmt = "SELECT Sum(CCur(SUB_TOTAL)) as refSubTotal from customer "
	SQLStmt = SQLStmt & "WHERE (customer.REFERER = '" & RSRef("REFERER") & "') "
	Set RSRefTotal = Connection.Execute(SQLStmt)
%>
<% 'If RSRefTotal("RefSubTotal") > 0 Then %>
<% If Trim(RSRef("REFERER")) = "REFERER_ID" OR Trim(RSRef("REFERER")) = "" Then %>
<% ' Do Nothing %>
<% Else %>

<tr>
	<td align="right"><strong><%= RSRef("REFERER") %></strong></td>
	<td align="right"><strong><%= RSRef("HTTP_REFERER") %></strong></td>

	<td align="right"><strong><%= FormatCurrency(RSRefTotal("refSubTotal"),2) %></strong><td>
</tr>
<% 'Else %>
<% End If %>
<%
	RSRef.MoveNext
	Loop

%>
<% End If %>
</table>

<%
	Else
	End If


%><html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Sales Reporting</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body style="font-family: Verdana">

<p align="center"> 

<% If Request("Report") = "1" Then %> </p>

<table border="0" cellpadding="2" cellspacing="2" width="95%" align="center">
  <% If RSSumDate.EOF Then %>
  <tr>
    <td colspan="4" align="center" bgcolor="#A09A8B"> <big><strong>There were no sales
    reported for the period <%= Request("StartDate") %> to <%= Request("EndDate") %></strong></big> </td>
  </tr>
  <% Else %>
  <tr>
    <td colspan="4" align="center" bgcolor="#A09A8B"> <big><strong>Sales For the Period <%= Request("StartDate") %> to <%= Request("EndDate") %></strong></big>
    </td>
  </tr>
  <tr>
    <td width="25%" align="center" bgcolor="#A09A8B">Date of Order</td>
    <td width="25%" align="center" bgcolor="#A09A8B">Order ID</td>
    <td width="25%" align="center" bgcolor="#A09A8B">Customer Name</td>
    <td width="25%" align="center" bgcolor="#A09A8B">Order Total &nbsp;</td>
  </tr>

<%

	Do While NOT RSSumDate.EOF
%>

  <tr>
    <td align="left"><%= RSSumDate("ORDER_DATE") %></td>
    <td align="left"><a href="salesreports.asp?OrderID=<%= RSSumDate("Customer_ID") %>&amp;Report=2"><%= RSSumDate("CUSTOMER_ID") %></a>
    </td>
    <td align="left"><%= RSSumDate("NAME") %></td>
    <td align="right"><%= RSSumDate("GRAND_TOTAL") %> &nbsp;</td>
  </tr>

<%
	RSSumDate.MoveNext
	Loop
%>
</table>
</center>

<% End If %>

<% ElseIf Request("Report") = "2" Then %>

  <table border="0" cellpadding="2" cellspacing="2" width="95%" align="center">
    <tr>
      <td colspan="4" bgcolor="#A09A8B" align="center">
      <strong>Transaction Detail Report for Order Number: <%= Request("OrderID") %>
      </strong></td>
    </tr>
	<td colspan="4">&nbsp;</td>
    <tr>
      <td colspan="4" bgcolor="#A09A8B"><big><u>Sold To</u></big></td>
    </tr>
    <tr>
      <td width="10%">Name:</td>
      <td width="40%"><%= RSCustDetail("NAME") %></td>
      <td width="10%">Order Date:</td>
      <td width="40"><%= FormatDateTime(RSCustDetail("ORDER_DATE"),vbShortDate) %></td>
    </tr>
    <tr>
      <td>Company:</td>
      <td><%= RSCustDetail("COMPANY") %></td>
      <td>Order Time:</td>
      <td><%= FormatDateTime(RSCustDetail("ORDER_DATE"),vbLongTime) %></td>
    </tr>
    <tr>
      <td>Address:</td>
      <td><%= RSCustDetail("ADDRESS_2") %></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>Address:</td>
      <td><%= RSCustDetail("ADDRESS_1") %></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td width="15%">City:</td>
      <td width="40%"><%= RSCustDetail("CITY") %></td>
      <td width="15%"></td>
      <td width="30%"></td>
      <td> </td>
    </tr>
    <tr>
      <td>State:</td>
      <td><%= RSCustDetail("STATE") %></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td>Zip:</td>
      <td><%= RSCustDetail("ZIP") %></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td>Country:</td>
      <td><%= RSCustDetail("COUNTRY") %></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td>Phone:</td>
      <td><%= RSCustDetail("PHONE") %></td>
      <td>Fax:</td>
      <td><%= RSCustDetail("FAX") %></td>
    </tr>
    <tr>
      <td>E-Mail:</td>
      <td><%= RSCustDetail("E_MAIL") %></td>
      <td></td>
      <td></td>
    </tr>
<% If (RSCustDetail("SHIP_STATE")) <> "" AND (RSCustDetail("SHIP_COUNTRY")) <> "" Then %>
    <tr>
      <td colspan="4"><big><u>Ship To</u></big></td>
    </tr>
    <tr>
      <td width="10%">Name:</td>
      <td width="40%"><%= RSCustDetail("SHIP_NAME") %></td>
      <td width="10%">&nbsp;</td>
      <td width="40">&nbsp;</td>
    </tr>
    <tr>
      <td>Company:</td>
      <td><%= RSCustDetail("SHIP_COMPANY") %></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>Address:</td>
      <td><%= RSCustDetail("SHIP_ADDRESS_2") %></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>Address:</td>
      <td><%= RSCustDetail("SHIP_ADDRESS_1") %></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td width="15%">City:</td>
      <td width="40%"><%= RSCustDetail("SHIP_CITY") %></td>
      <td width="15%"></td>
      <td width="30%"></td>
      <td> </td>
    </tr>
    <tr>
      <td>State:</td>
      <td><%= RSCustDetail("SHIP_STATE") %></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td>Zip:</td>
      <td><%= RSCustDetail("SHIP_ZIP") %></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td>Country:</td>
      <td><%= RSCustDetail("SHIP_COUNTRY") %></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td>Phone:</td>
      <td><%= RSCustDetail("SHIP_TELEPHONE") %></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
<% End If %>
    <tr>
      <td colspan="4">Special Instructions:</td>
    </tr>
    <tr>
      <td colspan="4"><%= RSCustDetail("SHIP_MESSAGE") %></td>
    </tr>	
    <% If Trim(RSCustDetail("PAYMENT_METHOD")) = "credit_card" Then %>

    <tr>
      <td bgcolor="#A09A8B" colspan="4"><big><u>Credit Card Information</u></big></td>
    </tr>
    <tr>
      <td>Card Type:</td>
      <td><%= RSCustDetail("CARD_TYPE") %></td>
      <td>Card Number:</td>
      <td><%= RSCustDetail("CARD_NO") %></td>
    </tr>
    <tr>
      <td>Card Name:</td>
      <td><%= RSCustDetail("NAME") %></td>
      <td>Expiration Date:</td>
      <td><%= RSCustDetail("CARD_EXP") %></td>
    </tr>

    <% ElseIf Trim(RSCustDetail("PAYMENT_METHOD")) = "e_check" Then %>
    <tr>
      <td bgcolor="#A09A8B" colspan="4"><big><u>E-Check Information</u></big></td>
    </tr>
    <tr>
      <td>Card Name:</td>
      <td><%= RSCustDetail("NAME") %></td>
      <td>Bank Name:</td>
      <td><%= RSCustDetail("BANK_NAME") %></td>
    </tr>
    <tr>
      <td>Routing Number:</td>
      <td><%= RSCustDetail("ROUTING_NO") %></td>
      <td>Account Number:</td>
      <td><%= RSCustDetail("CHK_ACCT_NO") %></td>
    </tr>
    <% 'ElseIf Trim(RSCustDetail("PAYMENT_METHOD")) = "purch_order" THEN %>

    <% 'ElseIf Trim(RSCustDetail("PAYMENT_METHOD")) = "phone_fax" THEN %>
    <% 'Else %>
    <% End If %>
  </table>
  </center></div><div align="center"><center>

  <table border="0" cellpadding="2" cellspacing="2" width="95%">
 
   <tr>
      <td bgcolor="#A09A8B" width="15%" align="center"><strong>Product Code</strong></td>
      <td bgcolor="#A09A8B" width="45%" align="center"><strong>Description</strong></td>
      <td bgcolor="#A09A8B" width="15%" align="center"><strong>Unit Price</strong></td>
      <td bgcolor="#A09A8B" width="10%" align="center"><strong>Quantity</strong></td>
      <td bgcolor="#A09A8B" width="15%" align="center"><strong>Total</strong></td>
    </tr>
    <%
	RSOrderDetail.MoveFirst
	Do While NOT RSOrderDetail.EOF
%>
    <tr>
      <td align="center"><%= RSOrderDetail("PRODUCT_ID") %></td>
      <td align="left"><%= RSOrderDetail("Description") %></td>
      <td align="right"><%= FormatCurrency(RSOrderDetail("PRICE")) %></td>
      <td align="center"><%= RSOrderDetail("QUANTITY") %></td>
      <td align="right"><%= FormatCurrency(RSOrderDetail("TOTAL")) %></td>
    </tr>
    <%
	RSOrderDetail.MoveNext
	
	Loop
%></center>
    <tr>
	<td colspan="5"><hr>
    </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td colspan="2">Sub Total</td>
      <td align="right"><%= RSCustDetail("SUB_TOTAL") %></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td colspan="2">Tax</td>
      <td align="right"><%= RSCustDetail("TAX") %></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
   
      <% If RSCustDetail("SHIPPING_AIR") > 1 Then %>
      <td colspan="2">Premium Shipping:</td>
      <td align="right"> <%= FormatCurrency(RSCustDetail("SHIPPING_AIR")) %></td>
      <% ElseIf RSCustDetail("SHIPPING_UPS") > 1 Then %>
      <td colspan="2">Standard Shipping:</td>
      <td align="right"> <%= FormatCurrency(RSCustDetail("SHIPPING_UPS")) %></td>
      <% End If %>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td colspan="2">Total</td>
      <td align="right"><%= RSCustDetail("GRAND_TOTAL") %></td>
    </tr>
    <tr>
      <td align="right" colspan="5"><form action="salesreports.asp?Report=3&amp;OrderID=<%= Request("OrderID") %>" method="post">
          <p><input type="submit" name="Create Invoice" value="PRINT INVOICE"></p>
        </form>
      </td>
    </tr>
    <tr>
      <td align="right" colspan="5"><form action="salesreports.asp?Report=5&amp;OrderID=<%= Request("OrderID") %>" method="Post">
          <p><input type="submit" value="DELETE RECORD"></p>
        </form>
      </td>
    </tr>
<% If Trim(RSCustDetail("PAYMENT_METHOD")) = "e_check" Then %>
	<% If Trim(RSCustDetail("ROUTING_NO")) <> "AUTHNET TRANSACTION" Then %>
    <tr>
      <td align="right" colspan="5"><form action="salesreports.asp?Report=6&amp;OrderID=<%= Request("OrderID") %>" method="Post">
          <p><input type="submit" value="PRINT CHECK"></p>
        </form>
      </td>
    </tr>
<% End If %>
<% End If %>
  </table>
  </center></div>

<% ElseIf Request("Report") = "3" Then %>

<!--#include file="invoice_head.htm"-->

<div align="center"><center><font color="#000000">

  <table border="0" cellpadding="3" cellspacing="0" width="95%">
   <tr>
      <td colspan="2"><big><u>Sold To</u></big></td>
    </tr>

    <tr>
      <td width="50%"><%= RSCustDetail("NAME") %></td>
      <td width="50%">Order Date: <%= FormatDateTime(RSCustDetail("ORDER_DATE"),vbShortDate) %></td>
    </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("COMPANY") %></td>
      <td width="50%">Order No: <%=RSCustDetail("Customer_ID") %></td>
    </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("ADDRESS_1") %></td>
      <td width="50%">Payment Method:
	<% If Trim(RSCustDetail("PAYMENT_METHOD")) = "credit_card" Then %> &nbsp;Credit Card
	<% ElseIf Trim(RSCustDetail("PAYMENT_METHOD")) = "e_check" Then %> &nbsp;Electronic Check
	<% ElseIf Trim(RSCustDetail("PAYMENT_METHOD")) = "purch_order" Then %>&nbsp;Purchase Order
	<% ElseIf Trim(RSCustDetail("PAYMENT_METHOD")) = "phone_fax" Then %>&nbsp; Phone or Fax
	<% Else %>
	<% End If %></td>
    </tr>
      <td width="50%"><%= RSCustDetail("ADDRESS_2") %></td>
      <td width="50%">&nbsp;</td>
    </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("CITY") %></td>
      <td width="50%"></td>
    </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("STATE") %></td>
      <td width="50%"></td>
    </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("ZIP") %></td>
      <td width="50%"></td>
    </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("COUNTRY") %></td>
      <td width="50%"></td>
    </tr>
<% If (RSCustDetail("SHIP_STATE")) <> "" AND (RSCustDetail("SHIP_COUNTRY")) <> "" Then %>
    <tr>
      <td colspan="2"><big><u>Ship To</u></big></td>
    </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("SHIP_NAME") %></td>
      <td width="50%">&nbsp:</td>
      </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("SHIP_COMPANY") %></td>
      <td width="50%">&nbsp;</td>
    </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("SHIP_ADDRESS_2") %></td>
      <td width="50%">&nbsp;</td></tr>
    <tr>
      <td width="50%"><%= RSCustDetail("SHIP_ADDRESS_1") %></td>
      <td width="50%">&nbsp;</td>
     </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("SHIP_CITY") %></td>
      <td width="50%">&nbsp:</td>
      </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("SHIP_STATE") %></td>
      <td width="50%">&nbsp:</td>
     </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("SHIP_ZIP") %></td>
      <td width="50%">&nbsp:</td>
     </tr>
    <tr>
      <td width="50%"><%= RSCustDetail("SHIP_COUNTRY") %></td>
      <td width="50%">&nbsp:</td>
     </tr>
    <tr>
<% End If %>
  </table>
  <table border="0" cellpadding="2" cellspacing="2" width="95%">
   <tr>
      <td bgcolor="#A09A8B" width="15%" align="center"><strong>Product Code</strong></td>
      <td bgcolor="#A09A8B" width="45%" align="center"><strong>Description</strong></td>
      <td bgcolor="#A09A8B" width="15%" align="center"><strong>Unit Price</strong></td>
      <td bgcolor="#A09A8B" width="10%" align="center"><strong>Quantity</strong></td>
      <td bgcolor="#A09A8B" width="15%" align="center"><strong>Total</strong></td>
    </tr>
  
  <%
	CurrentRecord = 0

	Do While NOT RSOrderDetail.EOF
%> 
    <tr>
      <td align="left"><%= RSOrderDetail("PRODUCT_ID") %></td>
      <td align=" left"><%= RSOrderDetail("Description") %></td>
      <td align="right"><%= FormatCurrency(RSOrderDetail("PRICE")) %> &nbsp;</td>
      <td align="right"><%= RSOrderDetail("QUANTITY") %> &nbsp;</td>
      <td align="right"><%= FormatCurrency(RSOrderDetail("TOTAL")) %> &nbsp;</td>
    </tr>
  
 <%
	RSOrderDetail.MoveNext
		
	CurrentRecord = CurrentRecord = 1
	
	Loop
%>
    <tr>
      <td colspan="5"><hr></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
      <td colspan="2">Sub Total</td>
      <td align="right"><%= RSCustDetail("SUB_TOTAL") %> &nbsp;</td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
      <td colspan="2">Tax</td>
      <td align="right"><%= RSCustDetail("TAX") %> &nbsp;</td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
      <% If RSCustDetail("SHIPPING_AIR") > 1 Then %><td colspan="2">Premium Shipping:</td>
      <td align="right"><%= FormatCurrency(RSCustDetail("SHIPPING_AIR")) %> &nbsp;</td>
      <% ElseIf RSCustDetail("SHIPPING_UPS") > 1 Then %><td colspan="2">Standard Shipping:</td>
      <td align="right"><%= FormatCurrency(RSCustDetail("SHIPPING_UPS")) %> &nbsp;</td>
      <% End If %>
    </tr>
    <tr>
      <td colspan="2"></td>
      <td colspan="2">Total </td>
      <td align="right"><%= RSCustDetail("GRAND_TOTAL") %> &nbsp;</td>
    </tr>
  </table>
  </font></center></div>

<% ElseIf Request("Report") = "4" Then %>

<% If SumGrandTotal = "0" Then %>

<p align="center"><big><strong>There were no sales for the period 
<%= Request("StartDate") %> to <%= Request("EndDate") %></strong></big></p>
<div align="center"><center><% Else %>

  <p align="center"><big><strong>Summary Sales Reports for the period <%= Request("StartDate") %> to <%= Request("EndDate") %></strong></big></p>
  <div align="center"><center><%
	CurrentRecord = 0

	Do While NOT RSSummary.EOF
%>

    <table border="0" cellpadding="3" cellspacing="0" width="95%">
      <tr>
        <td width="50%">Net Sales</td>
        <td width="50%"><%= FormatCurrency(SumSubTotal,2) %></td>
      </tr>
      <tr>
        <td width="50%">Total Tax</td>
        <td width="50%"><%= FormatCurrency(SumTax,2) %></td>
      </tr>
      <tr>
        <td width="50%">Total Shipping Standard</td>
        <td width="50%"><%= FormatCurrency(SumShipUPS,2) %></td>
      </tr>
      <tr>
        <td width="50%">Total Shipping Premium</td>
        <td width="50%"><%= FormatCurrency(SumShipAIR,2) %></td>
      </tr>
      <tr>
        <td width="50%">Total Shipping</td>
        <td width="50%"><%= FormatCurrency((SumShipAIR)+(SumShipUPS),2) %></td>
      </tr>
      <tr>
        <td width="50%">Total Gross Sales</td>
        <td width="50%"><%= FormatCurrency(SumGrandTotal,2) %></td>
      </tr>
      <%
	RSSummary.MoveNext
		
	CurrentRecord = CurrentRecord = 1
	
	Loop
%>
    </table>
    </center></div> 
<% End If %>
  <% ElseIf Request("Report") = "5" Then %> 


  <center><strong>Transaction Number 
  <%= Request("OrderID") %> has been deleted.</strong></center> 

  <% ElseIf Request("Report") = "6" Then %>

  <table width="90%" border="3" cellpadding="0" cellspacing="5" align="left">
    <tr>
      <td>

        <center>
        <table width="90%" border="1" valign="center">
          <tr>
            <td>
              <table width="90%">
                <tr>
                  <td><small><b><%= RSChck("NAME") %></b></small></td>
                </tr>
                <tr>

                  <td><small><b><%= RSChck("Address_2") %></b></small></td>
                </tr>

                <tr>
                  <% If RSChck("Address_1") > "" Then %>
                  <td>&nbsp;</td>
                  <td><small><b><%= RSChck("Address_1") %></b></small></td>
                </tr>
                <% Else %>
                <% End If %>
                <td><small><b><%= RSChck("City")&","&"&nbsp;"&RSChck("State")&"&nbsp;&nbsp;"&RSChck("Zip") %></b></small></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td nowrap><small><b>Pay To:</b></small>&nbsp;&nbsp;<u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u>&nbsp;&nbsp;
                </u></td>
            </tr>
            <tr>
              <td>&nbsp;
              </td>
            </tr>
            <tr>
              <td colspan="4" nowrap><small><b><u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </u>&nbsp;&nbsp;&nbsp;<b>Amount: &nbsp;&nbsp;<u><%= FormatCurrency(RSChck("GRAND_TOTAL")) %></u></b></small></td>
              </tr>

              <tr>

                <td><big><b><%= RSChck("BANK_NAME") %></b></big></td>
              </tr>
              <tr>

                <td>&nbsp;</td>
              </tr>
              <tr>

                <td colspan="4" nowrap><small><b>Memo:</b></small><u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u>
&nbsp;&nbsp;<u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;</u>&nbsp;&nbsp;
                </td>
              </tr>
              <tr>
                <td colspan="4" nowrap>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <font face="MICR 013 BT" size="5"><b>A<%= Trim(RSChck("ROUTING_NO")) %>A<%= Trim(RSChck("CHK_ACCT_NO")) %>B</font>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </center>
        </td>
      </tr>
    </table>
  </center>


<% End If %>
<% 'End If %>
  <% If Request("Report") = "3" OR Request("Report") = "6" Then %> 
  <%'Do Nothing %>
  <% Else %>

  <p align="center"><small><a href="prodadd.htm">Add Product</a> | <a href="proddelete.htm">Delete Product</a> | <a href="prodlist.asp">List Products</a> | <a href="prodedit.htm">Edit Product</a><br>
  <a href="reports.htm">Sales Reporting</a> | <a href="set_up.asp?Update=0">Store Set-Up</a></small></p>
  </div>

  <% End If %></p>
<% Connection.Close %>
</body>
</html>
