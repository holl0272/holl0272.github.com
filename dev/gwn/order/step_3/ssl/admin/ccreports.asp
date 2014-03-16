<%@ LANGUAGE="VBSCRIPT" %><%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront v.2.5.1
'
'   Author     :   LaGarde, Incorporated
'
'   Description:   Produces the Payment Processor Transaction Reports
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

	Dim DSN_Name
	DSN_Name = Session("DSN_Name")

	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""

	If Request.QueryString("Report") = "1" Then
	
	SQLStmt = "SELECT * FROM transactions, customer, orders "
	SQLStmt = SQLStmt & "WHERE (((customer.ORDER_DATE >= #" & Request("StartDate") & "#) "
	SQLStmt = SQLStmt & "AND (customer.ORDER_DATE <= #" & Request("EndDate") & "#)) AND "
	SQLStmt = SQLStmt & "(customer.CUSTOMER_ID = transactions.ORDER_ID) AND "
	SQLStmt = SQLStmt & "(customer.CUSTOMER_ID = orders.ORDER_ID)) "
	Set RSCCDate = Connection.Execute(SQLStmt)
	
	ElseIf Request.QueryString("Report") = "2" Then
	
	
	SQLStmt = "SELECT * FROM Customer WHERE CUSTOMER_ID = "
	SQLStmt = SQLStmt & "" & Request("Order_ID") & " "
	'Response.Write (SQLStmt)
	Set RSCustDetail = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT * FROM orders, product WHERE ORDER_ID = "
	SQLStmt = SQLStmt & "" & Request("Order_ID") & " AND "
	SQLStmt = SQLStmt & "orders.Product_ID = product.Product_ID "
	'Response.Write (SQLStmt)
	Set RSOrderDetail = Connection.Execute(SQLStmt)
	
	End If
%><html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Payment Processor Transactions Report</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="Microsoft Theme" content="none">
</head>

<body>
<div align="center"><center><% If Request("Report") = "1" Then %><% If RSCCDate.EOF Then %><center><strong>

  <p>There are no reports for the period <%= Request("StartDate") %> to <%= Request("EndDate") %></strong></center> <% Else %> <center><strong>Transaction
Summary Report for: <%= Request("StartDate") %> to <%= Request("EndDate") %></strong></center> <% 
	CurrentRecord = 0

	Do While NOT RSCCDate.EOF
%> </p>

  <table border="2" cellpadding="3" cellspacing="0" width="95%">
    <tr>
      <td width="25%"><strong>Order ID</td>
        <td width="25%"><strong>Order Date</td>
          <td width="25%"><strong>Cust. Trans. No.</td>
            <td width="25%"><strong>Merch. Trans. No.</td>
            </tr>
            <tr>
              <td>&nbsp;<a HREF="ccreports.asp?Report=2&amp;ORDER_ID=<%= RSCCDATE("ORDER_ID") %>"><%= RSCCDate("ORDER_ID") %></a></td>
              <td>&nbsp;<%= RSCCDate("ORDER_DATE") %></td>
              <td>&nbsp;<%= RSCCDate("CUST_TRANS_NO") %></td>
              <td>&nbsp;<%= RSCCDate("MERCH_TRANS_NO") %></td>
            </tr>
            <tr>
              <td><strong>AVS Code</td>
                <td><strong>Aux. Msg.</td>
                  <td><strong>Action Code</td>
                    <td><strong>Retrieval Code</td>
                    </tr>
                    <tr>
                      <td>&nbsp;<%= RSCCDate("AVS_CODE") %></td>
                      <td>&nbsp;<%= RSCCDate("AUX_MSG") %></td>
                      <td>&nbsp;<%= RSCCDate("ACTION_CODE") %></td>
                      <td>&nbsp;<%= RSCCDate("RETRIEVAL_CODE") %></td>
                    </tr>
                    <tr>
                      <td><strong>Auth No.</td>
                        <td><strong>Error Msg.</td>
                          <td><strong>Error Loc.</td>
                            <td><strong>Status</td>
                            </tr>
                            <tr>
                              <td>&nbsp;<%= RSCCDate("AUTH_NO") %></td>
                              <td>&nbsp;<%= RSCCDate("ERROR_MSG") %></td>
                              <td>&nbsp;<%= RSCCDate("ERROR_LOCATION") %></td>
                              <td>&nbsp;<%= RSCCDate("STATUS") %></td>
                            </tr>
                          </table>
                          <%
	RSCCDate.MoveNext
		
	CurrentRecord = CurrentRecord = 1
	
	Loop

%><% End If %><% ElseIf Request("Report") = "2" Then %>

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
<% End If %>
</center></div>

                      <p align="center"><small><small><a href="prodadd.htm">Add Product</a> | <a href="proddelete.htm">Delete Product</a> | <a href="prodlist.asp">List Products</a> | <a href="prodedit.htm">Edit Product</a><br>
                      <a href="reports.htm">Sales Reporting</a> | <a href="set_up.asp?Update=0">Store Set-Up</a></small></small></p>
                    </body>
                  </html>
