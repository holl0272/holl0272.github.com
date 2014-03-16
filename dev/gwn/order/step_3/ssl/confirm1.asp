<%@ LANGUAGE="VBScript" %>
<% Response.AddHeader "cache-control", "private" %>
<% Response.AddHeader "pragma", "no-cache" %>
<% Response.Expires = 0 %>
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront 99 v.3.0.1
'
'   Author     :   LaGarde, Incorporated
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
'   (c) Copyright 1998,1999 by LaGarde, Incorporated.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<%	

	If (Request("ECHODATA")) > "" Then

	ORDER_ID = Request("INVOICE")
	AUTHCODE = Request("AUTHCODE")
	DECLINEREASON = Request("DECLINEREASON")
	AVSDATA = Request("AVSDATA")
	TRANSID = Request("TRANSID")
	AUTHNET_AMOUNT = Request("AMOUNT")
	ORDER_ID = Request("INVOICE")
	CARD_TYPE = Request("METHOD")
	CUST_ID = Request("CUSTID")
	CARD_NAME = Request("NAME")
	PURCH_ORDER_NO = Request("RESPONSECODE")
	ADDRESS_1 = Request("ADDRESS")
	CITY = Request("CITY")
	STATE = Request("STATE")
	ZIP = Request("ZIP")
	COUNTRY = Request("COUNTRY")
	PHONE = Request("PHONE")
	FAX = Request("FAX")
	CUSTOMER = Request("EMAIL")
	DSN_NAME = Request("USER1")
	ADDRESS_2 = Request("USER2")
	COMPANY = Request("USER3")
	SHIP_NAME = Request("USER4")
	SHIP_COMPANY = Request("USER5")
	SHIP_ADDRESS_1 = Request("USER6")
	ShipMethod = Request("USER7")
	SHIP_CITY = Request("USER8")
	SHIP_STATE = Request("USER9")
	SHIP_COUNTRY = Request("USER10")
	SHIP_BODY = Replace(Request("SHIP_BODY"),"'","''")
	PAYMENT_METHOD = Request("METHOD")
	Set Connection = Server.CreateObject("ADODB.Connection")
		
	Connection.Open "DSN="&DSN_Name&""
	
	Else

	Dim DSN_Name
	DSN_Name = Session("DSN_Name")
	
	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""
	
	ORDER_ID = Request("ORDER_ID")
	CUST_NAME = Replace(Request("CUST_NAME"),"'","''")
	COMPANY = Replace(Request("COMPANY"),"'","''")
	ADDRESS_1 = Replace(Request("ADDRESS_1"),"'","''")
	ADDRESS_2 = Replace(Request("ADDRESS_2"),"'","''")
	CITY = Replace(Request("CITY"),"'","''")
	STATE = Replace(Request("STATE"),"'","''")
	COUNTRY = Replace(Request("COUNTRY"), "'","''")
	SHIP_NAME = Replace(Request("SHIP_NAME"),"'","''")
	SHIP_COMPANY = Replace(Request("SHIP_COMPANY"),"'","''")
	SHIP_ADDRESS_1 = Replace(Request("SHIP_ADDRESS_1"),"'","''")
	SHIP_ADDRESS_2 = Replace(Request("SHIP_ADDRESS_2"),"'","''")
	SHIP_CITY = Replace(Request("SHIP_CITY"),"'","''")
	SHIP_STATE = Replace(Request("SHIP_STATE"),"'","''")
	SHIP_COUNTRY = Replace(Request("SHIP_COUNTRY"), "'","''")
	SHIP_MESSAGE = Replace(Request("SHIP_MESSAGE"),"'","''")
	ZIP = Request("ZIP")
	PHONE = Request("PHONE")
	FAX = Request("FAX")
	CUSTOMER = Request("E_MAIL")
	PAYMENT_METHOD = Request("PAYMENT_METHOD")
	CARD_NAME = Replace(Request("CARD_NAME"),"'","''")
	CARD_TYPE = Request("CARD_TYPE")
	CARD_NO = Request("CARD_NO")
	CARD_EXP = Request("CARD_EXP")
	BANK_NAME = Replace(Request("BANK_NAME"),"'","''")
	ROUTING_NO = Request("BANK_CODE")
	CHK_ACCT_NO = Request("CHK_ACCT_NO")
	PURCH_ORDER_NO = Request("PURCH_ORDER_NO")
	SHIP_ZIP = Request("SHIP_ZIP")
	If Request("SHIPPING_STD") > 1 Then	
	ShipMethod = "STD"
	ElseIf Request("SHIPPING_PRM") > 1 Then
	ShipMethod = "PRM"
	End If
	End If

	SQLStmt = "SELECT GRAND_TOTAL from customer WHERE "
	SQLStmt = SQLStmt & "CUSTOMER_ID = " & ORDER_ID & ""
	RSOrderCheck = Connection.Execute(SQLStmt)
	If RSOrderCheck("GRAND_TOTAL") < 1 Then
	
	SQLStmt = "SELECT * FROM Admin"

	Set RSAdmin = Connection.Execute(SQLStmt)
	Login = RSAdmin("LOGIN")
	TaxCountry =  RSAdmin("TAX_COUNTRY")
	TaxState =  RSAdmin("TAX_STATE")
	StateTaxAmt = RSAdmin("STATE_TAX_AMOUNT")
	CountryTaxAmt = RSAdmin("COUNTRY_TAX_AMOUNT")
	TransactionMethod = RSAdmin("Transaction_Method")
	MailServer = Trim(RSAdmin("MAIL_SERVER"))
	MailMethod = Trim(RSAdmin("MAIL_METHOD"))
	PRIMARY = Trim(RSAdmin("PRIMARY_EMAIL"))
	SECONDARY = Trim(RSAdmin("SECONDARY_EMAIL"))
	SUBJECT = RSAdmin("EMAIL_SUBJECT")
	MESSAGE = RSAdmin("EMAIL_MESSAGE")
	ShippingA = CCur(RSAdmin("SHIPPING_A"))
	ShipAAmt = CCur(RSAdmin("SHIPA_AMOUNT"))
	ShippingB = CCur(RSAdmin("SHIPPING_B"))
	ShipABAmt = CCur(RSAdmin("SHIPAB_AMOUNT"))
	ShippingC = CCur(RSAdmin("SHIPPING_C"))
	ShipBCAmt = CCur(RSAdmin("SHIPBC_AMOUNT"))
	ShippingD = CCur(RSAdmin("SHIPPING_D"))
	ShipCDAmt = CCur(RSAdmin("SHIPCD_AMOUNT"))
	ShippingE = CCur(RSAdmin("SHIPPING_E"))
	ShipDEAmt = CCur(RSAdmin("SHIPDE_AMOUNT"))
	ShippingF = CCur(RSAdmin("SHIPPING_F"))
	ShipEFAmt = CCur(RSAdmin("SHIPEF_AMOUNT"))
	ShipF_UpAmt = CCur(RSAdmin("SHIPF_UP_AMOUNT"))

	SpShipAmt = CCur(RSAdmin("SPECIAL_SHIP_AMOUNT"))
	SpShipAAmt = CCur(ShipAAmt+(RSAdmin("SPECIAL_SHIP_AMOUNT")))
	SpShipBAmt = CCur(ShipABAmt+(RSAdmin("SPECIAL_SHIP_AMOUNT")))
	SpShipCAmt = CCur(ShipBCAmt+(RSAdmin("SPECIAL_SHIP_AMOUNT")))
	SpShipDAmt = CCur(ShipCDAmt+(RSAdmin("SPECIAL_SHIP_AMOUNT")))
	SpShipEAmt = CCur(ShipDEAmt+(RSAdmin("SPECIAL_SHIP_AMOUNT")))
	SpShipFAmt = CCur(ShipEFAmt+(RSAdmin("SPECIAL_SHIP_AMOUNT")))
	SpShipF_UpAmt = CCur(ShipF_upAmt+(RSAdmin("SPECIAL_SHIP_AMOUNT")))


	SQLStmt = "SELECT ORDERS.PRODUCT_ID, "
	SQLStmt = SQLStmt & "ORDERS.PRICE, "
	SQLStmt = SQLStmt & "ORDERS.TOTAL, "
	SQLStmt = SQLStmt & "ORDERS.QUANTITY, ORDERS.ID, PRODUCT.DESCRIPTION "
	SQLStmt = SQLStmt & "FROM ORDERS, PRODUCT "
	SQLStmt = SQLStmt & "WHERE ORDERS.ORDER_ID = " & ORDER_ID & " "
	SQLStmt = SQLStmt & "AND ORDERS.PRODUCT_ID = PRODUCT.PRODUCT_ID" 
	SET RSOrder = Connection.Execute(SQLStmt)
	
	Set RSProduct = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT (Sum(TOTAL)) AS SubTotal "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDER_ID & " "

	Set RSSumOrd = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT (Sum(TOTAL) * " & CountryTaxAmt & ") AS CountryTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDER_ID & " "
	Set RSCountryTax = Connection.Execute(SQLStmt)

	
	SQLStmt = "SELECT (Sum(TOTAL) * " & StateTaxAmt & ") AS StateTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDER_ID & " "
	Set RSStateTax = Connection.Execute(SQLStmt)


	
	If RSSumOrd("SubTotal") <= ShippingA Then
			Shipping = ShipAAmt
		ElseIf RSSumOrd("SubTotal") > ShippingA AND RSSumOrd("SubTotal") <= ShippingB Then
			Shipping = ShipABAmt
		ElseIf RSSumOrd("SubTotal") > ShippingB AND RSSumOrd("SubTotal") <= ShippingC Then
			Shipping = ShipBCAmt
		ElseIf RSSumOrd("SubTotal") > ShippingC AND RSSumOrd("SubTotal") <= ShippingD Then
			Shipping = ShipCDAmt
		ElseIf RSSumOrd("SubTotal") > ShippingD AND RSSumOrd("SubTotal") <= ShippingE Then
			Shipping = ShipDEAmt
		ElseIf RSSumOrd("SubTotal") > ShippingE AND RSSumOrd("SubTotal") <= ShippingF Then
			Shipping = ShipEFAmt
		ElseIf RSSumOrd("SubTotal") > ShippingG Then
			Shipping = ShipF_UpAmt
	End If

	
	If RSSumOrd("SubTotal") <= ShippingA Then
			SpShipping = SpShipAAmt
		ElseIf RSSumOrd("SubTotal") > ShippingA AND RSSumOrd("SubTotal") <= ShippingB Then
			SpShipping = SpShipBAmt
		ElseIf RSSumOrd("SubTotal") > ShippingB AND RSSumOrd("SubTotal") <= ShippingC Then
			SpShipping = SpShipCAmt
		ElseIf RSSumOrd("SubTotal") > ShippingC AND RSSumOrd("SubTotal") <= ShippingD Then
			SpShipping = SpShipDAmt
		ElseIf RSSumOrd("SubTotal") > ShippingD AND RSSumOrd("SubTotal") <= ShippingE Then
			SpShipping = SpShipEAmt
		ElseIf RSSumOrd("SubTotal") > ShippingE AND RSSumOrd("SubTotal") <= ShippingF Then
			SpShipping = SpShipEAmt
		ElseIf RSSumOrd("SubTotal") > ShippingF Then
			SpShipping = SpShipF_UpAmt
	End If
	
	
	If ShipMethod = "PRM" Then
		Shipping = cCur(SpShipping)
				
	ElseIf ShipMethod = "STD" Then
		Shipping = cCur(Shipping)
	End If
		
	SubTotal = RSSumOrd("SubTotal")
	
	If SHIP_COUNTRY = TaxCountry AND SHIP_STATE = TaxState Then
	Tax = ((CCur(RSStateTax("StateTax")))+(CCur(RSCountryTax("CountryTax"))))
		Tax = CCur(TAX)
		Tax = FormatCurrency(Tax,2)
		Grand_Total = ((SubTotal)+(Tax)+(Shipping))
	ElseIf SHIP_COUNTRY = "" Then
		If Country = TaxCountry AND State = TaxState Then
		Tax = ((CCur(RSStateTax("StateTax")))+(CCur(RSCountryTax("CountryTax"))))
		Tax = CCur(Tax)
		Tax = FormatCurrency(Tax,2)
		Grand_Total = ((SubTotal)+(Tax)+(Shipping))
		Else 
		Grand_Total = ((SubTotal)+(Shipping))
		End If
	Else
		Grand_Total = ((SubTotal)+(Shipping))

	End If
	If Request("PAYMENT_METHOD") = "phone_fax" Then
%>
<html>

<head>
<title>New Page </title>
</head>

<body>
<!--#include file="confirm_head.htm"-->

<table border="0" width="90%">
  <tr>
    <td width="50%">Customer Name</td>
    <td width="50%"><%= Request("Card_Name") %>
</td>
  </tr>
  <tr>
    <td width="50%">Address</td>
    <td width="50%"><%= Request("Address1") %>
</td>
  </tr>
  <tr>
    <td width="50%">Address</td>
    <td width="50%"><%= Request("Address2") %>
</td>
  </tr>
  <tr>
    <td width="50%">City</td>
    <td width="50%"><%= Request("City") %>
</td>
  </tr>
  <tr>
    <td width="50%">State</td>
    <td width="50%"><%= Request("State") %>
</td>
  </tr>
  <tr>
    <td width="50%">Country</td>
    <td width="50%"><%= Request("Country") %>
</td>
  </tr>
  <tr>
    <td width="50%">Zip</td>
    <td width="50%"><%= Request("Zip") %>
</td>
  </tr>
  <tr>
    <td width="50%">Phone</td>
    <td width="50%"><%= Request("Phone") %>
</td>
  </tr>
  <tr>
    <td width="50%">Fax</td>
    <td width="50%"><%= Request("Fax") %>
</td>
  </tr>
  <tr>
    <td width="50%">E-Mail</td>
    <td width="50%"><%= Request("E-Mail") %>
</td>
  </tr>
  <tr>
    <td width="50%">&nbsp;<p>Payment Method</td>
    <td width="50%">&nbsp;<p><input type="text" name="Payment_Method" size="60"></td>
  </tr>
  <tr>
    <td width="50%">Card Number</td>
    <td width="50%"><input type="text" name="<%= Request("CARD_NO") %>" size="60"></td>
  </tr>
  <tr>
    <td width="50%">Exp. Date</td>
    <td width="50%"><input type="text" name="<%= Request("EXP_DATE") %>" size="10"></td>
  </tr>
  <tr>
    <td width="50%">Bank Name</td>
    <td width="50%"><input type="text" name="<%= Request("BANK_NAME") %>" size="60"></td>
  </tr>
  <tr>
    <td width="50%">Bank Routing Number</td>
    <td width="50%"><input type="text" name="<%= Request("BANK_ROUTING_NO") %>" size="60"></td>
  </tr>
  <tr>
    <td width="50%">Checking Account Number</td>
    <td width="50%"><input type="text" name="<%= Request("CHK_ACCT_NO") %>" size="60"></td>
  </tr>
  <tr>
    <td width="50%">&nbsp;<p>Ship Name</td>
    <td width="50%">&nbsp;<p><%= Request("Ship_Name") %></td>
  </tr>
  <tr>
    <td width="50%">Ship Address</td>
    <td width="50%"><%= Request("Ship_Address") %>
</td>
  </tr>
  <tr>
    <td width="50%">City</td>
    <td width="50%"><%= Request("Ship_City") %>
</td>
  </tr>
  <tr>
    <td width="50%">State</td>
    <td width="50%"><%= Request("Ship_State") %>
</td>
  </tr>
  <tr>
    <td width="50%">Country</td>
    <td width="50%"><%= Request("Ship_Country") %>
</td>
  </tr>
  <tr>
    <td width="50%">Zip</td>
    <td width="50%"><%= Request("Ship_Zip") %>
</td>
  </tr>
  <tr>
    <td width="50%">Phone</td>
    <td width="50%"><%= Request("Ship_Phone") %>
</td>
  </tr>
</table>
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0" width="60%">
  <tr>
    <td width="50%" colspan="2"></td>
  </tr>
  <tr>
    <td width="50%" align="right">Order Number:</td>
    <td width="50%" align="right">&nbsp;<%= ORDER_ID %> </td>
  </tr>
  <tr>
    <td width="50%" align="right">Date: </td>
    <td width="50%" align="right">&nbsp;<%= Date() %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Order Amount: </td>
    <td width="50%" align="right">&nbsp;<%= FormatCurrency(RSSumOrd("SubTotal"),2) %></td>
  </tr>
  <tr>
    <td width="50%" align="right"><% If ShipMethod = "PRM" Then %>
<p>Premium Shipping:</td>
    <td width="50%" align="right"><%= FormatCurrency(Shipping,2) %>
</td>
    <td width="50%" align="right"><% ElseIf ShipMethod = "STD" Then %>
<p>Standard Shipping:</td>
    <td width="50%" align="right"><%= FormatCurrency(Shipping,2) %>
<% End If %>
</td>
  </tr>
<% If Tax > 0 Then %>
  <tr>
    <td width="50%" align="right">Tax:&nbsp; </td>
    <td width="50%" align="right">&nbsp;<%= FormatCurrency(TAX,2)%></td>
  </tr>
<% End If %>
  <tr>
    <td width="50%" align="right">Total Amount:</td>
    <td width="50%" align="right"><%= FormatCurrency(GRAND_TOTAL,2) %>
</td>
  </tr>
</table>
</center></div>
</body>
</html>
<% Else

	SQLStmt = "UPDATE CUSTOMER SET NAME = '" & CARD_NAME & "', "
	SQLStmt = SQLStmt & "COMPANY = '" & COMPANY & "', "
	SQLStmt = SQLStmt & "ADDRESS_1 = '" & ADDRESS_1 & "', "
	SQLStmt = SQLStmt & "ADDRESS_2 = '" & ADDRESS_2 & "', "
	SQLStmt = SQLStmt & "CITY = '" & CITY & "', "
	SQLStmt = SQLStmt & "STATE = '" & STATE & "', "
	SQLStmt = SQLStmt & "ZIP = '" & ZIP & "', "
	SQLStmt = SQLStmt & "COUNTRY = '" & COUNTRY & "', "
	SQLStmt = SQLStmt & "PHONE = '" & PHONE & "', "
	SQLStmt = SQLStmt & "FAX = '" & FAX & "', "
	SQLStmt = SQLStmt & "E_MAIL = '" & CUSTOMER & "', "
	SQLStmt = SQLStmt & "PAYMENT_METHOD = '" & PAYMENT_METHOD & "', "
	SQLStmt = SQLStmt & "CARD_TYPE = '" & CARD_TYPE & "', "
	SQLStmt = SQLStmt & "CARD_NO = '" & CARD_NO & "', "
	SQLStmt = SQLStmt & "CARD_EXP = '" & CARD_EXP & "', "
	SQLStmt = SQLStmt & "BANK_NAME = '" & BANK_NAME & "', "
	SQLStmt = SQLStmt & "ROUTING_NO = '" & ROUTING_NO & "', "
	SQLStmt = SQLStmt & "CHK_ACCT_NO = '" & CHK_ACCT_NO & "', "
	SQLStmt = SQLStmt & "PURCH_ORDER_NO = '" & PURCH_ORDER_NO & "', "
	SQLStmt = SQLStmt & "SUB_TOTAL = '" & FormatCurrency(RSSumOrd("SubTotal"),2) & "', "
	SQLStmt = SQLStmt & "TAX = '" & FormatCurrency(TAX,2) & "', "
	If ShipMethod = "STD" Then
	SQLStmt = SQLStmt & "SHIPPING_UPS = '" & FormatCurrency(Shipping,2) & "', "
	Else
	SQLStmt = SQLStmt & "SHIPPING_UPS = 0.00, "
	End If
	If ShipMethod = "PRM" Then
	SQLStmt = SQLStmt & "SHIPPING_AIR = '" & FormatCurrency(Shipping,2) & "', "
	Else
	SQLStmt = SQLStmt & "SHIPPING_AIR = 0.00, "
	End If
	SQLStmt = SQLStmt & "GRAND_TOTAL = '" & FormatCurrency(GRAND_TOTAL,2) & "' "
	SQLStmt = SQLStmt & ", SHIP_NAME = '" & SHIP_NAME & "', "
	SQLStmt = SQLStmt & "SHIP_COMPANY = '" & SHIP_COMPANY & "', "
	SQLStmt = SQLStmt & "SHIP_ADDRESS_1 = '" & SHIP_ADDRESS_1 & "', "
	SQLStmt = SQLStmt & "SHIP_ADDRESS_2 = '" & SHIP_ADDRESS_2 & "', "
	SQLStmt = SQLStmt & "SHIP_CITY = '" & SHIP_CITY & "', "
	SQLStmt = SQLStmt & "SHIP_STATE = '" & SHIP_STATE & "', "
	SQLStmt = SQLStmt & "SHIP_ZIP = '" & SHIP_ZIP & "', "
	SQLStmt = SQLStmt & "SHIP_COUNTRY = '" & SHIP_COUNTRY & "', "
	SQLStmt = SQLStmt & "SHIP_TELEPHONE = '" & SHIP_TELEPHONE & "' "
	SQLStmt = SQLStmt & ", SHIP_MESSAGE = '" & SHIP_MESSAGE & "'"
	SQLStmt = SQLStmt & " WHERE CUSTOMER_ID = " & ORDER_ID & ""
	'Response.Write sQLStmt
	Set RSConfirm = Connection.Execute(SQLStmt)


	If InStr((CUSTOMER),"@") Then
	CUSTOMER = CUSTOMER
	Else
	CUSTOMER = ""
	End If

	If InStr((SECONDARY),"@") Then
	SECONDARY = SECONDARY
	Else
	SECONDARY = ""
	End If

	SndPage = Request.ServerVariables("HTTP_REFERER")
	If InStr(SndPage, "process_order.asp")>0 Then
	SndPage = Left(SndPage, InStr(SndPage, "process_order.asp") - 1)
	End If
	Link = SndPage&"admin/salesreports.asp?REPORT=2&ORDERID="&ORDER_ID
	CR = Chr(10) & Chr(13)
	BODY1 = "Merchant Notification"& CR
	BODY2 = "Customer Order Confirmation"& CR
	BODY = "Order Number: " & ORDER_ID & "" & CR
	RSOrder.MoveFirst 
	Do While NOT RSOrder.EOF
        BODY = BODY & " Product ID: " & RSOrder("PRODUCT_ID") & "" & CR
        BODY = BODY & " Description: " & RSOrder("DESCRIPTION") & "" & CR
        BODY = BODY & " Price: " & FormatCurrency(RSOrder("PRICE"),2) & "" & CR
        BODY = BODY & " Quantity: " & RSOrder("QUANTITY") & "" & CR
        BODY = BODY & " Amount: " & FormatCurrency(RSOrder("TOTAL"),2) & "" & CR
	RSOrder.MoveNext
	Loop
	BODY = BODY & " Name: " & CARD_NAME & "" & CR
	BODY = BODY & " Company: " & COMPANY & "" & CR
	BODY = BODY & " Address1: " & ADDRESS_1 & "" & CR
	BODY = BODY & " Address2: " & ADDRESS_2 & "" & CR
	BODY = BODY & " City: " & CITY & "" & CR
	BODY = BODY & " State: " & STATE & "" & CR
	BODY = BODY & " Zip: " & ZIP & "" & CR
	BODY = BODY & " Country: " & Country & "" & CR
	BODY = BODY & " Phone: " & PHONE & "" & CR
	BODY = BODY & " E-Mail: " & CUSTOMER & "" & CR	
	BODY = BODY & " Lettering Color: " & Request("LET_COLOR") & "" & CR	
	BODY = BODY & " Reversible Lettering Color 1: " & Request("REV_LET_COL1") & "" & CR	
	BODY = BODY & " Reversible Color 1: " & Request("REV_COL1") & "" & CR	
	BODY = BODY & " Reversible Lettering Color 2: " & Request("REV_LET_COL2") & "" & CR	
	BODY = BODY & " Reversible Color 2: " & Request("REV_COL2") & "" & CR	
	BODY = BODY & " Style of Team Name: " & Request("STYLE_TEAM_NAME") & "" & CR	
	BODY = BODY & " Style of Individual Names: " & Request("STYLE_INDV_NAME") & "" & CR	
	BODY = BODY & " Team Name: " & Request("TEAM_NAME") & "" & CR	
	BODY = BODY & " Location of Team Name: " & Request("LOCATION_TEAM_NAME") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE1") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER1") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME1") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE2") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER2") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME2") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE3") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER3") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME3") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE4") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER4") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME4") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE5") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER5") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME5") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE6") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER6") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME6") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE7") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER7") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME7") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE8") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER8") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME8") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE9") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER9") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME9") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE10") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER10") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME10") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE11") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER11") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME11") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE12") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER12") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME12") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE13") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER13") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME13") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE14") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER14") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME14") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE15") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER15") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME15") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE16") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER16") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME16") & "" & CR		
	BODY = BODY & " Size: " & Request("SIZE17") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER17") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME17") & "" & CR		
	BODY = BODY & " Size: " & Request("SIZE18") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER18") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME18") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE19") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER19") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME19") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE20") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER20") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME20") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE21") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER21") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME21") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE22") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER22") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME22") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE23") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER23") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME23") & "" & CR	
	BODY = BODY & " Size: " & Request("SIZE24") & "" & CR	
	BODY = BODY & " Number: " & Request("NUMBER24") & "" & CR	
	BODY = BODY & " Name: " & Request("NAME24") & "" & CR	
	BODY = BODY & " Deliver Date: " & DELIVER_DATE & "" & CR
	BODY = BODY & " Payment Method: " & PAYMENT_METHOD & "" & CR
	BODY = BODY & " Account Name: " & CARD_NAME & "" & CR
	  If RSAdmin("MAIL_CC") = "on" Then
	      If PAYMENT_METHOD = "credit_card" Then
	         BODY = BODY & " Card Type: " & CARD_TYPE & "" & CR
	         BODY1 = BODY1 & " Card Number: " & CARD_NO & "" & CR
	         BODY1 = BODY1 & " Card Expiration: " & CARD_EXP & "" & CR
              ElseIf PAYMENT_METHOD = "e_check" Then
	          BODY = BODY & " Bank Name: " & BANK_NAME & "" & CR
	          BODY1 = BODY1 & " Routing Number: " & ROUTING_NO & "" & CR
	          BODY1 = BODY1 & " Account Number: " & ACCOUNT_NO & "" & CR
	      End If
          ElseIf PAYMENT_METHOD = "credit_card" OR PAYMENT_METHOD = "e_check" Then
	      BODY1 = BODY1 & " Retrieve Payment Detail: " & CR
			BODY1 = BODY1 &  LINK & CR
	  ElseIf PAYMENT_METHOD = "purch_order" THEN
	      BODY = BODY & " Purchase Order Number: " & PURCH_ORDER_NO & "" & CR
	Else	
	End If

	If PAYMENT_METHOD = "credit_card" Then
	      BODY2 = BODY2 & " Card Number:  ****-****-****-****" & CR
	      BODY2 = BODY2 & " Card Expiration:  **/** " & CR
	ElseIf PAYMENT_METHOD = "e_check" Then
	     BODY2 = BODY2 & " Routing Number: *********" & CR
	     BODY2 = BODY2 & " Account Number: *************" & CR
	End If
	BODY = BODY & " Order Sub Total: " & FormatCurrency(RSSumOrd("SubTotal"),2) & "" & CR
	If Tax > 0 Then
	BODY = BODY & " Sales Tax: " & FormatCurrency(TAX,2) & "" & CR
	End If
	If ShipMethod = "PRM" Then
	BODY = BODY & " Premium Shipping: " & FormatCurrency(SpShipping,2) & "" & CR
	Else
	BODY = BODY & " Standard Shipping " & FormatCurrency(Shipping,2) & "" & CR
	End If
	BODY = BODY & " Grand Total: " & FormatCurrency(GRAND_TOTAL) & "" & CR
	BODY = BODY & " Ship To: " & SHIP_NAME & "" & CR
	BODY = BODY & " Company Name: " & SHIP_COMPANY & "" & CR
	BODY = BODY & " Shipping Address1: " & SHIP_ADDRESS_1 & "" & CR
	BODY = BODY & " Shipping Address2: " & SHIP_ADDRESS_2 & "" & CR
	BODY = BODY & " Shipping City: " & SHIP_CITY & "" & CR
	BODY = BODY & " Shipping State: " & SHIP_STATE & "" & CR
	BODY = BODY & " Shipping Zip Code: " & SHIP_ZIP & "" & CR
	BODY = BODY & " Shipping Country: " & SHIP_COUNTRY & "" & CR
	BODY = BODY & " Shipping Telephone: " & SHIP_TELEPHONE & "" & CR
	BODY = BODY & " Special Instructions: " & Request("SHIP_MESSAGE") & "" & CR

	BODY1 = BODY&BODY1
	BODY2 = MESSAGE&BODY&BODY2

SendMail_M = "PrimaryMail"
If CUSTOMER <> "" Then
SendMail_C = "CustomerMail"
End If
If SECONDARY <> "" Then
SendMail_S = "SecondaryMail"
End If

%>
<!--#include file="mail.inc"-->
<!--#include file="confirm_head.htm"-->
<% End If %>
<%
'And finally, we call the confirm_foot.htm for the footer to the confirmation screen.
%>
<!--#include file="confirm_foot.htm"-->
<% Session.Abandon %>
<% Else %>
<% 
	RespPage = "order_complete.asp?DSN_Name="&DSN_Name
	Connection.Close
	'response.write RespPage
	Session.Abandon
	Response.Redirect RespPage %>
<% End If %>
