<%@ LANGUAGE="VBScript" %>

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


	
	RESPONSECODE = Request("RESPONSECODE")
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
	If Request("METHOD") = "ACH" Then
	PAYMENT_METHOD = "e_check"
	Else
	PAYMENT_METHOD = "creditcard"
	End If
	ADDRESS_1 = Request("ADDRESS")
	ADDRESS_2 = Request("USER2")
	CITY = Request("CITY")
	STATE = Request("STATE")
	ZIP = Request("ZIP")
	COUNTRY = Request("COUNTRY")
	PHONE = Request("PHONE")
	FAX = Request("FAX")
	CUSTOMER = Request("EMAIL")
	DSN_NAME = Request("USER1")
	COMPANY = Request("USER3")
	SHIP_NAME = Request("USER4")
	SHIP_COMPANY = Request("USER5")
	SHIP_ADDRESS = Replace(Request("USER6"),"%40","@")
	SHIP_ADDRESS = Split(SHIP_ADDRESS,"@")
	SHIP_ADDRESS_2 = SHIP_ADDRESS(0)
	SHIP_ADDRESS_1 = SHIP_ADDRESS(1)
	SHIP_ZIP = SHIP_ADDRESS(2)
	SHIP_TELEPHONE = SHIP_ADDRESS(3)
	SHIP_MESSAGE = SHIP_ADDRESS(4)
	ShipMethod = Request("USER7")
	SHIP_CITY = Request("USER8")
	SHIP_STATE = Request("USER9")
	SHIP_COUNTRY = Request("USER10")
	SHIP_ZIP = Request("CUSTID")
	SHIP_MESSAGE = Replace(Request("Description"),"'","''")
	
	Set Connection = Server.CreateObject("ADODB.Connection")
		
	Connection.Open "DSN="&DSN_Name&""

	If  Request("CustomerTransactionNumber") = "" Then
	CustomerTransactionNumber = ORDER_ID
	End If

	If Request("TRANSID") = "" Then
	TRANSID = "N/A"
	End If

    	If Request("AVSDATA") = "" Then
	AVSDATA = "N/A"
	End If
	
   	If Request("AUXMSG") = "" Then
	AUXMSG = "N/A"
	End If

	
	If Request("RESPONSECODE") = "" Then
	RESPONSECODE = "N/A"
	End If
		
   	If Request("RetrievalCode") = "" Then
	RetrievalCode = "N/A"
	End If

	If Request("AUTHCODE") = "" Then
	AUTHCODE = "N/A"
	End If

	If Request("DECLINEREASON") = "" Then
	DECLINEREASON = "N/A"
	End If

	If Request("ErrorLocation") = "" Then
	ErrorLocation = "N/A"
	End If

	If Request("Status") = "" Then
	Status = "N/A"
	End If
	
	
	SQLStmt = " INSERT into transactions (ORDER_ID, CUST_TRANS_NO, MERCH_TRANS_NO, "
	SQLStmt = SQLStmt & " AVS_CODE, AUX_MSG, ACTION_CODE, RETRIEVAL_CODE, "
	SQLStmt = SQLStmt & " AUTH_NO, ERROR_MSG, ERROR_LOCATION, STATUS) "
	SQLStmt = SQLStmt & " VALUES(" & ORDER_ID & ", "
	SQLStmt = SQLStmt & " '" & CustomerTransactionNumber & "', "
	SQLStmt = SQLStmt & " '" & TRANSID & "', "
	SQLStmt = SQLStmt & " '" & AVSDATA & "', '" & AuxMsg & "', "
	SQLStmt = SQLStmt & " '" & RESPONSECODE & "', '" & RetrievalCode & "', "
	SQLStmt = SQLStmt & " '" & AUTHCODE & "', '" & DECLINEREASON & "', "
	SQLStmt = SQLStmt & " '" & ErrorLocation & "', '" & Status & "')"
	'Response.Write SQLStmt
	Set RSCC = Connection.Execute(SQLStmt)


	
		
	SQLStmt = "SELECT GRAND_TOTAL from customer WHERE "
	SQLStmt = SQLStmt & "CUSTOMER_ID = " & ORDER_ID & ""
	RSOrderCheck = Connection.Execute(SQLStmt)
	If RSOrderCheck("GRAND_TOTAL") < 1 Then

	If Request("RESPONSECODE") = "A" OR Request("RESPONSECODE") = "P" OR Request("RESPONSECODE") = "X" OR Request("RESPONSECODE") = "T" Then
	
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
	SQLStmt = SQLStmt & "CARD_NO = 'AUTHNET TRANSACTION', "
	SQLStmt = SQLStmt & "CARD_EXP = 'AUTHNET TRANSACTION', "
	SQLStmt = SQLStmt & "BANK_NAME = 'AUTHNET TRANSACTION', "
	SQLStmt = SQLStmt & "ROUTING_NO = 'AUTHNET TRANSACTION', "
	SQLStmt = SQLStmt & "CHK_ACCT_NO = 'AUTHNET TRANSACTION', "
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
	SQLStmt = SQLStmt & "GRAND_TOTAL = '" & FormatCurrency(GRAND_TOTAL,2) & "', "
	SQLStmt = SQLStmt & "SHIP_NAME = '" & SHIP_NAME & "', "
	SQLStmt = SQLStmt & "SHIP_COMPANY = '" & SHIP_COMPANY & "', "
	SQLStmt = SQLStmt & "SHIP_ADDRESS_1 = '" & SHIP_ADDRESS_1 & "', "
	SQLStmt = SQLStmt & "SHIP_ADDRESS_2 = '" & SHIP_ADDRESS_2 & "', "
	SQLStmt = SQLStmt & "SHIP_CITY = '" & SHIP_CITY & "', "
	SQLStmt = SQLStmt & "SHIP_STATE = '" & SHIP_STATE & "', "
	SQLStmt = SQLStmt & "SHIP_ZIP = '" & SHIP_ZIP & "', "
	SQLStmt = SQLStmt & "SHIP_COUNTRY = '" & SHIP_COUNTRY & "', "
	SQLStmt = SQLStmt & "SHIP_TELEPHONE = '" & SHIP_TELEPHONE & "', "
	SQLStmt = SQLStmt & "SHIP_MESSAGE = '" & SHIP_MESSAGE & "' "
	SQLStmt = SQLStmt & "WHERE CUSTOMER_ID = " & ORDER_ID & ""
	'Response.Write SQLStmt
	Set RSConfirm = Connection.Execute(SQLStmt)


	If InStr((CUSTOMER),"@") Then
	CUSTOMER = CUSTOMER
	Else
	CUSTOMER = ""
	End If

	If InStr((SECONDARY),"@") Then
	SECONDARY = RSECONDARY
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





	BODY = BODY & " Transaction Type: " & TransactionMethod & "" & CR
	If Request("Method") = "ACH" Then
	BODY = BODY & "Payment: E-Check" & CR
	End If	

	BODY = BODY & " Customer Transaction Number: " & CustomerTransactionNumber & "" & CR
	BODY = BODY & " Merchant Transaction number: " & TRANSID & "" & CR
	BODY = BODY & " AVS Code: " & AVSDATA & "" & CR
	BODY = BODY & " Auxillary BODY: " & AuxMsg & "" & CR
	BODY = BODY & " ActionCode: " & RESPONSECODE & "" & CR
	BODY = BODY & " Retrieval Code: " & RetrievalCode & "" & CR
	BODY = BODY & " Authorization Number: " & AUTHCODE & "" & CR
	
	BODY = BODY & " Order Sub Total: " & FormatCurrency(RSSumOrd("SubTotal"),2) & "" & CR
	If Tax > 0 Then
	BODY = BODY & " " & TaxState & " Sales Tax: " & FormatCurrency(TAX,2) & "" & CR
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
If InStr((CUSTOMER),"@") Then
SendMail_C = "CustomerMail"
End If
If InStr((SECONDARY),"@") Then
SendMail_S = "SecondaryMail"
End If


%>
<!--#include file="mail.inc"-->

<!--#include file="confirm_head.htm"-->
<html>

<body>
<div align="center"><center><% If Request("METHOD") = "ACH" Then %><b>

<p>Please allow up to seven days for your check to be processed</b> 

<% End If %> </p>

<table border="0" cellpadding="0" cellspacing="0" width="60%">
  <tr>
    <td width="50%" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td width="50%" align="right">Order Number:</td>
    <td width="50%" align="right">&nbsp;<%= ORDER_ID %></td>
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
    <td width="50%" align="right"><% If ShipMethod = "PRM" Then %> Premium Shipping:</td>
    <td width="50%" align="right"><%= FormatCurrency(Shipping,2) %></td>
    <td width="50%" align="right"> <% ElseIf ShipMethod = "STD" Then %> Standard Shipping:</td>
    <td width="50%" align="right"><%= FormatCurrency(Shipping,2) %><% End If %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Tax:&nbsp; </td>
    <td width="50%" align="right">&nbsp;<%= FormatCurrency(TAX,2)%></td>
  </tr>
  <tr>
    <td width="50%" align="right">Total Amount:</td>
    <td width="50%" align="right"><%= FormatCurrency(GRAND_TOTAL,2) %></td>
  </tr>
  </font>
</table>
</center></div>
</body>
</html>
<%
'And finally, we call the confirm_foot.htm for the footer to the confirmation screen.
%><!--#include file="confirm_foot.htm"--><% Session.Abandon %>
<% ElseIf Request("RESPONSECODE") = "T" Then %>
<!--#include file="confirm_head.htm"-->
<html>

<body>
<div align="center"><center><b><strong>

<p>Your Test Order Was Successful</strong></b> </p>

<table border="0" cellpadding="0" cellspacing="0" width="60%">
  <tr>
    <td width="50%" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td width="50%" align="right">Order Number:</td>
    <td width="50%" align="right">&nbsp;<%= ORDER_ID %></td>
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
    <td width="50%" align="right"><% If SHIP_METHOD = "AIR" Then %> 2nd Day Air Shipping:</td>
    <td width="50%" align="right"><%= FormatCurrency(Shipping,2) %></td>
    <td width="50%" align="right"> <% ElseIf SHIP_METHOD = "UPS" Then %> UPS Shipping:</td>
    <td width="50%" align="right"><%= FormatCurrency(Shipping,2) %><% End If %></td>
  </tr>
  <tr>
    <td width="50%" align="right">Tax:&nbsp; </td>
    <td width="50%" align="right">&nbsp;<%= FormatCurrency(TAX,2)%></td>
  </tr>
  <tr>
    <td width="50%" align="right">Total Amount:</td>
    <td width="50%" align="right"><%= FormatCurrency(GRAND_TOTAL,2) %></td>
  </tr>
  </font>
</table>
</center></div>
</body>
</html>
<%
'And finally, we call the confirm_foot.htm for the footer to the confirmation screen.
%>
<!--#include file="confirm_foot.htm"-->
<% Session.Abandon %>

<% 	ElseIf Request("RESPONSECODE") = "D" OR Request("RESPONSECODE") = "R" Then

		FailPath = "ccerror.asp"
		
		CC_Response = FailPath&"?Order_ID="&Request("INVOICE")&"&DSN_Name="&Request("USER1")&"&CustomerTransactionNumber="&CustomerTransactionNumber&"&MerchantTransactionNumber="&MerchantTransactionNumber&"&AvsCode="&AvsCode&"&AuxMsg="&AuxMsg&"&ActionCode="&ActionCode&"&RetrievalCode="&RetrievalCode&"&AuthNum="&AuthNum&"&ErrorMsg="&DECLINEREASON&"&ErrorLocation="&ErrorLocation
	
		'Response.Write CC_Response
		Response.Redirect CC_Response 
	
	End If

%>

<% Else %>
<% 

RespPage = "order_complete.asp?DSN_NAME="&DSN_NAME
Response.Redirect RespPage

 %>
<% End If %>