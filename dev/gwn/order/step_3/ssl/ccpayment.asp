<%@ LANGUAGE="VBSCRIPT" %><%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront v.2.5.1
'
'   Author     :   LaGarde, Incorporated
'
'   Description:   This file invokes the CyberCash payment processing service.
'                  
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
	DSN_Name = Request("DSN_Name")
	
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


	SQLStmt = "SELECT GRAND_TOTAL from customer WHERE "
	SQLStmt = SQLStmt & "CUSTOMER_ID = " & ORDER_ID & ""
	RSOrderCheck = Connection.Execute(SQLStmt)
	'If RSOrderCheck("GRAND_TOTAL") < 1 Then
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
	
	
	If Shipping_PRM > 1 Then
		Shipping = cCur(SpShipping)
				
	ElseIf Shipping_STD > 1 Then
		Shipping = cCur(Shipping)
	End If
		
	SubTotal = RSSumOrd("SubTotal")
	
	If SHIP_COUNTRY = TaxCountry AND SHIP_STATE = TaxState Then
	Tax = ((RSStateTax("StateTax"))+(RSCountryTax("CountryTax")))
		Tax = CCur(TAX)
		Tax = FormatCurrency(Tax,2)
		Grand_Total = ((SubTotal)+(Tax)+(Shipping))
	ElseIf SHIP_COUNTRY = "" Then
		If Country = TaxCountry AND State = TaxState Then
		Tax = ((RSStateTax("StateTax"))+(RSCountryTax("CountryTax")))
		Tax = CCur(Tax)
		Tax = FormatCurrency(Tax,2)
		Grand_Total = ((SubTotal)+(Tax)+(Shipping))
		Else 
		Grand_Total = ((SubTotal)+(Shipping))
		End If
	Else
		Grand_Total = ((SubTotal)+(Shipping))

	End If

	Grand_Total = FormatCurrency(Grand_Total,2)
	ccCard_No = Request("CARD_NO")

	ccExp_Date = Request("CARD_EXP")

	ccName = Request("CARD_NAME")

	ccAddress = Request("ADDRESS_2")&","&Request("ADDRESS_1")

	ccCity = Request("CITY")

	ccState = Request("STATE")

	ccZip = Request("ZIP")

	ccCountry = Request("COUNTRY")

	SMPShost = Trim( RSAdmin("PAYMENT_SERVER") )
	Secret   = Trim( RSAdmin("LOGIN")          )
	Amt      = FormatCurrency(GRAND_TOTAL,2)
	Amt = Mid( Amt, 2 )
	OrderID = Request.QueryString("ORDER_ID")


	Set cc = Server.CreateObject("CyberCash.PaymentServer.1")

	cc.PaymentServerHost = SMPShost
	cc.PaymentServerSecret = Secret

	cc.Clear
	
	cc.OrderID = OrderID
	cc.Amount = Amt
	cc.CreditCardNumber = ccCard_No
	cc.ExpirationStr = ccExp_Date
	cc.Name = ccName
	cc.Address = ccAddress
	cc.City = ccCity
	cc.State = ccState
	cc.Zip = ccZip
	cc.Country = ccCountry

	MerchantType = Trim( RSAdmin("MERCHANT_TYPE") )
	MerchantType = UCase( MerchantType )

	If MerchantType = "AUTHCAPTURE" Then
	

	cc.AuthCapture
   	
	ElseIf MerchantType = "AUTHONLY" Then
	

	cc.AuthOnly
	
	End If

	
	Status = cc.Status

	CustomerTransactionNumber = cc.CustomerTransactionNumber
   	MerchantTransactionNumber = cc.MerchantTransactionNumber	
    	
	AvsCode = cc.AvsCode
	if Avscode = "" then AvsCode = "n/a"
	
    AuxMsg = cc.AuxMsg
	if Auxmsg = "" then AuxMsg = "n/a"
	
	ActionCode = cc.ActionCode
	if ActionCode = "" then ActionCode = "n/a"
	
    RetrievalCode = cc.RetrievalCode
	if RetrievalCode = "" then RetrievalCode = "n/a"
	
    AuthNum = cc.AuthNum
	if AuthNum = "" then AuthNum = "n/a"
	
	ErrorMsg = cc.ErrorMsg
	if ErrorMsg = "" then ErrorMsg = "n/a"

	ErrorLocation = cc.ErrorLocation
	if ErrorLocation = "" then ErrorLocation = "n/a"


	SQLStmt = " INSERT into transactions (ORDER_ID, CUST_TRANS_NO, MERCH_TRANS_NO, "
	SQLStmt = SQLStmt & " AVS_CODE, AUX_MSG, ACTION_CODE, RETRIEVAL_CODE, "
	SQLStmt = SQLStmt & " AUTH_NO, ERROR_MSG, ERROR_LOCATION, STATUS) "
	SQLStmt = SQLStmt & " VALUES('" & ORDER_ID & "', "
	SQLStmt = SQLStmt & " '" & CustomerTransactionNumber & "', "
	SQLStmt = SQLStmt & " '" & MerchantTransactionNumber & "', "
	SQLStmt = SQLStmt & " '" & AvsCode & "', '" & AuxMsg & "', "
	SQLStmt = SQLStmt & " '" & ActionCode & "', '" & RetrievalCode & "', "
	SQLStmt = SQLStmt & " '" & AuthNum & "', '" & ErrorMsg & "', "
	SQLStmt = SQLStmt & " '" & ErrorLocation & "', '" & Status & "')"
	
	Set RSCC = Connection.Execute(SQLStmt)
	
	CustomerTransactionNumber = Server.URLEncode(cc.CustomerTransactionNumber)	
	MerchantTransactionNumber = Server.URLEncode(cc.MerchantTransactionNumber)
    	
	AvsCode = Server.URLEncode(cc.AvsCode)
	if Avscode = "" then AvsCode = Server.URLEncode("n/a")
	
    	AuxMsg = Server.URLEncode(cc.AuxMsg)
	if Auxmsg = "" then AuxMsg = Server.URLEncode("n/a")
	
	ActionCode = Server.URLEncode(cc.ActionCode)
	if ActionCode = "" then ActionCode = Server.URLEncode("n/a")
	
    	RetrievalCode = Server.URLEncode(cc.RetrievalCode)

	if RetrievalCode = "" then RetrievalCode = Server.URLEncode("n/a")
	
    	AuthNum = Server.URLEncode(cc.AuthNum)

	if AuthNum = "" then AuthNum = Server.URLEncode("n/a")
	
    	ErrorMsg = Server.URLEncode(cc.ErrorMsg)
	
	if ErrorMsg = "" then ErrorMsg = Server.URLEncode("n/a")
	
    	ErrorLocation = Server.URLEncode(cc.ErrorLocation)
	
	if ErrorLocation = "" then ErrorLocation = Server.URLEncode("n/a")
	
	Dim CC_Response
	
	If Status = "success" Then

	SQLStmt = "UPDATE CUSTOMER SET NAME = '" & CARD_NAME & "', "
	SQLStmt = SQLStmt & "COMPANY = '" & COMPANY & "', "
	SQLStmt = SQLStmt & "ADDRESS_1 = '" & ADDRESS_1 & "', "
	SQLStmt = SQLStmt & "ADDRESS_2 = '" & ADDRESS_2 & "', "
	SQLStmt = SQLStmt & "CITY = '" & CITY & "', "
	SQLStmt = SQLStmt & "STATE = '" & STATE & "', "
	SQLStmt = SQLStmt & "ZIP = '" & Request("ZIP") & "', "
	SQLStmt = SQLStmt & "COUNTRY = '" & COUNTRY & "', "
	SQLStmt = SQLStmt & "PHONE = '" & Request("PHONE") & "', "
	SQLStmt = SQLStmt & "FAX = '" & Request("FAX") & "', "
	SQLStmt = SQLStmt & "E_MAIL = '" & Request("E_MAIL") & "', "
	SQLStmt = SQLStmt & "CARD_TYPE = '" & Request("CARD_TYPE") & "', "
	SQLStmt = SQLStmt & "CARD_NO = '" & Request("CARD_NO") & "', "
	SQLStmt = SQLStmt & "CARD_EXP = '" & Request("CARD_EXP") & "', "
	SQLStmt = SQLStmt & "SUB_TOTAL = '" & FormatCurrency(RSSumOrd("SubTotal"),2) & "', "
	SQLStmt = SQLStmt & "TAX = '" & FormatCurrency(TAX,2) & "', "
	If Request("Shipping_UPS") > 1 Then
	SQLStmt = SQLStmt & "SHIPPING_UPS = '" & FormatCurrency(Shipping,2) & "', "
	Else
	SQLStmt = SQLStmt & "SHIPPING_UPS = 0.00, "
	End If
	If Request("Shipping_AIR") > 1 Then
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
	SQLStmt = SQLStmt & "SHIP_ZIP = '" & Request("SHIP_ZIP") & "', "
	SQLStmt = SQLStmt & "SHIP_COUNTRY = '" & SHIP_COUNTRY & "', "
	SQLStmt = SQLStmt & "SHIP_TELEPHONE = '" & Request("SHIP_TELEPHONE") & "' "
	SQLStmt = SQLStmt & ", SHIP_MESSAGE = '" & SHIP_MESSAGE & "'"
	SQLStmt = SQLStmt & " WHERE CUSTOMER_ID = " & Request("ORDER_ID") & ""
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
	BODY = BODY & " E-Mail: " & EMAIL & "" & CR	
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
	BODY = BODY & " " & TaxState & " Sales Tax: " & FormatCurrency(TAX,2) & "" & CR
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
<html>

<body>

<table border="0" cellpadding="0" cellspacing="0" width="60%">
  <tr>
    <td width="50%" colspan="2"><p align="center"><font size="3"><b><u>Order Confirmation</u></b><br>
    </p>
    </td>
  </tr>
  <tr>
    <td width="50%" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td width="50%" align="right">Order Number:</td>
    <td width="50%" align="right">&nbsp;<%= Request.QueryString("ORDER_ID") %></td>
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
    <td width="50%" align="right"><% If Request("Shipping_AIR") = "" Then %> 2nd Day Air Shipping:</td>
    <td width="50%" align="right"><%= FormatCurrency(Shipping,2) %></td>
    <td width="50%" align="right"> <% ElseIf Request("Shipping_UPS") > 1 Then %> UPS Shipping:</td>
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
  <tr>
    <td><p align="center">&nbsp; <% Session.Abandon %> </p>
    </td>
  </tr>
</table>
</center>
</body>
</html>
<!--#include file="confirm_foot.htm"--><%		
	Else
		
		FailPath = "ccerror.asp"
		
		CC_Response = FailPath&"?CustomerTransactionNumber="&CustomerTransactionNumber&"&MerchantTransactionNumber="&MerchantTransactionNumber&"&AvsCode="&AvsCode&"&AuxMsg="&AuxMsg&"&ActionCode="&ActionCode&"&RetrievalCode="&RetrievalCode&"&AuthNum="&AuthNum&"&ErrorMsg="&ErrorMsg&"&ErrorLocation="&ErrorLocation
	
		Response.Redirect CC_Response 
		
	End If

	Session.Abandon

%>