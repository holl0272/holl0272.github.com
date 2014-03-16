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
	CUSTOMER = Request("E_MAIL")
	CARD_NAME = Replace(Request("CARD_NAME"),"'","''")
	SHIP_NAME = Replace(Request("SHIP_NAME"),"'","''")
	SHIP_COMPANY = Replace(Request("SHIP_COMPANY"),"'","''")
	SHIP_ADDRESS_1 = Replace(Request("SHIP_ADDRESS_1"),"'","''")
	SHIP_ADDRESS_2 = Replace(Request("SHIP_ADDRESS_2"),"'","''")
	SHIP_CITY = Replace(Request("SHIP_CITY"),"'","''")
	SHIP_STATE = Replace(Request("SHIP_STATE"),"'","''")
	SHIP_COUNTRY = Replace(Request("SHIP_COUNTRY"), "'","''")
	SHIP_MESSAGE = Replace(Request("SHIP_MESSAGE"),"'","''")

	SQLStmt = "SELECT * FROM Admin"

	Set RSAdmin = Connection.Execute(SQLStmt)

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

	If NOT RSOrder.EOF Then
	
	Set RSProduct = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT (Sum(TOTAL)) AS SubTotal "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDER_ID & " "

	Set RSSumOrd = Connection.Execute(SQLStmt)
	
	SQLStmt = "SELECT (Sum(TOTAL) * " & CountryTaxAmt & ") AS CountryTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDER_ID & " "
	'Response.Write SQLStmt
	Set RSCountryTax = Connection.Execute(SQLStmt)

	
	SQLStmt = "SELECT (Sum(TOTAL) * " & StateTaxAmt & ") AS StateTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDER_ID & " "
	'Response.Write SQLStmt
	Set RSStateTax = Connection.Execute(SQLStmt)
	Set RSTax = Connection.Execute(SQLStmt)


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
	

	If Request("Shipping_AIR") > 1 Then
		Shipping = cCur(SpShipping)
				
	ElseIf Request("Shipping_UPS") > 1 Then
		Shipping = cCur(Shipping)
	End If
		
	SubTotal = RSSumOrd("SubTotal")

		If SHIP_COUNTRY = TaxCountry AND SHIP_STATE = TaxState Then
	Tax = ((CCur(RSStateTax("StateTax")))+(CCur(RSCountryTax("CountryTax"))))
		Tax = CCur(TAX)
		Grand_Total = ((SubTotal)+(Tax)+(Shipping))
	ElseIf SHIP_COUNTRY = "" Then
		If Country = TaxCountry AND State = TaxState Then
		Tax = ((CCur(RSStateTax("StateTax")))+(CCur(RSCountryTax("CountryTax"))))
		Tax = CCur(Tax)
		Grand_Total = ((SubTotal)+(Tax)+(Shipping))
		Else 
		Grand_Total = ((SubTotal)+(Shipping))
		End If
	Else
		Grand_Total = ((SubTotal)+(Shipping))

	End If


	ccCard_No = Request("CARD_NO")

	ccExp_Date = Request("CARD_EXP")

	ccName = Request("CARD_NAME")

	ccAddress = Request("ADDRESS_2")&","&Request("ADDRESS_1")

	ccCity = Request("CITY")

	ccState = Request("STATE")

	ccZip = Request("ZIP")

	ccCountry = Request("COUNTRY")


Set AuthCtl = Server.CreateObject("AuthCtl.AuthCtlCtrl.1")
AuthCtl.Hub=1
AuthCtl.Protocol="SOCKETS"
AuthCtl.ApplName="PCAuthorize"
AuthCtl.HostName=RSAdmin("PAYMENT_SERVER")
AuthCtl.Port=54321
AuthCtl.Account = ccCard_No
AuthCtl.Amount = FormatCurrency(GRAND_TOTAL,2)
AuthCtl.ExpDate = ccExp_Date
AuthCtl.Invoice = ORDER_ID
AuthCtl.CustomerID = ORDER_ID
AuthCtl.AVSAddress = ccAddress
AuthCtl.AVSZip = ccZip
AuthCtl.Authorize

Error = AuthCtl.ErrMsg
Message = AuthCtl.RespText
Status = AuthCtl.RespStatus
RefNum = AuthCtl.RespRefNum
AVSCode = AuthCtl.RespAVSCode
AuthCode = AuthCtl.RespAuthCode

If RespStatus = "PD" OR RespStatus = "AA" Then 
Result = "Approve"

ElseIf RespStatus = "ND" Then
Result = "Decline"

ElseIf RespStatus = "NR" OR RespStatus = "E1" or RespStatus = "F1" Then
Result = "Call"

End If

	CustomerTransactionNumber = AuthCtl.RespRefNum
   	MerchantTransactionNumber = AuthCtl.RespAuthCode
    	
	AvsCode = AuthCtl.RespAVSCode
	if Avscode = "" then AvsCode = "n/a"
	
    AuxMsg = AuthCtl.RespText
	if Auxmsg = "" then AuxMsg = "n/a"
	
	ActionCode = AuthCtl.RespStatus
	if ActionCode = "" then ActionCode = "n/a"
	
    RetrievalCode = cc.RetrievalCode
	if RetrievalCode = "" then RetrievalCode = "n/a"
	
    AuthNum = AuthCtl.RespAuthCode
	if AuthNum = "" then AuthNum = "n/a"
	
	ErrorMsg = AuthCtl.ErrMsg
	if ErrorMsg = "" then ErrorMsg = "n/a"
	
	ErrorLocation = "n/a"


	
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
	
	
	
	If RespStatus = "PD" OR RespStatus = "AA" Then 


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
	SQLStmt = SQLStmt & "E_MAIL = '" & CUSTOMER & "', "
	SQLStmt = SQLStmt & "CARD_TYPE = '" & Request("CARD_TYPE") & "', "
	SQLStmt = SQLStmt & "CARD_NO = '" & Request("CARD_NO") & "', "
	SQLStmt = SQLStmt & "CARD_EXP = '" & Request("CARD_EXP") & "', "
	SQLStmt = SQLStmt & "SUB_TOTAL = '" & FormatCurrency(RSSumOrd("SubTotal"),2) & "', "
	SQLStmt = SQLStmt & "TAX = '" & FormatCurrency(TAX,2) & "', "
	If Request("Shipping_STD") > 1 Then
	SQLStmt = SQLStmt & "SHIPPING_UPS = '" & FormatCurrency(Shipping,2) & "', "
	Else
	SQLStmt = SQLStmt & "SHIPPING_UPS = 0.00, "
	End If
	If Request("Shipping_PRM") > 1 Then
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
	SQLStmt = SQLStmt & " WHERE CUSTOMER_ID = " & ORDER_ID & ""
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

<head>
<title></title>
</head>

<body>

<table border="0" cellpadding="0" cellspacing="0" width="60%">
  <tr>
    <td width="50%" colspan="2"><p align="center"><img src="images/sm_gwn_logo.jpg" alt="sm_gwn_logo.jpg (3952 bytes)" WIDTH="183" HEIGHT="48"><font size="3"></p>
    </font><p align="center"><font size="3"><b><u>Order Confirmation</u></b><br>
    </font></td>
  </tr>
  <tr>
    <td width="50%" colspan="2"></td>
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
    <td width="50%" align="right"><% If Request("Shipping_AIR") = "" Then %>
<p>2nd Day Air Shipping:</td>
    <td width="50%" align="right"><%= FormatCurrency(Shipping,2) %>
</td>
    <td width="50%" align="right"><% ElseIf Request("Shipping_UPS") > 1 Then %>
<p>UPS Shipping:</td>
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
  <tr>
    <td><p align="center">&nbsp; <% Session.Abandon %> </td>
  </tr>
</table>
</body>
</html>
<!--#include file="confirm_foot.htm"-->
<%		
	Else
		
		FailPath = "ccerror.asp"
	Connection.Close	
		CC_Response = FailPath&"?CustomerTransactionNumber="&CustomerTransactionNumber&"&MerchantTransactionNumber="&MerchantTransactionNumber&"&AvsCode="&AvsCode&"&AuxMsg="&AuxMsg&"&ActionCode="&ActionCode&"&RetrievalCode="&RetrievalCode&"&AuthNum="&AuthNum&"&ErrorMsg="&ErrorMsg&"&ErrorLocation="&ErrorLocation
	
		'Response.Write CC_Response
		Response.Redirect CC_Response 
		
	End If

	Session.Abandon

%>
