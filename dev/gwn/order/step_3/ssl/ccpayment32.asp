<%@ LANGUAGE="VBScript" %>

<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront 99 v.3.0.1
'
'   Author     :   LaGarde, Incorporated * Some parts of this file are
'		   copyright CyberCash, Inc.
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
'   (c) Copyright 1998 CyberCash. All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<%
	DSN_Name = Request("DSN_Name")
	
	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""
		

	SQLStmt = "SELECT * FROM Admin"

	Set RSAdmin = Connection.Execute(SQLStmt)
	

	Cust_Name = Server.URLEncode(Replace(Request("CARD_NAME"),"'","''"))
	Company = Server.URLEncode(Replace(Request("COMPANY"),"'","''"))
	Address_1 = Server.URLEncode(Replace(Request("ADDRESS_1"),"'","''"))
	Address_2 = Server.URLEncode(Replace(Request("ADDRESS_2"),"'","''"))
	City = Server.URLEncode(Replace(Request("CITY"),"'","''"))
	State = Server.URLEncode(Request("STATE"))
	Zip = Server.URLEncode(Request("ZIP"))
	Country = Server.URLEncode(Request("COUNTRY"))
	CardName = Server.URLEncode(Replace(Request("CARD_NAME"),"'","''"))
	CardType = Server.URLEncode(Request("CARD_TYPE"))
	CardNo = Request("CARD_NO")
	CardExp = Request("CARD_EXP")
	Phone = Request("PHONE")
	Fax = Request("FAX")
	EMail = Server.URLEncode(Request("E_MAIL"))
	Address = Server.URLEncode(Address_2&","&Address_1)
	
	If Request("PAYMENT_METHOD") = "e_check" Then
		scriptPath = "/directpaycheck.asp?"
	ElseIf Request("PAYMENT_METHOD") = "credit_card" Then
		scriptPath = "/directpaycredit.asp?"
	ElseIf Request("PAYMENT_METHOD") = "purch_order" Then
		PaymentPath = "confirm.asp"
	ElseIf Request("PAYMENT_METHOD") = "phone_fax" Then
		PaymentPath = "confirm.asp"
		
	End If

	ShipName = Server.URLEncode(Replace(Request("SHIP_NAME"),"'","''"))
	ShipCompany = Server.URLEncode(Replace(Request("SHIP_COMPANY"),"'","''"))
	ShipAddress1 = Replace(Request("SHIP_ADDRESS_1"),"'","''")
	ShipAddress2 = Replace(Request("SHIP_ADDRESS_2"),"'","''")
	ShipCity = Server.URLEncode(Replace(Request("SHIP_CITY"),"'","''"))
	ShipState = Server.URLEncode(Request("SHIP_STATE"))
	ShipCountry = Server.URLEncode(Request("SHIP_COUNTRY"))
	ShipZip = Server.URLEncode(Request("SHIP_ZIP"))
	ShipTelephone = Server.URLEncode(Request("SHIP_TELEPHONE"))
	SHIP_MESSAGE = Server.URLEncode(Replace(Request("SHIP_MESSAGE"),"'","''"))
	ShipAddress = Server.URLEncode(ShipAddress1&"@"&ShipAddress2&"@"&ShipZip&"@"&ShipTelephone&"@"&SHIP_MESSAGE)
	
	Login = Server.URLEncode(Trim(RSAdmin("LOGIN")))
	OrderID = Request("ORDER_ID")
	Subject = Server.URLEncode(RSAdmin("EMAIL_SUBJECT"))
	TaxCountry = Server.URLEncode(RSAdmin("TAX_COUNTRY"))
	TaxState = Server.URLEncode(RSAdmin("TAX_STATE"))
	StateTaxAmt = RSAdmin("STATE_TAX_AMOUNT")
	CountryTaxAmt = RSAdmin("COUNTRY_TAX_AMOUNT")
	TransactionMethod = RSAdmin("Transaction_Method")
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
	SQLStmt = SQLStmt & "WHERE ORDERS.ORDER_ID = " & Request.QueryString("ORDER_ID") & " "
	SQLStmt = SQLStmt & "AND ORDERS.PRODUCT_ID = PRODUCT.PRODUCT_ID" 
	
	SET RSOrder = Connection.Execute(SQLStmt)

	If NOT RSOrder.EOF Then
	
	Set RSProduct = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT (Sum(TOTAL)) AS SubTotal "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & Request.QueryString("ORDER_ID") & " "

	Set RSSumOrd = Connection.Execute(SQLStmt)


	SQLStmt = "SELECT (Sum(TOTAL) * " & CountryTaxAmt & ") AS CountryTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & Request.QueryString("ORDER_ID") & " "
	Set RSCountryTax = Connection.Execute(SQLStmt)

	
	SQLStmt = "SELECT (Sum(TOTAL) * " & StateTaxAmt & ") AS StateTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & Request.QueryString("ORDER_ID") & " "

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
	End If

	If Request("Shipping_PRM") > 1 Then
		Shipping = cCur(SpShipping)
		ShipMethod = "PRM"
				
	ElseIf Request("Shipping_STD") > 1 Then
		Shipping = cCur(Shipping)
		ShipMethod = "STD"
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

	Amount = Grand_Total
	
	
	
%>



<!-- #include file="mck-cgi/CCMerchantTest.inc" -->
<!-- #include file="mck-cgi/CCMerchantCustom.inc" -->
<!-- #include virtual="/mck-shared/CCVarBlock.inc" -->
<!-- #include virtual="/mck-shared/CCMckLib.inc" -->

<%
function TestDriver ()
   On Error Resume Next
   Dim nStatus, sOrderId, sPrice, sOrderValidTill, sPayload, sOrderDescr   
   Dim sCustomerId, sFailureUrl, sCustomization, sScript, sDebugMsg

   Dim DictQuery
   Set DictQuery = CreateObject("Scripting.Dictionary")

   Dim DictToken
   Set DictToken = CreateObject("Scripting.Dictionary")

   nStatus = InitConfig()
   if (nStatus = nE_ERROR) then
      TestDriver = sHTMLPage
      exit function
   end if

   CCDebug (" ")
   CCDebug ("Entering TestDriver")

   Dim sTempDiffTemplateLoc
   sTempDiffTemplateLoc = gDictConfig.Item("TEMPLATE_DIR") _
                        & "tempDifficulties.htm"

   Call GetQuery(DictQuery)
   
   sOrderId = Request("Order_ID")
   if (sOrderId = "") then
      sOrderId = GenerateOrderId()
   end if
   DictToken.Add "#ORDERID#", sOrderId
   
   sPrice = Amount
   if (sPrice = "") then
      ' Fail this.  The driver is not supposed to supply one
      sDebugMsg = MCKGetErrorMessage(nE_No_Amount)
      Call CCLogError (sDebugMsg)
      DictToken.Add "#MESSAGE#", sDebugMsg
      nStatus = FormatTemplate(sTempDiffTemplateLoc, DictToken)
      TestDriver = sHTMLPage
      exit function
   end if
  
   ' payload is optional
   sPayload = Trim(DictQuery.Item("payload"))
 
   ' ovt is optional
   sOrderValidTill = Trim(DictQuery.Item("ovt"))
 
   ' order-descr is optional
   sOrderDescr = Trim(DictQuery.Item("order-descr"))
 
   ' customer-id is optional
   sCustomerId = Trim(DictQuery.Item("customer-id"))
 

   ' Script name is a radio button ... name of the HTML template to 
   ' load and hence, which script we will test.
   sScript = gDictConfig.Item("TEMPLATE_DIR") & Trim(DictQuery.Item("script"))

   ' At this point, record what you know about the order so that
   ' the payment script can retrieve the information
   ' record the orderId as a known one... (fail if you can't record it)

   nStatus  = SaveOrderInfo (sOrderId, sPrice, sOrderValidTill, sPayload, _
                             sOrderDescr, sCustomerId)

   if (nStatus = nE_ERROR) then
      sDebugMsg = MCKGetErrorMessage(nE_No_OrderLog)
      Call CCLogError (sDebugMsg)
      DictToken.Add "#MESSAGE#", sDebugMsg
      nStatus = FormatTemplate(sTempDiffTemplateLoc, DictToken)
      TestDriver = sHTMLPage
      exit function
   end if

   'At this point, all of the data that needs to go into the "Payment Page"
   ' should have been collected.
   '
   ' PrintTemplate will expand a payment template:
   '   asp_cardsale.tem
   '   asp_nocardsale.tem
   '   asp_mswcardale.tem
   '   asp_ccwcardsale.tem
   '   asp_checksale.tem
   '
   ' that will drive the CGI script that you wish to test....
   '
   '  These parameters must be expanded (the CGI will validate these):
   '   #ORDERID#
   '   #TEST_PRMS#

   nStatus =  nE_NoErr

   'Call CCDebug ("Exiting TestDriver")
   'TestDriver = sHTMLPage
end function 

Call TestDriver

	customerID = "customerID="&OrderID
	order_ID = "orderID="&OrderID
	ordermo = "mo.order-id="&OrderID
	name = "cpi.card-name="&Cust_Name
	address = "cpi.card-address="&Address
	city = "cpi.card-city="&city
	state = "cpi.card-state="&state
	zip = "cpi.card-zip="&zip
	If country = "United+States" Then
	country = "cpi.card-country=US"
	ElseIf country = "Canada" Then 
	country = "cpi.card-country=CA"
	Else
	country = "cpi.card-country="&country
	End If
	phone = "phone="&phone
	fax = "fax="&fax
	email = "email="&email
	cardtype = "cpi.card-type="&CardType
	cardnumber = "cpi.card-number="&CardNo
	cardexp = "cpi.card-exp="&CardExp
	DSN_Name = "DSN_Name="&DSN_Name
	PaymentMethod = Request("PAYMENT_METHOD")
	Description = Server.URLEncode(Request("SHIP_MESSAGE"))
	chkacctno = "cpi.check-account-number="&Request("CHK_ACCT_NO")
	routeno = "cpi.check-bank-routing-number="&Request("BANK_CODE")
	chkacctname = "cpi.check-name="&Request("CARD_NAME")
	chkuse = "cpi.check-use=n/a"
	acctpaid = "cpi.account-paid=n/a"
	PaymentPath = RTrim(RSAdmin("PAYMENT_SERVER"))
	
	Connection.Close
	processing_path = PaymentPath&scriptPath&customerID&"&"&order_ID&"&"&ordermo&"&"&name&"&"&address&"&"&city&"&"&state&"&"&country&"&"&zip&"&"&phone&"&"&fax&"&"&email&"&"&cardtype&"&"&cardnumber&"&"&cardexp&"&"&DSN_Name&"&USER3="&Company&"&USER4="&ShipName&"&USER5="&ShipCompany&"&USER6="&ShipAddress&"&USER7="&ShipMethod&"&USER8="&ShipCity&"&USER9="&ShipState&"&USER10="&ShipCountry&"&DESCRIPTION="&Description&"&ACCTNO="&ChkAcctNo&"&ABACODE="&BankCode&"&BANKNAME="&BankName&"&PAYMENT_METHOD="&PaymentMethod
	'Response.Write processing_path
	Response.Redirect processing_path


%>
