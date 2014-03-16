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
	
	DSN_Name = Request("DSN_Name")
	
	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""
	
	SQLStmt = "SELECT * FROM Admin"

	Set RSAdmin = Connection.Execute(SQLStmt)

	MerchantType = Trim(RSAdmin("MERCHANT_TYPE"))


	If MerchantType = "normal_auth" Then MType = "NA" End If
	If MerchantType = "authonly" Then MType = "AO" End If
	If MerchantType = "authcapture" Then MType = "PA" End If
	If MerchantType = "Credit" Then	MType = "CR" End If


	'Response.Write MType

	CustName = Server.URLEncode(Replace(Request("CUST_NAME"),"'","''"))
	Company = Server.URLEncode(Replace(Request("COMPANY"),"'","''"))
	Address_1 = Server.URLEncode(Replace(Request("ADDRESS_1"),"'","''"))
	Address_2 = Server.URLEncode(Replace(Request("ADDRESS_2"),"'","''"))
	City = Server.URLEncode(Replace(Request("CITY"),"'","''"))
	State = Server.URLEncode(Request("STATE"))
	Zip = Server.URLEncode(Request("ZIP"))
	Country = Server.URLEncode(Request("COUNTRY"))
	CardName = Server.URLEncode(Replace(Request("CARD_NAME"),"'","''"))
	Phone = Server.URLEncode(Request("PHONE"))
	Fax = Server.URLEncode(Request("FAX"))
	EMail = Server.URLEncode(Request("E_MAIL"))
	Address = Address_2&"/"&Address_1
	If Request("Shipping_STD") = "" Then
	ShipMethod = "PRM"
	Else
	ShipMethod = "STD"
	End If
	If Request("PAYMENT_METHOD") = "e_check" Then
		CardType = "ACH"
		BankCode = Server.URLEncode(Request("BANK_CODE"))
		ChkAcctNo = Server.URLEncode(Request("CHK_ACCT_NO"))
		BankName = Server.URLEncode(Request("BANK_NAME"))
		PaymentPath = RTrim(RSAdmin("PAYMENT_SERVER"))
	ElseIf Request("PAYMENT_METHOD") = "credit_card" Then
		CardType = Request("CARD_TYPE")
		CardNo= Trim(Request("CARD_NO"))
		CardExp = Trim(Request("CARD_EXP"))
		PaymentPath = RTrim(RSAdmin("PAYMENT_SERVER"))
	ElseIf Request("PAYMENT_METHOD") = "purch_order" Then
		CardType = "purch_order"
		PaymentPath = "confirm.asp"
		RESPONSECODE = Request("PURCH_ORDER_NO")
	ElseIf Request("PAYMENT_METHOD") = "phone_fax" Then
		CardType = "phone_fax"
		PaymentPath = "confirm.asp"
		
	End If


	ShipName = Server.URLEncode(Replace(Request("SHIP_NAME"),"'","''"))
	ShipCompany = Server.URLEncode(Replace(Request("SHIP_COMPANY"),"'","''"))
	ShipAddress1 = Replace(Request("SHIP_ADDRESS_1"),"'","''")
	ShipAddress2 = Replace(Request("SHIP_ADDRESS_2"),"'","''")
	ShipCity = Server.URLEncode(Replace(Request("SHIP_CITY"),"'","''"))
	ShipState = Server.URLEncode(Request("SHIP_STATE"))
	ShipCountry = Server.URLEncode(Request("SHIP_COUNTRY"))
	ShipMessage = Server.URLEncode(Replace(Request("SHIP_MESSAGE"),"'","''"))
	ShipZip = Server.URLEncode(Request("SHIP_ZIP"))
	ShipPhone = Server.URLEncode(Request("SHIP_TELEPHONE"))
	ShipAddress = Server.URLEncode(ShipAddress1&"@"&ShipAddress2&"@"&ShipZip&"@"&ShipPhone&"@"&ShipMessage)
	
	Login = Server.URLEncode(Trim(RSAdmin("LOGIN")))
	OrderID = Request("ORDER_ID")
	Description = Server.URLEncode(Request("SHIP_MESSAGE"))
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
	SQLStmt = SQLStmt & "WHERE ORDERS.ORDER_ID = " & OrderID & " "
	SQLStmt = SQLStmt & "AND ORDERS.PRODUCT_ID = PRODUCT.PRODUCT_ID" 
	
	SET RSOrder = Connection.Execute(SQLStmt)

	If NOT RSOrder.EOF Then
	
	Set RSProduct = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT (Sum(TOTAL)) AS SubTotal "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & OrderID & " "

	Set RSSumOrd = Connection.Execute(SQLStmt)


	SQLStmt = "SELECT (Sum(TOTAL) * " & CountryTaxAmt & ") AS CountryTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & OrderID & " "
	Set RSCountryTax = Connection.Execute(SQLStmt)

	
	SQLStmt = "SELECT (Sum(TOTAL) * " & StateTaxAmt & ") AS StateTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & OrderID & " "

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


	If ShipMethod = "PRM" Then
		Shipping = cCur(SpShipping)
				
	ElseIf ShipMethod = "STD" Then
		Shipping = cCur(Shipping)
	End If
	
	SubTotal = RSSumOrd("SubTotal")
	'response.write "ShipCountry: "&ShipCountry&"ShipState: "&ShipState&"State: "&State&"Country: "&Country&"TaxState: "&TaxState&"TaxCountry: "&TaxCountry
	
	If ShipCountry = TaxCountry AND ShipState = TaxState Then
		Tax = ((RSStateTax("StateTax"))+(RSCountryTax("CountryTax")))
		Grand_Total = ((SubTotal)+(CCur(Tax))+(CCur(Shipping)))
		
	ElseIf ShipCountry = TaxCountry Then
		Tax = RSCountryTax("CountryTax")
		Grand_Total = ((SubTotal)+(CCur(Tax))+(CCur(Shipping)))
		
	ElseIf Country = TaxCountry AND State = TaxState Then
		Tax = ((RSStateTax("StateTax"))+(RSCountryTax("CountryTax")))
		Grand_Total = ((SubTotal)+(CCur(Tax))+(CCur(Shipping)))
		
	ElseIf Country = TaxCountry Then
		Tax = RSCountryTax("CountryTax")
		Grand_Total = ((SubTotal)+(CCur(Tax))+(CCur(Shipping)))
		
	Else
		Tax = "0.00"
		Grand_Total = (SubTotal+Shipping)
		
	End If

	Amount = FormatCurrency(Grand_Total,2)
	
	Connection.Close
	SndPayment = PaymentPath&"?RESPONSECODE="&RESPONSECODE&"&LOGIN="&Login&"&USER1="&DSN_Name&"&USER2="&Address_1&"&USER3="&Company&"&USER4="&ShipName&"&USER5="&ShipCompany&"&USER6="&ShipAddress&"&USER7="&ShipMethod&"&USER8="&ShipCity&"&USER9="&ShipState&"&USER10="&ShipCountry&"&CUSTID="&CustID&"&Invoice="&OrderID&"&DESCRIPTION="&Description&"&TYPE="&MType&"&AMOUNT="&Amount&"&ACCTNO="&ChkAcctNo&"&ABACODE="&BankCode&"&BANKNAME="&BankName&"&METHOD="&CardType&"&CARDNUM="&CardNo&"&EXPDATE="&CardExp&"&NAME="&CardName&"&ADDRESS="&Address_2&"&CITY="&City&"&STATE="&State&"&ZIP="&Zip&"&COUNTRY="&Country&"&PHONE="&Phone&"&FAX="&Fax&"&EMAIL="&EMail&"&ECHODATA=TRUE&DISABLERECEIPT=TRUE"
	'Response.Write SndPayment
	Response.Redirect SndPayment

%>