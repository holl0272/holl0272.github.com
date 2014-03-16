<%@ LANGUAGE="VBSCRIPT" %>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	DSN_Name = Request("DSN_Name")
	
	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""
		
	SQLStmt = "SELECT * FROM Admin"

	Set RSAdmin = Connection.Execute(SQLStmt)




	PAYMENT_METHOD = Request("Payment_Method")
	CARD_TYPE = Request("cpi.card-type")
	CARD_NAME = Request("cpi.card-name")
	CARD_NO = Request("cpi.card-number")
	CARD_EXP = Request("cpi.card-exp")
	ADDRESS = Request("cpi.card-address")
	CITY = Request("cpi.card-city")
	STATE = Server.URLEncode(Request("cpi.card-state"))
	ZIP = Request("cpi.card-zip")
	COUNTRY = Server.URLEncode(Request("cpi.card-country"))
	PHONE = Request("phone")
	FAX = Request("fax")
	CUSTOMER = Request("email")
	COMPANY = Request("USER3")
	ORDER_ID = Request("customerID")
	SHIP_NAME = Request("USER4")
	SHIP_COMPANY = Request("USER5")
	SHIP_ADDRESS = Request("USER6")
	SHIP_ADDRESS = Replace(SHIP_ADDRESS,"%40","@")
	
	SHIP_ADDRESS = Split(SHIP_ADDRESS,"@")
	SHIP_ADDRESS_2 = SHIP_ADDRESS(0)
	SHIP_ADDRESS_1 = SHIP_ADDRESS(1)
	SHIP_ZIP = SHIP_ADDRESS(2)
	SHIP_TELEPHONE = SHIP_ADDRESS(3)
	SHIP_MESSAGE = SHIP_ADDRESS(4)
	
	SHIP_CITY = Request("USER8")
	ShipState = Request("USER9")
	ShipCountry = Request("USER10")
	SHIP_METHOD = Request("USER7")
			
	SQLStmt = "SELECT GRAND_TOTAL from customer WHERE "
	SQLStmt = SQLStmt & "CUSTOMER_ID = " & ORDER_ID & ""
	RSOrderCheck = Connection.Execute(SQLStmt)
	If RSOrderCheck("GRAND_TOTAL") > 0 Then
	Response.Redirect "order_complete.htm"
	Else
	End If
	
	OrderID = Request("ORDERID")
	Subject = Server.URLEncode(RSAdmin("EMAIL_SUBJECT"))
	TaxCountry = Server.URLEncode(RSAdmin("TAX_COUNTRY"))
	TaxState = Server.URLEncode(RSAdmin("TAX_STATE"))
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
	SQLStmt = SQLStmt & "WHERE ORDERS.ORDER_ID = " & ORDERID & " "
	SQLStmt = SQLStmt & "AND ORDERS.PRODUCT_ID = PRODUCT.PRODUCT_ID" 
	
	SET RSOrder = Connection.Execute(SQLStmt)

	If NOT RSOrder.EOF Then
	
	Set RSProduct = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT (Sum(TOTAL)) AS SubTotal "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDERID & " "

	Set RSSumOrd = Connection.Execute(SQLStmt)


	SQLStmt = "SELECT (Sum(TOTAL) * " & CountryTaxAmt & ") AS CountryTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDERID & " "
	Set RSCountryTax = Connection.Execute(SQLStmt)

	SQLStmt = "SELECT (Sum(TOTAL) * " & StateTaxAmt & ") AS StateTax "
	SQLStmt = SQLStmt & "FROM ORDERS WHERE "
	SQLStmt = SQLStmt & " ORDER_ID = " & ORDERID & " "

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

	If SHIP_METHOD = "PRM" Then
		Shipping = cCur(SpShipping)
				
	ElseIf SHIP_METHOD = "STD" Then
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
	Amount = CCur(Grand_Total)
	
%>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">

<script LANGUAGE="vbscript" RUNAT="server">

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   Merchant Connection Kit (MCK)
'
'   Author     :   CYCH Development staff
'   Description:   CGI script to do SSL submission of a Credit Card payment
'                  mediated by a customer's browser
'
'   Notes      :   This may be edited to integrate with the merchant's
'                  storefront
'
'                         COPYRIGHT NOTICE
'
'   The contents of this file is protected under the United States
'   copyright laws as an unpublished work, and is confidential and
'   proprietary to CyberCash, Inc.  Its use or disclosure in whole or in
'   part without the expressed written permission of CyberCash, Inc. is
'   prohibited.
'
'   (c) Copyright 1993-97 by CyberCash, Inc.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' directpaycredit.asp
'   This script is launched when a consumer decides to buy something.
'   This services credit sales made with or without use of a wallet.
'   This script must be under SSL access ... e.g. do not send
'   card data over the web in the clear.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The "mck-shared" is a virtual directory which is set in your registry during
' configuration.  It's pointed to where the include files for Storefront
' ASPs reside.
'
' It's set in the HKEY_LOCAL_MACHINE->SYSTEM->CurrentControlSet->Services->
' W3SVC->Parameters->Virtual Roots.
'
<!-- #include file="CCMerchantTest.inc" -->
<!-- #include file="CCMerchantCustom.inc" -->
<!-- #include virtual="mck-shared/CCVarBlock.inc" -->
<!-- #include virtual="mck-shared/CCMckLib.inc" -->
<!-- #include virtual="mck-shared/CCMckDirectLib.inc" -->

function BuildDirectPayCredit ()
   On Error Resume Next

   Dim sProblemTemplate, sSuccessTemplate, sFailTemplate

   Dim DictPayment
   Set DictPayment = CreateObject("Scripting.Dictionary")

   Dim DictQuery
   Set DictQuery = CreateObject("Scripting.Dictionary")

   Dim DictToken
   Set DictToken = CreateObject("Scripting.Dictionary")

   Dim DictPOP
   Set DictPOP = CreateObject("Scripting.Dictionary")

   Dim nStatus, sPaymentURL, sSignature, sPOP, sFileName 
   Dim sStatus, sDebugMsg, sCode, nLogStatus, nPayloadStatus 
   Dim arDictKey, cCount, sName

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Some data initialization
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   nStatus = InitConfig()
   if (nStatus = nE_ERROR) then
      BuildDirectPayCredit = sHTMLPage  
      exit function
   end if

   CCDebug (" ")
   CCDebug ("Entering BuildDirectPayCredit")
   
   ' Set the file path for temporary difficulty template
   Dim sTempDiffTemplateLoc
   sTempDiffTemplateLoc = gDictConfig.Item("TEMPLATE_DIR") _
                        & "tempDifficulties.htm"
   
   ' A Note about message fields in paymentNVList.
   '    paymentNVList holds message parameters
   '    that will be passed to the Cash Register
   '
   ' These parameters are broken into three blocks according to
   ' a prefix in the argument field name.
   '
   ' prefix                 use
   '
   '  mo.        -- data driving the Cash Register
   '                 this include Payment data and some operational options
   '  cpi.       -- info about the Customer's "Payment Instrument: (card)
   '  mf.        -- anything additional you want to pass on to the
   '                  Fulfillment Center
   DictPayment.Add "mo.cybercash-id", gDictConfig.Item("CYBERCASH_ID")
   DictPayment.Add "mo.version", sMCKversion

   ' In direct connect mode, the MO and CPI blocks are not signed, since
   ' the message will be encrypted, which is a stonger insurance against
   ' tampering than the signature.. and because all of the relevant
   ' information comes back here anyway.
   DictPayment.Add "mo.signed-cpi", "no"

   sProblemTemplate = gDictConfig.Item("TEMPLATE_DIR") & "tempDifficulties.htm"

   ' In direct connect mode, you can generate a receipt or
   ' dispense a payload directly once the Cash Register responds,
   '
   ' or, you can redirect the response to a remote Fulfillment center
   ' and allow that Fulfillment Center to generate a receipt or dispense
   ' the payload.
   '
   ' Indicate a redirection by including the field:
   '   "mo.redirect-url"
   ' in your request.
   '
   ' If you redirect, we will use the customRedirectResponse.htm
   ' template to generate a redirection page.
   '
   ' You may use customReceipt.tem for a locally generated receipt
   ' and customFailureResponse.tem if the payment fails somehow.
   '
   ' you may customize these pages as you wish.

   ' If you are passing anything in the MF section of the message
   ' to a fulfillment center via a Redirect,
   ' You had better be an SSL connection (https://) in the redirectURL.
   if (Len(gDictConfig.Item("REDIRECT_URL")) > 0) then
      DictPayment.Add "mo.redirect-url", gDictConfig.Item("REDIRECT_URL") 
      DictToken.Add "#mo.redirect-url#", gDictConfig.Item("REDIRECT_URL") 
   end if

   ' For "Direct Connect" payments, the custom*.tem templates are
   ' *ALWAYS* used.
   if (Len(gDictConfig.Item("REDIRECT_URL")) > 0) then
      sSuccessTemplate = gDictConfig.Item("TEMPLATE_DIR") & "customRedirectResponse.htm"
   else
      sSuccessTemplate = gDictConfig.Item("TEMPLATE_DIR") & "customReceipt.htm"
   end if

   sFailTemplate = gDictConfig.Item("TEMPLATE_DIR") & "customFailureResponse.htm"

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Main code starts here 
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   ' Get the query string 
   Call GetQuery(DictQuery)

   ' Read the order-id from the message
	sOrderid = request("mo.order-id")
   'sOrderid = DictQuery.Item("mo.order-id")

   ' Payment pages feed into this, so the MO may be presnt as a MO Block.
   ' If it is, unpack it here...
   if (DictQuery.Exists("MO")) then
      ' Unpack the MO block and load it into the query vector
      ' Note:  You, the merchant, may sign such blocks.
      ' If you do, verify the signature here....
      Dim DictMO, sMO 
      Set DictMO = CreateObject("Scripting.Dictionary")

      sMO = DictQuery.Item("MO")
      DictQuery.Remove("MO")
      Call URLdecodeForm(sMO, DictMO)

      ' Read the order-id from the message
	sOrderid = Request("mo.order-id")
     'sOrderid = DictMO.Item("mo.order-id")
 
      if (DictMO.Exists("mo.sign")) then
         sSign = DictMO.Item("mo.sign")
         DictMO.Remove("mo.sign")       'Don't sign the sign

         sSignature = BuildSignature("mo", gDictConfig.Item("HASH_SECRET"), DictMO)

         if (sSign <> sSignature) then
            ' Gone!
            sDebugMsg = MCKGetErrorMessage(nE_MO_Signature)
            DictToken.Add "#MESSAGE#", sDebugMsg
            DictToken.Add "#ORDERID#", sOrderid
            Call CCLogError (sDebugMsg)
            nStatus = FormatTemplate(sProblemTemplateLoc, DictToken)
            BuildDirectPayCheck = sHTMLPage
            exit function
         end if

         arDictKey = DictMO.keys
         for cCount = 0 to DictMO.Count - 1
           sName = arDictKey(cCount)
           DictQuery.Add sName, DictMO.item(sName)
         Next  
      end if
   end if

   ' Load Payment MO information from query into the payment args list
   ' Key is the order id.  This function notifies and logs and exits
   ' if there is problem
   nStatus = LoadPaymentInfo(sOrderid, DictQuery, DictPayment, sProblemTemplate)
   if (nStatus = nE_ERROR) then
      BuildDirectPayCredit = sHTMLPage
      exit function
   end if

   ' The following tokens are in customReceipt.tem and customFailureResponse.tem
   ' To initialize here so they will be at the top of Token dictionary object.
   DictToken.Add "#pop.order-id#", request("OrderID")'sOrderid
   DictToken.Add "#mo.order-id#", request("mo.order-id")'sOrderid
   DictToken.Add "#mo.price#", Cint(GrandTotal)'100'DictPayment.Item("mo.price")

   ' load credit card info from form. 
   ' DON'T TOUCH THIS! This is the CR script that will process your payment.
   sPaymentURL = gDictConfig.Item("CCPS_HOST") & "directcardpayment.cgi"

   'DictPayment.Add "cpi.card-number",  Request("opi.card-number")'DictQuery.Item("cpi.card-number") 
   'DictPayment.Add "cpi.card-exp",     Request("cpi.card-exp")'DictQuery.Item("cpi.card-exp") 
   'DictPayment.Add "cpi.card-name",    Request("cpi.card-name")'DictQuery.Item("cpi.card-name") 
   'DictPayment.Add "cpi.card-address", Request("cpi.card-address")'DictQuery.Item("cpi.card-address") 
   'DictPayment.Add "cpi.card-city",    Request("cpi.card-city")'DictQuery.Item("cpi.card-city") 
   'DictPayment.Add "cpi.card-state",   Request("cpi.card-state")'DictQuery.Item("cpi.card-state") 
   'DictPayment.Add "cpi.card-zip",     Request("cpi.card-zip")'DictQuery.Item("cpi.card-zip") 
   'DictPayment.Add "cpi.card-country", Request("cpi.card-country")'DictQuery.Item("cpi.card-country") 

   DictPayment.Add "cpi.card-number",  DictQuery.Item("cpi.card-number") 
   DictPayment.Add "cpi.card-exp",     DictQuery.Item("cpi.card-exp") 
   DictPayment.Add "cpi.card-name",    DictQuery.Item("cpi.card-name") 
   DictPayment.Add "cpi.card-address", DictQuery.Item("cpi.card-address") 
   DictPayment.Add "cpi.card-city",    DictQuery.Item("cpi.card-city") 
   DictPayment.Add "cpi.card-state",   DictQuery.Item("cpi.card-state") 
   DictPayment.Add "cpi.card-zip",     DictQuery.Item("cpi.card-zip") 
   DictPayment.Add "cpi.card-country", DictQuery.Item("cpi.card-country") 


   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' You shouldn't need to change anything else below this
   ' unless you want to do some bookeeping at the end
   ' 
   ' Or you may add MF parameters here as well.
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' No signatures required for direct connect...
   ' NO SSL required to talk to the Cash Register either...
   ' Since the whole thing is encrypted, you can't mess with it
   ' and you can't read it!
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   ' If a payment page is driving this, you may have a signed MF blob here...
   ' Unpack it and check the signature....
   if (DictQuery.Exists("MF")) then
      ' Unpack the MF block and load it into the query vector
      ' Note:  You, the merchant, may sign such blocks.
      ' If you do, verify the signature here....
      Dim DictMF, sMF 
      Set DictMF = CreateObject("Scripting.Dictionary")

      sMF = DictQuery.Item("MF")
      DictQuery.Remove("MF")
      Call URLdecodeForm(sMF, DictMF)
    
      if (DictMF.Exists("mf.sign")) then
         sSign = DictMF.Item("mf.sign")
         DictMF.Remove("mf.sign")       'Don't sign the sign

         sSignature = BuildSignature("mf", gDictConfig.Item("HASH_SECRET"), DictMF)

         if (sSign <> sSignature) then
            ' Gone!
            sDebugMsg = MCKGetErrorMessage(nE_MF_Signature)
            DictToken.Add "#MESSAGE#", sDebugMsg
            DictToken.Add "#ORDERID#", sOrderid
            Call CCLogError (sDebugMsg)
            nStatus = FormatTemplate(sProblemTemplateLoc, DictToken)
            BuildDirectPayCheck = sHTMLPage
            exit function
         end if

         arDictKey = DictMF.keys
         for cCount = 0 to DictMF.Count - 1
           sName = arDictKey(cCount)
           DictQuery.Add sName, DictMF.item(sName)
         Next  
      end if
   end if

   ' Load any MF fields that may be in the query
   ' This will copy any query field with "mf." prefix into the payment vector
   Call BuildBlock("mf", DictQuery, DictPayment)

   ' Add anything additional you like to this message, we will pass it on to the FC!
   'DictPayment.Add "mf.your-field-here",  "some value"

   ' if you want to sign it, go ahead...
   ' if you want to do it some other way, that's also OK.  The cash
   ' register will pass this along, but it will not check it.
   '
   ' Adding a signature only makes sense if 
   '   1) You are redirecting to a Fulfillmetn Center
   '   2) You can check the signature at the Fulfillment Center

   if (Len(gDictConfig.Item("REDIRECT_URL")) > 0) then
      DictPayment.Add "mf.sign", BuildSignature("mf", _
                                 gDictConfig.Item("HASH_SECRET"), DictPayment) 
   end if

   '''''''''''''''''''''''''''
   ' Generate a Direct Payment
   '''''''''''''''''''''''''''

   Call doDirectPayment(sPaymentURL, DictPayment, DictPOP, DictToken)

   ' At this point, we do fulfillment ....
   ' Make absolutely everything available for template building
   '
   ' If you have anything else you need to record, like
   ' shipping address, etc.  You may do that here ...

   ' Your code goes here ...

   ' Log the result here ...
 
   sPOP = DictToken.Item("POP")
   Call LogNotification(sPOP, nLogStatus, sCode)

   ' What you "DO* about a logging failure depends on a number of things.
   if (nLogStatus <> nE_NoErr) then
      ' Log the failure in any case ...
      sDebugMsg = "Order ID " & sOrderid & MCKGetErrorMessage(nE_Fail_Notif) _
                & " - " & CStr(nLogStatus)       
      Call CCLogError (sDebugMsg)

      if (DictPOP.Exists("pop.payload")) then
         ' If logging failed, hard goods won't be delivered. This is fatal.
         DictToken.Add "#MESSAGE#", sDebugMsg
         DictToken.Add "#ORDERID#", sOrderid
         nStatus = FormatTemplate(sProblemTemplateLoc, DictToken)
         BuildDirectPayCheck = sHTMLPage
         exit function
      end if
   end if

   ' There are two possible ways to do fulfillment:
   ' 1) do it right here
   ' 2) redirect this response to a remote fulfillment location 
   ' Let's start with remote fulfillment ...  
   sStatus = DictPOP.Item("pop.status")

   if (Instr(1, sStatus, "success") = 0) then
      ' Bombed out ... Response using the customFailureResponse page
      Call FormatTemplate(sFailTemplate, DictToken)
   else
      if (Len(gDictConfig.Item("REDIRECT_URL")) > 0) then
         ' Use customRedirectResponse.tem to redirect browser to
         ' fulfillment site ...
         ' The redirect URL has a 2nd chance to log this, and
         ' since that is where fulfillment happends, this may end well...
         Call FormatTemplate(sSuccessTemplate, DictToken)
      else
         ' Here we generate a receipt or directly dispense a payload
        
         ' For a receipt, note that if the logging failed, we will lose
         ' the record of this sale ... This is a major problem. 

         if (DictPOP.Exists("pop.payload")) then
            sFileName = DictPOP.Item("pop.payload")

            nPayloadStatus = DispensePayload(sFileName)
            if (nPayloadStatus = nE_ERROR) then
               BuildDirectPayCredit = sHTMLPage
		exit function
            end if
         else
            Call FormatTemplate(sSuccessTemplate, DictToken)
		


<!-- #include file = "../confirm_head.htm" -->
Confirm = "<html><body><div align=center><center>"
Confirm = Confirm & "<h3><b>Order Confirmation</b>"
Confirm = Confirm & "<table border=0 cellpadding=0 cellspacing=0 width=60%>"
Confirm = Confirm & "<tr><td width=50% colspan=2>&nbsp;</td></tr><tr>"
Confirm = Confirm & "<td width=50% align=right>Order Number:</td>"
Confirm = Confirm & "<td width=50% align=right>&nbsp;"&ORDER_ID&"</td></tr>"
Confirm = Confirm & "<tr><td width=50% align=right>Date: </td>"
Confirm = Confirm & "<td width=50% align=right>&nbsp;"&Date()&"</td></tr>"
Confirm = Confirm & "<tr><td width=50% align=right>Order Amount: </td>"
Confirm = Confirm & "<td width=50% align=right>&nbsp;"&FormatCurrency(SubTotal,2)&"</td>"
Confirm = Confirm & "</tr><tr>"
Confirm = Confirm & "<td width=50% align=right>"
If SHIP_METHOD = "PRM" Then
Confirm = Confirm & "Premium Shipping:</td>"
Confirm = Confirm & "<td width=50% align=right>"&FormatCurrency(Shipping,2)&"</td>"
Confirm = Confirm & "<td width=50% align=right> "
ElseIf SHIP_METHOD = "STD" Then
Confirm = Confirm & " Standard Shipping:</td>"
Confirm = Confirm & "<td width=50% align=right>"&FormatCurrency(Shipping,2)&"</td>"
End If
If Tax > 0 Then
Confirm = Confirm & "</tr><tr><td width=50% align=right>Tax:&nbsp; </td>"
End If
Confirm = Confirm & "<td width=50% align=right>&nbsp;"&FormatCurrency(TAX,2)&"</td>"
Confirm = Confirm & "</tr><tr><td width=50% align=right>Total Amount:</td>"
Confirm = Confirm & "<td width=50% align=right>"&FormatCurrency(GRAND_TOTAL,2)&"</td>"
Confirm = Confirm & "</tr></font></table></center></div></body></html>"

Response.Write Confirm


         end if
      end if
   end if   

   Call CCDebug ("Exiting BuildDirectPayCredit")
   BuildDirectPayCredit = sHTMLPage


end function 

</script> 
<%
'*****************************************************

	
	If  Request("CustomerTransactionNumber") = "" Then
	CustomerTransactionNumber = "N/A"
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
	SetRSCC = Connection.Execute(SQLStmt)

	SQLStmt = "UPDATE CUSTOMER SET NAME = '" & CARD_NAME & "', "
	SQLStmt = SQLStmt & "COMPANY = '" & COMPANY & "', "
	SQLStmt = SQLStmt & "ADDRESS_1 = '" & ADDRESS & "', "
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
	If Ship_Method = "STD" Then
	SQLStmt = SQLStmt & "SHIPPING_UPS = '" & FormatCurrency(Shipping,2) & "', "
	Else
	SQLStmt = SQLStmt & "SHIPPING_UPS = 0.00, "
	End If
	If Ship_Method = "PRM" Then
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
	SQLStmt = SQLStmt & "SHIP_STATE = '" & ShipState & "', "
	SQLStmt = SQLStmt & "SHIP_ZIP = '" & SHIP_ZIP & "', "
	SQLStmt = SQLStmt & "SHIP_COUNTRY = '" & ShipCountry & "', "
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
<!--#include file="../mail.inc"-->

<%= BuildDirectPayCredit %>

