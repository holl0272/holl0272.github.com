<% 
Function sendOrderConfirmationEmails(byVal sCustEmail)

'	Call createMail("Confirm", sCustEmail)
	
Dim i, j
Dim paryAttributes
Dim paryOrderItem
Dim pstrBody
Dim pstrCustomerEmail
Dim pstrMerchantEmail
Dim pstrSubject

	With mclsCustomer
		'Build basic Email Body Info
		pstrBody = VbCrLf & "-----------------" & VbCrLf & "Sold To" & VbCrLf & "-----------------" & VbCrLf
		pstrBody = pstrBody & "" & .DisplayName & VbCrLf
		If Len(.custCompany) > 0 Then pstrBody = pstrBody & .custCompany & VbCrLf
		pstrBody = pstrBody & .custAddr1 & VbCrLf
		If Len(.custAddr2) > 0 Then pstrBody = pstrBody & .custAddr2 & VbCrLf
		pstrBody = pstrBody & .custCity & ", " & .custState & " " & .custZip & VbCrLf
		pstrBody = pstrBody & .countryName & VbCrLf
		pstrBody = pstrBody & .custPhone & VbCrLf
		If Len(.custFax) > 0 Then pstrBody = pstrBody & "Fax Number: " & .custFax & VbCrLf
		pstrBody = pstrBody & "Email Address: " & .custEmail & VbCrLf
	End	With	'mclsCustomer
		
	pstrBody = pstrBody & "Payment Method: " & sPaymentMethod & VbCrLf
		
	With mclsCustomerShipAddress
		pstrBody = pstrBody & VbCrLf & "-----------------" & VbCrLf & "Shipped To" & VbCrLf & "-----------------" & VbCrLf
		pstrBody = pstrBody & .DisplayName & VbCrLf
		If Len(.Company) > 0 Then pstrBody = pstrBody & .Company & vbCrLf
		pstrBody = pstrBody & .Addr1 & VbCrLf
		If Len(.Addr2) > 0 Then pstrBody = pstrBody & .Addr2 & VbCrLf
		pstrBody = pstrBody & .City & ", " & .State & " " & .Zip & VbCrLf
		pstrBody = pstrBody & .countryName & VbCrLf
		pstrBody = pstrBody & .Phone & VbCrLf
		If Len(.Fax) > 0 Then pstrBody = pstrBody & .Fax & VbCrLf
		If Len(.Email) > 0 Then pstrBody = pstrBody & .Email & VbCrLf
	End	With	'mclsCustomerShipAddress
		
	pstrBody = pstrBody & VbCrLf & "-----------------" & VbCrLf & "Purchase Summary" & VbCrLf & "-----------------" & VbCrLf
	pstrBody = pstrBody & "Order ID: " & mclsCartTotal.OrderID & VbCrLf
	
	For i = 0 To mclsCartTotal.UniqueOrderItemCount
		paryOrderItem = mclsCartTotal.OrderItem(i)
		If isArray(paryOrderItem) Then
			pstrBody = pstrBody & VbCrLf & "Item " & i + 1 & VbCrLf _
								& "Product ID: " & paryOrderItem(enOrderItem_prodID) & VbCrLf _
								& "Product Name: " & paryOrderItem(enOrderItem_prodName) & VbCrLf

			'Now for the attributes
			If paryOrderItem(enOrderItem_AttributeCount) > 0 Then
				paryAttributes = paryOrderItem(enOrderItem_AttributeArray)
				For j = 0 To paryOrderItem(enOrderItem_AttributeCount) - 1
					pstrBody = pstrBody & "  " & paryAttributes(j)(enAttributeItem_attrName) & ": " & paryAttributes(j)(enAttributeItem_attrdtName) & VbCrLf
				Next 'j
			End If	'isArray(paryOrderItemAttributes)
			
			pstrBody = pstrBody & "Product Price: " & FormatCurrency(paryOrderItem(enOrderItem_UnitPrice)) & VbCrLf _
								& "Quantity: " & paryOrderItem(enOrderItem_odrdttmpQuantity) & VbCrLf
								
			If paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) > 0 Then pstrBody = pstrBody & "BackOrdered Qty: " & paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) & " (of above qty)" & VbCrLf

			If paryOrderItem(enOrderItem_odrdttmpBackOrderQTY) > 0 Then
				pstrBody = pstrBody & "Gift Wrap Price: " & FormatCurrency(paryOrderItem(enOrderItem_gwPrice)) & " (of above qty)" & VbCrLf _
									& "Gift Wrap Qty: " & paryOrderItem(enOrderItem_odrdttmpGiftWrapQTY) & VbCrLf
			End If
		End If	'isArray(paryOrderItem)
	Next 'i

	Call setEmail_PromotionManager(pstrBody, maryDiscountSummary) 'this displays free products - see ssmodDiscounts.asp
	
	With mclsCartTotal

		If .Discount = 0 Then
			pstrBody = pstrBody & VbCrLf _
					& "SubTotal:    " & FormatCurrency(.SubTotalWithDiscount) & VbCrLf
		Else
			pstrBody = pstrBody & VbCrLf _
					& "SubTotal:    " & FormatCurrency(.SubTotal) & VbCrLf _
					& "Discount:    " & FormatCurrency(-1 * .Discount) & VbCrLf _
					& "SubTotal:    " & FormatCurrency(.SubTotalWithDiscount) & VbCrLf _
		End If

		pstrBody = pstrBody & "Shipping:    " & FormatCurrency(.Shipping) & "(" & .ShipMethodName & ")" & VbCrLf
			
		If .Handling > 0 Then pstrBody = pstrBody & "Handling:    " & FormatCurrency(.Handling) & VbCrLf
			  
		If sPaymentMethod = cstrCODTerm Then
			If CDbl(.COD) > 0 Then pstrBody = pstrBody & VbCrLf & "COD AMOUNT:  " & FormatCurrency(.COD) & VbCrLf
		End If

		pstrBody =	pstrBody & "State Tax:   " & FormatCurrency(.StateTax) & VbCrLf
		
		If .CountryTax > 0 Then pstrBody = pstrBody & "Country Tax: " & FormatCurrency(.CountryTax) & VbCrLf
		
		pstrBody =	pstrBody & "Grand Total: " & FormatCurrency(.CartTotal)

		Call ssGiftCertificate_adjustMail(pstrBody)	'see ssGiftCertificateRegister_common.asp
		
		If  Session("SpecialBilling") <> 0 then
			pstrBody = pstrBody & vbcrlf & "Billed Amount:" & FormatCurrency(Session("BillAmount"))
			pstrBody = pstrBody & vbcrlf & "Remaining Amount:" & FormatCurrency(Session("BackOrderAmount"))
		End if
	End With
		
 	'If adminShipType = 2 Then pstrBody = pstrBody & vbcrlf & "NOTICE: The shipping costs as shown above do not necessarily represent the carrier's published rates and may include additional charges levied by the merchant."

	pstrBody = pstrBody & VbCrLf & VbCrLf & "Special Instructions: " &  sShipInstructions
	
	pstrBody = pstrBody & VbCrLf & VbCrLf & WriteCustomFormValuesToEmail
	
	'Create Merchant Email Body
	If Len(cstrCCVFieldName) > 0 And Not cstrCCV_SaveToDB Then 
		pstrMerchantEmail = pstrBody & VbCrLf _
						  & "CCV: " & mstrPayCardCCV & vbcrlf _
						  & "Retrieve Order:" & VbCrLf _
						  & "--" & adminDomainName & "ssl/admin/sfReports1.asp?OrderID=" & mclsCartTotal.OrderID & VbCrLf _
						  & "--" & adminDomainName & "ssl/ssAdmin/ssOrderAdmin.asp?Action=ViewOrder&OrderID=" & mclsCartTotal.OrderID & VbCrLf
	Else
		pstrMerchantEmail = pstrBody & VbCrLf _
						  & "Retrieve Order:" & VbCrLf _
						  & "--" & adminDomainName & "ssl/admin/sfReports1.asp?OrderID=" & mclsCartTotal.OrderID & VbCrLf _
						  & "--" & adminDomainName & "ssl/ssAdmin/ssOrderAdmin.asp?Action=ViewOrder&OrderID=" & mclsCartTotal.OrderID & VbCrLf
	End If
	
	'Customer Email Body
	If cblnDisableLogin Then
		pstrCustomerEmail = adminEmailMessage & vbcrlf & pstrBody
	Else 
		pstrCustomerEmail = adminEmailMessage & vbcrlf & pstrBody & VbCrLf & "User Name: " & mclsCustomer.custEmail & VbCrLf & "Password: " & mclsCustomer.custPasswd & VbCrLf & "Use this information next time you order for quick access to your Customer Information."
	End If 
	pstrCustomerEmail = pstrCustomerEmail  & vbcrlf & "You may view your online order status at " & adminDomainName & "orderHistory.asp?orderID=" & mclsCartTotal.OrderID & "&email=" & mclsCustomer.custEmail & VbCrLf

	'Response.Write "<fieldset><legend>Confirmation Emails</legend><h4>Customer</h4>"& Replace(pstrCustomerEmail, vbcrlf, "<br />") & "<hr><h4>Merchant</h4>"& Replace(pstrMerchantEmail, vbcrlf, "<br />") & "</fieldset>"

	pstrSubject = adminEmailSubject
	Call createMail("-", mclsCustomer.custEmail & "|" & "" & "|" & "-" & "|" & pstrSubject & "|" & pstrCustomerEmail)
	'Use InvenNotification type so that CC (ie. Secondary) is cc'd if present)
	Call createMail("Confirm", adminPrimaryEmail & "|" & mclsCustomer.custEmail & "|" & adminSecondaryEmail & "|" & pstrSubject & "|" & pstrMerchantEmail)

End Function	'sendOrderConfirmationEmails
%>