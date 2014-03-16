<!--#include file="SFLib/ssGiftCertificate_class.asp"-->
<%
'********************************************************************************
'*   Gift Certificate Manager				                                    *
'*   Release Version:   1.01.002												*
'*   Release Date:		November 15, 2002										*
'*   Revision Date:		January 24, 2004										*
'*                                                                              *
'*   Release Notes: See ssGiftCertificate_class.asp                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'Common variables
Const cblnUseGiftCertificate = True

Dim mstrCertificate
Dim mdblssCertificateAmount
Dim mdblssGCOriginalTotalDue
Dim mdblssGCNewTotalDue
Dim mblnssGCCalculated

Dim mlngNumCertificatesToCreate
Dim maryGiftCertificatesToCreate()

Dim cstrGiftCertificate_ProductID
Dim cstrStoredCredit_ProductID
Dim cstrGiftCertificate_ToName_attrName
Dim cstrGiftCertificate_ToEmail_attrName
Dim cstrGiftCertificate_FromName_attrName
Dim cstrGiftCertificate_FromEmail_attrName
Dim cstrGiftCertificate_CertificateType_attrName
Dim cstrGiftCertificate_Message_attrName

Dim mdblAmountOfCertificateToCreate
Dim mblnGCUsed_sfReports
Dim mblnValidatedPayment
Dim mstrGiftCertificateRegistrationMessage

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	cstrGiftCertificate_ProductID =					"GiftCertificate"
	cstrStoredCredit_ProductID =					"StoredCredit"
	
	cstrGiftCertificate_ToName_attrName =			"To (Name):"
	cstrGiftCertificate_ToEmail_attrName =			"To (Email):"
	cstrGiftCertificate_FromName_attrName =			"From (Name):"
	cstrGiftCertificate_FromEmail_attrName =		"From (Email):"
	cstrGiftCertificate_Message_attrName =			"Message:"
	cstrGiftCertificate_CertificateType_attrName =	"Certificate Style:"
	'cstrGiftCertificate_CertificateType_attrName =	""

	Const cblnCombineCertificatePurchases = False			'Set to true to combine unique certificate purchases into one certificate, false not to
	Const cblnPermitMultipleCertificateRedemptions = True	'Set to true to permit multiple certificate redemptions on the same order, false not to
	Const cblnShowGCRegistrationBoxOnOrder = True			'Set to true to show registration box on order.asp; False not to
	Const cblnShowGCRegistrationLinkOnOrder = False			'Set to true to show link to registration pop-box on order.asp; False not to

	Const cbytAutoExpire = 0								'Number of months to expire certificate in; Set to 0 for no expiration
	Const cblnActivateOnPurchase = False					'Set to true to create active certificate immediately on purchase, no validation required; False not to
	Const cblnSendEmailOnPurchase = False					'Set to true to immediately send email to recipient; False not to
	Const cstrGC_FromEmail = ""					'Email address to send email from; leave blank to use the default storewide address
	Const cstrCCEmail = "test@isp.com"						'Email address to send cc email to; applies only if cblnSendEmailOnPurchase = True

'/
'/////////////////////////////////////////////////

	mdblAmountOfCertificateToCreate = 0
	mlngNumCertificatesToCreate = 0
	mdblssCertificateAmount = 0
	mdblssGCOriginalTotalDue = 0
	mdblssGCNewTotalDue = 0
	mblnssGCCalculated = False
	mblnValidatedPayment = False

'************************************************************************************************

Function DisplayCertificateCodes(byVal strCode)

Dim pstrTempCode

	pstrTempCode = Trim(strCode & "")
	
	If Len(pstrTempCode) > 0 Then
		pstrTempCode = Left(pstrTempCode, Len(pstrTempCode) - 1)
		pstrTempCode = Right(pstrTempCode, Len(pstrTempCode) - 1)
		pstrTempCode = Replace(pstrTempCode, ";", ", ")
	End If
	
	DisplayCertificateCodes = pstrTempCode

End Function	'DisplayCertificateCodes

'************************************************************************************************

Sub calculateGiftCertificate

Dim pclsssGiftCertificate

	If mblnssGCCalculated Then Exit Sub

	'Set defaults to current grand total
	mdblssGCOriginalTotalDue = CDbl(mclsCartTotal.CartTotal)
	mdblssGCNewTotalDue = CDbl(mclsCartTotal.CartTotal)
	
	If mblnssDebug_GiftCertificate Then Response.Write("calculateGiftCertificate - Cart Total: " & mdblssGCOriginalTotalDue & "<br />")
	
	mstrCertificate = visitorCertificateCodes
	If Len(mstrCertificate) > 0 Then
		Set pclsssGiftCertificate = New clsssGiftCertificate
		pclsssGiftCertificate.Connection = cnn
		If pclsssGiftCertificate.validateCertificate(mstrCertificate) Then
			mdblssCertificateAmount = pclsssGiftCertificate.CertificateValue
			'Now limit to order amount
			If mdblssGCOriginalTotalDue < mdblssCertificateAmount Then mdblssCertificateAmount = mdblssGCOriginalTotalDue

			mdblssGCNewTotalDue = mdblssGCOriginalTotalDue - mdblssCertificateAmount
			mblnssGCCalculated = True
			
		End If
		Set pclsssGiftCertificate = Nothing
	End If

End Sub	'calculateGiftCertificate

'************************************************************************************************

Sub ssGiftCertificate_adjustMail(byRef strEmailBody)

	If Not cblnUseGiftCertificate Then Exit Sub
	If Not mblnssGCCalculated Then Call calculateGiftCertificate

	If mdblssCertificateAmount > 0 Then
		strEmailBody = strEmailBody & vbcrlf _
					& vbcrlf _
					& "Certificate (" & DisplayCertificateCodes(mstrCertificate) & "): " & FormatCurrency(mdblssCertificateAmount, 2) & vbcrlf _
					& "Amount Billed:" & FormatCurrency(mdblssGCNewTotalDue, 2) & vbcrlf
	End If

End Sub	'ssGiftCertificate_adjustMail

'************************************************************************************************

Sub ssGiftCertificate_adjustMerchantMail(byVal strPath, byRef strEmailBody)

	If Not cblnUseGiftCertificate Then Exit Sub

	'create certificate if product ordered
	If mdblAmountOfCertificateToCreate > 0 Then

		strEmailBody = strEmailBody & vbcrlf _
					& vbcrlf _
					& "Gift Certificate Created: " & strPath & "admin/ssAdmin/ssGiftCertificateAdmin.asp" & vbcrlf
	End If

End Sub	'ssGiftCertificate_adjustMerchantMail

'************************************************************************************************

Sub ssGiftCertificate_SaveRedemption(byVal strOrderID)

Dim pclsssGiftCertificate
Dim paryCertificateDetail
Dim pdblRedemptionAmount
Dim i

	If Not cblnUseGiftCertificate Then Exit Sub
	If Not mblnssGCCalculated Then Call calculateGiftCertificate
	
	'update certificate usage
	If mdblssCertificateAmount > 0 Then
		pdblRedemptionAmount = CDbl(mdblssCertificateAmount)
		Set pclsssGiftCertificate = New clsssGiftCertificate
		pclsssGiftCertificate.Connection = cnn
		
		'added because semi-colons in code interfere with saving
		'debugprint "mstrCertificate",mstrCertificate
		If Instr(1, mstrCertificate, ";") > 0 Then 
			Dim paryCert
			paryCert = Split(mstrCertificate, ";")
			mstrCertificate = paryCert(1)
			
			For i = 1 To UBound(paryCert)
				If pdblRedemptionAmount > 0 Then
					If pclsssGiftCertificate.validateCertificate(paryCert(i)) Then
						Call pclsssGiftCertificate.setCertificateCustomerID(paryCert(i), iCustID)
						If pdblRedemptionAmount >= pclsssGiftCertificate.CertificateValue Then
							Call pclsssGiftCertificate.CreateRedemption(False, "", paryCert(i), enOrderRedemption, -1 * pclsssGiftCertificate.CertificateValue, True, strOrderID, "", "")
						Else
							Call pclsssGiftCertificate.CreateRedemption(False, "", paryCert(i), enOrderRedemption, -1 * pdblRedemptionAmount, True, strOrderID, "", "")
						End If
						pdblRedemptionAmount = pdblRedemptionAmount - pclsssGiftCertificate.CertificateValue
					End If
				End If
			Next 'i
			'debugprint "paryCert(i)",paryCert(i)
		Else
			Call pclsssGiftCertificate.setCertificateCustomerID(mstrCertificate, iCustID)
			Call pclsssGiftCertificate.CreateRedemption(False, "", mstrCertificate, enOrderRedemption, -1 * pdblRedemptionAmount, True, strOrderID, "", "")
		End If

		Set pclsssGiftCertificate = Nothing
	End If

	'create certificate if product ordered
	If mdblAmountOfCertificateToCreate > 0 Then
		Set pclsssGiftCertificate = New clsssGiftCertificate
		With pclsssGiftCertificate
			.Connection = cnn
			For i = 1 To UBound(maryGiftCertificatesToCreate)
				paryCertificateDetail = maryGiftCertificatesToCreate(i)(1)
				
				'Not all installations support text based attributes so use common fields instead
				If Len(paryCertificateDetail(enGC_ToEmail)) = 0 Then paryCertificateDetail(enGC_ToEmail) = mclsCustomerShipAddress.Email
				If Len(paryCertificateDetail(enGC_ToName)) = 0 Then paryCertificateDetail(enGC_ToName) = mclsCustomerShipAddress.DisplayName
				If Len(paryCertificateDetail(enGC_FromEmail)) = 0 Then paryCertificateDetail(enGC_FromEmail) = mclsCustomer.custEmail
				If Len(paryCertificateDetail(enGC_FromName)) = 0 Then paryCertificateDetail(enGC_FromName) = mclsCustomer.DisplayName
				If Len(paryCertificateDetail(enGC_Message)) = 0 Then paryCertificateDetail(enGC_Message) = sShipInstructions
				
				.ssGCToName = paryCertificateDetail(enGC_ToName)
				.ssGCIssuedToEmail = paryCertificateDetail(enGC_ToEmail)
				.ssGCFromName = paryCertificateDetail(enGC_FromName)
				.ssGCFromEmail = paryCertificateDetail(enGC_FromEmail)
				.ssGCMessage = paryCertificateDetail(enGC_Message)
				If cbytAutoExpire = 0 Then
					.ssGCExpiresOn = ""
				Else
					.ssGCExpiresOn = DateAdd("m", cbytAutoExpire, Date())
				End If
				.ssGCSingleUse = False
				
				'The next two aren't currently implemented
				.ssGCElectronic = False
				.ssGCFreeText = ""

				'If it is a self-issue "purchase card" then set the customer ID
				If Not paryCertificateDetail(enGC_CertificateOrCredit) Then .ssGCCustomerID = iCustID
				
				.createCertificate_New strOrderID
				
				mstrCertificate = .ssGCCode

				If mblnssDebug_GiftCertificate Then 
					Response.Write "<b>Created Gift Certificate " & i & ": " & .ssGCCode & "</b><br />"
					Response.Write "- Issued to: " & .ssGCIssuedToEmail & "<br />"
					Response.Write "- Certificate Type: " & maryGCRedemptionTypes(paryCertificateDetail(enGC_CertificateType)) & "<br />"
				End If	'mblnssDebug_GiftCertificate
				
				Call pclsssGiftCertificate.CreateRedemption(True, "", mstrCertificate, paryCertificateDetail(enGC_CertificateType), maryGiftCertificatesToCreate(i)(0), cblnActivateOnPurchase, strOrderID, "", "")
				Call SendCertificateEmail(pclsssGiftCertificate, paryCertificateDetail(enGC_CertificateType))
			Next 'i
		End With
		Set pclsssGiftCertificate = Nothing
	End If

End Sub	'ssGiftCertificate_SaveRedemption

'************************************************************************************************

Sub ssGiftCertificate_CheckCartForCertificates_MT(byRef aryOrderItem)
'Replaced Sub ssGiftCertificate_CheckCartForCertificates(strProductID, dblAmount, aAllProd, lngCounter) from original gc module
'Purpose: Checks cart contents during final cart roll-up in confirm.asp and collects gift certificate products into array for future processing

Dim pstrAttributeName
Dim pstrAttributeDetail
Dim paryCertificateDetail
Dim pstrCertTypes
Dim plngProductQty
Dim pdblProductValue
Dim i, j

	If mblnssDebug_GiftCertificate Then Response.Write "Checking product <b>" & aryOrderItem(enOrderItem_prodID) & "</b> to see if it is a certificate <b>" & cstrGiftCertificate_ProductID & "</b> - Result: " & CBool(cstrGiftCertificate_ProductID = aryOrderItem(enOrderItem_prodID)) & "<br />"
	If mblnssDebug_GiftCertificate Then Response.Write "Checking product <b>" & aryOrderItem(enOrderItem_prodID) & "</b> to see if it is a stored credit <b>" & cstrStoredCredit_ProductID & "</b> - Result: " & CBool(cstrStoredCredit_ProductID = aryOrderItem(enOrderItem_prodID)) & "<br />"
	If CBool(cstrGiftCertificate_ProductID = Trim(aryOrderItem(enOrderItem_prodID) & "")) Or CBool(cstrStoredCredit_ProductID = Trim(aryOrderItem(enOrderItem_prodID) & "")) Then

		mlngNumCertificatesToCreate = mlngNumCertificatesToCreate + 1
		ReDim Preserve maryGiftCertificatesToCreate(mlngNumCertificatesToCreate)

		If cblnCombineCertificatePurchases Then
			plngProductQty = 1
			pdblProductValue = aryOrderItem(enOrderItem_UnitPrice) * aryOrderItem(enOrderItem_odrdttmpQuantity)
		Else
			plngProductQty = aryOrderItem(enOrderItem_odrdttmpQuantity)
			pdblProductValue = aryOrderItem(enOrderItem_UnitPrice)
		End If
		
		If mblnssDebug_GiftCertificate Then 
			Response.Write "Combine Certificate Purchases: " & cblnCombineCertificatePurchases & "<br />"
			Response.Write "Product Qty: " & plngProductQty & "<br />"
			Response.Write "Product Price: " & pdblProductValue & "<br />"
			Response.Write "Num Certificates To Create: " & mlngNumCertificatesToCreate & "<br />"
		End If
		
		pstrCertTypes = Join(maryGCRedemptionTypes, "|") & "|"

		ReDim paryCertificateDetail(6)	'Set the certificate array
		
		'Set the default certificate type
		If CBool(cstrStoredCredit_ProductID = aryOrderItem(enOrderItem_prodID)) Then
			paryCertificateDetail(enGC_CertificateType) = enStoreCredit
		Else
			paryCertificateDetail(enGC_CertificateType) = clngDefaultCertificateType
		End If

		'This section only applies for Attribute Extender enabled carts
		'Attribute comes in as Attribute Detail Name + space + Attribute Value
		Dim paryAttributes
		If aryOrderItem(enOrderItem_AttributeCount) > 0 Then
			For i = 0 To aryOrderItem(enOrderItem_AttributeCount) - 1
				pstrAttributeName = aryOrderItem(enOrderItem_AttributeArray)(i)(enAttributeItem_attrName)
				pstrAttributeDetail = aryOrderItem(enOrderItem_AttributeArray)(i)(enAttributeItem_attrdtName)
				
				If mblnssDebug_GiftCertificate Then 
					Response.Write "<fieldset><legend>Attribute " & i & "</legend>"
					Response.Write "pstrAttributeName: " & pstrAttributeName & "<br />"
					Response.Write "pstrAttributeDetail: " & pstrAttributeDetail & "<br />"
					Response.Write "</fieldset>"
				End If	'mblnssDebug_GiftCertificate
				
				If Len(cstrGiftCertificate_ToName_attrName) > 0 And (pstrAttributeName = cstrGiftCertificate_ToName_attrName) Then
					paryCertificateDetail(enGC_ToName) = pstrAttributeDetail
				ElseIf Len(cstrGiftCertificate_ToEmail_attrName) > 0 And (pstrAttributeName = cstrGiftCertificate_ToEmail_attrName) Then
					paryCertificateDetail(enGC_ToEmail) = pstrAttributeDetail
				ElseIf Len(cstrGiftCertificate_FromName_attrName) > 0 And (pstrAttributeName = cstrGiftCertificate_FromName_attrName) Then
					paryCertificateDetail(enGC_FromName) = pstrAttributeDetail
				ElseIf Len(cstrGiftCertificate_FromEmail_attrName) > 0 And (pstrAttributeName = cstrGiftCertificate_FromEmail_attrName) Then
					paryCertificateDetail(enGC_FromEmail) = pstrAttributeDetail
				ElseIf Len(cstrGiftCertificate_Message_attrName) > 0 And (pstrAttributeName = cstrGiftCertificate_Message_attrName) Then
					paryCertificateDetail(enGC_Message) = pstrAttributeDetail
				ElseIf Len(cstrGiftCertificate_CertificateType_attrName) > 0 And (pstrAttributeName = cstrGiftCertificate_CertificateType_attrName) Then
					'Certificate type requires special handling
					If mblnssDebug_GiftCertificate Then Response.Write "Checking Certificate Type: " & cstrGiftCertificate_CertificateType_attrName & " (" & pstrAttributeDetail & ")<br />"
					If UBound(maryGCRedemptionTypes) > 1 Then
						For j = 2 To UBound(maryGCRedemptionTypes)	'start at 2 because 0 & 1 aren't certificate types
							If pstrAttributeDetail = maryGCRedemptionTypes(j) Then
								paryCertificateDetail(enGC_CertificateType) = j
								Exit For
							End If
						Next 'j
					End If
				End If

			Next 'i
			
			'Set the certificate/stored credit
			paryCertificateDetail(enGC_CertificateOrCredit) = CBool(cstrGiftCertificate_ProductID = aryOrderItem(enOrderItem_prodID))
			If Not paryCertificateDetail(enGC_CertificateOrCredit) Then
				paryCertificateDetail(enGC_CertificateType) = enStoreCredit	'Set the default certificate type
			End If
             
			If mblnssDebug_GiftCertificate Then 
				Response.Write "<fieldset><legend>GiftCertificatesToCreate: " & mlngNumCertificatesToCreate & "</legend>"
				Response.Write "- Amount: " & pdblProductValue & "<br />"
				For i = 0 To UBound(paryCertificateDetail)
					Select Case i
						Case enGC_ToName: Response.Write "- ToName"
						Case enGC_ToEmail: Response.Write "- ToEmail"
						Case enGC_FromName: Response.Write "- FromName"
						Case enGC_FromEmail: Response.Write "- FromEmail"
						Case enGC_CertificateType: Response.Write "- CertificateType"
						Case enGC_Message: Response.Write "- Message"
						Case enGC_CertificateOrCredit: Response.Write "- Certificate Or Credit"
					End Select					
					Response.Write ": " & paryCertificateDetail(i) & "<br />"
				Next 'i
				Response.Write "</fieldset>"
			End If	'mblnssDebug_GiftCertificate
        End If
         	
		maryGiftCertificatesToCreate(mlngNumCertificatesToCreate) = Array(pdblProductValue, paryCertificateDetail)
		For i = 2 To plngProductQty
			mlngNumCertificatesToCreate = mlngNumCertificatesToCreate + 1
			ReDim Preserve maryGiftCertificatesToCreate(mlngNumCertificatesToCreate)
			maryGiftCertificatesToCreate(mlngNumCertificatesToCreate) = maryGiftCertificatesToCreate(mlngNumCertificatesToCreate - 1)
			If mblnssDebug_GiftCertificate Then 
				Response.Write "GiftCertificatesToCreate: " & mlngNumCertificatesToCreate & "<br />"
				Response.Write "- Duplicate of : " & mlngNumCertificatesToCreate - 1 & "<br />"
			End If	'mblnssDebug_GiftCertificate
		Next 'i
	
		mdblAmountOfCertificateToCreate = mdblAmountOfCertificateToCreate + plngProductQty
		If mblnssDebug_GiftCertificate Then Response.Write "mdblAmountOfCertificateToCreate: " & mdblAmountOfCertificateToCreate & "<br />"
	End If	'cstrGiftCertificate_ProductID = aryOrderItem(enOrderItem_prodID)

End Sub	'ssGiftCertificate_CheckCartForCertificates_MT

'************************************************************************************************

Function GC_sfReports(lngOrderID, dblGrandTotal)

Dim pclsssGiftCertificate
Dim pblnIsActive
Dim pbytssGCRedemptionType

	If Len(dblGrandTotal & "") = 0 Then
		mdblssGCOriginalTotalDue = 0
	Else
		mdblssGCOriginalTotalDue = CDbl(dblGrandTotal)
	End If
	
	Set pclsssGiftCertificate = New clsssGiftCertificate
	pclsssGiftCertificate.Connection = cnn
	If pclsssGiftCertificate.LoadByOrder(lngOrderID) Then
		mstrCertificate = pclsssGiftCertificate.ssGCCode
		mdblssCertificateAmount = pclsssGiftCertificate.ssGCRedemptionAmount
		pblnIsActive = pclsssGiftCertificate.ssGCRedemptionActive
		pbytssGCRedemptionType = pclsssGiftCertificate.ssGCRedemptionType
		mdblssGCNewTotalDue = mdblssGCOriginalTotalDue + mdblssCertificateAmount
	End If
	Set pclsssGiftCertificate = Nothing

	mblnGCUsed_sfReports = CBool((mdblssCertificateAmount <> 0) And (pbytssGCRedemptionType = enOrderRedemption) And pblnIsActive)
	'GC_sfReports = mblnGCUsed_sfReports
	GC_sfReports = mblnGCUsed_sfReports And Not (mdblssGCNewTotalDue > 0)

End Function	'GC_sfReports

'************************************************************************************************

Sub displayGC_sfReports

If mblnGCUsed_sfReports Then
%>
<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
<tr>
	<td align="left"><b>Certificate (<a href="ssadmin/ssGiftCertificateAdmin.asp?Action=ViewByCode&ssGCCode=<%= DisplayCertificateCodes(mstrCertificate) %>"><%= DisplayCertificateCodes(mstrCertificate) %></a>):</b></td>
	<td align="right"><b><%= FormatCurrency(mdblssCertificateAmount, 2) %></b></td>
</tr>
<tr>
	<td align="left"><b>Amount Billed:</b></td>
	<td align="right"><b><%= FormatCurrency(mdblssGCNewTotalDue, 2) %></b></td>
</tr>
<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
<% 
End If

End Sub	'displayGC_sfReports

'************************************************************************************************

Function displayGC_sfReports_GCLink(lngOrderID, strProductID)

Dim pstrLink
Dim pstrTitle
Dim pclsssGiftCertificate
Dim pblnIsActive
Dim pbytssGCRedemptionType

	Set pclsssGiftCertificate = New clsssGiftCertificate
	pclsssGiftCertificate.Connection = cnn
	If pclsssGiftCertificate.LoadByOrder(lngOrderID) Then
		mstrCertificate = pclsssGiftCertificate.ssGCCode
		pblnIsActive = pclsssGiftCertificate.ssGCRedemptionActive
		pbytssGCRedemptionType = pclsssGiftCertificate.ssGCRedemptionType
		If pblnIsActive Then
			pstrTitle = " title=" & Chr(34) & "Certificate has been activated" & Chr(34)
		Else
			pstrTitle = " title=" & Chr(34) & "Certificate requires activation" & Chr(34)
		End If
	End If
	Set pclsssGiftCertificate = Nothing

	If cstrGiftCertificate_ProductID = strProductID Then
		pstrLink = "<a href=" & Chr(34) & "ssadmin/ssGiftCertificateAdmin.asp?Action=ViewByCode&ssGCCode=" & DisplayCertificateCodes(mstrCertificate) & Chr(34) & pstrTitle & ">" & strProductID & "</a>"
	Else
		pstrLink = strProductID
	End If
	
	displayGC_sfReports_GCLink = pstrLink

End Function	'displayGC_sfReports_GCLink

'************************************************************************************************

Function loadCertificatesByCustID(byVal strEmail, byRef aryCertificates)

'Input: Email address
'Output: array of certificates if found
'		 Format array(certificateCode, Amount Remaining, Expiration Date, Active, Expired)
'Notes: it is up to the calling function to deal with invalid certificates

Dim pclsssGiftCertificate

	Set pclsssGiftCertificate = New clsssGiftCertificate
	pclsssGiftCertificate.Connection = cnn
	
	If pclsssGiftCertificate.LoadByGCCustomerID(strEmail, False) Then
		aryCertificates = pclsssGiftCertificate.aryCertificates
		loadCertificatesByCustID = True
	Else
		loadCertificatesByCustID = False
		If mblnssDebug_GiftCertificate Then Response.Write "<h4>loadCertificatesByCustID:LoadByGCCustomerID (" & strEmail & ") = False</h4>"
	End If
	Set pclsssGiftCertificate = Nothing

End Function	'loadCertificatesByCustID

'************************************************************************************************

Sub checkForCertificateEntry(byVal strCertificateNumber)

Dim pclsssGiftCertificate
Dim pstrMessage
Dim pstrTemp

	mstrCertificate = strCertificateNumber

	If len(mstrCertificate) > 0 Then
	
		If isGCRegisteredForUse(strCertificateNumber) Then Exit Sub

		Set pclsssGiftCertificate = New clsssGiftCertificate
		pclsssGiftCertificate.Connection = cnn
		If pclsssGiftCertificate.validateCertificate(mstrCertificate) Then
		
			'Set the Certificate to the session variables
			pstrTemp = visitorCertificateCodes
			If len(pstrTemp) = 0 Then
				pstrTemp = ";" & mstrCertificate & ";"
			Else
				If instr(1,pstrTemp,";" & mstrCertificate & ";",1) = 0 Then pstrTemp = pstrTemp & mstrCertificate & ";"
			End If
			Call setVisitorPreference("visitorCertificateCodes", pstrTemp, True)
		Else
			mstrGiftCertificateRegistrationMessage = pclsssGiftCertificate.Message
		End If
		Set pclsssGiftCertificate = Nothing
	End If
	
End Sub	'checkForCertificateEntry

'************************************************************************************************

Function deleteGCRegisteredForUse(byVal strCertificateNumber)

Dim pblnResult
Dim pstrTemp

	pblnResult = False
	
	If isGCRegisteredForUse(strCertificateNumber) Then
		pstrTemp = ";" & strCertificateNumber & ";"
		pstrTemp = Replace(visitorCertificateCodes, pstrTemp, "")
		Call setVisitorPreference("visitorCertificateCodes", pstrTemp, True)
		pblnResult = True
	End If

	deleteGCRegisteredForUse = pblnResult
	
End Function	'deleteGCRegisteredForUse

'************************************************************************************************

Function isGCRegisteredForUse(byVal strCertificateNumber)

Dim pblnResult
Dim pstrTemp

	pblnResult = False
	
	If len(strCertificateNumber) > 0 Then
		pstrTemp = ";" & strCertificateNumber & ";"
		pblnResult = CBool(instr(1, visitorCertificateCodes, pstrTemp) > 0)
	End If

	isGCRegisteredForUse = pblnResult
	
End Function	'isGCRegisteredForUse

'***********************************************************************************************

Sub SendCertificateEmail(byRef objclsssGiftCertificate, byVal bytCertificateType)

Dim p_strSubject
Dim p_strBody
Dim pstrTrackingLink
Dim pstrCustName, pstrShipAddr

'On Error Resume Next

	If Not cblnSendEmailOnPurchase Then Exit Sub
	
	If mblnssDebug_GiftCertificate Then Response.Write("Sending Certificate Email for " & objclsssGiftCertificate.ssGCCode & ". . .<br />")
	
	If objclsssGiftCertificate.Load(objclsssGiftCertificate.ssGCCode) Then
		Call LoadEmailFile(p_strSubject, p_strBody, maryGCRedemptionEmailsAutomatic(bytCertificateType))
		
		'now replace the constants
		p_strSubject = customReplacements(p_strSubject, objclsssGiftCertificate)
		p_strBody = customReplacements(p_strBody, objclsssGiftCertificate)
		If mblnssDebug_GiftCertificate Then Response.Write("SendCertificateEmail - p_strBody: " & p_strBody & "<br />")

		Call createMail("-", objclsssGiftCertificate.ssGCIssuedToEmail & "|" & cstrGC_FromEmail & "|" & "" & "|" & p_strSubject & "|" & p_strBody)
		If Len(cstrCCEmail) > 0 Then Call createMail("-", cstrCCEmail & "|" & cstrGC_FromEmail & "|" & "" & "|" & p_strSubject & "|" & p_strBody)

		Call objclsssGiftCertificate.updateEmailSent(True)
	Else
		If mblnssDebug_GiftCertificate Then Response.Write("Unable to load " & objclsssGiftCertificate.ssGCCode & "<br />")
	End If

End Sub	'SendCertificateEmail

'************************************************************************************************

Sub LoadEmailFile(byRef strEmailSubject, byRef strEmailBody, byVal strTemplate)

Dim pobjFSO
Dim MyFile
Dim pstrTempLine
Dim pstrFilePath
Dim p_strSubject
Dim p_strBody

'On Error Resume Next

	pstrFilePath = Request.ServerVariables("PATH_TRANSLATED")
	pstrFilePath = Replace(Lcase(pstrFilePath),"confirm.asp","ssGCTemplates/" & strTemplate)

	Set pobjFSO = CreateObject("Scripting.FileSystemObject")
	Set MyFile =pobjFSO.OpenTextFile(pstrFilePath,1,True)

	p_strSubject = MyFile.ReadLine
	pstrTempLine = MyFile.ReadLine	'garbage line
	pstrTempLine = MyFile.ReadLine & vbcrlf
	Do While pstrTempLine <> "// DO NOT REMOVE THIS LINE //" AND NOT MyFile.AtEndOfStream
		p_strBody = p_strBody & pstrTempLine & vbcrlf
		pstrTempLine = MyFile.ReadLine
	Loop
	
	strEmailSubject = p_strSubject
	strEmailBody = p_strBody

	MyFile.Close
	Set MyFile = Nothing
	Set pobjFSO = Nothing

End Sub	'LoadEmailFile

'************************************************************************************************

%>
