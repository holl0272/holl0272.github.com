<%
'********************************************************************************
'*   Gift Certificate Manager				                                    *
'*   Release Version:   1.01.002												*
'*   Release Date:		November 15, 2002										*
'*   Revision Date:		November 6, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Version 1.02.002 (November 6, 2004)	                                    *
'*     - enhancement: added email sent confirmatin to result message			*
'*                                                                              *
'*   Version 1.02.001 (December 5, 2003)	                                    *
'*     - enhancement: added support for automatic email creation				*
'*                                                                              *
'*   Version 1.01.001 (October 13, 2003)	                                    *
'*     - Note: Implemented new versioning system (Major.Minor.Build)            *
'*     - enhancement: added support for additional text based fields			*
'*     - enhancement: added support for unlimited, custom certificate types		*
'*     - enhancement: added support for unlimited, customized email templates	*
'*     - enhancement: added strong protection for SQL Injection attacks			*
'*     - enhancement: added support for unlimited, custom certificate viewing	*
'*     - enhancement: added support to permit multiple certificate redemptions	*
'*     - enhancement: added support to collect certificate numbers on order.asp	*
'*                                                                              *
'*   Version 1.07 (March 21, 2003)		                                        *
'*     - enhancement: added certificate preview page                            *
'*                                                                              *
'*   Version 1.06 (February 26, 2003)                                           *
'*     - bug fix: Summary table paging fixed                                    *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Dim maryGCRedemptionTypes
Dim maryGCRedemptionEmailsAutomatic
Dim clngDefaultCertificateType

'/////////////////////////////////////////////////
'/
'/  User Parameters - Section 1
'/

	'Enumerations
	Const enOrderRedemption = 1			' DO NOT ALTER
	Const enStoreCredit = 2				' DO NOT ALTER
	Const enGiftCertificateStyle1 = 3
	Const enGiftCertificateStyle2 = 4
	Const enGiftCertificateStyle3 = 5	
	'Additional methods can be added here
	
	clngDefaultCertificateType = enGiftCertificateStyle1	'This should be one of the enumerations from above

	'Names of the various redemption types - DO NOT change the first two entries "","Order Redemption"
	maryGCRedemptionTypes = Array("","Order Redemption", _
								  "Store Credit", _
								  "Anniversary", _
								  "Thank You", _
								  "Traditional")
								  
	'Email files to use - DO NOT change the first two entries "","" - MUST have equal number of entries as maryGCRedemptionTypes
	maryGCRedemptionEmailsAutomatic = Array("","", _
											"ssGiftCertificateEmail_complete.txt", _
											"ssGiftCertificateEmail_complete.txt", _
											"ssGiftCertificateEmail_complete.txt", _
											"ssGiftCertificateEmail_complete.txt")
	
'/
'/////////////////////////////////////////////////

Class clsssGiftCertificate
'Assumptions:
'   Connection: defines a previously opened connection to the database

'Internal Class Variables
Private cstrDelimeter
Private pstrMessage
Private pobjCnn
Private pobjRS
Private pobjRS_Summary
Private pblnError
Private cblnCaseInsensitive
Private paryCertificates

'database variables

Private plngssGCID
Private pstrssGCCode
Private pdtssGCCreatedOn
Private pblnssGCElectronic
Private pdtssGCExpiresOn
Private pstrssGCIssuedToEmail
Private pdtssGCModifiedOn
Private pblnssGCSingleUse
Private plngssGCCustomerID

'added with v1.01.001
Private pstrssGCFreeText
Private pstrssGCToName
Private pstrssGCFromName
Private pstrssGCFromEmail
Private pstrssGCMessage

Private pdblssGCRedemptionAmount
Private pdtssGCRedemptionCreatedOn
Private pstrssGCRedemptionExternalNotes
Private plngssGCRedemptionID
Private pstrssGCRedemptionInternalNotes
Private plngssGCRedemptionOrderID
Private pbytssGCRedemptionType
Private pblnssGCRedemptionActive

'Parameters to create Certificate Numbers
Private cstrGCPrefix
Private cstrGCSuffix
Private clngGCMinNumber
Private clngGCMaxNumber
Private cbytGCLength
Private cblnGCUseOrderNumber
Private cblnGCGenerateRandom

'Derived Values
Private pdblCertificateValue
Private pblnExpired

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
    cblnCaseInsensitive = True
    
'/////////////////////////////////////////////////
'/
'/  User Parameters - Section 2
'/

	cstrGCPrefix = "GC"
	cstrGCSuffix = "XX"
	clngGCMinNumber = 0
	clngGCMaxNumber = 9999999
	cbytGCLength = 7
	cblnGCUseOrderNumber = False
	cblnGCGenerateRandom = True

'/
'/////////////////////////////////////////////////

End Sub

Private Sub class_Terminate()

	On Error Resume Next

	pobjRS_Summary.Close
	Set pobjRS_Summary = Nothing

	pobjRS.Close
	Set pobjRS = Nothing

	If Err.number <> 0 Then Err.Clear

End Sub

'***********************************************************************************************

Public Property Let Connection(objCnn)
    Set pobjCnn = objCnn
End Property

Public Property Get Recordset()
	If isObject(pobjRS) Then Set Recordset = pobjRS
End Property

Public Property Get SummaryRecordset()
    set SummaryRecordset = pobjRS_Summary
End Property

'***********************************************************************************************

Public Property Get aryCertificates()
    aryCertificates = paryCertificates
End Property

Public Property Get Message()
    Message = pstrMessage
End Property

Public Sub OutputMessage()

Dim i
Dim aError

    aError = Split(pstrMessage, cstrDelimeter)
    For i = 0 To UBound(aError)
        If pblnError Then
            Response.Write "<P align='center'><H4><FONT color=Red>" & aError(i) & "</FONT></H4></P>"
        Else
            Response.Write "<P align='center'><H4>" & aError(i) & "</H4></P>"
        End If
    Next 'i

End Sub 'OutputMessage

'Derived Values
Public Property Get CertificateValue()
    CertificateValue = pdblCertificateValue
End Property
Public Property Get Expired()
    Expired = pblnExpired
End Property

Public Property Let ssGCRedemptionActive(blnssGCRedemptionActive)
    pblnssGCRedemptionActive = blnssGCRedemptionActive
End Property

Public Property Get ssGCRedemptionActive()
    ssGCRedemptionActive = pblnssGCRedemptionActive
End Property

Public Property Let ssGCCode(strssGCCode)

Const p_intFieldLength = 20

    If (Len(strssGCCode) > p_intFieldLength)  then
        pstrssGCCode = Left(vData, p_intFieldLength)
    Else
        pstrssGCCode = strssGCCode
    End If

End Property

Public Property Get ssGCCode()
    ssGCCode = pstrssGCCode
End Property

Public Property Let ssGCCreatedOn(dtssGCCreatedOn)
    pdtssGCCreatedOn = dtssGCCreatedOn
End Property

Public Property Get ssGCCreatedOn()
    ssGCCreatedOn = pdtssGCCreatedOn
End Property

Public Property Let ssGCElectronic(blnssGCElectronic)
    pblnssGCElectronic = blnssGCElectronic
End Property

Public Property Get ssGCElectronic()
    ssGCElectronic = pblnssGCElectronic
End Property

Public Property Let ssGCExpiresOn(dtssGCExpiresOn)
    pdtssGCExpiresOn = dtssGCExpiresOn
End Property

Public Property Get ssGCExpiresOn()
    ssGCExpiresOn = pdtssGCExpiresOn
End Property

Public Property Get customExpirationDate()
	If isDate(pdtssGCExpiresOn) Then
	    customExpirationDate = pdtssGCExpiresOn
	Else
	    customExpirationDate = "No Expiration"
	End If
End Property

Public Property Let ssGCID(lngssGCID)
    plngssGCID = lngssGCID
End Property

Public Property Get ssGCID()
    ssGCID = plngssGCID
End Property

Public Property Let ssGCIssuedToEmail(strssGCIssuedToEmail)

Const p_intFieldLength = 255

    If (Len(strssGCIssuedToEmail) > p_intFieldLength)  then
        pstrssGCIssuedToEmail = Left(vData, p_intFieldLength)
    Else
        pstrssGCIssuedToEmail = strssGCIssuedToEmail
    End If

End Property

Public Property Get ssGCIssuedToEmail()
    ssGCIssuedToEmail = pstrssGCIssuedToEmail
End Property

Public Property Let ssGCModifiedOn(dtssGCModifiedOn)
    pdtssGCModifiedOn = dtssGCModifiedOn
End Property

Public Property Get ssGCModifiedOn()
    ssGCModifiedOn = pdtssGCModifiedOn
End Property

Public Property Let ssGCSingleUse(blnssGCSingleUse)
    pblnssGCSingleUse = blnssGCSingleUse
End Property

Public Property Get ssGCSingleUse()
    ssGCSingleUse = pblnssGCSingleUse
End Property

Public Property Let ssGCCustomerID(lngssGCCustomerID)
    plngssGCCustomerID = lngssGCCustomerID
End Property

Public Property Get ssGCCustomerID()
    ssGCCustomerID = plngssGCCustomerID
End Property

Public Property Get ssGCRedemptionAmount()
    ssGCRedemptionAmount = pdblssGCRedemptionAmount
End Property

Public Property Get ssGCRedemptionCreatedOn()
    ssGCRedemptionCreatedOn = pdtssGCRedemptionCreatedOn
End Property

Public Property Get ssGCRedemptionExternalNotes()
    ssGCRedemptionExternalNotes = pstrssGCRedemptionExternalNotes
End Property

Public Property Get ssGCRedemptionID()
    ssGCRedemptionID = plngssGCRedemptionID
End Property

Public Property Get ssGCRedemptionInternalNotes()
    ssGCRedemptionInternalNotes = pstrssGCRedemptionInternalNotes
End Property

Public Property Get ssGCRedemptionOrderID()
    ssGCRedemptionOrderID = plngssGCRedemptionOrderID
End Property

Public Property Get ssGCRedemptionType()
    ssGCRedemptionType = pbytssGCRedemptionType
End Property

'added with v1.01.001
Public Property Let ssGCFreeText(vntValue)
    pstrssGCFreeText = vntValue
End Property
Public Property Get ssGCFreeText()
    ssGCFreeText = pstrssGCFreeText
End Property

Public Property Let ssGCToName(vntValue)
    pstrssGCToName = vntValue
End Property
Public Property Get ssGCToName()
    ssGCToName = pstrssGCToName
End Property

Public Property Let ssGCFromName(vntValue)
    pstrssGCFromName = vntValue
End Property
Public Property Get ssGCFromName()
    ssGCFromName = pstrssGCFromName
End Property

Public Property Let ssGCFromEmail(vntValue)
    pstrssGCFromEmail = vntValue
End Property
Public Property Get ssGCFromEmail()
    ssGCFromEmail = pstrssGCFromEmail
End Property

Public Property Let ssGCMessage(vntValue)
    pstrssGCMessage = vntValue
End Property
Public Property Get ssGCMessage()
    ssGCMessage = pstrssGCMessage
End Property

'***********************************************************************************************

Private Sub LoadArray(objRS)

Dim i
Dim plngCounter
Dim pstrPrevCode

	With objRS
		If Not .EOF Then

			'This is outside the loop because the first entry is the original certificate type; the rest should be redemptions or additions
			pbytssGCRedemptionType = trim(.Fields("ssGCRedemptionType").Value)
			
			plngCounter = -1
			ReDim paryCertificates(0)
			For i = 1 To .RecordCount
				If pstrPrevCode <> Trim(.Fields("ssGCCode").Value) Then
					pstrPrevCode = Trim(.Fields("ssGCCode").Value)
					plngCounter = plngCounter + 1
					ReDim Preserve paryCertificates(plngCounter)
					paryCertificates(plngCounter) = Array(.Fields("ssGCCode").Value, 0, .Fields("ssGCExpiresOn").Value, .Fields("ssGCRedemptionActive").Value, False)
				End If
				.MoveNext
			Next	'i
			.MoveFirst
		End If
	End With 
	
	For i = 0 To UBound(paryCertificates)
		Call validateCertificate(paryCertificates(i)(0))
		paryCertificates(i)(1) = pdblCertificateValue
		paryCertificates(i)(4) = pblnExpired
	Next	'i
    
End Sub 'LoadArray

Private Sub LoadValues(objRS)

Dim i
Dim plngCounter
Dim pstrPrevCode

	With objRS
		If Not .EOF Then
			
			'This section is outside the loop because the first record is the original certificate type; the rest should be redemptions or additions
			pbytssGCRedemptionType = .Fields("ssGCRedemptionType").Value
			pstrssGCFreeText = Trim(.Fields("ssGCFreeText").Value)
			pstrssGCToName = Trim(.Fields("ssGCToName").Value)
			pstrssGCFromName = Trim(.Fields("ssGCFromName").Value)
			pstrssGCFromEmail = Trim(.Fields("ssGCFromEmail").Value)
			pstrssGCMessage = Trim(.Fields("ssGCMessage").Value)

			plngCounter = -1
			ReDim paryCertificates(0)
			For i = 1 To .RecordCount
				pstrssGCCode = Trim(.Fields("ssGCCode").Value)
				If pstrPrevCode <> pstrssGCCode Then
					pstrPrevCode = pstrssGCCode
					plngCounter = plngCounter + 1
					ReDim Preserve paryCertificates(plngCounter)
					
					pdtssGCCreatedOn = .Fields("ssGCCreatedOn").Value
					pblnssGCElectronic = .Fields("ssGCElectronic").Value
					pdtssGCExpiresOn = .Fields("ssGCExpiresOn").Value
					plngssGCID = .Fields("ssGCID").Value
					pstrssGCIssuedToEmail = Trim(.Fields("ssGCIssuedToEmail").Value)
					pdtssGCModifiedOn = .Fields("ssGCModifiedOn").Value
					pblnssGCSingleUse = .Fields("ssGCSingleUse").Value
					plngssGCCustomerID = .Fields("ssGCCustomerID").Value

					pdblssGCRedemptionAmount = Trim(.Fields("ssGCRedemptionAmount").Value)
					pblnssGCRedemptionActive = .Fields("ssGCRedemptionActive").Value
					pdtssGCRedemptionCreatedOn = Trim(.Fields("ssGCRedemptionCreatedOn").Value)
					pstrssGCRedemptionExternalNotes = trim(.Fields("ssGCRedemptionExternalNotes").Value)
					plngssGCRedemptionID = trim(.Fields("ssGCRedemptionID").Value)
					pstrssGCRedemptionInternalNotes = Trim(.Fields("ssGCRedemptionInternalNotes").Value)
					plngssGCRedemptionOrderID = Trim(.Fields("ssGCRedemptionOrderID").Value)
				
					paryCertificates(plngCounter) = Array(pstrssGCCode, 0, pdtssGCExpiresOn, pblnssGCRedemptionActive, False)
				End If
				
				.MoveNext
			Next	'i
			.MoveFirst
		End If
	End With 
	
	For i = 0 To UBound(paryCertificates)
		Call validateCertificate(paryCertificates(i)(0))
		paryCertificates(i)(1) = pdblCertificateValue
		paryCertificates(i)(4) = pblnExpired
	Next	'i
    
End Sub 'LoadValues

Private Sub LoadFromRequest

	'Dim vItem
	'Response.Write "<hr><h4>Form Data</h4>" & vbcrlf
	'For Each vItem in Request.Form
	'	Response.Write vItem & ": " & Request.Form(vItem) & "<br />" & vbcrlf
	'Next
	'Response.Write "<hr>" & vbcrlf
	'Response.Flush

    With Request.Form
        pblnssGCRedemptionActive = (.Item("ssGCRedemptionActive") = "ON")
        pstrssGCCode = Trim(.Item("ssGCCode"))
        pdtssGCCreatedOn = Trim(.Item("ssGCCreatedOn"))
        pblnssGCElectronic = (.Item("ssGCElectronic") = "ON")
        pdtssGCExpiresOn = Trim(.Item("ssGCExpiresOn"))
        plngssGCID = Trim(.Item("ssGCID"))
        pstrssGCIssuedToEmail = Trim(.Item("ssGCIssuedToEmail"))
        pdtssGCModifiedOn = Trim(.Item("ssGCModifiedOn"))
        pblnssGCSingleUse = (.Item("ssGCSingleUse") = "ON")
        plngssGCCustomerID = Trim(.Item("ssGCCustomerID"))

        pdblssGCRedemptionAmount = Trim(.Item("ssGCRedemptionAmount"))
        pdtssGCRedemptionCreatedOn = Trim(.Item("ssGCRedemptionCreatedOn"))
        pstrssGCRedemptionExternalNotes = Trim(.Item("ssGCRedemptionExternalNotes"))
        plngssGCRedemptionID = Trim(.Item("ssGCRedemptionID"))
        pstrssGCRedemptionInternalNotes = Trim(.Item("ssGCRedemptionInternalNotes"))
        plngssGCRedemptionOrderID = Trim(.Item("ssGCRedemptionOrderID"))
        pbytssGCRedemptionType = Trim(.Item("ssGCRedemptionType"))

		pstrssGCFreeText = trim(.Item("ssGCFreeText"))
		pstrssGCToName = trim(.Item("ssGCToName"))
		pstrssGCFromName = trim(.Item("ssGCFromName"))
		pstrssGCFromEmail = trim(.Item("ssGCFromEmail"))
		pstrssGCMessage = trim(.Item("ssGCMessage"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Find(lngssGCID)

Dim pstrSQL
Dim pblnTemp

'On Error Resume Next

	pblnTemp = False
	If Len(lngssGCID) > 0 Then
		pobjRS.Filter = "ssGCID=" & lngssGCID
		If Not pobjRS.EOF Then 
			Call LoadValues(pobjRS)
			pblnTemp = True
		End If
		pobjRS.Filter = ""
    End If
    
    Find = pblnTemp

End Function    'Find

'***********************************************************************************************

Public Function FindByCode(strssGCCode)

Dim pstrSQL
Dim pblnTemp

'On Error Resume Next

	pblnTemp = False
	If Len(strssGCCode) > 0 Then
		pobjRS.Filter = "ssGCCode='" & strssGCCode & "'"
		If Not pobjRS.EOF Then 
			Call LoadValues(pobjRS)
			pblnTemp = True
		End If
		pobjRS.Filter = ""
    End If
    
    FindByCode = pblnTemp

End Function    'FindByCode

'***********************************************************************************************

Public Function RemoveOrphanedGiftCertificates()

Dim pstrSQL
Dim p_objRS
Dim p_strOut
Dim i

'On Error Resume Next

	pstrSQL = "SELECT ssGCID, ssGCCode" _
			& " FROM ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode" _
			& " WHERE (((ssGiftCertificateRedemptions.ssGCRedemptionCGCode) Is Null))"

	Debugprint "pstrSQL",pstrSQL
    Set p_objRS = GetRS(pstrSQL)
    If Not (p_objRS.EOF Or p_objRS.BOF) Then
		Do While Not p_objRS.EOF
			pstrSQL = "Delete From ssGiftCertificates Where ssGCID = " & p_objRS.Fields("ssGCID").Value
			pobjCnn.Execute pstrSQL,,128
			pstrMessage = pstrMessage & "Certificate " & p_objRS.Fields("ssGCCode").Value & " was removed because it had no details. Please hit refresh.<br />"
			p_objRS.MoveNext
		Loop
        RemoveOrphanedGiftCertificates = True
    Else
		RemoveOrphanedGiftCertificates = False
    End If
    p_objRS.Close
    Set p_objRS = Nothing

End Function    'RemoveOrphanedGiftCertificates


'***********************************************************************************************

Public Function LoadByOrder(lngOrderID)

Dim pstrSQL
Dim p_objRS

'On Error Resume Next

	'SQL Injection protection
	If Len(lngOrderID) = 0 Or Not isNumeric(lngOrderID) Then
		LoadByOrder = False
		Exit Function
	End If

	pstrSQL = "SELECT ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
			& " FROM ssGiftCertificateRedemptions" _
			& " WHERE ssGiftCertificateRedemptions.ssGCRedemptionOrderID=" & lngOrderID

    Set p_objRS = CreateObject("adodb.Recordset")
	p_objRS.Open pstrSQL, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
    If Not (p_objRS.EOF Or p_objRS.BOF) Then
		pstrssGCCode = p_objRS.Fields("ssGCRedemptionCGCode").Value
		pdblssGCRedemptionAmount = CDbl(p_objRS.Fields("ssGCRedemptionAmount").Value)
		pblnssGCRedemptionActive = p_objRS.Fields("ssGCRedemptionActive").Value
		pbytssGCRedemptionType = p_objRS.Fields("ssGCRedemptionType").Value
        LoadByOrder = True
    Else
		LoadByOrder = False
    End If
    p_objRS.Close
    Set p_objRS = Nothing

End Function    'LoadByOrder

'***********************************************************************************************

Public Function Load(byVal strssGCCode)

Dim pstrSQL

'On Error Resume Next

	'SQL Injection protection
	If Not allPermissibleCharacters(strssGCCode, cstrPermissibleCharacters) Then
		Load = False
		Exit Function
	End If

	'pstrSQL = "SELECT ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificates.ssGCCustomerID, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCElectronic, ssGiftCertificates.ssGCCreatedOn, ssGiftCertificates.ssGCModifiedOn, ssGiftCertificateRedemptions.ssGCRedemptionID, ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sourceCustomer.custFirstName, sourceCustomer.custMiddleInitial, sourceCustomer.custLastName, sourceCustomer.custEmail" _
	pstrSQL = "SELECT ssGiftCertificates.*, ssGiftCertificateRedemptions.ssGCRedemptionID, ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sourceCustomer.custFirstName, sourceCustomer.custMiddleInitial, sourceCustomer.custLastName, sourceCustomer.custEmail" _
			& " FROM (((ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode) LEFT JOIN sfCustomers ON ssGiftCertificates.ssGCCustomerID = sfCustomers.custID) LEFT JOIN sfOrders ON ssGiftCertificateRedemptions.ssGCRedemptionOrderID = sfOrders.orderID) LEFT JOIN sfCustomers AS sourceCustomer ON sfOrders.orderCustId = sourceCustomer.custID" _
			& " WHERE ssGCCode='" & strssGCCode & "'"
    Set pobjRS = CreateObject("adodb.Recordset")

	pobjRS.Open pstrSQL, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
    If Not (pobjRS.EOF Or pobjRS.BOF) Then
        Call LoadValues(pobjRS)
        Load = True
    Else
		Load = False
    End If

End Function    'Load

'***********************************************************************************************

Public Function LoadByID(byVal lngssGCID)

Dim pstrSQL

On Error Resume Next

	'pstrSQL = "SELECT ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificates.ssGCCustomerID, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCElectronic, ssGiftCertificates.ssGCCreatedOn, ssGiftCertificates.ssGCModifiedOn, ssGiftCertificateRedemptions.ssGCRedemptionID, ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sourceCustomer.custFirstName, sourceCustomer.custMiddleInitial, sourceCustomer.custLastName, sourceCustomer.custEmail" _
	pstrSQL = "SELECT ssGiftCertificates.*, ssGiftCertificateRedemptions.ssGCRedemptionID, ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sourceCustomer.custFirstName, sourceCustomer.custMiddleInitial, sourceCustomer.custLastName, sourceCustomer.custEmail" _
			& " FROM (((ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode) LEFT JOIN sfCustomers ON ssGiftCertificates.ssGCCustomerID = sfCustomers.custID) LEFT JOIN sfOrders ON ssGiftCertificateRedemptions.ssGCRedemptionOrderID = sfOrders.orderID) LEFT JOIN sfCustomers AS sourceCustomer ON sfOrders.orderCustId = sourceCustomer.custID" _
			& " WHERE ssGCID=" & lngssGCID
    Set pobjRS = CreateObject("adodb.Recordset")
	pobjRS.Open pstrSQL, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
    If Not (pobjRS.EOF Or pobjRS.BOF) Then
        Call LoadValues(pobjRS)
        LoadByID = True
    Else
		LoadByID = False
    End If

End Function    'LoadByID

'***********************************************************************************************

Public Function LoadByGCCustomerID_old(byVal lngGCCustomerID, byVal blnIncludeEmail)

Dim pstrSQL

'On Error Resume Next

	'SQL Injection protection
	If Len(lngGCCustomerID) = 0 Or Not isNumeric(lngGCCustomerID) Then
		LoadByGCCustomerID = False
		Exit Function
	End If

	If blnIncludeEmail Then
		pstrSQL = "SELECT ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCModifiedOn, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificates.ssGCElectronic, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCCreatedOn, Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount) AS SumOfssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, ssGiftCertificates.ssGCCustomerID, Count(ssGiftCertificates.ssGCCode) AS Redemptions, sfCustomers.custID, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
				& " FROM (((ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode) LEFT JOIN sfCustomers ON ssGiftCertificates.ssGCCustomerID = sfCustomers.custID) LEFT JOIN sfOrders ON ssGiftCertificateRedemptions.ssGCRedemptionOrderID = sfOrders.orderID) LEFT JOIN sfCustomers AS sourceCustomer ON sfOrders.orderCustId = sourceCustomer.custID" _
				& " GROUP BY ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCModifiedOn, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificates.ssGCElectronic, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, ssGiftCertificates.ssGCCustomerID, sfCustomers.custID, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
				& " HAVING (ssGiftCertificates.ssGCCustomerID=" & Trim(lngGCCustomerID) & ")" _
				& " ORDER BY ssGiftCertificates.ssGCCode, ssGCID"
	Else	
		pstrSQL = "SELECT ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificates.ssGCCustomerID, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCElectronic, ssGiftCertificates.ssGCCreatedOn, ssGiftCertificates.ssGCModifiedOn, ssGiftCertificateRedemptions.ssGCRedemptionID, ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sourceCustomer.custFirstName, sourceCustomer.custMiddleInitial, sourceCustomer.custLastName, sourceCustomer.custEmail" _
				& " FROM (((ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode) LEFT JOIN sfCustomers ON ssGiftCertificates.ssGCCustomerID = sfCustomers.custID) LEFT JOIN sfOrders ON ssGiftCertificateRedemptions.ssGCRedemptionOrderID = sfOrders.orderID) LEFT JOIN sfCustomers AS sourceCustomer ON sfOrders.orderCustId = sourceCustomer.custID" _
				& " WHERE ssGCCustomerID=" & Trim(lngGCCustomerID)
	End If
	'debugprint "pstrSQL", pstrSQL
			
    Set pobjRS = CreateObject("adodb.Recordset")
	pobjRS.Open pstrSQL, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
    If Not (pobjRS.EOF Or pobjRS.BOF) Then
        Call LoadArray(pobjRS)
        LoadByGCCustomerID = True
    Else
		LoadByGCCustomerID = False
    End If

End Function    'LoadByGCCustomerID_old

'***********************************************************************************************

Public Function LoadByGCCustomerID(byVal lngGCCustomerID, byVal blnIncludeEmail)

Dim plngCounter
Dim pobjCmd
Dim pobjRS
Dim pstrSQL

'On Error Resume Next

	'SQL Injection protection
	If Len(lngGCCustomerID) = 0 Or Not isNumeric(lngGCCustomerID) Then
		LoadByGCCustomerID = False
		Exit Function
	End If

	If blnIncludeEmail Then
		pstrSQL = "SELECT ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCModifiedOn, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificates.ssGCElectronic, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCCreatedOn, Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount) AS SumOfssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, ssGiftCertificates.ssGCCustomerID, Count(ssGiftCertificates.ssGCCode) AS Redemptions, sfCustomers.custID, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
				& " FROM (((ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode) LEFT JOIN sfCustomers ON ssGiftCertificates.ssGCCustomerID = sfCustomers.custID) LEFT JOIN sfOrders ON ssGiftCertificateRedemptions.ssGCRedemptionOrderID = sfOrders.orderID) LEFT JOIN sfCustomers AS sourceCustomer ON sfOrders.orderCustId = sourceCustomer.custID" _
				& " GROUP BY ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCModifiedOn, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificates.ssGCElectronic, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, ssGiftCertificates.ssGCCustomerID, sfCustomers.custID, ssGiftCertificateRedemptions.ssGCRedemptionOrderID, ssGiftCertificateRedemptions.ssGCRedemptionInternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionExternalNotes, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
				& " HAVING (ssGiftCertificates.ssGCCustomerID=" & Trim(lngGCCustomerID) & ")" _
				& " ORDER BY ssGiftCertificates.ssGCCode, ssGCID"
		Set pobjRS = CreateObject("adodb.Recordset")
		pobjRS.Open pstrSQL, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
		If Not (pobjRS.EOF Or pobjRS.BOF) Then
			Call LoadArray(pobjRS)
			LoadByGCCustomerID = True
		Else
			LoadByGCCustomerID = False
		End If
		Call closeObj(pobjRS)
	Else	
		pstrSQL = "SELECT ssGiftCertificates.ssGCCode, Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount) AS SumOfssGCRedemptionAmount, ssGiftCertificates.ssGCExpiresOn" _
				& " FROM ssGiftCertificateRedemptions RIGHT JOIN ssGiftCertificates ON ssGiftCertificateRedemptions.ssGCRedemptionCGCode = ssGiftCertificates.ssGCCode" _
				& " WHERE ((ssGiftCertificates.ssGCExpiresOn Is Null) OR (ssGiftCertificates.ssGCExpiresOn>=?)) AND (ssGiftCertificateRedemptions.ssGCRedemptionActive=1 Or ssGiftCertificateRedemptions.ssGCRedemptionActive=-1)" _
				& " GROUP BY ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCCustomerID" _
				& " HAVING ((Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount)>0) AND (ssGiftCertificates.ssGCCustomerID)=?)"

		Set pobjCmd  = CreateObject("ADODB.Command")
		With pobjCmd
			.Commandtype = adCmdText
			.Commandtext = pstrSQL
			Set .ActiveConnection = cnn
			.Parameters.Append .CreateParameter("ssGCExpiresOn", adDBTimeStamp, adParamInput, 16, Now())
			.Parameters.Append .CreateParameter("ssGCCustomerID", adInteger, adParamInput, 4, lngGCCustomerID)
			Set pobjRS = .Execute
			
			plngCounter = -1
			With pobjRS
				If Not .EOF Then ReDim paryCertificates(0)
				Do While Not .EOF
					plngCounter = plngCounter + 1
					ReDim Preserve paryCertificates(plngCounter)
					paryCertificates(plngCounter) = Array(Trim(.Fields("ssGCCode").Value & ""), _
														  .Fields("SumOfssGCRedemptionAmount").Value, _
														  .Fields("ssGCExpiresOn").Value, _
														  True, _
														  False)
					.MoveNext
				Loop
				.Close
			End With	'pobjRS
			Set pobjRS = Nothing

		End With	'pobjCmd
		Set pobjCmd = Nothing
		LoadByGCCustomerID = CBool(plngCounter <> -1)
	End If
	'debugprint "pstrSQL", pstrSQL

End Function    'LoadByGCCustomerID

'***********************************************************************************************

Public Function LoadAll(aryFilter)

Dim pstrSQL
Dim pstrSQL_Where
Dim pstrSQL_Having
Dim pstrSQL_OrderBy
Dim i

'On Error Resume Next
	pstrSQL_Where = aryFilter(0)
	pstrSQL_Having = aryFilter(1)
	pstrSQL_OrderBy = aryFilter(2)

	pstrSQL = "SELECT ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCCreatedOn, Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount) AS CertificateValue, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, ssGiftCertificates.ssGCCustomerID" _
			& " FROM (((ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode) LEFT JOIN sfCustomers ON ssGiftCertificates.ssGCCustomerID = sfCustomers.custID) LEFT JOIN sfOrders ON ssGiftCertificateRedemptions.ssGCRedemptionOrderID = sfOrders.orderID) LEFT JOIN sfCustomers AS sourceCustomer ON sfOrders.orderCustId = sourceCustomer.custID" _
			& pstrSQL_Where _
			& " GROUP BY ssGiftCertificates.ssGCID, ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCIssuedToEmail, ssGiftCertificates.ssGCCreatedOn, ssGiftCertificateRedemptions.ssGCRedemptionActive, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, ssGiftCertificates.ssGCCustomerID" _
			& pstrSQL_Having _
			& pstrSQL_OrderBy
			
    Set pobjRS_Summary = CreateObject("adodb.Recordset")
    'Set pobjRS_Summary = GetRS(pstrSQL)
	With pobjRS_Summary
        .CursorLocation = 3 'adUseClient
	    
		If len(mlngMaxRecords) > 0 Then 
			.CacheSize = mlngMaxRecords
			.PageSize = mlngMaxRecords
		End If
		
		On Error Resume Next
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		
		If Err.number <> 0 Then
			If Err.number = 0 Then
				Response.Write "<h4><font color=red>Error " & Err.number & ": " & Err.Description & "</font></h4>"
			ElseIf Err.Description = "No current record." Then
				Response.Write "<h4><font color=green>Removing Orphaned GiftCertificates</font></h4>"
				Call RemoveOrphanedGiftCertificates
				.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
			Else
				Response.Write "<h4><font color=red>Error " & Err.number & ": " & Err.Description & "</font></h4>"
			End If
		End If
		
		mlngPageCount = .PageCount
		If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
	End With

	For i = 0 To 1
		If pobjRS_Summary.State = 1 Then
			If Not (pobjRS_Summary.EOF Or pobjRS_Summary.BOF) Then
				pstrssGCCode = pobjRS_Summary.Fields("ssGCCode").Value
				LoadAll = True
				Exit Function
			Else
				LoadAll = False
			End If
		Else
			'added to remove mismatched records
			If i = 0 Then
				Call RemoveOrphanedGiftCertificates
			Else
				Response.Write "<div class='FatalError'>You need to upgrade your database to use Gift Certificate Manager</div>" _
								& "<h3><a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=GiftCerificate'>Click here to upgrade</a></h3>"
			End If
		End If
	Next 'i
	LoadAll = False
	
End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(strssGCCode)

Dim sql
Dim rs

'On Error Resume Next

    sql = "Delete from ssGiftCertificates where ssGCCode = '" & strssGCCode & "'"
    pobjCnn.Execute sql, , 128

    sql = "Delete from ssGiftCertificateRedemptions where ssGCRedemptionCGCode = '" & strssGCCode & "'"
    pobjCnn.Execute sql, , 128

    If (Err.Number = 0) Then
        pstrMessage = "Record successfully deleted."
        Delete = True
    Else
        pstrMessage = Err.Description
        Delete = False
    End If

End Function    'Delete

'***********************************************************************************************

Public Function Update()

Dim sql
Dim objRS
Dim strErrorMessage
Dim blnAdd

'On Error Resume Next

    pblnError = False

    Call LoadFromRequest
    If ValidateValues Then
    
        If (Len(pstrssGCCode) = 0) Or (pstrssGCCode = "---")Then 
			pstrssGCCode = GetRandomCertificateNumber(plngssGCRedemptionOrderID)
			sql = "Select * from ssGiftCertificates where ssGCCode = '" & pstrssGCCode & "'"
			Set objRS = CreateObject("adodb.Recordset")
			objRS.open sql, pobjCnn, 1, 3
			blnAdd = True
			If objRS.EOF Then objRS.AddNew
        Else
			sql = "Select * from ssGiftCertificates where ssGCCode = '" & pstrssGCCode & "'"
			Set objRS = CreateObject("adodb.Recordset")
			objRS.open sql, pobjCnn, 1, 3
			If objRS.EOF Then 
				'sql = "Select * from ssGiftCertificates where ssGCCode = '" & GetRandomCertificateNumber(plngssGCRedemptionOrderID) & "'"
				'Set objRS = CreateObject("adodb.Recordset")
				'objRS.open sql, pobjCnn, 1, 3
				blnAdd = True
				objRS.addNew
				If objRS.EOF Then 
					pstrMessage = "Error creating Certificate number."
					pblnError = True
				End If
			Else
				blnAdd = False
			End If
        End If

        objRS.Fields("ssGCCode").Value = pstrssGCCode
		
        If isNull(objRS.Fields("ssGCCreatedOn").Value) Then objRS.Fields("ssGCCreatedOn").Value = Now()
        objRS.Fields("ssGCModifiedOn").Value = Now()

        If Len(pblnssGCElectronic) <> 0 Then 
            objRS.Fields("ssGCElectronic").Value = CBool(pblnssGCElectronic) * -1
        Else
            objRS.Fields("ssGCElectronic").Value = Null
        End If

        If Len(pblnssGCSingleUse) <> 0 Then 
            objRS.Fields("ssGCSingleUse").Value = CBool(pblnssGCSingleUse) * -1
        Else
            objRS.Fields("ssGCSingleUse").Value = Null
        End If

		With objRS
			.Fields("ssGCExpiresOn").Value = wrapSQLValue(pdtssGCExpiresOn, False, enDatatype_NA)
			.Fields("ssGCIssuedToEmail").Value = wrapSQLValue(pstrssGCIssuedToEmail, False, enDatatype_NA)
			.Fields("ssGCCustomerID").Value = wrapSQLValue(plngssGCCustomerID, False, enDatatype_NA)
		
			.Fields("ssGCFreeText").Value = wrapSQLValue(pstrssGCFreeText, False, enDatatype_NA)
			.Fields("ssGCToName").Value = wrapSQLValue(pstrssGCToName, False, enDatatype_NA)
			.Fields("ssGCFromName").Value = wrapSQLValue(pstrssGCFromName, False, enDatatype_NA)
			.Fields("ssGCFromEmail").Value = wrapSQLValue(pstrssGCFromEmail, False, enDatatype_NA)
			.Fields("ssGCMessage").Value = wrapSQLValue(pstrssGCMessage, False, enDatatype_NA)
		End With

        objRS.Update
        plngssGCID = objRS.Fields("ssGCID").Value

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The Certificate number you entered is already in use.<br />Please enter a different number.</H4><br />"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br />"
        End If
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = "The certificate was successfully added."
            Else
                pstrMessage = "The certificate was successfully updated."
            End If
        Else
            pblnError = True
        End If

        plngssGCID = objRS.Fields("ssGCID").Value
        objrs.Close
        Set objrs = Nothing
        
        'values may come in as an array
        Dim i
		Dim paryssGCRedemptionID
        
        If Len(plngssGCRedemptionID) = 0 Then
			paryssGCRedemptionID = Array ("")
        Else
			paryssGCRedemptionID = Split(plngssGCRedemptionID,",")
		End If
		
		For i = 0 To UBound(paryssGCRedemptionID)
		
			plngssGCRedemptionID = Trim(paryssGCRedemptionID(i))
			pbytssGCRedemptionType = Request.Form("ssGCRedemptionType." + plngssGCRedemptionID)
			pdblssGCRedemptionAmount = Request.Form("ssGCRedemptionAmount." + plngssGCRedemptionID)
			plngssGCRedemptionOrderID = Request.Form("ssGCRedemptionOrderID." + plngssGCRedemptionID)
			pstrssGCRedemptionInternalNotes = Request.Form("ssGCRedemptionInternalNotes." + plngssGCRedemptionID)
			pstrssGCRedemptionExternalNotes = Request.Form("ssGCRedemptionExternalNotes." + plngssGCRedemptionID)
			pblnssGCRedemptionActive = Request.Form("ssGCRedemptionActive." + plngssGCRedemptionID) = "ON"
			If Not IsNumeric(pdblssGCRedemptionAmount) And Len(pdblssGCRedemptionAmount) <> 0 Then
				pstrMessage = pstrMessage & "Please enter a number for the ssGCRedemptionAmount." & cstrDelimeter
			ElseIf Len(pdblssGCRedemptionAmount) = 0 Then
				pstrMessage = pstrMessage & "Please enter a value for the ssGCRedemptionAmount." & cstrDelimeter
			Else
				Call CreateRedemption(True, plngssGCRedemptionID, pstrssGCCode, pbytssGCRedemptionType, pdblssGCRedemptionAmount, pblnssGCRedemptionActive, plngssGCRedemptionOrderID, pstrssGCRedemptionInternalNotes, pstrssGCRedemptionExternalNotes)
			End If
		Next 'i

    End If

    Update = (not pblnError)

End Function    'Update

'***********************************************************************************************************************************

Public Function createCertificate(strOrderID, strssGCIssuedToEmail, blnssGCElectronic, dtssGCExpiresOn, blnssGCSingleUse, lngssGCCustomerID)

Dim sql
Dim objRS
Dim strErrorMessage

'On Error Resume Next

    pblnError = False

	pstrssGCCode = GetRandomCertificateNumber(strOrderID)
	sql = "Select * from ssGiftCertificates where ssGCCode = ''"
	Set objRS = CreateObject("adodb.Recordset")
	objRS.open sql, pobjCnn, 1, 3
	objRS.AddNew

	objRS.Fields("ssGCCode").Value = pstrssGCCode
		
    If Len(strssGCIssuedToEmail) <> 0 Then 
        objRS.Fields("ssGCIssuedToEmail").Value = strssGCIssuedToEmail
    Else
        objRS.Fields("ssGCIssuedToEmail").Value = Null
    End If
    
    objRS.Fields("ssGCCreatedOn").Value = Now()
    objRS.Fields("ssGCModifiedOn").Value = Now()

    If Len(blnssGCElectronic) <> 0 Then 
        objRS.Fields("ssGCElectronic").Value = blnssGCElectronic
    Else
        objRS.Fields("ssGCElectronic").Value = Null
    End If
    If Len(dtssGCExpiresOn) <> 0 Then 
        objRS.Fields("ssGCExpiresOn").Value = dtssGCExpiresOn
    Else
        objRS.Fields("ssGCExpiresOn").Value = Null
    End If

    If Len(blnssGCSingleUse) <> 0 Then 
        objRS.Fields("ssGCSingleUse").Value = blnssGCSingleUse
    Else
        objRS.Fields("ssGCSingleUse").Value = Null
    End If
    
    If Len(lngssGCCustomerID) <> 0 Then 
        objRS.Fields("ssGCCustomerID").Value = lngssGCCustomerID
    Else
        objRS.Fields("ssGCCustomerID").Value = Null
    End If

	objRS.Update
	plngssGCID = objRS.Fields("ssGCID").Value

End Function	'createCertificate

'***********************************************************************************************************************************

Private Function checkEmpty(byVal vntValue)

    If Len(vntValue) <> 0 Then 
        checkEmpty = vntValue
    Else
        checkEmpty = Null
    End If

End Function	'checkEmpty

'***********************************************************************************************************************************

Public Function createCertificate_New(strOrderID)

Dim sql
Dim objRS
Dim strErrorMessage

'On Error Resume Next

    pblnError = False

	pstrssGCCode = GetRandomCertificateNumber(strOrderID)
	sql = "Select * from ssGiftCertificates where ssGCCode = ''"
	Set objRS = CreateObject("adodb.Recordset")
	objRS.open sql, pobjCnn, 1, 3
	objRS.AddNew

	objRS.Fields("ssGCCode").Value = pstrssGCCode
		
	objRS.Fields("ssGCIssuedToEmail").Value = checkEmpty(pstrssGCIssuedToEmail)
	objRS.Fields("ssGCToName").Value = checkEmpty(pstrssGCToName)
	objRS.Fields("ssGCFromName").Value = checkEmpty(pstrssGCFromName)
	objRS.Fields("ssGCFromEmail").Value = checkEmpty(pstrssGCFromEmail)
	objRS.Fields("ssGCMessage").Value = checkEmpty(pstrssGCMessage)
	objRS.Fields("ssGCFreeText").Value = checkEmpty(pstrssGCFreeText)
	
    objRS.Fields("ssGCElectronic").Value = Abs(checkEmpty(pblnssGCElectronic))
    objRS.Fields("ssGCExpiresOn").Value = checkEmpty(pdtssGCExpiresOn)
    objRS.Fields("ssGCSingleUse").Value = checkEmpty(pblnssGCSingleUse)
    objRS.Fields("ssGCCustomerID").Value = checkEmpty(plngssGCCustomerID)

    objRS.Fields("ssGCCreatedOn").Value = Now()
    objRS.Fields("ssGCModifiedOn").Value = Now()
    
	objRS.Update
	plngssGCID = objRS.Fields("ssGCID").Value

End Function	'createCertificate_New

'***********************************************************************************************

Public Function CreateRedemption(blnComplete, lngssGCRedemptionID, strssGCRedemptionCGCode, bytssGCRedemptionType, dblssGCRedemptionAmount, blnssGCRedemptionActive, lngssGCRedemptionOrderID, strssGCRedemptionInternalNotes, strssGCRedemptionExternalNotes)


Dim sql
Dim objRS
Dim strErrorMessage
Dim plngTempID
Dim blnAdd

'On Error Resume Next

	If Len(strssGCRedemptionCGCode) = 0 Then strErrorMessage = "No Certificate number."
	If Len(bytssGCRedemptionType) = 0 Then strErrorMessage = "No certificate type."
	
	plngTempID = lngssGCRedemptionID
	If Len(plngTempID) = 0 Then plngTempID = 0

	If Len(strErrorMessage) = 0 Then
        
		sql = "Select * from ssGiftCertificateRedemptions where ssGCRedemptionID = " & plngTempID
		Set objRS = CreateObject("adodb.Recordset")
		objRS.open sql, pobjCnn, 1, 3
		If objRS.EOF Then
			objRS.AddNew
			blnAdd = True
		Else
			blnAdd = False
		End If

		objRS.Fields("ssGCRedemptionCGCode").Value = strssGCRedemptionCGCode
		objRS.Fields("ssGCRedemptionType").Value = bytssGCRedemptionType
			
		If Len(dblssGCRedemptionAmount) <> 0 Then 
			objRS.Fields("ssGCRedemptionAmount").Value = dblssGCRedemptionAmount
		Else
			objRS.Fields("ssGCRedemptionAmount").Value = Null
		End If

		If Len(lngssGCRedemptionOrderID) <> 0 Then 
			objRS.Fields("ssGCRedemptionOrderID").Value = lngssGCRedemptionOrderID
		Else
			objRS.Fields("ssGCRedemptionOrderID").Value = Null
		End If

        If Len(blnssGCRedemptionActive) <> 0 Then 
            objRS.Fields("ssGCRedemptionActive").Value = Abs(blnssGCRedemptionActive * -1)
        Else
            objRS.Fields("ssGCRedemptionActive").Value = False
        End If
        
		If blnComplete Then
			If Len(strssGCRedemptionExternalNotes) <> 0 Then 
				objRS.Fields("ssGCRedemptionExternalNotes").Value = strssGCRedemptionExternalNotes
			Else
				objRS.Fields("ssGCRedemptionExternalNotes").Value = Null
			End If
			If Len(strssGCRedemptionInternalNotes) <> 0 Then 
				objRS.Fields("ssGCRedemptionInternalNotes").Value = strssGCRedemptionInternalNotes
			Else
				objRS.Fields("ssGCRedemptionInternalNotes").Value = Null
			End If
		End If

		If isNull(objRS.Fields("ssGCRedemptionCreatedOn").Value) Then objRS.Fields("ssGCRedemptionCreatedOn").Value = Now()
		
		objRS.Update
		
		If Err.Number = -2147217887 Then
			If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
				strErrorMessage = "<H4>The data you entered is already in use.<br />Please enter a different data.</H4><br />"
				pblnError = True
			End If
		ElseIf Err.Number <> 0 Then
			Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br />"
		End If
	    
		plngTempID = objRS.Fields("ssGCRedemptionID").Value

		If Err.Number = 0 Then
			If blnAdd Then
				strErrorMessage = "The record was successfully added."
			Else
				strErrorMessage = "The record was successfully updated."
			End If
		Else
			pblnError = True
		End If
		
		objRS.Close
		Set objRS = Nothing

    End If
    
    pstrMessage = strErrorMessage

    CreateRedemption = CBool(Len(strErrorMessage) = 0)
    
    If Err.number <> 0 Then Err.Clear

End Function    'CreateRedemption

'***********************************************************************************************

Public Function setCertificateCustomerID(strssGCCode, lngssGCCustomerID)

Dim sql
Dim strErrorMessage

'On Error Resume Next

	If Len(strssGCCode) = 0 Then 
		strErrorMessage = "No certificate number."
		pblnError = True
	End If

	If Len(lngssGCCustomerID) = 0 Then 
		strErrorMessage = "No customer number."
		pblnError = True
	End If

	If Len(strErrorMessage) = 0 Then
        
		sql = "Update ssGiftCertificates Set ssGCCustomerID=" & lngssGCCustomerID & " where ssGCCode = '" & strssGCCode & "'"
		pobjCnn.Execute sql,,128
		
		If Err.Number <> 0 Then
			strErrorMessage = "Error: " & Err.Number & " - " & Err.Description & "<br />"
			pblnError = True
			Err.Clear
		Else
			strErrorMessage = "Certificate " & strssGCCode & " successfully updated."
			pblnError = False
		End If
	    
    End If
    
    pstrMessage = strErrorMessage

    setCertificateCustomerID = Not pblnError

End Function    'setCertificateCustomerID

'***********************************************************************************************

Public Function deleteRedemption(lngssGCRedemptionID)

Dim sql
Dim strErrorMessage

'On Error Resume Next

	If Len(lngssGCRedemptionID) = 0 Then 
		strErrorMessage = "No redemption id."
		pblnError = True
	End If

	If Len(strErrorMessage) = 0 Then
        
		sql = "Delete from ssGiftCertificateRedemptions where ssGCRedemptionID = " & lngssGCRedemptionID
		pobjCnn.Execute sql,,128
		
		If Err.Number <> 0 Then
			strErrorMessage = "Error: " & Err.Number & " - " & Err.Description & "<br />"
			pblnError = True
			Err.Clear
		Else
			strErrorMessage = "The redemption was successfully deleted."
		End If
	    
    End If
    
    pstrMessage = strErrorMessage

    deleteRedemption = pblnError

End Function    'deleteRedemption

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    If Len(pstrssGCCode) = 0 Then
        strError = strError & "Please enter a Certificate number." & cstrDelimeter
    End If
    
    If Not IsDate(pdtssGCExpiresOn) And Len(pdtssGCExpiresOn) <> 0 Then
        strError = strError & "Please enter a date for the expiration date." & cstrDelimeter
    End If

    pstrMessage = strError
    
    ValidateValues = (Len(strError) = 0)

End Function 'ValidateValues

'***********************************************************************************************************************************

Private Function RandomCertificate(strOrderID)

Dim pstrTemp

	Randomize()
	If cblnGCUseOrderNumber And Len(strOrderID) > 0 Then
		pstrTemp = strOrderID
		If Len(pstrTemp) < cbytGCLength Then pstrTemp = String(cbytGCLength - Len(pstrTemp),"0") & pstrTemp
	Else
		pstrTemp = CStr(Int(((clngGCMaxNumber - clngGCMinNumber + 1) * Rnd) + clngGCMinNumber))
		If Len(pstrTemp) < cbytGCLength Then pstrTemp = String(cbytGCLength - Len(pstrTemp),"0") & pstrTemp
	End If
	
	RandomCertificate = cstrGCPrefix & pstrTemp & cstrGCSuffix

End Function 'RandomCertificate

Private Function GetRandomCertificateNumber(strOrderID)

Dim p_objRS
Dim pstrSQL
Dim pstrCode
Dim pblnUnique
Dim i

	pblnUnique = False
	If cblnGCGenerateRandom OR cblnGCUseOrderNumber Then
		If cblnGCUseOrderNumber Then
			pstrCode = RandomCertificate(strOrderID)
		Else
			pstrCode = RandomCertificate("")
		End If
		pstrSQL = "Select ssGCID From ssGiftCertificates Where ssGCCode = '" & pstrCode & "'"
		Set p_objRS = CreateObject("adodb.Recordset")
		p_objRS.Open pstrSQL, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
		pblnUnique = p_objRS.EOF
	Else
		For i = 0 To 10
			pstrSQL = "Insert Into ssGiftCertificates ssGCCode='TempGCCode" & i & "'"
			
			On Error Resume Next
			pobjCnn.Execute pstrSQL,,128
			
			If Err.number = 0 Then
				pstrSQL = "Select ssGCID From ssGiftCertificates where ssGCCode='TempGCCode" & i & "'"
				Set p_objRS = CreateObject("adodb.Recordset")
				p_objRS.Open pstrSQL, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If Not p_objRS.EOF Then	
					pstrCode = RandomCertificate(p_objRS.Fields("ssGCID").Value)
					pstrSQL = "Update ssGiftCertificates Set ssGCCode='" & pstrCode & "' where ssGCID=TempGCCode" & p_objRS.Fields("ssGCID").Value
					pobjCnn.Execute pstrSQL,,128
					
				End If
			Else
				Err.Clear
			End If
			On Error Goto 0
			
			If Not pblnUnique Then Exit For
		Next 'i
	End If
	
	'if all else fails, generate a completely random code
	Do While Not pblnUnique
		pstrCode = RandomCertificate("")
		pstrSQL = "Select ssGCID From ssGiftCertificates Where ssGCCode = '" & pstrCode & "'"
		Set p_objRS = CreateObject("adodb.Recordset")
		p_objRS.Open pstrSQL, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
		pblnUnique = p_objRS.EOF
	Loop
		
	p_objRS.Close
	Set p_objRS = Nothing
	GetRandomCertificateNumber = pstrCode
	
	'Debugprint "pstrCode",pstrCode

End Function 'GetRandomCertificateNumber

'***********************************************************************************************************************************

Public Sub updateEmailSent(byVal blnSent)

Dim pstrSQL

	If Len(blnSent) = 0 Then
		pblnssGCElectronic = 0
	Else
		pblnssGCElectronic = CBool(blnSent)
	End If

	pstrSQL = "Update ssGiftCertificates Set ssGCElectronic=" & Abs(pblnssGCElectronic) & " Where ssGCID=" & plngssGCID
	'debugprint "pstrSQL", pstrSQL
	pobjCnn.Execute pstrSQL,,128
	
End Sub 'updateEmailSent

'***********************************************************************************************************************************

Public Function validateCertificate(ByVal strssGCRedemptionCGCode)

Dim pdtEndDate
Dim pblnValidCode
Dim pstrSQL
Dim p_objRS
Dim paryRedemptionCodes
Dim p_lngCounter
Dim pblnAtLeastOneValid

	If len(strssGCRedemptionCGCode) = 0 Then
		pstrMessage = "No certificate number was entered."
		validateCertificate = False
		Exit Function
	End If

	pblnExpired = False
	pdblCertificateValue = 0
	pblnAtLeastOneValid = False
	
	If mblnssDebug_GiftCertificate Then Response.Write "<fieldset><legend>Validating certificate code: <em>" & strssGCRedemptionCGCode & "</em></legend>"
	If Instr(1, strssGCRedemptionCGCode, ";") = 0 Then strssGCRedemptionCGCode = ";" & strssGCRedemptionCGCode & ";"
	paryRedemptionCodes = Split(strssGCRedemptionCGCode, ";")
	If mblnssDebug_GiftCertificate Then
		Response.Write "Certificate codes: " & strssGCRedemptionCGCode & "<br />"
		Response.Write "Certificate code count: " & UBound(paryRedemptionCodes) - 1 & "<br />"
	End If
	For p_lngCounter = 1 To UBound(paryRedemptionCodes) - 1
		pblnValidCode = True
		If len(paryRedemptionCodes(p_lngCounter)) = 0 Then
			pstrMessage = "No certificate number was entered."
			pblnValidCode = False
		Else
			pstrSQL = "SELECT ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionType" _
					& " FROM ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode" _
					& " WHERE ((ssGiftCertificateRedemptions.ssGCRedemptionCGCode='" & paryRedemptionCodes(p_lngCounter) & "') AND ((ssGiftCertificateRedemptions.ssGCRedemptionActive)<>0))"
			'Response.Write "validateCertificate SQL: " & pstrSQL & "<br />"
			Set p_objRS = CreateObject("ADODB.RECORDSET")
			With p_objRS
				.ActiveConnection = pobjCnn
				.CursorLocation = 2 'adUseClient
				.CursorType = 3 'adOpenStatic
				.LockType = 1 'adLockReadOnly
				.Source = pstrSQL
				.Open
			
				'Response.Write "validateCertificate .EOF: " & .EOF & "<br />"
				If mblnssDebug_GiftCertificate Then Response.Write "Database search result for <i>" & paryRedemptionCodes(p_lngCounter) & "</i>: " & CBool(Not .EOF) & "<br />"
				If .EOF Then
					pstrMessage = "An invalid certificate number was entered."
					pblnValidCode = False
				Else	'Valid Promotion Code so check for expiration
					If cblnCaseInsensitive Then
						If instr(1,LCase(Trim(.Fields("ssGCRedemptionCGCode").Value)),LCase(paryRedemptionCodes(p_lngCounter))) > 0 Then
							paryRedemptionCodes(p_lngCounter) = Trim(.Fields("ssGCRedemptionCGCode").Value)
						End If
					End If
					
					If mblnssDebug_GiftCertificate Then Response.Write "Check 2 for <i>" & paryRedemptionCodes(p_lngCounter) & "</i>: " & CBool(instr(1,Trim(.Fields("ssGCRedemptionCGCode").Value),Trim(paryRedemptionCodes(p_lngCounter)))) & "<br />"
					If instr(1,Trim(.Fields("ssGCRedemptionCGCode").Value),Trim(paryRedemptionCodes(p_lngCounter))) > 0 Then
						pdtEndDate = .Fields("ssGCExpiresOn").Value
						If len(pdtEndDate & "") > 0 Then
							If pdtEndDate < Now() Then 
								pstrMessage = "Certificate " & paryRedemptionCodes(p_lngCounter) & " has expired."
								pblnValidCode = False
								pblnExpired = True
							End If
						End If
						
						If Not pblnExpired Then
							Dim pblnUsed
							Dim i
							pblnUsed = False
							For i = 1 To .RecordCount
								pblnUsed = pblnUsed Or CBool(.Fields("ssGCRedemptionType").Value = 1)
								pdblCertificateValue = pdblCertificateValue + CDbl(.Fields("ssGCRedemptionAmount").Value)
								.MoveNext
							Next 'i
							.MoveFirst
							
							If .Fields("ssGCSingleUse").Value And pblnUsed Then
								pstrMessage = "Certificate previously used."
								pblnValidCode = False
							End If

							If pdblCertificateValue = 0 Then
								pstrMessage = "Certificate has no value remaining."
								pblnValidCode = False
							End If

						End If
					Else
						pstrMessage = "An invalid certificate number was entered."
						pblnValidCode = False
					End If
				End If	'.EOF
			End With	'p_objRS
		End If	'len(paryRedemptionCodes(p_lngCounter)) = 0
		If pblnValidCode Then pblnAtLeastOneValid = True
		If mblnssDebug_GiftCertificate Then Response.Write "Valid Code <i>" & paryRedemptionCodes(p_lngCounter) & "</i>: " & pblnValidCode & "<br />"
		If mblnssDebug_GiftCertificate Then Response.Write "AtLeastOneValid <i>" & paryRedemptionCodes(p_lngCounter) & "</i>: " & pblnAtLeastOneValid & "<br />"
	Next 'p_lngCounter
	If mblnssDebug_GiftCertificate Then Response.Write "</fieldset>"
	
	validateCertificate = pblnValidCode

End Function	'validateCertificate

'***********************************************************************************************************************************

Public Sub SendGCMail(ByVal sType, ByVal sInformation)

Dim parrInfo

	parrInfo = split(sInformation, "|")

	If Err.number <> 0 Then Err.Clear
	Call createMail(sType, sInformation)
	pstrMessage = pstrMessage & "<br />Email sent to " & parrInfo(0) & "." & cstrDelimeter
	If mblnssDebug_GiftCertificate Then Response.Write "<h3>SendGCMail: " & parrInfo(0) & "</h3>"

End Sub	'SendGCMail

'***********************************************************************************************************************************

Public Function verifyCertificate(ByVal strssGCRedemptionCGCode)

Dim pdtEndDate
Dim pblnValidCode
Dim pstrSQL
Dim p_objRS
Dim paryRedemptionCodes
Dim p_lngCounter

	If len(strssGCRedemptionCGCode) = 0 Then
		pstrMessage = "No certificate number was entered."
		verifyCertificate = False
		Exit Function
	End If


	pblnExpired = False
	pdblCertificateValue = 0
	pblnValidCode = True
	
	'Response.Write "<h3>strssGCRedemptionCGCode: " & strssGCRedemptionCGCode & "</h3>"
	If Instr(1, strssGCRedemptionCGCode, ";") = 0 Then strssGCRedemptionCGCode = ";" & strssGCRedemptionCGCode & ";"
	paryRedemptionCodes = Split(strssGCRedemptionCGCode, ";")
	For p_lngCounter = 1 To UBound(paryRedemptionCodes) - 1
	If len(paryRedemptionCodes(p_lngCounter)) = 0 Then
		pstrMessage = "No certificate number was entered."
		pblnValidCode = False
	Else
		pstrSQL = "SELECT ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificates.ssGCSingleUse, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionType" _
				& " FROM ssGiftCertificates LEFT JOIN ssGiftCertificateRedemptions ON ssGiftCertificates.ssGCCode = ssGiftCertificateRedemptions.ssGCRedemptionCGCode" _
				& " WHERE ((ssGiftCertificateRedemptions.ssGCRedemptionCGCode='" & paryRedemptionCodes(p_lngCounter) & "'))"
		'Response.Write "validateCertificate SQL: " & pstrSQL & "<br />"
		Set p_objRS = CreateObject("ADODB.RECORDSET")
		With p_objRS
			.ActiveConnection = pobjCnn
			.CursorLocation = 2 'adUseClient
			.CursorType = 3 'adOpenStatic
			.LockType = 1 'adLockReadOnly
			.Source = pstrSQL
			.Open
		
			'Response.Write "verifyCertificate .EOF: " & .EOF & "<br />"
			If .EOF Then
				pstrMessage = "An invalid certificate number was entered."
				pblnValidCode = False
			Else	'Valid Promotion Code so check for expiration
				If cblnCaseInsensitive Then
					If instr(1,LCase(Trim(.Fields("ssGCRedemptionCGCode").Value)),LCase(paryRedemptionCodes(p_lngCounter))) > 0 Then
						paryRedemptionCodes(p_lngCounter) = Trim(.Fields("ssGCRedemptionCGCode").Value)
					End If
				End If
				
				If instr(1,Trim(.Fields("ssGCRedemptionCGCode").Value),Trim(paryRedemptionCodes(p_lngCounter))) > 0 Then
					pdtEndDate = .Fields("ssGCExpiresOn").Value
					If len(pdtEndDate & "") > 0 Then
						If pdtEndDate < Now() Then 
							pstrMessage = "Certificate " & paryRedemptionCodes(p_lngCounter) & " has expired."
							pblnValidCode = False
							pblnExpired = True
						End If
					End If
					
					If Not pblnExpired Then
						Dim pblnUsed
						Dim i
						pblnUsed = False
						For i = 1 To .RecordCount
							pblnUsed = pblnUsed Or CBool(.Fields("ssGCRedemptionType").Value = 1)
							pdblCertificateValue = pdblCertificateValue + CDbl(.Fields("ssGCRedemptionAmount").Value)
							.MoveNext
						Next 'i
						.MoveFirst
						If .Fields("ssGCSingleUse").Value And pblnUsed Then
							pstrMessage = "Certificate previously used."
							pblnValidCode = False
						End If

						If pdblCertificateValue = 0 Then
							pstrMessage = "Certificate has no value remaining."
							pblnValidCode = False
						End If
					End If
				Else
					pstrMessage = "An invalid certificate number was entered."
					pblnValidCode = False
				End If
			End If
		End With
	End If
	Next 'p_lngCounter
	
	verifyCertificate = pblnValidCode

End Function	'verifyCertificate

'***********************************************************************************************

End Class   'clsPromotionssGiftCertificates

'************************************************************************************************

Function htmlOut(byVal strSource)

'Input: strSource in text format
'Output: string converted to HTML

Dim pstrTemp

	'Protect against nulls
	pstrTemp = Trim(strSource & "")
	
	pstrTemp = Replace(pstrTemp, vbcrlf, "<br />")
	
	htmlOut = pstrTemp

End Function	'htmlOut

'************************************************************************************************

Function customReplacements(ByVal strSource, ByRef objclsssGiftCertificate)

Dim p_strTemp

'On Error Resume Next

	p_strTemp = strSource

	p_strTemp = Replace(p_strTemp, "<CertificateValue>", FormatCurrency(objclsssGiftCertificate.CertificateValue,2))
	p_strTemp = Replace(p_strTemp, "<ssGCCode>", objclsssGiftCertificate.ssGCCode)
	p_strTemp = Replace(p_strTemp, "<ssGCExpiresOn>", Trim(objclsssGiftCertificate.ssGCExpiresOn & ""))

	p_strTemp = Replace(p_strTemp, "<ssGCToName>", Trim(objclsssGiftCertificate.ssGCToName & ""))
	p_strTemp = Replace(p_strTemp, "<ssGCToEmail>", Trim(objclsssGiftCertificate.ssGCIssuedToEmail & ""))
	p_strTemp = Replace(p_strTemp, "<ssGCFromName>", Trim(objclsssGiftCertificate.ssGCFromName & ""))
	p_strTemp = Replace(p_strTemp, "<ssGCFromEmail>", Trim(objclsssGiftCertificate.ssGCFromEmail & ""))
	p_strTemp = Replace(p_strTemp, "<ssGCMessage>", Trim(objclsssGiftCertificate.ssGCMessage & ""))
	p_strTemp = Replace(p_strTemp, "<ssGCFreeText>", Trim(objclsssGiftCertificate.ssGCFreeText & ""))
	
	customReplacements = p_strTemp

End Function	'customReplacements

'***********************************************************************************************

Function hasActiveCertificates()

Dim pobjRS
Dim pstrSQL

'On Error Resume Next

	If cblnSQLDatabase Then
		pstrSQL = "SELECT Top 1 ssGiftCertificates.ssGCCode" _
				& " FROM ssGiftCertificateRedemptions RIGHT JOIN ssGiftCertificates ON ssGiftCertificateRedemptions.ssGCRedemptionCGCode = ssGiftCertificates.ssGCCode" _
				& " GROUP BY ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
				& " HAVING ((ssGiftCertificates.ssGCExpiresOn>Now()) AND (ssGiftCertificateRedemptions.ssGCRedemptionActive<>0) AND (Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount)>0))"
	Else
		pstrSQL = "SELECT Top 1 ssGiftCertificates.ssGCCode" _
				& " FROM ssGiftCertificateRedemptions RIGHT JOIN ssGiftCertificates ON ssGiftCertificateRedemptions.ssGCRedemptionCGCode = ssGiftCertificates.ssGCCode" _
				& " GROUP BY ssGiftCertificates.ssGCCode, ssGiftCertificates.ssGCExpiresOn, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
				& " HAVING ((ssGiftCertificates.ssGCExpiresOn>Now()) AND (ssGiftCertificateRedemptions.ssGCRedemptionActive<>0) AND (Sum(ssGiftCertificateRedemptions.ssGCRedemptionAmount)>0))"
	End If
	
    Set pobjRS = CreateObject("adodb.Recordset")
	pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
	hasActiveCertificates = Not pobjRS.EOF
	pobjRS.Close
	Set pobjRS = Nothing

End Function    'hasActiveCertificates

'************************************************************************************************

Const enGC_ToName = 0
Const enGC_ToEmail = 1
Const enGC_FromName = 2
Const enGC_FromEmail = 3
Const enGC_CertificateType = 4
Const enGC_Message = 5
Const enGC_CertificateOrCredit = 6

Dim mblnssDebug_GiftCertificate
mblnssDebug_GiftCertificate = CBool(Len(Session("ssDebug_GiftCertificate")) > 0)
If mblnssDebug_GiftCertificate Then Response.Write "<h4>Gift Certificate debugging is on</h4>"
%>
