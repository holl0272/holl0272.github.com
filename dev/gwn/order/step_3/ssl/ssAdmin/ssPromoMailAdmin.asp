<%Option Explicit
'********************************************************************************
'*   Promotional Mail Manager							                        *
'*   Release Version:	1.00.005		                                        *
'*   Release Date:		September 21, 2002										*
'*   Revision Date:		October 28, 2004										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   1.00.006 Notes (June 3, 2005)			                                    *
'*     - Feature - Added ability to import mass customer lists					*
'*                                                                              *
'*   1.00.005 Notes (October 28, 2004)		                                    *
'*     - Feature - Added pop-up window to display results						*
'*     - Feature - Added ability to auto-select email addresses					*
'*                                                                              *
'*   1.00.004 Notes (September 24, 2003)                                        *
'*     - Feature - Added to product purchased filter							*
'*                                                                              *
'*   1.0.3 Notes (February 24, 2003)                                            *
'*     - Feature - "None" option added to Pricing Level filter                  *
'*                                                                              *
'*   1.0.2 Notes:					                                            *
'*     - Feature - added test mail function                                     *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = False

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Server.ScriptTimeout = 900		'in seconds. Adjust for large databases. Default is usually 90 seconds, most hosts limit it to 900 seconds

Const clngMaxEmailsToSendAtOneTime = 100
Const cbytDefault_UseUniqueEmails = 1
Const cbytDefault_AutoSelectCustomers = 1
Const cstrOrderBy = "sfCustomers.custEmail"	'sfCustomers.custEmail OR sfCustomers.custLastName
Const clngPageRefreshTime = 3000	'time in milliseconds
Const cbytTimePadding = 15	'time in seconds

'/
'/////////////////////////////////////////////////

Class clsCustomer

Private plngID
Private pstrEmail
Private pstrFirstName
Private pstrLastName
Private pstrDateLastOrder
Private pdblTotalOrders
Private plngOrderCount
Private pdblMaxOrder
Private pbytSubscribed
Private plngPrevOrderID
Private pblnDebug

	Private Sub class_Initialize()
	    pdblTotalOrders = 0
	    pdblMaxOrder = 0
	    plngOrderCount = 0
	    pbytSubscribed = 1
	    pblnDebug = False
	End Sub

	Public Sub AddOrder(objRS)

	Dim pstrTempDate
	Dim pdblOrderAmount

		With objRS
			plngID = .Fields("custID").Value
			pstrEmail = .Fields("custEmail").Value
			pstrFirstName = .Fields("custFirstName").Value
			pstrLastName = .Fields("custLastName").Value
			pstrTempDate = Trim(.Fields("orderDate").Value & "")
			pdblOrderAmount = Trim(.Fields("orderGrandTotal").Value & "")
			pbytSubscribed = pbytSubscribed * .Fields("custIsSubscribed").Value
			
			If plngPrevOrderID <> .Fields("orderDate").Value Then
				plngPrevOrderID = .Fields("orderDate").Value
				If Len(pdblOrderAmount) > 0 Then
					pdblOrderAmount = CDbl(pdblOrderAmount)
					plngOrderCount = plngOrderCount + 1
				Else
					pdblOrderAmount = 0
				End If
					
				pdblTotalOrders = pdblTotalOrders + pdblOrderAmount
					
				If pdblOrderAmount > pdblMaxOrder Then pdblMaxOrder = pdblOrderAmount
			End If

		End With
		
		If Len(pstrTempDate) > 0 Then
			If Len(pstrDateLastOrder) = 0 Then
				pstrDateLastOrder = pstrTempDate
			ElseIf CDate(pstrDateLastOrder) > CDate(pstrTempDate) Then
				pstrDateLastOrder = pstrTempDate
			End If
		End If
			
	End Sub

	Public Property Get ID
		ID = plngID
	End Property

	Public Property Get WholeName
		WholeName = pstrLastName & ", " & pstrFirstName
	End Property

	Public Property Get OrderCount
		OrderCount = plngOrderCount
	End Property

	Public Property Get FirstName
		FirstName = pstrFirstName
	End Property

	Public Property Get LastName
		LastName = pstrLastName
	End Property

	Public Property Get Email
		Email = pstrEmail
	End Property

	Public Property Get DateLastOrder
		If Len(pstrDateLastOrder) = 0 Then
			DateLastOrder = "-"
		Else
			DateLastOrder = FormatDateTime(pstrDateLastOrder,2)
		End If
	End Property

	Public Property Get MaxOrder
		MaxOrder = pdblMaxOrder
	End Property

	Public Property Get TotalOrders
		TotalOrders = pdblTotalOrders
	End Property

	Public Property Get Subscribed
		Subscribed = pbytSubscribed
	End Property

End Class	'clsCustomer


Class clsPromoMail
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pblnError
Private pblnDebug

Private prsOrders
Private prsOrderSummaries

'External Info
Private plngTotalMailToSend
Private plngNumEmailsSent
Private pblnUniqueEmailAddressesOnly
Private pblnAutoSelectCustomers

'ssorderID
'ssExternalNotes
'ssInternalNotes
'ssDatePaymentReceived
'ssDateOrderShipped
'ssShippedVia
'ssTrackingNumber
'ssOrderStatus
'ssDateEmailSent

Private plngssorderID
Private pstrssExternalNotes
Private pstrssInternalNotes
Private pdtssDatePaymentReceived
Private pdtssDateOrderShipped
Private pstrssPaidVia
Private pbytssShippedVia
Private pstrssTrackingNumber
Private pbytssOrderStatus
Private pdtssDateEmailSent

'Order Sent Email Parameters
Private pstrEmailTo
Private pstrEmailSubject
Private pstrEmailBody
Private pblnSendMail
Private pstrResults

Private pstrEmailFrom
Private pstrEmailFromName
Private pstrMailMethod
Private pstrMailServer
Private pstrSiteURL

Private pblnMailingInProcessChecked
Private plngMailingInProcess

Private enPromoMailStatus_ToSend
Private enPromoMailStatus_Fail
Private enPromoMailStatus_Success
Private enPromoMailStatus_Complete

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
    Call LoadDefaultEmailSettings
    pblnMailingInProcessChecked = False
	pblnDebug = False
	pblnUniqueEmailAddressesOnly = True
	pblnAutoSelectCustomers = False
	
	enPromoMailStatus_Complete = ""
	enPromoMailStatus_Fail = "0"
	enPromoMailStatus_Success = "1"
	enPromoMailStatus_ToSend = "2"
	
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	Call ReleaseObject(prsOrders)
	Call ReleaseObject(prsOrderSummaries)
End Sub

'***********************************************************************************************

	Public Property Get MailingInProcess()
		If Not pblnMailingInProcessChecked Then CheckForExistingMailing
	    MailingInProcess = plngMailingInProcess
	End Property

	Public Property Get EmailFrom()
	    EmailFrom = pstrEmailFrom
	End Property

	Public Property Let EmailFrom(strEmailFrom)
	    pstrEmailFrom = strEmailFrom
	End Property

	Public Property Let EmailFromName(strEmailFromName)
	    pstrEmailFromName = strEmailFromName
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

	Public Property Let UniqueEmailAddressesOnly(blnValue)
	    pblnUniqueEmailAddressesOnly = blnValue
	End Property

	Public Property Let AutoSelectCustomers(blnValue)
	    pblnAutoSelectCustomers = blnValue
	End Property

	Public Property Get OutputResults()
	    OutputResults = pstrResults
	End Property

'***********************************************************************************************

	Private Sub LoadFromRequest

	    With Request.Form
			plngssorderID = Trim(.Item("orderID"))
			pstrssExternalNotes = Trim(.Item("ssExternalNotes"))
			pstrssInternalNotes = Trim(.Item("ssInternalNotes"))
			pdtssDatePaymentReceived = Trim(.Item("ssDatePaymentReceived"))
			pdtssDateOrderShipped = Trim(.Item("ssDateOrderShipped"))
			pstrssPaidVia = Trim(.Item("ssPaidVia"))
			pbytssShippedVia = Trim(.Item("ssShippedVia"))
			pstrssTrackingNumber = Trim(.Item("ssTrackingNumber"))
			pbytssOrderStatus = Trim(.Item("ssOrderStatus"))
			
			pstrEmailTo = Trim(.Item("EmailTo"))
			pstrEmailSubject = Trim(.Item("EmailSubject"))
			pstrEmailBody = Trim(.Item("EmailBody"))
			pblnSendMail = Trim(.Item("SendEmail")) = "1"

	    End With

	End Sub 'LoadFromRequest

	'***********************************************************************************************

	Public Function isDBUpgraded

	Dim pobjRS
	Dim pstrSQL
	Dim pblnDBUpgraded

	On Error Resume Next

		pstrSQL = "SELECT ssPromoMailSent FROM sfCustomers Where custID=-1"

		Set	pobjRS = server.CreateObject("adodb.recordset")
		With pobjRS
	        .CursorLocation = 3 'adUseClient
	        .CursorType = 3 'adOpenStatic
	        .LockType = 1 'adLockReadOnly
			.Open pstrSQL, cnn

			If Err.number <> 0 Then
				debugprint "pstrSQL",pstrSQL
				pstrMessage = "Loading Error " & Err.number & ": " & Err.Description
				Err.Clear
				pblnDBUpgraded = False
				CheckForPriorMailing = False
			Else
				pblnDBUpgraded = True
			End If
		
		End With
		Call ReleaseObject(pobjRS)
		
		isDBUpgraded = pblnDBUpgraded

	End Function    'isDBUpgraded

	'***********************************************************************************************

	Public Function LoadOrderSummaries(strSQLParmeters)

	dim pstrSQL
	dim p_strWhere
	dim p_strHaving
	Dim parySQLParameter
	dim i
	dim sql
	Dim p_strJoinType	
	
	parySQLParameter = Split(strSQLParmeters,"|")
	p_strWhere = parySQLParameter(0)
	p_strHaving = parySQLParameter(1)
	p_strJoinType = parySQLParameter(2)
	
		If pblnAutoSelectCustomers Then
			If pblnUniqueEmailAddressesOnly Then
				pstrSQL = "SELECT Distinct sfCustomers.custEmail, sfCustomers.custID" _
						& " FROM sfCustomers " & p_strJoinType & " JOIN (sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCustomers.custID = sfOrders.orderCustId" _
						& p_strWhere _
						& " ORDER BY sfCustomers.custEmail DESC"
			Else
				pstrSQL = "SELECT sfCustomers.custEmail, sfCustomers.custID" _
						& " FROM sfCustomers " & p_strJoinType & " JOIN (sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCustomers.custID = sfOrders.orderCustId" _
						& p_strWhere _
						& " ORDER BY " & cstrOrderBy & ", sfOrders.orderDate DESC"
			End If	'pblnUniqueEmailAddressesOnly
		Else
			pstrSQL = "SELECT sfCustomers.custID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custIsSubscribed, sfCustomers.custEmail, sfOrders.orderDate, sfOrders.orderGrandTotal" _
					& " FROM sfCustomers " & p_strJoinType & " JOIN (sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCustomers.custID = sfOrders.orderCustId" _
					& p_strWhere _
					& " ORDER BY " & cstrOrderBy & ", sfOrders.orderDate DESC"
		End If	'pblnAutoSelectCustomers
		
		On Error Resume Next

		Set	prsOrderSummaries = server.CreateObject("adodb.recordset")
		With prsOrderSummaries
	        .CursorLocation = 3 'adUseClient
	        .CursorType = 3 'adOpenStatic
	        .LockType = 1 'adLockReadOnly
			.Open pstrSQL, cnn
			'debugprint "pstrSQL",pstrSQL

			If Err.number <> 0 Then
				debugprint "pstrSQL",pstrSQL
				pstrMessage = "Loading Error " & Err.number & ": " & Err.Description
				Err.Clear
				LoadOrderSummaries = False
				If Response.Buffer Then Response.Flush
				Exit Function
			End If
			
			LoadOrderSummaries = (Not .EOF)
			
		End With

	End Function    'LoadOrderSummaries

	'***********************************************************************************************

	Private Function CheckForExistingMailing()

	dim pstrSQL
	Dim prsEmailList

'	On Error Resume Next

		pstrSQL = "SELECT Count(ssPromoMailSent) As MailingInProcess FROM sfCustomers Where ssPromoMailSent = '" & enPromoMailStatus_ToSend & "'"
		set	prsEmailList = server.CreateObject("adodb.recordset")
		With prsEmailList
	        .CursorLocation = 3 'adUseClient
	        .CursorType = 3 'adOpenStatic
	        .LockType = 1 'adLockReadOnly
			.Open pstrSQL, cnn
			
			If .EOF Then
				plngMailingInProcess = 0
			Else
				plngMailingInProcess = ABS(.Fields("MailingInProcess").Value)
			End If

			.Close
		End With
		set	prsEmailList = Nothing

	End Function    'CheckForExistingMailing

	'***********************************************************************************************

	Public Sub Subscribe(byVal blnSubscribe)
	
	Dim i
	Dim plngCount
	Dim paryCustIDs
	Dim pobjRS
	Dim pstrEmails, paryEmail
	Dim pstrCustEmail
	Dim pstrTempResult
	Dim pstrSQL
	Dim pstrSQLWhere
	
'	On Error Resume Next

		paryCustIDs = Split(Request.Form("custID"),",")
		If isArray(paryCustIDs) Then


			pstrsqlWhere = "sfCustomers.custID In (" & wrapSQLValue(paryCustIDs(0), False, enDatatype_number)
			For i = 1 To UBound(paryCustIDs)
				pstrsqlWhere = pstrsqlWhere & ", " & wrapSQLValue(paryCustIDs(i), False, enDatatype_number)
			Next 'i
			pstrsqlWhere = pstrsqlWhere & ")"
			pstrSQL = "Update sfCustomers Set custIsSubscribed = " & Abs(blnSubscribe) & " Where " & pstrsqlWhere
			cnn.Execute pstrSQL,,128

			pstrSQL = "Select custEmail From sfCustomers Where " & pstrsqlWhere
			Set pobjRS = GetRS(pstrSQL)
			If Not pobjRS.EOF Then
				If blnSubscribe Then
					Response.Write "<fieldset><legend>The following customers have been subscribed</legend>"
				Else
					Response.Write "<fieldset><legend>The following customers have been unsubscribed</legend>"
				End If
				Response.Write pobjRS.GetString(2, , "", "<br />")
				Response.Write "</fieldset>"
			End If
			Call ReleaseObject(pobjRS)

		End If
		
	End Sub	'Subscribe

	'***********************************************************************************************

	Private Sub autoSetMailList
	
	Dim i
	Dim pobjDic
	Dim pstrKey
	Dim plngCount
	Dim pstrSQL
	Dim pstrWhere
	Dim plngCustID
	
		plngCount = 0
		plngTotalMailToSend = 0
		
		Call LoadOrderSummaries(SummaryFilter)
		Set pobjDic = Server.CreateObject("Scripting.Dictionary")
		For i = 1 To prsOrderSummaries.RecordCount
			plngCustID = prsOrderSummaries.Fields("custID").Value
			If pblnUniqueEmailAddressesOnly Then
				pstrKey = prsOrderSummaries.Fields("custEmail").Value
			Else
				pstrKey = "key_" & plngCustID
			End If

			If Not pobjDic.Exists(pstrKey) Then
				pobjDic.Add pstrKey, plngCustID
				
				If Len(pstrWhere) = 0 Then
					pstrWhere = " Where custID=" & Trim(plngCustID)
				Else
					pstrWhere = pstrWhere & " OR custID=" & Trim(plngCustID)
				End If
				plngCount = plngCount + 1
				plngTotalMailToSend = plngTotalMailToSend + 1
				
				If plngCount > 50 Then
					cnn.Execute "Update sfCustomers Set ssPromoMailSent = '" & enPromoMailStatus_ToSend & "' " & pstrWhere,,128
					plngCount = 0
					pstrWhere = ""
				End If
			End If
			
			prsOrderSummaries.MoveNext
		Next
		If Len(pstrWhere) > 0 Then cnn.Execute "Update sfCustomers Set ssPromoMailSent = '" & enPromoMailStatus_ToSend & "' " & pstrWhere,,128
	
	End Sub	'autoSetMailList

	'***********************************************************************************************

	Public Sub setMailList
	
	Dim i
	Dim plngCount
	Dim paryCustIDs
	Dim pstrSQL
	Dim pstrWhere
	
'	On Error Resume Next

		If ValidateValues Then
		
			If pblnAutoSelectCustomers Then
				Call autoSetMailList
			Else
				paryCustIDs = Split(Request.Form("custID"),",")
				If isArray(paryCustIDs) Then
					'set flags in database
					plngCount = 0
					plngTotalMailToSend = UBound(paryCustIDs) + 1
					cnn.Execute "Update sfCustomers Set ssPromoMailSent = '" & enPromoMailStatus_Complete & "'",,128

					For i = 0 To UBound(paryCustIDs)
						If Len(pstrWhere) = 0 Then
							pstrWhere = " Where custID=" & Trim(paryCustIDs(i))
						Else
							pstrWhere = pstrWhere & " OR custID=" & Trim(paryCustIDs(i))
						End If
						plngCount = plngCount + 1
						
						If plngCount > 50 Then
							cnn.Execute "Update sfCustomers Set ssPromoMailSent = '" & enPromoMailStatus_ToSend & "' " & pstrWhere,,128
							plngCount = 0
							pstrWhere = ""
						End If
					Next 'i
					If Len(pstrWhere) > 0 Then cnn.Execute "Update sfCustomers Set ssPromoMailSent = '" & enPromoMailStatus_ToSend & "' " & pstrWhere,,128
				Else
					strError = strError & "Please select a customer to send email to" & cstrDelimeter
				End If	'isArray(paryCustIDs)
			End If	'pblnAutoSelectCustomers

			Session("SendingEmail") = CStr(plngTotalMailToSend)

		End If	'ValidateValues
		
	End Sub	'setMailList

	'***********************************************************************************************

	Public Sub SendPromotionalMailing()
	
	Dim i
	Dim pstrEmailTo
	Dim pstrEmailBody
	Dim pstrEmailSubject
	Dim pstrchkEmails, parychkEmail
	Dim pstrEmails, paryEmail
	Dim pstrFirstNames, paryFirstNames
	Dim pstrLastNames, paryLastNames
	Dim pstrCustFirstName
	Dim pstrCustLastName
	Dim pstrCustEmail
	Dim pstrTempResult
	Dim pblnMailingComplete
	Dim pstrLocalResult
	Dim pblnSetNewMailing
	Dim pdtStartTime
	Dim ptmMaxTime
	Dim pElapsedTime
	
'	On Error Resume Next

		pdtStartTime = Now()
		ptmMaxTime = Server.ScriptTimeout - cbytTimePadding

		pblnSetNewMailing = (Len(Session("SendingEmail")) = 0) Or (Len(Request.Form("chkMailingInProcess")) = 0)
		
		If Request.Form("chkMailingInProcess") = 1 Then
			'no need to do anything
			Call WriteToOutputWindow("<h4>Continuing email . . .</h4>")
		ElseIf pblnSetNewMailing Then
			Call setMailList
			Call WriteToOutputWindow("<h4>Identifying email recipients . . .</h4>")
		End If

		If ValidateValues() Then

			'open the database to find the email addresses
			Dim prsMailing
			Set prsMailing = Server.CreateObject("ADODB.RECORDSET")
			'prsMailing.CursorLocation = 2 'adUseServer
			prsMailing.CursorLocation = 3 'adUseClient
			prsMailing.PageSize = clngMaxEmailsToSendAtOneTime
			prsMailing.Open "Select custID, custFirstName, custLastName, custEmail from sfCustomers where ssPromoMailSent = '" & enPromoMailStatus_ToSend & "' Order By " & cstrOrderBy, cnn, 3, 1    'adOpenForwardOnly, adLockReadOnly
			
			If Not prsMailing.EOF Then

				pstrResults = "<h4>Email was sent to the following people:</h4>"
				Call WriteToOutputWindow(pstrResults)
				
				plngTotalMailToSend = prsMailing.RecordCount
				
				For i = 1 to plngTotalMailToSend
					pstrCustEmail = Trim(prsMailing.Fields("custEmail").Value)
					pstrCustFirstName = Trim(prsMailing.Fields("custFirstName").Value & "")
					pstrCustLastName = Trim(prsMailing.Fields("custLastName").Value & "")

					pstrEmailSubject = MakeSubstitutions(mstrEmailSubject, pstrCustFirstName, pstrCustLastName, pstrCustEmail, False)
					pstrEmailBody = MakeSubstitutions(mstrEmailBody, pstrCustFirstName, pstrCustLastName, pstrCustEmail, mblnHTMLEmail)
					
					If pblnDebug Then
						Response.Write "<fieldset><legend>Email " & i & " - " & pstrCustFirstName & "</legend>"
						Response.Write pstrEmailSubject & "<hr />"
						Response.Write pstrEmailBody
						Response.Write "</fieldset>"
					End If
					
					pstrTempResult = ssSendMail(pstrEmailSubject, pstrEmailBody, pstrCustEmail, mblnHTMLEmail)
					If Len(pstrTempResult) = 0 Then
						pstrLocalResult = "<li>" & pstrCustLastName & ", " & pstrCustFirstName & ": " & pstrCustEmail & "</li>"
						pstrResults = pstrResults & pstrLocalResult
						cnn.Execute "Update sfCustomers Set ssPromoMailSent = '" & enPromoMailStatus_Success & "' Where custID=" & prsMailing.Fields("custID").Value,,128
					Else
						pstrLocalResult = "<li><font color=red>Error sending to " & pstrCustLastName & ", " & pstrCustFirstName & ": " & pstrCustEmail & " - " & pstrTempResult & "</font>" & "</li>"
						pstrResults = pstrResults & pstrLocalResult
						cnn.Execute "Update sfCustomers Set ssPromoMailSent = '" & enPromoMailStatus_Fail & "' Where custID=" & prsMailing.Fields("custID").Value,,128
					End If
					Call WriteToOutputWindow(pstrLocalResult)

					plngNumEmailsSent = plngNumEmailsSent + 1
					prsMailing.MoveNext
					
					pElapsedTime = DateDiff("s", pdtStartTime, Now())

					If i >= clngMaxEmailsToSendAtOneTime Or pElapsedTime > ptmMaxTime Then
						plngNumEmailsSent = CLng(Session("SendingEmail")) - prsMailing.RecordCount + plngNumEmailsSent
						If i >= clngMaxEmailsToSendAtOneTime Then
							pstrResults = pstrResults & "<h3>" & plngNumEmailsSent & " of " & Session("SendingEmail") & " emails sent. Mailing suspended due to limit of " & clngMaxEmailsToSendAtOneTime & " emails.</h3>"
						Else
							pstrResults = pstrResults & "<h3>" & plngNumEmailsSent & " of " & Session("SendingEmail") & " emails sent. Mailing suspended due to script timeout limitation (" & Server.ScriptTimeout & " seconds).</h3>"
						End If

						pstrResults = pstrResults & "<input type='hidden' name='mailContinue' id='mailContinue' value='1'>"
						pstrResults = pstrResults & "<script language='javascript'>document.frmData.Action.value='Sendmail'; window.setTimeout ('document.frmData.submit();'," & clngPageRefreshTime & ");</script>"
						pstrResults = pstrResults & "This page will automatically resubmit in " & FormatNumber(clngPageRefreshTime/1000, 2) & " seconds. Press here to continue <input type='button' name='btnContinueMailing' id='btnContinueMailing' value='immediately' onclick='SendEmail();'>"
						plngMailingInProcess = 1
						Exit For
					End If
				Next 'i
				
			End If
			
			pblnMailingComplete = prsMailing.EOF
			prsMailing.Close
			Set prsMailing = Nothing
			
			If pblnMailingComplete Then
			
			Dim pstrSuccessfulEmails
			Dim pstrUnsuccessfulEmails
			Dim plngSuccessfulEmails
			Dim plngUnsuccessfulEmails
			
				plngSuccessfulEmails = 0
				plngUnsuccessfulEmails = 0
			
				Set prsMailing = Server.CreateObject("ADODB.RECORDSET")
				prsMailing.CursorLocation = 2 'adUseClient
				prsMailing.Open "Select custID, custFirstName, custLastName, custEmail, ssPromoMailSent from sfCustomers where ssPromoMailSent <> '" & enPromoMailStatus_Complete & "' Order By " & cstrOrderBy, cnn, 3, 1    'adOpenForwardOnly, adLockReadOnly
				For i = 1 To prsMailing.RecordCount
					pstrCustEmail = Trim(prsMailing.Fields("custEmail").Value)
					pstrCustFirstName = Trim(prsMailing.Fields("custFirstName").Value & "")
					pstrCustLastName = Trim(prsMailing.Fields("custLastName").Value & "")
					If Trim(prsMailing.Fields("ssPromoMailSent").Value & "") = enPromoMailStatus_Success Then
						plngSuccessfulEmails = plngSuccessfulEmails + 1
						pstrLocalResult = "<li>" & pstrCustLastName & ", " & pstrCustFirstName & ": " & pstrCustEmail & "</li>"
						pstrSuccessfulEmails = pstrSuccessfulEmails & pstrLocalResult
					Else
						plngUnsuccessfulEmails = plngUnsuccessfulEmails + 1
						pstrLocalResult = "<li><font color=red>" & plngUnsuccessfulEmails & ") Error sending to " & pstrCustLastName & ", " & pstrCustFirstName & ": " & pstrCustEmail & " - " & pstrTempResult & "</font></li>"
						pstrUnsuccessfulEmails = pstrUnsuccessfulEmails & pstrLocalResult
					End If
					'Call WriteToOutputWindow(pstrLocalResult)
					prsMailing.MoveNext
				Next 'i
				prsMailing.Close
				Set prsMailing = Nothing
				
				If Len(pstrSuccessfulEmails) > 0 Then pstrSuccessfulEmails = "<ol>" & pstrSuccessfulEmails & "</ol>"
				If Len(pstrUnsuccessfulEmails) > 0 Then pstrUnsuccessfulEmails = "<ol>" & pstrUnsuccessfulEmails & "</ol>"
				
				cnn.Execute "Update sfCustomers Set ssPromoMailSent = '" & enPromoMailStatus_Complete & "'",,128
				
				pstrResults = "<fieldset id=promoMailResults><legend>Email Results</legend><strong>" & plngSuccessfulEmails & " emails successfully sent</strong>" & pstrSuccessfulEmails _
							& "<hr><strong>" & plngUnsuccessfulEmails & " emails unsuccessfully sent</strong>" & pstrUnsuccessfulEmails _
							& "<hr><h4>" & Session("SendingEmail") & " mails attempted. Mailing Complete</h4></fieldset>"	'& "<script language=javascript>ScrollToElem('promoMailResults');</script>"
				Call WriteToOutputWindow(pstrResults)
				
				Session("SendingEmail") = ""
			End If

		End If

	End Sub	'SendPromotionalMailing

	'***********************************************************************************************

	Private Function MakeSubstitutions(ByVal strEmailBody, ByVal strCustFirstName, ByVal strCustLastName, ByVal strEmail, ByVal blnHTML)

	Dim p_strBody
	Dim pstrCustFirstName
	Dim pstrCustLastName
	Dim pstrCustName
	Dim pstrUnsubscribeURL
	
	'On Error Resume Next

			pstrCustFirstName = strCustFirstName
			pstrCustLastName = strCustLastName
			pstrCustName = pstrCustFirstName & " " & pstrCustLastName
			p_strBody = strEmailBody
			pstrUnsubscribeURL = pstrSiteURL & "unsubscribe.asp?email=" & strEmail

			'HTML check removed since new version isn't encoding HTML
			'ALL Uppercase added since HTML editor is automatically changing case

			'If blnHTML Then
				pstrUnsubscribeURL = "<a href='" & pstrUnsubscribeURL & "'>here</a>"
				p_strBody = Replace(p_strBody,"&lt;customerFirstName&gt;",pstrCustFirstName)	' - this is the customer's first name
				p_strBody = Replace(p_strBody,"&lt;customerLastName&gt;",pstrCustLastName)	' - this is the customer's last name
				p_strBody = Replace(p_strBody,"&lt;customerName&gt;",pstrCustName)	' - this is the customer's first name, middle initial, last name
				p_strBody = Replace(p_strBody,"&lt;unsubscribeLink&gt;",pstrUnsubscribeURL)	' - this is the link to the customer order history page

				p_strBody = Replace(p_strBody,"&lt;CUSTOMERFIRSTNAME&gt;",pstrCustFirstName)	' - this is the customer's first name
				p_strBody = Replace(p_strBody,"&lt;CUSTOMERLASTNAME&gt;",pstrCustLastName)	' - this is the customer's last name
				p_strBody = Replace(p_strBody,"&lt;CUSTOMERNAME&gt;",pstrCustName)	' - this is the customer's first name, middle initial, last name
				p_strBody = Replace(p_strBody,"&lt;UNSUBSCRIBELINK&gt;",pstrUnsubscribeURL)	' - this is the link to the customer order history page
			'Else
				p_strBody = Replace(p_strBody,"<customerFirstName>",pstrCustFirstName)	' - this is the customer's first name
				p_strBody = Replace(p_strBody,"<customerLastName>",pstrCustLastName)	' - this is the customer's last name
				p_strBody = Replace(p_strBody,"<customerName>",pstrCustName)	' - this is the customer's first name, middle initial, last name
				p_strBody = Replace(p_strBody,"<unsubscribeLink>",pstrUnsubscribeURL)	' - this is the link to the customer order history page

				p_strBody = Replace(p_strBody,"<CUSTOMERFIRSTNAME>",pstrCustFirstName)	' - this is the customer's first name
				p_strBody = Replace(p_strBody,"<CUSTOMERLASTNAME>",pstrCustLastName)	' - this is the customer's last name
				p_strBody = Replace(p_strBody,"<CUSTOMERNAME>",pstrCustName)	' - this is the customer's first name, middle initial, last name
				p_strBody = Replace(p_strBody,"<UNSUBSCRIBELINK>",pstrUnsubscribeURL)	' - this is the link to the customer order history page
			'End If

			'added for new methods
			p_strBody = Replace(p_strBody,"{customerFirstName}",pstrCustFirstName)	' - this is the customer's first name
			p_strBody = Replace(p_strBody,"{customerLastName}",pstrCustLastName)	' - this is the customer's last name
			p_strBody = Replace(p_strBody,"{customerName}",pstrCustName)	' - this is the customer's first name, middle initial, last name
			p_strBody = Replace(p_strBody,"{unsubscribeLink}",pstrUnsubscribeURL)	' - this is the link to the customer order history page

			p_strBody = Replace(p_strBody,"{CUSTOMERFIRSTNAME}",pstrCustFirstName)	' - this is the customer's first name
			p_strBody = Replace(p_strBody,"{CUSTOMERLASTNAME}",pstrCustLastName)	' - this is the customer's last name
			p_strBody = Replace(p_strBody,"{CUSTOMERNAME}",pstrCustName)	' - this is the customer's first name, middle initial, last name
			p_strBody = Replace(p_strBody,"{UNSUBSCRIBELINK}",pstrUnsubscribeURL)	' - this is the link to the customer order history page

			MakeSubstitutions = p_strBody

	End Function	'MakeSubstitutions

	'***********************************************************************************************

	Private Function ssSendMail(strEmailSubject, strEmailBody, strEmailToAddr, blnHTML)

	Dim pobjMail
	Dim pstrSuccess

		If Len(Request.Form("chkTestMail")) > 0 Then
			With Response
				.Write "<table class=tbl border=1 cellspacing=0 cellpadding=3>" & vbcrlf
				.Write "<tr><th colspan=2 class=tblhdr>Email function in test mode</th></tr>" & vbcrlf
				.Write "<tr><td align=right>From:&nbsp;</td><td align=left>&nbsp;" & pstrEmailFrom & "</td></tr>" & vbcrlf
				.Write "<tr><td align=right>To:&nbsp;</td><td align=left>&nbsp;" & strEmailToAddr & "</td></tr>" & vbcrlf
				.Write "<tr><td align=right>Subject:&nbsp;</td><td align=left>&nbsp;" & strEmailSubject & "</td></tr>" & vbcrlf
				.Write "<tr><td align=right>Body:&nbsp;</td><td align=left>&nbsp;" & Replace(strEmailBody,vbcrlf,"<br />") & "</td></tr>" & vbcrlf
				.Write "</table>"
			End With
			ssSendMail = "Email not sent - Test mode selected"
		Else

			On Error Resume Next
		
			pstrSuccess = True
			If Len(pstrEmailFromName) = 0 Then pstrEmailFromName = pstrEmailFrom
	'		pstrMailMethod = "Simple Mail"
			Select Case pstrMailMethod
				Case "CDONTS Mail"
					Set pobjMail = Server.CreateObject("CDONTS.NewMail")
					With pobjMail
						If blnHTML Then
							.BodyFormat = 0
							.MailFormat = 0
						Else
							.BodyFormat = 1
							.MailFormat = 1
						End If
						.Send pstrEmailFrom, strEmailToAddr, strEmailSubject,strEmailBody
					End With
					Set pobjMail = Nothing
				Case "ASP Mail"
					Set pobjMail = Server.CreateObject ("smtpsvg.mailer")
					With pobjMail
						If blnHTML Then
							.ContentType = "text/html"
						End If
						.QMessage = False	'True	False
						.RemoteHost = pstrMailServer
						.AddRecipient strEmailToAddr, strEmailToAddr
						.FromAddress = pstrEmailFrom
						.FromName = pstrEmailFromName
						.Subject = strEmailSubject
						.BodyText = strEmailBody
						.SendMail
					End With
					Set pobjMail = Nothing
				Case Else
					Call createMail("",strEmailToAddr & "|" & pstrEmailFrom & "|" & "" & "|" & strEmailSubject & "|" & strEmailBody)
			End Select

			If Err.number <> 0 Then
				If Err.number = 424 Then
					pstrSuccess = "Could not create  " & pstrMailMethod & " object. Check to make sure this mail component is properly installed."
				Else
					pstrSuccess = "Error " & Err.number & ": " & Err.Description
				End If
				
				Err.Clear
			Else
				pstrSuccess = ""
			End If
		End If
		
		ssSendMail = pstrSuccess

	End Function	'ssSendMail

	'***********************************************************************************************

	Sub LoadDefaultEmailSettings

	Dim pobjRS
	Dim pstrSQL

		pstrSQL = "Select adminPrimaryEmail,adminMailMethod,adminMailServer,adminDomainName  from sfAdmin"
		Set pobjRS = Server.CreateObject("ADODB.RecordSet")
		With pobjRS
			.CursorLocation = 3 'adUseClient
			.Open pstrSQL, cnn, 3, 1
			If Len(pstrEmailFrom) = 0 Then pstrEmailFrom = .Fields("adminPrimaryEmail").Value
			pstrMailMethod = .Fields("adminMailMethod").Value
			pstrMailServer = .Fields("adminMailServer").Value
			pstrSiteURL = .Fields("adminDomainName").Value
			If Right(pstrSiteURL,1)  <> "/" Then pstrSiteURL = pstrSiteURL & "/"
			.Close
		End With
		Set pobjRS = Nothing
		
	End Sub	'LoadDefaultEmailSettings

	'***********************************************************************************************

	Public Sub OutputSummary()

	'On Error Resume Next

	Dim i
	Dim pstrName
	Dim pstrEmail
	Dim pstrSubscribed
	Dim pobjDic
	Dim pobjCustomer
	Dim plngID
	Dim vItem
	Dim pstrKey
	Dim plngItemCount
	Dim plngHeight
	
		plngItemCount = 0
		plngHeight = 300
		
		If plngMailingInProcess = 1 Then Exit Sub

		'Now for the table opener
		Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' bgcolor='whitesmoke' id='tblSummary' rules='none'>"

		'pblnUniqueEmailAddressesOnly = False	'True	'False
		'pblnAutoSelectCustomers = False
		If pblnAutoSelectCustomers Then
			plngItemCount = prsOrderSummaries.RecordCount

			'Now for the summary table contents
			Response.Write "<tr><td colspan='9'>"
			Response.Write "<div name='divSummary' style='height:" & plngHeight & "; overflow:scroll;'>"
			Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1'  rules='none'" _
						& "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary'" _
						& ">"
			If prsOrderSummaries.EOF Then
				Response.Write "<TR><TD align=center COLSPAN=9><h3>There are no customers meeting this criteria</h3></TD></TR>"
			Else
				Response.Write "<TR><TD align=center COLSPAN=9><h3>Customers to be emailed: " & plngItemCount & "</h3></TD></TR>"
			End If
			
			'Now for the summary table closer
			Response.Write "</TABLE></div>"
		Else
			'Now for the header row
			Response.Write "	<tr class='tblhdr'>"
			Response.Write "  <TH align='left' width='4%'><input type='checkbox' name='chkCheckAll' id='chkCheckAll1'  onclick='checkAll(this.form.custID, this.checked);checkAll(this.form.chkCheckAll2, this.checked);' value=''></TH>"
			Response.Write "  <TH align='left' width='19%'>Name</TH>" & vbCrLf
			Response.Write "  <TH align='left' width='19%'>Email</TH>" & vbCrLf
			Response.Write "  <TH align='left' width='14%'>Last Ordered</TH>" & vbCrLf
			Response.Write "  <TH align='left' width='12%'>Max Order</TH>" & vbCrLf
			Response.Write "  <TH align='left' width='9%'>Orders</TH>" & vbCrLf
			Response.Write "  <TH align='left' width='14%'>Total Orders</TH>" & vbCrLf
			Response.Write "  <TH align='left' width='5%'>Subscribed</TH>" & vbCrLf
			Response.Write "  <TH align='left' width='4%'>&nbsp;</TH>" & vbCrLf
			Response.Write "	</tr>"

			'Now for the summary table contents
			Response.Write "<tr><td colspan='9'>"
			Response.Write "<div name='divSummary' style='height:" & plngHeight & "; overflow:scroll;'>"
			Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1'  rules='none'" _
						& "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary'" _
						& ">"
			If prsOrderSummaries.EOF Then Response.Write "<TR><TD align=center COLSPAN=9><h3>There are no customers meeting this criteria</h3></TD></TR>"

			Set pobjDic = Server.CreateObject("Scripting.Dictionary")
			For i = 1 To prsOrderSummaries.RecordCount
				If pblnUniqueEmailAddressesOnly Then
					pstrKey = prsOrderSummaries.Fields("custEmail").Value
				Else
					pstrKey = "key_" & prsOrderSummaries.Fields("custID").Value
				End If

				If pobjDic.Exists(pstrKey) Then
					pobjDic.Item(pstrKey).AddOrder(prsOrderSummaries)
				Else
					Set pobjCustomer = New clsCustomer
					pobjCustomer.AddOrder(prsOrderSummaries)
					pobjDic.Add pstrKey, pobjCustomer
					plngItemCount = plngItemCount + 1
				End If
				prsOrderSummaries.MoveNext
			Next
			
			For Each vItem in pobjDic
				With pobjDic.Item(vItem)
					If CStr(.Subscribed & "") = "1" Then
						pstrSubscribed = "Subscribed"
					Else
						pstrSubscribed = "Not Subscribed"
					End If
					Response.Write "<TR>"
					Response.Write "<TD align='left' width='21%'><INPUT TYPE=CHECKBOX NAME=custID VALUE='" & .ID & "'>&nbsp;" & .WholeName
					Response.Write "<TD align='left' width='19%'>" & .Email & "&nbsp;</TD>"
					Response.Write "<TD align='center' width='13%'>" & .DateLastOrder & "&nbsp;</TD>"
					Response.Write "<TD align='right' width='12%'>" & FormatCurrency(.MaxOrder,2) & "&nbsp;</TD>"
					Response.Write "<TD align='right' width='8%'>" & .OrderCount & "&nbsp;</TD>"
					Response.Write "<TD align='right' width='12%'>" & FormatCurrency(.TotalOrders) & "&nbsp;</TD>"
					Response.Write "<TD align='right' width='15%'>" & pstrSubscribed & "&nbsp;</TD>"
					Response.Write "</TR>" & vbcrlf
				End With
			Next

			'Now for the summary table closer
			Response.Write "</TABLE></div>"

			'Now for the record summary
			Response.Write "	<tr class='tblhdr'>"
			Response.Write "  <TH align='left' colspan=2><input type='checkbox' name='chkCheckAll' id='chkCheckAll2'  onclick='checkAll(this.form.custID, this.checked);checkAll(this.form.chkCheckAll1, this.checked);' value=''></TH>" & vbCrLf
			Response.Write "  <TH align='left' colspan=3>Customers Returned: " & plngItemCount & "</TH>" & vbCrLf
			Response.Write "  <TH align='left' colspan=4><INPUT class='butn' id='btnSubscribe' name='btnSubscribe' type=button value='Subscribe' onclick='Subscribe(true);'>&nbsp;&nbsp;<INPUT class='butn' id='btnUnsubscribe' name='btnUnsubscribe' type=button value='Unsubscribe' onclick='Subscribe(false);'></TH>" & vbCrLf
			Response.Write "	</tr>"

		End If	'pblnAutoSelectCustomers

		'Now for the table closer
		Response.Write "</TABLE>"
			
	End Sub      'OutputSummary

	'***********************************************************************************************

	Function ValidateValues

	Dim strError

	    strError = ""
	    
	    If Len(mstrEmailFrom) = 0 Then strError = strError & "Please enter a <i>From</i> address." & cstrDelimeter
	    If Len(mstrEmailSubject) = 0 Then strError = strError & "Please enter a <i>Subject</i>." & cstrDelimeter
	    If Len(mstrEmailBody) = 0 Then strError = strError & "Please enter <i>email text</i>." & cstrDelimeter

	    pstrMessage = strError
	    ValidateValues = (Len(strError) = 0)

	End Function 'ValidateValues

	'***********************************************************************************************

End Class   'clsPromoMail

'***********************************************************************************************

Function SummaryFilter

Dim i
Dim paryTemp
Dim pstrMassUnsubscribe
Dim pstrsqlWhere
Dim pstrsqlHaving
Dim pstrTemp
Dim pstrJoinType	'used because of possibility of customers without orders

	'load the Product ID filter
	mstrProductID = Request.Form("ProductID")
	mstrProductIDExclude = Request.Form("ProductIDExclude")
	pstrMassUnsubscribe = Request.Form("massUnsubscribe")
	
	If Len(mstrProductID) > 0 Then
		paryTemp = Split(mstrProductID,", ")
		pstrsqlWhere = "(sfOrderDetails.odrdtProductID='" & paryTemp(0) & "'"
		For i = 1 To UBound(paryTemp)
			pstrsqlWhere = pstrsqlWhere & " OR sfOrderDetails.odrdtProductID='" & paryTemp(i) & "'"
		Next 'i
		pstrsqlWhere = pstrsqlWhere & ")"
		pstrJoinType = "INNER"
	End If
		
	If Len(mstrProductIDExclude) > 0 Then
		paryTemp = Split(mstrProductIDExclude,";")
		If Len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " AND (sfOrderDetails.odrdtProductID<>'" & paryTemp(0) & "'"
		Else
			pstrsqlWhere = "(sfOrderDetails.odrdtProductID<>'" & paryTemp(0) & "'"
		End If
		For i = 1 To UBound(paryTemp)
			pstrsqlWhere = pstrsqlWhere & " OR sfOrderDetails.odrdtProductID='" & paryTemp(i) & "'"
		Next 'i
		pstrsqlWhere = pstrsqlWhere & ")"
		pstrJoinType = "INNER"
	End If
'debugprint "mstrProductID", mstrProductID
'debugprint "mstrProductIDExclude", mstrProductIDExclude
'debugprint "pstrsqlWhere", pstrsqlWhere
	
	'load the Subscribed filter
	mradShowActive = Request.Form("radShowActive")
	If Len(mradShowActive) = 0 Then
		mradShowActive = 1
	Else
		mradShowActive = CLng(mradShowActive)
	End If
	If mradShowActive = 0 Then
		If Len(pstrsqlWhere) > 0  Then
			pstrsqlWhere = pstrsqlWhere & " AND (sfCustomers.custIsSubscribed=0)"
		Else
			pstrsqlWhere = "(sfCustomers.custIsSubscribed=0)"
		End If
		pstrsqlHaving = "(sfCustomers.custIsSubscribed=0)"
	ElseIf mradShowActive = 1 Then
		If Len(pstrsqlWhere) > 0  Then
			pstrsqlWhere = pstrsqlWhere & " AND (sfCustomers.custIsSubscribed=1)"
		Else
			pstrsqlWhere = "(sfCustomers.custIsSubscribed=1)"
		End If
		pstrsqlHaving = "(sfCustomers.custIsSubscribed=1)"
	End If
	
	If len(mbytDate_Filter) = 0 Then mbytDate_Filter = 0
	
	
	'load the date filter
	mbytDate_Filter = Request.Form("optDate_Filter")
	If len(mbytDate_Filter) = 0 Then mbytDate_Filter = 0

	mstrStartDate = Request.Form("StartDate")
	If len(mstrStartDate) > 0 then 
		if cblnSQLDatabase Then
			If Len(pstrsqlWhere) > 0  Then
				pstrsqlWhere = pstrsqlWhere & " AND (orderDate >= '" & mstrStartDate & " 12:00:00 AM')"
			Else
				pstrsqlWhere = "(orderDate >= '" & mstrStartDate & " 12:00:00 AM')"
			End If
		Else
			If Len(pstrsqlWhere) > 0  Then
				pstrsqlWhere = pstrsqlWhere & " AND (orderDate >= #" & mstrStartDate & " 12:00:00 AM#)"
			Else
				pstrsqlWhere = "(orderDate >= #" & mstrStartDate & " 12:00:00 AM#)"
			End If
		End If
		pstrJoinType = "INNER"
	End If
	
	mstrEndDate = Request.Form("EndDate")
	If len(mstrEndDate) > 0 then 
		if cblnSQLDatabase Then
			If Len(pstrsqlWhere) > 0  Then
				pstrsqlWhere = pstrsqlWhere & " AND (orderDate <= '" & mstrEndDate & " 11:59:59 PM')"
			Else
				pstrsqlWhere = "(orderDate <= '" & mstrEndDate & " 11:59:59 PM')"
			End If
		Else
			If Len(pstrsqlWhere) > 0  Then
				pstrsqlWhere = pstrsqlWhere & " AND (orderDate <= #" & mstrEndDate & " 11:59:59 PM#)"
			Else
				pstrsqlWhere = "(orderDate <= #" & mstrEndDate & " 11:59:59 PM#)"
			End If
		End If
		pstrJoinType = "INNER"
	End If

	'load the Manufacturer filter
	mlngManufacturerFilter = Request.Form("ManufacturerFilter")
	pstrTemp = Replace(mlngManufacturerFilter,"'","''")
	If Len(pstrTemp) > 0  Then
		If Len(pstrsqlWhere) > 0  Then
			pstrsqlWhere = pstrsqlWhere & " AND (sfOrderDetails.odrdtManufacturer='" & pstrTemp & "')"
		Else
			pstrsqlWhere = "(sfOrderDetails.odrdtManufacturer='" & pstrTemp & "')"
		End If
		pstrJoinType = "INNER"
	End If

	'load the Category filter
	mlngCategoryFilter = Request.Form("CategoryFilter")
	pstrTemp = Replace(mlngCategoryFilter,"'","''")
	If Len(pstrTemp) > 0  Then
		If Len(pstrsqlWhere) > 0  Then
			pstrsqlWhere = pstrsqlWhere & " AND (sfOrderDetails.odrdtCategory='" & pstrTemp & "')"
		Else
			pstrsqlWhere = "(sfOrderDetails.odrdtCategory='" & pstrTemp & "')"
		End If
		pstrJoinType = "INNER"
	End If

	'load the Vendor filter
	mlngVendorFilter = Request.Form("VendorFilter")
	pstrTemp = Replace(mlngVendorFilter,"'","''")
	If Len(pstrTemp) > 0  Then
		If Len(pstrsqlWhere) > 0  Then
			pstrsqlWhere = pstrsqlWhere & " AND (sfOrderDetails.odrdtVendor='" & pstrTemp & "')"
		Else
			pstrsqlWhere = "(sfOrderDetails.odrdtVendor='" & pstrTemp & "')"
		End If
		pstrJoinType = "INNER"
	End If
	
	'load the PricingLevel filter
	mlngPricingLevelFilter = Request.Form("PricingLevelFilter")
	If Len(mlngPricingLevelFilter) > 0  Then
		If Len(pstrsqlWhere) > 0  Then
			If mlngPricingLevelFilter = "None" then
				pstrsqlWhere = pstrsqlWhere & " and (sfCustomers.PricingLevelID is Null) "
				pstrsqlHaving = "(sfCustomers.PricingLevelID is Null)"
			Else
				'pstrsqlWhere = pstrsqlWhere & " and sfCustomers.PricingLevelID=" & CStr(CLng(mlngPricingLevelFilter) - 1)
				'pstrsqlHaving = "(sfCustomers.PricingLevelID=" & CStr(CLng(mlngPricingLevelFilter) - 1) & ")"
				pstrsqlWhere = pstrsqlWhere & " and sfCustomers.PricingLevelID=" & CStr(mlngPricingLevelFilter)
				pstrsqlHaving = "(sfCustomers.PricingLevelID=" & CStr(mlngPricingLevelFilter) & ")"
			End If
		Else
			If mlngPricingLevelFilter = "None" then
				pstrsqlWhere = " (sfCustomers.PricingLevelID is Null) "
				pstrsqlHaving = "(sfCustomers.PricingLevelID is Null)"
			Else
				'pstrsqlWhere = " sfCustomers.PricingLevelID=" & CStr(CLng(mlngPricingLevelFilter) - 1)
				'pstrsqlHaving = "(sfCustomers.PricingLevelID=" & CStr(CLng(mlngPricingLevelFilter) - 1) & ")"
				pstrsqlWhere = " sfCustomers.PricingLevelID=" & CStr(mlngPricingLevelFilter)
				pstrsqlHaving = "(sfCustomers.PricingLevelID=" & CStr(mlngPricingLevelFilter) & ")"
			End If
		End If
	End If

	'load the price filter
	mcurLowerPrice = Request.Form("LowerPrice")
	mcurUpperPrice = Request.Form("UpperPrice")
	If len(mcurLowerPrice) > 0 then
		If cblnSQLDatabase Then
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and convert(money,orderGrandTotal)>=" & mcurLowerPrice
			Else
				pstrsqlWhere = "convert(money,orderGrandTotal)>=" & mcurLowerPrice
			End If
'			sfOrderDetails.odrdtProductID
		Else
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " AND (cCur(orderGrandTotal)>=" & mcurLowerPrice & ")"
			Else
				pstrsqlWhere = " (cCur(orderGrandTotal)>=" & mcurLowerPrice & ")"
			End If
		End If
		pstrJoinType = "INNER"
	End If

	If len(mcurUpperPrice) > 0 then
		If cblnSQLDatabase Then
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and convert(money,orderGrandTotal)<=" & mcurUpperPrice
			Else
				pstrsqlWhere = "convert(money,orderGrandTotal)<=" & mcurUpperPrice
			End If
		Else
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and cCur(orderGrandTotal)<=" & mcurUpperPrice
			Else
				pstrsqlWhere = "cCur(orderGrandTotal)<=" & mcurUpperPrice
			End If
		End If
		pstrJoinType = "INNER"
	End If
	
	If Len(pstrMassUnsubscribe) > 0 Then
		If inStr(1, pstrMassUnsubscribe, vbcrlf) > 0 Then
			paryTemp = Split(pstrMassUnsubscribe, vbcrlf)
		ElseIf inStr(1, pstrMassUnsubscribe, vbTab) > 0 Then
			paryTemp = Split(pstrMassUnsubscribe, vbTab)
		Else
			paryTemp = Split(pstrMassUnsubscribe, ",")
		End If
		
		pstrsqlWhere = "sfCustomers.custEmail In (" & wrapSQLValue(paryTemp(0), False, enDatatype_string)
		For i = 1 To UBound(paryTemp)
			pstrsqlWhere = pstrsqlWhere & ", " & wrapSQLValue(paryTemp(i), False, enDatatype_string)
		Next 'i
		pstrsqlWhere = pstrsqlWhere & ")"
	End If

	If Len(pstrsqlWhere) > 0  Then pstrsqlWhere = " Where " & pstrsqlWhere
	If Len(pstrsqlHaving) > 0  Then pstrsqlHaving = " Having " & pstrsqlHaving
	If Len(pstrJoinType) = 0  Then pstrJoinType = "LEFT"

	SummaryFilter = pstrsqlWhere & "|" & pstrsqlHaving & "|" & pstrJoinType

	'Response.Write "<h4>" & pstrsqlWhere & "|" & pstrsqlHaving & "|" & pstrJoinType & "</h4>"
	
End Function    'SummaryFilter

Sub closeObj(objItem)
	ReleaseObject objItem
End Sub

'--------------------------------------------------------------------------------------------------
%>
<!--#include file="../SFLib/storeAdminSettings.asp"-->
<!--#include file="../SFLib/mail.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'**************************************************
'
'	Start Code Execution
'

On Error Goto 0	'added because of global error suppression in mail.asp

mstrPageTitle = "Promotional Mailing Administration"

'page variables
Dim mstrAction
Dim mclsPromoMail

Dim mblnShowFilter, mblnShowSummary
Dim mstrsqlWhere
Dim mstrShow

'Filter Elements
Dim mstrProductID
Dim mstrProductIDExclude

Dim mbytDate_Filter
Dim mstrStartDate, mstrEndDate

Dim mradUseUniqueEmails
Dim mradAutoSelectCustomers

Dim mcurLowerPrice
Dim mcurUpperPrice
Dim mlngManufacturerFilter
Dim mlngCategoryFilter
Dim mlngVendorFilter
Dim mlngPricingLevelFilter
Dim mradShowActive

Dim mstrEmailFrom
Dim mstrEmailFromName
Dim mstrEmailSubject
Dim mstrEmailBody
Dim mblnHTMLEmail

	
	Call WriteHeader("body_onload();",True)
	
    Set mclsPromoMail = New clsPromoMail
    With mclsPromoMail
    
		If Not .isDBUpgraded Then
			Set mclsPromoMail = Nothing
			Response.Redirect "ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=PromoMail"
		End If
		
		'Collect the form information
		mradUseUniqueEmails = Request.Form("UseUniqueEmails")
		If Len(mradUseUniqueEmails) = 0 Then mradUseUniqueEmails = cbytDefault_UseUniqueEmails
		
		mradAutoSelectCustomers = Request.Form("AutoSelectCustomers")
		If Len(mradAutoSelectCustomers) = 0 Then mradAutoSelectCustomers = cbytDefault_AutoSelectCustomers

		mstrEmailFrom = Request.Form("emailFrom")
		mstrEmailFromName = Request.Form("emailFromName")
		mstrEmailSubject = Request.Form("emailSubject")
		mstrEmailBody = Request.Form("emailBody")
		mblnHTMLEmail = (Request.Form("chkHTMLMail") = "1")
		
		If Len(mstrEmailFrom) > 0 Then 
			mclsPromoMail.EmailFrom = mstrEmailFrom
		Else
			mstrEmailFrom = mclsPromoMail.EmailFrom
		End If
		
		If Len(mstrEmailFromName) = 0 Then 
			mstrEmailFromName = mstrEmailFrom
		Else
			mclsPromoMail.EmailFromName = mstrEmailFromName
		End If

		mclsPromoMail.UniqueEmailAddressesOnly = CBool(mradUseUniqueEmails = "1")
		mclsPromoMail.AutoSelectCustomers = CBool(mradAutoSelectCustomers = "1")
	
		mstrShow = Request.Form("Show")
		If Len(mstrShow) = 0 Then mstrShow = "Filter"
		mstrAction = LoadRequestValue("Action")
		
		Select Case mstrAction
			Case "Subscribe"
				Call .Subscribe(True)
			Case "Unsubscribe"
				Call .Subscribe(False)
			Case "Sendmail"
				.SendPromotionalMailing
				mstrShow = "Results"
			Case "Filter"
				mstrShow = "Addresses"
		End Select

		Call .LoadOrderSummaries(SummaryFilter)

%>

<SCRIPT LANGUAGE="JavaScript">

/**********************************************************
*	Variables
**********************************************************

var theDataForm;
var mdicProductID = new ActiveXObject("Scripting.Dictionary");;
var mdicProductIDExclude = new ActiveXObject("Scripting.Dictionary");;

**********************************************************
*	Functions
**********************************************************

function body_onload()
function btnFilter_onclick(theButton)
function ChangeDate(theOption)
function checkIfHTML()
function DisplaySection(strSection)
function OpenFile(theFile)
function ReplaceEmailText()
function SendEmail()
function Subscribe(blnSubscribe)
function ValidInput(theForm)

**********************************************************
*	Begin Code
***********************************************************/

var theDataForm;
var mdicProductID = new ActiveXObject("Scripting.Dictionary");
var mdicProductIDExclude = new ActiveXObject("Scripting.Dictionary");
<% On Error Goto 0 %>
<%= setCustomDictionary(mstrProductID, ", ", "ProductID", "product") %>
<%= setCustomDictionary(mstrProductIDExclude, ", ", "ProductIDExclude", "product") %>

function body_onload()
{
	theDataForm = document.frmData;
	FillItem("ProductID");
	FillItem("ProductIDExclude");

	return DisplaySection(frmData.Show.value)
}

function btnFilter_onclick(theButton)
{
  if (ValidInput(theButton.form))
  {
	document.frmData.Action.value = "Filter";
	document.frmData.submit();
	return(true);
  }
}

function ChangeDate(theOption)
{

	switch (theOption.value)
	{
		case "0":
			document.frmData.StartDate.value= "";
			document.frmData.EndDate.value= "";
			break;
		case "1":
			document.frmData.StartDate.value= "<%= FormatDateTime(DateAdd("d",-1,Date())) %>";
			document.frmData.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "2":
			document.frmData.StartDate.value= "<%= FormatDateTime(DateAdd("ww",-1,Date())) %>";
			document.frmData.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "3":
			document.frmData.StartDate.value= "<%= FormatDateTime(DateAdd("m",-1,Date())) %>";
			document.frmData.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "4":
			document.frmData.StartDate.value= "<%= FormatDateTime(DateAdd("yyyy",-1,Date())) %>";
			document.frmData.EndDate.value= "<%= FormatDateTime(Date()) %>";
			break;
		case "5":
			document.frmData.StartDate.focus();
			break;
	}

}

function checkIfHTML()
{
var pblnHTML;
var pstrFile = document.frmData.emailBody.value;

	pblnHTML = (pstrFile.indexOf("<html>") != -1);
	pblnHTML = (pblnHTML || (pstrFile.indexOf("<HTML>") != -1))
	document.frmData.chkHTMLMail.checked = pblnHTML;
}

function DisplaySection(strSection)
{

var arySections = new Array("Filter","Addresses","Email","Results");

  frmData.Show.value = strSection;

 for (var i=0; i < arySections.length;i++)
 {
	if (arySections[i] == strSection)
	{
		document.all("tbl" + arySections[i]).style.display = "";
		document.all("td" + arySections[i]).className = "hdrSelected";
	}else{
		document.all("tbl" + arySections[i]).style.display = "none";
		document.all("td" + arySections[i]).className = "hdrNonSelected";
	}
 }	
 
	return(false);
}

function OpenFile(theFile)
//Initialize and script ActiveX controls not marked as safe needs to be set to Enable or Promopt
{

var pstrFilePath = theFile.value;
var fso = new ActiveXObject("Scripting.FileSystemObject");
var MyFile = fso.OpenTextFile(pstrFilePath, 1, false);
var pstrFile= "";
var pblnHTML;

	while (! MyFile.AtEndOfStream)
	{
		pstrFile = pstrFile + MyFile.ReadLine() + "\n";
	}
	MyFile.close();

	document.frmData.emailBody.value = pstrFile;
	pblnHTML = (pstrFile.indexOf("<html>") != -1);
	pblnHTML = (pblnHTML || (pstrFile.indexOf("<HTML>") != -1))
	document.frmData.chkHTMLMail.checked = pblnHTML;
}

function ReplaceEmailText()
{
var theString = document.frmData.emailBody.value;
var r;

	r = theString.replace(/<dateShipped>/i, document.frmData.ssDateOrderShipped.value);
	theString = r.replace(/<shipMethod>/i, document.frmData.ssShippedVia.options[document.frmData.ssShippedVia.selectedIndex].text);

	document.frmData.emailBody.value = theString;

}

function SendEmail()
{
	if (ValidInput(document.frmData))
	{
		document.frmData.Action.value = "Sendmail";
		openOutputWindow('ssOutputWindow.asp');
		document.frmData.submit();
	}
}

function Subscribe(blnSubscribe)
{
	if (blnSubscribe)
	{
		document.frmData.Action.value = "Subscribe";
	}else{
		document.frmData.Action.value = "Unsubscribe";
	}
	document.frmData.submit();
}

function ValidInput(theForm)
{
//  if (!isNumeric(theForm.prodWeight,false,"Please enter a number for the Order weight.")) {return(false);}

	var theSelect = theForm.ProductID;
	for (var i=0; i < theSelect.length;i++){theSelect.options[i].selected = true;}
	
	theSelect = theForm.ProductIDExclude;
	for (var i=0; i < theSelect.length;i++){theSelect.options[i].selected = true;}

    return(true);
}

//-->
</script>

<CENTER>
<TABLE border=0 cellPadding=5 cellSpacing=1 width="95%">
  <TR>
    <TH><div class="pagetitle "><%= mstrPageTitle %></div></TH>
 </TR>
</TABLE>

<%= .OutputMessage %>

<FORM action="ssPromoMailAdmin.asp" id=frmData name=frmData onsubmit="return ValidInput(this);" method=post>

<input type="hidden" id=Show name=Show value="<%= mstrShow %>">
<input type="hidden" id=Action name=Action value="">
<input type="hidden" id=blnShowSummary name=blnShowSummary value="">
<input type="hidden" id=blnShowFilter name=blnShowFilter value="">

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" id="tblMaster" rules="none">
	  <tr>
		<TD nowrap ID="tdFilter" class="hdrNonSelected" onclick="return DisplaySection('Filter');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Narrow search of email address">Step 1. Narrow Search</TD>
    <th nowrap width="1pt"></th>
		<TD nowrap ID="tdAddresses" class="hdrNonSelected" onclick="return DisplaySection('Addresses');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View addresses" >Step 2. Select Addresses</TD>
    <th nowrap width="1pt"></th>
		<TD nowrap ID="tdEmail" class="hdrNonSelected" onclick='return DisplaySection("Email");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Create promotional email" >Step 3. Create Email</TD>
    <th nowrap width="1pt"></th>
		<TD nowrap ID='tdResults' class="hdrNonSelected" onclick='return DisplaySection("Results");' onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Results of the promotional email" >Step 4. Results</TD>
    <th nowrap width="90%"></th>
		<TD nowrap class="hdrNonSelected" align=right><INPUT class="butn" type=image src="images/help.gif" value="?" onclick="OpenHelp('ssHelpFiles/PromotionalMail/help_PromotionalMailManager.htm')" id=btnHelp name=btnHelp></TD>
	  </tr>
  <tr>
    <td colspan="9" class="hdrSelected" height="1px"></td>
  </tr>
  <TR>
  <TD COLSPAN=9>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="0" id="tblFilter" rules="none">
<colgroup align="left">
<colgroup align="left">
  <TR>
    <TD valign="top">
		<fieldset>
			<legend>Include people who ordered products</legend>
			<select id="ProductID" name="ProductID" size="5" ondblclick="openMovementWindow('ProductID','product');" multiple></select>
			<a href="" onclick="openMovementWindow('ProductID','product'); return false;"><img src="images/properites.gif" border="0"></a><br />

	<p>From Category<br />
	<select size="1"  id=CategoryFilter name=CategoryFilter>
<% 	if len(mlngCategoryFilter) = 0 then
		Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
	else
		Response.Write "<option value=''>- All -</Option>" & vbcrlf
	end if
	Call MakeCombo("Select catName from sfCategories Order By catName","catName","catName",mlngCategoryFilter)
 %>
	</select></p>
    <p>From Manufacturer<br />
	<select size="1"  id=ManufacturerFilter name=ManufacturerFilter>
<% 	if len(mlngManufacturerFilter) = 0 then
		Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
	else
		Response.Write "<option value=''>- All -</Option>" & vbcrlf
	end if
	Call MakeCombo("Select mfgName from sfManufacturers Order By mfgName","mfgName","mfgName",mlngManufacturerFilter)
 %>
	</select></p>
	<p>From Vendor<br />
	<select size="1"  id=VendorFilter name=VendorFilter>
<% 	if len(mlngVendorFilter) = 0 then
		Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
	else
		Response.Write "<option value=''>- All -</Option>" & vbcrlf
	end if
	Call MakeCombo("Select vendName from sfVendors Order By vendName","vendName","vendName",mlngVendorFilter)
%>
	</select></p>
		</fieldset>
		
		<!--
        <p>Exclude people who ordered products<br />
		<select id="ProductIDExclude" name="ProductIDExclude" size="5" ondblclick="openMovementWindow('ProductIDExclude','product');" multiple></select>
		<a href="" onclick="openMovementWindow('ProductIDExclude','product'); return false;"><img src="images/properites.gif" border="0"></a>
		-->
		<input id="ProductIDExclude" name="ProductIDExclude" type="hidden" value="">
	</TD>
    <TD valign="top">
		<fieldset>
			<legend>Show Only Orders Placed</legend>
			<input type="radio" value="1" <%= isChecked(mbytDate_Filter="1") %> name="optDate_Filter" onclick="ChangeDate(this);">Day&nbsp;
			<input type="radio" value="2" <%= isChecked(mbytDate_Filter="2") %> name="optDate_Filter" onclick="ChangeDate(this);">Week&nbsp;
			<input type="radio" value="3" <%= isChecked(mbytDate_Filter="3") %> name="optDate_Filter" onclick="ChangeDate(this);">Month&nbsp;
			<input type="radio" value="4" <%= isChecked(mbytDate_Filter="4") %> name="optDate_Filter" onclick="ChangeDate(this);">Year&nbsp;
			<input type="radio" value="5" <%= isChecked(mbytDate_Filter="5") %> name="optDate_Filter" onclick="ChangeDate(this);">Custom&nbsp;
			<input type="radio" value="0" <%= isChecked(mbytDate_Filter="0") %> name="optDate_Filter" onclick="ChangeDate(this);">All&nbsp;<br />
	        
			<LABEL for="StartDate">Start Date:&nbsp;</LABEL><INPUT id=StartDate name=StartDate Value="<%= mstrStartDate %>">
			<A HREF="javascript:doNothing()" title="Select start date"
			onClick="setDateField(document.frmData.StartDate); top.newWin = window.open('calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
			<IMG SRC="images/calendar.gif" BORDER=0></A><br />

			<LABEL for="EndDate">&nbsp;&nbsp;End Date:&nbsp;</LABEL><INPUT id=EndDate name=EndDate Value="<%= mstrEndDate %>">
			<A HREF="javascript:doNothing()" title="Select end date"
			onClick="setDateField(document.frmData.EndDate); top.newWin = window.open('calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
			<IMG SRC="images/calendar.gif" BORDER=0></A>

		<p> With Totals between<br />
        <input type="text" id="LowerPrice" name="LowerPrice" size="10" value="<%= mcurLowerPrice %>" maxlength=15><br />
        And<br />
        <input type="text" id="UpperPrice" name="UpperPrice" size="10" value="<%= mcurUpperPrice %>" maxlength=15></p>
		</fieldset>

		<fieldset>
			<legend>Bulk email addresses</legend>
			<textarea name="massUnsubscribe" id="massUnsubscribe" rows="4" cols="35" title="Enter either a comma, tab, or line feed delimited list of email addresses."></textarea>
		</fieldset>
	</TD>
    <TD valign="top">
		<fieldset>
			<legend>Show Customers that are:</legend>
			<input type="radio" value="1" <%= isChecked(mradShowActive=1 or mradShowActive="") %> name="radShowActive">Subscribed<br />
			<input type="radio" value="0" <%= isChecked(mradShowActive=0) %> name="radShowActive">Not Subsubscribed<br />
			<input type="radio" value="2" <%= isChecked(mradShowActive=2) %> name="radShowActive">All</p>

			<% If cblnAddon_PricingLevelMgr Then %>
			<p>Assigned to Pricing Level<br />
			<select size="1" name="PricingLevelFilter" id="PricingLevelFilter">
			<%
				'added for Pricing Level Manager - No Pricing Level Filter
				if len(mlngPricingLevelFilter) = 0 then
					Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
				else
					Response.Write "<option value=''>- All -</Option>" & vbcrlf
				end if
				If mlngPricingLevelFilter = "None" then
					Response.Write "<option value='None' selected>- None -</Option>" & vbcrlf
				Else
					Response.Write "<option value='None'>- None -</Option>" & vbcrlf
				End If
				Call MakeCombo("Select PricingLevelID, PricingLevelName from PricingLevels Order By PricingLevelName","","",mlngPricingLevelFilter)
			%>
			</select></p>
			<% End If	'cblnAddon_PricingLevelMgr %>
			<table class="tbl" border=1 cellspacing=0 cellpadding=2>
			<tr><td>Your database may contain multiple entries for the same person or same email address.<br />
			<input type=radio name="UseUniqueEmails" id="UseUniqueEmails0" value="0" <%= isChecked(mradUseUniqueEmails="0") %>><label for="UseUniqueEmails0">Use all entries</label><br />
			<input type=radio name="UseUniqueEmails" id="UseUniqueEmails1" value="1" <%= isChecked(mradUseUniqueEmails="1") %>><label for="UseUniqueEmails1">Use only unique email addresses</label>
			</td></tr>
			</table>

			<p>
			<input type=radio name="AutoSelectCustomers" id="AutoSelectCustomers0" value="0" <%= isChecked(mradAutoSelectCustomers="0") %>><label for="AutoSelectCustomers0">Manually select customers</label><br />
			<input type=radio name="AutoSelectCustomers" id="AutoSelectCustomers1" value="1" <%= isChecked(mradAutoSelectCustomers="1") %>><label for="AutoSelectCustomers1">Automatically select customers</label>
			</p>
		</fieldset>
        
	</TD>
	<td valign="top" align="center">
	<INPUT class="butn" id=btnFilter name=btnFilter type=button value="Apply Filter" onclick="btnFilter_onclick(this);"><br />
	</td>
</tr>
</TABLE>
<br />

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="0" id="tblAddresses" rules="none">
  <TR>
    <TD valign="top"><%= .OutputSummary %></TD>
  </TR>
</table>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none" id="tblEmail">
	<COLGROUP align="left" width="25%">
	<COLGROUP align="left" width="75%">
  <tr>
    <td align="right">From: (email address)</td>
    <td><input type="text" name="emailFrom" size="75" VALUE="<%= mstrEmailFrom %>"></td>
  </tr>
  <tr>
    <td align="right">From: (display Name)</td>
    <td><input type="text" name="emailFromName" size="75" VALUE="<%= mstrEmailFromName %>"></td>
  </tr>
  <tr>
    <td align="right">Subject:</td>
    <td><input type="text" name="emailSubject" size="75" VALUE="<%= mstrEmailSubject %>"></td>
  </tr>
  <tr>
    <td align="right">Body:</td>
    <td>
      <textarea rows="10" name="emailBody" cols="80"><%= mstrEmailBody %></textarea><br />
	  <input type="file" id="emailFile" name="emailFile" onchange="OpenFile(this);" title="Select a file from your hard drive to use for your emailing.">
	  <a HREF="javascript:doNothing()" onClick="openACE(document.frmData.emailBody); checkIfHTML();" title="Edit this field with the HTML Editor"><img SRC="images/prop.bmp" BORDER=0></a>
    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input type=checkbox id=chkHTMLMail name=chkHTMLMail <% If mblnHTMLEmail Then Response.Write "checked" %> value="1">&nbsp;Send as HTML</td>
  </tr>
  <tr>
    <TD colspan=2 align=center><hr width=90%></TD>
  </tr>
  <tr>
    <TD>&nbsp;</TD>
	<TD><INPUT class='butn' id=btnSend name=btnSend type=button value='Send' onclick="SendEmail();">&nbsp;<input type="checkbox" value="1" name="chkTestMail" id="chkTestMail" <% If Len(Request.Form("chkTestMail")) > 0 Then Response.Write "checked" %>><label for="chkTestMail" title="Results are dumped to screen">Test Mailing</label></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
	<td>
	  <% If .MailingInProcess > 0 Then %>
	  <br />
	    <h3>There appears to be a mailing in progress. There are <em><%= .MailingInProcess %></em> names remaining.</h3>
	    <input type="checkbox" id="chkMailingInProcess" name=chkMailingInProcess value='1' checked>&nbsp;<label for="chkMailingInProcess">Uncheck this box to start a new mailing</label><br />
	  <% End If %>
	</td>
  </tr>
</table>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="0" id="tblResults" rules="none">
  <TR>
    <TD valign="top">
<%= .OutputResults %>
    </TD>
  </TR>
</table>

</FORM>


</TD>
</TR>

</TABLE>
<%
End With	'    mclsPromoMail
Set mclsPromoMail = Nothing
Call ReleaseObject(cnn)
If Response.Buffer Then Response.Flush
%>
<!--#include file="adminFooter.asp"-->
</CENTER>
</BODY>
</HTML>
