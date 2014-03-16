<%
'********************************************************************************
'*   Sandshot Software Common Email Class file
'*   
'*   File Version:		1.00.0001
'*   Revision Date:		August 18, 2004
'*
'*   1.00.001 (November 15, 2003)
'*   - Initial Release
'*
'********************************************************************************
'*	Dependencies
'********************************************************************************
'*
'*	 Requires modDatabase.asp v1.00.004 or later for remote email sending - applies to StoreFront 6.0 only
'*
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************

'***********************************************************************************************
'*	Email Enumerations
'***********************************************************************************************

Const enEmail_FileName = 0
Const enEmail_Subject = 1
Const enEmail_Body = 2
Const enEmail_TemplateName = 3

'***********************************************************************************************

Class clsEmail

'**********************************************************
'*	Function Reference
'**********************************************************
'
'Public Properties (Get/Let unless otherwise noted)
'	Public Property From
'	Public Property FromName
'	Public Property To
'	Public Property CC
'	Public Property BCC
'	Public Property Subject
'	Public Property Body
'	Public Property HTML
'	Public Property MailMethod
'	Public Property MailServer
'	Public Property Get subjectWithReplacements()
'	Public Property Get bodyWithReplacements()
'	Public Property Let ShowFailures(byVal value)
'	Private Sub Initialize
'
'INTERNAL ERROR HANDLING
'	Private Sub addError(byVal strErrorMessage)
'	Private Sub addMessage(byVal strErrorMessage)
'	Public Function ErrorMessages()
'
'Replacement Fields
'
'	Private Sub InitializeStandardReplacements()
'	Public Sub addCap(byVal strStart, byVal strFinish)
'	Public Sub clearReplacementValues()
'	Public Sub setReplacementValue(byVal strField, byVal strValue)
'	Public Function getReplacementValue(byVal strField)
'	Public Function makeReplacements(ByVal strSource)
'
'	Public Function Send
'
'	Public Function mailAsHTML()
'
'For remote sending Fields
'
'	Private Function sendRemoteEmail(byVal blnSendNow)
'	Private Function processEmailReturn(byVal strXML)
'	Public Property remoteEmailResults()
'
'Email templates
'
'	Public Function LoadEmailTemplates(ByVal strDirectoryPath, ByVal strDefaultFileToLoad, ByRef aryEmails)
'	Private Sub ParseEmailTemplate(ByRef objFile, ByRef strSubject, ByRef strBody, ByRef strTemplateName)

Private pstrFrom
Private pstrFromName
Private pstrTo
Private pstrCC
Private pstrBCC
Private pstrSubject
Private pstrBody
Private pblnHTML

Private pstrMailMethod
Private pstrMailServer
Private pstrMailServerUserName
Private pstrMailServerPassword

Private paryReplacementFields
Private paryReplacementFields_Repeating
Private pstrRepeatingItemTag
Private pstrRepeatingItemTemplate
Private pstrRepeatingItemReplacementText
Private paryEndCaps
Private pblnShowFailures

'***********************************************************************************************

Private Sub class_Initialize()
    pblnShowFailures = False
    Call InitializeStandardReplacements
    Call Initialize
End Sub

Private Sub class_Terminate()
    On Error Resume Next

	'For Remote Sending
	If isObject(xmlRoot) Then Set xmlRoot = Nothing
	If isObject(xmlDoc) Then Set xmlDoc = Nothing
End Sub

'***********************************************************************************************
'	INTERNAL ERROR HANDLING
'***********************************************************************************************
	Private paryErrors(0)
	Private paryMessages(0)
	'***********************************************************************************************

	Private Sub addError(byVal vntErrorMessage)
	
	Dim plngErrorCount
	Dim i
	
	On Error Resume Next
	
		plngErrorCount = ErrorCount
		If isArray(vntErrorMessage) Then
			If hasErrors Then
				ReDim Preserve paryErrors(plngErrorCount + UBound(vntErrorMessage) + 1)
			ElseIf UBound(vntErrorMessage) > 0 Then
				ReDim paryErrors(UBound(vntErrorMessage))
			End If
			
			For i = 0 To UBound(vntErrorMessage)
				paryErrors(i + plngErrorCount) = vntErrorMessage(i)
			Next 'i
		Else
			If hasErrors Then ReDim Preserve paryErrors(plngErrorCount + 1)
			paryErrors(plngErrorCount) = vntErrorMessage
		End If

	End Sub	'addError

	'***********************************************************************************************

	Private Sub addMessage(byVal vntMessage)
	
	Dim plngMessageCount
	Dim i
	
		plngMessageCount = MessageCount
		If isArray(vntMessage) Then
			If hasMessages Then
				ReDim Preserve paryMessages(plngMessageCount + UBound(vntMessage) + 1)
			ElseIf UBound(vntMessage) > 0 Then
				ReDim paryMessages(UBound(vntMessage))
			End If
			
			For i = 0 To UBound(vntMessage)
				paryMessages(i + plngMessageCount) = vntMessage(i)
			Next 'i
		Else
			If hasMessages Then ReDim Preserve paryMessages(plngMessageCount + 1)
			paryMessages(plngMessageCount) = vntMessage
		End If

	End Sub	'addMessage

	'***********************************************************************************************

	Public Property Get ErrorCount()
		If hasErrors Then
			ErrorCount = arrayLength(paryErrors) + 1
		Else
			ErrorCount = 0
		End If
	End Property 'ErrorCount

	'***********************************************************************************************

	Public Property Get MessageCount()
		If hasMessages Then
			MessageCount = arrayLength(paryMessages) + 1
		Else
			MessageCount = 0
		End If
	End Property 'MessageCount

	'***********************************************************************************************

	Public Property Get hasErrors()
		If CBool(UBound(paryErrors) > 0) Then
			hasErrors = True
		Else
			hasErrors = CBool(Len(paryErrors(0)) > 0)
		End If
	End Property 'hasErrors

	'***********************************************************************************************

	Public Property Get hasMessages()
		If CBool(UBound(paryMessages) > 0) Then
			hasMessages = True
		Else
			hasMessages = CBool(Len(paryMessages(0)) > 0)
		End If
	End Property 'hasMessages

	'***********************************************************************************************

	Public Property Get ErrorMessages()
		ErrorMessages = paryErrors
	End Property 'ErrorMessages

	'***********************************************************************************************

	Public Property Get Messages()
		Messages = paryMessages
	End Property 'Messages

	'***********************************************************************************************

	Public Function writeErrorMessages()

	Dim i
	Dim pstrOut

	    For i = 0 To UBound(paryErrors)
	        If Len(paryErrors(i)) > 0 Then pstrOut = pstrOut & "<P align='center'><H4><FONT color=Red>" & paryErrors(i) & "</FONT></H4></P>"
	    Next 'i
	    
	    For i = 0 To UBound(paryMessages)
	        If Len(paryMessages(i)) > 0 Then pstrOut = pstrOut & "<P align='center'><H4>" & paryMessages(i) & "</H4></P>"
	    Next 'i
	    
	    writeErrorMessages = pstrOut

	End Function 'writeErrorMessages

	'***********************************************************************************************

	Private Function checkLocalError(byVal strLocalErrorMessage)

	    If Len(Trim(strLocalErrorMessage)) = 0 Then
	        checkLocalError = False
	    Else
	        Call addError(strLocalErrorMessage)
	        checkLocalError = True
	    End If
	    
	End Function 'checkLocalError

	'***********************************************************************************************

	Private Function arrayLength(byVal ary)
		If isArray(ary) Then
			arrayLength = UBound(ary) + 1
		Else
			arrayLength = 0
		End If
	End Function 'arrayLength

'***********************************************************************************************
'	Public Properties
'***********************************************************************************************

	Public Property Let From(byVal value)
		pstrFrom = value
	End Property
	Public Property Get From()
	    From = pstrFrom
	End Property

	Public Property Let FromName(byVal value)
		pstrFromName = value
	End Property
	Public Property Get FromName()
	    FromName = pstrFromName
	End Property

	Public Property Let [To](byVal value)
		pstrTo = value
	End Property
	Public Property Get [To]()
	    [To] = pstrTo
	End Property

	Public Property Let CC(byVal value)
		pstrCC = value
	End Property
	Public Property Get CC()
	    CC = pstrCC
	End Property

	Public Property Let BCC(byVal value)
		pstrBCC = value
	End Property
	Public Property Get BCC()
	    BCC = pstrBCC
	End Property

	Public Property Let Subject(byVal value)
		pstrSubject = value
	End Property
	Public Property Get Subject()
	    Subject = pstrSubject
	End Property

	Public Property Let Body(byVal value)
		pstrBody = value
	End Property
	Public Property Get Body()
	    Body = pstrBody
	End Property

	Public Property Let HTML(byVal value)
		pblnHTML = value
	End Property
	Public Property Get HTML()
	    HTML = pblnHTML
	End Property

	Public Property Let MailMethod(byVal value)
		pstrMailMethod = value
	End Property
	Public Property Get MailMethod()
	    MailMethod = pstrMailMethod
	End Property

	Public Property Let MailServer(byVal value)

		'added to support mail server login
		Dim paryMailServer
		If InStr(1, value, "|") > 1 Then
			paryMailServer = Split(value, "|")
			pstrMailServer = paryMailServer(0)
			pstrMailServerUserName = paryMailServer(1)
			pstrMailServerPassword = paryMailServer(2)
		Else
			pstrMailServer = value
		End If

	End Property
	Public Property Get MailServer()
	    MailServer = pstrMailServer
	End Property
	Public Property Let MailServerUserName(byVal value)
		pstrMailServerUserName = value
	End Property
	Public Property Let MailServerPassword(byVal value)
		pstrMailServerPassword = value
	End Property

	Public Property Get subjectWithReplacements()
		subjectWithReplacements = makeReplacements(pstrSubject)
	End Property
	
	Public Property Get bodyWithReplacements()
		bodyWithReplacements = makeReplacements(pstrBody)
	End Property

	Public Property Let ShowFailures(byVal value)
		pblnShowFailures = value
	End Property

	'***********************************************************************************************

	Public Sub Initialize()
		pstrFrom = ""
		pstrFromName = ""
		pstrTo = ""
		pstrCC = ""
		pstrBCC = ""
		pstrSubject = ""
		pstrBody = ""
		pblnHTML = False
		pstrMailMethod = ""
		pstrMailServer = ""
		pstrRepeatingItemTag = ""
	End Sub

	'***********************************************************************************************

	Private Function validData()
	
		If Len(pstrFrom) = 0 Then addError "No from address provided"
		If Len(pstrTo) = 0 Then addError "No to address provided"
		If Len(pstrSubject) = 0 Then addError "No subject provided"
		If Len(pstrBody) = 0 Then addError "No body provided"
		
		validData = Not hasErrors

	End Function	'validData

'***********************************************************************************************
'	Replacement Fields
'***********************************************************************************************

	Private Sub InitializeStandardReplacements()
		ReDim paryReplacementFields(40)
		paryReplacementFields(0) = Array("orderNumber", "")

		paryReplacementFields(1) = Array("customerFirstName", "")
		paryReplacementFields(2) = Array("customerLastName", "")
		paryReplacementFields(3) = Array("customerMI", "")
		paryReplacementFields(4) = Array("customerName", "")
		paryReplacementFields(5) = Array("customerAddress", "")
		paryReplacementFields(6) = Array("customerAddress1", "")
		paryReplacementFields(7) = Array("customerAddress2", "")
		paryReplacementFields(8) = Array("customerCity", "")
		paryReplacementFields(9) = Array("customerState", "")
		paryReplacementFields(10) = Array("customerZip", "")
		paryReplacementFields(11) = Array("customerCountry", "")
		paryReplacementFields(12) = Array("customerCountryName", "")
		paryReplacementFields(13) = Array("customerPhone", "")
		paryReplacementFields(14) = Array("customerFax", "")
		paryReplacementFields(15) = Array("customerEmail", "")

		paryReplacementFields(16) = Array("recipientFirstName", "")
		paryReplacementFields(17) = Array("recipientLastName", "")
		paryReplacementFields(18) = Array("recipientMI", "")
		paryReplacementFields(19) = Array("recipientName", "")
		paryReplacementFields(20) = Array("recipientAddress", "")
		paryReplacementFields(21) = Array("recipientAddress1", "")
		paryReplacementFields(22) = Array("recipientAddress2", "")
		paryReplacementFields(23) = Array("recipientCity", "")
		paryReplacementFields(24) = Array("recipientState", "")
		paryReplacementFields(25) = Array("recipientZip", "")
		paryReplacementFields(26) = Array("recipientCountry", "")
		paryReplacementFields(27) = Array("recipientCountryName", "")
		paryReplacementFields(28) = Array("recipientPhone", "")
		paryReplacementFields(29) = Array("recipientFax", "")
		paryReplacementFields(30) = Array("recipientEmail", "")
		
		paryReplacementFields(31) = Array("storeName", "")
		paryReplacementFields(32) = Array("shipAddress", "")
		paryReplacementFields(33) = Array("trackingLink", "")
		paryReplacementFields(34) = Array("backorderMessage", "")
		paryReplacementFields(35) = Array("shipMethod", "")
		paryReplacementFields(36) = Array("dateShipped", "")
		paryReplacementFields(37) = Array("trackingNumber", "")
		paryReplacementFields(38) = Array("trackingMessage", "")
		
		paryReplacementFields(39) = Array("customerCompany", "")
		paryReplacementFields(40) = Array("recipientCompany", "")

		ReDim paryEndCaps(3)
		paryEndCaps(0) = Array("[", "]")
		paryEndCaps(1) = Array("{", "}")
		paryEndCaps(2) = Array("<", ">")
		paryEndCaps(3) = Array("&lt;", "&gt;")
		
		ReDim paryReplacementFields_Repeating(0)
		paryReplacementFields_Repeating(0) = Array("itemCounter", "")

	End Sub	'InitializeStandardReplacements

	'***********************************************************************************************

	Public Sub addCap(byVal strStart, byVal strFinish)
	
	Dim capCount
	
		capCount = UBound(paryEndCaps) + 1
		ReDim Preserve paryEndCaps(capCount)
		paryEndCaps(fieldCounter) = Array(strStart, strFinish)
		
	End Sub	'addCap

	'***********************************************************************************************

	Public Property Get RepeatingItemTemplate()
		RepeatingItemTemplate = pstrRepeatingItemTemplate
	End Property
	
	Public Property Let RepeatingItemTag(byVal Value)
		If Len(Value) > 0 Then Call LoadRepeatingItemTemplate(pstrBody, Value)
		pstrRepeatingItemTag = Value
	End Property

	'***********************************************************************************************

	Public Sub LoadRepeatingItemTemplate(ByVal strTemplate, ByVal strTemplateItemTag)

	Dim cstrStartTag
	Dim cstrEndTag
	Dim orderItemCounter
	Dim plngStart
	Dim plngEnd
	Dim pstrTemplateItem
	Dim pstrSource

		cstrStartTag = "<" & strTemplateItemTag & ">"
		cstrEndTag = "</" & strTemplateItemTag & ">"
		
		plngStart = InStr(1, strTemplate, cstrStartTag)
		If plngStart > 0 Then
			plngEnd = InStr(plngStart, strTemplate, cstrEndTag)
			If plngEnd > 1 Then
				pstrTemplateItem = Mid(strTemplate, plngStart + Len(cstrStartTag), plngEnd - plngStart - Len(cstrEndTag))
				'remove the trailing line feed. If present, it gets doubled up
				If Asc(Right(pstrTemplateItem, 1)) = 13 Then pstrTemplateItem = Left(pstrTemplateItem, Len(pstrTemplateItem) - 1)
				pstrSource = Left(strTemplate, plngStart - 1) & "{" & strTemplateItemTag & "}" & Right(strTemplate, Len(strTemplate) - plngEnd - Len(cstrEndTag))
				pstrRepeatingItemTemplate = pstrTemplateItem
				pstrBody = pstrSource
			End If
		End If
		
	End Sub	'LoadRepeatingItemTemplate

	'***********************************************************************************************

	Public Sub clearReplacementValues_Repeating()
	
	Dim fieldCounter
	
		For fieldCounter = 0 To UBound(paryReplacementFields_Repeating)
			paryReplacementFields_Repeating(fieldCounter)(1) = ""
		Next 'fieldCounter
		
	End Sub	'clearReplacementValues_Repeating

	'***********************************************************************************************

	Public Sub writeReplacementValues_Repeating()
	
	Dim fieldCounter
	
	    Response.Write "<fieldset><legend>ReplacementFields_Repeating Values</legend>"
		For fieldCounter = 0 To UBound(paryReplacementFields_Repeating)
		    Response.Write "(" & fieldCounter & ") " & paryReplacementFields_Repeating(fieldCounter)(0) & ": " & paryReplacementFields_Repeating(fieldCounter)(1) & "<BR>"
		Next 'fieldCounter
	    Response.Write "</fieldset>"
		
	End Sub	'writeReplacementValues_Repeating

	'***********************************************************************************************

	Public Sub ClearRepeatingItemReplacementText()
		pstrRepeatingItemReplacementText = ""
	End Sub	'ClearRepeatingItemReplacementText
	
	'***********************************************************************************************

	Public Sub SetRepeatingItemReplacementText()
	
	Dim p_strTemp
	
	Dim capCount
	Dim capCounter
	Dim fieldCount
	Dim fieldCounter
	Dim pLeftCap
	Dim pRightCap
	
		If Len(pstrRepeatingItemTemplate) = 0 Then Exit Sub
		
		p_strTemp = pstrRepeatingItemTemplate
		capCount = UBound(paryEndCaps)
		fieldCount = UBound(paryReplacementFields_Repeating)
		For capCounter = 0 To UBound(paryEndCaps)
			For fieldCounter = 0 To fieldCount
				p_strTemp = Replace(p_strTemp, paryEndCaps(capCounter)(0) & paryReplacementFields_Repeating(fieldCounter)(0) & paryEndCaps(capCounter)(1), paryReplacementFields_Repeating(fieldCounter)(1))
			Next 'fieldCounter
		Next 'capCounter
		
		'run the repeating item template through the top level replacement routine
		pstrRepeatingItemReplacementText = pstrRepeatingItemReplacementText & makeReplacements(p_strTemp)
		'response.Write "pstrRepeatingItemReplacementText: " & pstrRepeatingItemReplacementText & "<BR>"

	End Sub	'SetRepeatingItemReplacementText
	
	'***********************************************************************************************

	Public Sub setReplacementValue_Repeating(byVal strField, byVal strValue)
	
	Dim fieldCount
	Dim fieldCounter
	Dim pblnFieldFound

		pblnFieldFound = False
		fieldCount = UBound(paryReplacementFields_Repeating)
		For fieldCounter = 0 To fieldCount
			If paryReplacementFields_Repeating(fieldCounter)(0) = strField Then
				paryReplacementFields_Repeating(fieldCounter)(1) = strValue
				pblnFieldFound = True
				Exit For
			End If
		Next 'fieldCounter
		
		If Not pblnFieldFound Then
			ReDim Preserve paryReplacementFields_Repeating(fieldCounter)
			paryReplacementFields_Repeating(fieldCounter) = Array(strField, strValue)
		End If
		
	End Sub	'setReplacementValue_Repeating

	'***********************************************************************************************

	Public Function getReplacementValue_Repeating(byVal strField)
	
	Dim fieldCount
	Dim fieldCounter
	Dim pstrValue

		fieldCount = UBound(paryReplacementFields_Repeating)
		For fieldCounter = 0 To fieldCount
			If paryReplacementFields_Repeating(fieldCounter)(0) = strField Then
				pstrValue = paryReplacementFields_Repeating(fieldCounter)(1)
			End If
		Next 'fieldCounter
		
		getReplacementValue_Repeating = pstrValue
		
	End Function	'getReplacementValue_Repeating

	'***********************************************************************************************

	Public Sub clearReplacementValues()
	
	Dim fieldCounter
	
		For fieldCounter = 0 To UBound(paryReplacementFields)
			paryReplacementFields(fieldCounter)(1) = ""
		Next 'fieldCounter
		
	End Sub	'clearReplacementValues

	'***********************************************************************************************

	Public Sub writeReplacementValues()
	
	Dim fieldCounter
	
	    Response.Write "<fieldset><legend>Replacement Values</legend>"
		For fieldCounter = 0 To UBound(paryReplacementFields)
		    Response.Write "(" & fieldCounter & ") " & paryReplacementFields(fieldCounter)(0) & ": " & paryReplacementFields(fieldCounter)(1) & "<BR>"
		Next 'fieldCounter
	    Response.Write "</fieldset>"
		
	End Sub	'writeReplacementValues

	'***********************************************************************************************

	Public Sub setReplacementValue(byVal strField, byVal strValue)
	
	Dim fieldCount
	Dim fieldCounter
	Dim pblnFieldFound

		pblnFieldFound = False
		fieldCount = UBound(paryReplacementFields)
		For fieldCounter = 0 To fieldCount
			If paryReplacementFields(fieldCounter)(0) = strField Then
				paryReplacementFields(fieldCounter)(1) = strValue
				pblnFieldFound = True
				Exit For
			End If
		Next 'fieldCounter
		
		If Not pblnFieldFound Then
			ReDim Preserve paryReplacementFields(fieldCounter)
			paryReplacementFields(fieldCounter) = Array(strField, strValue)
		End If
		
	End Sub	'setReplacementValue

	'***********************************************************************************************

	Public Function getReplacementValue(byVal strField)
	
	Dim fieldCount
	Dim fieldCounter
	Dim pstrValue

		fieldCount = UBound(paryReplacementFields)
		For fieldCounter = 0 To fieldCount
			If paryReplacementFields(fieldCounter)(0) = strField Then
				pstrValue = paryReplacementFields(fieldCounter)(1)
			End If
		Next 'fieldCounter
		
		getReplacementValue = pstrValue
		
	End Function	'getReplacementValue

	'***********************************************************************************************

	Public Function RecipientName()
		RecipientName = getReplacementValue("recipientFirstName") & " " & getReplacementValue("recipientMI") & " " & getReplacementValue("recipientLastName")
	End Function    'RecipientName

	'***********************************************************************************************

	Public Function makeReplacements(ByVal strSource)

	Dim p_strTemp
	Dim pstrShipAddr
	
	Dim capCount
	Dim capCounter
	Dim fieldCount
	Dim fieldCounter
	Dim pLeftCap
	Dim pRightCap
	
		'build the standard replacements
		Call setReplacementValue("customerName", Replace(getReplacementValue("customerFirstName") & " " & getReplacementValue("customerMI") & " " & getReplacementValue("customerLastName"),"  "," "))
		pstrShipAddr = getReplacementValue("customerAddress2") & vbcrlf
		If Len(pstrShipAddr) = 0 Then
			pstrShipAddr = getReplacementValue("customerAddress1") & vbcrlf _
						 & getReplacementValue("customerCity") & ", " & getReplacementValue("customerState") & " " & getReplacementValue("customerZip")
		Else
			pstrShipAddr = getReplacementValue("customerAddress1") & vbcrlf _
						 & pstrShipAddr & vbcrlf _
						 & getReplacementValue("customerCity") & ", " & getReplacementValue("customerState") & " " & getReplacementValue("customerZip")
		End If
		Call setReplacementValue("customerAddress", pstrShipAddr)
		
		Call setReplacementValue("recipientName", Replace(RecipientName,"  "," "))

		pstrShipAddr = getReplacementValue("recipientAddress2")
		If Len(pstrShipAddr) = 0 Then
			pstrShipAddr = getReplacementValue("recipientAddress1") & vbcrlf _
						 & getReplacementValue("recipientCity") & ", " & getReplacementValue("recipientState") & " " & getReplacementValue("recipientZip")
		Else
			pstrShipAddr = getReplacementValue("recipientAddress1") & vbcrlf _
						 & pstrShipAddr & vbcrlf _
						 & getReplacementValue("recipientCity") & ", " & getReplacementValue("recipientState") & " " & getReplacementValue("recipientZip")
		End If
		Call setReplacementValue("recipientAddress", pstrShipAddr)
		
		If Len(pstrRepeatingItemTag) > 0 Then Call setReplacementValue(pstrRepeatingItemTag, pstrRepeatingItemReplacementText)

		p_strTemp = strSource
		capCount = UBound(paryEndCaps)
		fieldCount = UBound(paryReplacementFields)
		For capCounter = 0 To UBound(paryEndCaps)
			For fieldCounter = 0 To fieldCount
				p_strTemp = Replace(p_strTemp, paryEndCaps(capCounter)(0) & paryReplacementFields(fieldCounter)(0) & paryEndCaps(capCounter)(1), paryReplacementFields(fieldCounter)(1))
			Next 'fieldCounter
		Next 'capCounter
			
		makeReplacements = p_strTemp

	End Function	'makeReplacements

	'***********************************************************************************************

	Public Function Send()

	Dim i
	Dim paryAddress
	Dim pblnSuccess
	Dim pobjMail
	Dim pstrTempAddress

		pblnSuccess = True
		
		If validData() Then
			If Len(pstrFromName) = 0 Then pstrFromName = pstrFrom
			
			On Error Resume Next
			
			Select Case pstrMailMethod
				Case "ASP Mail"
					Set pobjMail = CreateObject ("smtpsvg.mailer")
					With pobjMail
						If pblnHTML Then
							.ContentType = "text/html"
						End If
						.QMessage = TRUE
						.RemoteHost = pstrMailServer
						.AddRecipient pstrTo, pstrTo
						.FromAddress = pstrFrom
						.FromName = pstrFromName
						.Subject = subjectWithReplacements
						.BodyText = bodyWithReplacements
						.SendMail
					End With
					Set pobjMail = Nothing
				Case "CDOSYS"
					Dim Configuration 'As New CDO.Configuration
					Dim Fields 'As ADODB.Fields
					
					Set Configuration = CreateObject("CDO.Configuration")
					'cdoSendUsingMethod, cdoSMTPServerPort, cdoSMTPServer, cdoSMTPConnectionTimeout, cdoSMTPAuthenticate, cdoURLProxyServer, cdoURLProxyBypass, cdoURLGetLatestVersion
					Set Fields = Configuration.Fields
					With Fields
						Fields(cdoSendUsingMethod) = 2
						'Fields(cdoSendUsingMethod) = 1
						Fields(cdoSMTPServerPort) = 25
						Fields(cdoSMTPServer) = pstrMailServer
						Fields(cdoSMTPConnectionTimeout) = 20
 						'Fields(cdoSMTPAuthenticate)      = 0
  						'Fields(cdoURLProxyServer)        = "server:80"
  						'Fields(cdoURLProxyBypass)        = "<local>"
  						Fields(cdoURLGetLatestVersion)   = True
						Fields(cdoSendUserName)         = pstrMailServerUserName
  						Fields(cdoSendPassword)         = pstrMailServerPassword
						.Update
					End With

					Set pobjMail = CreateObject("CDO.Message")
					With pobjMail
						Set .Configuration = Configuration
						.From = pstrFrom 
						.Subject= subjectWithReplacements
						
						If inStr(1, bodyWithReplacements, "<html") Or inStr(1, bodyWithReplacements, "<HTML") Then
							.HTMLBody = bodyWithReplacements 
						Else
							.TextBody = bodyWithReplacements 
						End If

						If Len(pstrTo) > 0 Then
							paryAddress = Split(pstrTo, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .To = pstrTempAddress
							Next 'i
						End If

						If Len(pstrCC) > 0 Then
							paryAddress = Split(pstrCC, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .Cc = pstrTempAddress
							Next 'i
						End If

						If Len(pstrBCC) > 0 Then
							paryAddress = Split(pstrBCC, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .Bcc = pstrTempAddress
							Next 'i
						End If

						.Send 
					End With
					Set pobjMail = Nothing
					Set Configuration = Nothing
					Set Fields = Nothing
				Case "CDONTS Mail"
					Set pobjMail = CreateObject("CDO.Message")
					With pobjMail
						.From = pstrFrom 
						.Subject= subjectWithReplacements
						
						If inStr(1, bodyWithReplacements, "<html") Or inStr(1, bodyWithReplacements, "<HTML") Then
							.HTMLBody = bodyWithReplacements 
						Else
							.TextBody = bodyWithReplacements 
						End If

						If Len(pstrTo) > 0 Then
							paryAddress = Split(pstrTo, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .To = pstrTempAddress
							Next 'i
						End If

						If Len(pstrCC) > 0 Then
							paryAddress = Split(pstrCC, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .Cc = pstrTempAddress
							Next 'i
						End If

						If Len(pstrBCC) > 0 Then
							paryAddress = Split(pstrBCC, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .Bcc = pstrTempAddress
							Next 'i
						End If

						.Send 
					End With
					Set pobjMail = Nothing
				Case "CDONTS Mail Win2000"
					Set pobjMail = CreateObject("CDONTS.NewMail")
					With pobjMail
						If pblnHTML Then
							.BodyFormat = 0
							.MailFormat = 0
						Else
							.BodyFormat = 1
							.MailFormat = 1
						End If

						'.Importance = 
						If Len(pstrCC) > 0 Then
							paryAddress = Split(pstrCC, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .Cc = pstrTempAddress
							Next 'i
						End If

						If Len(pstrBCC) > 0 Then
							paryAddress = Split(pstrBCC, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .Bcc = pstrTempAddress
							Next 'i
						End If

						.Send pstrFrom, pstrTo, subjectWithReplacements, bodyWithReplacements
					End With
					Set pobjMail = Nothing
				Case "J Mail"
					Set pobjMail = CreateObject ("JMail.SMTPMail")
					With pobjMail
						If pblnHTML Then
							.ContentType = "text/html"
						End If
						.ServerAddress = pstrMailServer
						.Sender = pstrFrom
						.SenderName = pstrFromName
						.AddRecipientEx pstrTo, pstrTo

						If Len(pstrCC) > 0 Then
							paryAddress = Split(pstrCC, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .AddRecipientCC = pstrTempAddress
							Next 'i
						End If

						If Len(pstrBCC) > 0 Then
							paryAddress = Split(pstrBCC, ";")
							For i = 0 To UBound(paryAddress)
								pstrTempAddress = paryAddress(i)
								If Len(pstrTempAddress) > 0 Then .AddRecipientBCC = pstrTempAddress
							Next 'i
						End If

						.Subject = subjectWithReplacements

						If inStr(1, bodyWithReplacements, "<html") Or inStr(1, bodyWithReplacements, "<HTML") Then
							.HTMLBody = bodyWithReplacements 
						Else
							.Body = bodyWithReplacements
						End If

						.Execute
					End With
					Set pobjMail = Nothing
				Case "Simple Mail 3.1"
					Set pobjMail = CreateObject ("ADISCON.SimpleMail.1")
					With pobjMail
						.MailServer = pstrMailServer
						.Sender = pstrFrom
						.SenderName = pstrFromName
						.Recipient = pstrTo
						.Subject = subjectWithReplacements
						If pblnHTML Then
							.ContentType = "text/html"
							.MessageHTMLText = bodyWithReplacements
						Else
							.MessageText = bodyWithReplacements
						End If
						.Send
					End With
					Set pobjMail = Nothing
				Case "sendRemote"
					pblnSuccess = sendRemoteEmail(False)
				Case "sendRemoteNow"
					pblnSuccess = sendRemoteEmail(True)
				Case "writeHTML"
					addMessage "Mail Method - Display Only Selected"
					Response.Write mailAsHTML
					pblnSuccess = True
				Case Else
					'this dumps into the modified mail.asp for SF5
					addMessage "Mail Method - Invalid mail method selected (<em>" & pstrMailMethod & "</em>)"
					pblnSuccess = False
			End Select

			If Err.number <> 0 Then
				If Err.number = 424 Then
					addError "Could not create  " & pstrMailMethod & " object. Check to make sure this mail component is properly installed."
				Else
					addError "Error " & Err.number & ": " & Err.Description
				End If
				Err.Clear
				pblnSuccess = False
			End If
			
			If pblnShowFailures And Not pblnSuccess Then Response.Write mailAsHTML
			'Response.Write mailAsHTML
		Else
			pblnSuccess = False
			'Response.Write "<fieldset><legend>Email Errors</legend>" & writeErrorMessages & "</fieldset>"
		End If	'validData()
		
		Send = pblnSuccess

	End Function	'Send

	'***********************************************************************************************

	Public Function mailAsHTML()
	
	Dim pstrOut
	
		pstrOut = "<table border=""1"" cellspacing=""0"" cellpadding=""3"" style=""border-color:black;border-collapse:collapse"">" & vbcrlf _
				& "<tr><th colspan=2>" & writeErrorMessages & "</th></tr>" & vbcrlf _
				& "<tr><td align=right>Method:&nbsp;</td><td align=left>&nbsp;" & pstrMailMethod & "</td></tr>" & vbcrlf _
				& "<tr><td align=right>Server:&nbsp;</td><td align=left>&nbsp;" & pstrMailServer & "</td></tr>" & vbcrlf _
				& "<tr><td align=right>From:&nbsp;</td><td align=left>&nbsp;" & pstrFrom & "</td></tr>" & vbcrlf _
				& "<tr><td align=right>To:&nbsp;</td><td align=left>&nbsp;" & pstrTo & "</td></tr>" & vbcrlf _
				& "<tr><td align=right>CC:&nbsp;</td><td align=left>&nbsp;" & pstrCC & "</td></tr>" & vbcrlf _
				& "<tr><td align=right>Bcc:&nbsp;</td><td align=left>&nbsp;" & pstrBCC & "</td></tr>" & vbcrlf _
				& "<tr><td align=right>Subject:&nbsp;</td><td align=left>&nbsp;" & subjectWithReplacements & "</td></tr>" & vbcrlf
		If pblnHTML Then
			pstrOut = pstrOut & "<tr><td align=right valign=top>Body:&nbsp;</td><td align=left>&nbsp;" & Replace(bodyWithReplacements, vbcrlf,"<br />") & "</td></tr>" & vbcrlf
		Else
			pstrOut = pstrOut & "<tr><td align=right valign=top>Body:&nbsp;</td><td align=left>&nbsp;<pre>" & bodyWithReplacements & "</pre></td></tr>" & vbcrlf
		End If
		pstrOut = pstrOut & "</table>"
		
		mailAsHTML = pstrOut

	End Function	'mailAsHTML

	'***************************************************************************************************************************************************************

	Dim xmlDoc
	Dim xmlRoot
	Dim pstrXMLResult

	'***************************************************************************************************************************************************************

	Public Property Get remoteEmailResults()
	    remoteEmailResults = pstrXMLResult
	End Property

	'***************************************************************************************************************************************************************

	Private Function sendRemoteEmail(byVal blnSendNow)

	Dim xmlNode
	Dim xmlEmail
	Dim pstrData
	Dim pstrResult
	Dim pblnSuccess

		pblnSuccess = True
		
		If Not isObject(xmlDoc) Then
			set xmlDoc = CreateObject("MSXML2.DOMDocument")
			' Create processing instruction and document root
			Set xmlNode = xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
			Set xmlNode = xmlDoc.insertBefore(xmlNode, xmlDoc.childNodes.Item(0))
		    
			' Create document root
			Set xmlRoot = xmlDoc.createElement("emails")
			Set xmlDoc.documentElement = xmlRoot
			xmlRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"
			Set xmlNode = Nothing
		End If	'Not isObject(xmlDoc)

		'Create email Node
		Set xmlEmail = xmlDoc.createElement("email")
		xmlRoot.appendChild xmlEmail

		If pblnHTML Then
			Call addNode(xmlDoc, xmlEmail, "format", "HTML")
		Else
			Call addNode(xmlDoc, xmlEmail, "format", "")
		End If	'pblnHTML
		
		Call addNode(xmlDoc, xmlEmail, "from", pstrFrom)
		Call addNode(xmlDoc, xmlEmail, "fromName", pstrFromName)
		Call addNode(xmlDoc, xmlEmail, "to", pstrTo)
		Call addNode(xmlDoc, xmlEmail, "cc", pstrCC)
		Call addNode(xmlDoc, xmlEmail, "bcc", pstrBCC)
		Call addNode(xmlDoc, xmlEmail, "subject", subjectWithReplacements)
		Call addNode(xmlDoc, xmlEmail, "body", bodyWithReplacements)

		Set xmlEmail = Nothing
		
		If blnSendNow Then
			pstrData = "action=sendEmail" _
					 & "&data=" & Server.URLEncode(xmlDoc.xml)
							 
			pstrXMLResult = RetrieveRemoteData(pstrMailServer, pstrData, True)		 
			pblnSuccess = processEmailReturn(pstrXMLResult)
		End If	'blnSendNow
		sendRemoteEmail = pblnSuccess

	End Function	'sendRemoteEmail

	'***************************************************************************************************************************************************************

	Private Function processEmailReturn(byVal strXML)

	Dim objXMLDoc
	Dim pstrData
	Dim objNodeList
	Dim i
	Dim pstrResult
	Dim pblnSuccess

		'debug.print "strXML", strXML
		pblnSuccess = True
		Set objXMLDoc = CreateObject("MSXML.DOMDocument")
		If objXMLDoc.LoadXML(strXML) Then
			Set objNodeList = objXMLDoc.GetElementsByTagName("Attempt")
			For i = 0 To objNodeList.Length - 1
				pstrResult = GetXMLElementValue(objNodeList(i), "Result")
				If InStr(1, pstrResult, "Fail:") > 0 Then
					pblnSuccess = False
					addError GetXMLElementValue(objNodeList(i), "email") & " - " & pstrResult
				End If
			Next 'i
			Set objNodeList = Nothing
		Else
			pblnSuccess = False
			addError "Unable to process email return - " & strXML
		End If
		Set objXMLDoc = Nothing
		
		processEmailReturn = pblnSuccess

	End Function	'processEmailReturn

	'***************************************************************************************************************************************************************

	Public Function DeleteEmailTemplate(ByVal strDirectoryPath, ByVal strFileToDelete)

	Dim pobjFSO
	Dim pblnResult
	
		pblnResult = False
		
		Set pobjFSO = CreateObject("Scripting.FileSystemObject")
		If pobjFSO.FolderExists(strDirectoryPath) Then
			If pobjFSO.FileExists(strDirectoryPath & strFileToDelete) Then
				pobjFSO.DeleteFile(strDirectoryPath & strFileToDelete)
				pblnResult = CBool(Err.number = 0)
				If Err.number <> 0 Then
					Response.Write "Error in ssclsEmail.asp:DeleteEmailTemplate. Error " & err.number & ": " & err.Description & "<br />"
				End If
			End If
		End If
		Set pobjFSO = Nothing
		
		DeleteEmailTemplate = pblnResult

	End Function	'DeleteEmailTemplate
	
	'***************************************************************************************************************************************************************

	Public Function EmailFileText(ByVal strTemplateName, ByVal strSubject, ByVal strBody)

	Dim pstrEmailFileText
	
		pstrEmailFileText = strTemplateName & vbcrlf _
						& "// Template Name Separator - DO NOT REMOVE THIS LINE //" & vbcrlf _
						& strSubject & vbcrlf _
						& "// Subject/Body Separator - DO NOT REMOVE THIS LINE //" & vbcrlf _
						& strBody & vbcrlf _
						& "// DO NOT REMOVE THIS LINE //" & vbcrlf
		
		EmailFileText = pstrEmailFileText

	End Function	'EmailFileText
	
	'***************************************************************************************************************************************************************

	Public Function UpdateEmailTemplate(ByVal strDirectoryPath, ByVal strFileToUpdate, ByVal lngIndex, ByRef aryEmails)

	Dim pobjFile
	Dim pobjFSO
	Dim pblnResult
	
		pblnResult = False
		
		Set pobjFSO = CreateObject("Scripting.FileSystemObject")
		If pobjFSO.FolderExists(strDirectoryPath) Then
			Set pobjFile = pobjFSO.CreateTextFile(strDirectoryPath & strFileToUpdate, True)
			pobjFile.Write EmailFileText (aryEmails(lngIndex)(3), aryEmails(lngIndex)(1), aryEmails(lngIndex)(2))
			pblnResult = CBool(Err.number = 0)
			pobjFile.Close
			Set pobjFile = Nothing
		End If
		Set pobjFSO = Nothing
		
		UpdateEmailTemplate = pblnResult

	End Function	'UpdateEmailTemplate
	
	'***************************************************************************************************************************************************************

	Public Function LoadEmailTemplates(ByVal strDirectoryPath, ByVal strDefaultFileToLoad, ByRef aryEmails)
	'Returns aryEmails(i)(3) where (3) --> Array("fileName", "subject", "body", "TemplateName")
	
	Dim pobjFSO
	Dim pobjFolder, pobjFiles
	Dim i
	Dim MyFile
	Dim p_strSubject
	Dim p_strBody
	Dim p_strTemplateName
	Dim pblnSuccess
	
	'On Error Resume Next

		pblnSuccess = True
		
		Set pobjFSO = CreateObject("Scripting.FileSystemObject")
		If InStr(1, strDirectoryPath, strDefaultFileToLoad) < 1 Or Len(strDefaultFileToLoad) = 0 Then
			i = 0
			
			On Error Resume Next
			If Err.number <> 0 Then Err.Clear
			Set pobjFolder = pobjFSO.GetFolder(strDirectoryPath)
			
			If Err.number = 0 Then
				On Error Goto 0
				Set pobjFiles = pobjFolder.Files
				ReDim aryEmails(pobjFiles.Count - 1)
				For Each MyFile In pobjFiles
					p_strSubject = ""
					p_strBody = ""

					aryEmails(i) = Array("fileName", "subject", "body", "TemplateName")
					
					aryEmails(i)(enEmail_FileName) = MyFile.Name
					p_strTemplateName = aryEmails(i)(enEmail_FileName)
					Set MyFile =pobjFSO.OpenTextFile(strDirectoryPath & MyFile.Name,1,True)
					Call ParseEmailTemplate(MyFile, p_strSubject, p_strBody, p_strTemplateName)
					MyFile.Close
					Set MyFile = Nothing
					
					aryEmails(i)(enEmail_Subject) = p_strSubject
					aryEmails(i)(enEmail_Body) = p_strBody
					aryEmails(i)(enEmail_TemplateName) = p_strTemplateName
					
					If LCase(aryEmails(i)(enEmail_FileName)) = LCase(strDefaultFileToLoad) Then
						pstrSubject = aryEmails(i)(enEmail_Subject)
						pstrBody = aryEmails(i)(enEmail_Body)
					End If
					
					i = i + 1
				Next 'MyFile
				Set pobjFiles = Nothing
			Else
				pblnSuccess = False
				addError "Error opening email templates. Error " & Err.number & " - " & Err.Description & "<br />" _
						& "Path: " & strDirectoryPath
			End If
			
			Set pobjFolder = Nothing
		Else
			p_strSubject = ""
			p_strBody = ""

			Set MyFile = pobjFSO.OpenTextFile(strDirectoryPath,1,True)
			Call ParseEmailTemplate(MyFile, p_strSubject, p_strBody, p_strTemplateName)
			MyFile.Close
			Set MyFile = Nothing
			
			ReDim aryEmails(0)
			aryEmails(0) = Array(strDefaultFileToLoad, p_strSubject, p_strBody, p_strTemplateName)
			pstrSubject = aryEmails(0)(enEmail_Subject)
			pstrBody = aryEmails(0)(enEmail_Body)
		End If
		Set pobjFSO = Nothing
		
		LoadEmailTemplates = pblnSuccess

	End Function	'LoadEmailTemplates

	'***************************************************************************************************************************************************************

	Private Sub ParseEmailTemplate(ByRef objFile, ByRef strSubject, ByRef strBody, ByRef strTemplateName)

	Dim pstrTempLine
	
		With objFile
			If Not .AtEndOfStream Then strSubject = .ReadLine
			If Not .AtEndOfStream Then pstrTempLine = .ReadLine	'garbage line
			
			
			If pstrTempLine = "// Template Name Separator - DO NOT REMOVE THIS LINE //" Then
				If Not .AtEndOfStream Then strTemplateName = strSubject
				If Not .AtEndOfStream Then strSubject = .ReadLine
				If Not .AtEndOfStream Then pstrTempLine = .ReadLine	'garbage line
			End If
			
			If Not .AtEndOfStream Then pstrTempLine = .ReadLine & vbcrlf
			
			If Len(strSubject) = 0 Then
				Response.Write "alert('" & strTemplateName & " is in an invalid format.');"
			End If
			
			Do While pstrTempLine <> "// DO NOT REMOVE THIS LINE //" AND NOT .AtEndOfStream
				strBody = strBody & pstrTempLine & vbcrlf
				pstrTempLine = .ReadLine
			Loop
		End With	'objFile
				
	End Sub	'LoadEmailFiles

End Class   'clsEmail

'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************

Class clsReplacement

Dim pdicReplacement

'***********************************************************************************************

Private Sub class_Initialize()
	Set pdicReplacement = CreateObject("scripting.dictionary")
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	If isObject(pdicReplacement) Then Set pdicReplacement = Nothing
End Sub

'***********************************************************************************************

Public Sub setReplacement(byVal strName, byVal strValue)
	If pdicReplacement.Exists(strName) Then
		pdicReplacement(strName) = strValue
	Else
		pdicReplacement.Add strName, strValue
	End If
End Sub

'***********************************************************************************************

Public Function getReplacment(byVal strSource)

Dim vItem
Dim pstrOut

	pstrOut = strSource
	For Each vItem in pdicReplacement
		pstrOut = Replace(pstrOut, vItem, pdicReplacement(vItem))
	Next
	
	getReplacment = pstrOut
	
End Function

End Class   'clsReplacement
%>
