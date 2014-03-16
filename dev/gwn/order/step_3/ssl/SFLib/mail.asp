<!--#include file="ssclsEmail.asp"--> 
<% 
'********************************************************************************
'*
'*   mail.asp
'*   Release Version:	1.00.001
'*   Release Date:		January 10, 2003
'*   Revision Date:		February 5, 2003
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins is confirm.asp
'*   from LaGarde, Incorporated. It has been heavily modified by Sandshot Software
'*   with permission from LaGarde, Incorporated who has NOT released any of the 
'*   original copyright protections. As such, this is a derived work covered by
'*   the copyright provisions of both companies.
'*
'*   LaGarde, Incorporated Copyright Statement
'*   The contents of this file is protected under the United States copyright
'*   laws and is confidential and proprietary to LaGarde, Incorporated.  Its 
'*   use ordisclosure in whole or in part without the expressed written 
'*   permission of LaGarde, Incorporated is expressly prohibited.
'*   (c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'*   
'*   Sandshot Software Copyright Statement
'*   The contents of this file are protected by United States copyright laws 
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.
'********************************************************************************

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Page Level variables
'**********************************************************

'**********************************************************
'*	Functions
'**********************************************************

'**********************************************************
'*	Begin Page Code
'**********************************************************

Sub createMail(byVal sType, byVal sInformation)

Dim paryInfo
Dim pstrEmailTo
Dim pstrEmailMessage
Dim pstrPrimaryMerchantEmail
Dim pstrSecondaryMerchantEmail
Dim pstrSubject

	'Load default settings, may be overwritten by inbound data
	pstrPrimaryMerchantEmail		= adminPrimaryEmail
	pstrSecondaryMerchantEmail		= adminSecondaryEmail
	pstrSubject		= adminEmailSubject
	pstrEmailMessage= adminEmailMessage
	
	paryInfo = split(sInformation, "|")
	
	If False Then
		Dim i
		Response.Write "<fieldset><legend>createMail</legend>"
		Response.Write "sType:  " & sType & "<br />"
		for i = 0 to ubound(paryInfo)
		Response.Write i & ":  " & paryInfo(i) & "<br />"
		next
		Response.Write "</fieldset>"
	End If
	
	Select Case sType
		Case "EmailFriend"
			pstrEmailTo = paryInfo(0)
			pstrPrimaryMerchantEmail = paryInfo(1)
			pstrSecondaryMerchantEmail = paryInfo(5)
			pstrSubject = paryInfo(4)
			pstrEmailMessage = paryInfo(2) & vbcrlf _
							 & adminDomainName & "detail.asp?product_id=" & server.urlencode((paryInfo(3)))
		Case "EmailWishList"
			pstrEmailTo = paryInfo(0)
			pstrPrimaryMerchantEmail = paryInfo(1)
			pstrSecondaryMerchantEmail = paryInfo(5)
			pstrSubject = paryInfo(4)
			pstrEmailMessage = paryInfo(2)
		Case "FPWD"
			pstrEmailTo = paryInfo(0)
			pstrSubject = "Requested Password"
			pstrEmailMessage = "Here is the password you requested." _
							 & "Your password for the e-mail account : " & pstrEmailTo & " is : " & paryInfo(1) & VbCrLf _
							 & "You may use it for login to " & adminDomainName & "MyAccount.asp?Action=Login&Email=" & pstrEmailTo & "&Password=" & paryInfo(1)
		Case "InvenNotification"
			pstrEmailTo = pstrPrimaryMerchantEmail 
			pstrSubject = paryInfo(0)
			pstrEmailMessage = paryInfo(1)
		Case "PromoMail"
			pstrEmailTo = paryInfo(0)
			pstrSubject = paryInfo(1)
			pstrEmailMessage = paryInfo(2) & vbcrlf _
							 & "To Remove yourself from the mailing list please go to this link:" & VbCrLf _ 
							 & adminDomainName & "unsubscribe.asp?email=" & paryInfo(0)
		Case Else
			pstrEmailTo = paryInfo(0)
			If ((len(paryInfo(1))>0) AND (paryInfo(1)<>"-")) Then
				pstrPrimaryMerchantEmail = paryInfo(1)
			ElseIf (paryInfo(1)="-") Then
				pstrPrimaryMerchantEmail = ""
			End If
			
			If ((len(paryInfo(2))>0) AND (paryInfo(2)<>"-")) Then
				pstrSecondaryMerchantEmail = paryInfo(2)
			ElseIf (paryInfo(2)="-") Then
				pstrSecondaryMerchantEmail = ""
			End If
			
			pstrSubject = paryInfo(3)
			pstrEmailMessage = paryInfo(4)

	End Select	'sType
	
	If sType = "PromoMail" Then pstrSecondaryMerchantEmail = ""	'suppress CC email for promo mails
		
	Dim pclsEmail
	Set pclsEmail = New clsEmail
	With pclsEmail
		.MailMethod = adminMailMethod
		.MailServer = adminMailServer
		.ShowFailures = True
		
		.From = pstrPrimaryMerchantEmail
		.To = pstrEmailTo
		.CC = ""
		.Subject = pstrSubject
		.Body = pstrEmailMessage
		.Send
		
	End With	'pclsEmail
	Set pclsEmail = Nothing

	If Err.number <> 0 Then Err.Clear
	
End Sub	'createMail

%>