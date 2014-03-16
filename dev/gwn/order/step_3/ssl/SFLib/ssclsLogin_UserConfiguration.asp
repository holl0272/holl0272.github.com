<%
'********************************************************************************
'*   Page Protector for StoreFront 5.0                                          *
'*   Release Version:   1.00.002                                                *
'*   Included with ssclsLogin.asp v1.00.002										*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

	'////////////////////////////////////////////////////////////////////////////////
	'//
	'//		USER CONFIGURATION

		'database mapping
		cstrTableName = "sfCustomers"							
		cstrFieldUserID = "custID"
		cstrFieldUserName = "custEmail"
		cstrFieldEmail = "custEmail"
		cstrFieldPassword = "custPasswd"
		cstrFieldGreeting = "custLastName" 
		
		'email password
		cstrEmailPasswordFrom = "PasswordReminder@mySite.com"
		cstrEmailPasswordSubject = "mySite Account Access - Password Reminder"
		cstrLoginURL = adminDomainName & "MyAccount.asp?Action=Login&Email=<<Email>>&Password=<<Password>>"
		cstrEmailPasswordBody = "This is a reminder for your account access on mySite.com." & vbcrlf _
							  & vbcrlf _
							  & "Email: <<Email>>" & vbcrlf _
							  & "Password: <<Password>>" & vbcrlf _
							  & vbcrlf _
							  & "You may login by visiting:" & vbcrlf _
							  & vbcrlf _
							  & "<<LoginURL>>" & vbcrlf _
							  & vbcrlf _
							  & "Please keep this for future reference."
		
	'//
	'////////////////////////////////////////////////////////////////////////////////
	
		pblnPLM = False

%>