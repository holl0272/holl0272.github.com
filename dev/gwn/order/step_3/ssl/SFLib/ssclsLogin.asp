<%
'********************************************************************************
'*   Page Protector for StoreFront 5.0                                          *
'*   Release Version:   1.00.003                                                *
'*   Release Date:      August 1, 2002											*
'*   Revision Date:     April 30, 2004											*
'*                                                                              *
'*   Revision History															*
'*                                                                              *
'*   1.00.003 (April 30, 2004)													*
'*   - SQL Injection review                                                     *
'*                                                                              *
'*   1.00.002 (August 6, 2003)													*
'*   - General clean-up                                                         *
'*   - Split out login and user configuration files                             *
'*   - Bug Fix - login would not acknowledge SF prior order history             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Class clsLogin

Private cstrTableName
Private cstrFieldUserID
Private cstrFieldUserName
Private cstrFieldEmail
Private cstrFieldPassword
Private cstrFieldGreeting

Private cstrEmailPasswordFrom
Private cstrEmailPasswordSubject
Private cstrLoginURL
Private cstrEmailPasswordBody

Private pstrUserID
Private pstrUsername
Private pstrPassword
Private pstrEmail
Private pstrGreeting

Private pblnPLM

Private pblnRememberMe
Private pstrRedirectPage

	'****************************************************************************************************************

	Private Sub Class_Initialize
	%>
	<!--#include file="ssclsLogin_UserConfiguration.asp"--> 
	<%
	End Sub

	'****************************************************************************************************************

	Private Sub Class_Terminate

	End Sub

	'****************************************************************************************************************

	Public Property Let UserID(strUserID)
		pstrUserID = strUserID
	End Property
	Public Property Get UserID
		UserID = pstrUserID
	End Property
	
	Public Property Let Username(strUsername)
		pstrUsername = strUsername
	End Property
	Public Property Get Username
		Username = pstrUsername
	End Property
	
	Public Property Let Password(strPassword)
		pstrPassword = strPassword
	End Property
	Public Property Get Password
		Password = pstrPassword
	End Property
	
	Public Property Let Email(strEmail)
		pstrEmail = strEmail
	End Property
	Public Property Get Email
		Email = pstrEmail
	End Property
	
	Public Property Get Greeting
		Greeting = pstrGreeting
	End Property
	
	Public Property Let RedirectPage(strRedirectPage)
		pstrRedirectPage = strRedirectPage
	End Property
	
	'***********************************************************************************************

	Public Function EmailPassword(byVal strUsername)

	Dim pobjRS
	Dim pstrSQL
	Dim Mailer
	Dim pstrBody
	Dim pstrTemp
	
	'On Error Resume Next

		pstrUsername = strUsername
		If len(pstrUsername) = 0 then
			EmailPassword = "<h3><b>Please enter an Email address</b></h3>"
			Exit Function
		ElseIf Not checkInput(pstrUsername, pstrPassword) Then
			EmailPassword = "<h3><b>There was a problem with your username. Please contact the system administrator.</b></h3>"
			Exit Function
		End If

		pstrSQL = "Select [" & cstrFieldUserID & "] as ssUserID, [" & cstrFieldEmail & "] as ssUserEmail, [" & cstrFieldPassword & "] as ssUserPassword, [" & cstrFieldGreeting & "] as ssFieldGreeting from " & cstrTableName & " where [" & cstrFieldUserName & "] = '" & pstrUsername & "'"
		set	pobjRS = CreateObject("adodb.recordset")
		pobjRS.CursorLocation = 2 'adUseClient
		pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly

		If pobjRS.EOF Then
			EmailPassword = "<h3><b>We have no record of this user</b></h3>"
		Else
			pstrUserID = trim(pobjRS.Fields("ssUserID").Value & "")
			pstrEmail = trim(pobjRS.Fields("ssUserEmail").Value & "")
			pstrPassword = trim(pobjRS.Fields("ssUserPassword").Value & "")

			pstrTemp = Replace(cstrLoginURL,"<<Email>>",pstrEmail)
			pstrTemp = Replace(pstrTemp,"<<Password>>",pstrPassword)

			pstrBody = Replace(cstrEmailPasswordBody,"<<Password>>",pstrPassword)
			pstrBody = Replace(pstrBody,"<<Email>>",pstrEmail)
			pstrBody = Replace(pstrBody,"<<LoginURL>>",pstrTemp)

				'Prepare string for modified mail routine
				' delimited with |
				'0 - sCustEmail
				'1 - sPrimary	- leave blank to use default, set to - not to send
				'2 - sSecondary	- leave blank to use default, set to - not to send
				'3 - sSubject
				'4 - sMessage

				Call createMail("",pstrEmail & "|" & "" & "|" & "-" & "|" & cstrEmailPasswordSubject & "|" & pstrBody)
			
			EmailPassword = "<h3><b>Your password has been sent to " & pstrEmail & "!</b></h3>"
		End If
		pobjRS.Close
		Set pobjRS = Nothing

	End Function	'EmailPassword

	'********************************************************************************
	
	Public Function ValidUserName(byVal strUsername, byVal strPassword)

	Dim pblnValidLogin
	Dim pclsCustomer
	Dim pstrOut

	'On Error Resume Next

		pblnValidLogin = False
		pstrUsername = strUsername
		pstrPassword = strPassword
		
		If len(pstrUsername) = 0 then
			If len(pstrPassword) = 0 then
				pstrOut = "<h3>Please log in.</h3>"
			Else
				pstrOut = "<h3><b>Please enter an Email address</b></h3>"
			End If
		ElseIf len(pstrPassword) = 0 then
			pstrOut = "<h3><b>Please enter a password</b></h3>"
		ElseIf Not checkInput(pstrUsername, pstrPassword) Then
			pstrOut = "<h3><b>There was a problem with your login. Please contact the system administrator.</b></h3>"
		Else
			Set pclsCustomer = New clsCustomer
			With pclsCustomer
				Set .Connection = cnn
				If .LoadCustomerByEmail(pstrUsername) Then
					If .custPasswd = pstrPassword Then
						pblnValidLogin = True
					ElseIf .LoadCustomerByEmailPassword(pstrUsername, pstrPassword) Then
					'this check added since email address is not required to be unique
					'this check could replace LoadCustomerByEmail initially but then the ability to separate invalid user/invalid password would be lost
						pblnValidLogin = True
					Else
						pstrOut = "<h3><b>You entered an invalid password. Please try again.</b></h3>"
					End If
				Else
					pstrOut = "<h3><b>You entered an invalid Email. Please try again.</b></h3>"
				End If
				
				If pblnValidLogin Then
					'pstrOut = "True"

					pstrUserID = .custID
					pstrPassword = .custPasswd
					pstrEmail = .custEmail
					pstrGreeting = .Greeting

					Session("custPricingLevel") = .PricingLevelID
					If Len(.clubExpDate) = 0 Then
						Call SetPromotionCodeToSession(.clubCode)
					ElseIf .clubExpDate >= Date() Then
						Call SetPromotionCodeToSession(.clubCode)
					End If		

					Call SetLoginParameters
					
					If Len(pstrRedirectPage) > 0 Then
						Call cleanup_dbconnopen	'This line needs to be included to close database connection
						Response.Redirect pstrRedirectPage
					End If
				End If
				
			End With	'pclsCustomer
			Set pclsCustomer = Nothing
		End If
		
		ValidUserName = pstrOut
		
	End Function	'ValidUserName

	'********************************************************************************
	
	Function ChangePassword(strUsername, strPassword, strNewPassword1, strNewPassword2)

	dim pstrNewPassword1, pstrNewPassword2
	dim pstrSQL, pobjRS

	'On Error Resume Next

		pstrUsername = strUsername
		pstrPassword = strPassword
		pstrNewPassword1 = strNewPassword1
		pstrNewPassword2 = strNewPassword2

		If len(pstrUsername) = 0 then
			ChangePassword = "<h3><b>Please enter a Email</b></h3>"
			Exit Function
		Elseif len(pstrPassword) = 0 then
			ChangePassword = "<h3><b>Please enter your current password</b></h3>"
			Exit Function
		Elseif len(pstrNewPassword1) = 0 then
			ChangePassword = "<h3><b>Please enter the new password</b></h3>"
			Exit Function
		Elseif len(pstrNewPassword2) = 0 then
			ChangePassword = "<h3><b>Please re-type your password</b></h3>"
			Exit Function
		Elseif (pstrNewPassword2 <> pstrNewPassword1) then
			ChangePassword = "<h3><b>The passwords you entered do not match.</b></h3>"
			Exit Function
		ElseIf Not checkInput(pstrUsername, pstrNewPassword1) Then
			ChangePassword = "<h3><b>There was a problem with your login. Please contact the system administrator.</b></h3>"
			Exit Function
		ElseIf Not checkInput(pstrUsername, pstrNewPassword2) Then
			ChangePassword = "<h3><b>There was a problem with your login. Please contact the system administrator.</b></h3>"
			Exit Function
		End If

		pstrSQL = "Select [" & cstrFieldUserID & "] as ssUserID, [" & cstrFieldEmail & "] as ssUserEmail, [" & cstrFieldPassword & "] as ssUserPassword, [" & cstrFieldGreeting & "] as ssFieldGreeting from " & cstrTableName & " where [" & cstrFieldUserName & "] = '" & pstrUsername & "'"
		set	pobjRS = CreateObject("adodb.recordset")
		pobjRS.CursorLocation = 2 'adUseClient
		pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		
		If Err.number <> 0 Then
			If instr(1,Err.Description,"cannot find the input table or query '" & cstrTableName & "'") <> 0 Then
				ChangePassword = "<div class='FatalError'>You need to upgrade your database to use integrated security</div>"
			Else
				ChangePassword = "<div class='FatalError'>Error: " & Err.number & " - " & Err.Description & "</div>"
			End If
		Else
			If pobjRS.eof or pobjRS.bof then
				ChangePassword = "<h3><b>You entered an invalid Email. Please try again.</b></h3>"
			Else
				Do While Not pobjRS.EOF
					If trim(pobjRS.Fields("ssUserPassword").Value) = pstrPassword Then
						pstrSQL = "Update " & cstrTableName & " set " & cstrFieldPassword & " = '" & pstrNewPassword1 & "' where " & cstrFieldUserName & " = '" & pstrUsername & "'"
						cnn.Execute pstrSQL,,128

						pstrUserID = Trim(pobjRS.Fields("ssUserID").Value & "")
						pstrPassword = Trim(pobjRS.Fields("ssUserPassword").Value & "")
						pstrEmail = Trim(pobjRS.Fields("ssUserEmail").Value & "")

						Call SetLoginParameters

						ChangePassword = "Username/Password Successfully Changed"
						pobjRS.Close
						set pobjRS = nothing
						Exit Function

						'ChangePassword = "<b>This functionality has been disabled.</b> The Username/Password would have been changed"
					End If
					pobjRS.MoveNext
				Loop

				ChangePassword = "<h3><b>You entered an invalid password. Please try again.</b></h3>"
			End If
		End If
		
		pobjRS.Close
		set pobjRS = nothing

	End Function 'ChangePassword

	'********************************************************************************
	
	Private Sub SetLoginParameters
	
		Call SetSessionLoginParameters(pstrUserID, pstrEmail)

		Call setGreeting(pstrGreeting)
		
		If Request.Form("rememberMe") = "1" Then
			Call setCookie_Email(pstrEmail, DateAdd("m", 1, Now()))
		Else
			Call setCookie_Email("", Now())
		End If
	
	End Sub	'SetLoginParameters

	'********************************************************************************
	
	Private Function checkInput(byRef strEmail, byRef strPassword)
	'Purpose: protect username (email address) and password from SQL Injection attacks
	
		checkInput = fncEmailValid(strEmail) And validatepassword(strPassword)
	
	End Function	'checkInput

	'Generic function to validate strings using regular expressions
	Function fncStringValid(strInput, strPattern)

	Dim MyRegExp

		Set MyRegExp = New RegExp
		MyRegExp.Pattern = ">" & strPattern & ">"
		fncStringValid = MyRegExp.Test(">" & strInput & ">")
	
	End Function

	Function fncEmailValid(strInput)
		fncEmailValid = fncStringValid(strInput, "[\w\.-]+@[\w\.-]+(\.(\w)+)+")
	End Function

	Function validatepassword(strPassword)
	
	Dim good_password_chars
	Dim i
	Dim c
	
		good_password_chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._!@#$%^&*()" 
		
		validatepassword = true
		
		For i = 1 to len(strPassword)
			c = mid(strPassword, i, 1 )
			if (InStr(good_password_chars, c ) = 0) then
				validatepassword = false
				exit function
			end if
		next
		
	End Function	'validatepassword
	
	'********************************************************************************
	
End Class	'clsLogin

'******************************************************************************************************************************
'******************************************************************************************************************************

Sub logOffCustomer()
	Session.Contents.Remove("login")
	Session.Contents.Remove("custPricingLevel")
	Call setVisitorLoggedInCustomerID(0)
End Sub	'logOffCustomer

'******************************************************************************************************************************

Sub ShowLoginForm(strMessage)

'mstrPrevPage = Request.QueryString("PrevPage")
'If Len(mstrPrevPage) = 0 Then mstrPrevPage = Request.Form("PrevPage")
%>
<!--#include file="ssclsLogin_LoginForm.asp"--> 
<% End Sub 'ShowLoginForm %>

<% Sub ShowChangePasswordForm(strMessage) %>
<!--#include file="ssclsLogin_ChangePasswordForm.asp"--> 
<% End Sub 'ShowChangePasswordForm %>

<%

Dim mblnShowMenu
Dim mclsLogin
Dim mstrAction
Dim mstrEmail
Dim mstrLoginMessage
Dim mstrLoginPageName
Dim mstrPassword
Dim mstrPrevPage

mstrLoginPageName = "myAccount.asp"
mblnShowMenu = False

'******************************************************************************************************************************

Sub ProtectThisPage(byVal strPageToProtect)

'***************************
'*
'*  Testing Only
'session("login") = ""
'*
'*
'***************************

	mstrAction = LoadRequestValue("Action")
	'Response.Write "mstrAction = " & mstrAction & "<br />"

	If mstrAction = "LogOff" Then Call logOffCustomer

	If Not isLoggedIn And (mstrAction <> "ChangePwd") And (mstrAction <> "EmailPwd") then

		mstrPrevPage = LoadRequestValue("PrevPage")
		If len(mstrPrevPage) = 0 Then mstrPrevPage = strPageToProtect

		mstrEmail = LoadRequestValue("Email")
		mstrPassword = LoadRequestValue("Password")
		Set mclsLogin = New clsLogin
		mstrLoginMessage = mclsLogin.ValidUserName(mstrEmail, mstrPassword)
		If len(mstrEmail) = 0 Then mstrEmail = Request.Cookies("Email")
		If  mstrLoginMessage = "True" Or Len(mstrLoginMessage) = 0 Then	
			If len(mstrPrevPage) = 0 then
				'Call ShowMenu
				mblnShowMenu = True
			Else
				Response.Clear
				Call cleanup_dbconnopen	'This line needs to be included to close database connection
				Response.Redirect mstrPrevPage
			End if	
		Else
			Call ShowMyAccountBreadCrumbsTrail("Login", False)
			Call ShowLoginForm(mstrLoginMessage)
		End If
		Set mclsLogin = Nothing
				
	ElseIf mstrAction = "ChangePwd" then
		If len(Request.Form("Action")) <> 0 Then 
			Set mclsLogin = New clsLogin
			mstrLoginMessage = mclsLogin.ChangePassword(Request.Form("Login"),Request.Form("Password"),Request.Form("NewPassword1"),Request.Form("NewPassword2"))
			Set mclsLogin = Nothing
		End If
		Call ShowMyAccountBreadCrumbsTrail("Change Password", False)
		Call ShowChangePasswordForm(mstrLoginMessage)
	ElseIf mstrAction = "EmailPwd" then
		mstrEmail = LoadRequestValue("Email")
		Set mclsLogin = New clsLogin
		mstrLoginMessage = mclsLogin.EmailPassword(mstrEmail)
		Set mclsLogin = Nothing
		Call ShowLoginForm(mstrLoginMessage)
	Else
		'Call ShowMenu
		mblnShowMenu = True
	End If

End Sub	'ProtectThisPage

'******************************************************************************************************************************

%>