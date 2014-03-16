<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Revision Date: November 26, 2004											*
'*   Version 1.01.001                                                           *
'*                                                                              *
'*   1.00.001 (November 26, 2004)                                               *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/


'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'visitorID is the key parmater used by the shopping cart; this replaces the previous Session.SessionID usage which was saved to Session("SessionID")
'	and Cookies("sfOrder")("SessionID")
'	ReturningOrder	- Cookie For NewOrder Page (last orderID)

'login stages
'1) Cookies("sfCustomer")("custID")
'	- maintained a year, not considered trustworthy at all
'	- set in 
'2) Session("") - only set if customer actually logs in, trustworthy
'3) visitorCustomerID - only set if customer actually logs in, trustworthy; this is automatically cleared in global.asa with session_onEnd

'Session variables - only present in event of valid login
	'Session("custPricingLevel")			- set in ssclsLogin
	'session("login")						- set in ssclsLogin - login name(email address)
	'Session("custGreeting") = pstrGreeting	- set in ssclsLogin
	'session("AdminLogin")					- set in ssLibrary/modLogin - array of user permissions

'Cookies
	'Response.Cookies("sfCustomer")("custID")
	'response.Cookies("Email") = pstrEmail	- set in ssclsLogin

'identifying temporary cart
'custID_cookie
'-- if empty pull the visitorID

'Other cookies - can eventually remove since these are saved to the visitor record


'**********************************************************
'*	Global variables
'**********************************************************

Dim iCustID

'**********************************************************
'*	Functions
'**********************************************************

'Sub preventPageCache
'Function SessionID()
'Sub setCookie_sfAddProduct()
'Sub setCookie_sfSearch()


'**********************************************************
'*	Begin Page Code
'**********************************************************

Sub preventPageCache

On Error Resume Next

	Response.Expires = 60
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	
	If Err.number <> 0 Then Err.Clear
	
End Sub	'preventPageCache

'******************************************************************************************************************************

Function custID_cookie()

Dim pvntID

	pvntID = getCookie_custID
	If Len(pvntID) > 0 And Not isNumeric(pvntID) Then
		Call setCookie_custID("", Now())
	   	pvntID = ""
	End If
	
	custID_cookie = pvntID

End Function	'custID_cookie

'**********************************************************

Function getCookie_custID()
	getCookie_custID = Request.Cookies("sfCustomer")("custID")
End Function

'**********************************************************

Sub setCookie_custID(byVal vntCustID, byVal dtExpires)
	On Error Resume Next	'necessary since under certain debug conditions page content may have been already written
	Response.Cookies("sfCustomer")("custID") = vntCustID
	Response.Cookies("sfCustomer").Expires = dtExpires
	If Err.number <> 0 Then Err.Clear
End Sub

'**********************************************************

Sub expireCookie_sfCustomer()
	Response.Cookies("sfCustomer").Expires = Now()
End Sub

'**********************************************************

Function getCookie_visitorID()

Dim pvntID

	pvntID = Trim(Request.Cookies("visitorID"))
	If Len(pvntID) > 0 Then
		If Not loadVisitorPreferences(pvntID) Then pvntID = ""
	End If
	
	If Len(pvntID) = 0 Then
		pvntID = createVisitor(Session.SessionID, Request.QueryString("REFERER"), Request.ServerVariables("HTTP_REFERER"), Request.ServerVariables("REMOTE_ADDR"))
		Call addSessionDebugMessage("Created visitor in getCookie_SessionID, new ID is " & pvntID)
		Call setCookie_visitorID(pvntID, Date() + 365)
	End If
	
	getCookie_visitorID = pvntID
	
End Function	'getCookie_visitorID

'**********************************************************

Sub setCookie_visitorID(byVal vntID, byVal dtExpires)
	Response.Cookies("visitorID") = vntID
	Response.Cookies("visitorID").Expires = dtExpires
End Sub

'**********************************************************

Sub setGreeting(byVal strGreeting)
	Session("custGreeting") = strGreeting
End Sub

'**********************************************************

Function getGreeting()
	getGreeting = Session("custGreeting")
End Function

'**********************************************************

Sub setCookie_Email(byVal vntID, byVal dtExpires)
	Response.Cookies("Email") = vntID
	Response.Cookies("Email").Expires = dtExpires
End Sub

'**********************************************************

Function getCookie_Email()
	getCookie_Email = Request.Cookies("Email")
End Function

'**********************************************************

Function getCookie_SessionID()

Dim pvntID

	pvntID = Request.Cookies("sfOrder")("SessionID")
	If Len(pvntID) = 0 Then pvntID = getCookie_visitorID

	getCookie_SessionID	= pvntID
	
End Function

'**********************************************************

Sub setCookie_SessionID(byVal vntID, byVal dtExpires)
	Response.Cookies("sfOrder")("SessionID") = vntID
	Response.Cookies("sfOrder").Expires = dtExpires
End Sub

'**********************************************************

Function SessionID()
	If Len(Session("SessionID") & "") = 0 Then Session("SessionID") = Session.SessionID
	SessionID = Session("SessionID")
End Function

'**********************************************************

Sub setSessionID(byVal lngSessionID)
	If Len(lngSessionID) > 0 Then Session("SessionID") = lngSessionID
End Sub

'**********************************************************

Sub setCookie_sfAddProduct()
	If Len(Request.Cookies("sfAddProduct")("Path")) = 0 Then
		On Error Resume Next
		Response.Cookies("sfAddProduct")("Path") = Request.ServerVariables("HTTP_REFERER")
		Response.Cookies("sfAddProduct").Expires = Date() + 1
	End If
End Sub

'**********************************************************

Function getCookie_sfAddProduct()
	getCookie_sfAddProduct	= Trim(Request.Cookies("sfAddProduct")("Path"))
End Function

'**********************************************************

Sub setCookie_ReturningOrder(byVal lngOrderID)
	Response.Cookies("ReturningOrder") = iOrderID
	Response.Cookies("ReturningOrder").Expires = Date() + 31   
End Sub

'**********************************************************

Function getCookie_sfSearch()
	getCookie_sfSearch	= Trim(Request.Cookies("sfSearch")("SearchPath"))
End Function

'**********************************************************

Sub setCookie_sfSearch()
	On Error Resume Next
	Response.Cookies("sfSearch")("SearchPath") = Request.ServerVariables("HTTP_REFERER")
	Response.Cookies("sfSearch").Expires = Date() + 1
	Response.Cookies("sfSearchCheck") = "Test"
End Sub

'**********************************************************

Sub write_KnownCookies()
	Response.Write("<fieldset style=""background:white""><legend>Known Cookies</legend><table border=1 cellspacing=0><tr><th>Item</th><th>Value</th></tr>")
	Response.Write("<tr><td>custID_cookie</td><td>" & custID_cookie & "</td></tr>")
	Response.Write("<tr><td>SessionID</td><td>" & SessionID & "</td></tr>")
	Response.Write("<tr><td>getCookie_visitorID</td><td>" & getCookie_visitorID & "</td></tr>")
	Response.Write("<tr><td>getCookie_SessionID</td><td>" & getCookie_SessionID & "</td></tr>")
	Response.Write("<tr><td>getCookie_sfAddProduct</td><td>" & getCookie_sfAddProduct & "</td></tr>")
	Response.Write("<tr><td>getCookie_sfSearch</td><td>" & getCookie_sfSearch & "</td></tr>")
	Response.Write "</table></fieldset>"
End Sub

'**********************************************************

Sub synchronizeCookies()

Dim pvntID

	pvntID = getCookie_visitorID
	If pvntID <> getCookie_SessionID Then
		Call setCookie_SessionID(pvntID,  DateAdd("d", 1, Now()))
		Call setSessionID(pvntID)
	ElseIf pvntID <> SessionID Then
		Call setSessionID(pvntID)
	End If

End Sub	'synchronizeCookies

'**********************************************************

Function getFromSession(byVal strSessionName)
	If Not isSessionItemExpired(strSessionName) Then getFromSession = Application(strSessionName)
End Function	'getFromSession

'******************************************************************************************************************************

Function isLoggedIn()

	'isLoggedIn = CBool(Len(Session("login")) > 0)
	If Len(CStr(visitorLoggedInCustomerID)) = 0 Then
		isLoggedIn = False
	ElseIf CBool(visitorLoggedInCustomerID <> 0) Then
		isLoggedIn = True
	Else
		isLoggedIn = False
	End If
	
End Function	'isLoggedIn

'******************************************************************************************************************************

Sub SetSessionLoginParameters(byVal lngCustID, byVal strEmail)

	Call setCookie_custID(lngCustID, DateAdd("d", 730, Now()))
	Call setCookie_SessionID(SessionID, DateAdd("d", 1, Now()))

	Session("login") = strEmail
	
	Call setVisitorLoggedInCustomerID(lngCustID)
	Call setVisitorCustomerID(lngCustID)
	Call setVisitorShippingLocationByCustomerID(lngCustID)

End Sub	'SetSessionLoginParameters
	
'**********************************************************

Function isSessionItemExpired(byVal strSessionName)

Dim pdtExpires
Dim pblnExpired

	pdtExpires = Session(strSessionName & "_Expires")
	If Len(CStr(pdtExpires)) = 0 Then
		pblnExpired = True
	ElseIf Not isDate(pdtExpires) Then
		Call removeFromSession(strSessionName)
		pblnExpired = True
	ElseIf CBool(pdtExpires < Now()) Then
		Call removeFromSession(strSessionName)
		pblnExpired = True
	Else
		pblnExpired = False
	End If
	
	'Response.Write "isSessionItemExpired (" & strSessionName & "_Expires): " & pdtExpires & ", Expired = " & pblnExpired & "<br />"
	isSessionItemExpired = pblnExpired

End Function	'isSessionItemExpired

'**********************************************************

Sub removeFromSession(byVal strSessionName)
	Session.Contents.Remove(strSessionName)
	Session.Contents.Remove(strSessionName & "_Expires")
End Sub	'removeFromSession

'**********************************************************

Sub saveToSession(byVal strSessionName, byRef vntToSave, byVal dtExpires)
	Session(strSessionName) = vntToSave
	Session(strSessionName & "_Expires") = dtExpires
End Sub	'saveToSession

'******************************************************************************************************************************

Function validCustIDCookie

Dim pblnValidCustID

	pblnValidCustID = False
	iCustID = custID_cookie
	If Len(iCustID) > 0 Then
		If CheckCustomerExists(iCustID) Then
			Call setCookie_custID(iCustID, Date() + 730)
			pblnValidCustID = True
		Else
			Call setCookie_custID("", Now())
		End If
	End If
	
	validCustIDCookie = pblnValidCustID

End Function	'validCustIDCookie

'******************************************************************************************************************************

Function CheckCustomerExists(byVal lngCustID)

Dim pstrSQL
Dim pobjRS

	If Len(lngCustID) = 0 Or Not isNumeric(lngCustID) Then
		CheckCustomerExists = False
	Else
		pstrSQL = "SELECT custID FROM sfCustomers WHERE custID = " & makeInputSafe(lngCustID)
		set pobjRS = GetRS(pstrSQL)
		CheckCustomerExists = NOT pobjRS.EOF
		Call closeObj(pobjRS)
	End If

End Function	'CheckCustomerExists

'******************************************************************************************************************************

Function isAdminLoggedIn()
	isAdminLoggedIn = isArray(Session("AdminLogin"))
End Function

'******************************************************************************************************************************

Function isAdminAutoLoginCookieSet()
	isAdminAutoLoginCookieSet = CBool(Len(Request.Cookies("adminUserID")) > 0)
End Function

'**********************************************************

Function isIPAnAdminIP
	Select Case Request.ServerVariables("REMOTE_ADDR")
		Case "127.0.0."
			isIPAnAdminIP = True
		Case Else
			'Response.Write "IP: " & Request.ServerVariables("REMOTE_ADDR") & "<br />"
			isIPAnAdminIP = False
	End Select
End Function

'******************************************************************************************************************************
%>