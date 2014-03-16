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

'You should record all known items in cache here
'CountryArray
'PaymentTypesArray
'sfManufacturers
'sfManufacturers_Expires
'sfVendors
'sfVendors_Expires
'ssCategorySearch
'ssCategorySearch_Expires
'StateArray

'**********************************************************
'*	Page Level variables
'**********************************************************

'**********************************************************
'*	Functions
'**********************************************************

'Function getFromCache(byVal strCacheName)
'Function isCacheItemExpired(byVal strCacheName)
'Sub removeFromCache(byVal strCacheName)
'Sub saveToCache(byVal strCacheName, byRef vntToSave, byVal dtExpires)


'**********************************************************
'*	Begin Page Code
'**********************************************************

'**********************************************************

Function getFromCache(byVal strCacheName)
	If Not isCacheItemExpired(strCacheName) Then getFromCache = Application(strCacheName)
End Function	'getFromCache

'**********************************************************

Function isCacheItemExpired(byVal strCacheName)

Dim pdtExpires
Dim pblnExpired

	pdtExpires = Application(strCacheName & "_Expires")
	If Len(CStr(pdtExpires)) = 0 Then
		pblnExpired = True
	ElseIf Not isDate(pdtExpires) Then
		Call removeFromCache(strCacheName)
		pblnExpired = True
	ElseIf CBool(pdtExpires < Now()) Then
		Call removeFromCache(strCacheName)
		pblnExpired = True
	Else
		pblnExpired = False
	End If
	
	'Response.Write "isCacheItemExpired (" & strCacheName & "_Expires): " & pdtExpires & ", Expired = " & pblnExpired & "<br />"
	isCacheItemExpired = pblnExpired

End Function	'isCacheItemExpired

'**********************************************************

Sub removeFromCache(byVal strCacheName)
	Application.Lock
	Application.Contents.Remove(strCacheName)
	Application.Contents.Remove(strCacheName & "_Expires")
	Application.UnLock
End Sub	'removeFromCache

'**********************************************************

Sub saveToCache(byVal strCacheName, byRef vntToSave, byVal dtExpires)
	Application.Lock
	Application(strCacheName) = vntToSave
	Application(strCacheName & "_Expires") = dtExpires
	Application.UnLock
End Sub	'saveToCache

'**********************************************************

%>