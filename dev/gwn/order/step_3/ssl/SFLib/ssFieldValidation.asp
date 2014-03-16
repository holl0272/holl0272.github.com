<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Revision Date: April 21, 2004												*
'*   Version 1.00.002                                                           *
'*                                                                              *
'*   1.00.002 (April 21, 2004)                                                  *
'*   - added check if Application settings are incorrect                        *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Const cstrPermissibleCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
	
'--------------------------------------------------------------------------------------------------

Function allPermissibleCharacters(byVal strToCheck, byVal strPermissibleCharacters)

Dim i
Dim char

	'Now protect against basic script attacks
	'Only allow characters and numbers as specified above
	For i = 1 to Len(strToCheck)
		char = mid(strToCheck, i, 1)
		If InStr(strPermissibleCharacters, char) = 0  Then
			If True Then Response.Write "Invalid Character in <b>" & strToCheck & "</b>. The character (<em>" & char & "</em>) is not in the approved list.<br />" & strPermissibleCharacters & "<br />If this site is live you should disable this message in ssl/SFLib/ssFieldValidation.asp<br />"
			allPermissibleCharacters = False
			Exit Function
		End If
	Next
	
	allPermissibleCharacters = True

End Function 'allPermissibleCharacters

'--------------------------------------------------------------------------------------------------

'Generic function to validate strings using regular expressions
Function fncStringValid(strInput, strPattern)

Dim MyRegExp

	Set MyRegExp = New RegExp
	MyRegExp.Pattern = ">" & strPattern & ">"
	fncStringValid = MyRegExp.Test(">" & strInput & ">")

End Function	'fncStringValid

'--------------------------------------------------------------------------------------------------

Function fncEmailValid(strInput)
	fncEmailValid = fncStringValid(strInput, "[\w\.-]+@[\w\.-]+(\.(\w)+)+")
End Function

'--------------------------------------------------------------------------------------------------

Function validatepassword(strPassword)

Dim good_password_chars
Dim i
Dim c

	good_password_chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789" 
	
	validatepassword = true
	
	For i = 1 to len(strPassword)
		c = mid(strPassword, i, 1 )
		if (InStr(good_password_chars, c ) = 0) then
			validatepassword = false
			exit function
		end if
	next
	
End Function	'validatepassword

'--------------------------------------------------------------------------------------------------

Function makeInputSafe(byVal str)
	makeInputSafe=replace(str & "", "'", "''")
End Function

'--------------------------------------------------------------------------------------------------

Function validNumber(byVal str)
	If len(str) > 0 And isNumeric(str) Then
		validNumber = True
	Else
		validNumber = False
	End If
End Function	'validNumber

'--------------------------------------------------------------------------------------------------

Function checkLength(byRef strValue, byVal maxLength, byVal emptyOK, byVal autoTrim)

Dim pblnResult
Dim pstrTemp

	pstrTemp = strValue
	If maxLength = 0 Then
		If emptyOK Then
			pblnResult = True
		Else
			pblnResult = CBool(Len(strValue) > 0)
		End If
	Else
		If Len(strValue) > maxLength And autoTrim Then pstrTemp = Left(strValue, maxLength)
		
		If emptyOK Then
			pblnResult = True
		Else
			pblnResult = CBool(Len(pstrTemp) > 0)
		End If
	End If
	
	strValue = pstrTemp
	checkLength = pblnResult
	
End Function	'checkLength

'--------------------------------------------------------------------------------------------------

%>