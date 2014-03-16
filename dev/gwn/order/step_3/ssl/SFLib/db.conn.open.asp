<%
'********************************************************************************
'*
'*   db.conn.open.asp - 
'*   Revision Date: November 26, 2004											*
'*   Version 1.01.001                                                           *
'*
'*   Release Notes:
'*
'*
'*   COPYRIGHT INFORMATION
'*
'*   This file's origins are search_results.asp APPVERSION: 50.4014.0.3
'*   from LaGarde, Incorporated. It has been heavily modified by Sandshot Software
'*   with permission from LaGarde, Incorporated who has NOT released any of the 
'*   original copyright protections. As such, this is a derived work covered by
'*   the copyright provisions of both companies.
'*
'*   LaGarde, Incorporated Copyright Statement                                                                           *
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
%>
<!--#include file="administratorMenus.asp"-->
<!--#include file="applicationManagement.asp"-->
<!--#include file="cookieSessionManagement.asp"-->
<!--#include file="ssclsDebug.asp"-->
<!--#include file="ssFieldValidation.asp"-->
<%
'**********************************************************
'	Developer notes
'**********************************************************

'This page should be included at the top of every page with a db call

'**********************************************************
'*	Page Level variables
'**********************************************************

'Datatype Enumerations
Const enDatatype_string = 0
Const enDatatype_number = 1
Const enDatatype_date = 2
Const enDatatype_boolean = 4

Dim cblnTrackHangingDBConnections
Dim cblnSF5AE
Dim cblnSQLDatabase
Dim cnn
Dim mstrValidationErrorMessage
Dim mstrCurrentPage
Dim mstrRootPath
Dim vDebug

Dim GlobalStartTime

cblnTrackHangingDBConnections = True

'**********************************************************
'*	Functions
'**********************************************************

'Sub abandonSession
'Function addValidationError(byVal strError)
'Sub CheckStoreConfigApplicationSettings
'Function CorrectEmptyValue(byVal vntValue, byVal vntValueIfEmpty)
'Sub closeObj(byRef objItem)
'Sub EmailErrorMessage(byRef objErr, byVal strLocation, byVal strExtra)
'Function GetRS(byVal strSQL)
'Function initializeDBConnection(byRef objCnn, byVal strConnection)
'Function LoadRequestValue(byVal strSource)
'Function makeSQLUpdate(byVal strFieldName, byVal vntFieldValue, byVal blnEmptyOK, byVal bytFieldType)
'Sub ShowStoreFrontVersion
'Function sqlSafe(byVal strSQL)
'Function wrapSQLValue(byVal vntFieldValue, byVal blnEmptyOK, byVal bytFieldType)
'Sub WriteFormVariables

'**********************************************************
'*	Begin Page Code
'**********************************************************

If CBool(Len(Session("ssDebug_vDebug")) > 0) Then
	vDebug = 1
Else
	vDebug = 0
End If

If ssDebug_General Then
	On Error Goto 0
Else
	On Error Resume Next
End If

'vDebug = 1

If Not initializeDBConnection(cnn, "") Then
	Server.Transfer Application("DBFailPage")
End If

Call LoadCustomStoreConfigurationSettings	'cannot load these until after database connection is established
Call CheckStoreConfigApplicationSettings	'this section added for shared SSL's where the application variables may not carry over

cblnSQLDatabase = CBool(Application("AppDatabase") <> "Access")
'cblnSQLDatabase = True				'Set this value to True for SQL Server databases, only need to set this manually for very early versions
'cblnSQLDatabase = False			'Set this value to False for Access databases, only need to set this manually for very early versions

cblnSF5AE = CBool(Application("AppName") = "StoreFrontAE")
'cblnSF5AE = True					'Set this value to True for AE Sites, only need to set this manually for very early versions
'cblnSF5AE = False					'Set this value to False for SE Sites, only need to set this manually for very early versions

'Call ShowStoreFrontVersion

If Request.QueryString("Action") = "ResetVisitor" Then
	Call logOffCustomer
	Call resetVisitorPreferences
End If

Call CheckAcceptBackOrder_dbconnopen

'**********************************************************
'*	Begin Function Definitions
'**********************************************************

Sub abandonSession
	
On Error Resume Next

	Session.Abandon 
	Call cleanup_dbconnopen

	Response.Redirect(adminDomainName)

End Sub	'abandonSession

'--------------------------------------------------------------------------------------------------

Function addValidationError(byVal strError)
	If Len(mstrValidationErrorMessage) = 0 Then
		mstrValidationErrorMessage = strError
	Else
		mstrValidationErrorMessage = mstrValidationErrorMessage & "|" & strError
	End If
End Function	'addValidationError
	
'--------------------------------------------------------------------------------------------------

Sub CheckAcceptBackOrder_dbconnopen
	
'On Error Resume Next

	Call CheckAcceptBackOrder
	If Err.number <> 0 Then Err.Clear

End Sub	'CheckAcceptBackOrder_dbconnopen

'--------------------------------------------------------------------------------------------------

'This sub exists because of a case where the initial values of the recordset were nulls until pre-read
Sub hack_preReadRecordset(byRef objRS)

Dim i
Dim pvnt

	For i = 1 To objRS.Fields.Count
		pvnt = objRS.Fields(i-1).Value
	Next

End Sub

'--------------------------------------------------------------------------------------------------

Function hasValidationError()
	hasValidationError = Len(mstrValidationErrorMessage) > 0
End Function	'hasValidationError

'--------------------------------------------------------------------------------------------------

Function hasInboundError()
	hasInboundError = Len(Session("validationError")) > 0
End Function	'hasInboundError

'--------------------------------------------------------------------------------------------------

Sub setInboundError()
	If hasInboundError Then
		addValidationError(Session("validationError"))
		Session.Contents.Remove("validationError")
	End If
End Sub	'setInboundError

'--------------------------------------------------------------------------------------------------

Function returnValidationErrorToSender(byVal strURL)
	Session("validationError") = mstrValidationErrorMessage
	Server.Transfer(strURL)
End Function	'hasValidationError

'--------------------------------------------------------------------------------------------------

Function displayValidationError

Dim paryErrors
Dim i
Dim pstrOut

	Call setInboundError
	If hasValidationError Then
		pstrOut = "<div align=left><table border=1 cellspacing=0 cellpadding=0 style='border-collapse: collapse'>" _
				& "<tr><th>Entry Error</th><tr><td class='Error'>There were some problems with the form you submitted. Please correct the following items:<ul>"
		paryErrors = Split(mstrValidationErrorMessage, "|")
		For i = 0 To UBound(paryErrors)
			pstrOut = pstrOut & "<li class='Error'>" & paryErrors(i) & "</li>" 
		Next 'i
		pstrOut = pstrOut & "</ul></td></tr></table></div>" 
	End If
	
	displayValidationError = pstrOut
	
End Function	'displayValidationError

'--------------------------------------------------------------------------------------------------

Sub cleanup_dbconnopen
	
	Call checkForUnhandledErrors

On Error Resume Next

	Call cleanup_ssclsDebug
	Call closeDBConnection(cnn)
	
End Sub	'cleanup_dbconnopen

'--------------------------------------------------------------------------------------------------

Sub CheckStoreConfigApplicationSettings

Dim pstrSQL
Dim pobjRS
Dim pblnSQLDatabase

	On Error Resume Next
	
	If Len(Application("AppDatabase")) = 0 Then
		pstrSQL = "Select invenId From sfInventory Where invenId = 0"
		Set	pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			If Err.number = 0 Then
				'-----------------------------------------------------
				'DO NOT REMOVE OR MODIFY THESE VARIABLES
				Application("AppName") ="StoreFrontAE" 
				Application("CartName") = "Wish List"
				Application("CartSaveButton") = "ADD TO WISH LIST"
				'-----------------------------------------------------
			Else
				'-----------------------------------------------------
				'DO NOT REMOVE OR MODIFY THESE VARIABLES
				Application("AppName") ="StoreFront" 
				Application("CartName") = "Saved Cart"
				Application("CartSaveButton") = "Save To Cart"
				'-----------------------------------------------------
				Err.Clear
			End If
		End With
		Call closeObj(pobjRS)
		
		pstrSQL = "SELECT prodID FROM sfProducts WHERE (prodDateAdded <= GETDATE()) AND (prodID is Null)"
		Set	pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 2 'adUseClient
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			If Err.number = 0 Then
				Application("AppDatabase") = "SQL Server"
			Else
				Application("AppDatabase") = "Access"
				Err.Clear
			End If
		End With
		Call closeObj(pobjRS)

	End If
	
End Sub	'CheckStoreConfigApplicationSettings

'--------------------------------------------------------------------------------------------------

Function CorrectEmptyValue(byVal vntValue, byVal vntValueIfEmpty)

	If Len(Trim(CStr(vntValue & ""))) = 0 Then
		CorrectEmptyValue = vntValueIfEmpty
	Else
		CorrectEmptyValue = vntValue
	End If

End Function	'CorrectEmptyValue

'--------------------------------------------------------------------------------------------------

Function checkFieldLength(byVal strValue, byVal lngMaxLength, byVal bytTrim)
'bytTrim can have the following values
'0 - trim left, empty ok
'1 - trim right, empty ok
'2 - trim left, replace zero length string with null
'3 - trim right, replace zero length string with null
Dim pstrOut

	pstrOut = strValue
	If lngMaxLength > 0 Then
		If Len(pstrOut) > lngMaxLength Then
			Select Case bytTrim
				Case 0, 2: 'left portion
					pstrOut = Left(pstrOut, lngMaxLength)
				Case 1, 3: 'right portion
					pstrOut = Right(pstrOut, lngMaxLength)
				Case Else:
				
			End Select
		End If	'Len(pstrOut) > lngMaxLength
	End If	'lngMaxLength > 0
	
	'now check for null replacements
	If Len(pstrOut) = 0 Then
		If bytTrim = 2 Or bytTrim = 3 Then pstrOut = Null
	End If
	
	checkFieldLength = pstrOut
	
End Function	'checkFieldLength

'--------------------------------------------------------------------------------------------------

Sub addParameter(byRef objCommand, byVal strParameterName, byVal lngParameterType, byVal strValue, byVal lngMaxLength, byVal bytTrim)
	objCommand.Parameters.Append objCommand.CreateParameter(strParameterName, lngParameterType, adParamInput, parameterFieldLength(strValue, lngMaxLength), checkFieldLength(strValue, lngMaxLength, bytTrim))
End Sub	'addParameter

'--------------------------------------------------------------------------------------------------

Function parameterFieldLength(byVal strValue, byVal lngMaxLength)

Dim plngFieldLength
Dim plngLength

	plngLength = Len(strValue)
	If plngLength = 0 Then
		plngFieldLength = lngMaxLength
	ElseIf plngLength < lngMaxLength Then
		plngFieldLength = plngLength
	Else
		plngFieldLength = lngMaxLength
	End If
	
	parameterFieldLength = plngFieldLength
	
End Function	'parameterFieldLength

'--------------------------------------------------------------------------------------------------

Sub closeObj(byRef objItem)

On Error Resume Next

	objItem.Close
	Set objItem = Nothing	
	If Err.number <> 0 Then Err.Clear

End Sub	'closeObj

'--------------------------------------------------------------------------------------------------

Function ConvertToBoolean(byVal vntValue, byVal blnDefault)

On Error Resume Next

	vntValue = cBool(vntValue)
	If Err.number <> 0 Then vntValue = blnDefault
	ConvertToBoolean = vntValue

End Function	'ConvertToBoolean

'--------------------------------------------------------------------------------------------------

Function CurrentPage
	If Len(mstrCurrentPage) = 0 Then
		mstrCurrentPage = LCase(Request.ServerVariables("SCRIPT_NAME"))
		mstrCurrentPage = Replace(mstrCurrentPage,cstrSubWebPath,"",1,1)
		mstrCurrentPage = Replace(mstrCurrentPage,"/","",1,1)
		'Response.Write "<h4>Current Page: " & mstrCurrentPage & "</h4>"
	End If
	CurrentPage = mstrCurrentPage
End Function	'CurrentPage

'--------------------------------------------------------------------------------------------------

Sub EmailErrorMessage(byRef objErr, byVal strLocation, byVal strExtra)

Dim pstrSubject
Dim pstrBody

'Format for email message
'
'subject: Error on {siteName} - 
'body	: The following error

	
End Sub	'EmailErrorMessage

'--------------------------------------------------------------------------------------------------

Function GetRS(byVal strSQL)

Dim rs

	If Err.number <> 0 Then Err.Clear
	
	set	rs = CreateObject("adodb.recordset")
	with rs
        .CursorLocation = 2 'adUseClient
        
        On Error Resume Next
		.Open strSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		If Err.number <> 0 And CBool(vDebug = 1) Then
			Response.Write "<font color=red>Error in GetRS: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
			Response.Write "<font color=red>Error in GetRS: sql = " & strSQL & "</font><br />" & vbcrlf
			Response.Flush
			Err.Clear
		End If
	end with
	set GetRS = rs

End Function		' GetRS

'--------------------------------------------------------------------------------------------------

Sub MakeCombo(byRef strSQL, byVal strText,byVal strValue,byVal strSelected)

Dim i
Dim pobjRS
Dim pblnRecordsetCreated
Dim pblnFound

	pblnFound = False
	pblnRecordsetCreated = False
'	set	pobjRS = CreateObject("adodb.recordset")
	If isObject(strSQL) Then
		set pobjRS = strSQL
		If pobjRS.recordcount > 0 then pobjRS.movefirst
	Else
		set pobjRS = GetRS(strSQL)
		pblnRecordsetCreated = True
	End If
	
	for i = 1 to pobjRS.recordcount
		if len(strSelected) > 0 then
			if trim(pobjrs.Fields(strValue)) <> trim(strSelected) then
				Response.Write "<option value=" & chr(34) & pobjRS.Fields(strValue).Value & chr(34) & ">" & pobjRS.Fields(strText).Value & "</option>" & vbcrlf
			else
				Response.Write "<option selected value=" & chr(34) & pobjRS.Fields(strValue).Value & chr(34) & ">" & pobjRS.Fields(strText).Value & "</option>" & vbcrlf
				pblnFound = True
			end if
		else
			Response.Write "<option value=" & chr(34) & pobjRS.Fields(strValue).Value & chr(34) & ">" & pobjRS.Fields(strText).Value & "</option>" & vbcrlf
		end if
		pobjRS.movenext
	next
	
	If Not pblnFound And Len(Trim(strSelected & "")) > 0 Then
		Response.Write "<option selected value=" & chr(34) & strSelected & chr(34) & ">" & strSelected & "</option>" & vbcrlf
	End If
	
	'Clean up only if the recordset was
	If pblnRecordsetCreated Then
		pobjRS.Close
		set pobjRS = nothing
	End If
	
End Sub		' MakeCombo

'--------------------------------------------------------------------------------------------------

Function isChecked(byVal vntValue)
	If ConvertToBoolean(vntValue, False) Then isChecked = "checked"
End Function	'isChecked

'--------------------------------------------------------------------------------------------------

Function isSelected(byVal vntValue)
	If ConvertToBoolean(vntValue, False) Then isSelected = "selected"
End Function	'isChecked

'--------------------------------------------------------------------------------------------------

Function initializeDBConnection(byRef objCnn, byVal strConnection)

Dim pstrConnection
Dim pblnResult

	GlobalStartTime = Timer
	Call DebugRecordSplitTime("initializeDBConnection . . .(ssl/SFLib/db.conn.open.asp?initializeDBConnection)")
	pblnResult = False
	pstrConnection = strConnection
	If Len(pstrConnection) = 0 Then pstrConnection = Application("DSN_NAME")

	If Err.number <> 0 Then Err.Clear
	
	Set objCnn = CreateObject("ADODB.Connection")
	On Error Resume Next
	objCnn.Open pstrConnection
	If objCnn.State = 1 Then
		pblnResult = True
		Call synchronizeCookies
		
		On Error Goto 0
		Call recordPageView
		
		If cblnTrackHangingDBConnections Then
			Dim plngNumConnections
			Application.Lock
			plngNumConnections = Application("ActiveDBConnections")
			If Len(Session("CurrentPage")) > 0 Then Call RecordHangingActiveDBConnections(Session("CurrentPage"))
			Session("CurrentPage") = CurrentPage
			If Len(plngNumConnections) = 0 Then
				plngNumConnections = 1
			Else
				plngNumConnections = plngNumConnections + 1
			End If
			Application("ActiveDBConnections") = plngNumConnections
			Application.UnLock
			
			'Reset mstrCurrentPage since subweb support is pulled from database
			mstrCurrentPage = ""
		End If
	Else
		pblnResult = False
	End If
	
	initializeDBConnection = pblnResult
	
	Call DebugRecordSplitTime("initializeDBConnection complete")

End Function		' initializeDBConnection

'--------------------------------------------------------------------------------------------------

Sub closeDBConnection(byRef objCnn)

	objCnn.Close
	Set objCnn = Nothing

	If cblnTrackHangingDBConnections Then
		Dim plngNumConnections
		Application.Lock
		plngNumConnections = Application("ActiveDBConnections")
		If Len(plngNumConnections) = 0 Then
			plngNumConnections = 0
		Else
			plngNumConnections = plngNumConnections - 1
		End If
		Application("ActiveDBConnections") = plngNumConnections
		Application.UnLock
		Session.Contents.Remove("CurrentPage")
		'If plngNumConnections <> 0 Then Response.Write "<h3><font color=red>Possible hanging connection(s): " & plngNumConnections & "</font></h3>"
	End If
	
	If Len(Session("ssDebug_Timing")) > 0 Then Response.Write "<h4>This page took " & FormatNumber(Timer - GlobalStartTime, 4) & " seconds to process.</h4>"

End Sub		'closeDBConnection

'--------------------------------------------------------------------------------------------------

Function jsOutputValue(byVal strValue)
	jsOutputValue = chr(34) & Replace(strValue & "", Chr(34), Chr(34) & " + cstrQuote + " & Chr(34)) & chr(34)
End Function

'--------------------------------------------------------------------------------------------------

Function LoadRequestValue(byVal strSource)

dim p_strTemp

	p_strTemp = Request.QueryString(strSource)
	If len(p_strTemp) = 0 Then p_strTemp = Request.Form(strSource)
	LoadRequestValue = p_strTemp
	
End Function	'LoadRequestValue

'--------------------------------------------------------------------------------------------------

Function makeSQLUpdate(byVal strFieldName, byVal vntFieldValue, byVal blnEmptyOK, byVal bytFieldType)

Dim pstrTempSQL

	pstrTempSQL = strFieldName & "=" & wrapSQLValue(vntFieldValue, blnEmptyOK, bytFieldType)
	makeSQLUpdate = pstrTempSQL

End Function	'makeSQLUpdate

'--------------------------------------------------------------------------------------------------

Sub ShowStoreFrontVersion

	Response.Write "<HR>"
	Response.Write "<p>The page you are viewing experienced a problem reading from the database. This can be due to a number of issues ranging from incorrect application settings to failure to perform a required database upgrade. The below information is provided to help to determine if it is your application settings.</p>"
	Response.Write "<H4>StoreFront Version: " & Application("AppName") & " (this should be StoreFront or StoreFrontAE)</H4>"
	If cblnSF5AE Then
		Response.Write "-- this script detects you have StoreFront 5.0 AE<br />"
	Else
		Response.Write "-- this script detects you have StoreFront 5.0 SE<br />"
	End If
	Response.Write "<H4>Database Type: " & Application("AppDatabase") & " (this should be Access or SQL)</H4>"
	If cblnSQLDatabase Then
		Response.Write "-- this script detects you are using a SQL Server database<br />"
	Else
		Response.Write "-- this script detects you are using an Access database<br />"
	End If
	Response.Write "<p>If any of the above settings are incorrect you should contact StoreFront Support or your developer. You can manually set these values in ssLibrary/modDatabase.asp.</p>"
	Response.Write "<HR>"
	Response.Flush
	
End Sub	'ShowStoreFrontVersion

'--------------------------------------------------------------------------------------------------

Function sqlSafe(byVal strSQL)

	If isNull(strSQL) Then
		sqlSafe = strSQL
	Else
		sqlSafe = Replace(strSQL, "'", "''")
	End If

End Function	'sqlSafe

'--------------------------------------------------------------------------------------------------

Function wrapSQLValue(byVal vntFieldValue, byVal blnEmptyOK, byVal bytFieldType)

Dim pstrTempSQL
Dim pvntTempValue

	Select Case CDbl(bytFieldType)
		Case enDatatype_string:	'string
			If blnEmptyOK Then
				pstrTempSQL = "'" & sqlSafe(vntFieldValue) & "'"
			Else
				If Len(Trim(vntFieldValue) & "") = 0 Then
					pstrTempSQL = "Null"
				Else
					pstrTempSQL = "'" & sqlSafe(vntFieldValue) & "'"
				End If
			End If
		Case enDatatype_number: 'number
			If Len(Trim(vntFieldValue) & "") = 0 Then
				pstrTempSQL = "Null"
			Else
				pstrTempSQL = sqlSafe(vntFieldValue)
			End If
		Case enDatatype_date: 'date
			If Len(Trim(vntFieldValue) & "") = 0 Then
				pstrTempSQL = "Null"
			Else
				If cblnSQLDatabase Then
					pstrTempSQL = "'" & sqlSafe(vntFieldValue) & "'"
				Else
					pstrTempSQL = "#" & sqlSafe(vntFieldValue) & "#"
				End If
			End If
		Case enDatatype_boolean: 'boolean
			If Len(Trim(vntFieldValue) & "") = 0 Then
				pstrTempSQL = "Null"
			Else
				pvntTempValue = Trim(LCase(CStr(vntFieldValue)))
				If (pvntTempValue = "yes") Or (pvntTempValue = "1") Or (pvntTempValue = "true") Then
					pvntTempValue = 1
				Else
					pvntTempValue = 0
				End If
				pstrTempSQL = sqlSafe(pvntTempValue)
			End If
	End Select
	
	wrapSQLValue = pstrTempSQL

End Function	'wrapSQLValue

'--------------------------------------------------------------------------------------------------

Function isXSSRisk(byVal strToCheck)

Dim pobjRegEx
Dim sBad

	sBad = "(<\s*(script|object|applet|embed|form)\s*>)"	' <  script xxx >
	sbad = sbad & "|" & "(<.*>)"							' >xxxxx<  warning includes hyperlinks and stuff between > and <
	sbad = sbad & "|" & "(&.{1,5};)"						' &xxxx;
	sbad = sbad & "|" & "eval\s*\("							' eval  (
 	sbad = sbad & "|" & "(event\s*=)"						' event  =

	'Now lets check for encoding
	sbad = Replace(sbad,"<", "(<|%60|<)")
	sbad = Replace(sbad,">", "(>|%62|>)")

	Set pobjRegEx = CreateObject("VBScript.RegExp") ' -> VB Script 5.0
	With pobjRegEx
		.IgnoreCase = True	'ignore case of string
		.Global = False		'stop on first hit
		.Pattern = sBad

		isXSSRisk = o.Test(strToCheck)
	End With
	Set pobjRegEx = Nothing
	
End Function	'isXSSRisk

'--------------------------------------------------------------------------------------------------

Sub WriteFormVariables

Dim pstrFormItem

	Response.Write "<fieldset><legend>Form Contents</legend>" & vbcrlf
	For Each pstrFormItem In Request.Form
		Response.Write pstrFormItem & ": " & Request.Form(pstrFormItem) & "<br />" & vbcrlf
	Next 'pstrFormItem
	Response.Write "</fieldset>" & vbcrlf
	
	Response.Write "<fieldset><legend>QueryString Contents</legend>" & vbcrlf
	For Each pstrFormItem In Request.QueryString
		Response.Write pstrFormItem & ": " & Request.QueryString(pstrFormItem) & "<br />" & vbcrlf
	Next 'pstrFormItem
	Response.Write "</fieldset>" & vbcrlf
	
End Sub	'WriteFormVariables

'--------------------------------------------------------------------------------------------------

Function getIdentity

Dim plngID
Dim pobjRS

	Set pobjRS = GetRS("SELECT @@Identity")
	If pobjRS.EOF Then
		plngID = -1
	Else
		plngID = pobjRS.Fields(0).Value
	End If
	pobjRS.Close
	Set pobjRS = Nothing
	
	getIdentity = plngID

End Function	'getIdentity

'--------------------------------------------------------------------------------------------------

Sub testScript()

Dim pobjCmd
Dim pobjRS
Dim pstrSQL

	pstrSQL = "SET NOCOUNT ON;INSERT INTO sfOrders(orderDate) VALUES(GETDATE());SELECT SCOPE_IDENTITY()"
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		Set .ActiveConnection = cnn
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		
		Set pobjRS = .Execute
		If Not pobjRS.EOF Then
			Response.Write "<h1>Result: " & pobjRS.Fields(0).Value & "</h1>"
		Else
			Response.Write "<h1>Fail</h1>"
		End If
		pobjRS.Close
		Set pobjRS = Nothing
	End With
	Set pobjCmd = Nothing

End Sub	'testScript

'--------------------------------------------------------------------------------------------------

Sub WriteCommandParameters(byRef objCMD, byRef strID, byVal strTitle)

Dim i

	With objCMD
		Response.Write "<fieldset><legend>" & strTitle & "</legend>"
		Response.Write "<table border=1 cellspacing=0>"
		Response.Write "<tr><td colspan=3>Commandtext: " & .Commandtext & "</td></tr>"
		Response.Write "<tr><th></th><th>Field</th><th>Value</th></tr>"
		For i = 0 To .Parameters.Count - 1
			Response.Write "<tr><td>" & i & "</td><td>" & .Parameters.item(i).Name & "</td><td>" & .Parameters.item(i).Value & "</td></tr>"
		Next 'i
		Response.Write "<tr><th></th><th>ID</th><th>" & strID & "</th></tr>"
		Response.Write "</table></fieldset>"
		'Response.Flush
	End With
	
End Sub	'WriteCommandParameters

'--------------------------------------------------------------------------------------------------

Sub checkForUnhandledErrors()

Dim pstrMailBody
Dim pstrSubject
Dim pstrOut

	If Err.number <> 0 And cblnTrackHangingDBConnections Then Call RecordHangingUnhandledError

	'If Request.ServerVariables("LOCAL_ADDR") = "127.0.0.1" Then Exit Sub
	
	If ssDebug_General Then
		If Err.number = 0 Then
			Response.Write "<h4>No Unhandled Errors</h4>"
		Else
			pstrSubject = "Unhandled Error on " & adminDomainName & CurrentPage & ": Error suppressed"
			
			pstrMailBody = "There was an unhandled error on " & adminDomainName & "." & vbcrlf _
							& "Time: " & FormatDateTime(Now, 3) & vbcrlf _
							& vbcrlf _
							& "Technical Information (for support personnel)" & vbcrlf _
							& vbcrlf _
							& "Error " & CStr(Err.Number) & ": " & Err.Description & vbcrlf _
							& "File: " & CurrentPage & vbcrlf _
							& vbcrlf _
							& "Server Information:" & vbcrlf _
							& "SERVER_NAME: " & Request.ServerVariables("SERVER_NAME") & vbcrlf _
							& "LOCAL_ADDR: " & Request.ServerVariables("LOCAL_ADDR") & vbcrlf _
							& "REMOTE_ADDR: " & Request.ServerVariables("REMOTE_ADDR") & vbcrlf _
							& vbcrlf _
							& "Browser Information:" & vbcrlf _
							& "HTTP_USER_AGENT: " & Request.ServerVariables("HTTP_USER_AGENT") & vbcrlf _
							& vbcrlf _
							& "Page:" & vbcrlf _
							& Request.ServerVariables("REQUEST_METHOD") & " " & Request.TotalBytes & " bytes to " & Request.ServerVariables("SCRIPT_NAME") & vbcrlf _
							& vbcrlf _
							& "Request.Form: " & Request.Form & vbcrlf _
							& vbcrlf _
							& "Request.QueryString: " & Request.QueryString & vbcrlf
			Response.Write "<fieldset><legend>" & pstrSubject & "</legend>"
			Response.Write Replace(pstrMailBody, vbcrlf, "<br />"	)
			Response.Write "</fieldset>"
		End If
	Else
		If Err.number = 0 Then Exit Sub
		pstrSubject = "Unhandled Error on " & adminDomainName & CurrentPage & ": Error suppressed"
		
		pstrMailBody = "There was an unhandled error on " & adminDomainName & "." & vbcrlf _
						& "Time: " & FormatDateTime(Now, 3) & vbcrlf _
						& vbcrlf _
						& "Technical Information (for support personnel)" & vbcrlf _
						& vbcrlf _
						& "Error " & CStr(Err.Number) & ": " & Err.Description & vbcrlf _
						& "File: " & CurrentPage & vbcrlf _
						& vbcrlf _
						& "Server Information:" & vbcrlf _
						& "SERVER_NAME: " & Request.ServerVariables("SERVER_NAME") & vbcrlf _
						& "LOCAL_ADDR: " & Request.ServerVariables("LOCAL_ADDR") & vbcrlf _
						& "REMOTE_ADDR: " & Request.ServerVariables("REMOTE_ADDR") & vbcrlf _
						& vbcrlf _
						& "Browser Information:" & vbcrlf _
						& "HTTP_USER_AGENT: " & Request.ServerVariables("HTTP_USER_AGENT") & vbcrlf _
						& vbcrlf _
						& "Page:" & vbcrlf _
						& Request.ServerVariables("REQUEST_METHOD") & " " & Request.TotalBytes & " bytes to " & Request.ServerVariables("SCRIPT_NAME") & vbcrlf _
						& vbcrlf _
						& "Request.Form: " & Request.Form & vbcrlf _
						& vbcrlf _
						& "Request.QueryString: " & Request.QueryString & vbcrlf

		Call createMail("", cstrPrimaryEmailToSendErrorTo & "|" & "" & "|" & "" & "|" & pstrSubject & "|" & pstrMailBody)
	End If

End Sub	'checkForUnhandledErrors

'--------------------------------------------------------------------------------------------------

Function getHiddenElement(byVal strName, byVal strValue)
	getHiddenElement = "<input type=""hidden"" name=" & Chr(34) & strName & Chr(34) & " id=" & Chr(34) & strName & Chr(34) & " value=" & Chr(34) & Server.HTMLEncode(strValue) & Chr(34) & ">"
End Function

'--------------------------------------------------------------------------------------------------

Function RootPath
	If Len(mstrRootPath) = 0 Then
		mstrRootPath = Request.ServerVariables("APPL_PHYSICAL_PATH")
	End If
	RootPath = mstrRootPath
End Function	'RootPath

'--------------------------------------------------------------------------------------------------

Sub writeQuerystringParametersToHiddenValues(byVal aryItemsToSkip)

Dim i
Dim pblnMatch
Dim pstrItemName

	For Each pstrItemName In Request.QueryString
		pblnMatch = True
		If isArray(aryItemsToSkip) Then
			For i = 0 To UBound(aryItemsToSkip)
				If pstrItemName = aryItemsToSkip(i) Then
					pblnMatch = False
					Exit For
				End If
			Next 'i
		ElseIf pstrItemName = aryItemsToSkip Then
			pblnMatch = False
		End If
		
		If pblnMatch Then Response.Write getHiddenElement(pstrItemName, Request.QueryString(pstrItemName))

	Next 'pstrItemName

End Sub	'writeQuerystringParametersToHiddenValues

'--------------------------------------------------------------------------------------------------
' FastString Code based on article located at http://www.eggheadcafe.com/articles/20011227.asp
' Usage:
'	Set stringBuilder = New FastString
'	stringBuilder.Reset
'	stringBuilder.Append ""
'	Result = stringBuilder.concat

Class FastString

Private stringArray
Dim growthRate
Dim numItems

	Private Sub Class_Initialize()
		growthRate = 50
		numItems = 0
		ReDim stringArray(growthRate)
	End Sub

	Private Sub Class_Terminate()
		Erase stringArray
	End Sub

	Public Sub Append(ByVal strValue)
		If numItems > UBound(stringArray) Then ReDim Preserve stringArray(UBound(stringArray) + growthRate)
		stringArray(numItems) = strValue & "" '& "" prevents type mismatch error if strValue is null. Performance hit is negligible.
		numItems = numItems + 1
	End Sub

	Public Sub Reset
		Erase stringArray
		Class_Initialize
	End Sub

	Public Function concat() 
		Redim Preserve stringArray(numItems) 
		concat = Join(stringArray, "")
	End Function

End Class	'FastString

%>
