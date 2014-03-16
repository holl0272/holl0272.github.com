<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Revision Date: October 13, 2002											*
'*   Version 2.0.2                                                              *
'*                                                                              *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Function sqlSafe(strSQL)

	If isNull(strSQL) Then
		sqlSafe = strSQL
	Else
		sqlSafe = Replace(strSQL,"'","''")
	End If

End Function	'sqlSafe

'--------------------------------------------------------------------------------------------------

Function wrapSQLValue(vntFieldValue, blnEmptyOK, bytFieldType)

Dim pstrTempSQL
Dim pvntTempValue

	Select Case bytFieldType
		Case 0:	'string
			If blnEmptyOK Then
				pstrTempSQL = "'" & sqlSafe(vntFieldValue) & "'"
			Else
				If Len(Trim(vntFieldValue) & "") = 0 Then
					pstrTempSQL = "Null"
				Else
					pstrTempSQL = "'" & sqlSafe(vntFieldValue) & "'"
				End If
			End If
		Case 1: 'number
			If Len(Trim(vntFieldValue) & "") = 0 Then
				pstrTempSQL = "Null"
			Else
				pstrTempSQL = sqlSafe(vntFieldValue)
			End If
		Case 2: 'date
			If Len(Trim(vntFieldValue) & "") = 0 Then
				pstrTempSQL = "Null"
			Else
				If cblnSQLDatabase Then
					pstrTempSQL = "'" & sqlSafe(vntFieldValue) & "'"
				Else
					pstrTempSQL = "#" & sqlSafe(vntFieldValue) & "#"
				End If
			End If
		Case 4: 'boolean
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

Function makeSQLUpdate(strFieldName, vntFieldValue, blnEmptyOK, bytFieldType)

Dim pstrTempSQL

	pstrTempSQL = strFieldName & "=" & wrapSQLValue(vntFieldValue, blnEmptyOK, bytFieldType)
	makeSQLUpdate = pstrTempSQL

End Function	'makeSQLUpdate

'--------------------------------------------------------------------------------------------------

Function sqlDateWrap(strValue)

Dim pstrTempValue

	If Len(strValue & "") = 0 Then
		sqlDateWrap = "Null"
	Else
		If cblnSQLDatabase Then
			sqlDateWrap = "'" & strValue & "'"
		Else
			sqlDateWrap = "#" & strValue & "#"
		End If
	End If

End Function	'sqlDateWrap

'--------------------------------------------------------------------------------------------------

Sub MakeCombo(strSQL,strText,strValue,strSelected)

Dim i
Dim pobjRS
Dim pblnRecordsetCreated
Dim pblnFound

	pblnFound = False
	pblnRecordsetCreated = False

	If isObject(strSQL) Then
		If strSQL.State <> 1 Then
			Response.Write "<option value=" & chr(34) & chr(34) & ">Error: Recordset does not exist</option>" & vbcrlf
			Exit Sub
		End If
		set pobjRS = strSQL
		If pobjRS.recordcount > 0 then pobjRS.movefirst
	Else
		set pobjRS = GetRS(strSQL)
		pblnRecordsetCreated = True
	End If

	for i = 1 to pobjRS.recordcount
		if len(strSelected) > 0 then
			if trim(pobjRS(strValue)) <> trim(strSelected) then
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

Function GetRS(strSQL)

Dim rs

	If Err.number <> 0 Then Err.Clear
	
	Call isValidConnection(cnn)
	
	set	rs = server.CreateObject("adodb.recordset")
	with rs
        .CursorLocation = 2 'adUseClient
        
        On Error Resume Next
		.Open strSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		If Err.number <> 0 Then
			Response.Write "<font color=red>Error in GetRS: Error " & Err.number & ": " & Err.Description & "</font><BR>" & vbcrlf
			Response.Write "<font color=red>Error in GetRS: sql = " & strSQL & "</font><BR>" & vbcrlf
			Response.Flush
			Err.Clear
		End If
	end with
	set GetRS = rs

End Function		' GetRS

'--------------------------------------------------------------------------------------------------

Function isValidConnection(byRef objCnn)
'Check for valid connection

Dim pblnValid

	If Err.number <> 0 Then Err.Clear
	pblnValid = True
	If isObject(objCnn) Then
        On Error Resume Next
		If objCnn.State <> 1 Then pblnValid = False
		If Err.number <> 0 Then
			pblnValid = False
			Err.Clear
		End If
        On Error Goto 0
	Else
		pblnValid = False
	End If
	
	If Not pblnValid Then Response.Write "<font color=red>Error: You do not have a valid connection to the database</font><BR>" & vbcrlf
	isValidConnection = pblnValid
	
End Function	'isValidConnection

'--------------------------------------------------------------------------------------------------

Function DebugPrint(strField,strFieldValue)
  response.write strField & " = " & strFieldValue & "<br>"
end Function

'--------------------------------------------------------------------------------------------------

Function WriteCurrency(strValue)

	If isNull(strValue) Then
		WriteCurrency = FormatCurrency(0,2)
	Elseif len(strValue) = 0 Then
		WriteCurrency = FormatCurrency(0,2)
	Else
		WriteCurrency = FormatCurrency(strValue,2)
	End If
	
End Function	'WriteCurrency

'--------------------------------------------------------------------------------------------------

Function EncodeString(strSource,blnHTML)

dim p_strTemp

	If isNull(strSource) Then
		EncodeString = ""
		Exit Function
	End If
	
	If blnHTML then
		EncodeString = server.HTMLEncode(strSource)
	else
		p_strTemp = Replace(strSource,chr(34),chr(34) & " + String.fromCharCode(34) + " & chr(34))	
		EncodeString = p_strTemp
	end if

End Function	'EncodeString

'--------------------------------------------------------------------------------------------------

Sub ReleaseObject(obj)

Dim pblnNoError
	
On Error Resume Next

	pblnNoError = (Err.number = 0)
	
	obj.Close
	Set obj = Nothing
	
	If pblnNoError And Err.number<> 0 Then Err.Clear
	
End Sub	'ReleaseObject

'--------------------------------------------------------------------------------------------------

Function LoadRequestValue(strSource)

dim p_strTemp

	p_strTemp = Request.QueryString(strSource)
	If len(p_strTemp) = 0 Then p_strTemp = Request.Form(strSource)
	LoadRequestValue = p_strTemp
	
End Function	'LoadRequestValue

'--------------------------------------------------------------------------------------------------

Function FileExists(strFilePath)

Dim pobjFSO

	'On Error Resume Next

	Set pobjFSO = CreateObject("Scripting.FileSystemObject")
	FileExists = pobjFSO.FileExists(strFilePath)
	Set pobjFSO = Nothing

End Function	'FileExists

'--------------------------------------------------------------------------------------------------

Sub ShowStoreFrontVersion

	Response.Write "<HR>"
	Response.Write "<p>The page you are viewing experienced a problem reading from the database. This can be due to a number of issues ranging from incorrect application settings to failure to perform a required database upgrade. The below information is provided to help to determine if it is your application settings.</p>"
	Response.Write "<H4>StoreFront Version: " & Application("AppName") & " (this should be StoreFront or StoreFrontAE)</H4>"
	If cblnSF5AE Then
		Response.Write "-- this script detects you have StoreFront 5.0 AE<br>"
	Else
		Response.Write "-- this script detects you have StoreFront 5.0 SE<br>"
	End If
	Response.Write "<H4>Database Type: " & Application("AppDatabase") & " (this should be Access or SQL)</H4>"
	If cblnSQLDatabase Then
		Response.Write "-- this script detects you are using a SQL Server database<br>"
	Else
		Response.Write "-- this script detects you are using an Access database<br>"
	End If
	Response.Write "<p>If any of the above settings are incorrect you should contact StoreFront Support or your developer. You can manually set these values in ssLibrary/modDatabase.asp.</p>"
	Response.Write "<HR>"
	Response.Flush
	
End Sub	'ShowStoreFrontVersion

'**************************************************
'
'	Start Code Execution
'

Dim mstrPageTitle
Dim cblnUseIntegratedSecurity
Dim cblnSQLDatabase
Dim cblnSF5AE

cblnUseIntegratedSecurity = False	'Set to true to use integrated security. Requires WebStoreManager to work
'session("login") = "Valid"			'Testing Only

cblnSQLDatabase = CBool(Application("AppDatabase") <> "Access")
'cblnSQLDatabase = True				'Set this value to True for SQL Server databases, only need to set this manually for very early versions
'cblnSQLDatabase = False			'Set this value to False for Access databases, only need to set this manually for very early versions

cblnSF5AE = CBool(Application("AppName") = "StoreFrontAE")
'cblnSF5AE = True					'Set this value to True for AE Sites, only need to set this manually for very early versions
'cblnSF5AE = False					'Set this value to False for SE Sites, only need to set this manually for very early versions

'If you are having issues with an add-on you should first verify if the settings are correct by setting the following line to True
If False Then Call ShowStoreFrontVersion

'If you directly call an admin page and get a "Data source name not found and no default driver specified" error
'comment out the following section of code

Dim cnn
Set cnn = Server.CreateObject("ADODB.Connection")
cnn.Open Session("DSN_NAME")

'Remove the comment from the line below and move the code down one line
'<!--#include file="../../../SFLib/db.conn.open.asp"-->
%>
