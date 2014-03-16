<%
'********************************************************************************
'*   Common Support File			                                            *
'*   Release Version:	2.01.006		                                        *
'*   Revision Date:		April 27, 2006											*
'*                                                                              *
'*   Release 2.01.006 (April 27, 2006)									        *
'*	   - Merged in new capabilities												*
'*																				*
'*   Release 2.01.001 (May 29, 2004)									        *
'*	   - First official release													*
'*																				*
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'**********************************************************
'*	Global variables
'**********************************************************

Dim debug
Dim cblnssMainDebug

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	'Const cstrDebugLog = "fpdb\test_debug_log.txt"
	Const cstrDebugLog = ""
	Const cblnOverwriteLog = True
	cblnssMainDebug = True

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Immediate Code Execution
'**********************************************************

	Call InitializeDebugging(cstrDebugLog, cblnssMainDebug)
	
	'Common overrides
	'debug.Enabled = True	'True	False
	debug.OverwriteLog = cblnOverwriteLog
	'debug.OutputFilePath = ""

	debug.DisplayInScreen = ssDebug_General	'True	False	'Applies to debugfile only

	'debug.Enabled = True
	'debug.PrintForm
	'Call DebugRecordTime("<b>Ending Page</b>")

'DebugRecordSplitTime(z)
	'debug.PrintServerVariables

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Functions
'**********************************************************
	'Sub InitializeDebugging(byVal strFilePath, byVal blnEnabled)
	'Sub cleanup_ssclsDebug()
	'Sub WriteDebugLog()
	'Sub DebugPrint(byVal strField, byVal strFieldValue)
	'Sub DebugPrintLater(byVal strField, byVal strFieldValue)
	'Sub DebugToJavascriptAlert(byVal strField, byVal strFieldValue)
	'Sub DebugRecordTime(byVal strMessage)
	'Sub OutputDebuggingNow()

Class clsDebug

	Private mb_Enabled
	Private md_BeginningTime
	Private md_StartTime
	Private md_SplitTime
	Private mo_Storage
	Private pstrFileContents
	Private pstrOutputFilePath
	Private pblnDisplayInScreen
	Private pblnNewFileWritten
	Private pblnOverwriteLog
	Private parySplitTimes
	Private plngSplitPosition

	'--------------------------------------------------------------------------------------------------

	Private Sub Class_Initialize()
		md_BeginningTime = Now()
		md_StartTime = Timer
		md_SplitTime = md_StartTime
		Set mo_Storage = CreateObject("Scripting.Dictionary")
		mb_Enabled = False
		pblnDisplayInScreen = False
		pblnNewFileWritten = False
		pblnOverwriteLog = False
		plngSplitPosition = -1
		ReDim parySplitTimes(0)
	End Sub

	Private Sub Class_Terminate()
		Call WriteDebugFile(False, "")
		Set mo_Storage = Nothing
	End Sub
	
	'--------------------------------------------------------------------------------------------------

	Public Property Get Enabled()
		Enabled = mb_Enabled
	End Property
	Public Property Let Enabled(bNewValue)
		mb_Enabled = bNewValue
	End Property

	Public Property Let OutputFilePath(bNewValue)
		pstrOutputFilePath = bNewValue
	End Property
	Public Property Get OutputFilePath
		OutputFilePath = pstrOutputFilePath
	End Property

	Public Property Let DisplayInScreen(bNewValue)
		pblnDisplayInScreen = bNewValue
	End Property
	Public Property Get DisplayInScreen
		DisplayInScreen = pblnDisplayInScreen
	End Property

	Public Property Let OverwriteLog(bNewValue)
		pblnOverwriteLog = bNewValue
	End Property
	Public Property Get OverwriteLog
		OverwriteLog = pblnOverwriteLog
	End Property

	'--------------------------------------------------------------------------------------------------

	Public Sub PrintLater(label, output)
		If Enabled Then Call mo_Storage.Add(label, output)
	End Sub

	'--------------------------------------------------------------------------------------------------

	Public Sub [End]()
		If Enabled Then
			Call PrintSummaryInfo()
			Call PrintCollection ("VARIABLE STORAGE", mo_Storage)
			Call PrintCollection ("QUERYSTRING COLLECTION", Request.QueryString())
			Call PrintCollection("FORM COLLECTION", Request.Form())
			Call PrintCollection ("COOKIES COLLECTION", Request.Cookies())
			Call PrintCollection ("SERVER VARIABLES COLLECTION", Request.ServerVariables())
		End If
	End Sub

	'--------------------------------------------------------------------------------------------------

	Public Sub PrintStorage()
		If Enabled Then Call PrintCollection ("VARIABLE STORAGE", mo_Storage)
	End Sub

	'--------------------------------------------------------------------------------------------------

	Public Sub PrintQuerystring()
		If Enabled Then Call PrintCollection ("QUERYSTRING COLLECTION", Request.QueryString())
	End Sub

	'--------------------------------------------------------------------------------------------------

	Public Sub PrintForm()
		If Enabled Then Call PrintCollection("FORM COLLECTION", Request.Form())
	End Sub

	'--------------------------------------------------------------------------------------------------

	Public Sub PrintCookies()
		If Enabled Then Call PrintCollection ("COOKIES COLLECTION", Request.Cookies())
	End Sub

	'--------------------------------------------------------------------------------------------------

	Public Sub PrintServerVariables()
		If Enabled Then Call  PrintCollection ("SERVER VARIABLES COLLECTION", Request.ServerVariables())
	End Sub

	'--------------------------------------------------------------------------------------------------

	Public Sub Print(byVal strField, byVal strFieldValue)
	
	Dim i
	
		If Not mb_Enabled Then Exit Sub
	
		If isArray(strFieldValue) Then
			For i = 0 To UBound(strFieldValue)
				If isArray(strFieldValue(i)) Then
					Call Print(strField & "(" & i & ")", strFieldValue(i))
				Else
					Output  strField & "(" & i & ") = " & strFieldValue(i) & "<br />" & vbcrlf
				End If
			Next 'i
		Else
			Output  strField & " = " & strFieldValue & "<br />" & vbcrlf
		End If
		
	End Sub	'Print

	'--------------------------------------------------------------------------------------------------

	Public Sub Print2(byVal strText)
	
	Dim i
	
		If Not mb_Enabled Then Exit Sub
	
		Output  strText
		
	End Sub	'Print

	'--------------------------------------------------------------------------------------------------

	Public Sub PrintSummaryInfo()

		If Not mb_Enabled Then Exit Sub
	
		Output "<hr>" & vbcrlf
		Output "<b>SUMMARY INFO</b><br />" & vbcrlf
		Output "Time of Request=" & md_BeginningTime & "<br />" & vbcrlf
		Output "Elapsed Time=" & ElapsedTime & " seconds<br />" & vbcrlf
		Output "Request Type=" & Request.ServerVariables("REQUEST_METHOD") & "<br />" & vbcrlf
		Output "Status Code=" & Response.Status & "<br />" & vbcrlf
		
		Output "<fieldset style=""background-color: white;color:black""><legend>" & Name & "</legend><table border=1 cellspacing=0 cellpadding=2><colgroup><col align=left /><col align=right /><col align=right /></colgroup><tr><th>Item</th><th>Value</th></tr>" & vbcrlf
		For Each vItem In Collection
			Output "<tr><td>" & vItem & "</td><td>" & Collection(vItem) & "</td></tr>" & vbcrlf
		Next
		Output "</table></fieldset>" & vbcrlf

	End Sub

	'--------------------------------------------------------------------------------------------------

	Public Function ElapsedTime()
	
	Dim EndTime
	
		EndTime = Timer

		'Watch for the midnight wraparound...
		If EndTime < md_StartTime Then EndTime = EndTime + (86400)

		ElapsedTime = EndTime - md_StartTime
		
	End Function	'ElapsedTime

	'--------------------------------------------------------------------------------------------------

	Public Function SplitTime()
	
	Dim EndTime
	
		EndTime = Timer

		'Watch for the midnight wraparound...
		If EndTime < md_SplitTime Then EndTime = EndTime + (86400)

		SplitTime = EndTime - md_SplitTime
		md_SplitTime = EndTime
		
	End Function	'SplitTime

	'--------------------------------------------------------------------------------------------------

	Public Sub PrintCollection(Byval Name, ByVal Collection)
	
	Dim vItem
	
		If Not mb_Enabled Then Exit Sub
	
		Output "<fieldset style=""background-color: white;color:black""><legend>" & Name & "</legend>" & vbcrlf
		For Each vItem In Collection
			Output vItem & "=" & Collection(vItem) & "<br />" & vbcrlf
		Next
		Output  "</fieldset><hr />" & vbcrlf
		
	End Sub	'PrintCollection

	'--------------------------------------------------------------------------------------------------
	
	Public Sub PrintApplication()
	
	Dim i
	Dim pstrKey
	Dim pvntValue
	
		Output "<fieldset style=""background-color: white;color:black""><legend>Application Contents (" & Application.Contents.Count + 1 & ")</legend><table border=1 cellspacing=0><tr><th>Item</th><th>Value</th></tr>"
		For i = 0 To Application.Contents.Count
			pstrKey = Application.Contents.Key(i)
			pvntValue = Application.Contents.Item(pstrKey)
			If isArray(pvntValue) Then
				Output "<tr><td>" & pstrKey & "</td>"
				Output "<td>Array: " & UBound(pvntValue) & "<br />"
				Call writeArray(pvntValue, "- ")
				Output "</td></tr>"
			Else
				Output "<tr><td>" & pstrKey & "</td><td>" & pvntValue & "</td></tr>"
			End If
		Next
		Output "</table></fieldset>"
	End Sub

	'--------------------------------------------------------------------------------------------------
	
	Public Sub PrintSession()
	
	Dim i
	Dim pstrKey
	Dim pvntValue
	
		Output "<fieldset style=""background-color: white;color:black""><legend>Session Contents (" & Session.Contents.Count + 1 & ")</legend><table border=1 cellspacing=0><tr><th>Item</th><th>Value</th></tr>"
		For i = 0 To Session.Contents.Count
			pstrKey = Session.Contents.Key(i)
			pvntValue = Session.Contents.Item(pstrKey)
			If isArray(pvntValue) Then
				Output "<tr><td>" & pstrKey & "</td>"
				Output "<td>Array: " & UBound(pvntValue) & "<br />"
				Call writeArray(pvntValue, "- ")
				Output "</td></tr>"
			Else
				Output "<tr><td>" & pstrKey & "</td><td>" & pvntValue & "</td></tr>"
			End If
		Next
		Output "</table></fieldset>"
	End Sub

	'--------------------------------------------------------------------------------------------------
	
	Private Sub writeArray(byRef ary, byVal strPrefix)
	
	Dim i
	
		For i = 0 To UBound(ary)
			On Error Resume Next
			If isArray(ary(i)) Then
				Call writeArray(ary(i), strPrefix & "- ")
				If Err.number <> 0 Then
					Output "Error writing array: " & UBound(ary, 2) & "<br />"
					Output "UBound: " & UBound(ary, 2) & "<br />"
				End If
			Else
				Output strPrefix & i & ": " & ary(i) & "<br />"
			End If
		Next 'i
	
	End Sub

	'--------------------------------------------------------------------------------------------------
	
	Private Function currentPath()
	
	Dim pstrPath
	
		pstrPath = Replace(LCase(Server.MapPath("clsDebug.asp")), LCase("clsDebug.asp"), "")
		
		currentPath = pstrPath

	End Function	'currentPath

	'--------------------------------------------------------------------------------------------------
	
	Public Sub WriteToFile(blnNewWindow, strFile)
	
	Dim fso
	Dim txtFile
	Dim pstrFilePath
	Dim pstrURL
	Dim vItem
	
		pstrFilePath = strFile
		If Len(pstrFilePath) = 0 Then pstrFilePath = logFilePath
		If Len(pstrFilePath) = 0 Then Exit Sub
	
		pstrFilePath = Replace(pstrFilePath, "<path>\", currentPath)
		pstrFilePath = Replace(pstrFilePath, "<path>", currentPath)
		
		On Error Resume Next
'		pstrFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & cstrDebugLog

		pstrURL = Request.ServerVariables("SERVER_NAME") & "/" & pstrOutputFilePath
		Set fso = CreateObject("Scripting.FileSystemObject")

		If Not pblnNewFileWritten And pblnOverwriteLog Then
			Set txtFile = fso.OpenTextFile(pstrFilePath,2,True)
		Else
			Set txtFile = fso.OpenTextFile(pstrFilePath,8,True)
		End If
		If Err.number <> 0 Then
			If Err.number = 76 Then
				Print "clsDebug - WriteToFile Error " & Err.number, Err.Description & " (" & pstrFilePath & ")"
			Else
				Print "clsDebug - WriteToFile Error " & Err.number,Err.Description
			End If
			Err.Clear
		End If
		For Each vItem In mo_Storage
			txtFile.WriteLine vItem & "=" & mo_Storage(vItem)
		Next
		txtFile.Close
		
		Set txtFile = Nothing
		Set fso = Nothing
		
		If blnNewWindow Then Response.Write "<script>window.open('http://" & pstrURL & "');</script>"
	
	End Sub	'WriteToFile

	'--------------------------------------------------------------------------------------------------
	
	Private Sub logFilePath(byVal strOverrideFilePath, byRef strOutputFilePath, byRef strURL)
	
	Dim pstrTempFilePath
	
		pstrTempFilePath = strOverrideFilePath
		If Len(pstrTempFilePath) = 0 Then pstrTempFilePath = pstrOutputFilePath
		If Len(pstrTempFilePath) = 0 Then Exit Sub

'		On Error Resume Next
		'pstrFilePath = Server.Mappath("../") & "\" & pstrOutputFilePath
		
		strOutputFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\" & pstrTempFilePath
		strURL = Request.ServerVariables("SERVER_NAME") & "/" & Replace(pstrTempFilePath, "\", "/")
		'Response.Write "logFilePath: " & strOutputFilePath & "<br />" & vbcrlf
	
	End Sub	'logFilePath

	'--------------------------------------------------------------------------------------------------

	Public Sub WriteDebugFile(byVal blnNewWindow, byVal strOverrideFilePath)

	Dim pstrFilePath
	Dim pstrURL

		If Not mb_Enabled Then Exit Sub

		If Len(pstrFileContents) = 0 Then Exit Sub
		Call logFilePath(strOverrideFilePath, pstrFilePath, pstrURL)
		If Len(pstrFilePath) = 0 Then Exit Sub
		
		If Not pblnNewFileWritten Then
			pstrFileContents = vbcrlf _
							& "*****************************************************************" & vbcrlf _
							& "" & vbcrlf _
							& "" & Now() & " -- New Entry -- " & vbcrlf _
							& "" & vbcrlf _
							& "*****************************************************************" & vbcrlf _
							& vbcrlf _
							& pstrFileContents
			Call WriteFile(pstrFilePath)
			pblnNewFileWritten = True
		Else
			Call WriteFile(pstrFilePath)
		End If
		
		If pblnDisplayInScreen And blnNewWindow Then Response.Write "<script>window.open('http://" & pstrURL & "');</script>"
	
	End Sub	'WriteDebugFile

	'--------------------------------------------------------------------------------------------------

	Private Sub WriteFile(byVal strFilePath)
	
	Dim fso
	Dim txtFile
	Dim vItem
	
		If Len(pstrFileContents) = 0 Then Exit Sub
		If Len(strFilePath) = 0 Then Exit Sub

		On Error Resume Next

		If Err.number <> 0 Then	Err.Clear

		Set fso = CreateObject("Scripting.FileSystemObject")

		If Not pblnNewFileWritten And pblnOverwriteLog Then
			Set txtFile = fso.OpenTextFile(strFilePath,2,True)
		Else
			Set txtFile = fso.OpenTextFile(strFilePath,8,True)
		End If
		
		If Err.number = 0 Then
			txtFile.Write pstrFileContents
			txtFile.Close
		Else
			Output "WriteFile Error " & Err.number & ": " & Err.Description & "<br />" & vbcrfl
		End If
		
		Set txtFile = Nothing
		Set fso = Nothing
		
		'Now clear the contents
		pstrFileContents = ""
		
	End Sub	'WriteFile

	'--------------------------------------------------------------------------------------------------
	
	Private Sub Output(byVal strOutput)
	
		If Not mb_Enabled Then Exit Sub
		If pblnDisplayInScreen Then Response.Write strOutput
		Call addToOutputFile(strOutput)
	
	End Sub	'Output

	'--------------------------------------------------------------------------------------------------

	Private Sub addToOutputFile(byVal strOutput)
	
	Dim pstrTemp

		If Len(cstrDebugLog) > 0 Then
			pstrTemp = Replace(strOutput, "<br />" & vbcrlf, vbcrlf)
			pstrTemp = Replace(pstrTemp, "<br />" & vbcrlf, vbcrlf)
			pstrFileContents = pstrFileContents & pstrTemp
		End If
		
	End Sub	'addToOutputFile

	'--------------------------------------------------------------------------------------------------

	Public Sub RecordTime(byVal strMessage)

	Dim pElapsedTime: pElapsedTime = ElapsedTime
	Dim pSplitTime: pSplitTime = SplitTime
	
		If CBool(Len(Session("ssDebug_TimingComplete")) > 0) Then
			Output "<fieldset style=""background-color: white;color:black""><legend>RecordTime</legend>"
			Output "Message: " & strMessage & "<br />" & vbcrlf
			Output "Start Time: " & FormatDateTime(md_BeginningTime, vbLongTime) & "<br />" & vbcrlf
			Output "Current Time: " & FormatDateTime(Now(), vbLongTime) & "<br />" & vbcrlf
			If pElapsedTime > 0 Then Output "Elapsed Time: " & FormatNumber(ElapsedTime, 4) & " seconds.<br />" & vbcrlf
			If pSplitTime > 0 Then Output "Split Time: " & FormatNumber(pSplitTime, 4) & " seconds.<br />" & vbcrlf
			Output "</fieldset>" & vbcrlf & vbcrlf
		End If
		
		Call SaveSplitTime(strMessage, ElapsedTime)
		
	End Sub	'RecordTime

	'--------------------------------------------------------------------------------------------------

	Public Sub RecordSplitTime(byVal strMessage)

	Dim pElapsedTime: pElapsedTime = ElapsedTime
	Dim pSplitTime: pSplitTime = SplitTime
	
		If CBool(Len(Session("ssDebug_TimingComplete")) > 0) Then
			Output "<fieldset style=""background-color: white;color:black""><legend>" & strMessage & "</legend>"
			If pElapsedTime > 0 Then Output "Elapsed Time: " & FormatNumber(ElapsedTime, 4) & " seconds.<br />" & vbcrlf
			If pSplitTime > 0 Then Output "Split Time: " & FormatNumber(pSplitTime, 4) & " seconds.<br />" & vbcrlf
			Output "</fieldset>" & vbcrlf & vbcrlf
		End If
		Call SaveSplitTime(strMessage, ElapsedTime)
		
	End Sub	'RecordSplitTime

	'--------------------------------------------------------------------------------------------------
	
	Private Sub SaveSplitTime(byVal strMessage, byVal dtTime)
	
		plngSplitPosition = plngSplitPosition + 1
		If plngSplitPosition > UBound(parySplitTimes) Then ReDim Preserve parySplitTimes(plngSplitPosition + 9)
		
		parySplitTimes(plngSplitPosition) = Array(strMessage, dtTime)
		
	End Sub	'SaveSplitTime

	'--------------------------------------------------------------------------------------------------
	
	Public Sub WriteSplitTimes()
	
	Dim i
	Dim pdtPreviousTime
	
		Output "<fieldset style=""background-color: white;color:black""><legend>Split Times</legend>"
		Output "<table border=1 cellpadding=2 cellspacing=0><colgroup><col align=left /><col align=center /><col align=center /></colgroup>"
		Output "<tr><th>Item</th><th>Time</th><th>Elapsed</th></tr>" & vbcrlf
		pdtPreviousTime = 0
		For i = 0 To plngSplitPosition
			Output "<tr><td>" & parySplitTimes(i)(0) & "</td>"
			Output "<td>" & FormatNumber(parySplitTimes(i)(1), 4) & "</td>"
			If parySplitTimes(i)(1) - pdtPreviousTime > 0 Then
				Output "<td>" & FormatNumber(parySplitTimes(i)(1) - pdtPreviousTime, 4) & "</td>"
			Else
				Output "<td>-</td>"
			End If
			Output "</tr>" & vbcrlf
			pdtPreviousTime = parySplitTimes(i)(1)
		Next 'i
		Output "</table>" & vbcrlf & vbcrlf
		Output "</fieldset>" & vbcrlf & vbcrlf
		
	End Sub	'WriteSplitTimes
	

End Class	'clsDebug

'********************************************************************************************
'********************************************************************************************

Sub InitializeDebugging(byVal strFilePath, byVal blnEnabled)

	Set debug = New clsDebug
	debug.Enabled = blnEnabled
	debug.OutputFilePath = strFilePath

End Sub	'InitializeDebugging

'********************************************************************************************

Sub WriteDebugLog()

	If isObject(debug) Then
		debug.WriteDebugFile False, ""
	End If

End Sub	'WriteDebugLog

'********************************************************************************************

Sub DebugPrint(byVal strField, byVal strFieldValue)
	If isObject(debug) Then
		debug.Print strField, strFieldValue
	Else
		If cblnssMainDebug Then Response.Write strField & " = " & strFieldValue & "<br />" & vbcrlf
	End If
End Sub	'DebugPrint

'********************************************************************************************

Sub Output(byVal strOut)
	If isObject(debug) Then
		debug.Print2 strOut
	Else
		If cblnssMainDebug Then Response.Write strOut
	End If
End Sub	'DebugPrint

'********************************************************************************************

Function DebugPrintRecordset(byVal strField, byRef objRS)

Dim i

	Output "<fieldset style=""background-color: white;color:black""><legend>" & strField & "</legend><table border=1 cellspacing=0><tr><th>Item</th><th>Value</th></tr>"
	For i = 1 To objRS.Fields.Count
		Output "<tr><td>(" & i & ") " & objRS.Fields(i-1).Name & "</td><td>&nbsp;" & objRS.Fields(i-1).Value & "</td></tr>"
	Next
	Output "</table></fieldset>"

End Function

'********************************************************************************************

Function DebugPrintRecordset_Complete(byVal strField, byRef objRS)

Dim i
Dim j

	j = 0
	
	Output "<fieldset style=""background-color: white;color:black""><legend>" & strField & "</legend>"
	Output "<table border=1 cellspacing=0><tr><th></th>"
	With objRS
		For i = 1 To .Fields.Count
			Output "<th>" & .Fields(i-1).Name & "</th>"
		Next
		Output "</tr>"
	
		Do While Not .EOF
			j = j + 1
			Output "<tr><td>" & j & ".</td>"
			For i = 1 To objRS.Fields.Count
				Output "<td>&nbsp;" & objRS.Fields(i-1).Value & "</td>"
			Next
		
			Output "</tr>"
			.MoveNext
		Loop
		.MoveFirst
	End With
	Output "</table></fieldset>"

End Function

'********************************************************************************************

Sub DebugPrintLater(byVal strField, byVal strFieldValue)
	If isObject(debug) Then
		debug.PrintLater strField, strFieldValue
	Else
		If cblnssMainDebug Then Response.Write strField & " = " & strFieldValue & "<br />" & vbcrlf
	End If
End Sub	'DebugPrintLater

'********************************************************************************************

Sub DebugToJavascriptAlert(byVal strField, byVal strFieldValue)
	If isObject(debug) Then
   		If cblnssMainDebug Then Response.Write "alert(" & Chr(34) & strField & ": " & strFieldValue & Chr(34) & ");" & vbcrlf
		Call DebugPrintLater(strField, strFieldValue)
	End If
End Sub	'DebugToJavascriptAlert

'********************************************************************************************

Sub DebugRecordTime(byVal strMessage)
	If isObject(debug) Then
		debug.RecordTime strMessage
		Call OutputDebuggingNow
	End If
End Sub	'DebugRecordTime

'********************************************************************************************

Sub DebugRecordSplitTime(byVal strMessage)
	If isObject(debug) Then
		debug.RecordSplitTime strMessage
		Call OutputDebuggingNow
	End If
End Sub	'DebugRecordSplitTime

'********************************************************************************************

Sub OutputDebuggingNow()
	Call WriteDebugLog
	If isObject(debug) Then
	'	If debug.DisplayInScreen Then Response.Flush
	Else
	'	If cblnssMainDebug Then Response.Flush
	End If
End Sub	'OutputDebuggingNow

'********************************************************************************************

Sub writeTimer(byVal strText)

Dim pdtCurrentTime

	pdtCurrentTime = Timer
	If Len(CStr(mdtStartTime)) = 0 Then
		mdtStartTime = pdtCurrentTime
		mdtLastTime = pdtCurrentTime
	End If
	
	Response.Write "<fieldset style=""background-color: white;color:black""><legend>" & strText & "</legend><div align=left>" _
				 & "Elapsed Time: " & pdtCurrentTime - mdtStartTime & " seconds<br />" _
				 & "Step Time: " & pdtCurrentTime - mdtLastTime & " seconds" _
				 & "</div></fieldset>"
				 
	mdtLastTime = pdtCurrentTime

End Sub

'********************************************************************************************

Sub cleanup_ssclsDebug

	If isObject(debug) Then
		debug.RecordTime "<b>Ending Page</b>"
		Call OutputDebuggingNow
		If ssDebug_General Then Call writeFinalGeneralDebugging
		debug.WriteSplitTimes
		debug.WriteToFile False, ""
		If CBool(Len(Session("ssDebug_ShowServerVariables")) > 0) Then debug.PrintCollection "SERVER VARIABLES COLLECTION", Request.ServerVariables()
		If CBool(Len(Session("ssDebug_ShowApplicationVariables")) > 0) Then debug.PrintApplication
		If CBool(Len(Session("ssDebug_ShowSessionVariables")) > 0) Then debug.PrintSession
		Set debug = Nothing
	End If

End Sub

'********************************************************************************************

Sub RecordHangingActiveDBConnections(byVal strPageName)

Dim pstrHangingPage

	pstrHangingPage = Application("HangingActiveDBConnections")
	If Len(pstrHangingPage) > 0 Then
		Application("HangingActiveDBConnections") = pstrHangingPage & "|" & Now() & ";" & strPageName
	Else
		Application("HangingActiveDBConnections") = Now() & ";" & strPageName
	End If

End Sub

'***************************************************************************************************************************************

Sub addSessionDebugMessage(byVal strMessage)
	'If True Then Exit Sub
	Dim pstrMessage
	pstrMessage = Session("globalDebuggingMessage")
	If Len(pstrMessage) = 0 Then
		Session("globalDebuggingMessage") = strMessage
	Else
		Session("globalDebuggingMessage") = pstrMessage & "|" & strMessage
	End If
End Sub	'addSessionDebugMessage

'********************************************************************************************

Sub RecordHangingUnhandledError()

Dim pstrUnhandledErrors
Dim pstrError

	pstrError = Now() & ";" & CurrentPage & ";" & Err.Number & ";" & Err.Description
	pstrUnhandledErrors = Application("UnhandledErrors")
	
	If Len(pstrUnhandledErrors) > 0 Then
		Application("UnhandledErrors") = pstrUnhandledErrors & "|" & pstrError
	Else
		Application("UnhandledErrors") = pstrError
	End If

End Sub

'********************************************************************************************

Sub writeFinalGeneralDebugging

	Response.Write "<fieldset style=""background-color: white;color:black"">"
	Response.Write "	<legend>General Debugging Information</legend>"
	Response.Write "	<table border=1 cellspacing=0>"
	Response.Write "	<tr><th colspan=2>Cookies</th></tr>"
	Response.Write "	<tr><td>Session.SessionID</td><td>&nbsp;" & Session.SessionID & "</td></tr>"
	Response.Write "	<tr><td>visitorLoggedInCustomerID</td><td>&nbsp;" & visitorLoggedInCustomerID & "</td></tr>"
	Response.Write "	<tr><td>Cookie_custID</td><td>&nbsp;" & custID_cookie & "</td></tr>"
	Response.Write "	<tr><td>Cookie_visitorID</td><td>&nbsp;" & getCookie_visitorID & "</td></tr>"
	Response.Write "	<tr><td>Cookie_SessionID</td><td>&nbsp;" & getCookie_SessionID & "</td></tr>"
	Response.Write "	<tr><td>visitorCertificateCodes</td><td>&nbsp;" & visitorCertificateCodes & "</td></tr>"
	Response.Write "	<tr><td>sfAddProduct</td><td>&nbsp;" & Replace(Replace(getCookie_sfAddProduct, "&", "<br />"), "http://localhost/MasterTemplate/", "") & "</td></tr>"
	Response.Write "	<tr><td>sfSearch</td><td>&nbsp;" & Replace(Replace(getCookie_sfSearch, "&", "<br />"), "http://localhost/MasterTemplate/", "") & "</td></tr>"
	Response.Write "	<tr><th colspan=2>Other</th></tr>"
	Response.Write "	<tr><td>SessionID</td><td>&nbsp;" & SessionID & "</td></tr>"
	Response.Write "	<tr><td>Viewed</td><td>&nbsp;" & Replace(visitorRecentlyViewedProducts, "|", "<br />") & "</td></tr>"
	Response.Write "	</table>"
	Response.Write "<a href=""ssDebuggingConsole.asp"">Debugging Console</a>"
	Response.Write "</fieldset>"

End Sub	'writeFinalGeneralDebugging

'********************************************************************************************

Function ssDebug_DyanmicProduct
	ssDebug_DyanmicProduct = CBool(Len(Session("ssDebug_DyanmicProduct")) > 0)
End Function

Function ssDebug_FraudScore
	ssDebug_FraudScore = CBool(Len(Session("ssDebug_FraudScore")) > 0)
End Function

Function ssDebug_General
	ssDebug_General = CBool(Len(Session("ssDebug_General")) > 0)
End Function

Function ssDebug_ProductDisplay
	ssDebug_ProductDisplay = CBool(Len(Session("ssDebug_ProductDisplay")) > 0)
End Function

Function ssDebug_Download
	ssDebug_Download = CBool(Len(Session("ssDebug_Download")) > 0)
End Function

Function cblnDebugCategorySearchTool
	cblnDebugCategorySearchTool = CBool(Len(Session("ssDebug_CategorySearchTool")) > 0)
End Function

Function cblnDebugCMS
	cblnDebugCMS = CBool(Len(Session("ssDebug_CMS")) > 0)
End Function

%>
