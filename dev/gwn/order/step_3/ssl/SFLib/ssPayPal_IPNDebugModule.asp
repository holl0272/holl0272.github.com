<% 
Const cblndebug_PayPalIPN = False	'True	False
'Const cstrDebugLog = "log\debug.txt"
Const cstrDebugLog = ""
Dim mstrWriteToFile

'**************************************************************************************

Sub Output(byVal strOutput)

	If cblndebug_PayPalIPN Then Response.Write strOutput
	Call addToOutputFile(strOutput)

End Sub	'Output

'**************************************************************************************

Sub addToOutputFile(byVal strOutput)

Dim pstrTemp

	If Len(cstrDebugLog) > 0 Then
		pstrTemp = Replace(strOutput, "<BR>" & vbcrlf, vbcrlf)
		pstrTemp = Replace(pstrTemp, "<br>" & vbcrlf, vbcrlf)
		mstrWriteToFile = mstrWriteToFile & pstrTemp
	End If
	
End Sub	'addToOutputFile

'**************************************************************************************

Sub WriteToFile(byVal blnNewWindow, byVal strTextToWrite, byVal blnOverwrite)

Dim fso
Dim txtFile
Dim pstrFilePath
Dim pstrURL
Dim vItem

	If Len(cstrDebugLog) = 0 Then Exit Sub
	If Len(strTextToWrite) = 0 Then Exit Sub

'		On Error Resume Next
	pstrFilePath = Server.Mappath("../") & "\" & cstrDebugLog
	Output "pstrFilePath: " & pstrFilePath & "<BR>" & vbcrlf
	pstrURL = Request.ServerVariables("SERVER_NAME") & "/" & Replace(cstrDebugLog, "\", "/")
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If blnOverwrite Then
		Set txtFile = fso.OpenTextFile(pstrFilePath,2,True)
	Else
		Set txtFile = fso.OpenTextFile(pstrFilePath,8,True)
	End If
	If Err.number <> 0 Then
		debugprint Err.number,Err.Description
	End If
	txtFile.Writeline vbcrlf _
						& "*****************************************************************" & vbcrlf _
						& "" & vbcrlf _
						& "" & Now() & " -- New Entry -- " & vbcrlf _
						& "" & vbcrlf _
						& "*****************************************************************" & vbcrlf
						
	txtFile.Write strTextToWrite
	txtFile.Close
	
	Set txtFile = Nothing
	Set fso = Nothing
	
	If blnNewWindow Then Response.Write "<script>window.open('http://" & pstrURL & "');</script>"

End Sub	'WriteToFile

'**************************************************************************************
%>