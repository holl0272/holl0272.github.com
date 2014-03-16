<% Option Explicit %>
<%
Dim pobjCnn
Dim pobjRS
Dim pstrConnection
Dim mblnInitializeConnection
Dim mstrResult
Dim mstrOutput

Const cstrOptionalRecordToCheck = "Select * from sfAdmin"

	pstrConnection = "Driver={mySQL};" & _ 
						"Server=???;" & _
						"Port=3306;" & _
						"Option=131072;" & _
						"Stmt=;" & _
						"Database=???;" & _
						"Uid=???;" & _
						"Pwd=???;"
						
	pstrConnection = ""
	pstrConnection = Application("DSN_NAME")
	If Len(pstrConnection) = 0 Then pstrConnection = Session("DSN_NAME")
	
	Set pobjCnn = Server.CreateObject("ADODB.Connection")
	
	On Error Resume Next
	
	pobjCnn.open pstrConnection
	If err.number = 0 Then
		mblnInitializeConnection = CBool(pobjCnn.State = 1)
		If mblnInitializeConnection Then
			If Len(cstrOptionalRecordToCheck) > 0 Then
				Set pobjRS = Server.CreateObject("ADODB.RECORDSET")
				pobjRS.Open cstrOptionalRecordToCheck, pobjCnn, 3, 1	'adOpenStatic, adLockReadOnly
				If pobjRS.State = 1 Then
					mstrOutput = "<div class='pagetitle2'>Database Check</div><div id='result'>Success</div>"
					mstrResult = ""
				Else
					mstrOutput = "<div id='result'>Failed</div>"
					mstrResult = "Could not open recordset"
				End If
				pobjRS.Close
				Set pobjRS = Nothing
			Else
				mstrOutput = "<div class='pagetitle2'>Database Check</div><div id='result'>Success</div>"
				mstrResult = ""
			End If
		Else
			If Err.number <> 0 Then
				mstrResult = "Error " & err.number & ": " & err.Description
				mstrOutput = "<div id='result'>Failed</div>"
			Else
				mstrResult = "Could not open database"
				mstrOutput = "<div id='result'>Failed</div>"
			End If
		End If
	Else
		mstrResult = "Error " & err.number & ": " & err.Description
		mstrOutput = "<div id='result'>Failed</div><div id='result'>Error " & err.number & ": " & err.Description & "</div>"
	End If
	
	pobjCnn.Close
	Set pobjCnn = Nothing

%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Sandshot Software Site Checking Page</title>
<link rel="stylesheet" href="../ssLibrary/ssStyleSheet.css" type="text/css">
<script language="vbscript">
	Sub body_onload
		window.parent.frMain.PageLoaded 3, True, "<%= mstrResult %>"	
	End Sub	'body_onload
</script>
</head>

<body onload="body_onload">
<%= mstrOutput %>
</body>
</html>
