<% Option Explicit %>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Sandshot Software Site Checking Page</title>
<link rel="stylesheet" href="../ssLibrary/ssStyleSheet.css" type="text/css">
<script language="vbscript">
	Sub body_onload
		window.parent.frMain.PageLoaded 2, True, ""	
	End Sub	'body_onload
</script>
</head>

<body onload="body_onload">
<%
	Response.Write "<div class='pagetitle2'>.asp Check</div><div id='result'>Success</div>" & vbcrlf
%>
	</BODY>
</HTML>