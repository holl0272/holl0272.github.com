<% Option Explicit 
'********************************************************************************
'*   Common Support File			                                            *
'*   Release Version:   2.00.0001												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		November 15, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True
Server.ScriptTimeout = 900
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/ssmodDBCleanup.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Database Clean-up</title>
<link rel="stylesheet" href="ssLibrary/ssStyleSheet.css" type="text/css">
</head>
<body>
<center>
<%
	Call CleanDB
	If Response.Buffer Then Response.Flush
    Call ReleaseObject(cnn)
%>
<p><a href="" onclick="window.close();">Close</a></p>
</center>
</body>
</html>
