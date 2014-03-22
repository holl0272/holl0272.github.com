<%@ LANGUAGE="VBScript" %>
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront 99 v.3.0.1
'
'   Author     :   LaGarde, Incorporated
'
'   Notes      :  There are no configurable elements in this file.
'                  
'
'                         COPYRIGHT NOTICE
'
'   The contents of this file is protected under the United States
'   copyright laws as an unpublished work, and is confidential and
'   proprietary to LaGarde, Incorporated.  Its use or disclosure in 
'   whole or in part without the expressed written permission of 
'   LaGarde, Incorporated is expressely prohibited.
'
'   (c) Copyright 1998,1999 by LaGarde, Incorporated.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
<%

	DSN_Name = Request("DSN_Name")
	
	Set Connection = Server.CreateObject("ADODB.Connection")

	Connection.Open "DSN="&DSN_Name&""

	SQLStmt = "SELECT DOMAIN_NAME FROM Admin"

	Set RSAdmin = Connection.Execute(SQLStmt)


SndPage = RSAdmin("DOMAIN_NAME")
Connection.Close
Session.Abandon
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="refresh" content="3; url=<%= SndPage %>">
<title>Order Complete</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body>

<p>&nbsp;</p>
<div align="center"><center>

<table border="0" cellpadding="10" cellspacing="10" width="80%">
  <tr>
    <td><p align="center"><strong>this order has already been completed.&nbsp; one moment
    while we prepare a new shopping session ....&nbsp;&nbsp;&nbsp; thank you for shopping
    gamewearnow.com!</strong></p>
    <p align="center"><img src="images/sm_gwn_logo.jpg" alt="wpe11.jpg (3952 bytes)" WIDTH="183" HEIGHT="48"></td>
  </tr>
</table>
</center></div>
</body>
</html>
