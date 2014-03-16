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

<!--#include file="ccerrormsg.htm"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">

<title>Bank Error Message</title>
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body>

<table width="65%" align="center" border="0">
<tr>
<td colspan="2">&nbsp;</td>
</tr>
	<tr>
		<td align="right"><strong>Customer Transaction Number: </strong></td>
		<td><strong><%= Request("CustomerTransactionNumber") %></strong></td>
	</tr>
	<tr>
		<td align="right"><strong>Bank Message: </strong></td>
		<td><strong><%= Request("ErrorMsg") %></strong></td>
	</tr>
	<tr>
		<td align="right"><strong>AVS Code: </strong></td>
		<td><strong><%= Request("AvsCode") %></strong></td>
	</tr>
</table>

</body>
</html>

