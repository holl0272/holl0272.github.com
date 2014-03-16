<%
option explicit
Response.Buffer = True
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.1003.0.1

'@FILENAME: sfreports6.asp
	 

'

'@DESCRIPTION:   web reporting tool

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws  and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	
%>
<SCRIPT language="javascript">
function helpMe(){
	var helpWin, loadHelp
	helpWin = window.open('help/daily_sm2f.htm','helpWin', 'scrollbars=1,resizable,location=0,status=0,toolbar=0,menubar=0,height=300,width=500')
	helpWin.focus()
}	
</script>
<!--#include file="../SFLib/incDesign_settings.asp"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/incGeneral.asp"-->
<%


If Request.Form("btnSubmit.x") <> "" Then
	Dim sOrderId, sSQL, rsOrders, sFirstName, sLastName, iCounter, sBgColor, sFontFace, sFontColor, sFontSize

	sOrderId = Request.Form("OrderID")
	sFirstName = Request.Form("FirstName")
	sLastName = Request.Form("LastName")

	sSQL = "Select custID, custFirstName, custLastName, custMiddleInitial, orderID, orderCustID, orderDate, orderGrandTotal " _
		   & "FROM sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId WHERE orderID LIKE '%" & sOrderId & "%'" _
		   & " AND custFirstName LIKE '%" & sFirstName & "%' AND custLastName LIKE '%" & sLastName & "%' and sfOrders.orderIsComplete = 1 Order By orderID"  
	Set rsOrders = CreateObject("ADODB.RecordSet")
	rsOrders.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
	%>
	<html>
	<head>
	<title>SF Reports Page</title>
	</head>

	<body background="<%= C_BKGRND %>" bgproperties="fixed" bgcolor="<%= C_BGCOLOR %>" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

	<form method="post" id="form1" name="form1">
	<table border="0" cellpadding="1" cellspacing="0" bgcolor="<%= C_BORDERCOLOR1 %>" width="<%= C_WIDTH %>" align="center">
	<tr>
	<td>

	    <table width="100%" border="0" cellspacing="1" cellpadding="3">
	    <tr>
	<%	If C_BNRBKGRND = "" Then %>
			<td align="middle" background="<%= C_BKGRND1 %>" bgcolor="<%= C_BGCOLOR1 %>"><b><font face="<%= C_FONTFACE1 %>" color="<%= C_FONTCOLOR1 %>" SIZE="<%= C_FONTSIZE1 %>"><%= C_STORENAME %></font></b></td>
	<%	Else %>
			<td align="middle" bgcolor="<%= C_BNRBGCOLOR %>"><img src="<%= C_BNRBKGRND %>" border="0"></td>
	<%	End If %>        
	    </tr>
	    <tr>
		<td align="middle" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">Transaction Details</font></b></td>        
	    </tr>
	    <tr>
		<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>"><b>Instructions: </b>This reporting tool will allow you to retrieve a detailed report for a single order.  Enter the order ID if you know it, or enter the customer's first or last name.  All matches will be displayed, and you will be able to select the one you are looking for.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a></td>    
	    </tr>
	    <tr>
	    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <tr>
	        <td width="15%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Order Number</font></b></td>        
			<td width="25%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Date</font></b></td>        
			<td width="35%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Customer Name</font></b></td>        
			<td width="25%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Order Total</font></b></td>        
	        </tr>
	        <%
	        If rsOrders.EOF Then
	        %>
				<tr>
				<td colspan="4" align="center" bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>"><font face="<%= C_FONTFACE5 %>" color="#ff0000" size="<%= C_FONTSIZE5+1 %>">There Were no Orders for your Search Criteria</font></td>
				</tr>
	        <%
	        Else
				iCounter = 1
				Do While Not rsOrders.EOF 
					If iCounter mod 2 = 0 Then
						sBgColor = C_ALTBGCOLOR1
						sFontFace = C_ALTFONTFACE1
						sFontColor = C_ALTFONTCOLOR1
						sFontSize = C_ALTFONTSIZE1
					Else
						sBgColor = C_ALTBGCOLOR2
						sFontFace = C_ALTFONTFACE2
						sFontColor = C_ALTFONTCOLOR2
						sFontSize = C_ALTFONTSIZE2
					End If
	        %>
					<tr>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><a href="sfReports1.asp?OrderID=<%= rsOrders.Fields("orderID") %>"><%= rsOrders.Fields("orderID") %></a></font></td>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= rsOrders.Fields("orderDate")%></font></td>
					<td bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= rsOrders.Fields("custFirstName") %>&nbsp;<%= rsOrders.Fields("custMiddleInitial") %>&nbsp;<%= rsOrders.Fields("custLastName") %></font></td>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= FormatCurrency(rsOrders.Fields("orderGrandTotal")) %></font></td>
					</tr>
	        <%
					iCounter = iCounter + 1
					rsOrders.MoveNext 
				Loop
			End If
	        %>
	        <tr>
	        <td width="100%" align="center" valign="top" colspan="3"></td>
	        </tr>
	        </table>
	    </td>
	    </tr>
	   tr>
		<td bgcolor="<%= C_BGCOLOR7 %>" background="<%= C_BKGRND7 %>"><font face="<%= C_FONTFACE7 %>" color="<%= C_FONTCOLOR7 %>" size="<%= C_FONTSIZE7 %>"><p align="center"><b><a href="ssAdmin/admin.asp">Site Administration</a> | <a href="sfReports.asp">Reports</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>
		</table>
	</td>
	</tr>
	</table>
	</form>
	</body>
	<%
	rsOrders.Close 
	Set rsOrders = nothing
	cnn.Close
	Set cnn = nothing
	%>
	</html>
<% Else %>
	<html>
	<head>
	<title>SF Reports Page</title>
	</head>

	<body background="<%= C_BKGRND %>" bgproperties="fixed" bgcolor="<%= C_BGCOLOR %>" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

	<form method="post" id="form1" name="form1">
	<table border="0" cellpadding="1" cellspacing="0" bgcolor="<%= C_BORDERCOLOR1 %>" width="<%= C_WIDTH %>" align="center">
	<tr>
	<td>

	    <table width="100%" border="0" cellspacing="1" cellpadding="3">
	    <tr>
	<%	If C_BNRBKGRND = "" Then %>
			<td align="middle" background="<%= C_BKGRND1 %>" bgcolor="<%= C_BGCOLOR1 %>"><b><font face="<%= C_FONTFACE1 %>" color="<%= C_FONTCOLOR1 %>" SIZE="<%= C_FONTSIZE1 %>"><%= C_STORENAME %></font></b></td>
	<%	Else %>
			<td align="middle" bgcolor="<%= C_BNRBGCOLOR %>"><img src="<%= C_BNRBKGRND %>" border="0"></td>
	<%	End If %>        
	    </tr>
	    <tr>
		<td align="middle" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">
        Retrieve Order</font></b></td>        
	    </tr>
	    <tr>
		<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>"><b>Instructions: </b>The following orders match your criteria.  To view the specifics of a single order, click that record's order ID.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a></td>    
	    </tr>
	    <tr>
	    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%" nowrap>
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <form method="post" name="frm1">
	        <tr>
	        <td align="right"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Order ID</font></td>
	        <td align="left">
            <input type="text" style="<%= C_FORMDESIGN %>" name="OrderID"></td>
	        </tr>
	        <tr>
	        <td align="right"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">First Name</font></td>
	        <td align="left">
            <input type="text" style="<%= C_FORMDESIGN %>" name="FirstName"></td>
	        </tr>
	        <tr>
	        <td align="right"><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Last Name</font></td>
	        <td align="left">
            <input type="text" style="<%= C_FORMDESIGN %>" name="LastName"></td>
	        </tr>
	        <tr>
	        <td width="100%" align="center" valign="top" colspan="3"></td>
	        </tr>
	        <tr>
	        <td colspan="2" width="100%" align="center" valign="top" colspan="4"><input type="image" name="btnSubmit" border="0" src="../<%= C_BTN18 %>" alt="Submit" WIDTH="108" HEIGHT="21"></td>
	        </tr>
	        </form>
	        </table>
	    </td>
	    </tr>
	        <tr>
		<td bgcolor="<%= C_BGCOLOR7 %>" background="<%= C_BKGRND7 %>"><font face="<%= C_FONTFACE7 %>" color="<%= C_FONTCOLOR7 %>" size="<%= C_FONTSIZE7 %>"><p align="center"><b><a href="ssAdmin/admin.asp">Site Administration</a> | <a href="sfReports.asp">Reports</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>
		</table>
	</td>
	</tr>
	</table>
	</form>
	</body>
	</html>
<% End If %>


