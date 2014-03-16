<%@ Language=VBScript %>
<%
option explicit
Response.Buffer = True
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.1003.0.1

'@FILENAME: sfreports3.asp
	 
'Access Version
'

'@DESCRIPTION:   web reporting tool

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	
%>
<HTML>
<title>SF Reports Page</title>
<HEAD>
</HEAD>
<SCRIPT language="javascript">
function helpMe(){
	var helpWin, loadHelp
	helpWin = window.open('help/daily_sm2d.htm','helpWin', 'scrollbars=1,resizable,location=0,status=0,toolbar=0,menubar=0,height=300,width=500')
	helpWin.focus()
}	
</script>
<!--#include file="../SFLib/incDesign_settings.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->

<%
Dim sStartDate, sEndDate, rsTrans, sSQL, sOrderId

sStartDate = MakeUSDate(Request.QueryString("startDate"))
sEndDate = MakeUSDate(Request.QueryString("endDate"))
sOrderId = Request.QueryString("OrderId")

Set rsTrans = CreateObject("ADODB.RecordSet")

If sOrderId = "" Then
	sSQL = "SELECT orderID, orderDate, trnsrspID, trnsrspOrderId, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode, trnsrspRetrievalCode, trnsrspAuthNo, trnsrspErrorMsg, trnsrspErrorLocation, trnsrspSuccess " _ 
		& "FROM sfOrders INNER JOIN sfTransactionResponse ON sfOrders.orderID = sfTransactionResponse.trnsrspOrderId WHERE orderDate BETWEEN " & wrapSQLValue(sStartDate, False, enDatatype_date) & " AND " & wrapSQLValue(sEndDate, False, enDatatype_date) & "  and sfOrders.orderIsComplete = 1 ORDER BY orderID"
Else
	sSQL = "SELECT orderID, orderDate, trnsrspID, trnsrspOrderId, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode, trnsrspRetrievalCode, trnsrspAuthNo, trnsrspErrorMsg, trnsrspErrorLocation, trnsrspSuccess " _ 
		& "FROM sfOrders INNER JOIN sfTransactionResponse ON sfOrders.orderID = sfTransactionResponse.trnsrspOrderId WHERE orderID = " & sOrderId & " and sfOrders.orderIsComplete = 1 ORDER BY orderID"
End If
rsTrans.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
%>
<body background="<%= C_BKGRND %>" bgproperties="fixed" bgcolor="<%= C_BGCOLOR %>" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">
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
	<td align="middle" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">Transaction Services</font></b></td>        
    </tr>
    <tr>
	<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>">Transaction information for each sale within the date range specified is listed below.  This report contains information returned by the payment processing service including error or authorization codes.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a></td>    
    </tr>
    <tr>
    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%" nowrap>
        <table border="0" width="100%" cellpadding="4" cellspacing="0">
<%if rsTrans.EOF and rsTrans.BOF then%>
   <tr>
				<td colspan=4 align="center" bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>"><font face="<%= C_FONTFACE5 %>" color="#ff0000" size="<%= C_FONTSIZE5+1 %>">There Were No Transactions Between <%= sStartDate %> And <%= sEndDate %></font></td>
				</tr>
<%else%>				

        <tr>
		<td width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>" colspan="4"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Report for <%= sStartDate %> to <%= sEndDate %></font></b></td>        
        </tr>
<%
Do While Not rsTrans.EOF
%>
	<tr>
	<td colspan=4 width="90%"><hr width="100%"></td>
	</tr>
	<tr>
	<td colspan=4>
	<table border="1" width="90%" align="center">
	<tr>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Order ID:&nbsp;<%= rsTrans.Fields("orderID") %></font></td>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Order Date:&nbsp;<%= rsTrans.Fields("orderDate") %></font></td>
	</tr>
	<tr>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Authorization #:&nbsp;<%= rsTrans.Fields("trnsrspAuthNo") %></font></td>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Success:&nbsp;<%= rsTrans.Fields("trnsrspSuccess") %></font></td>
	</tr>
	<tr>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Customer Transaction #:&nbsp;<%= rsTrans.Fields("trnsrspCustTransNo") %></font></td>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Merchant Transaction #:&nbsp;<%= rsTrans.Fields("trnsrspMerchTransNo") %></font></td>
	</tr>
	<tr>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">AVS Code:&nbsp;<%= rsTrans.Fields("trnsrspAVSCode") %></font></td>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">AUX Message:&nbsp;<%= rsTrans.Fields("trnsrspAUXMsg") %></font></td>
	</tr>
	<tr>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Action Code:&nbsp;<%= rsTrans.Fields("trnsrspActionCode") %></font></td>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Retrieval Code:&nbsp;<%= rsTrans.Fields("trnsrspRetrievalCode") %></font></td>
	</tr>
	<tr>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Error Message:&nbsp;<%= rsTrans.Fields("trnsrspErrorMsg") %></font></td>
	<td align="Left"><font face="<%= C_FONTFACE3 %>" size="<%= C_FONTSIZE3 %>">Error Location:&nbsp;<%= rsTrans.Fields("trnsrspErrorLocation") %></font></td>
	</tr>
	
	</table>
  </tr></td>
<%
	rsTrans.MoveNext  
Loop 
closeObj(rsTrans)
Call cleanup_dbconnopen	'This line needs to be included to close database connection
%>
        <tr>
        <td width="100%" align="center" valign="top" colspan="4"></td>
        </tr>
<%end if%>				        
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

</BODY>
</HTML>









