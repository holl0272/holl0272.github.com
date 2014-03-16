<%
option explicit
Response.Buffer = True
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.1003.0.1

'@FILENAME: sfreports1.asp
	
'Access Version
'

'@DESCRIPTION:   Sales details reporting 

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
<SCRIPT language="javascript">
function helpMe(){
	var helpWin, loadHelp
	helpWin = window.open('help/daily_sm2a.htm','helpWin', 'scrollbars=1,resizable,location=0,status=0,toolbar=0,menubar=0,height=300,width=500')
	helpWin.focus()
}	
function hideCredit(num) {
	if (document.frmHideCredit.chkHideCredit.checked) {
		document.frmHideCredit.txtCreditNumber.value = "****-****-****-" + num.slice(num.length-4, num.length)
	}
	else {
		document.frmHideCredit.txtCreditNumber.value = num
	}
}
</SCRIPT>
<!--#include file="../SFLib/incDesign_settings.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->
<!--#include file="../SFLib/incCC.asp"-->
<!--#include file="../SFLIB/incAE.asp"-->
<%
Dim sStartDate, sEndDate, sSQL, rsOrders, iCounter, sBgColor, sFontFace, sFontColor, sFontSize, sOrderID, bAddress
Dim rsOrderDetail, rsOrderShipDetail, rsOrderCredit, rsOrderProducts, rsOrderProdAtt, ccObj,iError,iTempDisc,rsTrans,rsAdmin,sTransMethod
Dim dblGW
Function getCardName(id)
	Dim rs,sSQL
	Set rs = CreateObject("ADODB.Recordset")
	sSQL = "SELECT transName FROM sfTransactionTypes WHERE transID = " & id
	rs.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
	getCardName = rs("transName")
	closeObj(rs)
End Function


sStartDate = MakeUSDate(Request.QueryString("startDate"))
sEndDate = MakeUSDate(Request.QueryString("endDate"))
sOrderID = Request.QueryString("OrderID")

If (sStartDate <> "" And sEndDate <> "") Then
	sSQL = "Select custID, custFirstName, custLastName, custMiddleInitial, orderID, orderCustID, orderDate, orderGrandTotal, orderPaymentMethod " _
		   & "FROM sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId " _
		   & " WHERE orderDate BETWEEN " & wrapSQLValue(sStartDate, True, enDatatype_date) & " AND " & wrapSQLValue(sEndDate, True, enDatatype_date) & " and sfOrders.orderIsComplete = 1 ORDER BY orderID"
	
	Set rsOrders = CreateObject("ADODB.RecordSet")
	rsOrders.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
%>
	<html>
	<head>
	<title>SF Reports Page</title>
	</head>

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
		<td align="middle" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">Sales Details</font></b></td>        
	    </tr>
	    <tr>
		<td bgcolor="<%= C_BGCOLOR3 %>" background="<%= C_BKGRND3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" size="<%= C_FONTSIZE3 %>">All of the orders placed within the date range you specified are listed below.  Choose <b>Transaction Details</b> to view the transaction service report for the selected record, or choose <b>Order Details</b> to view a full sales report for the selected record.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a></td>    
	    </tr>
	    <tr>
	    <td bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>" width="100%">
	        <table border="0" width="100%" cellpadding="4" cellspacing="0">
	        <%
	        If rsOrders.EOF Then
	        %>
				<tr>
				<td colspan=4 align="center" bgcolor="<%= C_BGCOLOR4 %>" background="<%= C_BKGRND4 %>"><font face="<%= C_FONTFACE5 %>" color="#ff0000" size="<%= C_FONTSIZE5+1 %>">There Were No Orders Between <%= sStartDate %> And <%= sEndDate %></font></td>
				</tr>
	        <%
	        Else%>
	        <tr>
	        <td width="5%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Order ID</font></b></td> 
	        <td width="10%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Order Details</font></b></td>        
			<td width="20%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Transaction Details</font></b></td>
			<td width="15%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Date</font></b></td>        
			<td width="30%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Customer Name</font></b></td>        
			<td width="20%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Order Total</font></b></td>  
			<td width="20%" align="center" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5 %>">Delete</font></td>      
	        </tr>
				<%iCounter = 1
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
					

	sSQL = "SELECT trnsrspID, trnsrspOrderId, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode, trnsrspRetrievalCode, trnsrspAuthNo, trnsrspErrorMsg, trnsrspErrorLocation, trnsrspSuccess " _ 
		& "FROM sfTransactionResponse WHERE trnsrspOrderId = " &  trim(rsOrders.Fields("orderID")) 
	Set rsTrans = cnn.execute(sSql)
					
	        %>
					<tr>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= rsOrders.Fields("orderID") %></font></td>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><a href="sfReports1.asp?OrderID=<%= rsOrders.Fields("orderID") %>">Order Details</a></font></td>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%If NOT rsTrans.EOF AND NOT rsTrans.BOF Then %><a href="sfReports3.asp?OrderID=<%= trim(rsOrders.Fields("orderID")) %>"><%= trim(rsOrders.Fields("orderPaymentMethod"))  %> Transaction</a></font></td><%Else%><%= trim(rsOrders.Fields("orderPaymentMethod")) %> Transaction</font></td>					<%End If%>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= trim(rsOrders.Fields("orderDate"))%></font></td>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= trim(rsOrders.Fields("custFirstName")) %>&nbsp;<%= trim(rsOrders.Fields("custMiddleInitial")) %>&nbsp;<%= trim(rsOrders.Fields("custLastName")) %></font></td>
					<td align="center" valign="top" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><%= FormatCurrency(rsOrders.Fields("orderGrandTotal")) %></font></td>
					<td valign="top" align="center" bgcolor="<%= sBgColor %>"><font face="<%= sFontFace %>" color="<%= sFontColor %>" SIZE="<%= sFontSize %>"><a href="sfreports1.asp?delete=1&OrderID=<%=rsOrders.Fields("orderID")%>&remove=1"><img src="../<%= C_BTN06 %>" border="0"></font></td>
	
					</tr>
	        <%
					iCounter = iCounter + 1
					rsOrders.MoveNext 
				Loop
			End If
	        %>
	        <tr>
	        <td width="100%" align="center" valign="top" colspan="6"></td>
	        </tr>
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
	</body>
	<%
	rsOrders.Close 
	Set rsOrders = nothing
	cnn.Close
	Set cnn = nothing
	%>
	</html>
<%

ElseIf Trim(Request.QueryString("DeleteOrd")) = "1" Then 

   Dim rsDelete 
   Dim rsDelete1 
   Dim rsDelete2 
   Dim rsDelete3 
   Dim vOrderID
Set rsDelete = CreateObject("ADODB.RecordSet")
Set rsDelete1 = CreateObject("ADODB.RecordSet")
Set rsDelete2 = CreateObject("ADODB.RecordSet")
Set rsDelete3 = CreateObject("ADODB.RecordSet")
Set rsOrderCredit = CreateObject("ADODB.RecordSet")

sOrderID =request.querystring("orderID")

   sSQL = "SELECT * FROM sfOrders" _
        & " WHERE orderID = " & sOrderId & " and orderIsComplete = 1"
   sfReports1_SQL1 'sfae
   
rsDelete.Open sSQL, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
vOrderId = rsDelete.Fields("orderAddrId")

sSQL = "SELECT * FROM sfOrderDetails WHERE odrdtOrderId = " & sOrderId
	sfReports1_SQL2 'SFAE
	
    rsDelete2.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText

sSQL = "SELECT * FrOM sfCPayments WHERE payID = " & Trim(rsDelete.Fields("orderPayId"))
  rsDelete3.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText


if rsDelete.BOF =false and rsDelete.EOF=false then
 rsDelete.Delete 
end if
if rsDelete2.BOF = false and rsDelete2.EOF = false then
 rsDelete2.Delete 
end if
if rsDelete3.BOF =false and rsDelete3.EOF=false then
  rsDelete3.Delete
end if

Set rsDelete = Nothing
Set rsDelete2 = Nothing
Set rsDelete3 = Nothing
%>
	<html>
	<head>
	<title>SF Reports Page</title>
	</head>

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
		<td><center><font color=white size=4><b>Order Number <%= sOrderID %> has been successfully deleted!</b></font></center></td>
		</tr>
		<tr>
		<td bgcolor="<%= C_BGCOLOR7 %>" background="<%= C_BKGRND7 %>"><font face="<%= C_FONTFACE7 %>" color="<%= C_FONTCOLOR7 %>" size="<%= C_FONTSIZE7 %>"><p align="center"><b><a href="ssAdmin/admin.asp">Site Administration</a> | <a href="sfReports.asp">Reports</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>
		</table>
 <%
Else
	
	Set rsOrderDetail = CreateObject("ADODB.RecordSet")
	sSQL = "SELECT * FROM sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId " _
		& "WHERE orderID = " & sOrderID
	rsOrderDetail.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	Set rsOrderShipDetail = CreateObject("ADODB.RecordSet")
	sSQL = "SELECT * FROM sfCShipAddresses WHERE cshpaddrID = " & rsOrderDetail.Fields("orderAddrId")
	rsOrderShipDetail.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText 
	If rsOrderDetail.Fields("orderAddrId") > 0 Then bAddress = 1
	
	Set rsOrderProducts = CreateObject("ADODB.RecordSet")
	sSQL = "SELECT * FROM sfOrderDetails WHERE odrdtOrderId = " & rsOrderDetail.Fields("orderID")
	sfReports1_SQL3 'SFAE
	'Response.Write Application("AppName") 
	'Response.End 
	rsOrderProducts.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	Set rsOrderProdAtt = CreateObject("ADODB.RecordSet")
	
	Set rsOrderCredit = CreateObject("ADODB.RecordSet")
	sSQL = "SELECT payCardType, payCardName, payCardNumber, payCardExpires FROM sfCPayments WHERE payID = " & Trim(rsOrderDetail.Fields("orderPayId"))

	rsOrderCredit.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	Set rsAdmin = CreateObject("ADODB.Recordset")
	rsAdmin.Open "SELECT adminTransMethod FROM sfAdmin", cnn, adOpenStatic, adLockReadOnly, adCmdText
	sTransMethod = trim(rsAdmin.Fields("adminTransMethod"))
	rsAdmin.Close
	Set rsAdmin = Nothing

	
	Err.number = 0
	On Error Resume Next
	Set ccObj = CreateObject("SFServer505.CCEncrypt")
	ccObj.putSeed(iCC)
	iError = Err.number 
	On Error GoTo 0
%>
<html>
<head>
</head>
<body bgColor="#ffffff">
<table bgColor="#ffffff" width="70%" border=0 cellspacing=0 cellpadding=5>
<form name="frmHideCredit">
<input type="hidden" name=hdnHideCredit" value="0">
<tr>
<td align="left"><b>Order Id:&nbsp;<%= rsOrderDetail.Fields("orderID") %></b></td>
<td align="right"><b>Order Date:&nbsp;<%= rsOrderDetail.Fields("orderDate")%></b></td>
</tr>
	<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
<tr>
<td align="left"><b>Sold To</b></td>
<td align="left"><b>Ship To</b></td>
</tr>
	<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
<tr>
<td align="right" width="50%">
	<table align="left" valign="top" width="100%">
	<tr>
	<td align="left" nowrap>First Name/MI:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custFirstName")%>&nbsp;<%= rsOrderDetail.Fields("custMiddleInitial")%></td>
	</tr>
	<tr>
	<td align="left">Last Name:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custLastName")%></td>
	</tr>
	<tr>
	<td align="left">Company:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custCompany")%></td>
	</tr>
	<tr>
	<td align="left" nowrap>Address Line 1:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custAddr1")%></td>
	</tr>
	<tr>
	<td align="left" nowrap>Address Line 2:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custAddr2")%></td>
	</tr>
	<tr>
	<td align="left">City:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custCity")%></td>
	</tr>
	<tr>
	<td align="left">State:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custState")%></td>
	</tr>
	<tr>
	<td align="left">Country:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custCountry")%></td>
	</tr>
	<tr>
	<td align="left">Zip or Postal Code:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custZip")%></td>
	</tr>
	<tr>
	<td align="left">Phone Number:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custPhone")%></td>
	</tr>
	<tr>
	<td align="left">Fax Number:</td>
	<td align="left">&nbsp;<%= rsOrderDetail.Fields("custFax")%></td>
	</tr>
	<tr>
	<td align="left" nowrap>Email Address:</td>
	<td align="left">&nbsp;<a href="mailto:<%= rsOrderDetail.Fields("custEmail")%>"><%= rsOrderDetail.Fields("custEmail") %></a></td>
	</tr>
	</table>
</td>
<td align="right" width="50%">
<% If bAddress = 1 Then %>
	<table align="left" valign="top" width="100%">
	<tr>
	<td align="left">First Name/MI:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipFirstName")%>&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipMiddleInitial")%></td>
	</tr>
	<tr>
	<td align="left">Last Name:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipLastName")%></td>
	</tr>
	<tr>
	<td align="left">Company:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipCompany")%></td>
	</tr>
	<tr>
	<td align="left">Address Line 1:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipAddr1")%></td>
	</tr>
	<tr>
	<td align="left">Address Line 2:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipAddr2")%></td>
	</tr>
	<tr>
	<td align="left">City:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipCity")%></td>
	</tr>
	<tr>
	<td align="left">State:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipState")%></td>
	</tr>
	<tr>
	<td align="left">Country:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipCountry")%></td>
	</tr>
	<tr>
	<td align="left">Zip or Postal Code:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipZip")%></td>
	</tr>
	<tr>
	<td align="left">Phone Number:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipPhone")%></td>
	</tr>
	<tr>
	<td align="left">Fax Number:</td>
	<td align="left">&nbsp;<%= rsOrderShipDetail.Fields("cshpaddrShipFax")%></td>
	</tr>
	<tr>
	<td align="left">Email Address:</td>
	<td align="left">&nbsp;<a href="mailto:<%= rsOrderShipDetail.Fields("cshpaddrShipEmail")%>"><%= rsOrderShipDetail.Fields("cshpaddrShipEmail") %></a></td>
	</tr>
	</table>
<% End If %>	
</td>
</tr>
<tr>
	<td>Special Instructions:</td>
	<td>&nbsp;</td>
	</tr>
	<tr>
	<td colspan="2"><%= rsOrderDetail.Fields("orderComments") %></td>
	</tr>
	<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
<tr>
<td colspan=2><b>Refferal Tracking Information</b></td>
</tr>
<tr>
<td colspan=2>

	<table align="left" valign="top" width="100%" cellpadding="0" cellspacing="0">
	<tr><td colspan="4"><hr width="100%" size="2"></td></tr>
	<tr>
	<td width=25% align="left">Trading Partner:&nbsp;</td>
	<td width=25% align="left"><%= rsOrderDetail.Fields("orderTradingPartner") %>&nbsp;</td>
	<td width=25% align="left">Remote Address:&nbsp;</td>
	<td width=25% align="left"><%= rsOrderDetail.Fields("orderRemoteAddress") %>&nbsp;</td>
	</tr>
	<tr>
	<td valign="top" align="left">Http Refferer:&nbsp;</td>
	<td colspan=3><%= mid(rsOrderDetail.Fields("orderHttpReferrer"),1, 50) %>&nbsp;</td>
	</tr>
	</table>
</td>
</tr>
	<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
<tr>
<td colspan=2><b>Payment Method: <%= Trim(rsOrderDetail.Fields("orderPaymentMethod")) %></b>
<% If Trim(rsOrderDetail.Fields("orderPaymentMethod")) = cstrECheckTerm Then %><a href="printEcheck.asp?orderid=<%= trim(rsOrderDetail.Fields("orderID")) %>">Print Check</a><%End If%>
</td>
</tr>
<%
	If (Trim(rsOrderDetail.Fields("orderPaymentMethod")) = "Credit Card" or Trim(rsOrderDetail.Fields("orderPaymentMethod")) = cstrPhoneFaxTerm) AND sTransMethod <> "15" AND sTransMethod <> "18" Then
%>	
	<tr>
	<td>


		<table>
		<tr>
		<td align="left">Name on Credit Card:</td>
		<td align="left">&nbsp;<%= rsOrderCredit.Fields("payCardName")%>
		</td>
		</tr>
		<tr>
		<td align="left">Credit Card Type:</td>
		<td align="left">&nbsp;<%= getCardName(rsOrderCredit.Fields("payCardType"))%></td>
		</tr>
		</table>
	</td>
	<td>
		<table>
		<tr>
		<td align="left">Credit Card Number:</td>
		<td align="left" nowrap>&nbsp;<input type="text" name="txtCreditNumber" value="<%If iError = 0 Then%><%= ccObj.decrypt(rsOrderCredit.Fields("payCardNumber"))%><%Else%><%=rsOrderCredit.Fields("payCardNumber")%><%End If%>" readonly size="20">&nbsp;<input type="checkbox" name="chkHideCredit" value="1" onClick="hideCredit('<%If iError = 0 Then%><%= ccObj.decrypt(rsOrderCredit.Fields("payCardNumber"))%><%Else%><%=rsOrderCredit.Fields("payCardNumber")%><%End If%>')">Mask Card #</td>
		<%closeObj(ccObj)%>
		</tr>
		<tr>
		<td align="left">Expiration Date:</td>
		<td align="left">&nbsp;<%= rsOrderCredit.Fields("payCardExpires")%></td>
		</tr>
		</table>
	</tr>
<%
		rsOrderCredit.Close 
		Set rsOrderCredit = nothing
	ElseIf rsOrderDetail.Fields("orderPaymentMethod") = cstrECheckTerm  or Trim(rsOrderDetail.Fields("orderPaymentMethod")) = cstrPhoneFaxTerm Then
%>
	<tr>
	<td>
		<table>
		<tr>
		<td align="right">Checking Account Number:</td>
		<td align="left">&nbsp;<%= rsOrderDetail.Fields("orderCheckAcctNumber")%></td>
		</tr>
		<tr>
		<td align="right">Check Number:</td>
		<td align="left">&nbsp;<%= rsOrderDetail.Fields("orderCheckNumber")%></td>
		</tr>
		</table>
	</td>
	<td>
		<table>
		<tr>
		<td align="right">Bank Name:</td>
		<td align="left">&nbsp;<%= rsOrderDetail.Fields("orderBankName")%></td>
		</tr>
		<tr>
		<td align="right">Routing Number:</td>
		<td align="left">&nbsp;<%= rsOrderDetail.Fields("orderRoutingNumber")%></td>
		</tr>
		</table>
	</tr>
	<tr>
	<td colspan=2>
	</td>
	</tr>
<%	ElseIf rsOrderDetail.Fields("orderPaymentMethod") = cstrPOTerm  or Trim(rsOrderDetail.Fields("orderPaymentMethod")) = cstrPhoneFaxTerm Then %>
	<tr>
	<td>
		<table>
		<tr>
		<td align="right">Purchase Order Name:</td>
		<td align="left">&nbsp;<%= rsOrderDetail.Fields("orderPurchaseOrderName")%></td>
		</tr>
		</table>
	</td>
	<td>
		<table>
		<tr>
		<td align="right">Purchase Order Number:</td>
		<td align="left">&nbsp;<%= rsOrderDetail.Fields("orderPurchaseOrderNumber")%></td>
		</tr>
		</table>
	</td>
	</tr>
<%	ElseIf rsOrderDetail.Fields("orderPaymentMethod") = "InternetCash" Then %>

<%
	End If
%>
<tr>
<td colspan=2>
	<table width="100%" border=1 cellpadding=5 cellspacing=0>
	<tr>
	<td><center><b>Product ID</b></center></td>
	<td><center><b>Product Name</b></center></td>
	<td><center><b>Quantity</b></center></td>
	<td><center><b>Product Price</b></center></td>
	<td><center><b>Product Total</b></center></td>
	</tr>
	
<% 
iTempDisc = 0
Do While Not rsOrderProducts.EOF 
%>
	<tr>
	<td valign="top"><%= rsOrderProducts.Fields("odrdtProductID") %>&nbsp;</td>
	<td valign="top"><%= rsOrderProducts.Fields("odrdtProductName") %><br>
		<table>
			<% 	
			   sSQL = "SELECT * FROM sfOrderAttributes WHERE odrattrOrderDetailId = " & rsOrderProducts.Fields("odrdtID")
			   rsOrderProdAtt.Open sSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
			   Do While Not rsOrderProdAtt.EOF
		    %>
					<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;<%= rsOrderProdAtt.Fields("odrattrName") %>: <%= rsOrderProdAtt.Fields("odrattrAttribute") %></td></tr>
			<%		rsOrderProdAtt.MoveNext 
			   Loop
			   rsOrderProdAtt.Close 
			%>
		</table>	
	</td>
	<td valign="top" align="right"><%= rsOrderProducts.Fields("odrdtQuantity") %>&nbsp;</td>
	<td valign="top" align="right"><%= FormatCurrency(cDbl(rsOrderProducts.Fields("odrdtSubTotal"))/cDbl(rsOrderProducts.Fields("odrdtQuantity"))) %>&nbsp;</td>
	<td valign="top" align="right"><%= FormatCurrency(rsOrderProducts.Fields("odrdtSubTotal")) %>&nbsp;</td>
	
	</tr>	
	<% If cblnSF5AE Then sfReports1_ShowProductDetails 'SFAE %>
<%	
	iTempDisc = iTempDisc + cDbl(rsOrderProducts.Fields("odrdtSubTotal"))
        If Application("AppName")    ="StoreFrontAE" 	then
		   If isNumeric(rsOrderProducts.Fields("odrdtGiftWrapPrice")) Then
			dblGW =  cDbl(rsOrderProducts.Fields("odrdtGiftWrapPrice"))
		   Else
			dblGW =  0
		   End If
		   iTempDisc = iTempDisc + dblGW

		End if
	rsOrderProducts.MoveNext 
  Loop 
%>
	</table>
</td>
</tr>
	<tr>
	<td colspan=2 align="right">
		<table>
			
		<%sfReports1_Coupon 'SFAE  'iTempDisc is adjusted in here%>
		<%
		iTempDisc = iTempDisc - cDbl(rsOrderDetail.Fields("orderAmount"))
		If iTempDisc <> 0 Then
		%>
	
		<tr>
		<td>Storewide Discount:&nbsp;&nbsp;&nbsp;</td>
		<td align="right">- <%= FormatCurrency(iTempDisc) %>&nbsp;&nbsp;</td>
		</tr>
		<% End If %>
		<tr>
		<td>Sub Total:&nbsp;&nbsp;&nbsp;</td>
		<td align="right"><%= FormatCurrency(rsOrderDetail.Fields("orderAmount")) %>&nbsp;&nbsp;</td>
		</tr>
		<tr>
		<td>State/Province Tax:&nbsp;&nbsp;&nbsp;</td>
		<td align="right"><%If rsOrderDetail.Fields("orderSTax") <> "" Then Response.Write FormatCurrency(rsOrderDetail.Fields("orderSTax")) %>&nbsp;&nbsp;</td>
		</tr>
		<tr>
		<td>Country Tax:&nbsp;&nbsp;&nbsp;</td>
		<td align="right"><%If rsOrderDetail.Fields("orderCTax") <> "" Then Response.Write FormatCurrency(rsOrderDetail.Fields("orderCTax")) %>&nbsp;&nbsp;</td>
		</tr>
		<tr>
		<%
		if rsOrderDetail.Fields("orderShipMethod") ="Free" then%>
		 <td>Free Shipping:&nbsp;&nbsp;&nbsp;</td>
		<%
		Elseif rsOrderDetail.Fields("orderShippingAmount")=0 then
		%>
		 <td>Shipping:&nbsp;&nbsp;&nbsp;</td>
		<%
		else
    	
    	%>
    	
		 <td><%= rsOrderDetail.Fields("orderShipMethod") %>:&nbsp;&nbsp;&nbsp;</td>
		<%
		end if
		%>
		<td align="right"><%If rsOrderDetail.Fields("orderShippingAmount") <> "" Then Response.Write FormatCurrency(rsOrderDetail.Fields("orderShippingAmount")) %>&nbsp;&nbsp;</td>
		</tr>
		<tr>
		<%
		if rsOrderDetail.Fields("orderPaymentMethod")= cstrCODTerm then
		%>
		<td>Handling/COD:&nbsp;&nbsp;&nbsp;</td>
		<%
		else
		%>
		<td>Handling:&nbsp;&nbsp;&nbsp;</td>
		<%
		end if
		%>
		
		<td align="right"><%If rsOrderDetail.Fields("orderHandling") <> "" Then	 Response.Write FormatCurrency(rsOrderDetail.Fields("orderHandling")) %>&nbsp;&nbsp;</td>
		</tr>
			<tr><td colspan="2"><hr width="100%" size="2"></td></tr>
		<tr>
		<td><b>Grand Total:&nbsp;&nbsp;&nbsp;</b></td>
		<td align="right"><%If rsOrderDetail.Fields("orderGrandTotal") <> "" Then Response.Write FormatCurrency(rsOrderDetail.Fields("orderGrandTotal")) %>&nbsp;&nbsp;</td>
		
		</tr>
		
		<%sfReports1_Billing'SFAE%>
		</table>
	</td>
	</tr>
	</table>
</td>
</tr>
</form>
<% If request.querystring("remove") = 1 Then %>
<tr>
<td colspan="2" align="right"><align="right"><a href="sfreports1.asp?deleteOrd=1&OrderID=<%=rsOrderDetail.Fields("orderID")%>"><img src="../<%= C_BTN06 %>" border="0"></a></td>
</tr>
<% End If %>
</table>
</body>

<%
	rsOrderProducts.Close 
	rsOrderDetail.Close 
	If bAddress = 1 Then
		rsOrderShipDetail.Close
	End If 
	Set rsOrderProdAtt = nothing 
	Set rsOrderDetail = nothing
	Set rsOrderShipDetail = nothing
	Set rsOrderProducts = nothing
End If
function SplitString(vData)
vdata = replace(vdata,"&","& ")
vdata = replace(vdata,"=","= ")
SplitString = vdata
end function


%>



