<%
Option Explicit
Response.Buffer = True 
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.1003.0.1

'@FILENAME: promo_mail.asp
	 
'Access Version

'@DESCRIPTION:   sends promotional mail

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
<!--#include file="../SfLib/sfsecurity.asp"-->
<!--#include file="../SFLib/incDesign.asp"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/mail.asp" -->
<script language=javascript>

function helpMe(){
	var helpWin, loadHelp
	helpWin = window.open('help/daily_sm4.htm','helpWin', 'scrollbars=1,resizable,location=0,status=0,toolbar=0,menubar=0,height=300,width=500')
	helpWin.focus()
}	

function checkPromoMail(form) {
	if ((form.startDate.value.length == 0) && (form.endDate.value.length == 0) && (form.product.value.length == 0)) {
		alert("Please enter search criteria.")
		form.startDate.focus();
		return false;
	}
	if ((form.startDate.value.length == 0) && (form.endDate.value.length != 0)) {
		alert("Please enter a starting date.");
		form.startDate.focus();
		return false;
	}
	if ((form.startDate.value.length != 0) && (form.endDate.value.length == 0)) {
		alert("Please enter an ending date.");
		form.endDate.focus();
		return false;
	}
	if (form.subject.value.length == 0) {
		alert("Please enter subject.");
		form.subject.focus();
		return false;		
	}
	if (form.message.value.length == 0) {
		alert("Please enter message.");
		form.message.focus();
		return false;		
	}
	return true;	
}
function checkRemove(form) {
	if (form.custEmail.value.length == "") {
		alert("Please enter an Email Address.")
		form.custEmail.focus()
		return false;
	}
	return true;
}
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>SF Store Promotional Mail Utility</title>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body background="<%= C_BKGRND %>" bgproperties="fixed" bgcolor="<%= C_BGCOLOR %>" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">
<table border='0' cellpadding='1' cellspacing='0' bgcolor="<%= C_BORDERCOLOR1 %>" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <%If C_BNRBKGRND = "" Then%>
	      <td colspan=4 align="center" background="<%= C_BKGRND1 %>" bgcolor="<%= C_BGCOLOR1 %>"><b><font face="<%= C_FONTFACE1 %>" color="<%= C_FONTCOLOR1 %>" SIZE="<%= C_FONTSIZE1 %>"><%= C_STORENAME %></font></b></td>
          <%Else%>
	      <td colspan=4 align="center" bgcolor="<%= C_BNRBGCOLOR %>"><img src="<%= C_BNRBKGRND %>" border="0"></td>
          <%End If%>
        </tr>
        <tr>
          <td colspan=4 align="center" background="<%= C_BKGRND2 %>" bgcolor="<%= C_BGCOLOR2 %>"><b><font face="<%= C_FONTFACE2 %>" color="<%= C_FONTCOLOR2 %>" SIZE="<%= C_FONTSIZE2 %>">StoreFront Web Store Promotional Mail Utility</font></b></td>
        </tr>
        <tr>
          <td align="left" background="<%= C_BKGRND3 %>" bgcolor="<%= C_BGCOLOR3 %>"><font face="<%= C_FONTFACE3 %>" color="<%= C_FONTCOLOR3 %>" SIZE="<%= C_FONTSIZE3 %>">Use Promo Mail to send promotional mailings to customers who have elected to subscribe to the web store mailing list.  Promotional emails will be sent to all customers who have 
          subscribed to your mailing list within the date range specified in the Date fields.</font>&nbsp;&nbsp;<a HREF="javascript:helpMe()"><img src="images/help.jpg" alt="Help" border="0"></a></td>
        </tr>
        <% If Request.Form("SendMail.x") = "" And Request.Form("Remove.x") = "" Then %> 
	    <tr>
	      <td width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5+1  %>">Promotional Mailing</font></b></td>
	    </tr>
	    <tr>
	      <td bgcolor="<%= C_BGCOLOR4 %>">
	        <table border="0" cellpadding="0" cellspacing="5" width="100%">
	          <form action=promo_mail.asp method=post id=form1 name=form1 onsubmit="return checkPromoMail(this);">
	            <tr><td><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Start Date:</font></td><td>
                  <input type=text name=startDate SIZE="20"></td><td><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">End Date:</font></td><td>
                  <input type=text name=endDate SIZE="20"></td></tr>
	            <tr><td><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Product ID:</font></td><td colspan=3><input type=text name=product size=20></td></tr>
	            <tr><td colspan=4><hr noshade width="90%"></td></tr>
	            <tr><td><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">Subject:</font></td><td colspan=3><input type=text name=subject size=58</td></tr>
	              <tr><td colspan=4><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><u>Mail Message</u></font></td></tr>
	              <tr><td colspan=4><textarea rows=10 cols=66 name=message></textarea></td></tr>
	              <tr><td colspan=4 align="center"><input type=image name="SendMail" border="0" src="../images/buttons/submit.gif" alt"Send Mail"></td></tr>
	            </form>
	          </table>
	        </td>
	      </tr>
	      <tr>
	        <td width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5+1 %>">Customer Removal</font></b></td>
	      </tr>
	      <tr>
	        <form action=promo_mail.asp method=post id=form1 name=form2 onsubmit="return checkRemove(this);">
	          <td align="center" bgcolor="<%= C_BGCOLOR4 %>">
	            <font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><b>Note:</b> To remove multiple addresses, enter the addresses 
                separated by commas.</font><br>
	            <font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">E-Mail Address(es):</font>&nbsp;<input type=text name=custEmail size=50><input type=image border="0" name="Remove" src="../images/buttons/submit.gif" alt"Remove Email(s)"></td>
	        </form>
	      </tr> 
          <% 
ElseIf Request.Form("SendMail.x") <> "" Then
	'*****************
	'** Run Mailing **
	'*****************
	Dim rsMailer,sInformation,sSubject,sMessage,arrMail,iNum,sName,tTemp,arrMailSent,j,sTemp
	Set rsMailer = Server.CreateObject("ADODB.Recordset")
	
	If Request.Form("product") <> "" and Request.Form("startDate") <> "" And Request.Form("endDate") <> "" Then
		SQL = "SELECT custEmail, custFirstName, custMiddleInitial, custLastName, odrdtProductID FROM sfCustomers INNER JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCustomers.custID = sfOrders.orderCustId " & _
		" WHERE (custIsSubscribed = 1) AND (custLastAccess >=  # " & MakeUSDate(Request("startDate")) & " #) AND (custLastAccess <= # " & MakeUSDate(Request("endDate")) & " #) AND odrdtProductID = '" & Request.Form("product") & "'"
	ElseIf Request.Form("product") <> "" Then
		SQL = "SELECT custEmail, custFirstName, custMiddleInitial, custLastName, odrdtProductID FROM sfCustomers INNER JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCustomers.custID = sfOrders.orderCustId " & _
		" WHERE (custIsSubscribed = 1) AND odrdtProductID = '" & Request.Form("product") & "'"
	Else
		SQL = "SELECT custEmail, custFirstName, custMiddleInitial, custLastName FROM sfCustomers WHERE (custIsSubscribed = 1) AND (custLastAccess >= # " & MakeUSDate(Request("startDate")) & " #) AND (custLastAccess <= # " & MakeUSDate(Request("endDate")) & " #)"
	End If
	rsMailer.CursorLocation = adUseClient
	rsMailer.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	sSubject = Request.Form("subject")
	sMessage = Request.Form("message")
	If Not (rsMailer.BOF And rsMailer.EOF) Then arrMail = rsMailer.GetRows 
	
	If isArray(arrMail) Then
		sTemp=""
		j = 0
		Redim arrMailSent(uBound(arrMail,2))
		For i=0 to uBound(arrMail,2)
			If arrMail(0,i) <> sTemp Then	
				sName = trim(arrMail(1,i)) & " " & trim(arrMail(2,i)) & " " & trim(arrMail(3,i))
				sInformation = arrMail(0,i) & "|" & sSubject & "|" & "Dear " & sName & "," & vbcrlf & vbtab & sMessage
				createMail "PromoMail", sInformation
				arrMailSent(j) = arrMail(0,i)
				j = j + 1
				sTemp = arrMail(0,i)
			End If
		Next
		Redim Preserve arrMailSent(j-1)
	End If 
	%>
	      <tr>
		    <td width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5+1  %>">Promotional Mailing Results</font></b></td>
	      </tr>
	      <tr>
	        <td bgcolor="<%= C_BGCOLOR4 %>">
	          <table border="0" cellpadding="0" cellspacing="5" width="100%">
	            <% If isArray(arrMailSent) Then %>
		        <tr><td align=center colspan=4><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">The promotional mailing was successfully completed.</font></td></tr>
	            <% Else %>
		        <tr><td align=center colspan=4><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>">There are no Subscribed Customers who 
                  Ordered between <%= Request.Form("startDate") %> and <%= Request.Form("endDate") %><% If Request.Form("product") <> "" Then %> for Product ID <%= Request.Form("product") %><%End If%></font></td></tr>
	            <% End If %>
	            <%
	If isArray(arrMailSent) Then
		iNum = uBound(arrMailSent)
		For i=0 to iNum Step 4
			Response.Write "<tr><td><font face=""" &  C_FONTFACE4 & """ color=""" & C_FONTCOLOR4 & """ size=""" & C_FONTSIZE4-1 & """>" & arrMailSent(i) & "</font></td>" 
			If iNum >=i+1 Then 
				Response.Write "<td><font face=""" &  C_FONTFACE4 & """ color=""" & C_FONTCOLOR4 & """ size=""" & C_FONTSIZE4-1 & """>" & arrMailSent(i+1) & "</font></td>"
			Else
				Response.Write "<td><font face=""" &  C_FONTFACE4 & """ color=""" & C_FONTCOLOR4 & """ size=""" & C_FONTSIZE4-1 & """>&nbsp;</font></td>"
			End If
			If iNum >=i+2 Then 
				Response.Write "<td><font face=""" &  C_FONTFACE4 & """ color=""" & C_FONTCOLOR4 & """ size=""" & C_FONTSIZE4-1 & """>" & arrMailSent(i+2) & "</font></td>"
			Else
				Response.Write "<td><font face=""" &  C_FONTFACE4 & """ color=""" & C_FONTCOLOR4 & """ size=""" & C_FONTSIZE4-1 & """>&nbsp;</font></td>"
			End If
			If iNum >=i+3 Then 
				Response.Write "<td><font face=""" &  C_FONTFACE4 & """ color=""" & C_FONTCOLOR4 & """ size=""" & C_FONTSIZE4-1 & """>" & arrMailSent(i+3) & "</font></td></tr>" 
			Else
				Response.Write "<td><font face=""" &  C_FONTFACE4 & """ color=""" & C_FONTCOLOR4 & """ size=""" & C_FONTSIZE4-1 & """>&nbsp;</font></td></tr>"
			End If
		Next 
	End If 
	%>
	            <tr><td align=center colspan=4><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><a href="promo_mail.asp">Back</a></font></td></tr>
	          </table>
	        </td>
	      </tr> 
	      <%
	closeObj(rsMailer)
ElseIf Request.Form("Remove.x") <> "" Then
	'**********************
	'** Remove recipient **
	'**********************
	Dim aAddresses, i, SQL

	aAddresses = Split(Request("custEmail"),",")
	SQL = "UPDATE sfCustomers SET custIsSubscribed = 0 WHERE "
	For i = 0 To UBound(aAddresses)-1
		SQL = SQL & "custEmail = '" & trim(aAddresses(i)) & "' OR "
	Next 
	SQL = SQL & "custEmail = '" & aAddresses(i) & "'"
	response.write SQL 
	cnn.Execute SQL
	%>
	      <tr>
		    <td width="100%" bgcolor="<%= C_BGCOLOR5 %>" background="<%= C_BKGRND5 %>"><b><font face="<%= C_FONTFACE5 %>" color="<%= C_FONTCOLOR5 %>" size="<%= C_FONTSIZE5+1  %>">Customer Removal</font></b></td>
	      </tr>
	      <tr>
	        <td bgcolor="<%= C_BGCOLOR4 %>">
	          <table border="0" cellpadding="0" cellspacing="5" width="100%">
	            <tr><td align=center colspan=4><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><%= Request("custEmail") %> was successfully removed from the list.</font></td></tr>
	            <tr><td align=center colspan=4><font face="<%= C_FONTFACE4 %>" color="<%= C_FONTCOLOR4 %>" size="<%= C_FONTSIZE4 %>"><a href="promo_mail.asp">Back</a></font></td></tr>
	          </table>
	        </td>
	      </tr>
          <% 
End If 
closeObj(cnn)
%>
	    <tr>
			    <td bgcolor="<%= C_BGCOLOR7 %>" background="<%= C_BKGRND7 %>"><font face="<%= C_FONTFACE7 %>" color="<%= C_FONTCOLOR7 %>" size="<%= C_FONTSIZE7 %>"><p align="center"><b><a href="MT_MenuSales.asp">Sales and Promotions</a> | <a href="Menu.asp">Merchant Tools Home</a> | <a href="../../search.asp"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>
</table>
</table>
    </body>



