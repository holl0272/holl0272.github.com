<% Option Explicit 
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version   2.0                                                      *
'*   Release Date      September 21, 2002				                        *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

Const cbytRemoveOldOrders = 30			'default number of days to remove old orders

'/
'/////////////////////////////////////////////////

Sub CleanUp

Dim pstrSQL
Dim DeleteOldOrdersDate
Dim rsOldOrders
Dim plngID
Dim mclsOrders
Dim pblnUpdated
Dim plngCount

	DeleteOldOrdersDate = DateAdd("d", -1 * mvalOldOrders, Now())

	'***********************************************************************************************

	If mblnOldOrders Then	'Delete old Orders
		plngCount = 0
		
		pstrSQL = "SELECT sfOrders.orderID, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssorderID AS ssOrderManagerOrderID" _
				& " FROM ssOrderManager RIGHT JOIN sfOrders ON ssOrderManager.ssorderID = sfOrders.orderID" _
				& " WHERE ((sfOrders.orderIsComplete=1) And (sfOrders.orderDate< " & wrapSQLValue(DeleteOldOrdersDate, True, enDatatype_date) & "))"
		set rsOldOrders = GetRS(pstrSQL)

		Do While Not rsOldOrders.EOF
			pblnUpdated = False
			plngID = rsOldOrders.Fields("orderID").Value

			If Len(rsOldOrders.Fields("ssOrderManagerOrderID").Value & "") = 0 Then
				pstrSQL = "Insert Into ssOrderManager (ssOrderID) Values (" & wrapSQLValue(plngID, False, enDatatype_number) & ")"
				pblnUpdated = True
				'Response.Write "pstrSQL = " & pstrSQL & "<br />"
				cnn.Execute pstrSQL,,128
			End If

			If Len(rsOldOrders.Fields("ssDatePaymentReceived").Value & "") = 0 Then
				pstrSQL = "Update ssOrderManager Set ssDatePaymentReceived=" & wrapSQLValue(rsOldOrders.Fields("orderDate").Value, True, enDatatype_date) _
						& " Where ssOrderID = " & wrapSQLValue(plngID, False, enDatatype_number)
				pblnUpdated = True
				Response.Write "pstrSQL = " & pstrSQL & "<br />"
				cnn.Execute pstrSQL,,128
			End If
			
			If Len(rsOldOrders.Fields("ssDateOrderShipped").Value & "") = 0 Then
				pstrSQL = "Update ssOrderManager Set ssDateOrderShipped=" & wrapSQLValue(rsOldOrders.Fields("orderDate").Value, True, enDatatype_date) _
						& " Where ssOrderID = " & wrapSQLValue(plngID, False, enDatatype_number)
				pblnUpdated = True
				Response.Write "pstrSQL = " & pstrSQL & "<br />"
				cnn.Execute pstrSQL,,128
			End If
				
			If pblnUpdated Then plngCount = plngCount + 1
			rsOldOrders.MoveNext
		Loop
		rsOldOrders.Close
		Set rsOldOrders = Nothing
		
		mstrMessage = mstrMessage & "<H3>" & plngCount & " Orders Updated</H3>"

	End If	'Delete old Orders

End Sub
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Dim mstrMessage
Dim mvalOldOrders
Dim mblnOldOrders

	mblnOldOrders = (Len(Request.Form("chkOldOrders")) > 0)
	If mblnOldOrders Then 
		mvalOldOrders = Request.Form("valOldOrders")
	Else
		mvalOldOrders = cbytRemoveOldOrders
	End If
	
	If (Len(Request.Form("Action")) > 0) Then 
		Call CleanUp
	End If

Call WriteHeader("",True)
%>
<SCRIPT LANGUAGE=javascript>
<!--

function ValidInput(theForm)
{
  if (!isNumeric(theForm.valOldOrders,false,"Please enter an number for the Incomplete Order days.")) {return(false);}
  
  return(true);
}

//-->
</SCRIPT>

<BODY>
<CENTER>
<%= mstrMessage %>
<FORM action='ssOrderAdmin_UpdatePastOrders.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<input type=hidden name="Action" value="Clean">
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
      <TR class='tblhdr'>
        <TH>Order Manager Prior Order Clean-Up Page</TH>
      </TR>
      <tr>
        <td><input type=checkbox name=chkOldOrders id="chkOldOrders">&nbsp;Mark orders paid for orders greater than&nbsp;<input name="valOldOrders" id="valOldOrders" value='<%= mvalOldOrders %>' size=3>&nbsp;days old</td>
      </tr>
  <TR>
    <TD align=center>
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset>&nbsp;&nbsp;
        <INPUT class='butn' id=btnUpdate name=btnUpdate type=submit value='Clean Database'>
    </TD>
  </TR>
</TABLE>
</FORM>

</CENTER>
</BODY>
</HTML>
<% 
	Set cnn=nothing
	Response.Flush
%>