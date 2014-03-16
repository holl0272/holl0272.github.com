<%Option Explicit
'********************************************************************************
'*   PayPal Payments															*
'*   Release Version:   3.00													*
'*   Release Date:		March 17, 2003											*
'*   Revision Date:		March 17, 2003											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clsPayPalPayments
'Assumptions:
'   pobjConn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private pbytActionCompleted_Completed
Private pbytActionCompleted_Denied
Private pbytActionCompleted_Failed
Private pbytActionCompleted_Pending
Private pstraddress_city
Private pstraddress_country
Private pstraddress_state
Private pstraddress_status
Private pstraddress_street
Private pstraddress_zip
Private pbytCategory
Private pstrcustom
Private pstrfirst_name
Private pstrinvoice
Private pstritem_name
Private pstritem_number
Private pstrlast_name
Private pstrnotify_version
Private pstrpayer_email
Private pstrpayer_status
Private pdtpayment_date
Private pdblpayment_fee
Private pdblpayment_gross
Private pstrpayment_status
Private pstrpayment_type
Private pstrpending_reason
Private pdblquantity
Private pstrreceiver_email
Private pstrtxn_id
Private pstrtxn_type
Private pstrverify_sign

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
End Sub

Private Sub class_Terminate()
    On Error Resume Next
    Set pRS = Nothing
End Sub

'***********************************************************************************************

Public Property Let Recordset(oRS)
    set pRS = oRS
End Property

Public Property Get Recordset()
    set Recordset = pRS
End Property


Public Property Get Message()
    Message = pstrMessage
End Property

Public Sub OutputMessage()

Dim i
Dim aError

    aError = Split(pstrMessage, cstrDelimeter)
    For i = 0 To UBound(aError)
        If pblnError Then
            Response.Write "<P align='center'><H4><FONT color=Red>" & aError(i) & "</FONT></H4></P>"
        Else
            Response.Write "<P align='center'><H4>" & aError(i) & "</H4></P>"
        End If
    Next 'i

End Sub 'OutputMessage


Public Property Get ActionCompleted_Completed()
    ActionCompleted_Completed = pbytActionCompleted_Completed
End Property

Public Property Get ActionCompleted_Denied()
    ActionCompleted_Denied = pbytActionCompleted_Denied
End Property

Public Property Get ActionCompleted_Failed()
    ActionCompleted_Failed = pbytActionCompleted_Failed
End Property

Public Property Get ActionCompleted_Pending()
    ActionCompleted_Pending = pbytActionCompleted_Pending
End Property

Public Property Get address_city()
    address_city = pstraddress_city
End Property

Public Property Get address_country()
    address_country = pstraddress_country
End Property

Public Property Get address_state()
    address_state = pstraddress_state
End Property

Public Property Get address_status()
    address_status = pstraddress_status
End Property

Public Property Get address_street()
    address_street = pstraddress_street
End Property

Public Property Get address_zip()
    address_zip = pstraddress_zip
End Property

Public Property Get Category()
    Category = pbytCategory
End Property

Public Property Get custom()
    custom = pstrcustom
End Property

Public Property Get first_name()
    first_name = pstrfirst_name
End Property

Public Property Get invoice()
    invoice = pstrinvoice
End Property

Public Property Get item_name()
    item_name = pstritem_name
End Property

Public Property Get item_number()
    item_number = pstritem_number
End Property

Public Property Get last_name()
    last_name = pstrlast_name
End Property

Public Property Get notify_version()
    notify_version = pstrnotify_version
End Property

Public Property Get payer_email()
    payer_email = pstrpayer_email
End Property

Public Property Get payer_status()
    payer_status = pstrpayer_status
End Property

Public Property Get payment_date()
    payment_date = pdtpayment_date
End Property

Public Property Get payment_fee()
    payment_fee = pdblpayment_fee
End Property

Public Property Get payment_gross()
    payment_gross = pdblpayment_gross
End Property

Public Property Get payment_status()
    payment_status = pstrpayment_status
End Property

Public Property Get payment_type()
    payment_type = pstrpayment_type
End Property

Public Property Get pending_reason()
    pending_reason = pstrpending_reason
End Property

Public Property Get quantity()
    quantity = pdblquantity
End Property

Public Property Get receiver_email()
    receiver_email = pstrreceiver_email
End Property

Public Property Get txn_id()
    txn_id = pstrtxn_id
End Property

Public Property Get txn_type()
    txn_type = pstrtxn_type
End Property

Public Property Get verify_sign()
    verify_sign = pstrverify_sign
End Property

'***********************************************************************************************

Private Sub LoadValues(objRS)

    pbytActionCompleted_Completed = trim(objRS.Fields("ActionCompleted_Completed").Value)
    pbytActionCompleted_Denied = trim(objRS.Fields("ActionCompleted_Denied").Value)
    pbytActionCompleted_Failed = trim(objRS.Fields("ActionCompleted_Failed").Value)
    pbytActionCompleted_Pending = trim(objRS.Fields("ActionCompleted_Pending").Value)
    pstraddress_city = trim(objRS.Fields("address_city").Value)
    pstraddress_country = trim(objRS.Fields("address_country").Value)
    pstraddress_state = trim(objRS.Fields("address_state").Value)
    pstraddress_status = trim(objRS.Fields("address_status").Value)
    pstraddress_street = trim(objRS.Fields("address_street").Value)
    pstraddress_zip = trim(objRS.Fields("address_zip").Value)
    pbytCategory = trim(objRS.Fields("Category").Value)
    pstrcustom = trim(objRS.Fields("custom").Value)
    pstrfirst_name = trim(objRS.Fields("first_name").Value)
    pstrinvoice = trim(objRS.Fields("invoice").Value)
    pstritem_name = trim(objRS.Fields("item_name").Value)
    pstritem_number = trim(objRS.Fields("item_number").Value)
    pstrlast_name = trim(objRS.Fields("last_name").Value)
    pstrnotify_version = trim(objRS.Fields("notify_version").Value)
    pstrpayer_email = trim(objRS.Fields("payer_email").Value)
    pstrpayer_status = trim(objRS.Fields("payer_status").Value)
    pdtpayment_date = trim(objRS.Fields("payment_date").Value)
    pdblpayment_fee = trim(objRS.Fields("payment_fee").Value)
    pdblpayment_gross = trim(objRS.Fields("payment_gross").Value)
    pstrpayment_status = trim(objRS.Fields("PayPalPayments.payment_status").Value)
    pstrpayment_type = trim(objRS.Fields("payment_type").Value)
    pstrpending_reason = trim(objRS.Fields("PayPalPayments.pending_reason").Value)
    pdblquantity = trim(objRS.Fields("quantity").Value)
    pstrreceiver_email = trim(objRS.Fields("receiver_email").Value)
    pstrtxn_id = trim(objRS.Fields("txn_id").Value)
    pstrtxn_type = trim(objRS.Fields("txn_type").Value)
    pstrverify_sign = trim(objRS.Fields("verify_sign").Value)

End Sub 'LoadValues

Private Sub LoadFromRequest

    With Request.Form
        pbytActionCompleted_Completed = Trim(.Item("ActionCompleted_Completed"))
        pbytActionCompleted_Denied = Trim(.Item("ActionCompleted_Denied"))
        pbytActionCompleted_Failed = Trim(.Item("ActionCompleted_Failed"))
        pbytActionCompleted_Pending = Trim(.Item("ActionCompleted_Pending"))
        pstraddress_city = Trim(.Item("address_city"))
        pstraddress_country = Trim(.Item("address_country"))
        pstraddress_state = Trim(.Item("address_state"))
        pstraddress_status = Trim(.Item("address_status"))
        pstraddress_street = Trim(.Item("address_street"))
        pstraddress_zip = Trim(.Item("address_zip"))
        pbytCategory = Trim(.Item("Category"))
        pstrcustom = Trim(.Item("custom"))
        pstrfirst_name = Trim(.Item("first_name"))
        pstrinvoice = Trim(.Item("invoice"))
        pstritem_name = Trim(.Item("item_name"))
        pstritem_number = Trim(.Item("item_number"))
        pstrlast_name = Trim(.Item("last_name"))
        pstrnotify_version = Trim(.Item("notify_version"))
        pstrpayer_email = Trim(.Item("payer_email"))
        pstrpayer_status = Trim(.Item("payer_status"))
        pdtpayment_date = Trim(.Item("payment_date"))
        pdblpayment_fee = Trim(.Item("payment_fee"))
        pdblpayment_gross = Trim(.Item("payment_gross"))
        pstrpayment_status = Trim(.Item("payment_status"))
        pstrpayment_type = Trim(.Item("payment_type"))
        pstrpending_reason = Trim(.Item("pending_reason"))
        pdblquantity = Trim(.Item("quantity"))
        pstrreceiver_email = Trim(.Item("receiver_email"))
        pstrtxn_id = Trim(.Item("txn_id"))
        pstrtxn_type = Trim(.Item("txn_type"))
        pstrverify_sign = Trim(.Item("verify_sign"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Load(strtxID)

Dim pstrSQL

'On Error Resume Next

	pstrSQL = "SELECT PayPalPayments.txn_id, PayPalPayments.receiver_email, PayPalPayments.item_name, PayPalPayments.item_number, PayPalPayments.quantity, PayPalPayments.invoice, PayPalPayments.custom, PayPalPayments.payment_status, PayPalPayments.pending_reason, PayPalPayments.payment_date, PayPalPayments.payment_gross, PayPalPayments.payment_fee, PayPalPayments.txn_type, PayPalPayments.first_name, PayPalPayments.last_name, PayPalPayments.address_street, PayPalPayments.address_city, PayPalPayments.address_state, PayPalPayments.address_zip, PayPalPayments.address_country, PayPalPayments.address_status, PayPalPayments.payer_email, PayPalPayments.payer_status, PayPalPayments.payment_type, PayPalPayments.notify_version, PayPalPayments.verify_sign, PayPalPayments.Category, PayPalPayments.ActionCompleted_Completed, PayPalPayments.ActionCompleted_Pending, PayPalPayments.ActionCompleted_Failed, PayPalPayments.ActionCompleted_Denied, PayPalIPNs.payment_status, PayPalIPNs.pending_reason, PayPalIPNs.DateIPNReceived, PayPalIPNs.PayPalIPNID" _
			& " FROM PayPalPayments LEFT JOIN PayPalIPNs ON PayPalPayments.txn_id = PayPalIPNs.txn_id"

	If Len(strtxID) > 0 Then
		pstrSQL = pstrSQL & " where PayPalPayments.txn_id='" & strtxID & "'"
	End If

	Set pRS = GetRS(pstrSQL)
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        Load = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(strtxn_id)

Dim sql

'On Error Resume Next

    sql = "Delete from PayPalPayments where txn_id = '" & strtxn_id & "'"
    pobjConn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "Record successfully deleted."
        Delete = True
    Else
        pstrMessage = Err.Description
        Delete = False
    End If

End Function    'Delete

'***********************************************************************************************
End Class   'clsPayPalPayments

'If Len(Session("login")) = 0 Then Response.Redirect "Admin.asp?PrevPage=" & Request.ServerVariables("SCRIPT_NAME")
mstrPageTitle = "PayPalPayments Administration"

%>
<!--#include file="AdminHeader.asp"-->
<!--#include file="PayPal_IPN_DB_Connection.asp"-->
<%
'Assumptions:
'   Connection: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsPayPalPayments
Dim mstrTXID

	mAction = LoadRequestValue("Action")
	mstrTXID = LoadRequestValue("txn_id")

    Set mclsPayPalPayments = New clsPayPalPayments
    
    Select Case mAction
        Case "Delete"
            mclsPayPalPayments.Delete mstrTXID
            Call ShowMessage
        Case "View"
            If mclsPayPalPayments.Load(mstrTXID) Then Call ShowDetail
        Case Else
            If mclsPayPalPayments.Load(mstrTXID) Then Call ShowDetail
    End Select
    
Sub ShowDetail

    With mclsPayPalPayments
%>

<SCRIPT LANGUAGE=javascript>
<!--

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + document.frmData.PromoTitle.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "Delete";
    theForm.submit();
    return(true);
    }
    Else
    {
    return(false);
    }
}

//-->
</SCRIPT>

<BODY>
<CENTER>
<H2><%= mstrPageTitle %></H2>
<%= .OutputMessage %>

<FORM action='PayPalPaymentsAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type=hidden id=txn_id name=txn_id value=<%= .txn_id %>>
<input type=hidden id=Action name=Action value='Update'>
<TABLE border=1 cellPadding=3 cellSpacing=0 width='95%'>
  <COLGROUP valign = top align=right />
  <COLGROUP valign = top align=left />
      <TR>
        <TD>Payment Date:</TD>
        <TD><%= .payment_date %></TD>
      </TR>
      <TR>
        <TD>Payment Type:</TD>
        <TD><%= .txn_type %></TD>
      </TR>
      <TR>
        <TD>Payment:</TD>
        <TD>
			<TABLE cellpadding="2" cellspacing="0" border="0" ID="Table1">
				<TR><TD align="right">Gross:</TD><TD align="right"><%= WriteCurrency(.payment_gross) %></TD></TR>
				<TR><TD align="right">Fee:</TD><TD align="right"><% If Not isNull(.payment_fee) Then Response.Write WriteCurrency(.payment_fee) %></TD></TR>
				<TR><TD align="right">Net:</TD><TD align="right"><% If Not isNull(.payment_fee) Then Response.Write WriteCurrency(.payment_gross - .payment_fee) %></TD></TR>
			</TABLE>
        </TD>
      </TR>
      <TR>
        <TD>Payment Status:</TD>
        <TD>
			<TABLE cellpadding="2" cellspacing="0" border="0" ID="Table2">
				<TR><TD align="right">Gross:</TD><TD align="right"><%= WriteCurrency(.payment_gross) %></TD></TR>
				<TR><TD align="right">Fee:</TD><TD align="right"><% If Not isNull(.payment_fee) Then Response.Write WriteCurrency(.payment_fee) %></TD></TR>
				<TR><TD align="right">Net:</TD><TD align="right"><% If Not isNull(.payment_fee) Then Response.Write WriteCurrency(.payment_gross - .payment_fee) %></TD></TR>
			</TABLE>
        </TD>
      </TR>
      <% If .Recordset.RecordCount = 1 Then %>
      <TR>
        <TD>Payment Status</LABEL></TD>
        <TD><%= .payment_status %>&nbsp;<% If Len(.pending_reason & "") > 0 Then Response.Write "(" & .pending_reason & ")" %></TD>
      </TR>
      <% Else %>
      <TR>
        <TD>Payment Status</LABEL></TD>
        <TD>
      </TR>
      <% 
			Do While Not .Recordset.EOF
				Response.Write .Recordset.Fields("PayPalIPNs.payment_status").Value
				If Len(.Recordset.Fields("PayPalIPNs.pending_reason").Value & "") > 0 Then Response.Write "(" & .Recordset.Fields("PayPalIPNs.pending_reason").Value & ")&nbsp;"
				Response.Write .Recordset.Fields("PayPalIPNs.pending_reason").Value	'PayPalIPNs.DateIPNReceived
				.RecordSet.MoveNext
			Loop
			.RecordSet.MoveFirst
%>
        </TD>
      </TR>
<%
         End If 
      %>
      <TR>
        <TD>&nbsp;<LABEL id=lbltxn_id for=txn_id>txn_id</LABEL></TD>
        <TD>&nbsp;<INPUT id=txn_id name=txn_id Value='<%= .txn_id %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>Paid by:</TD>
        <TD>
			<%= .first_name %>&nbsp;<%= .last_name %>&nbsp;(<%= .payer_status %>)<br/>
			<%= .address_street %><BR/>
			<%= .address_city %>,&nbsp;<%= .address_state %>&nbsp;<%= .address_zip %><BR/>
			<%= .address_country %><BR/>
			<%= .address_status %><BR/>
			<%= .payer_email %><BR/>
        </TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustom for=custom>custom</LABEL></TD>
        <TD>&nbsp;<TEXTAREA id=custom name=custom></TEXTAREA><%= .custom %></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblinvoice for=invoice>invoice</LABEL></TD>
        <TD>&nbsp;<INPUT id=invoice name=invoice Value='<%= .invoice %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblitem_name for=item_name>item_name</LABEL></TD>
        <TD>&nbsp;<INPUT id=item_name name=item_name Value='<%= .item_name %>' maxlength=100 size=60></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblitem_number for=item_number>item_number</LABEL></TD>
        <TD>&nbsp;<INPUT id=item_number name=item_number Value='<%= .item_number %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblpayment_type for=payment_type>payment_type</LABEL></TD>
        <TD>&nbsp;<INPUT id=payment_type name=payment_type Value='<%= .payment_type %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblquantity for=quantity>quantity</LABEL></TD>
        <TD>&nbsp;<INPUT id=quantity name=quantity Value='<%= .quantity %>'></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblreceiver_email for=receiver_email>receiver_email</LABEL></TD>
        <TD>&nbsp;<TEXTAREA id=receiver_email name=receiver_email></TEXTAREA><%= .receiver_email %></TD>
      </TR>
      <TR>
        <TD>&nbsp;</TD>
        <TD>&nbsp;<INPUT type="checkbox" id=Category name=Category Value="1" <% If .Category=1 Then Response.Write Checked %>><LABEL id="lblCategory" for=Category>This transaction is filed</LABEL></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblverify_sign for=verify_sign>verify_sign</LABEL></TD>
        <TD>&nbsp;<TEXTAREA id=verify_sign name=verify_sign></TEXTAREA><%= .verify_sign %></TD>
      </TR>
      <TR>
        <TD>Custom Actions:</TD>
        <TD>&nbsp;
			<% If Not isNull(.ActionCompleted_Completed) Then Response.Write "Completed action performed on " & FormatDateTime(.ActionCompleted_Completed) & ".<br />" %>
			<% If Not isNull(.ActionCompleted_Denied) Then Response.Write "Completed action performed on " & FormatDateTime(.ActionCompleted_Denied) & ".<br />" %>
			<% If Not isNull(.ActionCompleted_Failed) Then Response.Write "Completed action performed on " & FormatDateTime(.ActionCompleted_Failed) & ".<br />" %>
			<% If Not isNull(.ActionCompleted_Pending) Then Response.Write "Completed action performed on " & FormatDateTime(.ActionCompleted_Pending) & ".<br />" %>
        </TD>
      </TR>
  <TR>
    <TD>&nbsp;</TD>
    <TD>
        <INPUT id=btnReset name=btnReset type=reset value=Reset>&nbsp;&nbsp;
        <INPUT id=btnDelete name=btnDelete type=button value=Delete onclick='return btnDelete_onclick(this)'>
        <INPUT id=btnUpdate name=btnUpdate type=submit value='Save Changes'>
    </TD>
    <TD>&nbsp;</TD>
  </TR>
</TABLE>
</FORM>
<% 

    End With
End Sub	'ShowDetail 
%>

</CENTER>
</BODY>
</HTML>
<%
    Set mclsPayPalPayments = Nothing
    Set pobjConn = Nothing
    Response.Flush
%>
