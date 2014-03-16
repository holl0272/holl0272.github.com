<%
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
Private pstrinvoice
Private pstritem_name
Private pstritem_number
Private pstrfirst_name
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

Public Property Get invoice()
    invoice = pstrinvoice
End Property

Public Property Get item_name()
    item_name = pstritem_name
End Property

Public Property Get item_number()
    item_number = pstritem_number
End Property

Public Property Get first_name()
    first_name = pstrfirst_name
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
    pstrinvoice = trim(objRS.Fields("invoice").Value)
    pstritem_name = trim(objRS.Fields("item_name").Value)
    pstritem_number = trim(objRS.Fields("item_number").Value)
    pstrfirst_name = trim(objRS.Fields("first_name").Value)
    pstrlast_name = trim(objRS.Fields("last_name").Value)
    pstrnotify_version = trim(objRS.Fields("notify_version").Value)
    pstrpayer_email = trim(objRS.Fields("payer_email").Value)
    pstrpayer_status = trim(objRS.Fields("payer_status").Value)
    pdtpayment_date = trim(objRS.Fields("payment_date").Value)
    pdblpayment_fee = trim(objRS.Fields("payment_fee").Value)
    pdblpayment_gross = trim(objRS.Fields("payment_gross").Value)
    pstrpayment_status = trim(objRS.Fields("payment_status").Value)
    pstrpayment_type = trim(objRS.Fields("payment_type").Value)
    pstrpending_reason = trim(objRS.Fields("pending_reason").Value)
    pdblquantity = trim(objRS.Fields("quantity").Value)
    pstrreceiver_email = trim(objRS.Fields("receiver_email").Value)
    pstrtxn_id = trim(objRS.Fields("txn_id").Value)
    pstrtxn_type = trim(objRS.Fields("txn_type").Value)
    pstrverify_sign = trim(objRS.Fields("verify_sign").Value)

End Sub 'LoadValues

'***********************************************************************************************

Public Function LoadAll()

'On Error Resume Next

    Set pRS = GetRS("Select * from PayPalPayments " & mstrsqlWhere)
    
	If pRS.State <> 1 Then
		Response.Write "<div class='FatalError'>You need to upgrade your database to use PayPal Payments</div>" _
						& "<h3><a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=PayPalPayments'>Click here to upgrade</a></h3>"
		pstrMessage = "Loading Error " & Err.number & ": " & Err.Description
		Err.Clear
		Response.Flush
		LoadAll = False
		Exit Function
	End If

    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    Else
		LoadAll = False
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Load(strtxID)

Dim pstrSQL

'On Error Resume Next

	pstrSQL = "SELECT PayPalPayments.txn_id, PayPalPayments.receiver_email, PayPalPayments.item_name, PayPalPayments.item_number, PayPalPayments.quantity, PayPalPayments.invoice, PayPalPayments.custom, PayPalPayments.payment_status, PayPalPayments.pending_reason, PayPalPayments.payment_date, PayPalPayments.payment_gross, PayPalPayments.payment_fee, PayPalPayments.txn_type, PayPalPayments.first_name, PayPalPayments.last_name, PayPalPayments.address_street, PayPalPayments.address_city, PayPalPayments.address_state, PayPalPayments.address_zip, PayPalPayments.address_country, PayPalPayments.address_status, PayPalPayments.payer_email, PayPalPayments.payer_status, PayPalPayments.payment_type, PayPalPayments.notify_version, PayPalPayments.verify_sign, PayPalPayments.Category, PayPalPayments.ActionCompleted_Completed, PayPalPayments.ActionCompleted_Pending, PayPalPayments.ActionCompleted_Failed, PayPalPayments.ActionCompleted_Denied, PayPalIPNs.payment_status as IPNpaymentStatus, PayPalIPNs.pending_reason as IPNpendingReason, PayPalIPNs.DateIPNReceived, PayPalIPNs.PayPalIPNID" _
			& " FROM PayPalPayments LEFT JOIN PayPalIPNs ON PayPalPayments.txn_id = PayPalIPNs.txn_id"

	If Len(strtxID) > 0 Then
		pstrSQL = pstrSQL & " where PayPalPayments.txn_id='" & strtxID & "'"
	End If

	Set pRS = GetRS(pstrSQL)
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        Load = True
    Else
		Load = False
    End If

End Function    'Load

'***********************************************************************************************

Public Function LoadByItemNumber(strtxID)

Dim pstrSQL

'On Error Resume Next

	If Len(strtxID) = 0 Then
		LoadByItemNumber = False
		Exit Function
	End If

	pstrSQL = "SELECT PayPalPayments.txn_id, PayPalPayments.receiver_email, PayPalPayments.item_name, PayPalPayments.item_number, PayPalPayments.quantity, PayPalPayments.invoice, PayPalPayments.custom, PayPalPayments.payment_status, PayPalPayments.pending_reason, PayPalPayments.payment_date, PayPalPayments.payment_gross, PayPalPayments.payment_fee, PayPalPayments.txn_type, PayPalPayments.first_name, PayPalPayments.last_name, PayPalPayments.address_street, PayPalPayments.address_city, PayPalPayments.address_state, PayPalPayments.address_zip, PayPalPayments.address_country, PayPalPayments.address_status, PayPalPayments.payer_email, PayPalPayments.payer_status, PayPalPayments.payment_type, PayPalPayments.notify_version, PayPalPayments.verify_sign, PayPalPayments.Category, PayPalPayments.ActionCompleted_Completed, PayPalPayments.ActionCompleted_Pending, PayPalPayments.ActionCompleted_Failed, PayPalPayments.ActionCompleted_Denied, PayPalIPNs.payment_status as IPNpaymentStatus, PayPalIPNs.pending_reason as IPNpendingReason, PayPalIPNs.DateIPNReceived, PayPalIPNs.PayPalIPNID" _
			& " FROM PayPalPayments LEFT JOIN PayPalIPNs ON PayPalPayments.txn_id = PayPalIPNs.txn_id" _
			& " WHERE PayPalPayments.item_number='" & strtxID & "'"

	Set pRS = GetRS(pstrSQL)
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadByItemNumber = True
    Else
		LoadByItemNumber = False
    End If

End Function    'LoadByItemNumber

'***********************************************************************************************

Public Function Delete(strtxn_id)

Dim sql

'On Error Resume Next

    sql = "Delete from PayPalPayments where txn_id = '" & strtxn_id & "'"
	If Not isObject(pobjConn) Then Call InitializeConnection
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

Public Sub OutputSummary()

'On Error Resume Next

	Dim i
	Dim aSortHeader(9,1)
	Dim pstrOrderBy, pstrSortOrder, pstrTempSort
	Dim pstrTitle
	Dim pstrSelect, pstrHighlight
	Dim pstrID
	Dim pblnSelected
	Dim pblnClosed
	Dim pbytStartPoint
	Dim pbytEndPoint
	Dim pblnOddRow
	Dim pstrBGColor
	
		With Response

			If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
				pstrTempSort = "descending"
				pstrSortOrder = "ASC"
			Else
				pstrTempSort = "ascending"
				pstrSortOrder = "DESC"
			End If
			
			aSortHeader(1,0) = "Sort by Item Name in " & pstrTempSort & " order"
			aSortHeader(2,0) = "Sort by Item Number in " & pstrTempSort & " order"
			aSortHeader(3,0) = "Sort by Quantity in " & pstrTempSort & " order"
			aSortHeader(4,0) = "Sort by Last Names in " & pstrTempSort & " order"
			aSortHeader(5,0) = "Sort by Payment Date in " & pstrTempSort & " order"
			aSortHeader(6,0) = "Sort by Payment in " & pstrTempSort & " order"
			aSortHeader(7,0) = "Sort by Fee in " & pstrTempSort & " order"
			aSortHeader(8,0) = "Sort by Net Amount in " & pstrTempSort & " order"
			aSortHeader(9,0) = "Sort by Payment Status in " & pstrTempSort & " order"
				
			aSortHeader(1,1) = "Item Name"
			aSortHeader(2,1) = "Item Number"
			aSortHeader(3,1) = "Qty"
			aSortHeader(4,1) = "Last Name"
			aSortHeader(5,1) = "Payment Date"
			aSortHeader(6,1) = "Payment"
			aSortHeader(7,1) = "Fee"
			aSortHeader(8,1) = "Net"
			aSortHeader(9,1) = "Status"

			.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' id='tblSummary' >" & vbcrlf	'rules='none'
			.Write "	<tr class='tblhdr'>" & vbcrlf
			
			.Write "      <TH>&nbsp;</TH>" & vbcrlf
			if len(mstrOrderBy) > 0 Then
				pstrOrderBy = mstrOrderBy
			Else
				pstrOrderBy = "1"
			End If
		
			For i = 1 to UBound(aSortHeader)
				If cInt(pstrOrderBy) = i Then
					If (pstrSortOrder = "ASC") Then
						.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'DESC');" & chr(34) & _
										" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
										" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
										"&nbsp;&nbsp;&nbsp;<img src='images/up.gif' border=0 align=bottom></TH>" & vbCrLf
					Else
						.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'ASC');" & chr(34) & _
										" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
										" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
										"&nbsp;&nbsp;&nbsp;<img src='images/down.gif' border=0 align=bottom></TH>" & vbCrLf
					End If
				Else
				    .Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'" & pstrSortOrder & "');" & chr(34) & _
									" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
									" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & "</TH>" & vbCrLf
				End If
			Next 'i
'			.Write "<th>&nbsp;</th>"
			.Write "	</tr>"
		End With

	Dim mlngQty
	Dim mdblGross
	Dim mdblFee
	Dim pstrName
	Dim mdblGross_Item
	Dim mdblFee_Item
	Dim mlngQty_Item
	
	mlngQty = 0
	mdblGross = 0
	mdblFee = 0
	
    If prs.RecordCount > 0 Then
        prs.MoveFirst
		pblnOddRow = True
        For i = 1 To prs.RecordCount
			pstrName = ""
			If pblnOddRow Then
				pstrBGColor = "white"
			Else
				pstrBGColor = "lightgrey"
			End If
			Response.Write "<tr bgcolor=" & pstrBGColor & ">" & vbcrlf
			If prs.Fields("Category").Value = 1 Then
				Response.Write "<td align=center><input type='checkbox' name='txn_id' value='" & prs.Fields("txn_id").Value & "' checked></td>"
			Else
				Response.Write "<td align=center><input type='checkbox' name='txn_id' value='" & prs.Fields("txn_id").Value & "'></td>"
			End If
			
			Response.Write "<td>&nbsp;" & prs.Fields("item_name").Value & "&nbsp;</td>" & vbcrlf
			Response.Write "<td align=right>" & prs.Fields("item_number").Value & "&nbsp;&nbsp;</td>" & vbcrlf
			Response.Write "<td align=center>" & prs.Fields("quantity").Value & "&nbsp;</td>" & vbcrlf
			
			pstrName = Trim(prs.Fields("last_name").Value & "")
			If Len(pstrName) = 0 Then
				pstrName = prs.Fields("first_name").Value
			Else
				If Len(prs.Fields("first_name").Value & "") > 0 Then pstrName = pstrName & ", " & prs.Fields("first_name").Value
			End If
			Select Case Trim(prs.Fields("payer_status").Value & "")
				Case "verified"
					pstrName = "<span title='Verified'>" & pstrName & "</span>"
				Case "unverified"
					pstrName = "<span title='Unverified'><b>" & pstrName & "</b></span>"
				Case "intl_verified"
					pstrName = "<span title='International Verified'><i>" & pstrName & "</i></span>"
				Case "intl_unverified"
					pstrName = "<span title='International Unverified'><i><b>" & pstrName & "</b></i></span>"
			End Select
			
			mdblGross_Item = Trim(prs.Fields("payment_gross").Value & "")
			If Not isNumeric(mdblGross_Item) Then mdblGross_Item = 0
			
			mdblFee_Item = Trim(prs.Fields("payment_fee").Value & "")
			If Not isNumeric(mdblFee_Item) Then mdblFee_Item = 0
			
			mlngQty_Item = Trim(prs.Fields("quantity").Value & "")
			If Not isNumeric(mlngQty_Item) Then mlngQty_Item = 0

			Response.Write "<td>&nbsp;&nbsp;<a href='mailto:" & prs.Fields("payer_email").Value & "' title='Send an email to this person'>" & pstrName & "</a></td>" & vbcrlf
			Response.Write "<td align=right><a href='https://www.paypal.com/vst/id=" & prs.Fields("txn_id").Value & "' title='View details at PayPal.com' target='_blank'>" & prs.Fields("payment_date").Value & "</a>&nbsp;&nbsp;</td>" & vbcrlf
			Response.Write "<td align=right>" & FormatCurrency(mdblGross_Item,2) & "&nbsp;&nbsp;</td>" & vbcrlf
			If isNull(prs.Fields("payment_fee").Value) Then
				Response.Write "<td>&nbsp;</td><td>&nbsp;</td>" & vbcrlf
			Else
				Response.Write "<td align=right>" & FormatCurrency(mdblFee_Item,2) & "&nbsp;&nbsp;</td>" & vbcrlf
				Response.Write "<td align=right>" & FormatCurrency(mdblGross_Item - mdblFee_Item,2) & "&nbsp;&nbsp;</td>" & vbcrlf
			End If

			Select Case Trim(LCase(prs.Fields("payment_status").Value) & "")
				Case "completed"
					Response.Write "<td bgcolor=lightgreen align=center><a href='' onclick=" & Chr(34) & "OpenIPNDetail('" & prs.Fields("txn_id").Value & "'); return false;" & Chr(34) & " title='View IPN details'>Completed</a></td>" & vbcrlf
				Case "pending"
					Response.Write "<td bgcolor=yellow align=center><a href='' onclick=" & Chr(34) & "OpenIPNDetail('" & prs.Fields("txn_id").Value & "'); return false;" & Chr(34) & " title='View IPN details'>Pending&nbsp;(" & prs.Fields("pending_reason").Value & ")</a></td>" & vbcrlf
				Case "failed"
					Response.Write "<td bgcolor=red align=center><a href='' onclick=" & Chr(34) & "OpenIPNDetail('" & prs.Fields("txn_id").Value & "'); return false;" & Chr(34) & " title='View IPN details'>Failed&nbsp;(" & prs.Fields("pending_reason").Value & ")</a></td>" & vbcrlf
				Case "denied"
					Response.Write "<td bgcolor=red align=center><a href='' onclick=" & Chr(34) & "OpenIPNDetail('" & prs.Fields("txn_id").Value & "'); return false;" & Chr(34) & " title='View IPN details'>Denied&nbsp;(" & prs.Fields("pending_reason").Value & ")</a></td>" & vbcrlf
			End Select
			
'			Response.Write "<td>&nbsp;&nbsp;"
'			Select Case Trim(prs.Fields("address_status").Value & "")
'				Case ""
'					Response.Write "<span title='No address provided'>-</span>"
'				Case "confirmed"
'					Response.Write "<span title='Confirmed address'>C</span>"
'				Case "unconfirmed"
'					Response.Write "<span title='Unconfirmed address'>V</span>"
'			End Select
'			Response.Write "&nbsp;&nbsp;"
'			Select Case Trim(prs.Fields("payer_status").Value & "")
'				Case "verified"
'					Response.Write "<span title='Verified'>V</span>"
'				Case "unverified"
'					Response.Write "<span title='Unverified'>U</span>"
'				Case "intl_verified"
'					Response.Write "<span title='International Verified'>IV</span>"
'				Case "intl_unverified"
'					Response.Write "<span title='International Unverified'>IU</span>"
'			End Select
'			Response.Write "</td>" & vbcrlf
			
			Response.Write "</TR>" & vbcrlf
			
			If Not isNull(prs.Fields("quantity").Value) Then mlngQty = mlngQty + mlngQty_Item
			If Not isNull(prs.Fields("payment_gross").Value) Then mdblGross = mdblGross + mdblGross_Item
			If Not isNull(prs.Fields("payment_fee").Value) Then mdblFee = mdblFee + mdblFee_Item
            prs.MoveNext
            pblnOddRow = Not pblnOddRow
        Next
        
        Response.Write "<tr>" & vbcrlf
        Response.Write "  <th colspan=3>&nbsp;</th>" & vbcrlf
        Response.Write "  <th>" & mlngQty & "</th>" & vbcrlf
        Response.Write "  <th colspan=2>&nbsp;</th>" & vbcrlf
        Response.Write "  <th align=right>" & FormatCurrency(mdblGross,2) & "&nbsp;&nbsp;</th>" & vbcrlf
        Response.Write "  <th align=right>" & FormatCurrency(mdblFee,2) & "&nbsp;&nbsp;</th>" & vbcrlf
        Response.Write "  <th align=right>" & FormatCurrency(mdblGross - mdblFee,2) & "&nbsp;&nbsp;</th>" & vbcrlf
        Response.Write "  <th colspan=2>&nbsp;</th>" & vbcrlf
        Response.Write "</tr>" & vbcrlf
    Else
        Response.Write "<TR><TH colspan=11><h3>There are no Payments</h3></TH></TR>" & vbcrlf
    End If
%>
	</TABLE>
<%

End Sub      'OutputSummary

'***********************************************************************************************

Public Sub FileSelectedItems

Dim paryItems
Dim pstrItems
Dim i
Dim pstrSQL

	pstrItems = Request.Form("txn_id")
	If Len(pstrItems) > 0 Then
		paryItems = Split(pstrItems,",")
		If Not isObject(pobjConn) Then Call InitializeConnection
		For i = 0 To UBound(paryItems)
			pstrSQL = "Update PayPalPayments Set Category=1 Where txn_id='" & Trim(paryItems(i)) & "'"
			pobjConn.Execute pstrSQL,,128
		Next 'i
	End If

End Sub	'FileSelectedItems

'***********************************************************************************************

Public Sub FileItem(strTXID,blnFile)

Dim pstrSQL

	If Len(strTXID) > 0 Then
		If blnFile Then
			pstrSQL = "Update PayPalPayments Set Category=1 Where txn_id='" & Trim(strTXID) & "'"
		Else
			pstrSQL = "Update PayPalPayments Set Category=0 Where txn_id='" & Trim(strTXID) & "'"
		End If
		If Not isObject(pobjConn) Then Call InitializeConnection
		pobjConn.Execute pstrSQL,,128
	End If

End Sub	'FileSelectedItems

End Class   'clsPayPalPayments

'***********************************************************************************************

Dim mclsPayPalPayments
%>