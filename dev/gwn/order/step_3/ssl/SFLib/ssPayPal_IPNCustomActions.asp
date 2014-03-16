<% 
'********************************************************************************
'*   PayPal Payments															*
'*   Release Version:   3.00.002												*
'*   Release Date:		March 17, 2003											*
'*   Revision Date:		April 13, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Release 1.00.002 (April 13, 2004)											*
'*	 - Enhancement - added debugging code										*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'Insert your custom code for Complete actions
'--------------------------------------------

Const cblnHasCompletedAction = True
Const cblnHasPendingAction = True
Const cblnHasFailedAction = True
Const cblnHasDeniedAction = True

'--------------------------------------------

	Sub PerformCustomAction(strtxn_id,strPaymentStatus,strPendingReason,strPayerStatus, strAddressStatus, blnManual)

		Output "PerformCustomAction -- PerformCustomAction called on " & Now() & vbcrlf
		Output "PerformCustomAction	-- strtxn_id: " & strtxn_id & vbcrlf
		Output "PerformCustomAction	-- strPaymentStatus: " & strPaymentStatus & vbcrlf
		
		If Len(strtxn_id) = 0 Then Exit Sub
		Select Case LCase(strPaymentStatus)
			Case "completed":	Call PerformAction_Completed(strtxn_id,strPayerStatus, strAddressStatus, blnManual)
			Case "pending":		Call PerformAction_Pending(strtxn_id,strPendingReason,strPayerStatus, strAddressStatus, blnManual)
			Case "failed":		Call PerformAction_Failed(strtxn_id)
			Case "denied":		Call PerformAction_Denied(strtxn_id)
			Case Else:		Call PerformAction_Completed(strtxn_id,"Unknown - " & strPayerStatus, strAddressStatus, blnManual)
		End Select		

	End Sub 'PerformCustomAction

	'***********************************************************************************************

	Sub PerformAction_Completed(strtxn_id, strPayerStatus, strAddressStatus, blnManual)

	Dim pstrSQL
	Dim pobjRS
	Dim pblnDoCustomAction
	Dim plngOrderID
	Dim plngPos
	Dim pstrTempInvoice

		'pstrInvoice defined in ssPayPal_InstantNotification only; calling custom action from admin will err out
		On Error Resume Next
		pstrTempInvoice = pstrInvoice
		If Err.number <> 0 Then Err.Clear
		On Error Goto 0		
		
		pblnDoCustomAction = False
		
		If Len(strtxn_id) = 0 Then Exit Sub
		If Not cblnHasCompletedAction Then Exit Sub
		
		pstrSQL = "Select ActionCompleted_Completed, item_name from PayPalPayments where txn_id = '" & strtxn_id & "'"
		Output "PerformAction_Completed - pstrSQL: " & pstrSQL & "<br>" & vbcrlf

		Set pobjRS = server.CreateObject("adodb.Recordset")
		pobjRS.CursorLocation = 2	'adUseClient
		pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		Output "PerformAction_Completed -- PerformAction_Completed called on " & Now() & vbcrlf
		Output "PerformAction_Completed	-- pstrSQL: " & pstrSQL & vbcrlf
		Output "PerformAction_Completed	-- pobjRS.EOF: " & pobjRS.EOF & vbcrlf
		
		If Len(pstrTempInvoice) = 0 And Not pobjRS.EOF Then
			plngOrderID = Trim(pobjRS.Fields("item_name").Value & "")
			plngPos = InStrRev(plngOrderID, " ")
			If plngPos > 0 Then plngOrderID = Right(plngOrderID, Len(plngOrderID) - plngPos)
		Else
			plngOrderID = pstrInvoice
		End If

		If isNull(pobjRS.Fields("ActionCompleted_Completed").Value) Or cblndebug_PayPalIPN Then	'

			'Insert your custom code for Complete actions
			'--------------------------------------------
			'strPayerStatus - verified;unverified;intl_verified;intl_unverified
			'strAddressStatus -confirmed;unconfirmed
			
			'Set logic based on above parameters
			
			pblnDoCustomAction = True	'will do for all

			'Do not alter this line - permits manual action
			pblnDoCustomAction = pblnDoCustomAction Or blnManual
			
			If pblnDoCustomAction Then Call CustomIPN_StoreFront5(strtxn_id, plngOrderID)
			'If pblnDoCustomAction Then Call MyAction(lngOrderID)
			
			'--------------------------------------------

			pstrSQL = "Update PayPalPayments Set ActionCompleted_Completed = " & sqlDateWrap(Now()) & " Where txn_id='" & strtxn_id & "'"
			cnn.Execute pstrSQL,,128

		End If
		pobjRS.Close
		Set pobjRS = Nothing
			
	End Sub 'PerformAction_Completed

	'***********************************************************************************************

	Sub PerformAction_Pending(strtxn_id, strPendingReason, strPayerStatus, strAddressStatus, blnManual)

	Dim pstrSQL
	Dim pobjRS
	Dim pblnDoCustomAction

	'On Error Resume Next

		If Not cblnHasPendingAction Then Exit Sub
		pblnDoCustomAction = False
		pstrSQL = "Select ActionCompleted_Pending from PayPalPayments where txn_id = '" & strtxn_id & "' And ActionCompleted_Pending Is Null"
		Set pobjRS = server.CreateObject("adodb.Recordset")
		pobjRS.CursorLocation = 2	'adUseClient
		pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly

		If Not pobjRS.EOF Then

			'Insert your custom code for Pending actions
			'--------------------------------------------
			'strPendingReason - echeck;intl;verify;address;upgrade;unilateral;other
			'strPayerStatus - verified;unverified;intl_verified;intl_unverified
			'strAddressStatus -confirmed;unconfirmed
			
			'Set logic based on above parameters
			
			pblnDoCustomAction = True	'will do for all
			
			'Do not alter this line - permits manual action
			pblnDoCustomAction = pblnDoCustomAction Or blnManual
			
			If pblnDoCustomAction Then Call PerformAction_Completed(strtxn_id, strPayerStatus, strAddressStatus, blnManual)

			'--------------------------------------------

			pstrSQL = "Update PayPalPayments Set ActionCompleted_Pending = " & sqlDateWrap(Now()) & " Where txn_id='" & strtxn_id & "'"
			cnn.Execute pstrSQL,,128
		End If
		pobjRS.Close
		Set pobjRS = Nothing
			
	End Sub 'PerformAction_Pending

	'***********************************************************************************************

	Sub PerformAction_Failed(strtxn_id)

	Dim pstrSQL
	Dim pobjRS

	'On Error Resume Next

		If Not cblnHasFailedAction Then Exit Sub
		pstrSQL = "Select ActionCompleted_Failed from PayPalPayments where txn_id = '" & strtxn_id & "' And ActionCompleted_Failed Is Null"
		Set pobjRS = server.CreateObject("adodb.Recordset")
		pobjRS.CursorLocation = 2	'adUseClient
		pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly

		If Not pobjRS.EOF Then

			'Insert your custom code for Failed actions
			'--------------------------------------------
			

			
			'--------------------------------------------

			pstrSQL = "Update PayPalPayments Set ActionCompleted_Failed = " & sqlDateWrap(Now()) & " Where txn_id='" & strtxn_id & "'"
			cnn.Execute pstrSQL,,128
		End If
		pobjRS.Close
		Set pobjRS = Nothing
			
	End Sub 'PerformAction_Failed

	'***********************************************************************************************

	Sub PerformAction_Denied(strtxn_id)

	Dim pstrSQL
	Dim pobjRS

	'On Error Resume Next

		If Not cblnHasDeniedAction Then Exit Sub
		pstrSQL = "Select ActionCompleted_Denied from PayPalPayments where txn_id = '" & strtxn_id & "' And ActionCompleted_Denied Is Null"
		Set pobjRS = server.CreateObject("adodb.Recordset")
		pobjRS.CursorLocation = 2	'adUseClient
		pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly

		If Not pobjRS.EOF Then

			'Insert your custom code for Denied actions
			'--------------------------------------------
			

			
			'--------------------------------------------

			pstrSQL = "Update PayPalPayments Set ActionCompleted_Denied = " & sqlDateWrap(Now()) & " Where txn_id='" & strtxn_id & "'"
			cnn.Execute pstrSQL,,128
		End If
		pobjRS.Close
		Set pobjRS = Nothing
			
	End Sub 'PerformAction_Denied
	
	'***********************************************************************************************

	Sub CustomIPN_StoreFront5(byVal strtxn_id, byVal lngOrderID)

	Dim pstrSQL
	Dim pobjRS
	
'	On Error Resume Next

		'Check added for non-order related payments
		If Len(lngOrderID) = 0 Or Not isNumeric(lngOrderID) Then Exit Sub

		pstrSQL = "Select orderGrandTotal From sfOrders Where orderID = " & lngOrderID	'invoice used as Order Number
		Output "CustomIPN_StoreFront5 - pstrSQL: " & pstrSQL & vbcrlf
		Set pobjRS = server.CreateObject("adodb.Recordset")
		pobjRS.CursorLocation = 2	'adUseClient
		pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		Output "CustomIPN_StoreFront5 - pobjRS.EOF: " & pobjRS.EOF & vbcrlf
		If Not pobjRS.EOF Then
			'Found Order, now compare amount paid/amount due
			Dim pdblAmountDue, pdblAmountPaid
			Dim pstrErrorMessage
			
			If isNumeric(pobjRS.Fields("orderGrandTotal").Value) Then
				pdblAmountDue = CDbl(pobjRS.Fields("orderGrandTotal").Value)
			Else
				pstrErrorMessage = "Non-numeric amount due of " & pobjRS.Fields("orderGrandTotal").Value & "."
				pdblAmountDue = 0
			End If
			
			If isNumeric(pstrpayment_gross) Then
				pdblAmountPaid = CDbl(pstrpayment_gross)
			Else
				pstrErrorMessage = "Non-numeric amount paid of " & pstrpayment_gross & "."
				pdblAmountPaid = 0
			End If
			
			If pdblAmountPaid < pdblAmountDue Then
				pstrErrorMessage = pstrErrorMessage & vbcrlf _
								 & "Funds paid (" & FormatCurrency(pdblAmountPaid,2) & ") are less than amount due (" & FormatCurrency(pdblAmountDue,2) & ")"
			End If
			
			Output "CustomIPN_StoreFront5 - pdblAmountDue: " & pdblAmountDue & vbcrlf
			Output "CustomIPN_StoreFront5 - pdblAmountPaid: " & pdblAmountPaid & vbcrlf
			Output "CustomIPN_StoreFront5 - pstrErrorMessage: " & pstrErrorMessage & vbcrlf

			pstrSQL = "Select trnsrspSuccess From sfTransactionResponse Where trnsrspOrderId = " & lngOrderID
			Output "CustomIPN_StoreFront5 - pstrSQL: " & pstrSQL & vbcrlf
			Set pobjRS = server.CreateObject("adodb.Recordset")
			pobjRS.CursorLocation = 2	'adUseClient
			pobjRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			Output "CustomIPN_StoreFront5 - pobjRS.EOF: " & pobjRS.EOF & vbcrlf
				
			If pobjRS.EOF Or cblndebug_PayPalIPN Then
				pstrSQL = "Insert Into sfTransactionResponse" _
						& " (trnsrspOrderId, trnsrspCustTransNo, trnsrspMerchTransNo, trnsrspAVSCode, trnsrspAUXMsg, trnsrspActionCode ,trnsrspAuthNo ,trnsrspErrorLocation ,trnsrspSuccess ,trnsrspErrorMsg)" _
						& " Values (" _
						& "  " & lngOrderID & ", " _
						& " '" & strtxn_id & "', " _
						& " '" & pstrverify_sign & "', " _
						& " '" & "not applicable" & "', " _
						& " '" & "not applicable" & "', " _
						& " '" & pstrpayment_status & "', " _
						& " '" & strtxn_id & "', " _
						& " '" & "paypal" & "', " _
						& " '" & pstrpayment_status & "', " _
						& " '" & pstrErrorMessage & "')"
				Output "CustomIPN_StoreFront5 - pstrSQL: " & pstrSQL & vbcrlf
				cnn.Execute pstrSQL,,128
				
				'Now set the order complete
				pstrSQL = "Update sfOrders Set orderIsComplete=1 Where orderID=" & lngOrderID
				Output "CustomIPN_StoreFront5 - pstrSQL: " & pstrSQL & vbcrlf
				cnn.Execute pstrSQL,,128
				
				'Now if using Order Manager one should update the payment record
 				pstrSQL = "Insert Into ssOrderManager (ssorderID, ssDatePaymentReceived, ssPaidVia, ssInternalNotes)" _
 						& " Values (" _
 						& " " & lngOrderID & "," _
 						& " " & sqlDateWrap(pstrpayment_date) & "," _
 						& " 'PayPal'," _
 						& " '" & "Payment Recorded via PayPal IPN" & vbcrlf & pstrErrorMessage & "')"
				
				'Since not everybody will have Order Manager
				On Error Resume Next
				Output "CustomIPN_StoreFront5 - pstrSQL: " & pstrSQL & vbcrlf
				cnn.Execute pstrSQL,,128
				If Err.number <> 0 Then Err.Clear
				On Error Goto 0
				
				'One could send an email here but that is beyond this scope
				
			Else
				'this will happen for a payment update such as a check clearing
				pstrSQL = "Update sfTransactionResponse" _
						& " Set " _
						& "  trnsrspActionCode='" & pstrpayment_status & "', " _
						& "  trnsrspSuccess='" & strtxn_id & "', " _
						& "  trnsrspErrorMsg='" & pstrErrorMessage & "'" _
						& " Where trnsrspOrderId=" & lngOrderID
				Output "CustomIPN_StoreFront5 - pstrSQL: " & pstrSQL & vbcrlf
				cnn.Execute pstrSQL,,128

				'Now set the order complete
				pstrSQL = "Update sfOrders Set orderIsComplete=1 Where orderID=" & lngOrderID
				Output "CustomIPN_StoreFront5 - pstrSQL: " & pstrSQL & vbcrlf
				cnn.Execute pstrSQL,,128
			End If
		End If
		
		pobjRS.Close
		Set pobjRS = Nothing
			
	End Sub 'CustomIPN_StoreFront5
	
	'***********************************************************************************************

	Sub MyAction(lngOrderID)

	Dim mlngOrderID

		'On Error Resume Next

		mlngOrderID = lngOrderID
		If len(mlngOrderID) = 0 Then Exit Sub


	End Sub	'MyAction
%>


<% 
Sub WriteCustomActionTable

    With mclsPayPalPayments
		If cblnHasCompletedAction OR cblnHasPendingAction OR cblnHasFailedAction OR cblnHasDeniedAction Then
%>
<script language="javascript">
function PerformCustomAction(bytAction)
{
	document.frmData.CustomAction.value = bytAction;
	document.frmData.Action.value = "PerformCustomAction";
	document.frmData.submit();
}
</script>
	 <input type="hidden" id="CustomAction" name="CustomAction" value="">
     <tr>
       <td colspan="2">
         <table width="100%" border="1" cellpadding="2" cellspacing="0" ID="Table1">
			<tr><th colspan="3">Custom Actions</th></tr>
			<tr>
			  <th>&nbsp;</th>
			  <th>Accomplished On</th>
			  <th>Perfom</th>
			</tr>
			
			<% If cblnHasCompletedAction Then %>
			<tr>
			  <td align="center">Completed</td>
			  <td align="center"><% If Len(Trim(.ActionCompleted_Completed & "")) > 0 Then Response.Write FormatDateTime(.ActionCompleted_Completed) %>&nbsp;</td>
			  <td align="center"><a href="" onclick="PerformCustomAction(0); return false;" title="">Perform Completed Action</a></td>
			</tr>
			<% End If	'cblnHasCompletedAction %>
			
			<% If cblnHasPendingAction Then %>
			<tr>
			  <td align="center">Pending</td>
			  <td align="center"><% If Len(Trim(.ActionCompleted_Pending & "")) > 0 Then Response.Write FormatDateTime(.ActionCompleted_Pending) %>&nbsp;</td>
			  <td align="center"><a href="" onclick="PerformCustomAction(1); return false;" title="">Perform Pending Action</a></td>
			</tr>
			<% End If	'cblnHasPendingAction %>
			
			<% If cblnHasFailedAction Then %>
			<tr>
			  <td align="center">Failed</td>
			  <td align="center"><% If Len(Trim(.ActionCompleted_Failed & "")) > 0 Then Response.Write FormatDateTime(.ActionCompleted_Failed) %>&nbsp;</td>
			  <td align="center"><a href="" onclick="PerformCustomAction(2); return false;" title="">Perform Failed Action</a></td>
			</tr>
			<% End If	'cblnHasFailedAction %>
			
			<% If cblnHasDeniedAction Then %>
			<tr>
			  <td align="center">Completed</td>
			  <td align="center"><% If Len(Trim(.ActionCompleted_Denied & "")) > 0 Then Response.Write FormatDateTime(.ActionCompleted_Denied) %>&nbsp;</td>
			  <td align="center"><a href="" onclick="PerformCustomAction(3); return false;" title="">Perform Denied Action</a></td>
			</tr>
			<% End If	'cblnHasDeniedAction %>
			
         </table>
        </td>
      </tr>
     <% 
     End If
	End With
End Sub	'WriteCustomActionTable
%>
