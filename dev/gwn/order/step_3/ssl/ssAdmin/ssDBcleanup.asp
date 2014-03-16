<% Option Explicit 
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version:	2.00.002		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		September 25, 2004										*
'*                                                                              *
'*   Release 2.00.003 (September 25, 2004)										*
'*	   - Bug Fix - Updated dates to be compatible with non-U.S. dates			*
'*                                                                              *
'*   Release 2.00.002 (February 21, 2004)										*
'*	   - Bug Fix - Updated CC deletion routine for SQL Server compatibility		*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Dim cbytRemoveOldOrders
Dim cbytRemoveIncompleteOrders
Dim cbytRemoveSavedOrders
Dim cbytRemoveTempOrders
Dim cbytRemoveCCNumbers
Dim daysToKeepVisitors
Dim cstrCCReplace

'***********************************************************************************************************

Function InitializeDefaults

Dim pobjRS

	Set pobjRS = GetRS("Select * From sfAdmin")
	With pobjRS
		If Not .EOF Then
			cbytRemoveOldOrders = Trim(.Fields("hoursBetweenCleanings").Value & "")
			cbytRemoveIncompleteOrders = Trim(.Fields("daysToSaveIncompleteOrders").Value & "")
			cbytRemoveTempOrders = Trim(.Fields("daysToSaveTempOrders").Value & "")
			cbytRemoveSavedOrders = Trim(.Fields("daysToSaveSavedOrders").Value & "")
			cbytRemoveCCNumbers = Trim(.Fields("adminDeleteSchedule").Value & "")
			daysToKeepVisitors = Trim(.Fields("daysToKeepVisitors").Value & "")
			cstrCCReplace = Trim(.Fields("CCReplace").Value & "")
			
		End If

		If Len(cbytRemoveOldOrders) = 0 Then cbytRemoveOldOrders = 6
		If Len(cbytRemoveIncompleteOrders) = 0 Then cbytRemoveIncompleteOrders = 14
		If Len(cbytRemoveTempOrders) = 0 Then cbytRemoveTempOrders = 6
		If Len(cbytRemoveSavedOrders) = 0 Then cbytRemoveSavedOrders = 14
		If Len(cbytRemoveCCNumbers) = 0 Then cbytRemoveCCNumbers = 14
		If Len(daysToKeepVisitors) = 0 Then daysToKeepVisitors = 7
		.Close
	End With

End Function

'***********************************************************************************************************

Function ClearTransactionLogFile(byVal bytOption)

Dim pstrSQL
Dim pobjRS
Dim i
    
    
    pstrSQL = "sp_spaceused"
    Set pobjRS = GetRS(pstrSQL)
    If pobjRS.EOF Then
		Response.Write "EOF"
    Else
		'Response.Write "Database Name: " & pobjRS.Fields("database_name").Value & "<br />"
		'Response.Write "Size: " & pobjRS.Fields("database_size").Value & "<br />"
		'Response.Write "Unallocated Space: " & pobjRS.Fields("unallocated space").Value & "<br />"
		
		Select Case CStr(bytOption)
			Case "1":
				pstrSQL = "BACKUP LOG [" & pobjRS.Fields("database_name").Value & "] WITH TRUNCATE_ONLY"
				cnn.Execute pstrSQL,,128
				mstrMessage = mstrMessage & "Transaction log truncated.<br />"
			Case "2":
				pstrSQL = "DUMP TRANSACTION [" & pobjRS.Fields("database_name").Value & "] WITH NO_LOG"
				cnn.Execute pstrSQL,,128
				mstrMessage = mstrMessage & "Transaction log dumped.<br />"
			Case Else:
				pstrSQL = ""
		End Select
    End If
    Call ReleaseObject(pobjRS)

End Function	'ClearTransactionLogFile

'***********************************************************************************************************

Function CompactRepair(blnCreateBackUp)

Const dbProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
Dim pobjFSO
Dim pobjJRO
Dim pobjFile

Dim pstrDBConn

Dim pstrDBSource
Dim pstrDBTarget
Dim plngOrigFileSize
Dim plngNewFileSize

Dim plngPos1, plngPos2

	'Use existing connection information to determine path to database
	plngPos1 = Instr(1,cnn.ConnectionString,".mdb")
	If plngPos1 > 0 Then
		plngPos2 = InstrRev(cnn.ConnectionString,"=",plngPos1)
		If plngPos1 > plngPos2 Then	
			pstrDBSource = Mid(cnn.ConnectionString,plngPos2+1,plngPos1-plngPos2+3)
			pstrDBTarget = Replace(pstrDBSource,".mdb", "_" & Day(Date) & MonthName(Month(Date),True) & Year(Date) & ".mdb")
		End If
	Else
		mstrMessage = "This does not appear to be an Access Database!"
		CompactRepair = False
		Exit Function
	End If

	Set pobjFSO = Server.CreateObject("Scripting.FileSystemObject")
	If pobjFSO.FileExists(pstrDBTarget) Then pobjFSO.DeleteFile(pstrDBTarget)
	plngOrigFileSize = pobjFSO.GetFile(pstrDBSource).Size
	
	On Error Resume Next
 
	cnn.Close	'Close the connection since nobody can be connected for the compact/repair process to proceed
	Set pobjJRO = Server.CreateObject("JRO.JetEngine")
	If Err.number = -2147319779 Then
		mstrMessage = mstrMessage & "<H3><font color='red'>Error: JRO.JetEngine Library not found. Please contact your server administrator to install this.</font></H3>"
		CompactRepair = False
		Exit Function
	ElseIf Err.number <> 0 Then
		mstrMessage = mstrMessage & "<H3><font color='red'>Error: " & Err.number & ": " & Err.Description & "</font></H3>"
		CompactRepair = False
		Exit Function
	End If
		
	pobjJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pstrDBSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pstrDBTarget
	Set pobjJRO = nothing
	
	If Err.number = -2147467259  Then
		mstrMessage = mstrMessage & "<H3><font color='red'>Your database was in use so I could not compact the database.</font></H3>"
	ElseIf Err.number <> 0 Then
		mstrMessage = mstrMessage & "<H3><font color='red'>Error: " & Err.number & ": " & Err.Description & "</font></H3>"
	Else
		pobjFSO.CopyFile pstrDBTarget, pstrDBSource, True
		plngNewFileSize = pobjFSO.GetFile(pstrDBSource).Size
		mstrMessage = mstrMessage & "<H3>Database Successfully Compacted and Repaired! Original Size: " & plngOrigFileSize & " bytes. New Size: " & plngNewFileSize & " bytes.</H3>"
		If Not blnCreateBackUp Then 
			pobjFSO.DeleteFile(pstrDBTarget)
		Else
			mstrMessage = mstrMessage & "<br /><i>Your back-up copy is at <i>" & pstrDBTarget & "</i></H3>"
		End If
	End If
	
'	cnn.Open 'Technically should reopen the connection since that's the way it started but no reason to here since this is the last step
	
	Set pobjFSO = nothing
  
	CompactRepair = (Err.number = 0)
  
End Function	'CompactRepair

'***********************************************************************************************************

Sub CleanUp

Dim pstrSQL
Dim DeleteOldOrdersDate,DeleteIncompleteOrdersDate,DeleteSavedOrdersDate,DeleteTempOrdersDate,DeleteCCNumbers
Dim pcmdSelect
Dim pcmdDeleteOrderDetails, pcmdDeleteOrderAttrributes
Dim rsOldOrders,rsOldOrderDetails

	DeleteOldOrdersDate = DateAdd("d", -1 * mvalOldOrders, Now())
	DeleteIncompleteOrdersDate = DateAdd("d", -1 * mvalOrders, Now())
	DeleteSavedOrdersDate = DateAdd("d", -1 * mvalSaved, Now())
	DeleteTempOrdersDate = DateAdd("d", -1 * mvalTemp, Now())
	DeleteCCNumbers = DateAdd("d", -1 * mvalCC, Now())

	'***********************************************************************************************

	If mblnOldOrders Then	'Delete old Orders
		'order information saved to sfOrders, sfOrderDetails, sfOrderAttributes, sfCPayments, sfTransactionResponse
		'order information saved to sfOrdersAE, sfOrderDetailsAE for AE
		'information left in sfCustomers & sfCShipAddresses intentionally
		
		If True Then
			set rsOldOrders = GetRS("Select orderID, orderPayId from sfOrders where orderDate<=" & sqlDateWrap(makeISODate(DeleteOldOrdersDate)) & " AND orderIsComplete=1")
		Else
		set pcmdSelect = server.CreateObject("ADODB.Command")
		With pcmdSelect
			.CommandType = 1	'adCmdText
			.CommandText = "Select orderID, orderPayId from sfOrders where orderDate<=? AND orderIsComplete=1"
			.Parameters.Append .CreateParameter("DeleteOldOrdersDate",7,1,,DeleteOldOrdersDate)
			.ActiveConnection = cnn
			Set rsOldOrders = .Execute()
		End With
		End If

	
		If not rsOldOrders.EOF Then
			set pcmdSelect = server.CreateObject("ADODB.Command")
			With pcmdSelect
				.CommandType = 1	'adCmdText
				.CommandText = "Select odrdtID from sfOrderDetails Where odrdtOrderId=?"
				.Parameters.Append .CreateParameter("orderID",20,1,,rsOldOrders("orderID"))
				.ActiveConnection = cnn
				Set rsOldOrderDetails = pcmdSelect.Execute()
			End With
			set pcmdDeleteOrderDetails = server.CreateObject("ADODB.Command")
			
			If False Then
			
				With pcmdDeleteOrderDetails
					.CommandType = 1	'adCmdText
					.CommandText = "Delete from sfOrderDetails Where odrdtID=?"
					.Parameters.Append .CreateParameter("odrdtID",20,1,,rsOldOrders("orderID"))
					.ActiveConnection = cnn
				End With
					
				While not rsOldOrders.EOF
					pcmdSelect.Parameters("orderID").value = rsOldOrders("orderID")
					Set rsOldOrderDetails = pcmdSelect.Execute()
			
					If not rsOldOrderDetails.EOF Then
						set pcmdDeleteOrderAttrributes = server.CreateObject("ADODB.Command")
						With pcmdDeleteOrderAttrributes
							.CommandType = 1	'adCmdText
							.CommandText = "Delete from sfOrderAttributes Where odrattrOrderDetailId=?"
							.Parameters.Append .CreateParameter("odrattrOrderDetailId",20,1,,rsOldOrderDetails("odrdtID"))
							.ActiveConnection = cnn
							While not rsOldOrderDetails.EOF
								.Parameters("odrattrOrderDetailId").value = rsOldOrderDetails("odrdtID")
								.Execute
								rsOldOrderDetails.MoveNext
							Wend
						End With
					End If

					pcmdDeleteOrderDetails.Parameters("odrdtID").value = rsOldOrders("orderID")
					pcmdDeleteOrderDetails.Execute
						
					rsOldOrders.MoveNext
				Wend
				
				set pcmdSelect = server.CreateObject("ADODB.Command")
				With pcmdSelect
					.CommandType = 1	'adCmdText
					.CommandText = "Delete from sfOrders where orderIsComplete=1 AND orderDate<=?"
					.Parameters.Append .CreateParameter("DeleteIncompleteOrdersDate",7,1,,DeleteIncompleteOrdersDate)
					.ActiveConnection = cnn
					.Execute()
				End With
				
			Else
				'delete the order detail information
				While not rsOldOrderDetails.EOF
					pstrSQL = "Delete from sfOrderAttributes Where odrattrOrderDetailId=" & rsOldOrderDetails("odrdtID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128
					If cblnSF5AE Then 
						pstrSQL = "Delete from sfOrderDetailsAE Where odrdtAEID=" & rsOldOrderDetails("odrdtID")
						'Response.Write "pstrSQL = " & pstrSQL & "<br />"
						cnn.Execute pstrSQL,,128
					End If
					rsOldOrderDetails.MoveNext
				Wend

				'order information saved to sfOrders, sfOrderDetails, sfOrderAttributes, sfCPayments, sfTransactionResponse
				'order information saved to sfOrdersAE, sfOrderDetailsAE for AE
				While not rsOldOrders.EOF
					pstrSQL = "Delete from sfOrderDetails Where odrdtOrderId=" & rsOldOrders("orderID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128
					
					pstrSQL = "Delete from sfCPayments Where payID=" & rsOldOrders("orderPayId")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128
					
					pstrSQL = "Delete from sfTransactionResponse Where trnsrspOrderId=" & rsOldOrders("orderID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128
					
					If cblnSF5AE Then 
						pstrSQL = "Delete from sfOrdersAE Where orderAEID=" & rsOldOrders("orderID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
						cnn.Execute pstrSQL,,128
					End If
					
					pstrSQL = "Delete from sfOrders Where orderID=" & rsOldOrders("orderID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128

					rsOldOrders.MoveNext
				Wend

			End If

			mstrMessage = mstrMessage & "<H3>Old, Completed Orders Deleted</H3>"
		Else
			mstrMessage = mstrMessage & "<H3>There were no Old, Completed Orders to Delete</H3>"
		End If
	End If	'Delete old Orders

	'***********************************************************************************************

	If mblnOrders Then	'Delete old incomplete Orders
		set pcmdSelect = server.CreateObject("ADODB.Command")
		If True Then
			set rsOldOrders = GetRS("Select orderID, orderPayId from sfOrders where orderDate<=" & sqlDateWrap(makeISODate(DeleteOldOrdersDate)) & " AND (orderIsComplete<>1 OR orderIsComplete is Null)")
		Else
		With pcmdSelect
			.CommandType = 1	'adCmdText
			.CommandText = "Select orderID, orderPayId from sfOrders where orderDate<=? AND (orderIsComplete<>1 OR orderIsComplete is Null)"
			.Parameters.Append .CreateParameter("DeleteIncompleteOrdersDate",7,1,,DeleteIncompleteOrdersDate)
			.ActiveConnection = cnn
			Set rsOldOrders = .Execute()
		End With
		End If
	
		If not rsOldOrders.EOF Then
			set pcmdSelect = server.CreateObject("ADODB.Command")
			With pcmdSelect
				.CommandType = 1	'adCmdText
				.CommandText = "Select odrdtID from sfOrderDetails Where odrdtOrderId=?"
				.Parameters.Append .CreateParameter("orderID",20,1,,rsOldOrders("orderID"))
				.ActiveConnection = cnn
				Set rsOldOrderDetails = pcmdSelect.Execute()
			End With
			set pcmdDeleteOrderDetails = server.CreateObject("ADODB.Command")
			
			If False Then
			
				With pcmdDeleteOrderDetails
					.CommandType = 1	'adCmdText
					.CommandText = "Delete from sfOrderDetails Where odrdtID=?"
					.Parameters.Append .CreateParameter("odrdtID",20,1,,rsOldOrders("orderID"))
					.ActiveConnection = cnn
				End With
					
				While not rsOldOrders.EOF
					pcmdSelect.Parameters("orderID").value = rsOldOrders("orderID")
					Set rsOldOrderDetails = pcmdSelect.Execute()
			
					If not rsOldOrderDetails.EOF Then
						set pcmdDeleteOrderAttrributes = server.CreateObject("ADODB.Command")
						With pcmdDeleteOrderAttrributes
							.CommandType = 1	'adCmdText
							.CommandText = "Delete from sfOrderAttributes Where odrattrOrderDetailId=?"
							.Parameters.Append .CreateParameter("odrattrOrderDetailId",20,1,,rsOldOrderDetails("odrdtID"))
							.ActiveConnection = cnn
							While not rsOldOrderDetails.EOF
								.Parameters("odrattrOrderDetailId").value = rsOldOrderDetails("odrdtID")
								.Execute
								rsOldOrderDetails.MoveNext
							Wend
						End With
					End If

					pcmdDeleteOrderDetails.Parameters("odrdtID").value = rsOldOrders("orderID")
					pcmdDeleteOrderDetails.Execute
						
					rsOldOrders.MoveNext
				Wend
				
				set pcmdSelect = server.CreateObject("ADODB.Command")
				With pcmdSelect
					.CommandType = 1	'adCmdText
					.CommandText = "Delete from sfOrders where orderIsComplete=1 AND orderDate<=?"
					.Parameters.Append .CreateParameter("DeleteIncompleteOrdersDate",7,1,,DeleteIncompleteOrdersDate)
					.ActiveConnection = cnn
					.Execute()
				End With
				
			Else
				'delete the order detail information
				While not rsOldOrderDetails.EOF
					pstrSQL = "Delete from sfOrderAttributes Where odrattrOrderDetailId=" & rsOldOrderDetails("odrdtID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128
					If cblnSF5AE Then 
						pstrSQL = "Delete from sfOrderDetailsAE Where odrdtAEID=" & rsOldOrderDetails("odrdtID")
						'Response.Write "pstrSQL = " & pstrSQL & "<br />"
						cnn.Execute pstrSQL,,128
					End If
					rsOldOrderDetails.MoveNext
				Wend

				'order information saved to sfOrders, sfOrderDetails, sfOrderAttributes, sfCPayments, sfTransactionResponse
				'order information saved to sfOrdersAE, sfOrderDetailsAE for AE
				While not rsOldOrders.EOF
					pstrSQL = "Delete from sfOrderDetails Where odrdtOrderId=" & rsOldOrders("orderID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128
					
					pstrSQL = "Delete from sfCPayments Where payID=" & rsOldOrders("orderPayId")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128
					
					pstrSQL = "Delete from sfTransactionResponse Where trnsrspOrderId=" & rsOldOrders("orderID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128
					
					If cblnSF5AE Then 
						pstrSQL = "Delete from sfOrdersAE Where orderAEID=" & rsOldOrders("orderID")
					Response.Write "pstrSQL = " & pstrSQL & "<br />"
					'cnn.Execute pstrSQL,,128
					End If
					
					pstrSQL = "Delete from sfOrders Where orderID=" & rsOldOrders("orderID")
					'Response.Write "pstrSQL = " & pstrSQL & "<br />"
					cnn.Execute pstrSQL,,128

					rsOldOrders.MoveNext
				Wend

			End If

			mstrMessage = mstrMessage & "<H3>Old, Incomplete Orders Deleted</H3>"
		Else
			mstrMessage = mstrMessage & "<H3>There were no Old, Incomplete Orders to Delete</H3>"
		End If
	End If	'Delete old incomplete Orders

	'***********************************************************************************************

	If mblnTemp	 Then	'Delete old temporary Orders
		If True Then
			set rsOldOrders = GetRS("Select odrdttmpID from sfTmpOrderDetails where odrdttmpDate<=" & sqlDateWrap(makeISODate(DeleteOldOrdersDate)))
		Else
		set pcmdSelect = server.CreateObject("ADODB.Command")
		With pcmdSelect
			.CommandType = 1	'adCmdText
			.CommandText = "Select odrdttmpID from sfTmpOrderDetails where odrdttmpDate<=?"
			.Parameters.Append .CreateParameter("DeleteIncompleteOrdersDate",7,1,,DeleteTempOrdersDate)
			.ActiveConnection = cnn
			Set rsOldOrders = .Execute()
		End With
		End If
		
		If not rsOldOrders.EOF Then
		
			If False Then

				set pcmdDeleteOrderDetails = server.CreateObject("ADODB.Command")
				With pcmdDeleteOrderDetails
					.CommandType = 1	'adCmdText
					.CommandText = "Delete from sfTmpOrderAttributes Where odrattrtmpOrderDetailId=?"
					.Parameters.Append .CreateParameter("odrattrtmpOrderDetailId",20,1,,rsOldOrders("odrdttmpID"))
					.ActiveConnection = cnn

					While not rsOldOrders.EOF
						.Parameters("odrattrtmpOrderDetailId").value = rsOldOrders("odrdttmpID")
						.Execute()
						rsOldOrders.MoveNext
					Wend

				End With
					
				set pcmdSelect = server.CreateObject("ADODB.Command")
				With pcmdSelect
					.CommandType = 1	'adCmdText
					.CommandText = "Delete from sfTmpOrderDetails where odrdttmpDate<=?"
					.Parameters.Append .CreateParameter("DeleteIncompleteOrdersDate",7,1,,DeleteTempOrdersDate)
					.ActiveConnection = cnn
					.Execute()
				End With
			Else
				While not rsOldOrders.EOF
					cnn.Execute "Delete from sfTmpOrderAttributes Where odrattrtmpOrderDetailId=" & rsOldOrders("odrdttmpID"),,128
					rsOldOrders.MoveNext
				Wend

				cnn.Execute "Delete from sfTmpOrderDetails where odrdttmpDate<=" & sqlDateWrap(makeISODate(DeleteTempOrdersDate)),,128
				
			End If
			
			mstrMessage = mstrMessage & "<H3>Old, Incomplete Temporary Orders Deleted</H3>"
		Else
			mstrMessage = mstrMessage & "<H3>There were no Old, Incomplete Temporary Orders to Delete</H3>"
		End If
	End If		'Delete old temporary Orders

	'***********************************************************************************************

	If mblnSaved Then		'Delete old saved Orders
		set pcmdSelect = server.CreateObject("ADODB.Command")
		If True Then
			set rsOldOrders = GetRS("Select odrdtsvdID from sfSavedOrderDetails where odrdtsvdDate<=" & sqlDateWrap(makeISODate(DeleteOldOrdersDate)))
		Else
		With pcmdSelect
			.CommandType = 1	'adCmdText
			.CommandText = "Select odrdtsvdID from sfSavedOrderDetails where odrdtsvdDate<=?"
			.Parameters.Append .CreateParameter("odrdtsvdDate",7,1,,DeleteSavedOrdersDate)
			.ActiveConnection = cnn
			Set rsOldOrders = .Execute()
		End With
		End If
			
		If not rsOldOrders.EOF Then
		
			If False Then

				set pcmdDeleteOrderDetails = server.CreateObject("ADODB.Command")
				With pcmdDeleteOrderDetails
					.CommandType = 1	'adCmdText
					.CommandText = "Delete from sfSavedOrderAttributes Where odrattrsvdOrderDetailId=?"
					.Parameters.Append .CreateParameter("odrattrsvdOrderDetailId",20,1,,rsOldOrders("odrdtsvdID"))
					.ActiveConnection = cnn
					
					While not rsOldOrders.EOF
						.Parameters("odrattrsvdOrderDetailId").value = rsOldOrders("odrdtsvdID")
						.Execute()
						rsOldOrders.MoveNext
					Wend
					
				End With

				set pcmdSelect = server.CreateObject("ADODB.Command")
				With pcmdSelect
					.CommandType = 1	'adCmdText
					.CommandText = "Delete from sfSavedOrderDetails where odrdtsvdDate<=?"
					.Parameters.Append .CreateParameter("DeleteIncompleteOrdersDate",7,1,,DeleteSavedOrdersDate)
					.ActiveConnection = cnn
					.Execute()
				End With
				
			Else
				While not rsOldOrders.EOF
					cnn.Execute "Delete from sfSavedOrderAttributes Where odrattrsvdOrderDetailId=" & rsOldOrders("odrdtsvdID"),,128
					rsOldOrders.MoveNext
				Wend

				cnn.Execute "Delete from sfSavedOrderDetails where odrdtsvdDate<=" & sqlDateWrap(makeISODate(DeleteSavedOrdersDate)),,128
			
			End If
			
			mstrMessage = mstrMessage & "<H3>Old, Saved Orders Deleted</H3>"
		Else
			mstrMessage = mstrMessage & "<H3>There were no Old, Saved Orders to Delete</H3>"
		End If
	End If		'Delete old saved Orders

	'***********************************************************************************************

	If mblnCC Then		'Delete old CC numbers
		DeleteIncompleteOrdersDate = Date() - mvalCC
		If True Then
		
			If Len(cstrCCV) = 0 Then
				pstrSQL = "UPDATE sfCPayments " _
						& " SET sfCPayments.payCardNumber = '****'" _
						& " WHERE payID IN" _
						& " (SELECT orderPayId FROM sfOrders" _
						& "  WHERE (orderDate < " & sqlDateWrap(makeISODate(DeleteIncompleteOrdersDate)) & ") AND (orderPayId IS NOT NULL))"
			Else
				pstrSQL = "UPDATE sfCPayments " _
						& " SET sfCPayments.payCardNumber = '****', " & cstrCCV & "=Null"_
						& " WHERE payID IN" _
						& " (SELECT orderPayId FROM sfOrders" _
						& "  WHERE (orderDate < " & sqlDateWrap(makeISODate(DeleteIncompleteOrdersDate)) & ") AND (orderPayId IS NOT NULL))"
			End If
					
			'Response.Write("pstrSQL: " & pstrSQL & "<br />")
			'Response.Flush()
			cnn.Execute pstrSQL,,128
		Else
		pstrSQL = "UPDATE sfCustomers INNER JOIN sfCPayments ON sfCustomers.custID = sfCPayments.payCustId SET sfCPayments.payCardNumber = '****' WHERE (((sfCustomers.custLastAccess)<?))"
		set pcmdSelect = server.CreateObject("ADODB.Command")
		With pcmdSelect
			.CommandType = 1	'adCmdText
			.CommandText = pstrSQL
			.Parameters.Append .CreateParameter("DeleteIncompleteOrdersDate",7,1,,DeleteIncompleteOrdersDate)
			.ActiveConnection = cnn
			.Execute()
		End With
		End If

		If Len(cstrCCV) = 0 Then
			mstrMessage = mstrMessage & "<H3>Customer Credit Card Numbers Deleted</H3>"
		Else
			mstrMessage = mstrMessage & "<H3>Customer Credit Card Numbers and codes Deleted</H3>"
		End If
	End If		'Delete old CC numbers
	
	Set pcmdSelect = Nothing
	Set pcmdDeleteOrderDetails = Nothing
	Set pcmdDeleteOrderAttrributes = Nothing
	Set rsOldOrders = Nothing
	Set rsOldOrderDetails = Nothing

End Sub
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim mvalOldOrders, mvalOrders, mvalTemp, mvalSaved, mvalCC
Dim mblnOldOrders, mblnOrders, mblnTemp, mblnSaved, mblnCC
Dim mintCompact

	Call InitializeDefaults
	
	mblnOldOrders = (Len(Request.Form("chkOldOrders")) > 0)
	If mblnOldOrders Then 
		mvalOldOrders = Request.Form("valOldOrders")
	Else
		mvalOldOrders = cbytRemoveOldOrders
	End If
	
	mblnOrders = (Len(Request.Form("chkOrders")) > 0)
	If mblnOrders Then 
		mvalOrders = Request.Form("valOrders")
	Else
		mvalOrders = cbytRemoveIncompleteOrders
	End If
	
	mblnTemp = (Len(Request.Form("chkTemp")) > 0)
	If mblnTemp Then 
		mvalTemp = Request.Form("valTemp")
	Else
		mvalTemp = cbytRemoveTempOrders
	End If
	
	mblnSaved = (Len(Request.Form("chkSaved")) > 0)
	If mblnSaved Then 
		mvalSaved = Request.Form("valSaved")
	Else
		mvalSaved = cbytRemoveSavedOrders
	End If
	
	mblnCC = (Len(Request.Form("chkCC")) > 0)
	If mblnCC Then 
		mvalCC = Request.Form("valCC")
	Else
		mvalCC = cbytRemoveCCNumbers
	End If

	If (Len(Request.Form("Action")) > 0) Then 
		Call CleanUp
		mintCompact = cInt(Request.Form("Compact"))
		If mintCompact > 0 Then
			If cblnSQLDatabase Then
				Call ClearTransactionLogFile(mintCompact)
			Else
				Call CompactRepair(mintCompact=1)
			End If
		End If
	End If

Call WriteHeader("sortables_init()",True)
%>
<script language="javascript" type="text/javascript">
<!--

function ValidInput(theForm)
{
  if (!isNumeric(theForm.valOrders,false,"Please enter an number for the Incomplete Order days.")) {return(false);}
  if (!isNumeric(theForm.valTemp,false,"Please enter an number for the Temporary Order days.")) {return(false);}
  if (!isNumeric(theForm.valSaved,false,"Please enter an number for the Saved Order days.")) {return(false);}
  if (!isNumeric(theForm.valCC,false,"Please enter an number for the Credit Card days.")) {return(false);}
  
  return(true);
}

//-->
</script>

<CENTER>
<%= mstrMessage %>
<FORM action='ssDBcleanup.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<input type=hidden name="Action" value="Clean">
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
      <TR class='tblhdr'>
        <TH>Database Clean-Up Page</TH>
      </TR>
      <TR>
        <td>This page is used to clean up various database parameters. These clean-up routines are normally automatically handled upon accessing the admin pages.<hr></td>
      </TR>
      <tr>
        <td><input type=checkbox name=chkOldOrders id="chkOldOrders">&nbsp;Remove Orders Greater than&nbsp;<input name="valOldOrders" id="OldOrders" value='<%= mvalOldOrders %>' size=3>&nbsp;days old</td>
      </tr>
      <TR>
        <TD><input type=checkbox name=chkOrders id=chkOrders>&nbsp;Remove Incomplete Orders Greater than&nbsp;<input name=valOrders id=valOrders value='<%= mvalOrders %>' size=3>&nbsp;days old</TD>
      </TR>
      <TR>
        <TD><input type=checkbox name=chkTemp id=chkTemp>&nbsp;Remove Temporary Orders Greater than&nbsp;<input name=valTemp id=valTemp value='<%= mvalTemp %>' size=3>&nbsp;days old</TD>
      </TR>
      <TR>
        <TD><input type=checkbox name=chkSaved id=chkSaved>&nbsp;Remove Saved Orders Greater than&nbsp;<input name=valSaved id=valSaved value='<%= mvalSaved %>' size=3>&nbsp;days old</TD>
      </TR>
      <TR>
        <TD><input type=checkbox name=chkCC id=chkCC>&nbsp;Remove CC Numbers <% If Len(cstrCCV) > 0 Then %>and codes <% End If %>for Orders Greater than&nbsp;<input name=valCC id=valCC value='<%= mvalCC %>' size=3>&nbsp;days old (This will override your default settings.)</TD>
      </TR>
      <% If Not cblnSQLDatabase Then %>
      <TR>
        <TD>
        <fieldset><legend>Compact & Repair Database (Access Only)</legend>
          <input type="radio" name="compact" id="compact" value="0">&nbsp;Do Not Compact & Repair<br />
          <input type="radio" name="compact" id="compact" value="1">&nbsp;Compact & Repair - Create Back-up (database name + "_today's date" (Access Only)<br />
          <input type="radio" name="compact" id="compact" value="2">&nbsp;Compact & Repair - No Back-up (Access Only)<br />
        </fieldset>
        </TD>
      </TR>
      <% Else %>
      <tr>
        <td>
        <fieldset><legend>Database Details(SQL Server Only)</legend>
        <%
        Dim i, j
        Dim pstrSQL
        Dim pobjRS
        Dim pobjRSTables
        Dim paryTableInfo
        
		Dim rows
		Dim reserved
		Dim data
		Dim index_size
		Dim unused

		Dim rows_total
		Dim reserved_total
		Dim data_total
		Dim index_size_total
		Dim unused_total

		rows_total = 0
		reserved_total = 0
		data_total = 0
		index_size_total = 0
		unused_total = 0

        pstrSQL = "sp_spaceused"
        Set pobjRS = GetRS(pstrSQL)
        If pobjRS.EOF Then
			Response.Write "EOF"
        Else
			Response.Write "Database Name: " & pobjRS.Fields("database_name").Value & "<br />"
			Response.Write "Size: " & pobjRS.Fields("database_size").Value & "<br />"
			Response.Write "Unallocated Space: " & pobjRS.Fields("unallocated space").Value & "<br />"
			
			'For i = 0 To pobjRS.Fields.Count - 1
			'Response.Write pobjRS.Fields(i).Name & ": " & pobjRS.Fields(i).Value & "<br />"
			'Next 'i
        End If
        Call ReleaseObject(pobjRS)
        
		Response.Write "<table class=""tbl"" border=1 cellspacing=0 cellpadding=2 tag=""sortable"">" & vbCrLf
		Response.Write "<colgroup>"
		Response.Write "<col align=left>"
		Response.Write "<col align=right>"
		Response.Write "<col align=right>"
		Response.Write "<col align=right>"
		Response.Write "<col align=right>"
		Response.Write "<col align=right>"
		Response.Write "</colgroup>"
		Response.Write "<tr class=""tblhdr"">"
		Response.Write "<th tag=""sortable"">Table</th>"	'name
		Response.Write "<th tag=""sortable"">Rows</th>"	'rows
		Response.Write "<th tag=""sortable"">reserved (KB)</th>"	'reserved
		Response.Write "<th tag=""sortable"">data (KB)</th>"	'data
		Response.Write "<th tag=""sortable"">index_size (KB)</th>"	'index_size
		Response.Write "<th tag=""sortable"">unused (KB)</th>"	'unused
		Response.Write "</tr>"
		
        pstrSQL = "SELECT * FROM sysobjects WHERE type='U'"
        Set pobjRSTables = GetRS(pstrSQL)
        If pobjRSTables.EOF Then
			Response.Write "EOF"
        Else
			paryTableInfo = pobjRSTables.GetRows
			
			For j = 0 To UBound(paryTableInfo, 2)
				Response.Write "<tr>"
				Response.Write "<td>" & paryTableInfo(0, j) & "</td>"	'name
				
				pstrSQL = "sp_spaceused '" & paryTableInfo(0, j) & "'"
				Set pobjRS = GetRS(pstrSQL)

				If pobjRS.EOF Then
					Response.Write "<td colspan=3>No Data</td>"
				Else
					rows = Replace(pobjRS.Fields("rows").Value, " KB", "")
					reserved = Replace(pobjRS.Fields("reserved").Value, " KB", "")
					data = Replace(pobjRS.Fields("data").Value, " KB", "")
					index_size = Replace(pobjRS.Fields("index_size").Value, " KB", "")
					unused = Replace(pobjRS.Fields("unused").Value, " KB", "")

					rows_total = rows_total + CDbl(rows)
					reserved_total = reserved_total + CDbl(reserved)
					data_total = data_total + CDbl(data)
					index_size_total = index_size_total + CDbl(index_size)
					unused_total = unused_total + CDbl(unused)

					Response.Write "<td>" & rows & "</td>"
					Response.Write "<td>" & reserved & "</td>"
					Response.Write "<td>" & data & "</td>"
					Response.Write "<td>" & index_size & "</td>"
					Response.Write "<td>" & unused & "</td>"
				End If
				Call ReleaseObject(pobjRS)
				
				Response.Write "</tr>"				
			Next 'i

			Response.Write "<tr class=""tblhdr"">"
			Response.Write "<th>Total</th>"	'name
			Response.Write "<th>" & rows_total & "</th>"	'rows
			Response.Write "<th>" & reserved_total & "</th>"	'reserved
			Response.Write "<th>" & data_total & "</th>"	'data
			Response.Write "<th>" & index_size_total & "</th>"	'index_size
			Response.Write "<th>" & unused_total & "</th>"	'unused
			Response.Write "</tr>"

        End If
        Call ReleaseObject(pobjRSTables)
		Response.Write "</table>"
        
       
        %>Transaction Log:<br />
          <input type="radio" name="compact" id="compact3" value="0" checked>&nbsp;Do Nothing<br />
          <input type="radio" name="compact" id="compact4" value="1">&nbsp;Clear Log File - Truncate Log<br />
          <input type="radio" name="compact" id="compact5" value="2">&nbsp;Clear Log File - Dump Log<br />
       </fieldset>
        </td>
      </tr>
      <% End If	'Not cblnSQLDatabase %>
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