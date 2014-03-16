<%Option Explicit
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   2.00.0001												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		November 15, 2003										*
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

Class clsOrder
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pblnError

Private prsOrders
Private prsOrderSummaries

Private prsOrderTransactions

'database variables

'ssorderID
'ssExternalNotes
'ssInternalNotes
'ssDatePaymentReceived
'ssDateOrderShipped
'ssShippedVia
'ssTrackingNumber
'ssOrderStatus
'ssDateEmailSent

Private plngAltOrderID

Private plngssorderID
Private pstrssExternalNotes
Private pstrssInternalNotes
Private pdtssDatePaymentReceived
Private pdtssDateOrderShipped
Private pstrssPaidVia
Private pbytssShippedVia
Private pstrssTrackingNumber
Private pbytssOrderStatus
Private pdtssDateEmailSent

Private pblnssOrderFlagged
Private pdtssBackOrderDateNotified
Private pdtssBackOrderDateExpected
Private pstrssBackOrderMessage
Private pstrssBackOrderInternalMessage

'Order Sent Email Parameters
Private pstrEmailTo
Private pstrEmailSubject
Private pstrEmailBody
Private pblnSendMail

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	Call ReleaseObject(prsOrders)
	Call ReleaseObject(prsOrderSummaries)
	Call ReleaseObject(prsOrderTransactions)
End Sub

'***********************************************************************************************

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


	'***********************************************************************************************

	Public Function LoadOrderSummaries(strSQLParmeters)

	dim pstrSQL
	dim p_strWhere
	dim i
	dim sql

'	On Error Resume Next

		pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custLastName, sfOrderDetails.odrdtID, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssExported, ssOrderManager.ssOrderFlagged, ssOrderManager.ssBackOrderDateNotified, ssOrderManager.ssBackOrderDateExpected" _
				& " FROM ((sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
				& " GROUP BY sfOrders.orderID, sfCustomers.custLastName, sfOrderDetails.odrdtID, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, sfOrders.orderIsComplete, sfOrders.orderTradingPartner, ssOrderManager.ssExported, ssOrderManager.ssOrderFlagged, ssOrderManager.ssBackOrderDateNotified, ssOrderManager.ssBackOrderDateExpected" _
				& strSQLParmeters

		set	prsOrderSummaries = server.CreateObject("adodb.recordset")
		With prsOrderSummaries
	        .CursorLocation = 3 'adUseClient
	'        .CursorType = 3 'adOpenStatic
	        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
	        .LockType = 1 'adLockReadOnly
			.Open pstrSQL, cnn

			If Err.number <> 0 Then
				debugprint "pstrSQL",pstrSQL
				pstrMessage = "Loading Error " & Err.number & ": " & Err.Description
				Err.Clear
				LoadOrderSummaries = False
				Exit Function
			End If
			
			If Not .EOF Then plngAltOrderID = .Fields("orderID").Value
			
			LoadOrderSummaries = (Not .EOF)
		End With

	End Function    'LoadOrderSummaries

	'***********************************************************************************************

	Public Sub OutputSummary()

	'On Error Resume Next

	Dim i
	Dim aSortHeader(9,3)
	Dim pstrOrderBy, pstrSortOrder, pstrTempSort
	Dim pstrTitle
	Dim pstrSelect, pstrHighlight
	Dim pstrID
	Dim pblnSelected
	Dim pblnClosed
	Dim pbytStartPoint
	Dim pbytEndPoint

		With Response

				
			aSortHeader(1,1) = " "
			aSortHeader(2,1) = "Order Number"
			aSortHeader(3,1) = "Last Name"
			aSortHeader(4,1) = "Items"
			aSortHeader(5,1) = "Order Total"
			aSortHeader(6,1) = "Order Date"
			aSortHeader(7,1) = "Payment Received"
			aSortHeader(8,1) = "Order Shipped"
			aSortHeader(9,1) = "Back Ordered"

			
			If cblnUseOrderFlags Then 
				pbytStartPoint = 1
			Else
				pbytStartPoint = 2
			End If
			If cblnUseBackOrder Then
				pbytEndPoint = 9
			Else
				pbytEndPoint = 8
			End If

			.Write "<table cellpadding='4' cellspacing='0' border='1' id='tblSummary' rules='none'>"
			.Write "	<tr class='tblhdr'>"
			For i = pbytStartPoint to pbytEndPoint
			    .Write "  <TH>" & aSortHeader(i,1) & "</TH>" & vbCrLf
			Next 'i
			.Write "	</tr>"

	    If prsOrderSummaries.RecordCount > 0 Then
	        prsOrderSummaries.MoveFirst

	'Need to calculate current recordset page and upper bound to loop through
	dim plnguBound, plnglbound, pstrDisplay

		If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
		If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
		plnglbound = (mlngAbsolutePage - 1) * prsOrderSummaries.PageSize + 1
		plnguBound = mlngAbsolutePage * prsOrderSummaries.PageSize

        For i = 1 To prsOrderSummaries.RecordCount
			If pstrID <> trim(prsOrderSummaries("OrderID")) Then
				pstrID = trim(prsOrderSummaries("OrderID"))

				.Write " <TR>"

				If cblnUseOrderFlags Then 
					If prsOrderSummaries.Fields("ssOrderFlagged").Value = 1 Then
						.Write "<TD align=center><img src='images/MSGBOX03.ICO' alt='x' title='flagged for follow up' height='12' width='12'></TD>"
					Else
						.Write "<TD align=center>&nbsp;</TD>"
					End If
				End If
        		.Write "<TD align=center>" & pstrID & "&nbsp;</TD>"
        		.Write "<TD align=center>" & prsOrderSummaries.Fields("custLastName").Value & "&nbsp;</TD>"
        		.Write "<TD align=center>" & prsOrderSummaries.Fields("SumOfodrdtQuantity").Value & "&nbsp;</TD>"
        		.Write "<TD align=center>" & FormatCurrency(prsOrderSummaries.Fields("orderGrandTotal").Value,2) & "&nbsp;</TD>"
        		.Write "<TD align=center>" & prsOrderSummaries.Fields("orderDate").Value & "&nbsp;</TD>"
        		.Write "<TD align=center>" & prsOrderSummaries.Fields("ssDatePaymentReceived").Value & "&nbsp;</TD>"
        		.Write "<TD align=center>" & prsOrderSummaries.Fields("ssDateOrderShipped").Value & "&nbsp;</TD>"
				If cblnUseBackOrder Then .Write "<TD align=center>" & prsOrderSummaries.Fields("ssBackOrderDateExpected").Value & "&nbsp;</TD>"
	        	
        		.Write "</TR>"
        	End If
            prsOrderSummaries.MoveNext
        Next
        
	    Else
				.Write "<TR><TD align=center COLSPAN=6><h3>There are no Orders</h3></TD></TR>"
	    End If

			'Write the paging routine
			.Write "<tr class='tblhdr'><TH COLSPAN='" & pbytEndPoint+1 & "' align=center>"
			If prsOrderSummaries.RecordCount = 0 Then
				.Write "No Orders match your search criteria"
			Elseif prsOrderSummaries.RecordCount = 1 Then
				.Write "1 Order matches your search criteria"
			Else 
				.Write prsOrderSummaries.RecordCount & " Orders match your search criteria<br />"

			End If
			.Write "</TH></TR>"
			.Write "</TABLE>"
		End With
	End Sub      'OutputSummary

	'***********************************************************************************************

	Function ValidateValues()

	Dim strError

	    strError = ""
	    
		If False Then
			debugprint "plngssorderID",plngssorderID
			debugprint "pstrssExternalNotes",pstrssExternalNotes
			debugprint "pstrssInternalNotes",pstrssInternalNotes
			debugprint "pdtssDatePaymentReceived",pdtssDatePaymentReceived
			debugprint "pdtssDateOrderShipped",pdtssDateOrderShipped
			debugprint "pbytssShippedVia",pbytssShippedVia
			debugprint "pstrssTrackingNumber",pstrssTrackingNumber
			debugprint "pbytssOrderStatus",pbytssOrderStatus

			debugprint "pstrEmailTo",pstrEmailTo
			debugprint "pstrEmailSubject",pstrEmailSubject
			debugprint "pstrEmailBody",pstrEmailBody
			debugprint "pblnSendMail",pblnSendMail
		End If

	    If Len(plngssorderID) = 0 Then strError = strError & "Please enter a Order ID." & cstrDelimeter
	    If Len(pdtssDatePaymentReceived) > 0 And Not isDate(pdtssDatePaymentReceived) Then strError = strError & "The Payment Received Date must be a valid date format." & cstrDelimeter
	    If Len(pdtssDateOrderShipped) > 0 And Not isDate(pdtssDateOrderShipped) Then strError = strError & "The Order Shipped Date must be a valid date format." & cstrDelimeter

	    pstrMessage = strError
	    ValidateValues = (Len(strError) = 0)


	End Function 'ValidateValues

End Class   'clsOrder

'***********************************************************************************************

Function SummaryFilter

Dim pstrOrderBy
Dim pstrsqlWhere

' HAVING (((sfOrders.orderDate)>#10/1/2001#) AND ((ssOrderManager.ssDatePaymentReceived) Is Null) AND ((ssOrderManager.ssDateOrderShipped) Is Null) AND ((sfOrders.orderIsComplete)=1));
' Order By 

	pstrsqlWhere = " Having (sfOrders.orderIsComplete=1)"

	'load the text filter
	mbytText_Filter = Request.Form("optText_Filter")
	mstrText_Filter = Request.Form("Text_Filter")
	If len(mstrText_Filter) > 0 Then
		Select Case mbytText_Filter
			Case "0"	'Do Not Include
			Case "1"	'Order
				pstrsqlWhere = pstrsqlWhere & " AND (orderID Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "2"	'Last Name
				pstrsqlWhere = pstrsqlWhere & " AND (custLastName Like '%" & sqlSafe(mstrText_Filter) & "%')"
			Case "3"	'Affilliate
				pstrsqlWhere = pstrsqlWhere & " AND (orderTradingPartner Like '%" & mstrText_Filter & "%')"
		End Select	
	End If

	'load the radio filters
	mbytPayment_Filter = Request.Form("optPayment_Filter")
	mbytShipment_Filter = Request.Form("optShipment_Filter")
	mbytDate_Filter = Request.Form("optDate_Filter")
	mbytoptFlag_Filter	= Request.Form("optFlag_Filter")
	
	'set the defaults
	If len(mbytPayment_Filter) = 0 Then mbytShipment_Filter = 0
	If len(mbytShipment_Filter) = 0 Then mbytShipment_Filter = 1
	If len(mbytDate_Filter) = 0 Then mbytDate_Filter = 0
	If len(mbytoptFlag_Filter) = 0 Then mbytoptFlag_Filter = 0

	mstrStartDate = Request.Form("StartDate")
	If len(mstrStartDate) > 0 then 
		if cblnSQLDatabase Then
			pstrsqlWhere = pstrsqlWhere & " and (orderDate >= '" & mstrStartDate & " 12:00:00 AM')"
		Else
			pstrsqlWhere = pstrsqlWhere & " and (orderDate >= #" & mstrStartDate & " 12:00:00 AM#)"
		End If
	End If
	
	mstrEndDate = Request.Form("EndDate")
	If len(mstrEndDate) > 0 then 
		If cblnSQLDatabase Then
			pstrsqlWhere = pstrsqlWhere & " and (orderDate <= '" & mstrEndDate & " 11:59:59 PM')"
		Else
			pstrsqlWhere = pstrsqlWhere & " and (orderDate <= #" & mstrEndDate & " 11:59:59 PM#)"
		End If
	End If


	Select Case mbytPayment_Filter
		Case "0"	'Do Not Include
		Case "1"	'Active
			pstrsqlWhere = pstrsqlWhere & " and ssDatePaymentReceived is Not Null"
		Case "2"	'Inactive
			pstrsqlWhere = pstrsqlWhere & " and ssDatePaymentReceived is Null"
	End Select	

	Select Case mbytoptFlag_Filter
		Case "0"	'Do Not Include
		Case "1"	'Flagged
			pstrsqlWhere = pstrsqlWhere & " and ssOrderFlagged=1"
		Case "2"	'unflagged
			pstrsqlWhere = pstrsqlWhere & " and (ssOrderFlagged=0 OR ssOrderFlagged is Null)"
	End Select	

	Select Case mbytShipment_Filter
		Case "0"	'Do Not Include
		Case "1"	'Active
			pstrsqlWhere = pstrsqlWhere & " and ssDateOrderShipped is Null"
		Case "2"	'Inactive
			pstrsqlWhere = pstrsqlWhere & " and ssDateOrderShipped is Not Null"
	End Select	
	
	'Build  the Order By
	mstrOrderBy = Request.Form("OrderBy")
	If len(mstrOrderBy) = 0 Then mstrOrderBy = 0
	
	mstrSortOrder = Request.Form("SortOrder")
	If len(mstrSortOrder) = 0 Then mstrSortOrder = "Desc"

	dim paryOrderBy(8)
	paryOrderBy(0) = "orderDate"	'Default
	paryOrderBy(1) = "ssOrderFlagged"
	paryOrderBy(2) = "OrderID"
	paryOrderBy(3) = "custLastName"
	paryOrderBy(4) = "Sum(sfOrderDetails.odrdtQuantity)"
	paryOrderBy(5) = "orderGrandTotal"
	paryOrderBy(6) = "orderDate"
	paryOrderBy(7) = "ssDatePaymentReceived"
	paryOrderBy(8) = "ssDateOrderShipped"

	pstrOrderBy = " Order By " & paryOrderBy(mstrOrderBy) & " " & mstrSortOrder 
'debugprint "SummaryFilter",	pstrsqlWhere  & pstrOrderBy
	SummaryFilter = pstrsqlWhere  & pstrOrderBy
	
End Function    'SummaryFilter

'--------------------------------------------------------------------------------------------------

Function ConvertToBoolean(vntValue)

On Error Resume Next

	vntValue = cBool(vntValue)
	If Err.number <> 0 Then vntValue = False
	ConvertToBoolean = vntValue

End Function	'ConvertToBoolean

Sub closeObj(objItem)
	ReleaseObject objItem
End Sub


'--------------------------------------------------------------------------------------------------
%>
<!--#include file="../SFLib/mail.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/incCC.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
<%
'**************************************************
'
'	Start Code Execution
'

mstrPageTitle = "Order Administration"

'page variables
Dim mAction
Dim mclsOrder
Dim mlngOrderID

Dim mblnShowFilter, mblnShowSummary
Dim mstrsqlWhere, mstrSortOrder,mstrOrderBy

'Display setting
Dim mbytDisplay

'Filter Elements
Dim mbytText_Filter
Dim mstrText_Filter

Dim mstrStartDate, mstrEndDate

Dim mbytShipment_Filter
Dim mbytPayment_Filter
Dim mbytDate_Filter
Dim mbytoptFlag_Filter

'Paging Elements
Dim mlngPageCount,mlngAbsolutePage
Dim mlngMaxRecords

	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	mbytDisplay = Request.Form("optDisplay")
	If len(mbytDisplay) = 0 Then mbytDisplay = 0
	
    Set mclsOrder = New clsOrder
    With mclsOrder
    
	mAction = LoadRequestValue("Action")
	mlngOrderID = Request.Form("OrderID")

	Call .LoadOrderSummaries(SummaryFilter)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<LINK href="ssLibrary/ssStyleSheet.css" type="text/css" rel="stylesheet">
<title>Order Summary</title>
</head>
<body>

<BODY>
<CENTER>

<% Response.Write .OutputSummary %>

</CENTER>
</BODY>
</HTML>
<%
    End With

    
    Set cnn = Nothing
    Response.Flush
%>