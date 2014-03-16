<!--#include File = "../../../SFLib/modEncryption.asp"-->
<!--#include File = "ssOrderAdmin_SupportingSettings.asp"-->
<%
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

Class clsOrder
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pblnError

Private paryCustomValues_OrderManager
Private paryCustomValues_Order
Private prsOrders
Private prsOrderSummaries

Private plngOrderSummaryCount

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
'ssInternalOrderStatus
'ssDateEmailSent

Private plngAltOrderID

Private plngssorderID
Private plngorderStoreID
Private pstrssExternalNotes
Private pstrssInternalNotes
Private pdtssDatePaymentReceived
Private pdtssDateOrderShipped
Private pstrssPaidVia
Private pbytssShippedVia
Private pstrssTrackingNumber
Private pbytssOrderStatus
Private pbytssInternalOrderStatus
Private pdtssDateEmailSent
Private plngPriorShippedOrders

Private pblnssOrderFlagged
Private pblnssExported
Private pblnssExportedPayment
Private pblnssExportedShipping
Private pdtssBackOrderDateNotified
Private pdtssBackOrderDateExpected
Private pstrssBackOrderMessage
Private pstrssBackOrderInternalMessage
Private pstrssBackOrderTrackingNumber

'Order Sent Email Parameters
Private pstrEmailTo
Private pstrEmailSubject
Private pstrEmailBody
Private pblnSendMail

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
    plngOrderSummaryCount = 0
    plngPriorShippedOrders = 0
	Call InitializeCustomValues
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	Call ReleaseObject(prsOrders)
	Call ReleaseObject(prsOrderSummaries)
	Call ReleaseObject(prsOrderTransactions)
End Sub

'***********************************************************************************************
%>
<!--#include file="../ssLibrary/ssmodCommonError.asp"-->
<%
'***********************************************************************************************

	Public Property Get OrderSummaryCount()
	    OrderSummaryCount = plngOrderSummaryCount
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

'***********************************************************************************************

Private Sub InitializeCustomValues

Exit Sub

	ReDim paryCustomValues_Order(0)
	
	'format: Display Text, field name, field value(must be ""), DisplayType, DisplayLength, sqlSource
	paryCustomValues_Order(0) = Array("Gift Message","orderGiftMessage","",enDisplayType_textarea,"","",enDatatype_string)

	ReDim paryCustomValues_OrderManager(0)
	
	'format: Display Text, field name, field value(must be ""), DisplayType, DisplayLength, sqlSource
	paryCustomValues_OrderManager(0) = Array("Ship As Available","ShipAsAvailable","",enDisplayType_checkbox,"","",enDatatype_boolean)

End Sub	'InitializeCustomValues

Private Sub LoadCustomValues(objRS)

Dim i

	If isArray(paryCustomValues_OrderManager) Then
		For i = 0 To UBound(paryCustomValues_OrderManager)
			paryCustomValues_OrderManager(i)(2) = objRS.Fields(paryCustomValues_OrderManager(i)(1)).Value
		Next 'i
	End If
	
	If isArray(paryCustomValues_Order) Then
		For i = 0 To UBound(paryCustomValues_Order)
			paryCustomValues_Order(i)(2) = objRS.Fields(paryCustomValues_Order(i)(1)).Value
		Next 'i
	End If
	
End Sub	'LoadCustomValues

Private Sub UpdateCustomValues(byRef objRS)

Dim i

	If isArray(paryCustomValues_OrderManager) Then
		For i = 0 To UBound(paryCustomValues_OrderManager)
			Call setRSFieldValue(objRS, paryCustomValues_OrderManager(i)(1), paryCustomValues_OrderManager(i)(2))
			'objRS.Fields(paryCustomValues_OrderManager(i)(1)).Value = paryCustomValues_OrderManager(i)(2)
		Next 'i
	End If

End Sub	'UpdateCustomValues

Private Sub UpdateCustomOrderValues(byRef objRS)

Dim i

	If isArray(paryCustomValues_Order) Then
		For i = 0 To UBound(paryCustomValues_Order)
			Call setRSFieldValue(objRS, paryCustomValues_Order(i)(1), paryCustomValues_Order(i)(2))
			'objRS.Fields(paryCustomValues_Order(i)(1)).Value = paryCustomValues_Order(i)(2)
		Next 'i
	End If

End Sub	'UpdateCustomOrderValues

Private Sub LoadCustomValuesFromRequest()

Dim i

	If isArray(paryCustomValues_OrderManager) Then
		For i = 0 To UBound(paryCustomValues_OrderManager)
			paryCustomValues_OrderManager(i)(2) = Trim(Request.Form(paryCustomValues_OrderManager(i)(1)))
		Next 'i
	End If
	
	If isArray(paryCustomValues_Order) Then
		For i = 0 To UBound(paryCustomValues_Order)
			paryCustomValues_Order(i)(2) = Trim(Request.Form(paryCustomValues_Order(i)(1)))
		Next 'i
	End If
	
End Sub	'LoadCustomValuesFromRequest

Public Property Get CustomValues()
    CustomValues = paryCustomValues_OrderManager
End Property

Public Property Get CustomOrderValues()
    CustomOrderValues = paryCustomValues_Order
End Property

'***********************************************************************************************

	Public Property Let orderStoreID(vntValue)
		plngorderStoreID = vntValue
	End Property

	Public Property Get orderStoreID()
	    orderStoreID = plngorderStoreID
	End Property

	Public Property Get ssOrderID()
	    ssOrderID = plngssOrderID
	End Property

	Public Property Get ssExternalNotes()
	    ssExternalNotes = pstrssExternalNotes
	End Property

	Public Property Get ssInternalNotes()
	    ssInternalNotes = pstrssInternalNotes
	End Property

	Public Property Get ssDatePaymentReceived()
	    ssDatePaymentReceived = pdtssDatePaymentReceived
	End Property

	Public Property Get ssDateOrderShipped()
	    ssDateOrderShipped = pdtssDateOrderShipped
	End Property

	Public Property Get ssPaidVia()
	    ssPaidVia = pstrssPaidVia
	End Property

	Public Property Get ssShippedVia()
	    ssShippedVia = pbytssShippedVia
	End Property

	Public Property Get ssTrackingNumber()
	    ssTrackingNumber = pstrssTrackingNumber
	End Property

	Public Property Get ssOrderStatus()
	    ssOrderStatus = pbytssOrderStatus
	End Property
	
	Public Property Get ssInternalOrderStatus()
	    ssInternalOrderStatus = pbytssInternalOrderStatus
	End Property
	
	Public Property Get ssDateEmailSent()
		ssDateEmailSent = pdtssDateEmailSent
	End Property

	Public Property Get ssOrderFlagged()
	    ssOrderFlagged = pblnssOrderFlagged
	End Property

	Public Property Get ssExported()
	    ssExported = pblnssExported
	End Property

	Public Property Get ssExportedPayment()
	    ssExportedPayment = pblnssExportedPayment
	End Property

	Public Property Get ssExportedShipping()
	    ssExportedShipping = pblnssExportedShipping
	End Property

	Public Property Get ssBackOrderDateNotified()
	    ssBackOrderDateNotified = pdtssBackOrderDateNotified
	End Property

	Public Property Get ssBackOrderDateExpected()
	    ssBackOrderDateExpected = pdtssBackOrderDateExpected
	End Property

	Public Property Get ssBackOrderMessage()
	    ssBackOrderMessage = pstrssBackOrderMessage
	End Property

	Public Property Get ssBackOrderInternalMessage()
	    ssBackOrderInternalMessage = pstrssBackOrderInternalMessage
	End Property

	Public Property Get ssBackOrderTrackingNumber()
	    ssBackOrderTrackingNumber = pstrssBackOrderTrackingNumber
	End Property
	
	Public Property Get PriorShippedOrders()
		PriorShippedOrders = plngPriorShippedOrders
	End Property
	
'***********************************************************************************************

	Public Property Get rsOrders()
		If isObject(prsOrders) Then Set rsOrders = prsOrders
	End Property

	Public Property Get rsOrderSummaries()
		If isObject(rsOrderSummaries) Then Set rsOrderSummaries = prsOrderSummaries
	End Property

'***********************************************************************************************

	Private Sub LoadValues

		If rsOrders.EOF Then Exit Sub
		
		With rsOrders
			plngssOrderID = trim(.Fields("OrderID").Value)
			If Len(cstrStoreIDFieldName) > 0 Then plngorderStoreID = trim(.Fields(cstrStoreIDFieldName).Value)
			pstrssExternalNotes = trim(.Fields("ssExternalNotes").Value)
			pstrssInternalNotes = trim(.Fields("ssInternalNotes").Value)
			pdtssDatePaymentReceived = trim(.Fields("ssDatePaymentReceived").Value)
			pdtssDateOrderShipped = trim(.Fields("ssDateOrderShipped").Value)
			pstrssPaidVia = trim(.Fields("ssPaidVia").Value)
			pbytssShippedVia = trim(.Fields("ssShippedVia").Value)
			pstrssTrackingNumber = trim(.Fields("ssTrackingNumber").Value)
			pbytssOrderStatus = trim(.Fields("ssOrderStatus").Value)
			pbytssInternalOrderStatus = trim(.Fields("ssInternalOrderStatus").Value)
			pdtssDateEmailSent = .Fields("ssDateEmailSent").Value

			pblnssOrderFlagged = .Fields("ssOrderFlagged").Value
			pblnssExported= .Fields("ssExported").Value
			pblnssExportedPayment= .Fields("ssExportedPayment").Value
			pblnssExportedShipping= .Fields("ssExportedShipping").Value
			
			pdtssBackOrderDateNotified = .Fields("ssBackOrderDateNotified").Value
			pdtssBackOrderDateExpected = .Fields("ssBackOrderDateExpected").Value
			pstrssBackOrderMessage = .Fields("ssBackOrderMessage").Value
			pstrssBackOrderInternalMessage = .Fields("ssBackOrderInternalMessage").Value
			
			On Error Resume Next
			pstrssBackOrderTrackingNumber = .Fields("ssBackOrderTrackingNumber").Value
			If Err.number <> 0 Then
				Response.Clear
				Call DisplayUpgradeError(Err, "pstrSQL: " & pstrSQL)
				Err.Clear
				Response.Flush
			End If
		End With
		Call LoadCustomValues(rsOrders)

	End Sub 'LoadValues

	Private Sub LoadFromRequest

	    With Request.Form
			plngssorderID = Trim(.Item("orderID"))
			pstrssExternalNotes = Trim(.Item("ssExternalNotes"))
			pstrssInternalNotes = Trim(.Item("ssInternalNotes"))
			pdtssDatePaymentReceived = Trim(.Item("ssDatePaymentReceived"))
			pdtssDateOrderShipped = Trim(.Item("ssDateOrderShipped"))
			pstrssPaidVia = Trim(.Item("ssPaidVia"))
			pbytssShippedVia = Trim(.Item("ssShippedVia"))
			pstrssTrackingNumber = Trim(.Item("ssTrackingNumber"))
			pbytssOrderStatus = Trim(.Item("ssOrderStatus"))
			pbytssInternalOrderStatus = Trim(.Item("ssInternalOrderStatus"))
			
			pstrEmailTo = Trim(.Item("EmailTo"))
			pstrEmailSubject = Trim(.Item("EmailSubject"))
			pstrEmailBody = Trim(.Item("EmailBody"))
			pblnSendMail = Trim(.Item("SendEmail")) = "1"

			pblnssOrderFlagged = Trim(.Item("ssOrderFlagged")) = "1"
			pblnssExported = Trim(.Item("ssExported")) = "1"
			pblnssExportedPayment = Trim(.Item("ssExportedPayment")) = "1"
			pblnssExportedShipping = Trim(.Item("ssExportedShipping")) = "1"
			
			pdtssBackOrderDateNotified = Trim(.Item("ssBackOrderDateNotified"))
			pdtssBackOrderDateExpected = Trim(.Item("ssBackOrderDateExpected"))
			pstrssBackOrderMessage = Trim(.Item("ssBackOrderMessage"))
			pstrssBackOrderInternalMessage = Trim(.Item("ssBackOrderInternalMessage"))
			pstrssBackOrderTrackingNumber = Trim(.Item("ssBackOrderTrackingNumber"))
	    End With
	    Call LoadCustomValuesFromRequest

	End Sub 'LoadFromRequest

	'***********************************************************************************************

	Private Sub DisplayUpgradeError(byRef objError, byVal strMessage)

		Response.Write "<div class='FatalError'>You need to upgrade your database to use Order Manager</div>" _
						& "<h3><a href='ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=OrderManager'>Click here to upgrade</a></h3>"
		Response.Write "<br>" & strMessage
			
	End Sub    'DisplayUpgradeError

	'***********************************************************************************************
	%>
	<!--#include file="ssOrderAdmin_OrderSummary.asp"-->
	<%
	'***********************************************************************************************

	Public Function LoadOutstandingPayments(arySQLParameters)

	Dim pstrSQL
	Dim p_strWhere
	Dim p_strGroupBy
	Dim p_sqlHaving
	Dim p_OrderBy

	On Error Resume Next

		p_strWhere = getValueFromArray(arySQLParameters, 0, "")
		p_strGroupBy = getValueFromArray(arySQLParameters, 1, "")
		p_sqlHaving = getValueFromArray(arySQLParameters, 2, "")
		p_OrderBy = getValueFromArray(arySQLParameters, 3, "")

		'paging routine
		mlngAbsolutePage = LoadRequestValue("AbsolutePage")
		If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
		mlngMaxRecords = LoadRequestValue("PageSize")
		If len(mlngMaxRecords) = 0 Then mlngMaxRecords = clngDefaultRecords	

		pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custFirstName, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderFlagged, ssOrderManager.ssPaidVia" _
				& " FROM sfCShipAddresses RIGHT JOIN (((sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId" _
				& " " & p_strWhere _
				& " GROUP BY sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custEmail, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderStatus, ssOrderManager.ssInternalOrderStatus, ssOrderManager.ssOrderFlagged, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipEmail, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCustomers.custCompany, sfCustomers.custAddr1, sfCustomers.custAddr2, sfOrders.orderIsComplete, sfOrders.orderVoided, sfOrders.orderTradingPartner, ssOrderManager.ssPaidVia " & p_strGroupBy _
				& " " & p_sqlHaving _
				& " " & p_OrderBy

'				& " FROM ((sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
'				& " GROUP BY sfOrders.orderID, sfCustomers.custLastName, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderFlagged, sfOrders.orderIsComplete, sfOrders.orderTradingPartner, ssOrderManager.ssPaidVia " & p_strGroupBy _
		set	prsOrderSummaries = server.CreateObject("adodb.recordset")
		With prsOrderSummaries
	        .CursorLocation = 3 'adUseClient
	'        .CursorType = 3 'adOpenStatic
	        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
	        .LockType = 1 'adLockReadOnly
			If (len(mlngMaxRecords) > 0) AND (mlngMaxRecords <>0) Then 
				.CacheSize = mlngMaxRecords
				.PageSize = mlngMaxRecords
			End If
			.Open pstrSQL, cnn

			If Err.number <> 0 Then
				debugprint "pstrSQL",pstrSQL
				pstrMessage = "Loading Error " & Err.number & ": " & Err.Description
				Err.Clear
				LoadOrderSummaries = False
				Exit Function
			End If
			
			'If mlngMaxRecords = 0 Then mlngMaxRecords = .RecordCount
			plngOrderSummaryCount = .RecordCount
			mlngPageCount = .PageCount
			If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
			
			If Not .EOF Then plngAltOrderID = .Fields("orderID").Value
			
			LoadOutstandingPayments = (Not .EOF)
			
		End With

	End Function    'LoadOutstandingPayments

	'***********************************************************************************************

	Sub ShowOutstandingPayments()
	
	Dim plngID
	Dim i
	
		With Response
			.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' bgcolor='whitesmoke' id='tblOutstandingPayments' rules='none'>"

	    If prsOrderSummaries.RecordCount > 0 Then
	    
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='left'>"
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='right'>"
			.Write "<colgroup align='right'>"
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='center'>"

			.Write "<tr class='tblhdr'>"
			.Write "<TH>Order&nbsp;Number</TH>"
			.Write "<TH>Last&nbsp;Name</TH>"
			.Write "<TH>Items</TH>"
			.Write "<TH>Order&nbsp;Total</TH>"
			.Write "<TH>Order&nbsp;Date&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TH>"
			.Write "<TH>Date&nbsp;Payment&nbsp;Received</TH>"
			.Write "<TH>Paid&nbsp;Via</TH>"
			.Write "</tr>"
			
	        prsOrderSummaries.MoveFirst

	        For i = 1 To prsOrderSummaries.RecordCount
	        
				plngID = trim(prsOrderSummaries("OrderID"))
				.Write " <TR class='Active'>"
	        	.Write "<TD><input type=hidden name=payOrderID value=" & plngID & ">" & plngID & "&nbsp;</TD>"
	        	.Write "<TD>" & prsOrderSummaries.Fields("custLastName").Value & ", " & prsOrderSummaries.Fields("custFirstName").Value & "&nbsp;</TD>"
	        	.Write "<TD>" & prsOrderSummaries.Fields("SumOfodrdtQuantity").Value & "&nbsp;</TD>"
	        	.Write "<TD>" & WriteCurrency(prsOrderSummaries.Fields("orderGrandTotal").Value) & "&nbsp;&nbsp;&nbsp;</TD>"
	        	.Write "<TD>" & FormatDateTime(prsOrderSummaries.Fields("orderDate").Value,1) & "&nbsp;</TD>"
	        	.Write "<TD><input name=payDateID style='HEIGHT: 12pt; font-size: xx-small;' value='" & prsOrderSummaries.Fields("ssDatePaymentReceived").Value & "' ondblclick='this.value=" & Chr(34) & FormatDateTime(Date()) & Chr(34) & "'></TD>"
	        	.Write "<TD><input name=PaidVia style='HEIGHT: 12pt; font-size: xx-small;' value='" & prsOrderSummaries.Fields("ssPaidVia").Value & "' ondblclick='this.value=" & Chr(34) & "Credit Card" & Chr(34) & "'></TD>"
	        	.Write "</TR>"
       	
	            prsOrderSummaries.MoveNext
	        Next
	    Else
				.Write "<TR><TD align=center COLSPAN=6><h3>There are no outstanding payments</h3></TD></TR>"
	    End If
			.Write "</TABLE>"
		End With

	End Sub	'ShowOutstandingPayments

	'***********************************************************************************************

	Sub ShowOutstandingShipments()
	
	Dim plngID
	Dim i
	
		With Response
	
			.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' bgcolor='whitesmoke' id='tblShowOutstandingShipments' rules='none'>"

	    If prsOrderSummaries.RecordCount > 0 Then
	    
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='left'>"
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='right'>"
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='center'>"
			.Write "<colgroup align='center'>"

			.Write "<tr class='tblhdr'>"
			.Write "<TH>&nbsp;Order</TH>"
			.Write "<TH>Last&nbsp;Name</TH>"
			.Write "<TH>Items</TH>"
			.Write "<TH>Order&nbsp;Total</TH>"
			.Write "<TH>Ordered</TH>"
			.Write "<TH>Paid&nbsp;on</TH>"
			.Write "<TH>Shipped&nbsp;On</TH>"
			.Write "<TH>Ship&nbsp;Via</TH>"
			.Write "<TH>Tracking</TH>"
			.Write "<TH>Email</TH>"
			.Write "</tr>"
			
	        prsOrderSummaries.MoveFirst

	        For i = 1 To prsOrderSummaries.RecordCount
	        
				plngID = trim(prsOrderSummaries("OrderID"))
				.Write " <TR class='Active'>"
	        	.Write "<TD>"
	        	.Write "   <input type=hidden name='shipOrderID' value='" & plngID & "'>"
	        	.Write "   <input type=hidden name='isDirty' id='isDirty" & plngID & "' value='0'>"
	        	.Write plngID & "&nbsp;"
	        	.Write "</TD>"
	        	.Write "<TD>" & prsOrderSummaries.Fields("custLastName").Value & ", " & prsOrderSummaries.Fields("custFirstName").Value & "&nbsp;</TD>"
	        	.Write "<TD>" & prsOrderSummaries.Fields("SumOfodrdtQuantity").Value & "&nbsp;</TD>"
	        	.Write "<TD>" & WriteCurrency(prsOrderSummaries.Fields("orderGrandTotal").Value) & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
	        	.Write "<TD>" & FormatDateTime(prsOrderSummaries.Fields("orderDate").Value,2) & "&nbsp;</TD>"
	        	If Len(prsOrderSummaries.Fields("ssDatePaymentReceived").Value & "") > 0 Then
	        		.Write "<TD>" & FormatDateTime(prsOrderSummaries.Fields("ssDatePaymentReceived").Value,2) & "</TD>"
				Else
	        		.Write "<TD>-</TD>"
				End If
	        	.Write "<TD><input name=shipDate size=10 style='HEIGHT: 12pt; font-size: xx-small;' value='" & prsOrderSummaries.Fields("ssDateOrderShipped").Value & "' ondblclick='this.value=this.tag;' tag=" & Chr(34) & FormatDateTime(Date()) & Chr(34) & " onchange='MakeSummaryDirty(" & plngID & ");'></TD>"
	        	.Write "<TD><select name=shipVia size=1 onchange='MakeSummaryDirty(" & plngID & ");'>" & ShipMethodsAsOptions(prsOrderSummaries.Fields("ssShippedVia").Value) & "</select></TD>"
	        	.Write "<TD><input name=shipTrackingNumber size=12 style='HEIGHT: 12pt; font-size: xx-small;' value='" & prsOrderSummaries.Fields("ssTrackingNumber").Value & "' onchange='setShipmentDefaults(" & plngID & "); MakeSummaryDirty(" & plngID & ");'></TD>"
				If Len(prsOrderSummaries.Fields("ssDateEmailSent").Value & "") = 0 Then
	        		.Write "<TD><input type=checkbox name=shipMail." & plngID & " style='HEIGHT: 12pt;' value=1 onchange='MakeSummaryDirty(" & plngID & ");'></TD>"
	        	Else
	        		.Write "<TD><input type=checkbox name=shipMail." & plngID & " style='HEIGHT: 12pt;' value=1 checked onchange='MakeSummaryDirty(" & plngID & ");'></TD>"
	        	End If
	        	.Write "</TR>"

	            prsOrderSummaries.MoveNext
	        Next
	    Else
				.Write "<TR><TD align=center COLSPAN=6><h3>There are no outstanding shipments</h3></TD></TR>"
	    End If

		.Write "</TABLE>"
		End With

	End Sub	'ShowOutstandingShipments

	'***********************************************************************************************

	Public Function LoadOutstandingShipments(arySQLParameters)

	Dim pstrSQL
	Dim p_strWhere
	Dim p_strGroupBy
	Dim p_sqlHaving
	Dim p_OrderBy

	On Error Resume Next

		p_strWhere = getValueFromArray(arySQLParameters, 0, "")
		p_strGroupBy = getValueFromArray(arySQLParameters, 1, "")
		p_sqlHaving = getValueFromArray(arySQLParameters, 2, "")
		p_OrderBy = getValueFromArray(arySQLParameters, 3, "")

		'paging routine
		mlngAbsolutePage = LoadRequestValue("AbsolutePage")
		If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
		mlngMaxRecords = LoadRequestValue("PageSize")
		If len(mlngMaxRecords) = 0 Then mlngMaxRecords = clngDefaultRecords	

		pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custFirstName, Sum(sfOrderDetails.odrdtQuantity) AS SumOfodrdtQuantity, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderFlagged, ssOrderManager.ssShippedVia, ssOrderManager.ssTrackingNumber, ssOrderManager.ssDateEmailSent" _
				& " FROM sfCShipAddresses RIGHT JOIN (((sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCShipAddresses.cshpaddrID = sfOrders.orderAddrId" _
				& " " & p_strWhere _
				& " GROUP BY sfOrders.orderID, sfCustomers.custLastName, sfCustomers.custFirstName, sfCustomers.custEmail, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderStatus, ssOrderManager.ssInternalOrderStatus, ssOrderManager.ssOrderFlagged, sfCShipAddresses.cshpaddrShipLastName, sfCShipAddresses.cshpaddrShipFirstName, sfCShipAddresses.cshpaddrShipEmail, sfCShipAddresses.cshpaddrShipCompany, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCustomers.custCompany, sfCustomers.custAddr1, sfCustomers.custAddr2, sfOrders.orderIsComplete, sfOrders.orderVoided, sfOrders.orderTradingPartner, ssOrderManager.ssShippedVia, ssOrderManager.ssTrackingNumber, ssOrderManager.ssDateEmailSent " & p_strGroupBy _
				& " " & p_sqlHaving _
				& " " & p_OrderBy

'				& " FROM ((sfCustomers INNER JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
'				& " GROUP BY sfOrders.orderID, sfCustomers.custLastName, sfOrders.orderGrandTotal, sfOrders.orderDate, ssOrderManager.ssDatePaymentReceived, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssBackOrderDateExpected, ssOrderManager.ssOrderFlagged, sfOrders.orderIsComplete, sfOrders.orderTradingPartner, ssOrderManager.ssShippedVia, ssOrderManager.ssTrackingNumber, ssOrderManager.ssDateEmailSent " & p_strGroupBy _
		set	prsOrderSummaries = server.CreateObject("adodb.recordset")
		With prsOrderSummaries
	        .CursorLocation = 3 'adUseClient
	'        .CursorType = 3 'adOpenStatic
	        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
	        .LockType = 1 'adLockReadOnly
			If (len(mlngMaxRecords) > 0) AND (mlngMaxRecords <>0) Then 
				.CacheSize = mlngMaxRecords
				.PageSize = mlngMaxRecords
			End If
			.Open pstrSQL, cnn

			If Err.number <> 0 Then
				debugprint "pstrSQL",pstrSQL
				pstrMessage = "Loading Error " & Err.number & ": " & Err.Description
				Err.Clear
				LoadOutstandingShipments = False
				Exit Function
			End If
			
			If mlngMaxRecords = 0 Then mlngMaxRecords = .RecordCount
			plngOrderSummaryCount = .RecordCount
			mlngPageCount = .PageCount
			If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
			
			If Not .EOF Then plngAltOrderID = .Fields("orderID").Value
			
			LoadOutstandingShipments = (Not .EOF)
		End With

	End Function    'LoadOutstandingShipments

	'***********************************************************************************************

	Public Function Load(lngOrderID)

	dim pstrTempID
	dim pstrSQL
	Dim pobjRS

'	On Error Resume Next

		If Len(lngOrderID) = 0 Then
			pstrTempID = plngAltOrderID
		Else
			pstrTempID = lngOrderID
		End If
		
		If Len(pstrTempID) = 0 Then
			Load = False
			Exit Function
		End If
	
		If cblnSF5AE Then
			pstrSQL = "SELECT sfOrders.*, " _
					& "		  sfOrderDetails.*, " _
					& "		  sfOrderAttributes.*, " _
					& "		  sfProducts.prodID, " _
					& "		  sfCPayments.*, " _
					& "		  sfCustomers.*, " _
					& "		  sfCShipAddresses.*, " _
					& "		  ssOrderManager.*, " _
					& "		  sfLocalesCountry1.loclctryName AS shipToCountryName, sfLocalesCountry.loclctryName AS billToCountryName, sfOrdersAE.orderCouponCode, sfOrdersAE.orderBillAmount, sfOrdersAE.orderBackOrderAmount, sfOrdersAE.orderCouponDiscount, sfOrderDetailsAE.odrdtGiftWrapPrice, sfOrderDetailsAE.odrdtGiftWrapQTY, sfOrderDetailsAE.odrdtBackOrderQTY, sfOrderDetailsAE.odrdtAttDetailID" _
					& " FROM (((sfLocalesCountry AS sfLocalesCountry1 RIGHT JOIN (sfLocalesCountry RIGHT JOIN ((((((sfCustomers RIGHT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) LEFT JOIN sfCShipAddresses ON sfOrders.orderAddrId = sfCShipAddresses.cshpaddrID) LEFT JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) ON sfLocalesCountry.loclctryAbbreviation = sfCustomers.custCountry) ON sfLocalesCountry1.loclctryAbbreviation = sfCShipAddresses.cshpaddrShipCountry) LEFT JOIN sfOrdersAE ON sfOrders.orderID = sfOrdersAE.orderAEID) LEFT JOIN sfOrderDetailsAE ON sfOrderDetails.odrdtID = sfOrderDetailsAE.odrdtAEID) LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
					& " WHERE sfOrders.orderID=" & pstrTempID

		Else
			pstrSQL = "SELECT sfOrders.*, " _
					& "       sfOrderDetails.*, " _
					& "       sfOrderAttributes.*, " _
					& "       sfProducts.prodID, " _
					& "       sfCPayments.*, " _
					& "       sfCustomers.*, " _
					& "       sfCShipAddresses.*, " _
					& "       ssOrderManager.*, " _
					& "       sfLocalesCountry1.loclctryName AS shipToCountryName, sfLocalesCountry.loclctryName AS billToCountryName " _
					& " FROM (sfLocalesCountry AS sfLocalesCountry1 RIGHT JOIN (sfLocalesCountry RIGHT JOIN ((((sfCustomers RIGHT JOIN sfOrders ON sfCustomers.custID = sfOrders.orderCustId) LEFT JOIN sfCShipAddresses ON sfOrders.orderAddrId = sfCShipAddresses.cshpaddrID) LEFT JOIN sfCPayments ON sfOrders.orderPayId = sfCPayments.payID) LEFT JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) ON sfLocalesCountry.loclctryAbbreviation = sfCustomers.custCountry) ON sfLocalesCountry1.loclctryAbbreviation = sfCShipAddresses.cshpaddrShipCountry) LEFT JOIN ((sfOrderDetails LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID) LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId) ON sfOrders.orderID = sfOrderDetails.odrdtOrderId" _
					& " WHERE orderID=" & pstrTempID
		End If
		'debugprint "pstrSQL", pstrSQL
		
		Select Case 3
			Case 0:	'Sort by product name only
					pstrSQL = pstrSQL & " Order By odrdtProductName Asc, odrattrOrderDetailId, odrattrID"	'odrattrOrderDetailId necessary to keep the attributes in the right place
			Case 1: 'Sort by product name, attribute name
					pstrSQL = pstrSQL & " Order By odrdtProductName Asc, odrattrOrderDetailId, odrattrName"	'odrattrOrderDetailId necessary to keep the attributes in the right place
			Case Else	'Sort by the order they were added to the cart
					pstrSQL = pstrSQL & " Order By odrdtID, odrattrOrderDetailId"	'odrattrOrderDetailId necessary to keep the attributes in the right place
		End Select
		'debugprint "pstrSQL",pstrSQL 

		Set prsOrders = GetRS(pstrSQL)

		'debugprint "prsOrders.State",prsOrders.State 
		If prsOrders.State <> 1 Then 
			Call ShowStoreFrontVersion
			pstrMessage = "Could not load orders."
			Load = False
		Else
			Call LoadValues
			Load = (Not prsOrders.EOF)
			
			pstrSQL = "SELECT Count(sfOrders.orderID) AS CountOforderID" _
					& " FROM ssOrderManager RIGHT JOIN sfOrders ON ssOrderManager.ssorderID = sfOrders.orderID" _
					& " GROUP BY ssOrderManager.ssDateOrderShipped, sfOrders.orderCustId" _
					& " HAVING (((Count(sfOrders.orderID))>=1) AND ((ssOrderManager.ssDateOrderShipped) Is Not Null) AND (sfOrders.orderCustId=" & prsOrders.Fields("custID").Value & "))"
			Set pobjRS = GetRS(pstrSQL)
			If Not pobjRS.EOF Then
				plngPriorShippedOrders = pobjRS.Fields("CountOforderID").Value
			End If
			Call ReleaseObject(pobjRS)
			
		End If

	End Function    'Load

	'***********************************************************************************************

	Public Function DeleteOrder(lngOrderID)

	Dim sql
	Dim pblnDeletionEnabled

	'On Error Resume Next

		If len(lngOrderID) = 0 Then Exit Function
	
		pblnDeletionEnabled = False
		pblnDeletionEnabled = True

		'Delete any order attributes
		sql = deleteRelatedTableSQL("sfOrderDetails", "odrdtOrderId", "odrdtOrderId", enDatatype_number, lngOrderID, "sfOrderAttributes", "odrattrOrderDetailId")
		cnn.Execute sql, , 128
		If cblnSF5AE Then
			sql = deleteRelatedTableSQL("sfOrderDetails", "odrdtOrderId", "odrdtOrderId", enDatatype_number, lngOrderID, "sfOrderDetailsAE", "odrdtAEID")
			If pblnDeletionEnabled Then cnn.Execute sql, , 128
		End If

		'Delete the credit card payments
		sql = deleteRelatedTableSQL("sfOrders", "orderPayId", "orderID", enDatatype_number, lngOrderID, "sfCPayments", "payID")
		If pblnDeletionEnabled Then cnn.Execute sql, , 128

		'Delete the transaction responses
		sql = "Delete from sfTransactionResponse where trnsrspOrderId=" & wrapSQLValue(lngOrderID, False, enDatatype_number)
		If pblnDeletionEnabled Then cnn.Execute sql, , 128

		'Delete individual order details
		sql = "Delete from sfOrderDetails where odrdtOrderId=" & wrapSQLValue(lngOrderID, False, enDatatype_number)
		If pblnDeletionEnabled Then cnn.Execute sql, , 128
		If cblnSF5AE Then
			sql = "Delete from sfOrdersAE where orderAEID=" & wrapSQLValue(lngOrderID, False, enDatatype_number)
			If pblnDeletionEnabled Then cnn.Execute sql, , 128
		End If
	
		'Can get rid of top level records now
		sql = "Delete from sfOrders where OrderID=" & lngOrderID
		If pblnDeletionEnabled Then cnn.Execute sql, , 128
	
		If (Err.Number = 0) Then
			If pblnDeletionEnabled Then
				pstrMessage = "Order " & lngOrderID & " was successfully deleted."
			Else
				pstrMessage = "Order " & lngOrderID & " was not deleted because this function has been disabled."
			End If
		    DeleteOrder = True
		Else
		    pstrMessage = Err.Description
		    DeleteOrder = False
		End If
		
	End Function	'DeleteOrder

	'***********************************************************************************************

	Public Function checkPaymentMethodChange()

	Dim orderCustId
	Dim orderPayId
	Dim origPaymentType
	Dim pobjRS
	Dim newPaymentType
	
		origPaymentType = LoadRequestValue("origPaymentType")
		newPaymentType = LoadRequestValue("newPaymentType")
		
		If origPaymentType <> newPaymentType Then
			If newPaymentType = "Credit Card" Then
				'orderCustId
				'orderPayId
				Set pobjRS = GetRS("Select orderCustId From sfOrders Where orderID=" & plngssOrderID)
				orderCustId = pobjRS.Fields("orderCustId").Value
				Call ReleaseObject(pobjRS)

				cnn.Execute "Insert Into sfCPayments (payCustId, payCardType, payCardName, payIsActive) Values (" & orderCustId & ", '2', 'New Payment For Order " & plngssOrderID & "', 0)",,128
				Set pobjRS = GetRS("Select payID From sfCPayments Where payCustId=" & orderCustId & " And payCardType='2' And payCardName='New Payment For Order " & plngssOrderID & "' And payIsActive=0")
				orderPayId = pobjRS.Fields("payID").Value
				Call ReleaseObject(pobjRS)
				
				cnn.Execute "Update sfOrders Set orderPayId=" & orderPayId & " Where orderID=" & plngssOrderID,,128
			End If
			
			'Could remove original transaction but let the user do that manually
			'If origPaymentType = "Credit Card" Then
			
			cnn.Execute "Update sfOrders Set orderPaymentMethod='" & newPaymentType & "' Where orderID=" & plngssOrderID,,128
	
		End If	'origPaymentType <> newPaymentType
		
	End Function	'checkPaymentMethodChange
	
	'***********************************************************************************************

	Public Function Update()

	Dim sql
	Dim rs
	Dim strErrorMessage
	Dim blnAdd
	Dim pstrOrigprodID

	'On Error Resume Next

	    pblnError = False
	    Call LoadFromRequest

	    strErrorMessage = ValidateValues
	    If ValidateValues Then
	    
	        If Len(plngssOrderID) = 0 Then plngssOrderID = 0
	        
	        'this added to mark orders complete
	        sql = "Select * from sfOrders where orderID = " & plngssOrderID
	        Set rs = server.CreateObject("adodb.Recordset")
	        rs.open sql, cnn, 1, 3
	        If Not rs.EOF Then
				If Trim(Request.Form("orderIsComplete")) = "1" Then
					rs.Fields("orderIsComplete").Value = 1
				Else
					rs.Fields("orderIsComplete").Value = 0
				End If

				'this added to mark orders void
				If Trim(Request.Form("orderVoided")) = "1" Then
					rs.Fields("orderVoided").Value = 1
				Else
					rs.Fields("orderVoided").Value = 0
				End If
				
				Call UpdateCustomOrderValues(rs)
				rs.Update
				
	        End If

	        sql = "Select * from ssOrderManager where ssOrderID = " & plngssOrderID
	        Set rs = server.CreateObject("adodb.Recordset")
	        rs.open sql, cnn, 1, 3
	        If rs.EOF Then
	            rs.AddNew
	            rs.Fields("ssOrderID").Value = plngssOrderID
	            blnAdd = True
	        Else
	            blnAdd = False
	        End If

			rs.Fields("ssExternalNotes").Value = wrapSQLValue(pstrssExternalNotes, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssInternalNotes").Value = wrapSQLValue(pstrssInternalNotes, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssDatePaymentReceived").Value = wrapSQLValue(pdtssDatePaymentReceived, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssPaidVia").Value = wrapSQLValue(pstrssPaidVia, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssDateOrderShipped").Value = wrapSQLValue(pdtssDateOrderShipped, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssShippedVia").Value = wrapSQLValue(pbytssShippedVia, False, enDatatype_NA)	'NA type used to prevent quote wraps
			If Len(pstrssTrackingNumber) > 100 Then pstrssTrackingNumber = Left(pstrssTrackingNumber, 100)
			rs.Fields("ssTrackingNumber").Value = wrapSQLValue(pstrssTrackingNumber, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssOrderStatus").Value = wrapSQLValue(pbytssOrderStatus, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssInternalOrderStatus").Value = wrapSQLValue(pbytssInternalOrderStatus, False, enDatatype_NA)	'NA type used to prevent quote wraps

	        If Len(pblnssOrderFlagged) <> 0 Then rs.Fields("ssOrderFlagged").Value = Abs(CBool(pblnssOrderFlagged))
	        If Len(pblnssExported) <> 0 Then rs.Fields("ssExported").Value = Abs(CBool(pblnssExported))
	        If Len(pblnssExportedPayment) <> 0 Then rs.Fields("ssExportedPayment").Value = Abs(CBool(pblnssExportedPayment))
	        If Len(pblnssExportedShipping) <> 0 Then rs.Fields("ssExportedShipping").Value = Abs(CBool(pblnssExportedShipping))

			rs.Fields("ssBackOrderDateNotified").Value = wrapSQLValue(pdtssBackOrderDateNotified, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssBackOrderDateExpected").Value = wrapSQLValue(pdtssBackOrderDateExpected, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssBackOrderMessage").Value = wrapSQLValue(pstrssBackOrderMessage, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssBackOrderInternalMessage").Value = wrapSQLValue(pstrssBackOrderInternalMessage, False, enDatatype_NA)	'NA type used to prevent quote wraps
			rs.Fields("ssBackOrderTrackingNumber").Value = wrapSQLValue(pstrssBackOrderTrackingNumber, False, enDatatype_NA)	'NA type used to prevent quote wraps

			If pblnSendMail Then rs.Fields("ssDateEmailSent").Value = Date()

			Call UpdateCustomValues(rs)
			
	        rs.Update
	        
	        Call checkPaymentMethodChange

			If pblnSendMail Then 
				If Request.Form("StockEmail") = 1 Then
					'Call LoadEmailFiles(pstrEmailSubject,pstrEmailBody)
					pstrEmailSubject = ""
					pstrEmailBody = ""
				Else
					pstrEmailSubject = Request.Form("emailSubject")
					pstrEmailBody = Request.Form("emailBody")
				End If
				Call SendEmail(pstrEmailSubject, pstrEmailBody, plngssOrderID, "")
			End If

	        If Err.Number = -2147217887 Then
	            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
	                pstrMessage = "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
	                pblnError = True
	            End If
	        ElseIf Err.Number <> 0 Then
	            pblnError = True
	            pstrMessage = "Error: " & Err.Number & " - " & Err.Description & "<BR>"
	        Else
                pstrMessage = "The changes to " & plngssOrderID & " were successfully saved."
	        End If
	        
	        rs.Close
	        Set rs = Nothing
	        
	    Else
	        pblnError = True
	    End If

	    Update = (not pblnError)

	End Function    'Update
	
	Private Function SendEmail(byVal strEmailSubject, byVal strEmailBody, byVal lngOrderID, byVal strTemplate)

	Dim p_objRS
	Dim pstrSQL
	Dim pstrEmailTo
	Dim pstrEmailPrimary
	Dim pstrEmailSecondary
	
	Dim pstrTempEmailSubject
	Dim pstrTempEmailBody
	
	pstrEmailPrimary = ""
	pstrEmailSecondary = ""

'	On Error Resume Next

		pstrSQL = "SELECT sfOrders.orderID, sfCustomers.custFirstName, sfCustomers.custMiddleInitial, sfCustomers.custLastName, sfCustomers.custEmail, sfCShipAddresses.cshpaddrShipAddr1, sfCShipAddresses.cshpaddrShipAddr2, sfCShipAddresses.cshpaddrShipCity, sfCShipAddresses.cshpaddrShipState, sfCShipAddresses.cshpaddrShipZip, sfCShipAddresses.cshpaddrShipCountry, sfCShipAddresses.cshpaddrShipEmail, ssOrderManager.ssDateOrderShipped, ssOrderManager.ssShippedVia, ssOrderManager.ssTrackingNumber, ssOrderManager.ssBackOrderMessage" _
				& "       FROM (sfCustomers INNER JOIN (sfOrders INNER JOIN ssOrderManager ON sfOrders.orderID = ssOrderManager.ssorderID) ON sfCustomers.custID = sfOrders.orderCustId) INNER JOIN sfCShipAddresses ON sfOrders.orderAddrId = sfCShipAddresses.cshpaddrID" _
				& " WHERE orderID=" & lngOrderID
		If Len(cstrStoreIDFieldName) > 0 Then pstrSQL = Replace(pstrSQL, "sfOrders.orderID,", "sfOrders.orderID,sfOrders." & cstrStoreIDFieldName & ",")

		Set p_objRS = GetRS(pstrSQL)
		If Not p_objRS.EOF Then
			If Len(pstrEmailTo & "") = 0 Then pstrEmailTo = Trim(p_objRS.Fields("custEmail").Value & "")
			If Len(cstrStoreIDFieldName) > 0 Then mclsOrder.orderStoreID = p_objRS.Fields(cstrStoreIDFieldName).Value
			If Len(pstrEmailTo & "") = 0 Then
				SendEmail = False
			Else
				If Len(pstrEmailBody) > 0 Then
					pstrTempEmailSubject = pstrEmailSubject
					pstrTempEmailBody = pstrEmailBody
					pstrTempEmailSubject = customReplacements(pstrEmailSubject, p_objRS, True)
					pstrTempEmailBody = customReplacements(pstrEmailBody, p_objRS, True)
				Else
					Call LoadEmails(strTemplate, pstrTempEmailSubject,pstrTempEmailBody, p_objRS, True)
				End If

				'Prepare string for modified mail routine
				' delimited with |
				'0 - sCustEmail
				'1 - sPrimary	- leave blank to use default, set to - not to send
				'2 - sSecondary	- leave blank to use default, set to - not to send
				'3 - sSubject
				'4 - sMessage

				Call createMail("",pstrEmailTo & "|" & EmailFromAddress(pstrEmailPrimary) & "|" & EmailFromAddress(pstrEmailSecondary) & "|" & pstrTempEmailSubject & "|" & pstrTempEmailBody)

				SendEmail = True
			End If
		Else
			'This will occur if only order tracking information is pressent
			'Response.Write "<font color='red'>This order does not exist in the sfOrders table so no email could be sent.</font></br>"			
			SendEmail = False
		End If
		
		p_objRS.Close
		Set p_objRS = Nothing

	End Function	'SendEmail

	'***********************************************************************************************

	Private Function orderExists(byVal lngOrderID)

	Dim pstrSQL
	Dim pobjRS
	
		pstrSQL = "Select orderID from sfOrders Where orderID = " & wrapSQLValue(lngOrderID, False, enDatatype_number)
		Set pobjRS = GetRS(pstrSQL)
		orderExists = CBool(Not pobjRS.EOF)
		pobjRS.close
		Set pobjRS = Nothing
	
	End Function	'orderExists

	'***********************************************************************************************

	Private Function checkOrderManagerEntryExists(byVal lngOrderID, byVal blnCreateNewRecord)

	Dim pstrSQL
	Dim pobjRS
	Dim pblnExists
	
		pstrSQL = "Select ssOrderID from ssOrderManager Where ssOrderID = " & wrapSQLValue(lngOrderID, False, enDatatype_number)
		Set pobjRS = GetRS(pstrSQL)
		If pobjRS.EOF Then
			If blnCreateNewRecord Then
				pstrSQL = "Insert Into ssOrderManager (ssOrderID) Values (" & wrapSQLValue(lngOrderID, False, enDatatype_number) & ")"
				cnn.Execute pstrSQL,,128
				pblnExists = True
			Else
				pblnExists = False
			End If
		Else
			pblnExists = True
		End If
		pobjRS.close
		Set pobjRS = Nothing
	
		checkOrderManagerEntryExists = pblnExists
	
	End Function	'checkOrderManagerEntryExists

	'***********************************************************************************************

	Public Function UpdatePayments()

	Dim plngID
	Dim sql
	Dim paryOrderIDs
	Dim paryDates
	Dim paryPaidVia
	Dim i
	Dim pdtPaidOn
	Dim pstrPaidVia
	Dim plngRecordsUpdated

	'Initialize the arrays
	paryOrderIDs = Split(Request.Form("payOrderID"),",")
	paryDates = Split(Request.Form("payDateID"),",")
	paryPaidVia = Split(Request.Form("paidVia"),",")
	plngRecordsUpdated = 0

	'On Error Resume Next

	    pblnError = False

		For i = 0 To UBound(paryOrderIDs)
			plngID = getValueFromArray(paryOrderIDs, i, "")
			pdtPaidOn = getValueFromArray(paryDates, i, "")
			pstrPaidVia = getValueFromArray(paryPaidVia, i, "")

			If Len(plngID)>0 And Len(pdtPaidOn)>0 Then
				If checkOrderManagerEntryExists(plngID, True) Then
					sql = "Update ssOrderManager Set" _
						& " ssDatePaymentReceived=" & wrapSQLValue(pdtPaidOn, True, enDatatype_date) & "," _
						& " ssPaidVia=" & wrapSQLValue(pstrPaidVia, True, enDatatype_string) _
						& " Where ssOrderID = " & wrapSQLValue(plngID, False, enDatatype_number)
					cnn.Execute sql,,128
					plngRecordsUpdated = plngRecordsUpdated + 1
				Else
					pblnError = True
				End If
	        End If
		Next	'i

		If pblnError Then
			pstrMessage = "There was an error updating the payments. Please try again."
        Else
			If plngRecordsUpdated = 0 Then
				pstrMessage = "No payments were updated. Please try again."
			ElseIf plngRecordsUpdated = 1 Then
				pstrMessage = "The payment was successfully recorded."
			Else
				pstrMessage = plngRecordsUpdated & " payments were successfully recorded."
			End If
        End If
		
	    UpdatePayments = (not pblnError)

	End Function    'UpdatePayments

	'***********************************************************************************************

	Public Function UpdateShipments()

	Dim sql
	Dim paryOrderIDs
	Dim paryShipDates
	Dim paryShipVia
	Dim paryTracking
	Dim paryDirty
	Dim i
	
	Dim plngID
	Dim pdtShipDate
	Dim pstrShipVia
	Dim pstrTracking
	Dim pblnDirty
	Dim plngRecordsUpdated
	
	'Initialize the arrays
	paryOrderIDs = Split(Request.Form("shipOrderID"),",")
	paryShipDates = Split(Request.Form("shipDate"),",")
	paryShipVia = Split(Request.Form("shipVia"),",")
	paryTracking = Split(Request.Form("shipTrackingNumber"),",")
	paryDirty = Split(Request.Form("isDirty"),",")
	plngRecordsUpdated = 0
	
	'On Error Resume Next
	    pblnError = False

		For i=0 To UBound(paryOrderIDs)
			plngID = getValueFromArray(paryOrderIDs, i, "")
			pdtShipDate = getValueFromArray(paryShipDates, i, "")
			pstrShipVia = getValueFromArray(paryShipVia, i, "")
			pstrTracking = getValueFromArray(paryTracking, i, "")
			pblnDirty = getValueFromArray(paryDirty, i, False)
			If cStr(pblnDirty) = "1" Then pblnDirty = True
			
			If Not isDate(pdtShipDate) Then pdtShipDate = Date()

			If Len(plngID)>0 And pblnDirty Then
				If checkOrderManagerEntryExists(plngID, True) Then
					pblnSendMail = Trim(Request.Form("shipMail." & plngID)) = "1"

					If pblnSendMail Then
						sql = "Update ssOrderManager Set" _
							& " ssDateOrderShipped=" & wrapSQLValue(pdtShipDate, True, enDatatype_date) & "," _
							& " ssShippedVia=" & wrapSQLValue(pstrShipVia, True, enDatatype_number) & "," _
							& " ssTrackingNumber=" & wrapSQLValue(pstrTracking, True, enDatatype_string) & "," _
							& " ssDateEmailSent=" & wrapSQLValue(Now(), True, enDatatype_date) _
							& " Where ssOrderID = " & wrapSQLValue(plngID, False, enDatatype_number)
					Else
						sql = "Update ssOrderManager Set" _
							& " ssDateOrderShipped=" & wrapSQLValue(pdtShipDate, True, enDatatype_date) & "," _
							& " ssShippedVia=" & wrapSQLValue(pstrShipVia, True, enDatatype_number) & "," _
							& " ssTrackingNumber=" & wrapSQLValue(pstrTracking, True, enDatatype_string) _
							& " Where ssOrderID = " & wrapSQLValue(plngID, False, enDatatype_number)
					End If
					
					cnn.Execute sql,,128
					plngRecordsUpdated = plngRecordsUpdated + 1

					If pblnSendMail Then Call SendEmail("", "", plngID, cstrEmailTemplate_Shipment)

				Else
					pblnError = True
				End If
	        End If

		Next

		If pblnError Then
			pstrMessage = "There was an error updating the shipments. Please try again."
        Else
			If plngRecordsUpdated = 0 Then
				pstrMessage = "No shipments were updated. Please try again."
			ElseIf plngRecordsUpdated = 1 Then
				pstrMessage = "The shipment was successfully recorded."
			Else
				pstrMessage = plngRecordsUpdated & " shipments were successfully recorded."
			End If
        End If
			
	    UpdateShipments = (not pblnError)

	End Function    'UpdateShipments

	'***********************************************************************************************

	Public Function ImportShipments()

	Dim sql
	Dim paryOrderIDs
	Dim paryShipDates
	Dim paryShipVia
	Dim paryTracking
	Dim paryDirty
	Dim i
	Dim pblnupdatePayment
	Dim pstrMailMessage
	
	Dim plngID
	Dim pdtShipDate
	Dim pstrShipVia
	Dim pstrTracking
	Dim pblnDirty
	Dim plngRecordsUpdated
	
	'Initialize the arrays
	paryOrderIDs = Split(Request.Form("shipOrderID"),",")
	paryShipDates = Split(Request.Form("shipDate"),",")
	paryShipVia = Split(Request.Form("shipVia"),",")
	paryTracking = Split(Request.Form("shipTrackingNumber"),",")
	paryDirty = Split(Request.Form("isDirty"),",")
	plngRecordsUpdated = 0
	
	pblnupdatePayment = CBool(Request.Form("updatePayment") = "1")

	'On Error Resume Next
	    pblnError = False

		For i=0 To UBound(paryOrderIDs)
			plngID = Trim(getValueFromArray(paryOrderIDs, i, ""))
			pdtShipDate = Trim(getValueFromArray(paryShipDates, i, ""))
			pstrShipVia = Trim(getValueFromArray(paryShipVia, i, ""))
			pstrTracking = Trim(getValueFromArray(paryTracking, i, ""))
			pblnDirty = getValueFromArray(paryDirty, i, False)
			If cStr(pblnDirty) = "1" Then pblnDirty = True

			If Len(plngID)>0 And pblnDirty Then
				If orderExists(plngID) Then
					If checkOrderManagerEntryExists(plngID, True) Then
						pblnSendMail = Trim(Request.Form("shipMail." & plngID)) = "1"

						If pblnSendMail Then
							sql = "Update ssOrderManager Set" _
								& " ssDateOrderShipped=" & wrapSQLValue(pdtShipDate, True, enDatatype_date) & "," _
								& " ssShippedVia=" & wrapSQLValue(pstrShipVia, True, enDatatype_number) & "," _
								& " ssTrackingNumber=" & wrapSQLValue(pstrTracking, True, enDatatype_string) & "," _
								& " ssDateEmailSent=" & wrapSQLValue(Now(), True, enDatatype_date) _
								& " Where ssOrderID = " & wrapSQLValue(plngID, False, enDatatype_number)
						Else
							sql = "Update ssOrderManager Set" _
								& " ssDateOrderShipped=" & wrapSQLValue(pdtShipDate, True, enDatatype_date) & "," _
								& " ssShippedVia=" & wrapSQLValue(pstrShipVia, True, enDatatype_number) & "," _
								& " ssTrackingNumber=" & wrapSQLValue(pstrTracking, True, enDatatype_string) _
								& " Where ssOrderID = " & wrapSQLValue(plngID, False, enDatatype_number)
						End If	'pblnSendMail
						
						cnn.Execute sql,,128
						plngRecordsUpdated = plngRecordsUpdated + 1

						If False Then
							Response.Write "OrderID :" & plngID & "<BR>"
							Response.Write "Shipped Via :" & pstrShipVia & "<BR>"
							Response.Write "Date Order Shipped :" & pdtShipDate & "<BR>"
							Response.Write "Tracking Number :" & Trim(paryTracking(i)) & "<BR>"
							Response.Write "Send Mail :" & pblnSendMail & "<BR>"
							Response.Write "Update Payment :" & pblnupdatePayment & "<BR>"
							Response.Write "<hr>"
							Response.Flush
						End If

						'automatically update payments
						If pblnupdatePayment Then
							sql = "Update ssOrderManager Set" _
								& " ssDateOrderShipped=" & wrapSQLValue(pdtShipDate, True, enDatatype_date) & "," _
								& " ssShippedVia=" & wrapSQLValue(pstrShipVia, True, enDatatype_number) & "," _
								& " ssTrackingNumber=" & wrapSQLValue(pstrTracking, True, enDatatype_string) & "," _
								& " ssDateEmailSent=" & wrapSQLValue(Now(), True, enDatatype_date) _
								& " Where ssOrderID = " & wrapSQLValue(plngID, False, enDatatype_number)
							'cnn.Execute sql,,128
						End If	'pblnupdatePayment
						
						pstrMailMessage = ""
						If pblnSendMail Then 
							If SendEmail("", "", plngID, cstrEmailTemplate_Shipment) Then
								pstrMailMessage = "Email sent.</br>"
								pstrMessage = pstrMessage & "<br/><font color='black'>Order " & plngID & " shipment information updated. Email sent to customer.</font>"
							Else
								pstrMailMessage = "<font color='red'>This order does not exist in the sfOrders table so no email could be sent.</font></br>"
								pstrMessage = pstrMessage & "<br/><font color='black'>Order " & plngID & " shipment information updated.</font> <font color='red'>Error sending email to customer.</font>"
							End If
						Else
							pstrMessage = pstrMessage & "<br/><font color='black'>Order " & plngID & " shipment information updated.</font>"
						End If	'pblnSendMail

					Else
						pstrMessage = pstrMessage & "<br/><font color='red'>There was an error creating the order status record for Order " & plngID & ".</font>"
						pblnError = True
					End If	'checkOrderManagerEntryExists
				Else
					pstrMessage = pstrMessage & "<br/><font color='red'>Order " & plngID & " does not exist in the sfOrders table so it could not be imported.</font>"
					pblnError = True
				End If	'orderExists
	        End If	'Len(plngID)>0 And pblnDirty

		Next

		If pblnError Then
			pstrMessage = "There was an error updating the shipments. Please try again." & pstrMessage
        Else
			If plngRecordsUpdated = 0 Then
				pstrMessage = "No shipments were updated. Please try again." & pstrMessage
			ElseIf plngRecordsUpdated = 1 Then
				pstrMessage = "The shipment was successfully recorded." & pstrMessage
			Else
				pstrMessage = plngRecordsUpdated & " shipments were successfully recorded." & pstrMessage
			End If
        End If
			
	    ImportShipments = (not pblnError)

	End Function    'ImportShipments

	'***********************************************************************************************

	Public Sub OutputSummary()

	'On Error Resume Next

	Dim i
	Dim pstrOrderBy, pstrSortOrder, pstrTempSort
	Dim pstrTitle
	Dim pstrSelect
	Dim pstrHighlight
	Dim pstrTRClass
	Dim pstrTRImage
	Dim pstrID
	Dim pblnSelected
	Dim pblnClosed
	Dim pbytStartPoint
	Dim pbytEndPoint
	
	Dim plngOrderCount: plngOrderCount = 0
	Dim plngItemCount: plngItemCount = 0
	Dim pdblSales: pdblSales = 0

	Dim pstrOrderDate
	Dim pstrssPaymentReceived
	Dim pstrssOrderShipped
	Dim pstrssBackOrderExpected

	Dim aSortHeader(9,4)
	Dim enOutput_Title:				enOutput_Title = 0				' 
	Dim enOutput_ColumnName:		enOutput_ColumnName = 1			' 
	Dim enOutput_ColumnWidth:		enOutput_ColumnWidth = 2		' 
	Dim enOutput_HeaderWidth:		enOutput_HeaderWidth = 3		' 
	Dim enOutput_ColumnAlignment:	enOutput_ColumnAlignment = 4	' 

		With Response

			If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
				pstrTempSort = "descending"
				pstrSortOrder = "ASC"
			Else
				pstrTempSort = "ascending"
				pstrSortOrder = "DESC"
			End If
			
			aSortHeader(1, enOutput_Title) = "Sort by flagged orders"
			aSortHeader(2, enOutput_Title) = "Sort by Order Numbers in " & pstrTempSort & " order"
			aSortHeader(3, enOutput_Title) = "Sort by Last Names in " & pstrTempSort & " order"
			aSortHeader(4, enOutput_Title) = "Sort by Item Quantities in " & pstrTempSort & " order"
			aSortHeader(5, enOutput_Title) = "Sort by Order Totals in " & pstrTempSort & " order"
			aSortHeader(6, enOutput_Title) = "Sort by Order Date in " & pstrTempSort & " order"
			aSortHeader(7, enOutput_Title) = "Sort by Payment Received Dates in " & pstrTempSort & " order"
			aSortHeader(8, enOutput_Title) = "Sort by Order Shipped Dates in " & pstrTempSort & " order"
			aSortHeader(9, enOutput_Title) = "Sort by Back Ordered Dates in " & pstrTempSort & " order"
				
			aSortHeader(1, enOutput_ColumnName) = "&nbsp;"
			aSortHeader(2, enOutput_ColumnName) = "Order Number"
			aSortHeader(3, enOutput_ColumnName) = "Last Name"
			aSortHeader(4, enOutput_ColumnName) = "Items"
			aSortHeader(5, enOutput_ColumnName) = "Order&nbsp;Total"
			aSortHeader(6, enOutput_ColumnName) = "Order Date"
			aSortHeader(7, enOutput_ColumnName) = "Payment Received"
			aSortHeader(8, enOutput_ColumnName) = "Order Shipped"
			aSortHeader(9, enOutput_ColumnName) = "Back Ordered"

			If cblnUseBackOrder Then
				'column header widths
				aSortHeader(0, enOutput_ColumnWidth) = 1
				aSortHeader(1, enOutput_ColumnWidth) = 1
				aSortHeader(2, enOutput_ColumnWidth) = 7
				aSortHeader(3, enOutput_ColumnWidth) = 23
				aSortHeader(4, enOutput_ColumnWidth) = 8
				aSortHeader(5, enOutput_ColumnWidth) = 12
				aSortHeader(6, enOutput_ColumnWidth) = 12
				aSortHeader(7, enOutput_ColumnWidth) = 12
				aSortHeader(8, enOutput_ColumnWidth) = 12
				aSortHeader(9, enOutput_ColumnWidth) = 12

				'scrolling column widths
				aSortHeader(0, enOutput_HeaderWidth) = 1
				aSortHeader(1, enOutput_HeaderWidth) = 1
				aSortHeader(2, enOutput_HeaderWidth) = 8
				aSortHeader(3, enOutput_HeaderWidth) = 23
				aSortHeader(4, enOutput_HeaderWidth) = 3
				aSortHeader(5, enOutput_HeaderWidth) = 9
				aSortHeader(6, enOutput_HeaderWidth) = 21
				aSortHeader(7, enOutput_HeaderWidth) = 13
				aSortHeader(8, enOutput_HeaderWidth) = 11
				aSortHeader(9, enOutput_HeaderWidth) = 10
			Else
				'column header widths
				aSortHeader(0, enOutput_ColumnWidth) = 2
				aSortHeader(1, enOutput_ColumnWidth) = 2
				aSortHeader(2, enOutput_ColumnWidth) = 10
				aSortHeader(3, enOutput_ColumnWidth) = 18
				aSortHeader(4, enOutput_ColumnWidth) = 14
				aSortHeader(5, enOutput_ColumnWidth) = 14
				aSortHeader(6, enOutput_ColumnWidth) = 14
				aSortHeader(7, enOutput_ColumnWidth) = 14
				aSortHeader(8, enOutput_ColumnWidth) = 14
				aSortHeader(9, enOutput_ColumnWidth) = 14

				'scrolling column widths
				aSortHeader(0, enOutput_HeaderWidth) = 3
				aSortHeader(1, enOutput_HeaderWidth) = 9
				aSortHeader(2, enOutput_HeaderWidth) = 9
				aSortHeader(3, enOutput_HeaderWidth) = 20
				aSortHeader(4, enOutput_HeaderWidth) = 14
				aSortHeader(5, enOutput_HeaderWidth) = 10
				aSortHeader(6, enOutput_HeaderWidth) = 19
				aSortHeader(7, enOutput_HeaderWidth) = 14
				aSortHeader(8, enOutput_HeaderWidth) = 14
				aSortHeader(9, enOutput_HeaderWidth) = 14
			End If

			'column alignments
			aSortHeader(0, enOutput_ColumnAlignment) = "center"	'checkbox
			aSortHeader(1, enOutput_ColumnAlignment) = "center"	'flag
			aSortHeader(2, enOutput_ColumnAlignment) = "center"	'order number
			aSortHeader(3, enOutput_ColumnAlignment) = "left"	'last name
			aSortHeader(4, enOutput_ColumnAlignment) = "center"	'items
			aSortHeader(5, enOutput_ColumnAlignment) = "right"	'total
			aSortHeader(6, enOutput_ColumnAlignment) = "center"	'order date
			aSortHeader(7, enOutput_ColumnAlignment) = "center"	'payment received
			aSortHeader(8, enOutput_ColumnAlignment) = "center"	'order shipped
			aSortHeader(9, enOutput_ColumnAlignment) = "center"	'back ordered
			
			If clngScrollableTableHeight <= 0 Then
				For i = 0 To UBound(aSortHeader)
					aSortHeader(i, enOutput_ColumnWidth) = ""	'aSortHeader(i, enOutput_HeaderWidth)
				Next 'i
				aSortHeader(3, enOutput_ColumnWidth) = "100%"	'aSortHeader(i, enOutput_HeaderWidth)
			End If	'

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

			.Write "<table class='tbl' width='100%' cellpadding='3' cellspacing='0' border='0' bgcolor='whitesmoke' id='tblSummary'>"	' rules='none'
			For i = 0 to pbytEndPoint
				If (Not cblnUseOrderFlags) And (i = 1) Then
					'do nothing
				Else
					.Write "  <colgroup align='" & aSortHeader(i, enOutput_ColumnAlignment) & "' width='" & aSortHeader(i, enOutput_ColumnWidth) & "%'>" & vbcrlf
					'.Write "  <colgroup align='" & aSortHeader(i, enOutput_ColumnAlignment) & "' width='" & aSortHeader(i, enOutput_HeaderWidth) & "%'>" & vbcrlf
				End If
			Next 'i
			.Write "	<tr class='tblhdr'>"
			
			'.Write "<TH>&nbsp;</TH>"
			.Write "<TH><input type='checkbox' name='chkCheckAll' id='chkCheckAll1'  onclick='checkAll(theDataForm.chkssOrderID, this.checked);checkAll(theDataForm.chkCheckAll2, this.checked);' value=''></TH>"
			if len(mstrOrderBy) > 0 Then
				pstrOrderBy = mstrOrderBy
			Else
				pstrOrderBy = "1"
			End If
		
			For i = pbytStartPoint to pbytEndPoint
				If cInt(pstrOrderBy) = i Then
					If (pstrSortOrder = "ASC") Then
						.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'DESC');" & chr(34) & _
										" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
										" title='" & aSortHeader(i, enOutput_Title) & "'>" & aSortHeader(i, enOutput_ColumnName) & _
										"&nbsp;&nbsp;&nbsp;<img src='images/up.gif' border=0 align=bottom></TH>" & vbCrLf
					Else
						.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'ASC');" & chr(34) & _
										" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
										" title='" & aSortHeader(i, enOutput_Title) & "'>" & aSortHeader(i, enOutput_ColumnName) & _
										"&nbsp;&nbsp;&nbsp;<img src='images/down.gif' border=0 align=bottom></TH>" & vbCrLf
					End If
				Else
				    .Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'" & pstrSortOrder & "');" & chr(34) & _
									" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
									" title='" & aSortHeader(i, enOutput_Title) & "'>" & aSortHeader(i, enOutput_ColumnName) & "</TH>" & vbCrLf
				End If
			Next 'i
			.Write "	</tr>"
	' 
			'Now for the summary table contents
			If clngScrollableTableHeight > 0 Then
				.Write "<tr><td colspan='" & pbytEndPoint+1 & "'>"
				.Write "<div name='divSummary' style='height:" & clngScrollableTableHeight & "; overflow:scroll;'>"	'
				.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none'" _
						& "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary'" _
						& ">"

				For i = 0 to pbytEndPoint
					If (Not cblnUseOrderFlags) And (i = 1) Then
						'do nothing
					Else
						.Write "<colgroup align='" & aSortHeader(i, enOutput_ColumnAlignment) & "' width='" & aSortHeader(i, enOutput_HeaderWidth) & "%' style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>"
					End If
				Next 'i
			End If	'clngScrollableTableHeight > 0
						
	    If prsOrderSummaries.RecordCount > 0 Then
	        prsOrderSummaries.MoveFirst

	'Need to calculate current recordset page and upper bound to loop through
	dim plnguBound, plnglbound, pstrDisplay

		If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
		If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
		If True Then
			plnglbound = (mlngAbsolutePage - 1) * prsOrderSummaries.PageSize + 1
			plnguBound = mlngAbsolutePage * prsOrderSummaries.PageSize
		Else
			plnglbound = (mlngAbsolutePage - 1) * prsOrderSummaries.PageSize + 1
			plnguBound = mlngAbsolutePage * prsOrderSummaries.PageSize
		End If

		If plnguBound > prsOrderSummaries.RecordCount Then plnguBound = prsOrderSummaries.RecordCount
			prsOrderSummaries.AbsolutePosition = plnglbound
	        For i = plnglbound To plnguBound
				If pstrID <> trim(prsOrderSummaries("OrderID")) Then
					pstrID = trim(prsOrderSummaries("OrderID"))
					pstrTitle = "Click to view " & prsOrderSummaries("OrderID")
					pstrHighlight = "title='" & pstrTitle & "' " _
								  & "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
								  & "onmouseout='doMouseOutRow(this); ClearTitle();' " _
							      & "onmousedown='//ViewOrder(" & chr(34) & pstrID & chr(34) & ");'"

					pstrSelect = "title='" & pstrTitle & "' " _
							   & "onmouseover='DisplayTitle(this);' " _
							   & "onmouseout='ClearTitle();' " _
							   & "onmousedown='ViewOrder(" & chr(34) & pstrID & chr(34) & ");'"
				
					pblnSelected = (pstrID = plngssOrderID)
					pblnClosed = ConvertToBoolean(prsOrderSummaries.Fields("ssDateOrderShipped").Value, False)

					pstrOrderDate = customFormatDateTime(prsOrderSummaries.Fields("OrderDate").Value, 2, "")
					pstrssPaymentReceived = customFormatDateTime(prsOrderSummaries.Fields("ssDatePaymentReceived").Value, 2, "")
					pstrssOrderShipped = customFormatDateTime(prsOrderSummaries.Fields("ssDateOrderShipped").Value, 2, "")
					pstrssBackOrderExpected = customFormatDateTime(prsOrderSummaries.Fields("ssBackOrderDateExpected").Value, 2, "-")
					
					If pblnSelected Then
						pstrTRClass = "Selected"
					Else
						pstrTRClass = maryInternalOrderStatuses(correctEmptyValue(getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssInternalOrderStatus")), 0))(1)
						If Len(pstrTRClass) = 0 Then
							If pblnClosed Then
								pstrTRClass = "Inactive"
							Else
								pstrTRClass = "Active"
							End If
						End If
					End If
					pstrTRImage = maryInternalOrderStatuses(correctEmptyValue(getRSFieldValue_Unknown(prsOrderSummaries.Fields("ssInternalOrderStatus")), 0))(2)
					If Len(pstrTRImage) > 0 Then pstrTRImage = "<img src='" & pstrTRImage & "' border=0 />"

					If pblnSelected Then
					    .Write "<TR class='" & pstrTRClass & "' onmouseover='doMouseOverRow(this);' onmouseout='doMouseOutRow(this);' id='selectedSummaryItem'>"
					Else
						.Write " <TR class='" & pstrTRClass & "' " & pstrHighlight & ">"
					End If
					
					.Write " <TD style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'><input type='checkbox' name='chkssOrderID' id='chkssOrderID' value='" & Server.HTMLEncode(pstrID) & "'" & isChecked(pblnSelected) & ">" & pstrTRImage & "</TD>"
					
					If cblnUseOrderFlags Then 
						If prsOrderSummaries.Fields("ssOrderFlagged").Value = 1 Then
							.Write "<TD style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'><img src='images/MSGBOX03.ICO' alt='x' title='flagged for follow up' height='12'></TD>"
						Else
							.Write "<TD style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>&nbsp;</TD>"
						End If
					End If	'cblnUseOrderFlags
					
	        		.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1' " & pstrSelect & ">&nbsp;&nbsp;<u>" & pstrID & "</u>&nbsp;</TD>"
	        		.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & prsOrderSummaries.Fields("custLastName").Value & ", " & prsOrderSummaries.Fields("custFirstName").Value & "&nbsp;</TD>"
	        		.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & prsOrderSummaries.Fields("SumOfodrdtQuantity").Value & "&nbsp;</TD>"
	        		.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & WriteCurrency(prsOrderSummaries.Fields("orderGrandTotal").Value) & "&nbsp;&nbsp;</TD>"
	        		.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & pstrOrderDate & "&nbsp;</TD>"
	        		.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & pstrssPaymentReceived & "&nbsp;</TD>"
	        		.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & pstrssOrderShipped & "&nbsp;</TD>"
					If cblnUseBackOrder Then .Write "<TD style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & pstrssBackOrderExpected & "&nbsp;</TD>"
	        	
					plngOrderCount = plngOrderCount + 1
					plngItemCount = plngItemCount + prsOrderSummaries.Fields("SumOfodrdtQuantity").Value
					If isNumeric(prsOrderSummaries.Fields("orderGrandTotal").Value) Then pdblSales = pdblSales + prsOrderSummaries.Fields("orderGrandTotal").Value
	
	        		.Write "</TR>" & vbcrlf
	        	End If

	        	'added to display note
	        	If Len(cstrDisplayMemoField) > 0 Then
	        		If Len(prsOrderSummaries.Fields(cstrDisplayMemoField).Value & "") > 0 Then
		        		.Write "<tr><td colspan=3>&nbsp;</td><td COLSPAN='" & pbytEndPoint-2 & "' align=left>" & prsOrderSummaries.Fields(cstrDisplayMemoField).Value & "</td></tr>"		
	        		End If
	        	End If

	            prsOrderSummaries.MoveNext
	        Next

			.Write "<tr><td COLSPAN='" & pbytEndPoint+1 & "' align=center><hr width='100%'></td></tr>"		

			.Write " <TD style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'><input type='checkbox' name='chkCheckAll' id='chkCheckAll2'  onclick='checkAll(theDataForm.chkssOrderID, this.checked);checkAll(theDataForm.chkCheckAll1, this.checked);' value=''></TD>"
			If cblnUseOrderFlags Then 
				.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1' colspan=2 align=center>" & plngOrderCount & " orders</TD>"
			Else
				.Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1' colspan=1 align=center>" & plngOrderCount & " orders</TD>"
			End If
	        .Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>&nbsp;</TD>"
	        .Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & plngItemCount & "</TD>"
	        .Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>" & WriteCurrency(pdblSales) & "&nbsp;</TD>"
	        .Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>&nbsp;</TD>"
	        .Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>&nbsp;</TD>"
			If cblnUseBackOrder Then .Write "<TD style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>&nbsp;</TD>"
	        .Write "<TD nowrap style='border-left-width: 1; border-right-width: 1; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1'>&nbsp;</TD>"

	    Else
				.Write "<TR><TD align=center COLSPAN=6><h3>No orders meet your search criteria.</h3></TD></TR>"
				.Write "<input type='hidden' id='PageSize' name='PageSize' value='" & mlngMaxRecords & "'>"
	    End If
	    
		If clngScrollableTableHeight > 0 Then .Write "</TABLE></div>"

			'Write the paging routine
			.Write "<tr class='tblhdr'><TH COLSPAN='" & pbytEndPoint+1 & "' align=center>"

			If prsOrderSummaries.RecordCount = 0 Then
				.Write "No Orders match your search criteria"
			Elseif prsOrderSummaries.RecordCount = 1 Then
				.Write "1 Order matches your search criteria"
			Else 
				.Write prsOrderSummaries.RecordCount & " Orders match your search criteria<br>"

				dim pstrCheck
				pstrCheck = "return isInteger(this, true, ""Please enter a positive integer for the recordset page size."");"
				.Write "Show&nbsp;<input type='Text' id='PageSize' name='PageSize' value='" & mlngMaxRecords & "' maxlength='4' size='4' style='text-align:center;' onblur='" & pstrCheck & "' ondblclick='this.value=" & prsOrderSummaries.RecordCount & ";'>&nbsp;<a href='' class='tblhdr' onclick='document.frmData.submit(); return false;' title='Set records to show'>orders</a> at a time.&nbsp;&nbsp;"

				If cInt(mlngAbsolutePage) <> 1 And mlngPageCount > 1 Then
					Response.Write "<a href='#' onclick='return ViewPage(" & cInt(mlngAbsolutePage) - 1 & ");'><<</a> "
				End If

				For i=1 to mlngPageCount
					plnglbound = (i-1) * mlngMaxRecords + 1
					plnguBound = i * mlngMaxRecords
					if plnguBound > prsOrderSummaries.RecordCount Then plnguBound = prsOrderSummaries.RecordCount
					
					If cblnShowPageNumbers Then
						pstrDisplay = i & "&nbsp;"
					Else
						pstrDisplay = plnglbound & " - " & plnguBound & "&nbsp;"
					End If
					If i = cInt(mlngAbsolutePage) Then
						Response.Write pstrDisplay
					Else
						Response.Write "<a href='#' onclick='return ViewPage(" & i & ");'>" & pstrDisplay & "</a> "
					End If
				Next
				
				If cInt(mlngAbsolutePage) <> mlngPageCount And mlngPageCount > 1 Then
					Response.Write "<a href='#' onclick='return ViewPage(" & cInt(mlngAbsolutePage) + 1 & ");'>>></a> "
				End If

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
			debugprint "pbytssInternalOrderStatus",pbytssInternalOrderStatus

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
	
'***********************************************************************************************

	Public Function CheckOrderChange()
	
	Dim pblnBaseOrderChanged
	Dim pblnquantityUpdated
	Dim pstrSQL
	Dim plngNumItems
	Dim i
	Dim pstrDeletions
	Dim paryDeletions
	
	'sfOrder changes
	Dim pstrorderAmount
	Dim pstrDiscount
	Dim pstrorderSTax
	Dim pstrorderCTax
	Dim pstrorderShipMethod
	Dim pstrorderShippingAmount
	Dim pstrorderHandling
	Dim pstrorderGrandTotal

	'sfOrder changes
	Dim pstrorderCouponCode
	Dim pstrorderCouponDiscount
	Dim pstrorderBillAmount
	Dim pstrorderBackOrderAmount
	
	'sfOrderDetails changes
	Dim pstrodrdtID
	Dim pstrodrdtProductID
	Dim pstrodrdtProductName
	Dim pstrodrdtQuantity
	Dim pstrunitPrice
	Dim pstrodrdtSubTotal
	Dim pstrBuyersClubPointsIssued
	Dim paryBuyersClubPointsIssued
	
	'sfOrderDetails changes
	Dim paryodrdtID
	Dim paryodrdtProductID
	Dim paryodrdtProductName
	Dim paryodrdtQuantity
	Dim paryunitPrice
	Dim paryodrdtSubTotal
	
	'for inventory changes
	Dim pstrorigQty
	Dim pstrodrdtAttDetailID
	Dim plngQtyDelta
	
	'Dim vItem
	'For Each vItem In Request.Form
	'Response.Write vItem & ": " & Request.Form(vItem) & "<BR>" & vbcrlf
	'Next
	'Response.Flush
	
		pblnBaseOrderChanged = CBool(Len(Trim(Request.Form("baseOrderChanged"))) > 0)
		pblnquantityUpdated = CBool(Trim(Request.Form("baseOrderChanged")) = "1")
	
		If Not pblnBaseOrderChanged Then Exit Function
		'Response.Write "<h3>Updating order</h3>"

		'Update the base order
		pstrorderAmount = Trim(Request.Form("orderAmount"))
		pstrDiscount = Trim(Request.Form("Discount"))
		pstrorderSTax = Trim(Request.Form("orderSTax"))
		pstrorderCTax = Trim(Request.Form("orderCTax"))
		pstrorderShipMethod = Trim(Request.Form("orderShipMethod"))
		pstrorderShippingAmount = Trim(Request.Form("orderShippingAmount"))
		pstrorderHandling = Trim(Request.Form("orderHandling"))
		pstrorderGrandTotal = Trim(Request.Form("orderGrandTotal"))

		If Len(pstrorderAmount) = 0 Then pstrorderAmount = 0
		If CDbl(pstrorderAmount) <> Round(pstrorderAmount, 2) Then pstrorderAmount = Round(pstrorderAmount, 2)
		If Len(pstrDiscount) = 0 Then pstrDiscount = 0
		If Len(pstrorderSTax) = 0 Then pstrorderSTax = 0
		If Len(pstrorderCTax) = 0 Then pstrorderCTax = 0
		If Len(pstrorderShipMethod) = 0 Then pstrorderShipMethod = "-"
		If Len(pstrorderShippingAmount) = 0 Then pstrorderShippingAmount = 0
		If Len(pstrorderHandling) = 0 Then pstrorderHandling = 0
		If Len(pstrorderGrandTotal) = 0 Then pstrorderGrandTotal = 0
	
		If False Then
			Response.Write "pstrorderAmount: " & pstrorderAmount & "<BR>" & vbcrlf
			Response.Write "pstrDiscount: " & pstrDiscount & "<BR>" & vbcrlf
			Response.Write "pstrorderSTax: " & pstrorderSTax & "<BR>" & vbcrlf
			Response.Write "pstrorderCTax: " & pstrorderCTax & "<BR>" & vbcrlf
			Response.Write "pstrorderShipMethod: " & pstrorderShipMethod & "<BR>" & vbcrlf
			Response.Write "pstrorderShippingAmount: " & pstrorderShippingAmount & "<BR>" & vbcrlf
			Response.Write "pstrorderHandling: " & pstrorderHandling & "<BR>" & vbcrlf
			Response.Write "pstrorderGrandTotal: " & pstrorderGrandTotal & "<BR>" & vbcrlf
		End If
	
		pstrSQL = "Update sfOrders Set " _
				& " orderAmount='" & Replace(pstrorderAmount,"'","''") & "', " _
				& " orderSTax='" & Replace(pstrorderSTax,"'","''") & "', " _
				& " orderCTax='" & Replace(pstrorderCTax,"'","''") & "', " _
				& " orderShipMethod='" & Replace(pstrorderShipMethod,"'","''") & "', " _
				& " orderShippingAmount='" & Replace(pstrorderShippingAmount,"'","''") & "', " _
				& " orderHandling='" & Replace(pstrorderHandling,"'","''") & "', " _
				& " orderGrandTotal='" & Replace(pstrorderGrandTotal,"'","''") & "'" _
				& " Where orderID=" & mlngOrderID
		cnn.Execute pstrSQL,,128
	
		'Now update AE
		If cblnSF5AE Then
			pstrorderCouponCode = Trim(Request.Form("orderCouponCode"))
			pstrorderCouponDiscount = Trim(Request.Form("orderCouponDiscount"))
			pstrorderBillAmount = Trim(Request.Form("orderBillAmount"))
			pstrorderBackOrderAmount = Trim(Request.Form("orderBackOrderAmount"))

			If Len(pstrorderCouponCode & pstrorderCouponDiscount) > 0 Then
				pstrSQL = "Select * from sfOrdersAE Where orderAEID=" & mlngOrderID
				Dim pobjsfOrdersAE
				Set pobjsfOrdersAE = Server.CreateObject("ADODB.RECORDSET")
				pobjsfOrdersAE.Open pstrSQL, cnn, 1, 3
				If pobjsfOrdersAE.EOF Then
					pobjsfOrdersAE.AddNew
					pobjsfOrdersAE.Fields("orderAEID").Value = mlngOrderID
				End If

				If Len(pstrorderCouponCode) > 0 Then pobjsfOrdersAE.Fields("orderCouponCode").Value = pstrorderCouponCode
				If Len(pstrorderCouponDiscount) > 0 Then pobjsfOrdersAE.Fields("orderCouponDiscount").Value = pstrorderCouponDiscount
				If Len(pstrorderBillAmount) > 0 Then pobjsfOrdersAE.Fields("orderBillAmount").Value = pstrorderBillAmount
				If Len(pstrorderBackOrderAmount) > 0 Then pobjsfOrdersAE.Fields("orderBackOrderAmount").Value = pstrorderBackOrderAmount

				'pobjsfOrdersAE.Fields("orderBillAmount").Value = pstrorderGrandTotal
				
				pobjsfOrdersAE.Update
				pobjsfOrdersAE.Close
				Set pobjsfOrdersAE = Nothing
			End If
			
			'AE Order details
			Dim pstrodrdtAEID
			Dim pstrodrdtGiftWrapPrice
			Dim pstrodrdtGiftWrapQTY
			Dim pstrodrdtBackOrderQTY

			Dim paryodrdtAEID
			Dim paryodrdtGiftWrapPrice
			Dim paryodrdtGiftWrapQTY
			Dim paryodrdtBackOrderQTY

			pstrodrdtAEID = Trim(Request.Form("odrdtAEID"))
			pstrodrdtGiftWrapPrice = Trim(Request.Form("odrdtGiftWrapPrice"))
			pstrodrdtGiftWrapQTY = Trim(Request.Form("odrdtGiftWrapQTY"))
			pstrodrdtBackOrderQTY = Trim(Request.Form("odrdtBackOrderQTY"))

			If Len(pstrodrdtGiftWrapPrice & pstrodrdtGiftWrapQTY & pstrodrdtBackOrderQTY) > 0 Then
				paryodrdtAEID = Split(pstrodrdtAEID, ",")
				paryodrdtGiftWrapPrice = Split(pstrodrdtGiftWrapPrice, ",")
				paryodrdtGiftWrapQTY = Split(pstrodrdtGiftWrapQTY, ",")
				paryodrdtBackOrderQTY = Split(pstrodrdtBackOrderQTY, ",")
				plngNumItems = UBound(paryodrdtAEID)
			Else
				plngNumItems = -1
			End If
			
			Dim pstrTempBackOrderQTY
			Dim pstrTempGiftWrapPrice
			Dim pstrTempGiftWrapQTY
			
			For i = 0 To plngNumItems
				pstrTempBackOrderQTY = getValueFromArray(paryodrdtBackOrderQTY, i, 0)
				pstrTempGiftWrapPrice = getValueFromArray(paryodrdtGiftWrapPrice, i, 0)
				pstrTempGiftWrapQTY = getValueFromArray(paryodrdtGiftWrapQTY, i, 0)
				
				'Response.Write "pstrTempGiftWrapPrice:" & pstrTempGiftWrapPrice & ":<BR>"
				pstrSQL = "Insert Into sfOrderDetailsAE (odrdtAEID, odrdtGiftWrapPrice, odrdtGiftWrapQTY, odrdtBackOrderQTY) " _
						& " Values (" _
						& " " & wrapSQLValue(paryodrdtAEID(i), False, enDatatype_number) & ", " _
						& " " & wrapSQLValue(pstrTempGiftWrapPrice, False, enDatatype_string) & ", " _
						& " " & wrapSQLValue(pstrTempGiftWrapQTY, False, enDatatype_string) & ", " _
						& " " & wrapSQLValue(pstrTempBackOrderQTY, False, enDatatype_string) & " " _
						& " )"
				On Error Resume Next
				cnn.Execute pstrSQL,,128
				
				If Err.number <> 0 Then
					Err.Clear
					pstrSQL = "Update sfOrderDetailsAE Set " _
							& " odrdtGiftWrapPrice='" & Replace(pstrTempGiftWrapPrice,"'","''") & "', " _
							& " odrdtGiftWrapQTY=" & pstrTempGiftWrapQTY & ", " _
							& " odrdtBackOrderQTY=" & pstrTempBackOrderQTY & " " _
							& " Where odrdtAEID=" & paryodrdtAEID(i)
					cnn.Execute pstrSQL,,128
				End If
			Next 'i

			'Now check for AE order item deletions
			pstrDeletions = Trim(Request.Form("deleteodrdtAEID"))
			paryDeletions = Split(pstrDeletions, ",")
			plngNumItems = UBound(paryDeletions)
			For i = 0 To plngNumItems
				If Len(paryDeletions(i)) > 0 Then
					pstrSQL = "Delete From sfOrderDetailsAE Where odrdtAEID=" & paryDeletions(i)
					'debugprint "pstrSQL", pstrSQL
					cnn.Execute pstrSQL,,128
				End If
			Next 'i
		
		End If	'cblnSF5AE
	
		'sfOrderDetails changes
		pstrodrdtID = Trim(Request.Form("odrdtID"))
		pstrodrdtProductID = Trim(Request.Form("odrdtProductID"))
		pstrodrdtProductName = Trim(Request.Form("odrdtProductName"))
		pstrodrdtQuantity = Trim(Request.Form("odrdtQuantity"))
		pstrunitPrice = Trim(Request.Form("unitPrice"))
		pstrodrdtSubTotal = Trim(Request.Form("odrdtSubTotal"))

		pstrorigQty = Trim(Request.Form("origQty"))
		pstrodrdtAttDetailID = Trim(Request.Form("odrdtAttDetailID"))
		
		'Now split into arrays
		paryodrdtID = Split(pstrodrdtID, ",")
		pstrodrdtProductID = Split(pstrodrdtProductID, ",")
		pstrodrdtProductName = Split(pstrodrdtProductName, ",")
		pstrodrdtQuantity = Split(pstrodrdtQuantity, ",")
		pstrunitPrice = Split(pstrunitPrice, ",")
		pstrodrdtSubTotal = Split(pstrodrdtSubTotal, ",")
		
		pstrBuyersClubPointsIssued = Trim(Request.Form("buyersClubPointsIssued"))
		paryBuyersClubPointsIssued = Split(pstrBuyersClubPointsIssued, ",")

		pstrorigQty = Split(pstrorigQty, ",")
		pstrodrdtAttDetailID = Split(pstrodrdtAttDetailID, ",")
	
		plngNumItems = UBound(paryodrdtID)
		For i = 0 To plngNumItems
			If Len(paryBuyersClubPointsIssued(i)) = 0 Or Not isNumeric(paryBuyersClubPointsIssued(i)) Then paryBuyersClubPointsIssued(i) = 0
			If Instr(1,paryodrdtID(i),"newProduct") > 0 Then
			
				Dim odrdtProductName
				Dim odrdtCategory
				Dim odrdtManufacturer
				Dim odrdtVendor
				Dim pobjRS
				
				If cblnSF5AE Then
					pstrSQL = "SELECT sfProducts.prodName, sfCategories.catName, sfManufacturers.mfgName, sfVendors.vendName" _
							& " FROM sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN (sfCategories RIGHT JOIN sfProducts ON sfCategories.catID = sfProducts.prodCategoryId) ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId" _
							& " WHERE sfProducts.prodID=" & wrapSQLValue(pstrodrdtProductID(i), True, enDatatype_string)
				Else
					pstrSQL = "SELECT sfProducts.prodName, sfCategories.catName, sfManufacturers.mfgName, sfVendors.vendName" _
							& " FROM sfVendors RIGHT JOIN (sfManufacturers RIGHT JOIN (sfCategories RIGHT JOIN sfProducts ON sfCategories.catID = sfProducts.prodCategoryId) ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId" _
							& " WHERE sfProducts.prodID=" & wrapSQLValue(pstrodrdtProductID(i), True, enDatatype_string)
				End If
				
				Set pobjRS = GetRS(pstrSQL)
				If Not pobjRS.EOF Then
					odrdtProductName = Trim(pobjRS.Fields("prodName").Value & "")
					odrdtCategory = Trim(pobjRS.Fields("catName").Value & "")
					odrdtManufacturer = Trim(pobjRS.Fields("mfgName").Value & "")
					odrdtVendor = Trim(pobjRS.Fields("vendName").Value & "")
				End If
				Call ReleaseObject(pobjRS)
				If Len(pstrodrdtProductName(i)) = 0 Then pstrodrdtProductName(i) = odrdtProductName
				
				pstrSQL = "Insert Into sfOrderDetails (odrdtOrderId, odrdtQuantity, odrdtSubTotal, odrdtCategory, odrdtManufacturer, odrdtVendor,odrdtProductName, odrdtPrice, odrdtProductID, buyersClubPointsIssued) " _
						& " Values (" _
						& " " & mlngOrderID & ", " _
						& " " & Trim(pstrodrdtQuantity(i)) & ", " _
						& " " & wrapSQLValue(pstrodrdtSubTotal(i), True, enDatatype_string) & ", " _
						& " " & wrapSQLValue(odrdtCategory, True, enDatatype_string) & ", " _
						& " " & wrapSQLValue(odrdtManufacturer, True, enDatatype_string) & ", " _
						& " " & wrapSQLValue(odrdtVendor, True, enDatatype_string) & ", " _
						& " " & wrapSQLValue(pstrodrdtProductName(i), True, enDatatype_string) & ", " _
						& " " & wrapSQLValue(pstrunitPrice(i), True, enDatatype_string) & ", " _
						& " " & wrapSQLValue(paryBuyersClubPointsIssued(i), True, enDatatype_number) & ", " _
						& " " & wrapSQLValue(pstrodrdtProductID(i), True, enDatatype_string) _
						& " )"
				
				cnn.Execute pstrSQL,,128
			Else
				pstrSQL = "Update sfOrderDetails Set " _
						& " odrdtQuantity=" & Trim(pstrodrdtQuantity(i)) & ", " _
						& " odrdtSubTotal='" & Trim(Replace(pstrodrdtSubTotal(i),"'","''")) & "', " _
						& " odrdtProductName='" & Trim(Replace(pstrodrdtProductName(i),"'","''")) & "', " _
						& " odrdtPrice='" & Trim(Replace(pstrunitPrice(i),"'","''")) & "', " _
						& " buyersClubPointsIssued=" & Replace(paryBuyersClubPointsIssued(i),"'","''") & ", " _
						& " odrdtProductID='" & Trim(Replace(pstrodrdtProductID(i),"'","''")) & "' " _
						& " Where odrdtID=" & paryodrdtID(i)
				cnn.Execute pstrSQL,,128
				
				Call updateDownload_OrderDetail(paryodrdtID(i))
			End If
				
			If cblnSF5AE And cblnUpdateInventoryOnChanges And False Then
				If pstrodrdtQuantity(i) <> pstrorigQty(i) Then
					If Len(pstrorigQty(i)) = 0 Then
						plngQtyDelta = pstrodrdtQuantity(i)
					Else
						plngQtyDelta = pstrodrdtQuantity(i) - pstrorigQty(i)
					End If
					
					plngQtyDelta = plngQtyDelta * -1	'Flip the sign since a return adds to the inventory
					Call updateInventoryQtyDelta(pstrodrdtProductID(i), Replace(pstrodrdtAttDetailID(i), "|", ","), plngQtyDelta)
				End If
			End If	'
			
		Next 'i
	
		'Now check for order item deletions
		pstrDeletions = Trim(Request.Form("deleteodrdtID"))
		paryDeletions = Split(pstrDeletions, ",")
		plngNumItems = UBound(paryDeletions)
		For i = 0 To plngNumItems
			If Len(paryDeletions(i)) > 0 Then
				pstrSQL = "Delete From sfOrderDetails Where odrdtID=" & paryDeletions(i)
				'debugprint "pstrSQL", pstrSQL
				cnn.Execute pstrSQL,,128
				
				pstrSQL = "Delete From sfOrderAttributes Where odrattrOrderDetailId=" & paryDeletions(i)
				cnn.Execute pstrSQL,,128
				
				If cblnSF5AE Then
					pstrSQL = "Delete From sfOrderDetailsAE Where odrdtAEID=" & paryDeletions(i)
					cnn.Execute pstrSQL,,128
				End If

				If cblnSF5AE And cblnUpdateInventoryOnChanges Then
					plngQtyDelta = pstrorigQty(i)
					Call updateInventoryQtyDelta(pstrodrdtProductID(i), Replace(pstrodrdtAttDetailID(i), "|", ","), plngQtyDelta)
				End If	'

			End If
		Next 'i
	
	End Function	'CheckOrderChange

End Class   'clsOrder

'***********************************************************************************************
'***********************************************************************************************

	Function setExportedStatus(byVal strOrderIDs, byVal strExportField, byVal bytExported)
	
	Dim i
	Dim paryOrderIDs
	Dim pblnResult
	Dim pstrSQL
	
		If Len(strOrderIDs) = 0 Then Exit Function
		
		paryOrderIDs = Split(strOrderIDs, ", ")
		For i = 0 To UBound(paryOrderIDs)
			If Err.number <> 0 Then Err.Clear
			
			On Error Resume Next
			pstrSQL = "Insert Into ssOrderManager (ssorderID, " & strExportField & ") Values (" & wrapSQLValue(paryOrderIDs(i), False, enDatatype_number) & "," & wrapSQLValue(bytExported, True, enDatatype_number) & ")" 
			cnn.Execute pstrSQL,,128
			
			If Err.number = 0 Then
				pblnResult = True
			Else
				Err.Clear
				pstrSQL = "Update ssOrderManager Set " & strExportField & "=" & wrapSQLValue(bytExported, True, enDatatype_number) & " Where ssorderID=" & wrapSQLValue(paryOrderIDs(i), False, enDatatype_number)
				cnn.Execute pstrSQL,,128
				If Err.number = 0 Then
					pblnResult = True
				Else
					Err.Clear
					pblnResult = False
				End If
			End If

		Next 'i
		
		setExportedStatus = pblnResult
	
	End Function	'setExportedStatus

'***********************************************************************************************

	Function DecryptCardNumber(ByVal strCC, ByVal blnObfuscate)

	Dim pintCCSeed	'this is here to error early
	Dim pobjCCEncrypt
	Dim pstrCardNumber
	Dim pstrCheckChar
	
		strCC = Trim(strCC & "")
		If Len(strCC) = 0 Then Exit Function
		
		On Error Resume Next
		pintCCSeed = iCC	'iCC is delcared in SFLib/incCC.asp
		If Err.number <> 0 Then
			pintCCSeed = 0
			Err.Clear
		End If
		
		'Now Decrypt the card
		'If the first letter is and E it is SF encryption
		If Len(strCC) > 0 Then pstrCheckChar = Left(strCC, 1)
		If strCC = "****-***-***-****" Then
			pstrCardNumber = strCC
		ElseIf isNumeric(strCC) Then
			pstrCardNumber = strCC
		ElseIf pstrCheckChar <> "E" Then
			pstrCardNumber = EnDeCrypt(strCC, cstrRC4Key)
		Else
			If cblnUseSF505Dll Then
				Set pobjCCEncrypt = Server.CreateObject("SFServer505.CCEncrypt")
				If Err.number <> 0 Then
					Response.Write "<h3><font color=red>It appears you don't have the correct version of the sfServer.dll installed. The latest version is SFServer505. If you're still using a version of StoreFront prior to 50.5 you should alter the <i>Const cblnUseSF505Dll</i> setting in ssOrderAdmin_common.asp</font></h3>"
					Response.Write "Error " & Err.number & ": " & err.Description & "<br />"
					Err.Clear
					Set pobjCCEncrypt = Server.CreateObject("SFServer.CCEncrypt")
					If Err.number <> 0 Then
						Response.Write "<hr>Attempted to use old version of the sfServer.dll<br />"
						Response.Write "<h3><font color=red>It appears you don't have any version of the sfServer.dll installed.</font></h3>"
						Response.Write "Error " & Err.number & ": " & err.Description & "<br />"
						Err.Clear
						Set pobjCCEncrypt = Server.CreateObject("SFServer.CCEncrypt")
					End If
				End If
			Else
				Set pobjCCEncrypt = Server.CreateObject("SFServer.CCEncrypt")
			End If
			
			If Err.number = 0 Then
				pobjCCEncrypt.putSeed(pintCCSeed)
				pstrCardNumber = pobjCCEncrypt.decrypt(CStr(strCC))
				
				If Err.number <> 0 Then Err.Clear
			Else
				Response.Write "<h3><font color=red>It appears you don't have any version of the sfServer.dll installed.</font></h3>"
				pstrCardNumber = strCC
				Err.Clear
			End If
			Set pobjCCEncrypt = Nothing
		End If
		
		If blnObfuscate And Len(pstrCardNumber) >= 4 Then
			pstrCardNumber = "****-****-****-" & Right(pstrCardNumber, 4)
		ElseIf Not isAllowedToViewCC And Len(pstrCardNumber) >= 4 Then
			pstrCardNumber = "****-****-****-" & Right(pstrCardNumber, 4)
		End If
		'If blnObfuscate And Len(pstrCardNumber) >= 4 Then pstrCardNumber = Left(pstrCardNumber, 4) & "-****-****-" & Right(pstrCardNumber, 4)
		
		DecryptCardNumber = pstrCardNumber

	End Function	'DecryptCardNumber

'***********************************************************************************************

	Sub LoadEmails(ByVal strFileName, ByRef strEmailSubject, ByRef strEmailBody, ByRef objRS, ByRef blnComplete)

	Dim i

	'On Error Resume Next

		Call LoadEmailFiles(maryEmails)
		For i = 0 To UBound(maryEmails)
			maryEmails(i)(enEmail_Subject) = customReplacements(maryEmails(i)(enEmail_Subject), objRS, blnComplete)
			maryEmails(i)(enEmail_Body) = customReplacements(maryEmails(i)(enEmail_Body), objRS, blnComplete)

			If CBool(maryEmails(i)(enEmail_FileName) = strFileName) Or CBool((i = UBound(maryEmails)) And (Len(strFileName) = 0)) Then
				strEmailSubject = customReplacements(maryEmails(i)(enEmail_Subject), objRS, blnComplete)
				strEmailBody = customReplacements(maryEmails(i)(enEmail_Body), objRS, blnComplete)
			End If
		Next 'i

	End Sub	'LoadEmails

'***********************************************************************************************

	Function customReplacements(ByVal strSource, ByRef objRS, ByRef blnComplete)

	Dim p_strTemp
	Dim pstrTrackingLink
	Dim pstrCustName
	Dim pstrShipAddr
	Dim pstrShipTime

	'On Error Resume Next

		pstrTrackingLink = orderHistoryURL & "?OrderID=" & Trim(objRS.Fields("orderID").Value) & "&email=" & Trim(objRS.Fields("custEmail").Value)
		pstrCustName = Replace(Trim(objRS.Fields("custFirstName").Value) & " " & Trim(objRS.Fields("custMiddleInitial").Value) & " " & Trim(objRS.Fields("custLastName").Value),"  "," ")
		pstrShipAddr = Trim(objRS.Fields("cshpaddrShipAddr1").Value) & vbcrlf
		If Len(Trim(objRS.Fields("cshpaddrShipAddr2").Value)) > 0 Then pstrShipAddr = pstrShipAddr & Trim(objRS.Fields("cshpaddrShipAddr2").Value)
		pstrShipAddr = pstrShipAddr & Trim(objRS.Fields("cshpaddrShipCity").Value) & ", " & Trim(objRS.Fields("cshpaddrShipState").Value) & " " & Trim(objRS.Fields("cshpaddrShipZip").Value)
			
		p_strTemp = strSource
		p_strTemp = Replace(p_strTemp,"<customerFirstName>",pstrCustName)	' - this is the customer's first name
		p_strTemp = Replace(p_strTemp,"<customerLastName>",pstrCustName)	' - this is the customer's last name
		p_strTemp = Replace(p_strTemp,"<customerMIName>",pstrCustName)	' - this is the customer's middle initial
		p_strTemp = Replace(p_strTemp,"<customerName>",pstrCustName)	' - this is the customer's first name, middle initial, last name
		p_strTemp = Replace(p_strTemp,"<shipAddress>",pstrShipAddr)	' - this is the order ship address
		p_strTemp = Replace(p_strTemp,"<trackingLink>",pstrTrackingLink)	' - this is the link to the customer order history page
		p_strTemp = Replace(p_strTemp,"<orderNumber>",Trim(objRS.Fields("orderID").Value))	' - this is the order number
		p_strTemp = Replace(p_strTemp,"<backorderMessage>",Trim(objRS.Fields("ssBackOrderMessage").Value) & "")	' - this is the order number
		p_strTemp = Replace(p_strTemp,"<customerEmail>",Trim(objRS.Fields("custEmail").Value) & "")	' - this is the order number
		p_strTemp = Replace(p_strTemp,"<recipientEmail>",Trim(objRS.Fields("cshpaddrShipEmail").Value) & "")	' - this is the order number

		If blnComplete Then
			If Len(objRS.Fields("ssDateOrderShipped").Value & "") = 0 Then
				pstrShipTime = FormatDateTime(Date())
			ElseIf isDate(objRS.Fields("ssDateOrderShipped").Value) Then
				pstrShipTime = FormatDateTime(objRS.Fields("ssDateOrderShipped").Value)
			End If
			
			p_strTemp = Replace(p_strTemp,"<dateShipped>", pstrShipTime)	' - this is the date the order was shipped
			p_strTemp = Replace(p_strTemp,"<shipMethod>",ShipIDToName(Trim(objRS.Fields("ssShippedVia").Value)))	' -	this is the shipping carrier used
		End If
		
		customReplacements = p_strTemp

	End Function	'customReplacements

	'***********************************************************************************************
	
	Dim maryEmails

	'***********************************************************************************************
	
	Function getItemFromStoreConfiguration(byVal lngElementID, byVal strDefault)
	
	Dim i
	Dim pstrTemp
	Dim pblnFound:	pblnFound = False
	
		On Error Resume Next
		
		If Err.number <> 0 Then Err.Clear
		
		For i = 0 To UBound(maryStores)
			If CBool(LCase(Trim(mclsOrder.orderStoreID & "")) = CStr(maryStores(i)(enStore_ID))) Then
				pstrTemp = maryStores(i)(lngElementID)
				pblnFound = True
				Exit For
			End If
			
			If Not pblnFound Then
				pstrTemp = maryStores(1)(lngElementID)
			End If
		Next 'i

		If Err.number <> 0 Or Len(pstrTemp) = 0 Then
			pstrTemp = strDefault
			Err.Clear
		End If
		
		getItemFromStoreConfiguration = pstrTemp
			
	End Function	'getItemFromStoreConfiguration

	'***********************************************************************************************
	
	Function emailTemplateDirectory()
		emailTemplateDirectory = getItemFromStoreConfiguration(enStore_EmailDirectory, "emailTemplates/")
	End Function	'emailTemplateDirectory

	'***********************************************************************************************
	
	Function getStoreNameFromID(byRef strStoreID)
		getStoreNameFromID = getItemFromStoreConfiguration(enStore_Name, "")
	End Function	'getStoreNameFromID

	'***********************************************************************************************
	
	Function orderReportsURL()
		orderReportsURL = getItemFromStoreConfiguration(enStore_ReportsURL, cstrWebSite)
	End Function	'orderReportsURL

	'***********************************************************************************************
	
	Function orderHistoryURL()
		orderHistoryURL = getItemFromStoreConfiguration(enStore_OrderHistoryURL, cstrWebSite)
	End Function	'orderHistoryURL

	'***********************************************************************************************
	
	Function PackingSlipURL()
		PackingSlipURL = getItemFromStoreConfiguration(enStore_PackingSlipURL, cstrPackingSlipTemplate)
	End Function	'PackingSlipURL

	'***********************************************************************************************
	
	Function EmailFromAddress(byVal strDefault)
		EmailFromAddress = getItemFromStoreConfiguration(enStore_EmailFrom, strDefault)
	End Function	'EmailFromAddress

	'***********************************************************************************************
	
	Function getStoreIDByOrderID(byVal lngOrderID)
	
	Dim pstrSQL
	Dim pobjRS
	Dim pvntResult
	
		If Len(cstrStoreIDFieldName) = 0 Or Len(lngOrderID) = 0 Then Exit Function
		
		pstrSQL = "Select " & cstrStoreIDFieldName & " From sfOrders Where OrderID=" & lngOrderID
		Set pobjRS = GetRS(pstrSQL)
		pvntResult = pobjRS.Fields(cstrStoreIDFieldName).Value
		pobjRS.Close
		Set pobjRS = Nothing
		
		getStoreIDByOrderID = pvntResult

	End Function	'getStoreIDByOrderID
	
	'***********************************************************************************************
	
	Sub LoadEmailFiles(ByRef aryEmails)

	Dim pobjFSO
	Dim pobjFolder, pobjFiles
	Dim i
	Dim MyFile
	Dim pstrTempLine
	Dim pstrFilePath
	Dim p_strSubject
	Dim p_strBody
	
	On Error Resume Next

		pstrFilePath = Request.ServerVariables("PATH_TRANSLATED")
		pstrFilePath = Replace(Lcase(pstrFilePath),"ssorderadmin.asp","")
		pstrFilePath = Replace(Lcase(pstrFilePath),"ssorderadminpmport.asp","")
		pstrFilePath = pstrFilePath & emailTemplateDirectory

		Set pobjFSO = CreateObject("Scripting.FileSystemObject")
		i = 0
		Set pobjFolder = pobjFSO.GetFolder(pstrFilePath)
		Set pobjFiles = pobjFolder.Files
		ReDim aryEmails(pobjFiles.Count - 1)
		For Each MyFile In pobjFiles
			p_strBody = ""
			aryEmails(i) = Array("fileName", "subject", "body")
			
			aryEmails(i)(enEmail_FileName) = MyFile.Name
			Set MyFile =pobjFSO.OpenTextFile(pstrFilePath & MyFile.Name,1,True)

			p_strSubject = MyFile.ReadLine
			pstrTempLine = MyFile.ReadLine	'garbage line
			pstrTempLine = MyFile.ReadLine & vbcrlf
			Do While pstrTempLine <> "// DO NOT REMOVE THIS LINE //" AND NOT MyFile.AtEndOfStream
				p_strBody = p_strBody & pstrTempLine & vbcrlf
				pstrTempLine = MyFile.ReadLine
			Loop
			
			aryEmails(i)(enEmail_Subject) = p_strSubject
			aryEmails(i)(enEmail_Body) = p_strBody

			MyFile.Close
			Set MyFile = Nothing
			
			i = i + 1
		Next 'MyFile
		Set pobjFiles = Nothing
		Set pobjFolder = Nothing
		
		Set pobjFSO = Nothing

	End Sub	'LoadEmailFiles

	'***********************************************************************************************

	Sub closeObj(objItem)
		ReleaseObject objItem
	End Sub

	'***********************************************************************************************

'***********************************************************************************************
' Added for Gift Certificates
'***********************************************************************************************

'Enumerations
Const enOrderRedemption = 1
Const enGiftCertificate = 2
Const enStoreCredit = 3

Dim mstrCertificate
Dim mdblssCertificateAmount
Dim mdblssGCNewTotalDue

	Function GC_LoadByOrder(lngOrderID, dblGrandTotal)

	Dim pstrSQL
	Dim p_objRS

	'On Error Resume Next

		pstrSQL = "SELECT ssGiftCertificateRedemptions.ssGCRedemptionCGCode, ssGiftCertificateRedemptions.ssGCRedemptionAmount, ssGiftCertificateRedemptions.ssGCRedemptionType, ssGiftCertificateRedemptions.ssGCRedemptionActive" _
				& " FROM ssGiftCertificateRedemptions" _
				& " WHERE ssGCRedemptionType=1 AND ssGiftCertificateRedemptions.ssGCRedemptionOrderID=" & lngOrderID

		Set p_objRS = server.CreateObject("adodb.Recordset")
		p_objRS.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		If Not (p_objRS.EOF Or p_objRS.BOF) Then
			If Trim(p_objRS.Fields("ssGCRedemptionType").Value & "") <> cStr(enGiftCertificate) Then
				mstrCertificate = Trim(p_objRS.Fields("ssGCRedemptionCGCode").Value & "")
				mdblssCertificateAmount = Trim(p_objRS.Fields("ssGCRedemptionAmount").Value & "")
				If isNumeric(mdblssCertificateAmount) Then
					mdblssCertificateAmount = CDbl(mdblssCertificateAmount)
					mdblssGCNewTotalDue = dblGrandTotal + mdblssCertificateAmount
				End If
				GC_LoadByOrder = True
			Else
				GC_LoadByOrder = False
			End If
		Else
			GC_LoadByOrder = False
		End If
		p_objRS.Close
		Set p_objRS = Nothing

	End Function    'GC_LoadByOrder

Dim mclsOrder
%>
