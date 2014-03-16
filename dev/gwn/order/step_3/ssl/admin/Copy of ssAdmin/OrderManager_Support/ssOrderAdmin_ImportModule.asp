<%
'********************************************************************************
'*   Order Manager for StoreFront 6.0
'*   Support File
'*   File Version:		1.00.003
'*   Revision Date:		October 17, 2004
'*
'*   1.00.003 (October 17, 2004)
'*   - Updated to account for csv files wrapped in quotes
'*
'*   1.00.002 (September 28, 2004)
'*   - Updated to skip import for empty rows in import file
'*
'*   1.00.001 (December 26, 2003)
'*   - Initial Release
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<% Sub ShowImportTracking()  %>
<% Dim i %>
<SCRIPT language="vbscript">

Const cblnUseNewShipMethodIDs = True	'Set to False if you have a prior version of Order Manager and used UPS as item 1
Dim maryShipMethods(3)
If cblnUseNewShipMethodIDs Then
	maryShipMethods(0) = Array("UPS","http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=1&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=")
	maryShipMethods(1) = Array("U.S.P.S.","http://www.framed.usps.com/cgi-bin/cttgate/ontrack.cgi?tracknbr=")
Else
	maryShipMethods(1) = Array("UPS","http://wwwapps.ups.com/etracking/tracking.cgi?tracknums_displayed=1&TypeOfInquiryNumber=T&HTMLVersion=4.0&sort_by=status&InquiryNumber1=")
	maryShipMethods(0) = Array("U.S.P.S.","http://www.framed.usps.com/cgi-bin/cttgate/ontrack.cgi?tracknbr=")
End If
maryShipMethods(2) = Array("FedEx","http://www.fedex.com/cgi-bin/tracking?action=track&language=english&cntry_code=us&tracknumbers=")
maryShipMethods(3) = Array("Canada Post","http://204.104.133.7/scripts/tracktrace.dll?MfcIsApiCommand=TraceE&referrer=CPCNewPop&i_num=")

Dim maryCells()
	
	'Tracking Import Settings
	Dim cTrackingColumn_OrderNumber
	Dim cTrackingColumn_ShippingMethod
	Dim cTrackingColumn_TrackingNumber
	Dim cTrackingColumn_Message
	Dim cTrackingColumn_ShipmentDate
	Dim cTrackingColumn_SendEmail
	Dim cstrDefault_CarrierID
	
	'this section mirrors the data entered in the _common file
	Const cTrackingOrderNumberPrefix = "<%= cTrackingOrderNumberPrefix %>"
	Const cblnDefault_SendEmail = <%= cblnDefault_SendEmail %>

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function OpenFile(theFile)
	'Initialize and script ActiveX controls not marked as safe needs to be set to Enable or Promopt

	Dim fso
	Dim MyFile
	Dim pstrFilePath
	Dim pstrFile
	Dim pstrErrorMessage
	Dim pstrTempLine
	
		pstrFilePath = theFile.value
		On Error Resume Next
		
		Set fso = CreateObject("Scripting.FileSystemObject")

		If Err.number = 429 Then
			pstrErrorMessage = "You do not have the security settings set properly for this item. " & vbcrlf _
							& "To enable this functionality do the following: "  & vbcrlf _
							& "  - In the Internet Explorer toolbar select Tools --> Internet Options "  & vbcrlf _
							& "  - Select the security tab "  & vbcrlf _
							& "  - Select Custom Level "  & vbcrlf _
							& "  - Find the option 'Initialize and Script ActiveX Components not marked as safe' "  & vbcrlf _
							& "    Change this setting to Prompt "  & vbcrlf _
							& "  - Select OK and OK "
			msgbox(pstrErrorMessage)
		ElseIf Err.number > 0 Then
			pstrErrorMessage = "There was an error opening the file " & pstrFilePath & ". " & vbcrlf _
							& "Error " & Err.number & ": " & Err.Description & vbcrlf
			msgbox(pstrErrorMessage)
		Else
			Set MyFile = fso.OpenTextFile(pstrFilePath, 1, True)
			Do While Not MyFile.AtEndOfStream
				pstrTempLine = MyFile.ReadLine
				pstrFile = pstrFile & pstrTempLine & vbcrlf
			Loop
			MyFile.close
			Set MyFile = Nothing
		End If

		Set fso = Nothing
		
		OpenFile = pstrFile

	End Function	'OpenFile

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function ClearTrackingNumbers()
	
	Dim theTable
	Dim i
	Dim plngRows
	
		Set theTable = document.all("tblImport")
		plngRows = theTable.rows.length
		If plngRows > 1 Then
			For i = plngRows-1 To 1 Step -1
				theTable.deleteRow(i)
			Next 'i
		End If

	End Function	'ClearTrackingNumbers

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function GetTrackingNumbers(theFile)

	Dim i
	Dim paryShipments
	Dim pobjPackage
	Dim pobjNodeList
	Dim pobjXMLDoc
	Dim pstrFile
	Dim pstrLine
	Dim pstrMailMethod
	Dim pstrOut
	Dim pstrTemp
	
		showLoadingMessage "x"
		pstrFile = OpenFile(theFile)

		'added since Endicia can produce invalid xml
		If Left(pstrFile, 1) = "<" Then
			pstrFile = Replace(pstrFile, "&", "&amp;")
		End If
		
		Set pobjXMLDoc = CreateObject("MSXML.DOMDocument")
		pobjXMLDoc.async = false
		If pobjXMLDoc.LoadXML(pstrFile) Then
			Set pobjNodeList = pobjXMLDoc.selectNodes("DAZzle/Package")
			plngPackages = pobjNodeList.length - 1
			If plngPackages < 0 Then
				Set pobjNodeList = pobjXMLDoc.selectNodes("DAZzleLog/Record")
				plngPackages = pobjNodeList.length - 1

				pstrFile = ""
				For i = 0 To plngPackages
					Set pobjPackage = pobjNodeList.item(i)
					
					pstrMailMethod = pobjPackage.selectSingleNode("MailClass").Text
					pstrTemp = pobjPackage.selectSingleNode("TransactionDateTime").Text
					pstrTemp = Left(pstrTemp, 10)

					'pstrFile = pstrFile & pobjPackage.attributes.item(0).value & "," & pobjPackage.selectSingleNode("PIC").Text & "," & pstrTemp  & vbcrlf
					'pstrLine = pobjPackage.selectSingleNode("TransactionID").Text & "," _
					pstrLine = pobjPackage.selectSingleNode("ReferenceID").Text & "," _
							 & pobjPackage.selectSingleNode("PIC").Text & "," _
							 & pstrTemp & "," _
							 & pstrMailMethod
					
					pstrFile = pstrFile & pstrLine & vbcrlf
					'msgbox "Postmark Date: " & pstrTemp
				Next
			Else
				pstrFile = ""
				For i = 0 To plngPackages
					Set pobjPackage = pobjNodeList.item(i)
					
					pstrTemp = pobjPackage.selectSingleNode("PostmarkDate").Text
					pstrTemp = Mid(pstrTemp, 5, 2) & "/" & Mid(pstrTemp, 7, 2) & "/" & Mid(pstrTemp, 1, 4)

					pstrFile = pstrFile & pobjPackage.attributes.item(0).value& "," & pobjPackage.selectSingleNode("PIC").Text & "," & pstrTemp  & vbcrlf
					'msgbox "Postmark Date: " & pstrTemp
				Next		
			End If
			Set pobjPackage = Nothing
			
			Set pobjNodeList = Nothing
			setElementValue frmData.TrackingImportProfile, "Endicia XML"
			setPredefinedImportTemplates "Endicia XML"
		Else
			'msgbox pstrFile
		End If
		Set pobjXMLDoc = Nothing

		document.frmData.dataFile.value = pstrFile
		paryShipments = Split(pstrFile,vbcrlf)
		document.frmData.dataFile.rows = UBound(paryShipments)
		Call LoadTrackingNumbers(pstrFile)
		
		showLoadingMessage ""

	End Function	'GetTrackingNumbers

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function processCustomDate(strDate, strDefaultDate)
	
	Dim pblnValidDate
	Dim pstrTemp
	Dim pstrYear
	Dim pstrMonth
	Dim pstrDay
	
		If isDate(strDate) Then
			pstrTemp = strDate
		ElseIf Len(strDate) = 0 Then
			pstrTemp = strDefaultDate
		Else
			pblnValidDate = False
			'Try UPS Custom formatting
			If Len(strDate) >= 8 Then
				pstrYear = Left(strDate, 4)
				pstrMonth = Mid(strDate, 5, 2)
				pstrDay = Mid(strDate, 7, 2)
				pstrTemp = pstrMonth & "/" & pstrDay & "/" & pstrYear
				
				pblnValidDate = isDate(pstrTemp)
			Else
			
			End If
			If Not pblnValidDate Then pstrTemp = strDefaultDate
		End If
		
		processCustomDate = pstrTemp
	
	End Function	'processCustomDate

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function LoadTrackingNumbers(pstrFile)

	Dim paryShipments
	Dim paryShipment
	Dim i, j
	Dim pstrCellHTML
	dim pNewRow
	dim pNewCell
	dim ptempDate
	
	Dim plngOrderNumber
	Dim pstrShipVia
	Dim plngShipCarrierID
	Dim pstrShipTrackingNumber
	Dim pstrShipDate
	Dim pstrShipMethod
	Dim pstrSendMail
	Dim pstrDataRow
	Dim pstrDelimeter
	
		Call setPredefinedImportTemplates(getRadio(document.frmData.TrackingImportProfile))
		Call ClearTrackingNumbers
	
		If Instr(1, pstrFile, vbTab) Then
			pstrDelimeter = vbTab
		ElseIf Instr(1, pstrFile, ",") Then
			pstrDelimeter = ","
		Else
			pstrDelimeter = ","
		End If

		paryShipments = Split(pstrFile, vbcrlf)
		
		ptempDate = Date()
		If isNumeric(ptempDate) Then
			ptempDate = ptempDate/1000000
			ptempDate = Left(ptempDate,2) & "/" & Mid(ptempDate,3,2) & "/" & Right(ptempDate,4)
		End If
		
		For i = 0 To UBound(paryShipments)-1
			pstrDataRow = Trim(paryShipments(i))
			If Len(pstrDataRow) > 0 Then
				Set pNewRow = document.all("tblImport").insertRow()
				
				paryShipment = Split(pstrDataRow, pstrDelimeter)
				
				plngOrderNumber = stripQuotes(getArrayValue(paryShipment, cTrackingColumn_OrderNumber, ""))
				plngOrderNumber = Replace(plngOrderNumber, cTrackingOrderNumberPrefix, "")
				pstrShipTrackingNumber = stripQuotes(getArrayValue(paryShipment, cTrackingColumn_TrackingNumber, ""))

				pstrShipMethod = stripQuotes(getArrayValue(paryShipment, cTrackingColumn_ShippingMethod, ""))
				plngShipCarrierID = ShipCodeToCarrierID(pstrShipMethod)
				If Len(plngShipCarrierID) = 0 Then plngShipCarrierID = cstrDefault_CarrierID
				pstrShipVia = ShipIDToName(plngShipCarrierID)
				
				pstrShipDate = stripQuotes(getArrayValue(paryShipment, cTrackingColumn_ShipmentDate, ""))
				pstrShipDate = processCustomDate(pstrShipDate, ptempDate)
				
				pstrSendMail = stripQuotes(getArrayValue(paryShipment, cTrackingColumn_SendEmail, ""))
				If Len(pstrSendMail) = 0 Then pstrSendMail = cblnDefault_SendEmail
				
				On Error Resume Next
				If CBool(pstrSendMail) Then
					pstrSendMail = " checked"
				Else
					pstrSendMail = ""
				End If
				If err.number <> 0 Then Err.Clear
				On Error Goto 0
				
				If Len(plngOrderNumber) > 0 And isNumeric(plngOrderNumber) Then
					For j = 0 To 5
						Select Case j
							Case 0: pstrCellHTML = "<!--<INPUT type=checkbox id=Update name=Update." & plngOrderNumber & " value=1 checked>-->" _
												& "<INPUT type=hidden id=shipOrderID name=shipOrderID value='" & plngOrderNumber & "'>" _
												& "<INPUT type=hidden id=isDirty name=isDirty value='1'>" _
												& "<INPUT type=hidden id=shipVia name=shipVia value='" & plngShipCarrierID & "'>" _
												& "<INPUT type=hidden id=shipTrackingNumber name=shipTrackingNumber value='" & pstrShipTrackingNumber & "'>" _
												& "<INPUT type=hidden id=shipDate name=shipDate value='" & pstrShipDate & "'>"
							Case 1: pstrCellHTML = "<!--<INPUT type=checkbox id=shipMail name=shipMail." & plngOrderNumber & " value='1'" & pstrSendMail & ">-->"
							Case 2: pstrCellHTML = plngOrderNumber
							Case 3: pstrCellHTML = pstrShipVia & " (" & pstrShipMethod & ")"
							Case 4: pstrCellHTML = pstrShipTrackingNumber
							Case 5: pstrCellHTML = pstrShipDate
						End Select
						Set pNewCell = pNewRow.insertCell()
						pNewCell.innerHTML = pstrCellHTML
					Next 'j
				End If
			End If	'Len(pstrDataRow) > 0
		Next 'i

	End Function	'LoadTrackingNumbers

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function GetEmail(theFile)

	Dim pstrFile

		pstrFile = OpenFile(theFile)
		document.frmData.emailBody.value = pstrFile

	End Function	'GetEmail

	'--------------------------------------------------------------------------------------------------

	Function getArrayValue(byRef ary, byRef lngIndex, byRef vntDefaultToUse)
		If isArray(ary) Then
			If UBound(ary) >= lngIndex And lngIndex > -1 Then
				getArrayValue = Trim(ary(lngIndex) & "")
			Else
				getArrayValue = vntDefaultToUse
			End If
		Else
			If Len(ary) > 0 Then
				getArrayValue = ary
			Else
				getArrayValue = vntDefaultToUse
			End If
		End If
	End Function	'getArrayValue

	'--------------------------------------------------------------------------------------------------

	Sub setPredefinedImportTemplates(byVal strTemplate)
	
	Dim pstrEmailTemplateToUse

		pstrEmailTemplateToUse = "<%= cstrDefaultOrderShippedEmailTemplate %>"
		
		Select Case strTemplate
		<%
			Response.Write vbcrlf
			For i = 1 To UBound(maryImportTemplates)
				Response.Write "			Case " & Chr(34) & maryImportTemplates(i)(0) & Chr(34) & vbcrlf
				Response.Write "				cstrDefault_CarrierID = " & maryImportTemplates(i)(1) & vbcrlf
				Response.Write "				cTrackingColumn_OrderNumber = " & maryImportTemplates(i)(2) & vbcrlf
				Response.Write "				cTrackingColumn_TrackingNumber = " & maryImportTemplates(i)(3) & vbcrlf
				Response.Write "				cTrackingColumn_ShipmentDate = " & maryImportTemplates(i)(4) & vbcrlf
				Response.Write "				cTrackingColumn_ShippingMethod = " & maryImportTemplates(i)(5) & vbcrlf
				Response.Write "				cTrackingColumn_Message = " & maryImportTemplates(i)(6) & vbcrlf
				Response.Write "				cTrackingColumn_SendEmail = " & maryImportTemplates(i)(7) & vbcrlf
				Response.Write "				pstrEmailTemplateToUse = " & Chr(34) & maryImportTemplates(i)(8) & Chr(34) & vbcrlf
			Next 'i
			
			Response.Write "			Case Else" & vbcrlf
			Response.Write "				cstrDefault_CarrierID = " & maryImportTemplates(0)(1) & vbcrlf
			Response.Write "				cTrackingColumn_OrderNumber = " & maryImportTemplates(0)(2) & vbcrlf
			Response.Write "				cTrackingColumn_TrackingNumber = " & maryImportTemplates(0)(3) & vbcrlf
			Response.Write "				cTrackingColumn_ShipmentDate = " & maryImportTemplates(0)(4) & vbcrlf
			Response.Write "				cTrackingColumn_ShippingMethod = " & maryImportTemplates(0)(5) & vbcrlf
			Response.Write "				cTrackingColumn_Message = " & maryImportTemplates(0)(6) & vbcrlf
			Response.Write "				cTrackingColumn_SendEmail = " & maryImportTemplates(0)(7) & vbcrlf
			Response.Write "				pstrEmailTemplateToUse = " & Chr(34) & maryImportTemplates(0)(8) & Chr(34) & vbcrlf
		%>
		End Select
		letSelectValue document.frmData.emailFile, pstrEmailTemplateToUse
		//msgbox strTemplate & ": " & pstrEmailTemplateToUse
		
		document.frmData.cTrackingColumn_OrderNumber.value = cTrackingColumn_OrderNumber
		document.frmData.cTrackingColumn_ShippingMethod.value = cTrackingColumn_ShippingMethod
		document.frmData.cTrackingColumn_TrackingNumber.value = cTrackingColumn_TrackingNumber
		document.frmData.cTrackingColumn_ShipmentDate.value = cTrackingColumn_ShipmentDate
		document.frmData.cTrackingColumn_Message.value = cTrackingColumn_Message
		document.frmData.cTrackingColumn_SendEmail.value = cTrackingColumn_SendEmail
		document.frmData.cstrDefault_CarrierID.value = cstrDefault_CarrierID
		
		'Figure out the max column
		Dim plngNumColumns
		
		plngNumColumns = 0
		If plngNumColumns < cTrackingColumn_OrderNumber Then plngNumColumns = cTrackingColumn_OrderNumber
		If plngNumColumns < cTrackingColumn_ShippingMethod Then plngNumColumns = cTrackingColumn_ShippingMethod
		If plngNumColumns < cTrackingColumn_TrackingNumber Then plngNumColumns = cTrackingColumn_TrackingNumber
		If plngNumColumns < cTrackingColumn_ShipmentDate Then plngNumColumns = cTrackingColumn_ShipmentDate
		If plngNumColumns < cTrackingColumn_Message Then plngNumColumns = cTrackingColumn_Message
		If plngNumColumns < cTrackingColumn_SendEmail Then plngNumColumns = cTrackingColumn_SendEmail

		'now write out the keys
		
		Dim i
		ReDim paryKeys(plngNumColumns)
		Dim pstrTemp
		For i = 0 To plngNumColumns - 1
			paryKeys(i) = ""
		Next 'i
		For i = 0 To plngNumColumns - 1
			Select Case i
				Case cTrackingColumn_OrderNumber:		paryKeys(i) = "Order Number (" & i + 1 & ")"
				Case cTrackingColumn_ShippingMethod:	paryKeys(i) = "Shipping Method (" & i + 1 & ")"
				Case cTrackingColumn_TrackingNumber:	paryKeys(i) = "Tracking Number (" & i + 1 & ")"
				Case cTrackingColumn_ShipmentDate:		paryKeys(i) = "Shipment Date (" & i + 1 & ")"
				Case cTrackingColumn_Message:			paryKeys(i) = "Message (" & i + 1 & ")"
				Case cTrackingColumn_SendEmail:			paryKeys(i) = "Send Email (" & i + 1 & ")"
				Case Else:								paryKeys(i) = "-"
			End Select
		Next 'i
		paryKeys(i) = ShipIDToName(cstrDefault_CarrierID) & " (" & cstrDefault_CarrierID & ")"

		pstrTemp = "<table border=1 cellpadding=2 cellspacing=0 class=tbl><tr class=tblhdr><th>Item</th><th>Column</th></tr>" _
				 & "<tr><td>Order Number</td><td>" & CStr(cTrackingColumn_OrderNumber + 1) & "</td>" _
				 & "<tr><td>ShippingMethod</td><td>" & CStr(cTrackingColumn_ShippingMethod + 1) & "</td>" _
				 & "<tr><td>TrackingNumber</td><td>" & CStr(cTrackingColumn_TrackingNumber + 1) & "</td>" _
				 & "<tr><td>ShipmentDate</td><td>" & CStr(cTrackingColumn_ShipmentDate + 1) & "</td>" _
				 & "<tr><td>Default Carrier to use</td><td>" & ShipIDToName(cstrDefault_CarrierID) & "</td>" _
				 & "</table>"

		'		 & "Message: " & CStr(cTrackingColumn_Message + 1) & "<br />" _
		'		 & "SendEmail: " & CStr(cTrackingColumn_SendEmail + 1) & "<br />" _
		'pstrTemp = pstrTemp & "Elements: " & CStr(plngNumColumns + 1) & "<br />" & paryKeys(0)
		'For i = 1 To plngNumColumns - 1
		'	pstrTemp = pstrTemp & ", " & paryKeys(i)
		'Next 'i
		'pstrTemp = pstrTemp & "<br />Default Carrier to use: " & paryKeys(plngNumColumns)
		document.all("divFileKeys").innerHTML = pstrTemp

	End Sub	'setPredefinedImportTemplates

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function ShipIDToName(lngID)
		If lngID <= UBound(maryShipMethods) Then ShipIDToName = maryShipMethods(lngID)(0)
	End Function	'ShipIDToName

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function ShipNameToID(strName)
	
	Dim i
	
        For i=0 to UBound(maryShipMethods)
			If Trim(strName) = Trim(maryShipMethods(i)(0)) Then
				ShipNameToID = i
				Exit Function
			End If
        Next
        ShipNameToID = -1
	End Function	'ShipNameToID

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function ShipCodeToCarrierID(strShipMethod)
	
	Dim pbytCarrierID

		Select Case strShipMethod
			Case "1DM","1DA","1DAPI","1DP","2DM","2DA","3DS","GND","STD","XPR","XDM","XPD","Next Day Air", _
				 "UPS"
				pbytCarrierID = ShipNameToID("UPS")
			Case "Express", _
				 "Priority", _
				 "Parcel", _
				 "First Class", _
				 "Global Express Guaranteed (GXG) Document Service", _
				 "Global Express Guaranteed (GXG) Non-Document Service", _
				 "Global Express Mail (EMS)", _
				 "Global Priority Mail - Flat-rate Envelope (large):", _
				 "Global Priority Mail - Flat-rate Envelope (small)", _
				 "Global Priority Mail - Variable Weight Envelope (single)", _
				 "Airmail Letter Post", _
				 "Airmail Parcel Post", _
				 "Economy (Surface) Letter Post", _
				 "Economy (Surface) Parcel Post", _
				 "USPS", "U.S.P.S."
				pbytCarrierID = ShipNameToID("U.S.P.S.")
			Case "U.S. Domestic FedEx Home Delivery Package", "FedEx"
				pbytCarrierID = ShipNameToID("FedEx")
			Case "01","03","05","06","20","70","80","83","86"	'FDXE
				pbytCarrierID = ShipNameToID("FedEx")
			Case "01i","03i","06i","70i","86i"	'FDXE - international
				pbytCarrierID = ShipNameToID("FedEx")
			Case "92i"	'FDXG
				pbytCarrierID = ShipNameToID("FedEx")
			Case "90","92"	'FDXG
				pbytCarrierID = ShipNameToID("FedEx")
			Case "Canada Post"
				pbytCarrierID = ShipNameToID("Canada Post")
			Case Else
				pbytCarrierID = -1
				pbytCarrierID = 0
		End Select
		
		ShipCodeToCarrierID = pbytCarrierID
		
	End Function	'ShipCodeToCarrierID

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Sub SetImportTable

		dim pNewRow
		dim pbytRowCount

		Set pNewRow = document.all("tblImport").Rows(1)
		pbytRowCount = pNewRow.Cells.length-1
		ReDim maryCells(pbytRowCount)
		For i = 0 to pbytRowCount
			maryCells(i) = pNewRow.Cells(i).innerHTML
		Next 'i
		
	End Sub	'SetImportTable

	'--------------------------------------------------------------------------------------------------

	Function stripQuotes(byVal strValue)

	Dim pstrOut

		pstrOut = strValue
		If Left(pstrOut, 1) = Chr(34) Then pstrOut = Right(pstrOut, Len(pstrOut) - 1)
		If Right(pstrOut, 1) = Chr(34) Then pstrOut = Left(pstrOut, Len(pstrOut) - 1)
		stripQuotes = pstrOut

	End Function	'stripQuotes

</SCRIPT>
<script LANGUAGE="JavaScript">

//tipMessage[...]=[title,text]
tipMessage['importStep1']=["Step 1. Select Tracking Numbers file", "Select a tracking number file from your hard drive. You may also paste the contents of the file in the <em>Data File</em> section"]
tipMessage['importStep2']=["Step 2. Select Tracking Numbers file format", "Pick a format the import file is in. If a data file has been selected the results will automatically appear"]
tipMessage['importStep3']=["Step 3. Click to review", "Click here to preview what will be imported."]
tipMessage['importStep4']=["Step 4. Import tracking numbers", "Click here to upload the tracking information to the database"]

tipMessage['updateShipment']=["Import Options", "Check to set orders as shipped after importing the tracking numbers."]
tipMessage['updatePayment']=["Import Options", "Check to set orders as paid after importing the tracking numbers."]
tipMessage['sendEmail']=["Import Options", "Check to send <em>Order Shipped</em> email. The email template will be that selected below."]
tipMessage['saveMessageBody']=["Import Options", "Check to save email body sent to the tracking message field. Note: You must check the send email option to use this option"]
tipMessage['emailFile']=["Import Options", "Select the email template to use for this mailing"]
tipMessage['lblupdateShipment']=tipMessage['updateShipment']
tipMessage['lblupdatePayment']=tipMessage['updatePayment']
tipMessage['lblsendEmail']=tipMessage['sendEmail']
tipMessage['lblsaveMessageBody']=tipMessage['saveMessageBody']
tipMessage['lblemailFile']=tipMessage['emailFile']
</script>

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none" id="tblImportTracking">
  <colgroup>
	<col align="left" valign="top" />
	<col align="left" valign="top" />
	<col align="left" valign="top" width="100%" />
  </colgroup>
  <TR>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<TD>
	<TABLE class="tbl" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblImport" width="100%">
		<colgroup>
			<col align="center" />
			<col align="center" />
			<col align="center" />
			<col align="center" />
			<col align="center" />
			<col align="center" />
		</colgroup>
      <TR class="tblhdr">
        <TD><!--Update-->&nbsp;</TD>
        <TD><!--Send email--></TD>
        <TD>Order ID</TD>
        <TD>Shipping Method</TD>
        <TD>Tracking Number</TD>
        <TD>Shipping Date</TD>
      </TR>
	</TABLE>
	</td>
  </TR>
  <tr>
	<td><span id="importStep1" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Step&nbsp;1.</span>&nbsp;</td>
    <td>Select Tracking Numbers file:&nbsp;</td>
    <td>
	  <input type=file id="trackingNumbersFile" name="trackingNumbersFile" onchange="javascript:GetTrackingNumbers(this);" size="20">
    </td>
  </tr>
  <tr>
	<td><span id="importStep2" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Step&nbsp;2.</span>&nbsp;</td>
    <td>Select Tracking Numbers file format:&nbsp;</td>
    <td>
		<% For i = 0 To UBound(maryImportTemplates) %>
		<input type="radio" name="TrackingImportProfile" id="TrackingImportProfile<%= i %>" value="<%= maryImportTemplates(i)(0) %>" <%= isChecked(maryImportTemplates(i)(0) = cstrDefault_ImportProfile) %> onclick="LoadTrackingNumbers(frmData.dataFile.value)">&nbsp;<label for="TrackingImportProfile<%= i %>" onclick="LoadTrackingNumbers(frmData.dataFile.value)"><%= maryImportTemplates(i)(0) %></label><br />
		<% Next 'i %>
    </td>
  </tr>
  <tr>
	<td><span id="importStep3" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Step 3.</span>&nbsp;</td>
    <td>Click to review:&nbsp;</td>
    <td><div style="cursor: hand;" id="divPreview" onclick="LoadTrackingNumbers(frmData.dataFile.value)" title="Click to preview tracking what information will be imported from tracking file.">Preview</div></td>
  </tr>
  <tr>
	<td colspan="2">&nbsp;</td>
    <td>
	  <input type="checkbox" id="updateShipment" name="updateShipment" value="1" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();" checked>&nbsp;<label for="updateShipment" id="lblupdateShipment" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Automatically mark order as shipped</label><br />
	  <input type="checkbox" id="updatePayment" name="updatePayment" value="1" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();" checked>&nbsp;<label for="updatePayment" id="lblupdatePayment" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Automatically update unrecorded payments</label><br />
	  <input type="checkbox" id="sendEmail" name="sendEmail" value="1" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">&nbsp;<label for="sendEmail" id="lblsendEmail" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Automatically send shipment email</label><br />
	  <input type="checkbox" id="saveMessageBody" name="saveMessageBody" value="1" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">&nbsp;<label for="saveMessageBody" id="lblsaveMessageBody" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Save email message body to message field</label><br />
      <label for="emailFile" id="lblemailFile" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Email Template To Use</label>
      <select name="emailFile" ID="emailFile" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">
      <%
      Dim pstrFilePath
      Dim pclsEmail
      Dim paryEmail
      	pstrFilePath = ssAdminPath & cstrEmailTemplateFolder
	'debug.print "ssAdminPath", ssAdminPath
	Set pclsEmail = New clsEmail
	Call pclsEmail.LoadEmailTemplates(pstrFilePath, cstrDefaultOrderShippedEmailTemplate, paryEmail)
	Set pclsEmail = Nothing
	If isArray(paryEmail) Then
		For i = 0 To UBound(paryEmail)
		%><option value="<%= paryEmail(i)(enEmail_FileName) %>"><%= paryEmail(i)(enEmail_TemplateName) %></option><%
		Next 'i
	End If
      %>
		<option value="" selected>- Use Import File Default -</option>
      </select>

    </td>
  </tr>
  <!--
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><div style="cursor: hand;" id="Div1" onclick="ClearTrackingNumbers()" title="Click to clear contents of table. Used if you need to reset table if you make manual changes to the tracking file below.">Reset Table</div></td>
  </tr>
  -->
  <TR>
	<td><span id="importStep4" onmouseover="stm(tipMessage[this.id],tipStyle['dataEntry']);" onMouseOut="htm();">Step 4.</span>&nbsp;</td>
    <TD>Import tracking numbers:&nbsp;</TD>
	<TD><INPUT class="butn" id="btnSend" name="btnSend" type="submit" value="Import"></td>
  </TR>
  <TR>
    <TD colspan=3 align=center><HR width=90%></TD>
  </TR>
  <tr>
	<td>&nbsp;</td>
	<td></td>
	<td><div id="divFileKeys"></div></td>
  </tr>
  <TR>
	<td>&nbsp;</td>
    <td>Data file:&nbsp;</td>
    <TD><textarea name="dataFile" id="dataFile" rows="1" cols="60"></textarea></TD>
  </TR>
</table>
<input type="hidden" name="cTrackingColumn_OrderNumber" id="cTrackingColumn_OrderNumber" value="">
<input type="hidden" name="cTrackingColumn_Message" id="cTrackingColumn_Message" value="">
<input type="hidden" name="cTrackingColumn_ShippingMethod" id="cTrackingColumn_ShippingMethod" value="">
<input type="hidden" name="cTrackingColumn_TrackingNumber" id="cTrackingColumn_TrackingNumber" value="">
<input type="hidden" name="cTrackingColumn_ShipmentDate" id="cTrackingColumn_ShipmentDate" value="">
<input type="hidden" name="cTrackingColumn_SendEmail" id="cTrackingColumn_SendEmail" value="">
<input type="hidden" name="cstrDefault_CarrierID" id="cstrDefault_CarrierID" value="">
<% End Sub 'ShowImportTracking %>
