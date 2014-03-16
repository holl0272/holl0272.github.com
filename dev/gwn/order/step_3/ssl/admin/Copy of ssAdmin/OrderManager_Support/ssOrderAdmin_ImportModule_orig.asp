<%
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version   2.00.001	                                                *
'*   Release Date      December 16, 2003										*
'*   Revision Date     December 16, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************
%>
<% Sub ShowImportTracking()  %>

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

	Function GetTrackingNumbers(theFile)

	Dim pstrFile
	Dim pblnUseCommas
	Dim paryShipments
	Dim paryShipment
	Dim i, j
	Dim pstrCellHTML
	Dim paryValues(3)
	dim pNewRow
	dim pNewCell
	dim ptempDate

		pstrFile = OpenFile(theFile)
		pblnUseCommas = Instr(1,pstrFile,",")
		paryShipments = Split(pstrFile,vbcrlf)
		'document.frmData.trackingNumbers.value = pstrFile
		
		For i = 0 To UBound(paryShipments)-1
		
			Set pNewRow = document.all("tblImport").insertRow()
			
			If pblnUseCommas Then
				paryShipment = Split(paryShipments(i),",")
			Else
				paryShipment = Split(paryShipments(i),vbTab)
			End If
			
			paryValues(3) = Date()
			For j = 0 To UBound(paryShipment)
				'paryValues(j) = paryShipment(j)
				paryValues(j) = stripQuotes(paryShipment(j))
			Next 'j
			If Not isDate(paryValues(3)) Then
				ptempDate = paryValues(3)
				If isNumeric(ptempDate) Then
					ptempDate = ptempDate/1000000
					paryValues(3) = Left(ptempDate,2) & "/" & Mid(ptempDate,3,2) & "/" & Right(ptempDate,4)
					'msgbox(ptempDate)
				End If
			End If

			For j = 0 To 5
				Select Case j
					Case 0: pstrCellHTML = "<!--<INPUT type=checkbox id=Update name=Update." & paryValues(0) & " value=1 checked>-->" _
										 & "<INPUT type=hidden id=shipOrderID name=shipOrderID value='" & paryValues(0) & "'>" _
										 & "<INPUT type=hidden id=isDirty name=isDirty value='1'>" _
										 & "<INPUT type=hidden id=shipVia name=shipVia value='" & ShipCodeToCarrierID(paryValues(1)) & "'>" _
										 & "<INPUT type=hidden id=shipTrackingNumber name=shipTrackingNumber value='" & paryValues(2) & "'>" _
										 & "<INPUT type=hidden id=shipDate name=shipDate value='" & paryValues(3) & "'>"
					Case 1: pstrCellHTML = "<INPUT type=checkbox id=shipMail name=shipMail." & paryValues(0) & " value='1' checked>"
					Case 2: pstrCellHTML = paryValues(0)
					Case 3: pstrCellHTML = ShipIDToName(ShipCodeToCarrierID(paryValues(1)))	& "(" & paryValues(1) & ")"
					Case 4: pstrCellHTML = paryValues(2)
					Case 5: pstrCellHTML = paryValues(3)
				End Select
				Set pNewCell = pNewRow.insertCell()
				pNewCell.innerHTML = pstrCellHTML
			Next 'j

		Next 'i

	End Function	'GetTrackingNumbers

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function GetEmail(theFile)

	Dim pstrFile

		pstrFile = OpenFile(theFile)
		document.frmData.emailBody.value = pstrFile

	End Function	'GetEmail

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

<table class="tbl" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none" id="tblImportTracking">
	<COLGROUP align="left" />
	<COLGROUP align="left" width="100%" />
  <TR>
	<TD>&nbsp;</td>
	<TD>
	<TABLE class="tbl" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblImport" width="100%">
	  <COLGROUP align=center></COLGROUP>
	  <COLGROUP align=center></COLGROUP>
	  <COLGROUP align=center></COLGROUP>
	  <COLGROUP align=center></COLGROUP>
	  <COLGROUP align=center></COLGROUP>
      <TR>
        <TD><!--Update-->&nbsp;</TD>
        <TD>Send email</TD>
        <TD>Order ID</TD>
        <TD>Shipping Method</TD>
        <TD>Tracking Number</TD>
        <TD>Shipping Date</TD>
      </TR>
	</TABLE>
	</td>
  <TR>
  <tr>
    <td>Step&nbsp;1.&nbsp;Select&nbsp;Tracking&nbsp;Numbers&nbsp;file&nbsp;:&nbsp;</td>
    <td>
	  <input type=file id="trackingNumbersFile" name="trackingNumbersFile" onchange="javascript:GetTrackingNumbers(this);" size="20">
    </td>
  </tr>
  <tr>
    <td>Step 2. Click to review:&nbsp;</td>
    <td><div style="cursor: hand;" id="divPreview" onclick="GetTrackingNumbers(frmData.trackingNumbersFile)">Preview</div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
	  <input type="checkbox" id="updatePayment" name="updatePayment" value="1" checked>Automatically update unrecorded payments
    </td>
  </tr>
  <TR>
    <TD>Step 3. Import tracking numbers:&nbsp;</TD>
	<TD><INPUT class="butn" id="btnSend" name="btnSend" type="submit" value="Import"></td>
  </TR>
  <TR>
    <TD colspan=2 align=center><HR width=90%></TD>
  </TR>
</table>
<% End Sub 'ShowImportTracking %>
