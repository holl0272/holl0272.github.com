<%
Option Explicit
Response.Buffer = False
Server.ScriptTimeout = 900

'********************************************************************************
'*   Customer Manager for StoreFront 5.0                                        *
'*   Release Version:	2.00.004		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		March 16, 2005											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   2.00.004 (March 16, 2005)													*
'*   - Enhancement - added tabbed interface										*
'*                                                                              *
'*   2.00.003 (January 14, 2004)                                                *
'*   - Bug fix - update routine modified to use nulls instead of empty values   *
'*                                                                              *
'*   2.00.002 (November 6, 2003)                                                *
'*   - Added Pricing Level Manager support                                      *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	'NONE
	
'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************


Class clsItem
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pblnError
Private pstrMessage
Private pobjRS
Private pvntID
Private plngFixedContentType

'Variables specific to table
Private pstrTableName
Private pstrDisplayField
Private paryCustomValues

Private pstrTableFilter
Private pstrTableSortField
Private pstrTableSortOrder

'Added for export capability
Private pstrExportTemplateDirectory
Private pstrKeyFieldName
Private pstrKeyFieldIsNumeric

Private pstrPageTitle
Private pstrPageVersion
Private pstrURL

'***********************************************************************************************

Private Sub class_Initialize()
	Call InitializeCustomValues
End Sub

Private Sub class_Terminate()

Dim i

    On Error Resume Next
	Call ReleaseObject(pobjRS)
	For i = 0 To UBound(paryCustomValues)
		Call ReleaseObject(paryCustomValues(i)(enCustomField_sqlSource))
	Next 'i

End Sub

'***********************************************************************************************

Public Property Get FixedContentType
	FixedContentType = plngFixedContentType
End Property
Public Property Get tableName
	tableName = pstrTableName
End Property
Public Property Get KeyFieldName
	KeyFieldName = pstrKeyFieldName
End Property
Public Property Get KeyFieldIsNumeric
	KeyFieldIsNumeric = pstrKeyFieldIsNumeric
End Property
Public Property Get ExportTemplateDirectory
	ExportTemplateDirectory = pstrExportTemplateDirectory
End Property
Public Property Get PageTitle
	PageTitle = pstrPageTitle
End Property
Public Property Get PageVersion
	PageVersion = pstrPageVersion
End Property
Public Property Get URL
	URL = pstrURL
End Property

'***********************************************************************************************

Private Sub InitializeCustomValues

Dim i

	pstrTableName = "sfCustomers"
	pstrKeyFieldName = "custID"
	pstrKeyFieldIsNumeric = True
	pstrExportTemplateDirectory = "sfCustomers/"	'Note: this must end in a /, use - to not display export capability
	pstrDisplayField = "custLastName"
	pstrTableSortField = "custID"
	pstrTableSortOrder = "Asc"
	pstrPageTitle = "Customer Administration"
	pstrPageVersion = "2.00.004"
	pstrURL = "sfCustomerAdmin.asp"
	
	'format: Display Text, field name, field value(must be ""), DisplayType, DisplayLength, sqlSource, Datatype, Show in summary

	'Datatype Enumerations - defined in modDatabase.asp
	'enDatatype_string, enDatatype_number, enDatatype_date, enDatatype_boolean

	ReDim paryCustomValues(20)
	paryCustomValues(0) = Array("custID", "custID", "", enDisplayType_hidden, "4", "", enDatatype_number, False, "", "Customer ID, hidden field")
	paryCustomValues(1) = Array("Last Name", "custLastName", "", enDisplayType_textbox, "20", "", enDatatype_string, True, "", "Customer last name")
	paryCustomValues(2) = Array("FirstName", "custFirstName",   "", enDisplayType_textbox, "20", "", enDatatype_string, True, "")
	paryCustomValues(3) = Array("MI", "custMiddleInitial", "", enDisplayType_textbox, "3", "", enDatatype_string, False, "")
	paryCustomValues(4) = Array("Company", "custCompany", "", enDisplayType_textbox, "20", "", enDatatype_string, True, "")
	paryCustomValues(5) = Array("Addr1", "custAddr1", "", enDisplayType_textbox, "30", "", enDatatype_string, False, "")
	paryCustomValues(6) = Array("Addr2", "custAddr2", "", enDisplayType_textbox, "30", "", enDatatype_string, False, "")
	paryCustomValues(7) = Array("City", "custCity", "", enDisplayType_textbox, "20", "", enDatatype_string, False, "")
	paryCustomValues(8) = Array("State", "custState", "", enDisplayType_select, "4", "SELECT loclstAbbreviation, loclstName FROM sfLocalesState WHERE loclstLocaleIsActive=1 Order By loclstName", enDatatype_string, False, "")
	paryCustomValues(9) = Array("Zip", "custZip", "", enDisplayType_textbox, "20", "", enDatatype_string, False, "")
	paryCustomValues(10) = Array("Country", "custCountry", "", enDisplayType_select, "4", "SELECT loclctryAbbreviation, loclctryName FROM sfLocalesCountry WHERE loclctryLocalIsActive=1 Order By loclctryName", enDatatype_string, False, "")
	paryCustomValues(11) = Array("Email", "custEmail", "", enDisplayType_textbox, "50", "", enDatatype_string, True, "")
	paryCustomValues(12) = Array("Phone", "custPhone", "", enDisplayType_textbox, "10", "", enDatatype_string, False, "")
	paryCustomValues(13) = Array("FAX", "custFAX", "", enDisplayType_textbox, "10", "", enDatatype_string, False, "")
	paryCustomValues(14) = Array("Password", "custPasswd", "", enDisplayType_textbox, "10", "", enDatatype_string, False, "")
	paryCustomValues(15) = Array("Is Subscribed", "custIsSubscribed", "", enDisplayType_checkbox, "20", "", enDatatype_number, True, "")
	paryCustomValues(16) = Array("Times Accessed", "custTimesAccessed", "", enDisplayType_textbox, "10", "", enDatatype_number, True, "")
	paryCustomValues(17) = Array("Last Access", "custLastAccess", "", enDisplayType_textbox_WithDateSelect, "20", "", enDatatype_date, True, "")
	paryCustomValues(18) = Array("Pricing Level", "PricingLevelID", "", enDisplayType_select, "4", "Select PricingLevelID, PricingLevelName from PricingLevels Order By PricingLevelName", enDatatype_number, False)
	paryCustomValues(19) = Array("Club Code", "clubCode", "", enDisplayType_select, "4", "Select PromoCode, PromoTitle From Promotions Order By PromoTitle", enDatatype_string, False, "")
	paryCustomValues(20) = Array("Club Expr. Date", "clubExpDate", "", enDisplayType_textbox_WithDateSelect, "10", "", enDatatype_date, False, "")

	'--------------------------------------------------------------------------------------------------------------------

	For i = 0 To UBound(paryCustomValues)
		If Len(paryCustomValues(i)(enCustomField_sqlSource)) > 0 Then Set paryCustomValues(i)(enCustomField_sqlSource) = GetRS(paryCustomValues(i)(enCustomField_sqlSource))
	Next 'i

End Sub	'InitializeCustomValues

'***********************************************************************************************

Public Function getIndexByFieldName(strFieldName)

Dim i

	If isArray(paryCustomValues) Then
		For i = 0 To UBound(paryCustomValues)
			If paryCustomValues(i)(enCustomField_FieldName) = strFieldName Then
				getIndexByFieldName = i
				Exit Function
			End If
		Next 'i
	End If
	
	getIndexByFieldName = -1
	
End Function	'getIndexByFieldName

'***********************************************************************************************

Private Sub LoadCustomValues(objRS)

Dim i

	If Not isArray(paryCustomValues) Then Exit Sub
	For i = 0 To UBound(paryCustomValues)
		paryCustomValues(i)(enCustomField_FieldValue) = objRS.Fields(paryCustomValues(i)(enCustomField_FieldName)).Value
	Next 'i
	
End Sub	'LoadCustomValues

'***********************************************************************************************

Private Sub LoadCustomValuesFromRequest()

Dim i

	If Not isArray(paryCustomValues) Then Exit Sub
	For i = 0 To UBound(paryCustomValues)
		paryCustomValues(i)(enCustomField_FieldValue) = Trim(Request.Form(paryCustomValues(i)(enCustomField_FieldName)))
	Next 'i
	
End Sub	'LoadCustomValuesFromRequest

'***********************************************************************************************

Public Property Get Message()
    Message = pstrMessage
End Property

'***********************************************************************************************

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

Public Property Get CustomValues()
    CustomValues = paryCustomValues
End Property

'***********************************************************************************************

Public Property Let SortField(strValue)
    pstrTableSortField = strValue
End Property

Public Property Let SortOrder(strValue)
    pstrTableSortOrder = strValue
End Property

Public Property Let TableFilter(strValue)
    pstrTableFilter = strValue
End Property

'***********************************************************************************************

Public Property Get rsItems()
    If isObject(pobjRS) Then Set rsItems = pobjRS
End Property

Public Property Get Records()
    If isObject(pobjRS) Then Set Records = pobjRS
End Property
'***********************************************************************************************

Public Function Load()

dim pstrSQL
dim p_strWhere
dim i
dim sql

'On Error Resume Next

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	
	set	pobjRS = server.CreateObject("adodb.recordset")
	With pobjRS
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server
        pstrSQL = "SELECT * From " & pstrTableName
        
        If Len(pstrTableFilter) > 0 Then pstrSQL = pstrSQL & " Where " & pstrTableFilter
        If Len(pstrTableSortField) > 0 Then pstrSQL = pstrSQL & " Order By " & pstrTableSortField & " " & pstrTableSortOrder

		If len(mlngMaxRecords) > 0 Then 
			.CacheSize = mlngMaxRecords
			.PageSize = mlngMaxRecords
		End If

		On Error Resume Next
		
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		If Err.number <> 0 Then
			Response.Write "<fieldset><legend>Error loading summary</legend>"
			Response.Write "Error " & err.number & ": " & err.Description & "<br />"
			Response.Write "SQL: " & pstrSQL & "<br />"
			Response.Write "No filter applied<br />"
			Response.Write "</fieldset>"
			err.Clear
			.Open pstrBaseSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		End If
		mlngPageCount = .PageCount
		If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
		
		Dim plnglbound
		If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
		If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
		plnglbound = (mlngAbsolutePage - 1) * pobjRS.PageSize + 1
		If Not pobjRS.EOF Then pobjRS.AbsolutePosition = plnglbound

	End With

	If Not pobjRS.EOF Then 
		pvntID = pobjRS.Fields(paryCustomValues(0)(enCustomField_FieldName)).Value
		Call LoadCustomValues(pobjRS)
	End If
    Load = (Not pobjRS.EOF)

End Function    'Load

'***********************************************************************************************

Private Sub LoadValues(objRS)

Dim i

	If Not isArray(paryCustomValues) Then Exit Sub
	For i = 0 To UBound(paryCustomValues)
		paryCustomValues(i)(enCustomField_FieldValue) = Trim(objRS.Fields(paryCustomValues(i)(enCustomField_FieldName)).Value & "")
	Next 'i

End Sub 'LoadValues

'***********************************************************************************************

Public Function Find(lngID)

Dim pstrSQL

'On Error Resume Next

    With pobjRS
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
				pvntID = lngID
				pstrSQL = paryCustomValues(0)(enCustomField_FieldName) & "=" & wrapSQLValue(pvntID, False, paryCustomValues(0)(enCustomField_DataType)) 
                .Find pstrSQL
            Else
                .MoveLast
            End If
            If Not .EOF Then LoadValues (pobjRS)
        End If
    End With

End Function    'Find

'***********************************************************************************************

Public Function DeleteSelected(byVal strIDs)

Dim i
Dim paryIDs
Dim pstrSQL

'On Error Resume Next

	If len(strIDs) = 0 Then Exit Function
	
	paryIDs = Split(strIDs, ",")
	For i = 0 To UBound(paryIDs)
		Call Delete(paryIDs(i))
	Next
    
End Function    'Delete

'***********************************************************************************************

Public Function Delete(byVal vntID)

Dim pstrSQL

'On Error Resume Next

	If len(vntID) = 0 Then
		Exit Function
	Else
		vntID = Trim(vntID)
	End If
	
	pstrSQL = "Delete From " & pstrTableName & " Where " & paryCustomValues(0)(enCustomField_FieldName) & "=" & wrapSQLValue(vntID, False, paryCustomValues(0)(enCustomField_DataType)) 
	cnn.Execute pstrSQL, , 128
	
    If (Err.Number = 0) Then
        pstrMessage = "Deletion Successful"
        Delete = True
    Else
        pstrMessage = Err.Description
        Delete = False
    End If
    
End Function    'Delete

'***********************************************************************************************

Public Function Update()

Dim pstrSQL
Dim objRS
Dim strErrorMessage
Dim blnAdd
Dim pstrOrigprodID
Dim p_strTableName, p_strFieldName
Dim i

'On Error Resume Next

    pblnError = False
    Call LoadCustomValuesFromRequest

    'strErrorMessage = ValidateValues
    If Len(strErrorMessage) = 0 Then
    
		Select Case paryCustomValues(0)(enCustomField_DataType)
			Case enDatatype_string
				pstrSQL = "SELECT * From " & pstrTableName & " Where " & paryCustomValues(0)(enCustomField_FieldName) & "=" & wrapSQLValue(paryCustomValues(0)(enCustomField_FieldValue), True, enDatatype_string) 
			Case enDatatype_number
				If Len(paryCustomValues(0)(enCustomField_FieldValue)) = 0 Then
					pstrSQL = "SELECT * From " & pstrTableName & " Where " & paryCustomValues(0)(enCustomField_FieldName) & "=0"
				Else
					pstrSQL = "SELECT * From " & pstrTableName & " Where " & paryCustomValues(0)(enCustomField_FieldName) & "=" & paryCustomValues(0)(enCustomField_FieldValue)
				End If
		End Select

        Set objRS = server.CreateObject("adodb.Recordset")
		objRS.CursorLocation = 3
        objRS.open pstrSQL, cnn, 1, 3
        If objRS.EOF Then
            objRS.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

		If paryCustomValues(0)(enCustomField_DataType) = enDatatype_string Then objRS.Fields(paryCustomValues(i)(enCustomField_FieldName)).Value = paryCustomValues(i)(enCustomField_FieldValue)
		For i = 1 To UBound(paryCustomValues)
			If paryCustomValues(i)(enCustomField_DisplayType) = enDisplayType_checkbox Then
				'debugprint paryCustomValues(i)(enCustomField_FieldName),paryCustomValues(i)(enCustomField_FieldValue)
				If CBool(Len(paryCustomValues(i)(enCustomField_FieldValue)) > 0) Then
					objRS.Fields(paryCustomValues(i)(enCustomField_FieldName)).Value = 1
				Else
					objRS.Fields(paryCustomValues(i)(enCustomField_FieldName)).Value = 0
				End If
			Else
				'debugprint paryCustomValues(i)(enCustomField_FieldName),paryCustomValues(i)(enCustomField_FieldValue)
				
				On Error Resume Next
				'objRS.Fields(paryCustomValues(i)(enCustomField_FieldName)).Value = paryCustomValues(i)(enCustomField_FieldValue)
				objRS.Fields(paryCustomValues(i)(enCustomField_FieldName)).Value = wrapSQLValue(paryCustomValues(i)(enCustomField_FieldValue), False, enDatatype_NA)
				If Err.number = -2147217887 Then
					pstrMessage = pstrMessage & "<font color=red>Error updating " & paryCustomValues(i)(enCustomField_DisplayText) & ": value '<i>" & paryCustomValues(i)(enCustomField_FieldValue) & "</i>' is invalid. Please make sure it is the right type, has a value, or is not too long.<br />"					
				ElseIf Err.number <> 0 Then
					pstrMessage = pstrMessage & "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />"
					pstrMessage = pstrMessage & "<font color=red>Error updating " & paryCustomValues(i)(enCustomField_DisplayText) & ": value '<i>" & paryCustomValues(i)(enCustomField_FieldValue) & "</i>' is invalid. Please make sure it is the right type, has a value, or is not too long.<br />"					
				End If
				On Error Goto 0
			End If
		Next 'i

		objRS.Update
		
		'need to check for {id} replacement
		If blnAdd Then
			For i = 1 To UBound(paryCustomValues)
				If paryCustomValues(i)(enCustomField_DataType) = enDatatype_string Then
					If InStr(1, paryCustomValues(i)(enCustomField_FieldValue), "{id}") > 0 Then
						paryCustomValues(i)(enCustomField_FieldValue) = Replace(paryCustomValues(i)(enCustomField_FieldValue), "{id}", objRS.Fields(paryCustomValues(0)(enCustomField_FieldName)).Value)
						objRS.Fields(paryCustomValues(i)(enCustomField_FieldName)).Value = wrapSQLValue(paryCustomValues(i)(enCustomField_FieldValue), False, enDatatype_NA)
						objRS.Update
					End If
				End If
			Next 'i
		End If
		
        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = pstrMessage &  cstrdelimeter & "<H4>The data you entered is already in use.<br />Please enter a different data.</H4><br />"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br />"
        End If
		If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstrMessage & cstrdelimeter & paryCustomValues(getIndexByFieldName(pstrDisplayField))(enCustomField_FieldValue) & " was successfully added."
            Else
                pstrMessage = pstrMessage & cstrdelimeter & "The changes to " & paryCustomValues(getIndexByFieldName(pstrDisplayField))(enCustomField_FieldValue) & " were successfully saved."
            End If
        Else
            pblnError = True
        End If

        objRS.Close
		Set objRS = Nothing
    Else
        pblnError = True
    End If

    Update = (not pblnError)

End Function    'Update

'***********************************************************************************************

Sub ShowTemplateSelections

Dim paryExportTemplates

	If pstrExportTemplateDirectory <> "-" Then
		Call getFileNamesInFolder(ssAdminPath & "exportTemplates/" & pstrExportTemplateDirectory, ".xsl", paryExportTemplates)
		
		Response.Write "<tr>"
		Response.Write "<td colspan=""" & UBound(paryCustomValues) + 1 & """>"
		Response.Write "<table class=""tbl"" width=""100%"" cellpadding=""3"" cellspacing=""0"" border=""1"" rules=""none"" id=""tblSummaryFunctions"">"
		Response.Write "<tr>"
		Response.Write "<td valign=""middle"">"
		
		If isAllowedToDeleteItems Then Response.Write "<a href="""" onclick=""DeleteSelected(); return false;"">Delete Selected Items</a>"

		Response.Write "</td>"
		Response.Write "<td align=""right"" valign=""middle"">"
		If isArray(paryExportTemplates) Then
			Response.Write "<select name=""ExportTemplates"" id=""ExportTemplates"">"
			Response.Write "<option value="""" selected>Select a Template</option>"
			For i = 0 To UBound(paryExportTemplates)
				Response.Write "<option value=""" & paryExportTemplates(i) & """>" & Replace(paryExportTemplates(i), ".xsl", "") & "</option>"
			Next 'i
			Response.Write "</select>"
			Response.Write "<input class=""butn"" name=""btnView"" id=""btnView"" type=button value=""View"" onclick=""viewSelected(''); return false;"" title=""View selected items"">&nbsp;&nbsp;"
			Response.Write "<input class=""butn"" name=""btnDownload"" id=""btnDownload"" type=image src=""images/save.gif"" value=""Download"" onclick=""downloadSelected(''); return false;"" title=""Download selected items using the selected template"">&nbsp;&nbsp;"
			Response.Write "<input class=""butn"" name=""btnPrint"" id=""btnPrint"" type=image src=""images/print.gif"" value=""Print"" onclick=""printSelected(); return false;"" title=""Print selected items"">&nbsp;&nbsp;"
		Else
			Response.Write "<div style=""color:red;font-size:14pt;background-color:yellow;border:solid 1pt black"">No templates exist in the " & "exportTemplates/" & pstrExportTemplateDirectory & " directory.</div>"
		End If	'isArray(paryExportTemplates)
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "</td>"
		Response.Write "</tr>"
		
	End If	'pstrExportTemplateDirectory <> "-"
End Sub	'ShowTemplateSelections

'***********************************************************************************************

Public Sub OutputSummary()

'On Error Resume Next

Dim i, j
Dim pstrTitle
Dim pstrSelect, pstrHighlight
Dim pstrID
Dim pblnSelected
Dim pbytNumColumns
Dim pblnItemLinkDisplayed

	pbytNumColumns = 0
	With Response

		.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' bgcolor='whitesmoke' id='tblSummary' rules='none'>"	'
		Call ShowTemplateSelections
		.Write "	<tr class='tblhdr'>"
		.Write "    <th><input type='checkbox' name='chkCheckAll' id='chkCheckAll'  onclick='checkAll(this.form.chkItemID, this.checked);' value=''></th>"
		For i = 0 To UBound(paryCustomValues)
			If paryCustomValues(i)(enCustomField_DisplayType) <> enDisplayType_hidden And paryCustomValues(i)(enCustomField_ShowInSummary) Then
				pbytNumColumns = pbytNumColumns + 1
				If pstrTableSortField = paryCustomValues(i)(enCustomField_FieldName) Then
					If (pstrTableSortOrder = "ASC") Then
						.Write "  <TH align=left style='cursor:hand;' onclick=""" & "SortColumn('" & paryCustomValues(i)(enCustomField_FieldName) & "','DESC');" & """ onMouseOver='HighlightColor(this);' onMouseOut='deHighlightColor(this);'"">" _
							& paryCustomValues(i)(enCustomField_DisplayText) & "&nbsp;&nbsp;&nbsp;<img src='images/up.gif' border=0 align=bottom></TH>" & vbCrLf
					Else
						.Write "  <TH align=left style='cursor:hand;' onclick=""" & "SortColumn('" & paryCustomValues(i)(enCustomField_FieldName) & "','ASC');" & """ onMouseOver='HighlightColor(this);' onMouseOut='deHighlightColor(this);'"">" _
							& paryCustomValues(i)(enCustomField_DisplayText) & "&nbsp;&nbsp;&nbsp;<img src='images/down.gif' border=0 align=bottom></TH>" & vbCrLf
					End If
				Else
					.Write "  <TH align=left style='cursor:hand;' onclick=""" & "SortColumn('" & paryCustomValues(i)(enCustomField_FieldName) & "','DESC');" & """ onMouseOver='HighlightColor(this);' onMouseOut='deHighlightColor(this);'"">" _
						& paryCustomValues(i)(enCustomField_DisplayText) & "</TH>" & vbCrLf
				End If
			'.Write "<TH align=left>" & paryCustomValues(i)(enCustomField_DisplayText) & "</TH>"
			End If
		Next
		.Write "	</tr>"

		
		If pobjRS.RecordCount > 0 Then
			pobjRS.MoveFirst

			'Need to calculate current recordset page and upper bound to loop through
			dim plnguBound, plnglbound, pstrDisplay

			If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
			If mlngAbsolutePage = 0 Then mlngAbsolutePage = 1
			plnglbound = (mlngAbsolutePage - 1) * pobjRS.PageSize + 1
			plnguBound = mlngAbsolutePage * pobjRS.PageSize

			If plnguBound > pobjRS.RecordCount Then plnguBound = pobjRS.RecordCount
				pobjRS.AbsolutePosition = plnglbound
				For i = plnglbound To plnguBound
		        
					pstrID = trim(pobjRS.Fields(paryCustomValues(0)(enCustomField_FieldName)).Value)
					pstrTitle = "Click to view " & Replace(pobjRS.Fields(paryCustomValues(1)(enCustomField_FieldName)).Value & "","'","")
					pstrSelect = "title='" & pstrTitle & "' " _
							& "onmouseover='doMouseOverRow(this); DisplayTitle(this);' " _
							& "onmouseout='doMouseOutRow(this); ClearTitle();' " _
							& "onmousedown='viewItem(" & chr(34) & pstrID & chr(34) & ");'"
					
					pblnSelected = CBool(Trim(pstrID) = Trim(pvntID))

					If pblnSelected Then
						.Write "<TR class='Selected'>"
					Else
						'.Write " <TR class='Inactive' " & pstrSelect & ">"
						.Write " <TR class='Inactive'>"
					End If
					pblnItemLinkDisplayed = False
					
					.Write "<TD><input type=checkbox name=chkItemID id=chkItemID value=" & Chr(34) & Server.HTMLEncode(pobjRS.Fields(paryCustomValues(0)(enCustomField_FieldName)).Value) & Chr(34) & ">&nbsp;</TD>"
					For j = 0 To UBound(paryCustomValues)
						If paryCustomValues(j)(enCustomField_DisplayType) <> enDisplayType_hidden And paryCustomValues(j)(enCustomField_ShowInSummary) Then 
							If paryCustomValues(j)(enCustomField_DisplayType) = enDisplayType_checkbox Then
								If isNull(pobjRS.Fields(paryCustomValues(j)(enCustomField_FieldName)).Value) Then
									.Write "<TD>False</TD>"
								ElseIf CBool(pobjRS.Fields(paryCustomValues(j)(enCustomField_FieldName)).Value) Then
									.Write "<TD>True</TD>"
								Else
									.Write "<TD>False</TD>"
								End If
							ElseIf paryCustomValues(j)(enCustomField_DisplayType) <> enDisplayType_select Then
								If pblnItemLinkDisplayed Then
									.Write "<TD>"
								Else
									.Write "<TD " & pstrSelect & ">"
									pblnItemLinkDisplayed = True
								End If
								.Write trim(pobjRS.Fields(paryCustomValues(j)(enCustomField_FieldName)).Value) & "</TD>"
							Else
								.Write "<TD>" & getSelectText(paryCustomValues(j)(enCustomField_sqlSource), pobjRS.Fields(paryCustomValues(j)(enCustomField_FieldName)).Value) & "</TD>"
							End If
						End If
					Next 'j
        			.Write "</TR>"
		        	
					pobjRS.MoveNext
				Next
			Else
					.Write "<TR><TD align=center COLSPAN=" & pbytNumColumns & "><h3>There are no Items</h3></TD></TR>"
			End If
    
			.Write "<tr class='tblhdr'><TH COLSPAN=" & pbytNumColumns + 1 & " align=center>"
			
			If pobjRS.RecordCount = 0 Then
				.Write "No Items match your search criteria"
			Elseif pobjRS.RecordCount = 1 Then
				.Write "1 Item matches your search criteria"
			Else 
				.Write pobjRS.RecordCount & " Items match your search criteria<br />"

			dim pstrCheck
			pstrCheck = "if (isInteger(this, true, ""Please enter a positive integer for the recordset page size."")){btnFilter_onclick(this);}else{return false;}"
			.Write "Show&nbsp;<input type='text' id='PageSize' name='PageSize' value='" & pobjRS.PageSize & "' maxlength='4' size='4' style='text-align:center;' onchange='" & pstrCheck & "'>&nbsp;records at a time.&nbsp;&nbsp;"

			If mlngPageCount > 1 Then
				Response.Write "&nbsp;Goto&nbsp;<select name=pageSelect id=pageSelect onchange='return ViewPage(this.selectedIndex+1);'>"
				For i=1 to mlngPageCount
					plnglbound = (i-1) * mlngMaxRecords + 1
					plnguBound = i * mlngMaxRecords
					if plnguBound > pobjRS.RecordCount Then plnguBound = pobjRS.RecordCount
					Response.Write "<option " & isSelected(i = cInt(mlngAbsolutePage)) & ">" & "Page " & i & " (" & plnglbound & " - " & plnguBound & ")</option>"
				Next
				Response.Write "</select>"
			End If	'mlngPageCount > 1
		End If
		.Write "</TH></TR>"
		.Write "</TABLE>"
	End With
End Sub      'OutputSummary

'******************************************************************************************************************************************************************

End Class   'clsItem
	
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'**********************************************************
'*	Functions
'**********************************************************

'Sub WriteFormOpener
'Sub WriteItemDetail
'Sub WriteCustomTable
'Function CustomDisplayText(byVal lngPos, byRef ary)
'Sub CustomOutput(byVal lngPos, byRef ary)
'Sub WriteFooterTable
'Sub WritePageHeader
'Function LoadFilter
'Sub WriteItemFilter()

'**********************************************************
'*	Page Level variables
'**********************************************************

	Dim maryCustomValues
	Dim mblnAutoShowTable
	Dim mblnShowDetail
	Dim mblnShowFilter
	Dim mblnShowHeader
	Dim mblnShowSummary
	Dim mbytSummaryTableHeight
	Dim mclsItem
	Dim mlngAbsolutePage
	Dim mlngMaxRecords
	Dim mlngPageCount
	Dim mradTextSearch
	Dim mstrAction
	Dim mstrItemTitle
	Dim mstrShow
	Dim mstrSortField
	Dim mstrSortOrder
	Dim mstrsqlWhere
	Dim mstrTextSearch
	Dim mvntID

'**********************************************************
'*	Begin Page Code
'**********************************************************

	mlngMaxRecords = LoadRequestValue("PageSize")
	If len(mlngMaxRecords) = 0 Then mlngMaxRecords = 50

	mblnShowHeader = True
	mblnShowDetail = True

	mstrShow = Request.Form("Show")
	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "false")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")
	
	mstrAction = LoadRequestValue("Action")
	If Len(mstrAction) = 0 Then mstrAction = "Filter"
	mlngAbsolutePage = LoadRequestValue("AbsolutePage")
	
	mstrSortField = LoadRequestValue("SortField")
	mstrSortOrder = LoadRequestValue("SortOrder")

    Set mclsItem = New clsItem
    With mclsItem
		mstrPageTitle = .PageTitle
		maryCustomValues = .CustomValues
		mvntID = Trim(Request.Form(maryCustomValues(0)(enCustomField_FieldName)))

		Select Case mstrAction
			Case "New", "Update"
				.Update
			Case "Delete"
				.Delete mvntID
				mvntID = ""
			Case "DeleteSelected"
				.DeleteSelected(Request.Form("chkItemID"))
				mvntID = ""
			Case "viewItem"
				mvntID = LoadRequestValue("ViewID")
			Case "Filter"
				mvntID = ""
		End Select
	    
	    If Len(mstrSortField) > 0 Then .SortField = mstrSortField
	    If Len(mstrSortOrder) > 0 Then .SortOrder = mstrSortOrder
	    .TableFilter = LoadFilter
		If .Load Then 
			maryCustomValues = .CustomValues
			If Len(mvntID) = 0 Then mvntID = maryCustomValues(0)(enCustomField_FieldValue)
			.Find mvntID
			maryCustomValues = .CustomValues
			mstrItemTitle = maryCustomValues(1)(enCustomField_FieldValue)
		End If
	
		Call WriteHeader("body_onload();",True)
%>
<script LANGUAGE=javascript>
<!--

var theDataForm;
var strDetailTitle = "<%= mstrItemTitle %>";
var blnIsDirty;
var strSubSection = "Status";

function MakeDirty(theItem)
{
var theForm = theItem.form;

	theForm.btnReset.disabled = false;
	blnIsDirty = true;
}

function body_onload()
{
	theDataForm = document.frmData;
	blnIsDirty = false;
	document.all("spanprodName").innerHTML = strDetailTitle;

<%
If mblnShowSummary Then
	Response.Write "DisplayMainSection('Summary');" & vbcrlf
ElseIf mblnShowFilter Then
	Response.Write "DisplayMainSection('Filter');" & vbcrlf
Else
	If mblnShowHeader Then Response.Write "DisplayMainSection('itemDetail');" & vbcrlf
	Response.Write "ScrollToElem('selectedSummaryItem');" & vbcrlf
	
	'Response.Write "DisplaySection(" & chr(34) & mstrShow & chr(34) & ");"
End If
%>
}

function DisplaySection(strSection)
{
<% 'Response.Write "return false;" %>

<% 
Dim pstrTempHeaderRow

pstrTempHeaderRow = "'General'"
pstrTempHeaderRow = pstrTempHeaderRow & ",'Custom'"
pstrTempHeaderRow = pstrTempHeaderRow & ",'ProductSales'"
%>
var arySections = new Array(<%= pstrTempHeaderRow %>);

	frmData.Show.value = strSection;

	for (var i=0; i < arySections.length;i++)
	{
		if (arySections[i] == strSection)
		{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "";
				document.all("td" + arySections[i]).className = "hdrSelected";
			}
		}else{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "none";
				document.all("td" + arySections[i]).className = "hdrNonSelected";
			}
		}
	}
}

function DisplayMainSection(strSection)
{

	var arySections = new Array('Filter', 'Summary', 'itemDetail');

	for (var i=0; i < arySections.length;i++)
	{
		if (arySections[i] == strSection)
		{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "";
				document.all("td" + arySections[i]).className = "hdrSelected";
			}
		}else{
			if (document.all("tbl" + arySections[i]) != null)
			{
				document.all("tbl" + arySections[i]).style.display = "none";
				document.all("td" + arySections[i]).className = "hdrNonSelected";
			}
		}
	}
	
	if (document.all("tblSummaryFunctions") != null)
	{
 		if (strSection == "Summary")
		{
			document.all("tblSummaryFunctions").style.display = "";
		}else{
			document.all("tblSummaryFunctions").style.display = "none";
		}
	}

	return(false);
}


function btnNewItem_onclick(theButton)
{
var theForm = theButton.form;

	SetDefaults(theForm);
    document.all("spanprodName").innerHTML = theDataForm.btnUpdate.value;

}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete this?");
    if (blnConfirm)
    {
    theForm.Action.value = "Delete";
    theForm.submit();
    return(true);
    }
    else
    {
    return(false);
    }
}

function SetDefaults(theForm)
{
<%  
Dim i

If isArray(maryCustomValues) Then 
	For i = 0 To UBound(maryCustomValues)
		Response.Write "theForm." & maryCustomValues(i)(enCustomField_FieldName) & ".value = " & Chr(34) & Chr(34) & ";" & vbcrlf
	Next 'i
End If
%>
    
    
return(true);
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.Action.value = "Filter";
	theDataForm.SortField.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.submit();
	return false;
}

function ViewPage(theValue)
{
	theDataForm.AbsolutePage.value = theValue;
	theDataForm.Action.value = "Filter";
	theDataForm.submit();
	return false;
}

function viewItem(theValue)
{
	theDataForm.ViewID.value = theValue;
	theDataForm.Action.value = "viewItem";
	theDataForm.submit();
	return false;
}

function ValidInput(theForm)
{
var  strSection = frmData.Show.value;

	theDataForm.submit();
    return(true);
}

function DeleteSelected()
{
	if (! anyChecked(theDataForm.chkItemID))
	{
		alert("Please select at least one item to delete.");
		return false;
	}
	
	var blnConfirm = confirm("Are you sure you wish to delete the selected item(s)?");
	if (blnConfirm)
	{
		theDataForm.Action.value = 'DeleteSelected';
		theDataForm.submit();
	}else{
		return false;
	}

}

function downloadSelected(strCustomAction, strExportTemplate, strExportField)
{
	var originalAction = theDataForm.action;
	
	if (! anyChecked(theDataForm.chkItemID))
	{
		alert("Please select at least one item to download.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	if (strExportTemplate == null)
	{
		strCustomAction = 'ssGenericTableExport.asp';
		if (theDataForm.ExportTemplates.selectedIndex == 0)
		{
			alert("Please select a template.");
			theDataForm.ExportTemplates.focus();
			return false;
		}
		strExportTemplate = theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value
	}
	
	if (strCustomAction == '')
	{
		strCustomAction = 'ssGenericTableExport.asp';
	}
	
	if (strExportField == null)
	{
		strExportField = '';
	}
	
	theDataForm.action='ssGenericTableExport.asp';
	theDataForm.Action.value = 'downloadOrders' + '|' + strExportTemplate + '|' + strExportField;
	theDataForm.target='docOrders';
	theDataForm.submit();
	theDataForm.action=originalAction;
	theDataForm.target='';
	return false;
}

function printSelected(strCustomExport)
{
	var originalAction = theDataForm.action;

	if (strCustomExport == null)
	{
		strCustomExport = 'ssGenericTableExport.asp';
		if (theDataForm.ExportTemplates.selectedIndex == 0)
		{
			alert("Please select a template.");
			theDataForm.ExportTemplates.focus();
			return false;
		}
	}
	
	if (! anyChecked(theDataForm.chkItemID))
	{
		alert("Please select at least one item to print.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	theDataForm.action = strCustomExport;
	theDataForm.Action.value = 'printOrders' + '|' + theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value;
	theDataForm.target='printOrders';
	theDataForm.submit();
	theDataForm.action=originalAction;
	theDataForm.target='';
	return false;

}

function viewSelected(strCustomAction, strExportTemplate, strExportField)
{
	var originalAction = theDataForm.action;

	if (! anyChecked(theDataForm.chkItemID))
	{
		alert("Please select at least one item to view.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	if (strExportTemplate == null)
	{
		strCustomAction = 'ssGenericTableExport.asp';
		if (theDataForm.ExportTemplates.selectedIndex == 0)
		{
			alert("Please select a template.");
			theDataForm.ExportTemplates.focus();
			return false;
		}
		strExportTemplate = theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value
	}
	
	if (strCustomAction == '')
	{
		strCustomAction = 'ssGenericTableExport.asp';
	}
	
	if (strExportField == null)
	{
		strExportField = '';
	}
	
	theDataForm.action=strCustomAction;
	theDataForm.Action.value = 'viewOrders' + '|' + strExportTemplate + '|' + strExportField;
	theDataForm.target='docOrders';
	theDataForm.submit();
	theDataForm.action=originalAction;
	theDataForm.target='';
	return false;
}

//-->
</script>
<center>
<%
End With

Call WriteFormOpener
Response.Write mclsItem.OutputMessage
%>
<table id="tblMainPage" class="tbl" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="1" cellspacing="0" border="0" rules="none">
  <tr>
	<th id="tdFilter" class="hdrSelected" nowrap onclick="return DisplayMainSection('Filter');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="Set your filter criteria here.">&nbsp;Filter&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tdSummary" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('Summary');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View items which meet the specified filter criteria">&nbsp;Summaries&nbsp;</th>
	<th nowrap width="2pt"></th>
	<th id="tditemDetail" class="hdrNonSelected" nowrap onclick="return DisplayMainSection('itemDetail');" onMouseOver="window.status = this.title" onMouseOut="window.status = ''" title="View the selected item's detail">&nbsp;Detail&nbsp;</th>
	<th width="90%" align=right><span class="pagetitle2"><%= mstrPageTitle %></span>&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/OrderManager/help_OrderManager.htm')" id="btnHelp" name="btnHelp" title="Release Version <%= mclsItem.PageVersion %>"></th>
  </tr>
  <tr>
	<td colspan="6" class="hdrSelected" height="1px"></td>
  </tr>
  <tr>
	<td colspan="6" style="border-style: solid; border-color: steelblue; border-width: 1; padding: 1">
	<%
		Call WriteItemFilter
		If (len(mstrAction) > 0 or mblnAutoShowTable) Then Response.Write mclsItem.OutputSummary
		If mblnShowDetail Then Call WriteItemDetail
	%>
	</td>
  </tr>
</table>
</FORM>
</center>
<!--#include file="adminFooter.asp"-->
</BODY>
</HTML>
<%

	Call ReleaseObject(cnn)
    If Response.Buffer Then Response.Flush

'************************************************************************************************************************************
'
'	SUPPORTING ROUTINES
'
'************************************************************************************************************************************

Sub WriteFormOpener
%>
<form id="frmData" name="frmData" onsubmit="return ValidInput(this);" method="post" action="<%= mclsItem.URL %>">
<input type=hidden id="ViewID" name="ViewID">
<input type=hidden id=Action name=Action value="Update">
<input type=hidden id=blnShowSummary name=blnShowSummary value="">
<input type=hidden id=blnShowFilter name=blnShowFilter value="">
<input type=hidden id=Show name=Show value="<%= mstrShow %>">
<input type=hidden id=AbsolutePage name=AbsolutePage value="<%= mlngAbsolutePage %>">
<input type=hidden id="SortField" name="SortField" value="<%= mstrSortField %>">
<input type=hidden id="SortOrder" name="SortOrder" value="<%= mstrSortOrder %>">
<input type=hidden id="tableName" name="tableName" value="<%= mclsItem.tableName %>">
<input type=hidden id="tableKeyFieldName" name="tableKeyFieldName" value="<%= mclsItem.KeyFieldName %>">
<input type=hidden id="tableKeyFieldIsNumeric" name="tableKeyFieldIsNumeric" value="<%= mclsItem.KeyFieldIsNumeric %>">
<input type=hidden id="ExportTemplateDirectory" name="ExportTemplateDirectory" value="<%= mclsItem.ExportTemplateDirectory %>">

<% End Sub	'WriteFormOpener %>

<% Sub WriteItemDetail %>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblitemDetail">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
  <tr class="tblhdr">
	<th align=center><span id="spanprodName"></span>&nbsp;</th>
  </tr>
  <tr>
    <td>
	<% 
		Call WriteCustomTable
		Call WriteFooterTable
	%>
	</td>
  </tr>
</table>
<%
End Sub	'WriteItemDetail

'************************************************************************************************************************************

Sub WriteCustomTable

Dim i

If Not isArray(maryCustomValues) Then Exit Sub
%>

<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblCustom">
	<colgroup align="left" width="25%">
	<colgroup align="left" width="75%">
	<% For i = 0 To UBound(maryCustomValues) 
		If maryCustomValues(i)(enCustomField_DisplayType) <> enDisplayType_hidden Then
	%>
      <tr>
        <td class="Label"><%= CustomDisplayText(i, maryCustomValues) %>:&nbsp;</td>
        <td><% Call CustomOutput(i, maryCustomValues) %></td>
      </tr>
    <% 
		Else
			Call CustomOutput(i, maryCustomValues)
		End If 
	  Next 'i
	%>
</table>
<% 
End Sub	'WriteCustomTable 

'**************************************************************************************************************************************************

Function CustomDisplayText(byVal lngPos, byRef ary)

	Select Case ary(lngPos)(enCustomField_FieldName)
		Case "custPasswd":
			CustomDisplayText = "<a href='../../../myAccount.asp?PrevPage=quickOrder.asp&Action=&Email=" & ary(mclsItem.getIndexByFieldName("custEmail"))(enCustomField_FieldValue) & "&Password=" & ary(lngPos)(enCustomField_FieldValue) & "' title='Impersonate this customer'>" & ary(lngPos)(enCustomField_DisplayText) & "</a>"
		Case "custEmail":
			CustomDisplayText = "<a href='mailTo:email=" & ary(lngPos)(enCustomField_FieldValue) & "' title='Send an email to this customer'>" & ary(lngPos)(enCustomField_DisplayText) & "</a>"
		Case "clubCode":
			CustomDisplayText = "<a href='ssPromotionsAdmin.asp' title='View Promotions'>" & ary(lngPos)(enCustomField_DisplayText) & "</a>"
		Case "PricingLevelID":
			CustomDisplayText = "<a href='ssPricingLevelAdmin.asp' title='View Pricing Levels'>" & ary(lngPos)(enCustomField_DisplayText) & "</a>"
		Case Else:
			If Len(getValueFromArray(ary(lngPos), enCustomField_DisplayTip, "")) = 0 Then
				'CustomDisplayText = ary(lngPos)(enCustomField_DisplayText)
				CustomDisplayText = "<label for=""" & ary(lngPos)(enCustomField_FieldName) & """>" & ary(lngPos)(enCustomField_DisplayText) & "</label>"
			Else
				CustomDisplayText = "<label for=""" & ary(lngPos)(enCustomField_FieldName) & """ onmouseover=""tipMessage['" & getValueFromArray(ary(lngPos), enCustomField_FieldName, "") & "']=['Data Entry Help','" & getValueFromArray(ary(lngPos), enCustomField_DisplayTip, "") & "'];showDataEntryTip(this);"" onmouseout=""htm();"">" & ary(lngPos)(enCustomField_DisplayText) & "</label>"
			End If
	End Select

End Function	'CustomDisplayText

'**************************************************************************************************************************************************

Sub CustomOutput(byVal lngPos, byRef ary)

Dim j
Dim pbytPricingLevelID

	Select Case ary(lngPos)(enCustomField_FieldName)
		Case "clubCode":
        %>
			<select size="1" name="<%= ary(lngPos)(enCustomField_FieldName) %>"  id="<%= ary(lngPos)(enCustomField_FieldName) %>">
			<% 	
				If len(ary(lngPos)(enCustomField_FieldValue)) = 0 then
					Response.Write "<option value='' selected>- None -</Option>" & vbcrlf
				Else
					Response.Write "<option value=''>- None -</Option>" & vbcrlf
				End If

				If ary(lngPos)(enCustomField_sqlSource).RecordCount > 0 Then ary(lngPos)(enCustomField_sqlSource).MoveFirst
				For j = 1 to ary(lngPos)(enCustomField_sqlSource).recordcount
					pbytPricingLevelID = trim(ary(lngPos)(enCustomField_sqlSource).Fields("PromoCode").Value)
					If len(ary(lngPos)(enCustomField_FieldValue)) > 0 then
						If pbytPricingLevelID <> trim(ary(lngPos)(enCustomField_FieldValue)) then
							Response.Write "<option value=" & chr(34) & pbytPricingLevelID & chr(34) & ">" & ary(lngPos)(enCustomField_sqlSource).Fields("PromoTitle").Value & "</option>" & vbcrlf
						Else
							Response.Write "<option selected value=" & chr(34) & pbytPricingLevelID & chr(34) & ">" & ary(lngPos)(enCustomField_sqlSource).Fields("PromoTitle").Value & "</option>" & vbcrlf
						End If
					Else
						Response.Write "<option value=" & chr(34) & pbytPricingLevelID & chr(34) & ">" & ary(lngPos)(enCustomField_sqlSource).Fields("PromoTitle").Value & "</option>" & vbcrlf
					End If	'len(ary(lngPos)(enCustomField_FieldValue)) > 0
					ary(lngPos)(enCustomField_sqlSource).movenext
				Next

				'Call MakeCombo(ary(lngPos)(enCustomField_sqlSource),"","",ary(lngPos)(enCustomField_FieldValue))
				%>
				</select>
		<%
		Case "PricingLevelID":
        %>
			<select size="1" name="<%= ary(lngPos)(enCustomField_FieldName) %>"  id="<%= ary(lngPos)(enCustomField_FieldName) %>">
			<% 	
				If len(ary(lngPos)(enCustomField_FieldValue)) = 0 then
					Response.Write "<option value='' selected>- None -</Option>" & vbcrlf
				Else
					Response.Write "<option value=''>- None -</Option>" & vbcrlf
				End If

				If ary(lngPos)(enCustomField_sqlSource).RecordCount > 0 Then ary(lngPos)(enCustomField_sqlSource).MoveFirst
				For j = 1 to ary(lngPos)(enCustomField_sqlSource).recordcount
					pbytPricingLevelID = trim(ary(lngPos)(enCustomField_sqlSource).Fields("PricingLevelID").Value)
					If len(ary(lngPos)(enCustomField_FieldValue)) > 0 then
						If pbytPricingLevelID <> trim(ary(lngPos)(enCustomField_FieldValue)) then
							Response.Write "<option value=" & chr(34) & pbytPricingLevelID & chr(34) & ">" & ary(lngPos)(enCustomField_sqlSource).Fields("PricingLevelName").Value & "</option>" & vbcrlf
						Else
							Response.Write "<option selected value=" & chr(34) & pbytPricingLevelID & chr(34) & ">" & ary(lngPos)(enCustomField_sqlSource).Fields("PricingLevelName").Value & "</option>" & vbcrlf
						End If
					Else
						Response.Write "<option value=" & chr(34) & pbytPricingLevelID & chr(34) & ">" & ary(lngPos)(enCustomField_sqlSource).Fields("PricingLevelName").Value & "</option>" & vbcrlf
					End If	'len(ary(lngPos)(enCustomField_FieldValue)) > 0
					ary(lngPos)(enCustomField_sqlSource).movenext
				Next

				'Call MakeCombo(ary(lngPos)(enCustomField_sqlSource),"","",ary(lngPos)(enCustomField_FieldValue))
				%>
				</select>
		<%
		Case Else:
			Response.Write writeHTMLFormElement(ary(lngPos)(enCustomField_DisplayType), ary(lngPos)(enCustomField_DisplayLength), ary(lngPos)(enCustomField_FieldName), ary(lngPos)(enCustomField_FieldName), ary(lngPos)(enCustomField_FieldValue), ary(lngPos)(enCustomField_sqlSource), "MakeDirty(this);") 
	End Select

End Sub	'CustomOutput

'**************************************************************************************************************************************************

Sub WriteFooterTable
%>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFooter">
  <tr>
    <td>&nbsp;</td>
    <td>
        <input class='butn' title='Create a new Item' id=btnNewItem name=btnNewItem type=button value='New' onclick='return btnNewItem_onclick(this)'>&nbsp;
        <input class='butn' title="Reset" id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)' disabled>&nbsp;&nbsp;
        <% If isAllowedToDeleteItems Then %><input class='butn' title="Delete this Item" id=btnDelete name=btnDelete type=button value='Delete' onclick='return btnDelete_onclick(this)'><% End If %>
        <input class='butn' title="Save changes" id=btnUpdate name=btnUpdate type=button value='Save Changes' onclick='return ValidInput(this.form);'>
    </td>
  </tr>
</table>
<%
End Sub	'WriteFooterTable

'************************************************************************************************************************************

Sub WritePageHeader
%>
<table border=0 cellPadding=5 cellSpacing=1 width="95%" ID="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
    <th>&nbsp;</th>
    <th align='right'>
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Filter">Hide Filter</div></a><br />
		<a href="#"><div id="divSummary" onclick="return DisplaySummary();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Summary">Hide Summary</div></a>
	</th>
  </tr>
</table>
<% 
End Sub	'WritePageHeader 

'***********************************************************************************************

Function LoadFilter

Dim pstrSelFilter
Dim pstrradFilter
Dim pstrsqlFragment
Dim pstrsqlWhere

	'modified so could link in directly
	mradTextSearch = LoadRequestValue("radTextSearch")
	mstrTextSearch = trim(LoadRequestValue("TextSearch"))
	If (Len(mradTextSearch) > 0) And (Len(mstrTextSearch) > 0) Then
		pstrsqlWhere =  maryCustomValues(mradTextSearch)(enCustomField_FieldName) & " Like '%" & sqlSafe(mstrTextSearch) & "%'"
	End If

	For i = 0 To UBound(maryCustomValues)
		If  maryCustomValues(i)(enCustomField_ShowInSummary) Then
		Select Case maryCustomValues(i)(enCustomField_DisplayType)
			Case enDisplayType_hidden:
				pstrSelFilter = Trim(Request.Form("selFilter" & i ))
				'Added to force filter to use hidden contentType
				If maryCustomValues(i)(enCustomField_FieldName) = "contentContentType" Then
					If Len(pstrSelFilter) = 0 Then pstrSelFilter = mclsItem.FixedContentType
				End If
				If len(pstrSelFilter) > 0 then
					If len(pstrsqlWhere) > 0 Then
						pstrsqlWhere = pstrsqlWhere & " and " & maryCustomValues(i)(enCustomField_FieldName) & "=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
						'pstrsqlWhere = pstrsqlWhere & " and " & maryCustomValues(i)(enCustomField_FieldName) & "=" & pstrSelFilter
					Else
						pstrsqlWhere = maryCustomValues(i)(enCustomField_FieldName) & "=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
						'pstrsqlWhere = maryCustomValues(i)(enCustomField_FieldName) & "=" & pstrSelFilter
					End If
				End If
			Case enDisplayType_textbox_WithDateSelect
				pstrSelFilter = Trim(Request.Form("startDateFilter" & i ))
				If Len(pstrSelFilter) > 0 Then
					If len(pstrsqlWhere) > 0 Then
						pstrsqlWhere = pstrsqlWhere & " and " & maryCustomValues(i)(enCustomField_FieldName) & ">=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
					Else
						pstrsqlWhere = maryCustomValues(i)(enCustomField_FieldName) & ">=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
					End If
				End If

				pstrSelFilter = Trim(Request.Form("endDateFilter" & i ))
				If Len(pstrSelFilter) > 0 Then
					If len(pstrsqlWhere) > 0 Then
						pstrsqlWhere = pstrsqlWhere & " and " & maryCustomValues(i)(enCustomField_FieldName) & "<=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
					Else
						pstrsqlWhere = maryCustomValues(i)(enCustomField_FieldName) & "<=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
					End If
				End If
			Case enDisplayType_select:
				pstrSelFilter = Trim(Request.Form("selFilter" & i ))
				If len(pstrSelFilter) > 0 then
					If len(pstrsqlWhere) > 0 Then
						pstrsqlWhere = pstrsqlWhere & " and " & maryCustomValues(i)(enCustomField_FieldName) & "=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
						'pstrsqlWhere = pstrsqlWhere & " and " & maryCustomValues(i)(enCustomField_FieldName) & "=" & pstrSelFilter
					Else
						pstrsqlWhere = maryCustomValues(i)(enCustomField_FieldName) & "=" & wrapSQLValue(pstrSelFilter, True, maryCustomValues(i)(enCustomField_DataType))
						'pstrsqlWhere = maryCustomValues(i)(enCustomField_FieldName) & "=" & pstrSelFilter
					End If
				End If
			Case enDisplayType_checkbox:
				pstrradFilter = Trim(Request.Form("radFilter" & i ))
				If len(pstrradFilter) > 0 Then
					If CStr(pstrradFilter) = "1" Then
						pstrsqlFragment = "(" & maryCustomValues(i)(enCustomField_FieldName) & "=1 Or " & maryCustomValues(i)(enCustomField_FieldName) & "=-1)"
					ElseIf CStr(pstrradFilter) = "2" Then
						pstrsqlFragment = maryCustomValues(i)(enCustomField_FieldName) & "=0"
					Else
						pstrsqlFragment = ""
					End If
					
					If len(pstrsqlFragment) > 0 Then
						If len(pstrsqlWhere) > 0 Then
							pstrsqlWhere = pstrsqlWhere & " and " & pstrsqlFragment
						Else
							pstrsqlWhere = pstrsqlFragment
						End If
					End If
				End If
		End Select
		End If
	Next 'i

	LoadFilter = pstrsqlWhere
	'Response.Write "pstrsqlWhere: " & pstrsqlWhere & "<BR>"
End Function    'LoadFilter

'******************************************************************************************************************************************************************

Sub WriteItemFilter()

Dim i
Dim plngradTextCounter: plngradTextCounter = 0
Dim plng
%>
<script LANGUAGE=javascript>
<!--

function btnFilter_onclick(theButton)
{
var theForm = theButton.form;

  theForm.Action.value = "Filter";
  theForm.submit();
  return(true);
}

//-->
</script>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
<colgroup align="left">
<colgroup align="left">
  <tr>
    <td valign="top">
        Filter on<br />
		<%
		For i = 0 To UBound(maryCustomValues)
			If  maryCustomValues(i)(enCustomField_ShowInSummary) Then
			Select Case maryCustomValues(i)(enCustomField_DisplayType)
				Case enDisplayType_hidden
				Case enDisplayType_select
				Case enDisplayType_radio
				Case enDisplayType_textbox
					plngradTextCounter = plngradTextCounter + 1
				%>
				<input type="radio" value="<%= i %>" <% If mradTextSearch=CStr(i) Then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch<%= i %>"><label for="radTextSearch<%= i %>"><%= maryCustomValues(i)(enCustomField_DisplayText) %></label><br />
				<%
				Case enDisplayType_checkbox
				Case enDisplayType_listbox
				Case Else
			End Select
			End If	'maryCustomValues(i)(enCustomField_ShowInSummary)
		Next 'i
		
		%>
        <input type="radio" value="" <% if mradTextSearch="" then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch"><label for="radTextSearch">Do Not Include</label>
        <br />containing the text<br />
        <input type=enDisplayType_textbox name="TextSearch" size="20" value="<%= EncodeString(mstrTextSearch,True) %>" ID="TextSearch">
	</td>
	
	<td valign="top" align="center">
		<%
		Dim pstrSelFilter
		For i = 0 To UBound(maryCustomValues)
			If  maryCustomValues(i)(enCustomField_ShowInSummary) Then
			Select Case maryCustomValues(i)(enCustomField_DisplayType)
				Case enDisplayType_hidden
				Case enDisplayType_select
				%>
				<p>
				Filter by <%= maryCustomValues(i)(enCustomField_DisplayText) %><br />
				<select size="1" name="selFilter<%= i %>"  id="selFilter<%= i %>">
				<% 	
				pstrSelFilter = LoadRequestValue("selFilter" & i )
				If i = 1 Then
					If Len(pstrSelFilter) = 0 And Len(Request.Form) = 0 Then pstrSelFilter = 3
				End If
				If len(pstrSelFilter) = 0 then
					Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
				Else
					Response.Write "<option value=''>- All -</Option>" & vbcrlf
				End If
				Call MakeCombo(maryCustomValues(i)(enCustomField_sqlSource),"","",pstrSelFilter)
				%>
				</select></p>
				<%
				Case enDisplayType_radio
				Case enDisplayType_textbox
				Case enDisplayType_textbox_WithDateSelect
				
				Dim mstrStartDate, mstrEndDate
				
					mstrStartDate = Trim(Request.Form("startDateFilter" & i ))
					mstrEndDate = Trim(Request.Form("endDateFilter" & i ))
				%>
				<fieldset style="text-align:left; width:25%">
					<legend> <%= maryCustomValues(i)(enCustomField_DisplayText) %> </legend>
					<label for="startDateFilter<%= i %>">Start Date:&nbsp;</label><input id="startDateFilter<%= i %>" name="startDateFilter<%= i %>" Value="<%= mstrStartDate %>">
					<a HREF="javascript:doNothing()" title="Select start date" onClick="setDateField(document.frmData.startDateFilter<%= i %>); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
					<img SRC="images/calendar.gif" BORDER=0></a><br />

					<label for="endDateFilter<%= i %>">&nbsp;&nbsp;End Date:&nbsp;</label><input id="endDateFilter" name="endDateFilter<%= i %>" Value="<%= mstrEndDate %>">
					<a HREF="javascript:doNothing()" title="Select end date" onClick="setDateField(document.frmData.endDateFilter<%= i %>); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
					<img SRC="images/calendar.gif" BORDER=0></a><br />
				</fieldset>
				<%
				Case enDisplayType_checkbox
					pstrSelFilter = Trim(Request.Form("radFilter" & i ))
				%>
				<input type=radio name="radFilter<%= i %>" id="radFilter<%= i %>_1" value="1" <%= isChecked(CStr(pstrSelFilter) = "1") %>>&nbsp;<label for="radFilter<%= i %>_1">Is <%= maryCustomValues(i)(enCustomField_DisplayText) %></label>
				<input type=radio name="radFilter<%= i %>" id="radFilter<%= i %>_2" value="2" <%= isChecked(CStr(pstrSelFilter) = "2") %>>&nbsp;<label for="radFilter<%= i %>_2">Not <%= maryCustomValues(i)(enCustomField_DisplayText) %></label>
				<input type=radio name="radFilter<%= i %>" id="radFilter<%= i %>_0" value="" <%= isChecked(CStr(pstrSelFilter) = "") %>>&nbsp;<label for="radFilter<%= i %>_0">Do Not Include</label><br />
				<%
				Case enDisplayType_listbox
				Case Else
			End Select
			End If	'maryCustomValues(i)(enCustomField_ShowInSummary)
		Next 'i
		
		%>

	</td>
	<td>
	  <input class="butn" id=btnFilter name=btnFilter type=button value="Apply Filter" onclick="btnFilter_onclick(this);"><br />
	</td>
  </tr>
</table>
<% End Sub	'WriteItemFilter %>